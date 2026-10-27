//! What a sheet's filter is: ours (with the typed criteria it was made from, so
//! it can be reapplied or cleared) or one the editor only carries along.
use super::dom::Dom;
use serde_json::{Value, json};
use sha2::{Digest, Sha256};

const URI: &str = workbook_format!(URN_OWNED_FILTER_V1);
type Said<T> = Result<T, String>;

/// A filter the editor carries but cannot reapply; the sentence becomes the
/// view's reason.
fn said<T>(text: &str) -> Said<T> {
    Err(crate::coded("filter-unsupported", text))
}
fn digest(text: &str) -> String {
    format!("{:x}", Sha256::digest(text.as_bytes()))
}
fn exact(value: &Value, keys: &[&str]) -> bool {
    value
        .as_object()
        .is_some_and(|o| o.len() == keys.len() && keys.iter().all(|k| o.contains_key(*k)))
}
fn is_minus_zero(value: &Value) -> bool {
    exact(value, &["$workbookNumber"]) && value["$workbookNumber"] == "-0"
}
fn decoded(value: &Value) -> Value {
    if is_minus_zero(value) {
        json!(-0.0)
    } else {
        value.clone()
    }
}
fn column_number(letters: &str) -> usize {
    letters
        .bytes()
        .fold(0, |n, c| n * 26 + (c - b'A' + 1) as usize)
}
fn is_column(text: &str) -> bool {
    (1..=3).contains(&text.len()) && text.bytes().all(|b| b.is_ascii_uppercase())
}
/// first row, last row, left column, right column, all from zero.
fn bounds(range: &str) -> Said<(usize, usize, usize, usize)> {
    let bad = || crate::coded("filter-unsupported", "Invalid filter range");
    let (a, b) = range.split_once(':').ok_or_else(bad)?;
    let split = |cell: &str| -> Option<(usize, usize)> {
        let at = cell.find(|c: char| c.is_ascii_digit())?;
        let (letters, digits) = cell.split_at(at);
        (is_column(letters)
            && !digits.starts_with('0')
            && digits.bytes().all(|b| b.is_ascii_digit()))
        .then(|| Some((digits.parse::<usize>().ok()?, column_number(letters))))
        .flatten()
    };
    let ((first, left), (last, right)) = (split(a).ok_or_else(bad)?, split(b).ok_or_else(bad)?);
    let (first, last, left, right) = (first - 1, last - 1, left - 1, right - 1);
    if first >= last || last >= 1048576 || left > right || right >= 16384 {
        return said("Invalid filter range");
    }
    Ok((first, last, left, right))
}
fn criteria(range: &str, column: &Value, values: &[Value]) -> Said<(usize, usize, usize, usize)> {
    let b = bounds(range)?;
    let column = column
        .as_str()
        .filter(|c| is_column(c))
        .ok_or_else(|| crate::coded("filter-unsupported", "Invalid filter column"))?;
    let c = column_number(column) - 1;
    if c < b.2 || c > b.3 {
        return said("Filter column is outside the requested range");
    }
    let sound = |v: &Value| match v {
        Value::Null | Value::Bool(_) => true,
        Value::String(s) => s.len() <= 32768,
        Value::Number(n) => n.as_f64().is_some_and(f64::is_finite),
        _ => false,
    };
    if values.len() > 1000 || !values.iter().all(sound) {
        return said("Invalid filter values");
    }
    Ok(b)
}
/// The typed record is checked as strictly as when it was written.
fn validate(model: &Value) -> Said<Vec<Value>> {
    if !exact(model, &["schema", "range", "column", "values", "baseline"])
        || model["schema"] != 1
        || !exact(&model["baseline"], &["format", "rows"])
    {
        return said("Invalid owned filter provenance");
    }
    let values: Vec<Value> = model["values"]
        .as_array()
        .ok_or_else(|| crate::coded("filter-unsupported", "Invalid filter values"))?
        .iter()
        .map(decoded)
        .collect();
    let b = criteria(
        model["range"].as_str().unwrap_or(""),
        &model["column"],
        &values,
    )?;
    let source = "Invalid filter source defaults";
    let format = model["baseline"]["format"].as_object().ok_or(source)?;
    let name_ok = |k: &str| {
        let local = k.split_once(':').map_or(k, |(_, l)| l);
        let word =
            |s: &str| !s.is_empty() && s.chars().all(|c| c.is_ascii_alphanumeric() || c == '_');
        word(local) && k.split_once(':').is_none_or(|(p, _)| word(p))
    };
    if format.iter().any(|(k, v)| !v.is_string() || !name_ok(k))
        || format
            .get("zeroHeight")
            .is_some_and(|v| !["0", "1", "true", "false"].contains(&v.as_str().unwrap_or("")))
    {
        return said(source);
    }
    let baseline = "Invalid filter visibility baseline";
    let rows = model["baseline"]["rows"].as_array().ok_or(baseline)?;
    if rows.len() > 1048576 {
        return said(baseline);
    }
    let mut previous = b.0 as i64;
    for item in rows {
        let pair = item.as_array().filter(|p| p.len() == 2).ok_or(baseline)?;
        let at = pair[0].as_i64().ok_or(baseline)?;
        let hidden_ok = pair[1].is_null()
            || ["0", "1", "true", "false"].contains(&pair[1].as_str().unwrap_or("?"));
        if at <= previous || at > b.1 as i64 || !hidden_ok {
            return said(baseline);
        }
        previous = at;
    }
    Ok(values)
}

struct Record {
    values: Vec<Value>,
    range: String,
    column: String,
    revision: String,
}

fn attr_names(dom: &Dom, node: usize) -> Vec<&str> {
    dom.nodes[node]
        .attrs
        .iter()
        .map(|a| a.qname.as_str())
        .collect()
}
fn read_record(dom: &Dom, filter: usize) -> Said<Option<Record>> {
    let nodes = dom.child_elements(filter);
    let exts: Vec<usize> = nodes
        .iter()
        .copied()
        .filter(|n| dom.is(*n, "extLst"))
        .collect();
    let button_only = |n: usize| {
        dom.is(n, "filterColumn")
            && dom.child_elements(n).is_empty()
            && attr_names(dom, n)
                .iter()
                .all(|k| ["colId", "hiddenButton", "showButton"].contains(k) || k.contains(':'))
    };
    let passive = |n: usize| button_only(n) || dom.is(n, "sortState");
    if exts.is_empty() {
        if nodes.iter().any(|n| !passive(*n)) {
            return said("Imported filter criteria have no owned visibility provenance");
        }
        return Ok(None);
    }
    let unknown = "Unknown filter extensions";
    if exts.len() != 1 || dom.child_elements(exts[0]).len() != 1 {
        return said(unknown);
    }
    let ext = dom.child_elements(exts[0])[0];
    let records = dom.child_elements(ext);
    if !dom.is(ext, "ext") || dom.attr(ext, "uri") != URI || records.len() != 1 {
        return said(unknown);
    }
    let state = records[0];
    if dom.nodes[state].ns != URI
        || dom.nodes[state].local != "filterState"
        || attr_names(dom, state).join(",") != "revision"
        || !dom.child_elements(state).is_empty()
    {
        return said("Invalid owned filter record");
    }
    let text = dom.text(state);
    let revision = dom.attr(state, "revision").to_string();
    if text.len() > 16 * 1024 * 1024
        || revision.len() != 64
        || !revision
            .bytes()
            .all(|b| matches!(b, b'0'..=b'9' | b'a'..=b'f'))
        || digest(&text) != revision
    {
        return said("Filter revision does not match its provenance");
    }
    let model: Value = serde_json::from_str(&text)
        .map_err(|_| crate::coded("filter-unsupported", "Invalid owned filter record"))?;
    let values = validate(&model)?;
    if model["range"].as_str() != Some(dom.attr(filter, "ref")) {
        return said("Filter range differs from owned criteria");
    }
    let range = model["range"].as_str().unwrap_or("").to_string();
    let column = model["column"].as_str().unwrap_or("").to_string();
    // The criteria in the file must be exactly what the typed record says.
    let selected: Vec<&Value> = values
        .iter()
        .filter(|v| !v.is_null() && v.as_str() != Some(""))
        .collect();
    let expected_blank = values.iter().any(|v| v.is_null() || v.as_str() == Some(""));
    let left = bounds(&range)?.2;
    let expected_column = (column_number(&column) - 1).wrapping_sub(left).to_string();
    let ordinary: Vec<usize> = nodes
        .iter()
        .copied()
        .filter(|n| *n != exts[0] && !passive(*n))
        .collect();
    let expected_count = usize::from(!values.is_empty());
    let different = "Saved filter criteria differ from typed owned criteria";
    let mut actual = Vec::new();
    for n in ordinary {
        let kids = dom.child_elements(n);
        let plain_column = dom.is(n, "filterColumn")
            && dom.attribute(n, "colId").is_some()
            && attr_names(dom, n)
                .iter()
                .all(|k| ["colId", "hiddenButton", "showButton"].contains(k))
            && kids.len() == 1;
        if !plain_column {
            return said("Unknown filter criteria");
        }
        let set = kids[0];
        let entries = dom.child_elements(set);
        let plain_set = dom.is(set, "filters")
            && attr_names(dom, set).iter().all(|k| *k == "blank")
            && entries.iter().all(|v| {
                dom.is(*v, "filter")
                    && attr_names(dom, *v).join(",") == "val"
                    && dom.child_elements(*v).is_empty()
            });
        if !plain_set {
            return said("Unknown filter criteria");
        }
        let blank = dom.attr(set, "blank");
        actual.push((
            dom.attr(n, "colId").to_string(),
            (!blank.is_empty()).then(|| blank.to_string()),
            entries
                .iter()
                .map(|v| dom.attr(*v, "val").to_string())
                .collect::<Vec<_>>(),
        ));
    }
    if actual.len() != expected_count {
        return said(different);
    }
    for (col, blank, got) in &actual {
        let expected_values: Vec<String> = selected
            .iter()
            .map(|v| match v {
                Value::Bool(b) => (u8::from(*b)).to_string(),
                Value::String(s) => s.clone(),
                other => other.to_string(),
            })
            .collect();
        let numbers_match = got.len() == selected.len()
            && got.iter().zip(&selected).all(|(v, s)| match s {
                Value::Number(n) => {
                    let plain = {
                        // ^-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?$
                        let t = v.strip_prefix('-').unwrap_or(v);
                        let (m, e) = t
                            .find(['e', 'E'])
                            .map_or((t, None), |i| (&t[..i], Some(&t[i + 1..])));
                        let digits =
                            |s: &str| !s.is_empty() && s.bytes().all(|b| b.is_ascii_digit());
                        m.split_once('.')
                            .map_or(digits(m), |(a, b)| digits(a) && digits(b))
                            && e.is_none_or(|e| digits(e.strip_prefix(['-', '+']).unwrap_or(e)))
                    };
                    plain
                        && v.parse::<f64>()
                            .ok()
                            .zip(n.as_f64())
                            .is_some_and(|(a, b)| a.to_bits() == b.to_bits())
                }
                _ => true,
            });
        let texts_match = got
            .iter()
            .zip(&selected)
            .zip(&expected_values)
            .all(|((v, s), want)| s.is_number() || v == want);
        if *col != expected_column
            || *blank != expected_blank.then(|| "1".to_string())
            || got.len() != selected.len()
            || !numbers_match
            || !texts_match
        {
            return said(different);
        }
    }
    Ok(Some(Record {
        values,
        range,
        column,
        revision,
    }))
}

/// A filter's definition for the view, and whether it carries our typed record.
pub(super) fn inspect(dom: &Dom, filter: usize) -> (Value, bool) {
    let reference = dom.attr(filter, "ref");
    match read_record(dom, filter) {
        Ok(Some(record)) => {
            let active = !record.values.is_empty();
            let mut definition = json!({
                "range": record.range, "column": record.column, "values": record.values,
                "owned": true, "revision": record.revision, "canReapply": active, "canClear": active,
            });
            if !active {
                definition["reason"] = json!("No active filter criteria");
            }
            (definition, true)
        }
        Ok(None) => (
            json!({"range": reference, "owned": false, "revision": null, "canReapply": false,
                "canClear": false, "reason": "No active owned filter criteria"}),
            false,
        ),
        Err(reason) => (
            json!({"range": reference, "owned": false, "revision": null, "canReapply": false,
                "canClear": false, "reason": crate::split_code(&reason).1}),
            false,
        ),
    }
}

/// Which of a sheet's filters the view shows (the sheet's own, then its tables').
pub(super) fn choose(found: Vec<(Value, bool)>, imported_visibility: bool) -> Option<Value> {
    let owned = found.iter().filter(|(_, record)| *record).count();
    if owned > 1 {
        return Some(
            json!({"range": "", "owned": false, "revision": null, "canReapply": false,
            "canClear": false, "reason": "Multiple owned filters on one worksheet are protected"}),
        );
    }
    let selected = if owned == 1 {
        found.iter().find(|(_, record)| *record)
    } else {
        found
            .iter()
            .find(|(d, _)| d["reason"] != "No active owned filter criteria")
            .or(found.first())
    };
    match selected {
        Some((definition, _)) => Some(definition.clone()),
        None if imported_visibility => Some(json!({"range": "", "owned": false, "revision": null,
            "canReapply": false, "canClear": false,
            "reason": "Imported filtering visibility has no owned provenance"})),
        None => None,
    }
}
