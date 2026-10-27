//! Source-owned typed criteria and lexical visibility provenance. The caller
//! never supplies restoration rows or an accepted filter definition.
use crate::xml::{NS, attrs, xml_text};
use crate::{Result, err, field, local};
use quick_xml::{
    NsReader, Writer,
    events::{BytesStart, Event},
};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use std::{
    collections::BTreeMap,
    fs::{self, File},
    io::BufReader,
    path::Path,
};
use zip::ZipArchive;
pub(super) const URI: &str = workbook_format!(URN_OWNED_FILTER_V1);
pub(super) struct Prepared {
    pub input: Value,
    pub text: String,
    pub revision: String,
    pub clear: bool,
    pub remove: bool,
    pub baseline: BTreeMap<u32, Option<String>>,
    pub default_hidden: bool,
    pub previous_filter_bytes: usize,
    /// What the filter being replaced says that is not a criterion and must be
    /// kept: columns whose dropdown button is hidden or shown (column number
    /// within the range, and the attributes), and how the range was last sorted.
    pub buttons: Vec<(u32, Vec<(String, String)>)>,
    pub sorted: Vec<u8>,
}
/// The parts of a stored filter that are not criteria: button visibility per
/// column and the "sorted by" record.
fn passive(bytes: &[u8]) -> Result<(Vec<(u32, Vec<(String, String)>)>, Vec<u8>)> {
    let mut reader = quick_xml::Reader::from_reader(bytes);
    let mut buttons = Vec::new();
    let mut sorted = Writer::new(Vec::new());
    let (mut depth, mut inside) = (0usize, 0usize);
    loop {
        let event = reader.read_event().map_err(err)?;
        match &event {
            Event::Start(e) | Event::Empty(e) => {
                let empty = matches!(event, Event::Empty(_));
                let name = String::from_utf8_lossy(e.local_name().as_ref()).into_owned();
                if depth == 1 && name == "filterColumn" {
                    let a = clean_attrs(e)?;
                    let kept = ["hiddenButton", "showButton"]
                        .iter()
                        .filter_map(|key| a.get(*key).map(|v| ((*key).to_owned(), v.clone())))
                        .collect::<Vec<_>>();
                    if let (Some(column), false) = (
                        a.get("colId").and_then(|v| v.parse::<u32>().ok()),
                        kept.is_empty(),
                    ) {
                        buttons.push((column, kept));
                    }
                }
                if inside > 0 || (depth == 1 && name == "sortState") {
                    // Written without a prefix, like the filter that will hold it.
                    let mut element = BytesStart::new(name);
                    for attribute in e.attributes() {
                        let attribute = attribute.map_err(err)?;
                        element.push_attribute(attribute);
                    }
                    sorted
                        .write_event(if empty {
                            Event::Empty(element)
                        } else {
                            Event::Start(element)
                        })
                        .map_err(err)?;
                    if !empty {
                        inside += 1;
                    }
                }
                if !empty {
                    depth += 1;
                }
            }
            Event::End(e) => {
                depth = depth.saturating_sub(1);
                if inside > 0 {
                    let name = String::from_utf8_lossy(e.local_name().as_ref()).into_owned();
                    sorted
                        .write_event(Event::End(quick_xml::events::BytesEnd::new(name)))
                        .map_err(err)?;
                    inside -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok((buttons, sorted.into_inner()))
}
fn exact(v: &Value, keys: &[&str]) -> bool {
    v.as_object()
        .is_some_and(|o| o.len() == keys.len() && keys.iter().all(|k| o.contains_key(*k)))
}
fn clean_attrs(e: &quick_xml::events::BytesStart<'_>) -> Result<BTreeMap<String, String>> {
    // Namespace declarations and attributes from another namespace (Excel's
    // xr:uid revision ids) say nothing about the filter itself.
    Ok(attrs(e)?
        .into_iter()
        .filter(|(k, _)| !k.starts_with("xmlns") && !k.contains(':'))
        .collect())
}
fn digest(text: &str) -> String {
    format!("{:x}", Sha256::digest(text.as_bytes()))
}
fn values(model: &Value) -> Result<Vec<Value>> {
    let list = model["values"]
        .as_array()
        .filter(|v| v.len() <= 1000)
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid filter values"))?;
    list.iter()
        .map(|v| {
            if exact(v, &["$workbookNumber"]) && v["$workbookNumber"] == "-0" {
                Ok(json!(-0.0))
            } else if v.is_null()
                || v.is_boolean()
                || v.is_string() && v.as_str().unwrap().len() <= 32768
                || v.as_f64().is_some_and(f64::is_finite)
            {
                Ok(v.clone())
            } else {
                Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid typed filter value",
                ))
            }
        })
        .collect()
}
fn validate(model: &Value) -> Result<(BTreeMap<u32, Option<String>>, bool)> {
    if !exact(model, &["schema", "range", "column", "values", "baseline"])
        || model["schema"] != 1
        || !exact(&model["baseline"], &["format", "rows"])
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid owned filter provenance",
        ));
    }
    let (first, last) = crate::bulk::bounds(field(model, "range")?)?;
    if first.0 >= last.0 {
        return Err(crate::coded(
            "filter-bad-input",
            "Filter requires a header and body",
        ));
    }
    let col = crate::xml::position(&format!("{}1", field(model, "column")?))?.1;
    if col < first.1 || col > last.1 {
        return Err(crate::coded(
            "filter-unsupported",
            "Filter column is outside the requested range",
        ));
    }
    values(model)?;
    let format = model["baseline"]["format"]
        .as_object()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid filter source defaults"))?;
    if format
        .iter()
        // Excel writes prefixed attributes such as x14ac:dyDescent on nearly every sheet.
        .any(|(k, v)| {
            let name = k.split_once(':').map_or(k.as_str(), |(prefix, local)| {
                if prefix.is_empty()
                    || !prefix
                        .bytes()
                        .all(|b| b.is_ascii_alphanumeric() || b == b'_')
                {
                    ""
                } else {
                    local
                }
            });
            name.is_empty()
                || !name.bytes().all(|b| b.is_ascii_alphanumeric() || b == b'_')
                || !v.is_string()
        })
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid filter source defaults",
        ));
    }
    let hidden = format.get("zeroHeight").map(Value::as_str).flatten();
    if hidden.is_some_and(|v| !["0", "1", "true", "false"].contains(&v)) {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid default row visibility",
        ));
    }
    let mut baseline = BTreeMap::new();
    let mut previous = first.0;
    for row in model["baseline"]["rows"]
        .as_array()
        .filter(|r| r.len() <= 1048576)
        .ok_or_else(|| {
            crate::coded(
                "internal-invalid-field",
                "Invalid filter visibility baseline",
            )
        })?
    {
        let row = row.as_array().filter(|r| r.len() == 2).ok_or_else(|| {
            crate::coded(
                "internal-invalid-field",
                "Invalid filter visibility baseline",
            )
        })?;
        let r = row[0]
            .as_u64()
            .filter(|r| *r > u64::from(previous) && *r <= u64::from(last.0))
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid filter baseline row"))?
            as u32;
        let value = if row[1].is_null() {
            None
        } else {
            Some(
                row[1]
                    .as_str()
                    .filter(|v| ["0", "1", "true", "false"].contains(v))
                    .ok_or_else(|| {
                        crate::coded("internal-invalid-field", "Invalid lexical row visibility")
                    })?
                    .to_owned(),
            )
        };
        baseline.insert(r, value);
        previous = r;
    }
    Ok((baseline, hidden.is_some_and(|v| ["1", "true"].contains(&v))))
}
struct Source {
    format: BTreeMap<String, String>,
    rows: BTreeMap<u32, Option<String>>,
    filter: Option<Vec<u8>>,
    unknown_mode: bool,
}
fn scan_source<R: std::io::BufRead>(reader: R, worksheet: bool) -> Result<Source> {
    let mut reader = NsReader::from_reader(reader);
    let mut buffer = Vec::new();
    let mut depth = 0;
    let mut format = None;
    let mut rows = BTreeMap::new();
    let mut filter = None;
    let mut capture: Option<(usize, Writer<Vec<u8>>)> = None;
    let mut unknown_mode = false;
    loop {
        crate::cancel::check()?;
        let (ns, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        let main = ns == quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(NS));
        match &event {
            Event::Start(e) | Event::Empty(e) => {
                let empty = matches!(event, Event::Empty(_));
                let name = e.local_name();
                if main && depth == 1 && name.as_ref() == b"autoFilter" {
                    if filter.is_some() || capture.is_some() {
                        return Err(crate::coded("filter-ambiguous", "Ambiguous saved filters"));
                    }
                    if empty {
                        let mut w = Writer::new(Vec::new());
                        w.write_event(event.clone()).map_err(err)?;
                        filter = Some(w.into_inner())
                    } else {
                        capture = Some((depth, Writer::new(Vec::new())))
                    }
                }
                if worksheet && main && depth == 1 && name.as_ref() == b"sheetFormatPr" {
                    if format.is_some() {
                        return Err(crate::coded(
                            "filter-ambiguous",
                            "Ambiguous source defaults",
                        ));
                    }
                    format = Some(clean_attrs(e)?)
                }
                if worksheet && main && depth == 1 && name.as_ref() == b"sheetPr" {
                    unknown_mode = clean_attrs(e)?
                        .get("filterMode")
                        .is_some_and(|v| ["1", "true"].contains(&v.as_str()))
                }
                if worksheet && main && depth == 2 && name.as_ref() == b"row" {
                    let a = clean_attrs(e)?;
                    let row = a
                        .get("r")
                        .ok_or_else(|| crate::coded("internal-missing-field", "Missing row index"))?
                        .parse::<u32>()
                        .map_err(err)?
                        .checked_sub(1)
                        .filter(|r| *r < 1048576)
                        .ok_or_else(|| crate::coded("filter-malformed", "Invalid row index"))?;
                    let hidden = a.get("hidden").cloned();
                    if hidden
                        .as_deref()
                        .is_some_and(|v| !["0", "1", "true", "false"].contains(&v))
                    {
                        return Err(crate::coded(
                            "internal-unknown-reference",
                            "Unknown source row visibility",
                        ));
                    }
                    if rows.insert(row, hidden).is_some() {
                        return Err(crate::coded("filter-duplicate", "Duplicate source row"));
                    }
                }
                if let Some((_, w)) = &mut capture {
                    w.write_event(event.clone()).map_err(err)?;
                    if w.get_ref().len() > 16 * 1024 * 1024 {
                        return Err(crate::coded(
                            "filter-limit",
                            "Filter metadata exceeds its budget",
                        ));
                    }
                }
                if !empty {
                    depth += 1
                }
            }
            Event::End(_) => {
                depth -= 1;
                if let Some((start, w)) = &mut capture {
                    w.write_event(event.clone()).map_err(err)?;
                    if depth == *start {
                        filter = Some(capture.take().unwrap().1.into_inner())
                    }
                }
            }
            Event::Eof => break,
            _ => {
                if let Some((_, w)) = &mut capture {
                    w.write_event(event.clone()).map_err(err)?
                }
            }
        }
        buffer.clear();
    }
    Ok(Source {
        format: format.unwrap_or_default(),
        rows,
        filter,
        unknown_mode,
    })
}
fn record(bytes: &[u8]) -> Result<Option<(Value, String, String)>> {
    let mut first = quick_xml::Reader::from_reader(bytes);
    let event = first.read_event().map_err(err)?;
    let mut normalized = Vec::new();
    if let Event::Start(e) | Event::Empty(e) = &event {
        let q = String::from_utf8_lossy(e.name().as_ref()).into_owned();
        let binding = q
            .split_once(':')
            .map(|(p, _)| format!("xmlns:{p}"))
            .unwrap_or("xmlns".into());
        if !attrs(e)?.contains_key(&binding) {
            let mut e = e.clone().into_owned();
            e.push_attribute((binding.as_str(), std::str::from_utf8(NS).unwrap()));
            let mut writer = Writer::new(Vec::new());
            writer
                .write_event(if matches!(event, Event::Empty(_)) {
                    Event::Empty(e)
                } else {
                    Event::Start(e)
                })
                .map_err(err)?;
            normalized = writer.into_inner();
            normalized.extend_from_slice(&bytes[first.buffer_position() as usize..]);
        }
    }
    let bytes = if normalized.is_empty() {
        bytes
    } else {
        normalized.as_slice()
    };
    let mut reader = NsReader::from_reader(bytes);
    let mut buffer = Vec::new();
    let mut root = None;
    let mut stack = Vec::<String>::new();
    let mut revision = None;
    let mut text = String::new();
    let mut columns = Vec::<(String, Option<String>, Vec<String>)>::new();
    let mut extensions = 0;
    loop {
        let (ns, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        let main = ns == quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(NS));
        let owned = ns
            == quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(URI.as_bytes()))
            || ns
                == quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(
                    workbook_format!(URN_OWNED_FILTER_V1).as_bytes(),
                ));
        match &event {
            Event::Start(e) | Event::Empty(e) => {
                let empty = matches!(event, Event::Empty(_));
                let name = String::from_utf8_lossy(e.local_name().as_ref()).into_owned();
                let a = clean_attrs(e)?;
                match (stack.len(), name.as_str()) {
                    (0, "autoFilter") if main && a.len() == 1 && a.contains_key("ref") => {
                        root = a.get("ref").cloned()
                    }
                    (1, "filterColumn")
                        if main
                            && a.contains_key("colId")
                            && (a.len() == 1
                                || !empty
                                    && a.keys().all(|k| {
                                        ["colId", "hiddenButton", "showButton"]
                                            .contains(&k.as_str())
                                    })) =>
                    {
                        columns.push((a["colId"].clone(), None, Vec::new()))
                    }
                    // A column that only hides or shows its dropdown button filters
                    // nothing, and a record of how the range was last sorted is an
                    // indicator, not a criterion. Neither makes the filter unknown.
                    (1, "filterColumn")
                        if main
                            && empty
                            && a.contains_key("colId")
                            && a.keys().all(|k| {
                                ["colId", "hiddenButton", "showButton"].contains(&k.as_str())
                            }) => {}
                    (1, "sortState") if main => {}
                    (2, "sortCondition")
                        if main && stack.last().is_some_and(|parent| parent == "sortState") => {}
                    (2, "filters") if main && a.keys().all(|k| k == "blank") => {
                        if let Some(v) = a.get("blank") {
                            if v != "1" {
                                return Err(crate::coded(
                                    "internal-unknown-reference",
                                    "Unknown saved blank criteria",
                                ));
                            }
                            columns
                                .last_mut()
                                .ok_or_else(|| {
                                    crate::coded("internal-missing-field", "Missing filter column")
                                })?
                                .1 = Some(v.clone())
                        }
                    }
                    (3, "filter") if main && a.len() == 1 && a.contains_key("val") && empty => {
                        columns
                            .last_mut()
                            .ok_or_else(|| {
                                crate::coded("internal-missing-field", "Missing filter column")
                            })?
                            .2
                            .push(a["val"].clone())
                    }
                    (1, "extLst") if main && a.is_empty() => extensions += 1,
                    (2, "ext")
                        if main
                            && a.len() == 1
                            && a.get("uri").is_some_and(|value| {
                                format_matches!(value.as_str(), URN_OWNED_FILTER_V1)
                            }) => {}
                    (3, "filterState")
                        if owned
                            && a.len() == 1
                            && a.contains_key("revision")
                            && revision.is_none() =>
                    {
                        revision = a.get("revision").cloned()
                    }
                    _ => {
                        return Err(crate::coded(
                            "internal-unknown-reference",
                            "Unknown imported filter criteria or extension",
                        ));
                    }
                }
                if !empty {
                    stack.push(name)
                }
            }
            Event::End(_) => {
                stack.pop();
            }
            Event::Text(t) => {
                let value = xml_text(t.as_ref())?;
                if stack.last().is_some_and(|s| s == "filterState") {
                    text.push_str(&value)
                } else if !value.trim().is_empty() {
                    return Err(crate::coded("filter-malformed", "Mixed saved filter text"));
                }
            }
            Event::CData(t) if stack.last().is_some_and(|s| s == "filterState") => {
                text.push_str(&t.decode().map_err(err)?)
            }
            Event::Eof => break,
            Event::Comment(_) => {}
            _ => {
                return Err(crate::coded(
                    "filter-unsupported",
                    "Unsupported saved filter XML",
                ));
            }
        }
        if text.len() > 16 * 1024 * 1024 {
            return Err(crate::coded(
                "filter-limit",
                "Filter metadata exceeds its budget",
            ));
        }
        buffer.clear();
    }
    let Some(revision) = revision else {
        if columns.is_empty() && extensions == 0 {
            return Ok(None);
        }
        return Err(crate::coded(
            "filter-unsupported",
            "Imported filter criteria have no owned visibility provenance",
        ));
    };
    if extensions != 1
        || revision.len() != 64
        || !revision
            .bytes()
            .all(|b| b.is_ascii_digit() || (b'a'..=b'f').contains(&b))
        || digest(&text) != revision
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Filter revision differs from owned provenance",
        ));
    }
    let model: Value = serde_json::from_str(&text).map_err(err)?;
    validate(&model)?;
    if root.as_deref() != model["range"].as_str() {
        return Err(crate::coded(
            "internal-mismatch",
            "Saved filter range differs from owned criteria",
        ));
    }
    let values = values(&model)?;
    if values.is_empty() {
        if !columns.is_empty() {
            return Err(crate::coded(
                "filter-malformed",
                "Unexpected active filter criteria",
            ));
        }
    } else {
        let (first, _) = crate::bulk::bounds(field(&model, "range")?)?;
        let col = crate::xml::position(&format!("{}1", field(&model, "column")?))?.1;
        let selected = values
            .iter()
            .filter(|v| !v.is_null() && *v != &json!(""))
            .collect::<Vec<_>>();
        if columns.len() != 1
            || columns[0].0 != (col - first.1).to_string()
            || columns[0].1.as_deref()
                != values.iter().any(|v| v.is_null() || v == "").then_some("1")
            || columns[0].2.len() != selected.len()
        {
            return Err(crate::coded(
                "internal-mismatch",
                "Saved filter criteria differ from typed criteria",
            ));
        }
        for (v, actual) in selected.iter().zip(&columns[0].2) {
            let equal = if let Some(n) = v.as_f64() {
                actual.parse::<f64>().is_ok_and(|a| {
                    a == n && (n != 0.0 || a.is_sign_negative() == n.is_sign_negative())
                })
            } else if let Some(b) = v.as_bool() {
                actual == if b { "1" } else { "0" }
            } else {
                v.as_str() == Some(actual.as_str())
            };
            if !equal {
                return Err(crate::coded(
                    "internal-mismatch",
                    "Saved filter criteria differ from typed criteria",
                ));
            }
        }
    }
    Ok(Some((model, text, revision)))
}
pub(super) fn prepare(root: &Path, input: &Value) -> Result<Prepared> {
    let kind = field(input, "kind")?;
    let follow = kind != "filterRows";
    if follow
        && input.as_object().unwrap().keys().any(|k| {
            ![
                "action",
                "operation",
                "kind",
                "root",
                "source",
                "version",
                "sheet",
                "view",
                "expectedFilterRevision",
            ]
            .contains(&k.as_str())
        })
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid accepted filter command fields",
        ));
    }
    let version = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing filter version"))?;
    let actual = crate::package_edit::auditing_metadata(root)?;
    let supplied = input["view"]["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing source sheet metadata"))?
        .iter()
        .map(|s| json!({"name":s["name"],"path":s["path"]}))
        .collect::<Vec<_>>();
    if actual["sheets"] != json!(supplied)
        || !crate::package_edit::same_records(&actual["tables"], &input["view"]["tables"])
        || !crate::package_edit::same_records(&actual["namedRanges"], &input["view"]["namedRanges"])
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Filter metadata differs from accepted source",
        ));
    }
    let source_path = local(root, field(input, "source")?)?;
    let observed = crate::hash(&source_path)?;
    let index: Value =
        serde_json::from_slice(&fs::read(local(root, "index.json")?).map_err(err)?).map_err(err)?;
    if field(input, "source")? == "source.xlsx" {
        if version != 0 || observed != index["revision"] {
            return Err(crate::coded(
                "filter-bad-input",
                "Filter source does not bind the accepted version",
            ));
        }
    } else {
        let generated = crate::export::write(
            root,
            &json!({"version":version,"revision":index["revision"],"calculationMetadata":true}),
        )?;
        if generated["revision"] != observed {
            return Err(crate::coded(
                "filter-bad-input",
                "Filter source does not bind the accepted version",
            ));
        }
    }
    let sheet = field(input, "sheet")?;
    let mut zip = ZipArchive::new(File::open(&source_path).map_err(err)?).map_err(err)?;
    let source = scan_source(BufReader::new(zip.by_name(sheet).map_err(err)?), true)?;
    let name =
        actual["sheets"]
            .as_array()
            .unwrap()
            .iter()
            .find(|s| s["path"] == sheet)
            .ok_or_else(|| {
                crate::coded("internal-unknown-reference", "Unknown source worksheet")
            })?["name"]
            .clone();
    let tables = actual["tables"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|t| t["sheet"] == name)
        .collect::<Vec<_>>();
    let mut candidates = Vec::new();
    if let Some(bytes) = source.filter.as_ref() {
        candidates.push((None, bytes.clone()))
    }
    for t in &tables {
        let s = scan_source(
            BufReader::new(zip.by_name(field(t, "path")?).map_err(err)?),
            false,
        )?;
        if let Some(bytes) = s.filter {
            candidates.push((Some(field(t, "path")?.to_owned()), bytes))
        }
    }
    let mut owned = None;
    for (path, bytes) in &candidates {
        if let Some(r) = record(bytes)? {
            if owned.is_some() {
                return Err(crate::coded(
                    "filter-protected",
                    "Multiple owned worksheet filters are protected",
                ));
            }
            owned = Some((path.clone(), r))
        }
    }
    let mut request = input.clone();
    let mut model = if let Some((_, (m, _, revision))) = owned.as_ref() {
        if follow {
            if input["expectedFilterRevision"].as_str() != Some(revision) {
                return Err(crate::coded("filter-stale", "Filter definition is stale"));
            }
            if kind != "removeFilter" && values(m)?.is_empty() {
                return Err(crate::coded(
                    "filter-bad-input",
                    "No active filter criteria",
                ));
            }
            request["range"] = m["range"].clone();
            request["column"] = m["column"].clone();
            request["values"] = json!(values(m)?);
        } else if input["range"] != m["range"] {
            return Err(crate::coded(
                "filter-unsupported",
                "Remove the existing owned filter before changing its range",
            ));
        }
        m.clone()
    } else {
        if follow {
            return Err(crate::coded(
                "filter-bad-input",
                "Filter has no owned visibility provenance",
            ));
        }
        if source.unknown_mode {
            return Err(crate::coded(
                "filter-unsupported",
                "Imported filtering visibility has no owned provenance",
            ));
        }
        let (first, last) = crate::bulk::bounds(field(input, "range")?)?;
        json!({"schema":1,"range":input["range"],"column":input["column"],"values":input["values"],"baseline":{"format":source.format,"rows":source.rows.iter().filter(|(r,_)|**r>first.0&&**r<=last.0).map(|(r,v)|json!([r,v])).collect::<Vec<_>>()}})
    };
    if kind == "filterRows" {
        model["column"] = request["column"].clone();
        model["values"] = request["values"].clone();
    }
    if let Some(bytes) = source.filter.as_ref() {
        let mut reader = quick_xml::Reader::from_reader(bytes.as_slice());
        let event = reader.read_event().map_err(err)?;
        if let Event::Start(e) | Event::Empty(e) = event {
            if attrs(&e)?.get("ref").map(String::as_str) != request["range"].as_str() {
                return Err(crate::coded(
                    "filter-unsupported",
                    "Replacing unrelated filters is unsupported",
                ));
            }
        }
    }
    let clear = ["clearFilter", "removeFilter"].contains(&kind)
        || request["values"].as_array().is_some_and(Vec::is_empty);
    if clear {
        model["values"] = json!([]);
        request["values"] = json!([])
    }
    for v in model["values"]
        .as_array_mut()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing filter values"))?
    {
        if v.as_f64().is_some_and(|n| n == 0.0 && n.is_sign_negative()) {
            *v = json!({"$workbookNumber":"-0"})
        }
    }
    let (baseline, default_hidden) = validate(&model)?;
    let format: Value = json!(source.format);
    if model["baseline"]["format"] != format {
        return Err(crate::coded(
            "internal-mismatch",
            "Worksheet defaults differ from the owned visibility baseline",
        ));
    }
    let text = model.to_string();
    if text.len() > 16 * 1024 * 1024 {
        return Err(crate::coded(
            "filter-limit",
            "Filter metadata exceeds its budget",
        ));
    }
    let revision = digest(&text);
    let target = if tables.is_empty() {
        None
    } else {
        Some(
            field(
                tables
                    .iter()
                    .find(|t| t["ref"] == request["range"])
                    .ok_or_else(|| {
                        crate::coded(
                            "filter-bad-input",
                            "Filter must match the complete table range",
                        )
                    })?,
                "path",
            )?
            .to_owned(),
        )
    };
    if target.is_some() && source.filter.is_some() {
        return Err(crate::coded(
            "filter-conflict",
            "Overlapping worksheet and table filters are protected",
        ));
    }
    let previous = candidates.iter().find(|(p, _)| *p == target);
    let previous_filter_bytes = previous.map(|(_, b)| b.len()).unwrap_or(0);
    let (buttons, sorted) = match previous {
        Some((_, bytes)) => passive(bytes)?,
        None => (Vec::new(), Vec::new()),
    };
    Ok(Prepared {
        input: request,
        text,
        revision,
        clear,
        remove: kind == "removeFilter",
        baseline,
        default_hidden,
        previous_filter_bytes,
        buttons,
        sorted,
    })
}
impl Prepared {
    pub fn metadata_budget(&self, root: &Path, new_filter_bytes: usize) -> Result<()> {
        let mut zip = ZipArchive::new(File::open(local(root, "metadata.xlsx")?).map_err(err)?)
            .map_err(err)?;
        let mut total = 0u64;
        for i in 0..zip.len() {
            total = total
                .checked_add(zip.by_index(i).map_err(err)?.size())
                .ok_or_else(|| crate::coded("filter-limit", "Filter metadata budget overflow"))?;
        }
        if total
            .saturating_sub(self.previous_filter_bytes as u64)
            .saturating_add(new_filter_bytes as u64)
            > 16 * 1024 * 1024
        {
            return Err(crate::coded(
                "filter-limit",
                "Filtered workbook metadata exceeds its existing budget",
            ));
        }
        Ok(())
    }

    pub fn lexical(&self, row: u32, filtered: bool) -> Option<String> {
        let value = self.baseline.get(&row).cloned().flatten();
        let hidden = value.as_deref().is_some_and(|s| ["1", "true"].contains(&s))
            || value.is_none() && self.default_hidden;
        if filtered && !hidden {
            Some("1".into())
        } else {
            value
        }
    }
    pub fn inherited_hidden(&self, row: u32) -> bool {
        self.lexical(row, false)
            .as_deref()
            .is_some_and(|s| ["1", "true"].contains(&s))
            || self.lexical(row, false).is_none() && self.default_hidden
    }
}
#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;
    #[test]
    fn total_metadata_budget_includes_existing_parts_and_new_owned_record() {
        let dir = tempfile::tempdir().unwrap();
        let file = File::create(dir.path().join("metadata.xlsx")).unwrap();
        let mut zip = zip::ZipWriter::new(file);
        zip.start_file("metadata.xml", zip::write::SimpleFileOptions::default())
            .unwrap();
        zip.write_all(&vec![b' '; 16 * 1024 * 1024 - 100]).unwrap();
        zip.finish().unwrap();
        let p = Prepared {
            input: json!({}),
            text: String::new(),
            revision: String::new(),
            clear: false,
            remove: false,
            baseline: BTreeMap::new(),
            default_hidden: false,
            previous_filter_bytes: 0,
            buttons: Vec::new(),
            sorted: Vec::new(),
        };
        assert!(p.metadata_budget(dir.path(), 100).is_ok());
        assert!(p.metadata_budget(dir.path(), 101).is_err());
    }
}
