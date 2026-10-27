//! Bounded mutation evidence for MCP. Compare changed tiles near storage and
//! return at most 100 cell records plus an exact count for each work page.
use crate::range::read_db;
use crate::xml::address;
use crate::{Result, database, err};
use rusqlite::{Connection, OptionalExtension, params};
use serde_json::{Value, json};
use std::{
    collections::{BTreeMap, HashMap},
    path::Path,
};
type ReceiptDefinition = (u32, u32, String, String, bool);
#[derive(Default)]
struct ReceiptDefinitions {
    items: HashMap<(u64, u64, String, String), Option<ReceiptDefinition>>,
    bytes: usize,
}
/// A compact normalization changes representation, not authored formulas.
/// Canonicalize a raw side only when the opposite version has a completed
/// normalization proof and this side's actual master still agrees with it.
fn receipt_cell(
    db: &Connection,
    version: u64,
    other: u64,
    sheet: &str,
    reference: &str,
    cell: &Value,
    cache: &mut ReceiptDefinitions,
) -> Result<Value> {
    if cell["formulaType"] != "shared" {
        return Ok(cell.clone());
    }
    let id = cell["sharedFormulaId"].to_string();
    let key = (version, other, sheet.to_owned(), id.clone());
    let definition = if let Some(value) = cache.items.get(&key) {
        value.clone()
    } else {
        let candidate:Option<(u32,u32,String,bool)>=db.prepare_cached("WITH RECURSIVE ancestry(id,depth) AS (SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN ancestry a ON v.id=a.id WHERE v.parent IS NOT NULL) SELECT g.master_r,g.master_c,g.formula,g.invariant FROM ancestry a JOIN shared_normalizations g ON g.version=a.id WHERE g.sheet=?2 AND g.group_id=?3 ORDER BY a.depth LIMIT 1").map_err(err)?.query_row(params![other,sheet,id],|r|Ok((r.get(0)?,r.get(1)?,r.get(2)?,r.get(3)?))).optional().map_err(err)?;
        let mut proved = None;
        if let Some((mr, mc, formula, invariant)) = candidate {
            let source_ref:Option<String>=db.prepare_cached("SELECT json_extract(formula_attrs,'$.ref') FROM cells WHERE sheet=?1 AND r=?2 AND c=?3 AND json_extract(formula_attrs,'$.t')='shared' AND json_quote(json_extract(formula_attrs,'$.si'))=?4").map_err(err)?.query_row(params![sheet,mr,mc,id],|r|r.get(0)).optional().map_err(err)?.flatten();
            let master = read_db(
                db,
                &json!({"sheet":sheet,"range":address(mr,mc),"version":version,"limit":1}),
            )?;
            let master = &master["cells"][0][1];
            if master["formulaType"] == "shared"
                && master["sharedFormulaId"] == cell["sharedFormulaId"]
                && master["formula"].as_str() == Some(&formula)
                && let Some(bounds) = source_ref
                && master["formulaReference"].as_str() == Some(&bounds)
            {
                proved = Some((mr, mc, formula, bounds, invariant));
            }
        }
        let cost =
            sheet.len() + id.len() + proved.as_ref().map_or(0, |d| d.2.len() + d.3.len()) + 192;
        if cache.items.len() < 64 && cache.bytes + cost <= 65536 {
            cache.items.insert(key, proved.clone());
            cache.bytes += cost;
        }
        proved
    };
    let Some((mr, mc, formula, bounds, invariant)) = definition else {
        return Ok(cell.clone());
    };
    let at = crate::xml::position(reference)?;
    let (a, b) = crate::bulk::bounds(&bounds)?;
    if at.0 < a.0
        || at.0 > b.0
        || at.1 < a.1
        || at.1 > b.1
        || (at != (mr, mc) && cell["formula"].as_str().is_some_and(|f| !f.is_empty()))
    {
        return Ok(cell.clone());
    }
    let shifted = if invariant {
        json!(formula)
    } else {
        crate::formula::rewrite(&json!({"items":[{"formula":formula,"copyRows":i64::from(at.0)-i64::from(mr),"copyColumns":i64::from(at.1)-i64::from(mc)}]}))?["formulas"][0].clone()
    };
    let mut value = cell.clone();
    value["formula"] = shifted;
    value["value"] = Value::Null;
    Ok(value)
}
fn receipt_value(cell: &Value) -> Value {
    if cell["formula"].is_string() || cell.get("formulaType").is_some() {
        if cell["calculation"]["status"] == "calculated" {
            cell["calculation"]["value"].clone()
        } else {
            Value::Null
        }
    } else {
        cell["value"].clone()
    }
}
fn semantic(cell: &Value, styles: &Value) -> Value {
    let mut style = styles[cell["styleId"].as_u64().unwrap_or(0) as usize]
        .as_object()
        .cloned()
        .unwrap_or_default();
    for key in [
        "format",
        "bold",
        "italic",
        "underline",
        "fontFamily",
        "fontSize",
        "fontColor",
        "fillColor",
        "horizontalAlign",
        "verticalAlign",
        "wrapText",
        "border",
    ] {
        if let Some(v) = cell.get(key) {
            style.insert(key.into(), v.clone());
        }
    }
    let mut calculation = cell["calculation"].clone();
    if let Some(map) = calculation.as_object_mut() {
        map.remove("revision");
        map.remove("nativePartial");
    }
    json!({"value":receipt_value(cell),"formula":cell["formula"],"style":style,"calculation":calculation})
}
/// Saved overlays can contain JSON integers while an imported cell contains
/// the equivalent float. Compare that representation change exactly, without
/// rounding distinct large integers through f64 or ignoring authored changes.
fn semantic_equal(a: &Value, b: &Value) -> bool {
    if a == b {
        return true;
    }
    match (a, b) {
        (Value::Number(a), Value::Number(b)) if a.is_f64() != b.is_f64() => {
            let (integer, float) = if a.is_f64() {
                (b, a.as_f64().unwrap())
            } else {
                (a, b.as_f64().unwrap())
            };
            if float.fract() != 0.0 {
                return false;
            }
            if let Some(integer) = integer.as_u64() {
                // u64::MAX rounds to 2^64 as f64; exclude that boundary before
                // the saturating float-to-integer cast can hide an inequality.
                float >= 0.0 && float < u64::MAX as f64 && float as u64 == integer
            } else if let Some(integer) = integer.as_i64() {
                float >= i64::MIN as f64 && float < 0.0 && float as i64 == integer
            } else {
                false
            }
        }
        (Value::Array(a), Value::Array(b)) => {
            a.len() == b.len() && a.iter().zip(b).all(|(a, b)| semantic_equal(a, b))
        }
        (Value::Object(a), Value::Object(b)) => {
            a.len() == b.len()
                && a.iter()
                    .all(|(key, a)| b.get(key).is_some_and(|b| semantic_equal(a, b)))
        }
        _ => false,
    }
}
pub fn diff(root: &Path, input: &Value) -> Result<Value> {
    let db = database(root)?;
    let before = input["before"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing before version"))?;
    let after = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing after version"))?;
    let offset = input["offset"].as_u64().unwrap_or(0);
    let sql = "WITH RECURSIVE old(id) AS (SELECT ?1 UNION ALL SELECT v.parent FROM versions v JOIN old a ON v.id=a.id WHERE v.parent IS NOT NULL), new(id) AS (SELECT ?2 UNION ALL SELECT v.parent FROM versions v JOIN new a ON v.id=a.id WHERE v.parent IS NOT NULL), different(id) AS (SELECT id FROM old WHERE id NOT IN(SELECT id FROM new) UNION SELECT id FROM new WHERE id NOT IN(SELECT id FROM old)) SELECT DISTINCT sheet,r/32*32,c/32*32 FROM changes WHERE version IN (SELECT id FROM different) ORDER BY sheet,r/32*32,c/32*32 LIMIT 8 OFFSET ?3";
    let mut stmt = db.prepare_cached(sql).map_err(err)?;
    let tiles = stmt
        .query_map(rusqlite::params![before, after, offset], |r| {
            Ok((
                r.get::<_, String>(0)?,
                r.get::<_, u32>(1)?,
                r.get::<_, u32>(2)?,
            ))
        })
        .map_err(err)?
        .collect::<std::result::Result<Vec<_>, _>>()
        .map_err(err)?;
    let mut changes = Vec::new();
    let mut count = 0;
    let mut definitions = ReceiptDefinitions::default();
    for (sheet, row, col) in &tiles {
        let request = json!({"sheet":sheet,"range":format!("{}:{}",address(*row,*col),address(row+31,col+31)),"limit":2048,"version":before});
        let old = read_db(&db, &request)?;
        let mut request = request;
        request["version"] = json!(after);
        let new = read_db(&db, &request)?;
        let old: BTreeMap<&str, &Value> = old["cells"]
            .as_array()
            .unwrap()
            .iter()
            .map(|p| (p[0].as_str().unwrap(), &p[1]))
            .collect();
        let new: BTreeMap<&str, &Value> = new["cells"]
            .as_array()
            .unwrap()
            .iter()
            .map(|p| (p[0].as_str().unwrap(), &p[1]))
            .collect();
        let coordinates: std::collections::BTreeSet<&str> =
            old.keys().chain(new.keys()).copied().collect();
        for reference in coordinates {
            let a = receipt_cell(
                &db,
                before,
                after,
                sheet,
                reference,
                old.get(reference).copied().unwrap_or(&Value::Null),
                &mut definitions,
            )?;
            let b = receipt_cell(
                &db,
                after,
                before,
                sheet,
                reference,
                new.get(reference).copied().unwrap_or(&Value::Null),
                &mut definitions,
            )?;
            if !semantic_equal(
                &semantic(&a, &input["styles"]),
                &semantic(&b, &input["styles"]),
            ) {
                count += 1;
                if changes.len() < 100 {
                    changes.push(json!({"sheet":sheet,"cell":reference,"before":{"value":receipt_value(&a),"formula":a["formula"]},"after":{"value":receipt_value(&b),"formula":b["formula"],"calculation":b["calculation"]}}));
                }
            }
        }
    }
    Ok(
        json!({"changes":changes,"count":count,"nextOffset":if tiles.len()==8 {Some(offset+8)}else{None}}),
    )
}
pub fn summary(root: &Path, input: &Value) -> Result<Value> {
    let db = database(root)?;
    let mut request = input.clone();
    request["range"] = json!("A1:XFD1048576");
    request["formulas"] = json!(true);
    request["limit"] = json!(2048);
    let mut counts = json!({"calculated":0,"unresolved":0,"error":0,"circular":0});
    loop {
        crate::cancel::check()?;
        let page = read_db(&db, &request)?;
        for pair in page["cells"].as_array().unwrap() {
            let status = pair[1]["calculation"]["status"]
                .as_str()
                .unwrap_or("unresolved");
            let status = if counts.get(status).is_some() {
                status
            } else {
                "unresolved"
            };
            counts[status] = json!(counts[status].as_u64().unwrap() + 1);
        }
        if page["nextCursor"].is_null() {
            break;
        }
        request["after"] = page["nextCursor"].clone();
    }
    Ok(counts)
}

/// Compare immutable accepted bases with a bounded merge of cell coordinates.
/// The host may address only another child of this same managed owner directory.
pub fn bases(root: &Path, input: &Value) -> Result<Value> {
    let owner = root
        .parent()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing workbook owner"))?;
    if !owner.join("owner.json").is_file() {
        return Err(crate::coded(
            "workbook-bad-input",
            "Cross-base receipts require a managed owner",
        ));
    }
    let before_root = crate::local(owner, crate::field(input, "beforeBase")?)?;
    if std::fs::canonicalize(&before_root).map_err(err)?.parent() != Some(owner) {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid receipt base",
        ));
    }
    let old_db = database(&before_root)?;
    let new_db = database(root)?;
    let make = |sheet: &Value, version: &Value| json!({"sheet":sheet,"version":version,"range":"A1:XFD1048576","after":input["after"],"limit":1024});
    let old = if input["beforeSheet"].is_string() {
        read_db(&old_db, &make(&input["beforeSheet"], &input["before"]))?
    } else {
        json!({"cells":[]})
    };
    let new = if input["sheet"].is_string() {
        read_db(&new_db, &make(&input["sheet"], &input["version"]))?
    } else {
        json!({"cells":[]})
    };
    let a = old["cells"].as_array().unwrap();
    let b = new["cells"].as_array().unwrap();
    let (mut i, mut j, mut visited, mut count) = (0, 0, 0, 0);
    let mut changes = Vec::new();
    let mut cursor = None;
    let semantic = |cell: &Value, styles: &Value| {
        let mut value = semantic(cell, styles);
        if cell["formula"].is_string() || cell.get("formulaType").is_some() {
            value["value"] = if cell["calculation"]["status"] == "calculated" {
                cell["calculation"]["value"].clone()
            } else {
                Value::Null
            };
        }
        value
    };
    while visited < 1024 && (i < a.len() || j < b.len()) {
        crate::cancel::check()?;
        let left = a
            .get(i)
            .map(|p| crate::xml::position(p[0].as_str().unwrap()))
            .transpose()?;
        let right = b
            .get(j)
            .map(|p| crate::xml::position(p[0].as_str().unwrap()))
            .transpose()?;
        let position = match (left, right) {
            (Some(a), Some(b)) => a.min(b),
            (Some(a), None) | (None, Some(a)) => a,
            _ => break,
        };
        let before = if left == Some(position) {
            let cell = &a[i][1];
            i += 1;
            cell
        } else {
            &Value::Null
        };
        let after = if right == Some(position) {
            let cell = &b[j][1];
            j += 1;
            cell
        } else {
            &Value::Null
        };
        let reference = address(position.0, position.1);
        cursor = Some(reference.clone());
        visited += 1;
        if !semantic_equal(
            &semantic(before, &input["beforeStyles"]),
            &semantic(after, &input["styles"]),
        ) {
            count += 1;
            if changes.len() < 100 {
                changes.push(json!({"cell":reference,"before":{"value":before["value"],"formula":before["formula"]},"after":{"value":after["value"],"formula":after["formula"],"calculation":after["calculation"]}}));
            }
        }
    }
    Ok(
        json!({"changes":changes,"count":count,"nextCursor":if visited==1024{cursor}else{None},"visited":visited}),
    )
}

#[cfg(test)]
mod tests {
    use super::semantic_equal;
    use serde_json::{Value, json};

    #[test]
    fn receipt_numeric_equality_is_exact_across_integer_and_float_storage() {
        for (a, b, expected) in [
            ("1", "1.0", true),
            ("-1", "-1.0", true),
            ("0", "-0.0", true),
            ("1", "1.5", false),
            ("9007199254740992", "9007199254740992.0", true),
            ("9007199254740993", "9007199254740992.0", false),
            ("9007199254740993", "9007199254740992", false),
            ("-9007199254740993", "-9007199254740992.0", false),
            ("-9223372036854775808", "-9223372036854775808.0", true),
            ("-9223372036854775807", "-9223372036854775808.0", false),
            ("-9223372036854775808", "-9223372036854777856.0", false),
            ("9223372036854775807", "9223372036854775808.0", false),
            ("18446744073709549568", "18446744073709549568.0", true),
            ("18446744073709551615", "18446744073709551616.0", false),
        ] {
            let a: Value = serde_json::from_str(a).unwrap();
            let b: Value = serde_json::from_str(b).unwrap();
            assert_eq!(semantic_equal(&a, &b), expected, "{a} vs {b}");
            assert_eq!(semantic_equal(&b, &a), expected, "{b} vs {a}");
        }
    }

    #[test]
    fn receipt_equality_preserves_nested_calculation_style_and_value_changes() {
        let a = json!({"value":1,"formula":"A1+1","style":{"bold":true,"fontSize":11},"calculation":{"value":2,"status":"calculated"}});
        let b = json!({"value":1.0,"formula":"A1+1","style":{"bold":true,"fontSize":11.0},"calculation":{"value":2.0,"status":"calculated"}});
        assert!(semantic_equal(&a, &b));
        for (path, value) in [
            ("/value", json!("1")),
            ("/value", json!(2)),
            ("/formula", json!("A1+2")),
            ("/style/bold", json!(false)),
            ("/style/fontSize", json!(12)),
            ("/calculation/value", json!(3)),
            ("/calculation/status", json!("error")),
        ] {
            let mut changed = b.clone();
            *changed.pointer_mut(path).unwrap() = value;
            assert!(!semantic_equal(&a, &changed), "{path}");
        }
    }
}
