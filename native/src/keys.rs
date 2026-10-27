//! Business-key scans stay near storage; only the requested row window returns.
use crate::bulk::{bounds, literal};
use crate::range::read_db;
use crate::xml::{address, position};
use crate::{Result, database, err, field};
use rusqlite::params;
use serde_json::{Value, json};
use std::path::Path;
fn identity(value: &Value, unique: bool) -> Result<String> {
    Ok(match value {
        Value::Null => "null".into(),
        Value::Bool(v) => format!("b:{v}"),
        Value::String(v) => format!("s:{v}"),
        Value::Number(n) => {
            let n = n
                .as_f64()
                .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid numeric key"))?;
            if !n.is_finite() {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid numeric key",
                ));
            }
            format!(
                "n:{}",
                if unique && n == 0.0 {
                    "0".into()
                } else {
                    n.to_string()
                }
            )
        }
        Value::Object(v) if v.get("$workbookNumber") == Some(&json!("-0")) => {
            if unique {
                "n:0".into()
            } else {
                "n:-0".into()
            }
        }
        _ => {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid business key",
            ));
        }
    })
}
pub fn query(root: &Path, input: &Value) -> Result<Value> {
    let db = database(root)?;
    let sheet = field(input, "sheet")?;
    let (first, last) = bounds(field(input, "range")?)?;
    let col = position(&format!("{}1", field(input, "column")?))?.1;
    if col < first.1 || col > last.1 {
        return Err(crate::coded(
            "workbook-unsupported",
            "Key column is outside the requested range",
        ));
    }
    let values = input["values"]
        .as_array()
        .filter(|v| v.len() <= 1000)
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid key values"))?;
    let unique = input["unique"].as_bool().unwrap_or(false);
    let keys = values
        .iter()
        .map(|v| identity(v, unique))
        .collect::<Result<Vec<_>>>()?;
    db.execute_batch("CREATE TEMP TABLE key_matches(r INTEGER PRIMARY KEY,k TEXT NOT NULL); CREATE TEMP TABLE key_seen(k TEXT PRIMARY KEY);").map_err(err)?;
    let mut counts = vec![0u64; keys.len()];
    let mut first_rows = vec![None; keys.len()];
    let mut record = |row: u32, cell: &Value| -> Result<()> {
        if cell.get("formula").is_some()
            || cell.get("formulaType").is_some()
            || cell["scenarioResult"] == true
        {
            return Err(crate::coded(
                "workbook-bad-input",
                "Business keys must be literal identifiers",
            ));
        }
        let key = identity(&literal(cell, true)?, unique)?;
        if unique {
            let inserted = db
                .execute("INSERT OR IGNORE INTO key_seen(k) VALUES(?1)", [&key])
                .map_err(err)?;
            if inserted == 0 {
                return Err(crate::coded(
                    "workbook-duplicate",
                    "Duplicate key in requested range",
                ));
            }
        }
        let mut matched = false;
        for (index, requested) in keys.iter().enumerate() {
            if requested == &key {
                counts[index] += 1;
                first_rows[index].get_or_insert(row);
                matched = true;
            }
        }
        if matched {
            db.execute(
                "INSERT INTO key_matches(r,k) VALUES(?1,?2)",
                params![row, key],
            )
            .map_err(err)?;
        }
        Ok(())
    };
    let mut after = None;
    let mut next = first.0;
    loop {
        crate::cancel::check()?;
        let page = read_db(
            &db,
            &json!({"sheet":sheet,"version":input["version"],"range":format!("{}:{}",address(first.0,col),address(last.0,col)),"after":after,"limit":2048}),
        )?;
        for item in page["cells"].as_array().unwrap() {
            let row = position(item[0].as_str().unwrap())?.0;
            while next < row {
                record(next, &json!({"value":null}))?;
                next += 1;
            }
            record(row, &item[1])?;
            next = row + 1;
        }
        after = page["nextCursor"].as_str().map(str::to_owned);
        if after.is_none() {
            break;
        }
    }
    while next <= last.0 {
        record(next, &json!({"value":null}))?;
        next += 1;
    }
    let count: u64 = db
        .query_row("SELECT count(*) FROM key_matches", [], |r| r.get(0))
        .map_err(err)?;
    let limit = input["limit"].as_u64().unwrap_or(500).min(501);
    let offset = input["offset"].as_u64().unwrap_or(0);
    let mut statement = db
        .prepare_cached("SELECT r FROM key_matches ORDER BY r LIMIT ?1 OFFSET ?2")
        .map_err(err)?;
    let rows = statement
        .query_map(params![limit, offset], |r| r.get::<_, u32>(0))
        .map_err(err)?
        .collect::<std::result::Result<Vec<_>, _>>()
        .map_err(err)?;
    Ok(json!({"rows":rows,"count":count,"matches":counts,"firstRows":first_rows}))
}
