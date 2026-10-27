//! Bounded work pages: scan near storage and return only matching cells.
use crate::range::read_public_db as read_db;
use crate::{Result, database, err, field};
use rusqlite::params;
use serde_json::{Value, json};
use std::path::Path;

pub fn search(root: &Path, input: &Value) -> Result<Value> {
    let needle = field(input, "query")?.to_lowercase();
    if needle.is_empty() || needle.len() > 32768 {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid search text",
        ));
    }
    let mode = field(input, "mode")?;
    if !["raw", "formula", "display"].contains(&mode) {
        return Err(crate::coded(
            "query-unsupported",
            "Unsupported native search mode",
        ));
    }
    let match_limit =
        match input.get("matchLimit") {
            None | Some(Value::Null) => None,
            Some(value) => Some(value.as_u64().filter(|v| *v > 0 && *v <= 1000).ok_or_else(
                || crate::coded("internal-invalid-field", "Invalid search match limit"),
            )?),
        };
    let formats: Value = if mode == "display" {
        serde_json::from_slice(
            &std::fs::read(root.join("display-formats.json")).map_err(crate::err)?,
        )
        .map_err(crate::err)?
    } else {
        Value::Null
    };
    let db = database(root)?;
    let mut request = input.clone();
    request["range"] = input
        .get("range")
        .cloned()
        .unwrap_or(json!("A1:XFD1048576"));
    request["limit"] = json!(2048);
    let mut cells = Vec::new();
    let mut matches = 0;
    let mut scanned = 0;
    let mut cursor = Value::Null;
    // Each request has a fixed work budget so host cancellation can intervene
    // between requests, even when the first match is at the end of the sheet.
    for _ in 0..4 {
        let page = read_db(&db, &request)?;
        for pair in page["cells"].as_array().unwrap() {
            let cell = &pair[1];
            scanned += 1;
            let value = if mode == "formula" {
                cell.get("formula").unwrap_or(&Value::Null)
            } else {
                cell.get("rawValue")
                    .filter(|v| !v.is_null())
                    .unwrap_or(&cell["value"])
            };
            let content = if mode == "formula" {
                match value.as_str() {
                    Some(formula) => format!("={}", formula.trim_start_matches('=')),
                    None => String::new(),
                }
            } else if mode == "display" {
                crate::display::cell(
                    cell,
                    &formats[cell["styleId"].as_u64().unwrap_or(0) as usize],
                    input["date1904"] == true,
                )?
            } else {
                match value {
                    Value::Null => String::new(),
                    Value::String(s) => s.clone(),
                    Value::Object(o) if o.contains_key("$workbookNumber") => {
                        o["$workbookNumber"].as_str().unwrap_or("").to_owned()
                    }
                    _ => value.to_string(),
                }
            };
            if content.to_lowercase().contains(&needle) {
                matches += 1;
                if cells.len() < 1000 {
                    cells.push(pair.clone());
                }
                if match_limit.is_some_and(|limit| matches >= limit) {
                    return Ok(
                        json!({"cells":cells,"matches":matches,"scanned":scanned,"nextCursor":pair[0]}),
                    );
                }
            }
        }
        cursor = page["nextCursor"].clone();
        if cursor.is_null() {
            break;
        }
        request["after"] = cursor.clone();
    }
    Ok(json!({"cells":cells,"matches":matches,"scanned":scanned,"nextCursor":cursor}))
}

/// Selection summaries scan accepted versions in bounded pages; no cell values
/// cross IPC. Text/booleans/blanks are excluded from numeric statistics. Formula
/// errors and unresolved values are counted explicitly, never treated as zero.
pub fn selection(root: &Path, input: &Value) -> Result<Value> {
    let db = database(root)?;
    let mut request = input.clone();
    request["limit"] = json!(2048);
    request["offset"] = json!(0);
    request["after"] = Value::Null;
    let (mut count, mut numeric, mut errors, mut unresolved, mut unsafe_numbers) =
        (0u64, 0u64, 0u64, 0u64, 0u64);
    let mut sum = 0.0;
    let (mut min, mut max): (Option<f64>, Option<f64>) = (None, None);
    loop {
        crate::cancel::check()?;
        let page = read_db(&db, &request)?;
        for pair in page["cells"]
            .as_array()
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid cell page"))?
        {
            let cell = &pair[1];
            let formula = cell.get("formula").is_some_and(|v| !v.is_null())
                || cell.get("formulaType").is_some_and(|v| !v.is_null())
                || cell["scenarioResult"] == true
                || cell["calculation"]["pending"] == true;
            let value = if formula {
                count += 1;
                match cell["calculation"]["status"].as_str() {
                    Some("calculated") => &cell["calculation"]["value"],
                    Some("error" | "circular") => {
                        errors += 1;
                        continue;
                    }
                    _ => {
                        unresolved += 1;
                        continue;
                    }
                }
            } else {
                let value = &cell["value"];
                if value.is_null() || value.as_str() == Some("") {
                    continue;
                }
                count += 1;
                if cell["dataType"] == "error" {
                    errors += 1;
                    continue;
                }
                value
            };
            if let Some(number) = value.as_f64() {
                if !formula && crate::bulk::literal(cell, false).is_err() {
                    unsafe_numbers += 1;
                    continue;
                }
                numeric += 1;
                sum = crate::decimal::add(sum, number)?;
                min = Some(min.map_or(number, |v| v.min(number)));
                max = Some(max.map_or(number, |v| v.max(number)));
            } else if value.is_object() {
                unsafe_numbers += 1;
            }
        }
        let cursor = &page["nextCursor"];
        if cursor.is_null() {
            break;
        }
        request["after"] = cursor.clone();
    }
    let safe = unsafe_numbers == 0;
    Ok(
        json!({"count":count,"numericCount":numeric,"sum":if safe {Some(sum)} else {None},"average":if safe && numeric>0 {Some(crate::decimal::divide(sum,numeric as f64)?)} else {None},"min":if safe {min} else {None},"max":if safe {max} else {None},"errors":errors,"unresolved":unresolved,"unsafeNumbers":unsafe_numbers}),
    )
}

/// Distinct filter choices stay in quota-bounded, file-backed TEMP storage.
/// Each request deliberately rescans its pinned version: no unique-value map
/// or persistent cache is retained. Header exclusion belongs to the range.
pub fn filter_values(root: &Path, input: &Value) -> Result<Value> {
    let _span = crate::diagnostics::span("filter_values.scan");
    let sheet = field(input, "sheet")?;
    let (first, last) = crate::bulk::bounds(field(input, "range")?)?;
    let column = field(input, "column")?;
    if column.is_empty() || column.len() > 3 || !column.bytes().all(|b| b.is_ascii_uppercase()) {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid filter column",
        ));
    }
    let col = crate::xml::position(&format!("{column}1"))?.1;
    if col < first.1 || col > last.1 {
        return Err(crate::coded(
            "query-unsupported",
            "Filter column is outside the requested range",
        ));
    }
    let query = match input.get("query") {
        None => "",
        Some(Value::String(s)) if s.len() <= 32768 => s,
        _ => {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid filter search text",
            ));
        }
    }
    .to_lowercase();
    let offset = match input.get("offset") {
        None => 0,
        Some(v) => v
            .as_u64()
            .filter(|v| *v <= 9_007_199_254_740_991)
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid filter offset"))?,
    };
    let limit = match input.get("limit") {
        None => 100,
        Some(v) => v
            .as_u64()
            .filter(|v| *v > 0 && *v <= 100)
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid filter page limit"))?,
    };
    let version = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing filter version"))?;
    let index: Value =
        serde_json::from_slice(&std::fs::read(root.join("index.json")).map_err(err)?)
            .map_err(err)?;
    if !index["sheets"]
        .as_array()
        .is_some_and(|sheets| sheets.iter().any(|s| s["path"] == sheet))
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown filter worksheet",
        ));
    }
    let db = database(root)?;
    db.execute_batch("CREATE TEMP TABLE filter_values(k TEXT PRIMARY KEY,value TEXT NOT NULL,label TEXT NOT NULL,n INTEGER NOT NULL,first_row INTEGER NOT NULL) WITHOUT ROWID;").map_err(err)?;
    let mut protected = json!({"errors":0,"unresolved":0,"special":0,"unsafeNumbers":0});
    let mut after = Value::Null;
    let mut next = first.0;
    let range = format!(
        "{}:{}",
        crate::xml::address(first.0, col),
        crate::xml::address(last.0, col)
    );
    loop {
        crate::cancel::check()?;
        let page = read_db(
            &db,
            &json!({"sheet":sheet,"version":version,"range":range,"after":after,"limit":2048}),
        )?;
        let cells = page["cells"]
            .as_array()
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid filter cell page"))?;
        crate::diagnostics::count("filter_values.pages", 1);
        crate::diagnostics::maximum("filter_values.page_cells", cells.len() as u64);
        for pair in cells {
            crate::cancel::check()?;
            let row = crate::xml::position(pair[0].as_str().ok_or_else(|| {
                crate::coded("internal-invalid-field", "Invalid filter cell address")
            })?)?
            .0;
            if row < next || row > last.0 {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid filter cell order",
                ));
            }
            if row > next {
                filter_group(&db, &Value::Null, next, u64::from(row - next))?;
                crate::diagnostics::count("filter_values.blank_positions", u64::from(row - next));
            }
            match filter_literal(&pair[1]) {
                Ok(value) => filter_group(&db, &value, row, 1)?,
                Err(kind) => protected[kind] = json!(protected[kind].as_u64().unwrap() + 1),
            }
            crate::diagnostics::count("filter_values.stored_cells", 1);
            next = row + 1;
        }
        after = page["nextCursor"].clone();
        if after.is_null() {
            break;
        }
    }
    if next <= last.0 {
        let count = u64::from(last.0 - next) + 1;
        filter_group(&db, &Value::Null, next, count)?;
        crate::diagnostics::count("filter_values.blank_positions", count);
    }
    crate::cancel::check()?;
    // instr matches a literal search string (including '%' and '_'), and the
    // stored lowercase label provides Unicode-aware case-insensitive search.
    let total: u64 = db
        .query_row(
            "SELECT count(*) FROM filter_values WHERE instr(label,?1)>0",
            [&query],
            |row| row.get(0),
        )
        .map_err(err)?;
    let mut statement=db.prepare_cached("SELECT k,value,n FROM filter_values WHERE instr(label,?1)>0 ORDER BY first_row,k LIMIT ?2 OFFSET ?3").map_err(err)?;
    let mut rows = statement
        .query(params![query, limit, offset])
        .map_err(err)?;
    let mut values = Vec::new();
    while let Some(row) = rows.next().map_err(err)? {
        crate::cancel::check()?;
        let value: Value =
            serde_json::from_str(&row.get::<_, String>(1).map_err(err)?).map_err(err)?;
        values.push(json!({"key":row.get::<_,String>(0).map_err(err)?,"value":value,"count":row.get::<_,u64>(2).map_err(err)?}));
    }
    crate::diagnostics::count("filter_values.distinct_values", total);
    crate::diagnostics::count("filter_values.returned_values", values.len() as u64);
    let end = offset + values.len() as u64;
    let result = json!({"values":values,"total":total,"nextOffset":if end<total {Some(end)} else {None},"protected":protected});
    if serde_json::to_vec(&result).map_err(err)?.len() > 8 * 1024 * 1024 {
        return Err(crate::coded(
            "query-limit",
            "Filter value page exceeds the transport byte bound",
        ));
    }
    Ok(result)
}

/// Use the same qualified scalar extraction as filter apply. Error cells must
/// never become text choices, and missing calculated values must not be blank.
fn filter_literal(cell: &Value) -> std::result::Result<Value, &'static str> {
    if cell.get("formulaType").is_some() || cell["scenarioResult"] == true {
        return Err("special");
    }
    if cell.get("formula").is_some() {
        match cell["calculation"]["status"].as_str() {
            Some("calculated") if cell["calculation"].get("value").is_some() => {}
            Some("error" | "circular") => return Err("errors"),
            _ => return Err("unresolved"),
        }
    } else if cell["dataType"] == "error" {
        return Err("errors");
    }
    let value = crate::bulk::literal(cell, true).map_err(|_| "unsafeNumbers")?;
    if value.is_object() || value.as_f64().is_some_and(|n| !n.is_finite()) {
        return Err("unsafeNumbers");
    }
    if value.is_array() {
        return Err("special");
    }
    Ok(value)
}

fn filter_group(db: &rusqlite::Connection, value: &Value, row: u32, count: u64) -> Result<()> {
    let (key, label, wire) = match value {
        Value::Null => ("null".into(), "(blank)".into(), Value::Null),
        Value::String(text) => (format!("s:{text}"), text.to_lowercase(), value.clone()),
        Value::Bool(b) => (format!("b:{b}"), b.to_string(), value.clone()),
        Value::Number(number) => {
            let number = number
                .as_f64()
                .filter(|n| n.is_finite())
                .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid filter number"))?;
            // Exact binary64 identities retain adjacent representable values
            // and signed zero; apply currently distinguishes +0 and -0.
            (
                format!("n:{:016x}", number.to_bits()),
                number.to_string(),
                if number == 0.0 && number.is_sign_negative() {
                    json!({"$workbookNumber":"-0"})
                } else {
                    value.clone()
                },
            )
        }
        _ => {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid filter scalar",
            ));
        }
    };
    db.prepare_cached("INSERT INTO filter_values(k,value,label,n,first_row) VALUES(?1,?2,?3,?4,?5) ON CONFLICT(k) DO UPDATE SET n=n+excluded.n,first_row=min(first_row,excluded.first_row)").map_err(err)?.execute(params![key,wire.to_string(),label,count,row]).map_err(err)?;
    Ok(())
}
