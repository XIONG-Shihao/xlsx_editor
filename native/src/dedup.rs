//! Duplicate-record analysis over an immutable accepted version. Only bounded
//! pages enter memory; keys and the stable first-row mapping live in TEMP SQLite.
use crate::bulk::{bounds, meaning};
use crate::xml::{address, position};
use crate::{Result, database, err, field};
use rusqlite::{Connection, OptionalExtension, params};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use std::path::Path;

mod write;
pub fn apply(root: &Path, input: &Value) -> Result<Value> {
    write::apply(root, input)
}

pub(crate) struct Config {
    pub first: (u32, u32),
    pub last: (u32, u32),
    pub body: u32,
    pub columns: Vec<u32>,
    pub version: u64,
}

pub(crate) fn config(input: &Value) -> Result<Config> {
    let (first, last) = bounds(field(input, "range")?)?;
    let positions = u64::from(last.0 - first.0 + 1) * u64::from(last.1 - first.1 + 1);
    if positions > 1_000_000 {
        return Err(crate::coded(
            "duplicates-duplicate",
            "Duplicate removal is limited to 1,000,000 selected positions",
        ));
    }
    let header = input["hasHeaders"].as_bool().ok_or_else(|| {
        crate::coded(
            "duplicates-bad-input",
            "Choose whether the range has headers",
        )
    })?;
    let body = first.0 + u32::from(header);
    if body > last.0 {
        return Err(crate::coded(
            "duplicates-duplicate",
            "Duplicate removal needs at least one data row",
        ));
    }
    let raw = input["columns"]
        .as_array()
        .filter(|v| !v.is_empty() && v.len() <= 16)
        .ok_or_else(|| crate::coded("duplicates-duplicate", "Choose 1–16 duplicate key columns"))?;
    let mut columns = Vec::with_capacity(raw.len());
    for item in raw {
        let label = item.as_str().ok_or_else(|| {
            crate::coded("internal-invalid-field", "Invalid duplicate key column")
        })?;
        if label.is_empty() || label.len() > 3 || !label.bytes().all(|b| b.is_ascii_uppercase()) {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid duplicate key column",
            ));
        }
        let column = position(&format!("{label}1"))?.1;
        if column < first.1 || column > last.1 || columns.contains(&column) {
            return Err(crate::coded(
                "duplicates-duplicate",
                "Duplicate key columns must be distinct and inside the range",
            ));
        }
        columns.push(column);
    }
    Ok(Config {
        first,
        last,
        body,
        columns,
        version: input["version"].as_u64().ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing duplicate preview version",
            )
        })?,
    })
}

/// Exact typed values: errors and text differ; blank and empty text differ.
/// Decimal spellings compare by mathematical value without rounding raw source.
/// Styling is moved with the kept record, and is not part of a duplicate key.
fn key(cell: &Value) -> Result<Value> {
    if cell["formulaType"].is_string() || cell["scenarioResult"] == true {
        return Err(crate::coded(
            "duplicates-duplicate",
            "Duplicate keys with special formulas or scenario outputs are protected",
        ));
    }
    let derived = cell["formula"].is_string() || cell["calculation"]["pending"] == true;
    let value = if derived {
        if cell["calculation"]["pending"] == true {
            return Err(crate::coded(
                "duplicates-bad-input",
                "Calculate pending duplicate keys before previewing",
            ));
        }
        match cell["calculation"]["status"].as_str() {
            Some("calculated") => &cell["calculation"]["value"],
            Some("error") => {
                let error = cell["calculation"]["error"].as_str().ok_or_else(|| {
                    crate::coded(
                        "internal-invalid-field",
                        "Invalid calculated duplicate error",
                    )
                })?;
                return Ok(json!(["error", error]));
            }
            _ => {
                return Err(crate::coded(
                    "duplicates-unsupported",
                    "Calculate unresolved duplicate keys before previewing",
                ));
            }
        }
    } else {
        if cell["dataType"] == "error" {
            return Ok(json!([
                "error",
                cell["value"].as_str().ok_or_else(|| crate::coded(
                    "internal-invalid-field",
                    "Invalid duplicate error"
                ))?
            ]));
        }
        &cell["value"]
    };
    if value.is_null() {
        return Ok(json!(["blank"]));
    }
    if let Some(text) = value.as_str() {
        return Ok(json!(["text", text]));
    }
    if let Some(boolean) = value.as_bool() {
        return Ok(json!(["boolean", boolean]));
    }
    if let Some(number) = value.as_f64() {
        if !number.is_finite() {
            return Err(crate::coded(
                "duplicates-protected",
                "Nonfinite duplicate keys are protected",
            ));
        }
        let canonical = number.to_string();
        let raw = if derived {
            canonical.as_str()
        } else {
            cell["rawValue"].as_str().unwrap_or(&canonical)
        };
        let (negative, digits, power) = meaning(raw).ok_or_else(|| {
            crate::coded("internal-invalid-field", "Invalid numeric duplicate key")
        })?;
        return Ok(json!(["number", negative, digits, power]));
    }
    Err(crate::coded(
        "duplicates-duplicate",
        "Unsupported duplicate key type",
    ))
}

fn store(
    db: &Connection,
    digest: &mut Sha256,
    row: u32,
    keys: &[Value],
    body: u32,
    kept: &mut u64,
) -> Result<()> {
    crate::cancel::check()?;
    let encoded = serde_json::to_string(keys).map_err(err)?;
    if encoded.len() > 1_048_576 {
        return Err(crate::coded(
            "duplicates-duplicate",
            "Duplicate key row exceeds 1 MiB",
        ));
    }
    digest.update(row.to_le_bytes());
    digest.update((encoded.len() as u64).to_le_bytes());
    digest.update(encoded.as_bytes());
    let first: Option<u32> = db
        .prepare_cached("SELECT first_row FROM duplicate_first WHERE k=?1")
        .map_err(err)?
        .query_row([&encoded], |r| r.get(0))
        .optional()
        .map_err(err)?;
    let first = match first {
        Some(first) => first,
        None => {
            db.prepare_cached("INSERT INTO duplicate_first VALUES(?1,?2)")
                .map_err(err)?
                .execute(params![encoded, row])
                .map_err(err)?;
            row
        }
    };
    let destination = if first == row {
        let destination = u64::from(body) + *kept;
        *kept += 1;
        Some(destination)
    } else {
        None
    };
    db.prepare_cached("INSERT INTO duplicate_rows VALUES(?1,?2,?3)")
        .map_err(err)?
        .execute(params![row, first, destination])
        .map_err(err)?;
    Ok(())
}

/// Leave the first/source/destination mapping in this connection for the
/// independent writer to recompute and validate at acceptance.
pub(crate) fn prepare(root: &Path, db: &Connection, input: &Value) -> Result<Value> {
    let config = config(input)?;
    let sheet = field(input, "sheet")?;
    let index: Value =
        serde_json::from_slice(&std::fs::read(root.join("index.json")).map_err(err)?)
            .map_err(err)?;
    if !index["sheets"]
        .as_array()
        .is_some_and(|sheets| sheets.iter().any(|s| s["path"] == sheet))
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown duplicate worksheet",
        ));
    }
    let source = crate::hash(&root.join("source.xlsx"))?;
    if index["revision"].as_str() != Some(source.as_str()) {
        return Err(crate::coded(
            "duplicates-duplicate",
            "Duplicate preview source changed since indexing",
        ));
    }
    db.execute_batch("CREATE TEMP TABLE duplicate_first(k TEXT PRIMARY KEY,first_row INTEGER NOT NULL) WITHOUT ROWID; CREATE TEMP TABLE duplicate_rows(source INTEGER PRIMARY KEY,first_row INTEGER NOT NULL,destination INTEGER UNIQUE);").map_err(err)?;
    let canonical = json!({"source":source,"version":config.version,"sheet":sheet,
        "first":config.first,"last":config.last,"body":config.body,"columns":config.columns,
        "comparison":"exact-typed-values-v1"});
    let mut digest = Sha256::new();
    digest.update(serde_json::to_vec(&canonical).map_err(err)?);
    let blank = vec![json!(["blank"]); config.columns.len()];
    let mut keys = blank.clone();
    let mut row = config.body;
    let mut kept = 0;
    let mut after = Value::Null;
    let range = format!(
        "{}:{}",
        address(config.body, config.first.1),
        address(config.last.0, config.last.1)
    );
    loop {
        crate::cancel::check()?;
        let page = crate::range::read_public_db(
            db,
            &json!({"sheet":sheet,"version":config.version,
            "range":range,"limit":2048,"after":after}),
        )?;
        for pair in page["cells"]
            .as_array()
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid duplicate cell page"))?
        {
            let (r, c) = position(pair[0].as_str().ok_or_else(|| {
                crate::coded("internal-invalid-field", "Invalid duplicate cell address")
            })?)?;
            if r < row || r > config.last.0 {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid duplicate cell order",
                ));
            }
            while row < r {
                store(db, &mut digest, row, &keys, config.body, &mut kept)?;
                row += 1;
                keys.clone_from(&blank);
            }
            if let Some(i) = config.columns.iter().position(|column| *column == c) {
                keys[i] = key(&pair[1])?;
            }
        }
        after = page["nextCursor"].clone();
        if after.is_null() {
            break;
        }
    }
    while row <= config.last.0 {
        store(db, &mut digest, row, &keys, config.body, &mut kept)?;
        row += 1;
        keys.clone_from(&blank);
    }
    let total = u64::from(config.last.0 - config.body + 1);
    Ok(
        json!({"fingerprint":format!("{:x}",digest.finalize()),"totalRows":total,
        "keptRows":kept,"removedRows":total-kept,"comparison":"exact-typed-values-v1"}),
    )
}

pub fn preview(root: &Path, input: &Value) -> Result<Value> {
    let o = input.as_object().ok_or_else(|| {
        crate::coded("internal-invalid-field", "Invalid duplicate preview fields")
    })?;
    if [
        "action",
        "operation",
        "root",
        "version",
        "sheet",
        "range",
        "hasHeaders",
        "columns",
    ]
    .iter()
    .any(|k| !o.contains_key(*k))
        || o.keys().any(|k| {
            ![
                "action",
                "operation",
                "root",
                "version",
                "sheet",
                "range",
                "hasHeaders",
                "columns",
                "offset",
                "limit",
            ]
            .contains(&k.as_str())
        })
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid duplicate preview fields",
        ));
    }

    let offset = match input.get("offset") {
        None => 0,
        Some(value) => value.as_u64().filter(|v| *v <= 1_000_000).ok_or_else(|| {
            crate::coded("internal-invalid-field", "Invalid duplicate preview offset")
        })?,
    };
    let limit = match input.get("limit") {
        None => 100,
        Some(value) => value
            .as_u64()
            .filter(|v| *v > 0 && *v <= 100)
            .ok_or_else(|| {
                crate::coded(
                    "internal-invalid-field",
                    "Invalid duplicate preview page limit",
                )
            })?,
    };
    let db = database(root)?;
    let mut result = prepare(root, &db, input)?;
    let mut statement = db.prepare_cached("SELECT source,first_row FROM duplicate_rows WHERE source!=first_row ORDER BY source LIMIT ?1 OFFSET ?2").map_err(err)?;
    let mut rows = statement.query(params![limit, offset]).map_err(err)?;
    let mut samples = Vec::new();
    while let Some(row) = rows.next().map_err(err)? {
        // UI row labels are one-based; SQLite/source coordinates are zero-based.
        samples.push(json!({"row":row.get::<_,u32>(0).map_err(err)?+1,
            "keptRow":row.get::<_,u32>(1).map_err(err)?+1}));
    }
    let next = offset + samples.len() as u64;
    result["samples"] = json!(samples);
    result["offset"] = json!(offset);
    result["nextOffset"] = if next < result["removedRows"].as_u64().unwrap() {
        json!(next)
    } else {
        Value::Null
    };
    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn exact_source_decimal_identity_without_binary_rounding() {
        let number = |raw: &str| json!({"value":raw.parse::<f64>().unwrap(),"dataType":"number","rawValue":raw});
        assert_eq!(key(&number("1.00")).unwrap(), key(&number("1e0")).unwrap());
        assert_eq!(key(&number("-0")).unwrap(), key(&number("0.0")).unwrap());
        assert_ne!(
            key(&number("9007199254740992")).unwrap(),
            key(&number("9007199254740993")).unwrap()
        );
        assert_ne!(
            key(&number("0.1234567890123456789")).unwrap(),
            key(&number("0.1234567890123456788")).unwrap()
        );
    }

    #[test]
    fn keys_keep_text_boolean_blank_empty_and_error_distinctions() {
        let inputs = [
            json!({"value":null}),
            json!({"value":""}),
            json!({"value":"1"}),
            json!({"value":1,"rawValue":"1"}),
            json!({"value":true}),
            json!({"value":"#N/A"}),
            json!({"value":"#N/A","dataType":"error"}),
            json!({"value":"A"}),
            json!({"value":"a"}),
            json!({"value":"A "}),
        ];
        let mut keys = inputs
            .iter()
            .map(|cell| serde_json::to_string(&key(cell).unwrap()).unwrap())
            .collect::<Vec<_>>();
        keys.sort();
        keys.dedup();
        assert_eq!(keys.len(), inputs.len());
        assert_eq!(
            key(&json!({"formula":"1/0","calculation":{"status":"error","error":"#DIV/0!"}}))
                .unwrap(),
            key(&json!({"value":"#DIV/0!","dataType":"error"})).unwrap()
        );
    }

    #[test]
    fn stored_pending_special_and_unresolved_results_are_never_keys() {
        for cell in [
            json!({"formula":"1","value":1,"calculation":{"status":"stored","pending":true}}),
            json!({"formula":"UNKNOWN()","value":1,"calculation":{"status":"unresolved"}}),
            json!({"formulaType":"array","value":1}),
            json!({"scenarioResult":true,"value":1}),
        ] {
            assert!(key(&cell).is_err());
        }
    }
}
