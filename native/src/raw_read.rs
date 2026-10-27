//! Raw MCP read contract over a selected-sheet SQLite index. No calculation or cell graph.
use crate::{
    Result, database, err, field,
    xml::{address, spreadsheet_text},
};
use rusqlite::{OptionalExtension, params};
use serde_json::{Value, json};
use std::path::Path;
fn ordered(fields: &[(&str, Value)]) -> String {
    ordered_format(fields, None)
}
fn ordered_format(fields: &[(&str, Value)], format: Option<&str>) -> String {
    format!(
        "{{{}}}",
        fields
            .iter()
            .map(|(key, value)| format!(
                "{}:{}",
                serde_json::to_string(key).unwrap(),
                if let Some(format) = format
                    && *key == "numberFormat"
                {
                    format.to_owned()
                } else if let Value::Number(n) = value {
                    crate::display::number(n.as_f64().unwrap())
                } else {
                    value.to_string()
                }
            ))
            .collect::<Vec<_>>()
            .join(",")
    )
}
struct Record {
    kind: String,
    raw: Option<String>,
    inline: String,
    formula: Option<String>,
    attrs: Option<String>,
    style: usize,
    explicit: bool,
}
fn item(
    record: Option<Record>,
    r: u32,
    c: u32,
    input: &Value,
    meta: &Value,
    formats: &[String],
) -> Result<String> {
    let style = record.as_ref().map(|r| r.style);
    let reference = address(r, c);
    let hidden = meta["hiddenRows"].as_array().is_some_and(|a| {
        a.binary_search_by_key(&u64::from(r), |v| v.as_u64().unwrap_or(u64::MAX))
            .is_ok()
    });
    let col_hidden = meta["hiddenColumns"].as_array().is_some_and(|a| {
        a.iter().any(|v| {
            v[0].as_u64().unwrap_or(u64::MAX) <= u64::from(c + 1)
                && v[1].as_u64().unwrap_or(0) >= u64::from(c + 1)
        })
    });
    let mut fields = vec![
        ("kind", json!("cell")),
        ("sheet", input["name"].clone()),
        ("address", json!(reference)),
        ("rowHidden", json!(hidden)),
        ("columnHidden", json!(col_hidden)),
    ];
    let mut attributes_text = String::new();
    let (kind, value, raw, formula, cache, format, expected, expected_formula, storage) =
        if let Some(record) = record {
            let raw = record.raw.as_deref();
            let f = record.formula.is_some();
            let missing = raw.is_none() || (raw == Some("") && record.kind != "str");
            let kind: String;
            let mut value = json!(raw);
            if f && missing {
                kind = match record.kind.as_str() {
                    "n" => "number",
                    "b" => "boolean",
                    "str" => "string",
                    "e" => "error",
                    "d" => "iso-date",
                    s => s,
                }
                .into();
                value = Value::Null;
            } else {
                match record.kind.as_str() {
                    "s" => {
                        kind = "string".into();
                        value = json!(record.inline)
                    }
                    "inlineStr" => {
                        kind = "string".into();
                        value = json!(record.inline)
                    }
                    "b" => {
                        if !matches!(raw, Some("0" | "1")) {
                            return Err(crate::coded(
                                "internal-invalid-field",
                                "Invalid boolean cell",
                            ));
                        }
                        kind = "boolean".into();
                        value = json!(raw == Some("1"));
                    }
                    "n" => {
                        kind = if raw.is_none() || raw == Some("") {
                            value = Value::Null;
                            "blank"
                        } else {
                            "number"
                        }
                        .into();
                    }
                    "str" => {
                        kind = "string".into();
                        value = raw
                            .map(spreadsheet_text)
                            .transpose()?
                            .map(|value| Value::String(value.into_owned()))
                            .unwrap_or(Value::Null);
                    }
                    "e" => kind = "error".into(),
                    "d" => kind = "iso-date".into(),
                    _ => kind = format!("unsupported:{}", record.kind),
                }
            }
            if record.explicit
                && input["formats"]
                    .as_array()
                    .is_none_or(|a| record.style >= a.len())
            {
                return Err(crate::coded(
                    "package-unsupported",
                    "Cell style index is unavailable",
                ));
            }
            let format = input["formats"]
                .get(record.style)
                .cloned()
                .unwrap_or(json!({"id":"0","code":null,"source":"builtin-id"}));
            attributes_text = record.attrs.clone().unwrap_or_else(|| "{}".into());
            let attrs: Value = record
                .attrs
                .as_deref()
                .map(serde_json::from_str)
                .transpose()
                .map_err(err)?
                .unwrap_or(json!({}));
            let expression = record.formula.filter(|f| !f.is_empty());
            let formula = if f {
                json!({"expression":expression,"attributes":attrs,"interpretation":if expression.is_some(){"source-expression"}else{"shared-or-special-formula-unresolved"}})
            } else {
                Value::Null
            };
            let cache = if f {
                json!({"value":value,"raw":raw,"type":kind,"freshness":if missing{"missing"}else{"unverified"}})
            } else {
                Value::Null
            };
            let expected = if let Some(raw) = raw
                && kind == "number"
            {
                raw.parse::<f64>()
                    .ok()
                    .filter(|n| n.is_finite())
                    .map_or(Value::Null, |n| json!(n))
            } else {
                value.clone()
            };
            let expected_formula = if f {
                json!(format!("={}", expression.as_deref().unwrap_or("null")))
            } else {
                Value::Null
            };
            (
                kind,
                if f { Value::Null } else { value },
                json!(raw),
                formula,
                cache,
                format,
                expected,
                expected_formula,
                "present",
            )
        } else {
            (
                "blank".into(),
                Value::Null,
                Value::Null,
                Value::Null,
                Value::Null,
                Value::Null,
                Value::Null,
                Value::Null,
                "absent",
            )
        };
    fields.extend([
        ("type", json!(kind)),
        ("value", value),
        ("rawValue", raw.clone()),
        ("formula", formula.clone()),
        ("cachedValue", cache.clone()),
        ("numberFormat", format.clone()),
    ]);
    if storage == "present" {
        fields.push(("displayedText", Value::Null));
    }
    fields.extend([
        ("storage", json!(storage)),
        ("exactValue", raw),
        ("expectedValue", expected),
        (
            "guardEncoding",
            json!(if kind == "number" {
                "IEEE-754 number for the editing engine; exactValue retains the stored text"
            } else {
                "typed value"
            }),
        ),
        ("expectedFormula", expected_formula),
    ]);
    let format_text = style.map(|index| {
        formats
            .get(index)
            .map(String::as_str)
            .unwrap_or(r#"{"id":"0","code":null,"source":"builtin-id"}"#)
    });
    let mut text = ordered_format(&fields, format_text);
    // Keep the documented JS field order in the search text, including nested records.
    if !formula.is_null() {
        let from = formula.to_string();
        let to = format!(
            "{{\"expression\":{},\"attributes\":{},\"interpretation\":{}}}",
            formula["expression"], attributes_text, formula["interpretation"]
        );
        text = text.replace(&from, &to);
    }
    if !cache.is_null() {
        let from = cache.to_string();
        let to = ordered(&[
            ("value", cache["value"].clone()),
            ("raw", cache["raw"].clone()),
            ("type", cache["type"].clone()),
            ("freshness", cache["freshness"].clone()),
        ]);
        text = text.replace(&from, &to);
    }
    Ok(text)
}
pub fn query(root: &Path, input: &Value) -> Result<Value> {
    let (first, last) = crate::bulk::bounds(field(input, "range")?)?;
    let columns = u64::from(last.1 - first.1 + 1);
    let size = u64::from(last.0 - first.0 + 1) * columns;
    let offset = input["offset"].as_u64().unwrap_or(0);
    let limit = input["limit"]
        .as_u64()
        .filter(|n| *n > 0 && *n <= 1000)
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid raw read limit"))?;
    let needle = input["query"].as_str().unwrap_or("").to_lowercase();
    if !needle.is_empty() && size > 50_000_000 {
        return Err(crate::coded(
            "package-limit",
            "Raw search work budget exceeds 50000000 positions; narrow the range",
        ));
    }
    let db = database(root)?;
    let index: Value =
        serde_json::from_slice(&std::fs::read(root.join("index.json")).map_err(err)?)
            .map_err(err)?;
    let sheet = field(input, "sheet")?;
    let meta = index["sheets"]
        .as_array()
        .and_then(|a| a.iter().find(|s| s["path"] == sheet))
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing raw sheet"))?;
    let mut statement=db.prepare_cached("SELECT c.kind,c.raw,c.inline,c.formula,c.formula_attrs,c.style,c.style_explicit,s.value FROM cells c LEFT JOIN strings s ON c.kind='s' AND s.id=CAST(c.raw AS INTEGER) WHERE c.sheet=?1 AND c.r=?2 AND c.c=?3").map_err(err)?;
    let read_record = |row: &rusqlite::Row<'_>| -> rusqlite::Result<Record> {
        let kind: String = row.get(0)?;
        let raw: Option<String> = row.get(1)?;
        let inline = if kind == "s" {
            row.get::<_, Option<String>>(7)?
        } else {
            Some(row.get::<_, String>(2)?)
        };
        if kind == "s"
            && (raw
                .as_deref()
                .is_none_or(|s| s.is_empty() || !s.bytes().all(|b| b.is_ascii_digit()))
                || inline.is_none())
        {
            return Err(rusqlite::Error::InvalidParameterName(
                "Invalid shared-string index".into(),
            ));
        }
        Ok(Record {
            kind,
            raw,
            inline: inline.unwrap_or_default(),
            formula: row.get(3)?,
            attrs: row.get(4)?,
            style: row.get(5)?,
            explicit: row.get(6)?,
        })
    };
    // Search streams the occupied rows once; absent positions still participate in
    // the raw JSON contract without a separate SQLite point query per coordinate.
    let mut scan = db.prepare_cached("SELECT c.kind,c.raw,c.inline,c.formula,c.formula_attrs,c.style,c.style_explicit,s.value,c.r,c.c FROM cells c LEFT JOIN strings s ON c.kind='s' AND s.id=CAST(c.raw AS INTEGER) WHERE c.sheet=?1 AND c.r BETWEEN ?2 AND ?3 AND c.c BETWEEN ?4 AND ?5 ORDER BY c.r,c.c").map_err(err)?;
    let mut rows = scan
        .query_map(params![sheet, first.0, last.0, first.1, last.1], |row| {
            Ok((
                row.get::<_, u32>(8)?,
                row.get::<_, u32>(9)?,
                read_record(row)?,
            ))
        })
        .map_err(err)?;
    let mut next = if needle.is_empty() {
        None
    } else {
        rows.next().transpose().map_err(err)?
    };
    let formats: Vec<String> = input["formats"]
        .as_array()
        .map(|formats| {
            formats
                .iter()
                .map(|format| {
                    ordered(&[
                        ("id", format["id"].clone()),
                        ("code", format["code"].clone()),
                        ("source", format["source"].clone()),
                    ])
                })
                .collect()
        })
        .unwrap_or_default();
    let mut items = Vec::new();
    let mut matched = 0u64;
    let start = if needle.is_empty() {
        offset.min(size)
    } else {
        0
    };
    let end = if needle.is_empty() {
        offset.saturating_add(limit).min(size)
    } else {
        size
    };
    for position in start..end {
        crate::cancel::check()?;
        let r = first.0 + (position / columns) as u32;
        let c = first.1 + (position % columns) as u32;
        let record = if needle.is_empty() {
            statement
                .query_row(params![sheet, r, c], read_record)
                .optional()
                .map_err(err)?
        } else if next
            .as_ref()
            .is_some_and(|(nr, nc, _)| (*nr, *nc) == (r, c))
        {
            let record = next.take().map(|(_, _, record)| record);
            next = rows.next().transpose().map_err(err)?;
            record
        } else {
            None
        };
        let text = item(record, r, c, input, meta, &formats)?;
        if needle.is_empty() {
            items.push(serde_json::from_str::<Value>(&text).map_err(err)?)
        } else if text.to_lowercase().contains(&needle) {
            if matched >= offset && items.len() < (limit as usize) {
                items.push(serde_json::from_str::<Value>(&text).map_err(err)?)
            }
            matched += 1;
        }
    }
    let total = if needle.is_empty() { size } else { matched };
    Ok(
        json!({"items":items,"total":total,"offset":offset,"nextOffset":if offset.saturating_add(limit)<total{Some(offset+limit)}else{None}}),
    )
}
