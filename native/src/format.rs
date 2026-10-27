//! Compact range formatting. Only one tile is decoded at a time; acceptance is
//! one SQLite transaction, including batches containing overlapping ranges.
use crate::range::{STYLE_FIELDS, inherited_edit, merge_edit, read_db, store_calculated_change};
use crate::xml::{address, position};
use crate::{Result, database, err, field, local};
use serde_json::{Value, json};
use std::{fs::File, io::Read, path::Path};

fn style_count(root: &Path) -> Result<u64> {
    let mut zip = zip::ZipArchive::new(File::open(local(root, "metadata.xlsx")?).map_err(err)?)
        .map_err(err)?;
    let mut bytes = Vec::new();
    zip.by_name("xl/styles.xml")
        .map_err(err)?
        .read_to_end(&mut bytes)
        .map_err(err)?;
    let mut reader = quick_xml::Reader::from_reader(bytes.as_slice());
    let (mut inside, mut count) = (false, 0);
    loop {
        crate::cancel::check()?;
        use quick_xml::events::Event;
        match reader.read_event().map_err(err)? {
            Event::Start(e) if e.local_name().as_ref() == b"cellXfs" => inside = true,
            Event::End(e) if e.local_name().as_ref() == b"cellXfs" => inside = false,
            Event::Start(e) | Event::Empty(e) if inside && e.local_name().as_ref() == b"xf" => {
                count += 1
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(count)
}

fn validate(format: &Value, styles: u64) -> Result<()> {
    let map = format
        .as_object()
        .filter(|m| !m.is_empty())
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing range format"))?;
    for (key, value) in map {
        let valid = match key.as_str() {
            "bold" | "italic" | "underline" | "wrapText" => value.is_boolean(),
            "styleId" => value.as_u64().is_some_and(|id| id < styles),
            "fontSize" => value.as_f64().is_some_and(|n| (6.0..=72.0).contains(&n)),
            "fontFamily" => value.as_str().is_some_and(|s| {
                !s.trim().is_empty()
                    && s.encode_utf16().count() <= 100
                    && !s.chars().any(|c| {
                        c == '\u{fffe}'
                            || c == '\u{ffff}'
                            || c < ' ' && !['\t', '\r', '\n'].contains(&c)
                    })
            }),
            "fontColor" | "fillColor" => value
                .as_str()
                .is_some_and(|s| s.len() == 6 && s.bytes().all(|b| b.is_ascii_hexdigit())),
            "horizontalAlign" => value
                .as_str()
                .is_some_and(|s| ["left", "center", "right"].contains(&s)),
            "verticalAlign" => value
                .as_str()
                .is_some_and(|s| ["top", "center", "bottom"].contains(&s)),
            "border" => value.as_str().is_some_and(|s| ["none", "all"].contains(&s)),
            "numberFormat" => value.as_str().is_some_and(|s| {
                [
                    "general",
                    "number",
                    "percent",
                    "date",
                    "text",
                    "currency",
                    "accounting",
                    "thousands",
                    "decimal0",
                    "decimal1",
                    "decimal2",
                ]
                .contains(&s)
            }),
            _ => false,
        };
        if !valid {
            return Err(crate::coded(
                "internal-invalid-field",
                format!("Invalid range format: {key}"),
            ));
        }
    }
    Ok(())
}

fn display_format(name: &str) -> Value {
    match name {
        "number" | "decimal2" => json!(2),
        "percent" => json!(10),
        "date" => json!(14),
        "text" => json!(49),
        "currency" => json!("$#,##0.00"),
        "accounting" => json!(44),
        "thousands" => json!(3),
        "decimal0" => json!(1),
        "decimal1" => json!("0.0"),
        _ => json!(0),
    }
}

pub fn apply(root: &Path, input: &Value) -> Result<Value> {
    let commands = input["commands"]
        .as_array()
        .filter(|c| !c.is_empty())
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing format commands"))?;
    let index: Value =
        serde_json::from_slice(&std::fs::read(local(root, "index.json")?).map_err(err)?)
            .map_err(err)?;
    let styles = style_count(root)?;
    let mut db = database(root)?;
    let tx = db.transaction().map_err(err)?;
    let mut parent = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing parent version"))?;
    let mut changed = 0u64;
    if tx
        .query_row("SELECT count(*) FROM versions WHERE id=?1", [parent], |r| {
            r.get::<_, u32>(0)
        })
        .map_err(err)?
        != 1
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown workbook version",
        ));
    }
    for command in commands {
        let sheet = field(command, "sheet")?;
        if !index["sheets"]
            .as_array()
            .unwrap()
            .iter()
            .any(|s| s["path"] == sheet)
        {
            return Err(crate::coded(
                "internal-unknown-reference",
                "Unknown worksheet",
            ));
        }
        validate(&command["format"], styles)?;
        let format = &command["format"];
        let mut bounds = field(command, "range")?.split(':');
        let first = position(
            bounds
                .next()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing range"))?,
        )?;
        let last = bounds.next().map(position).transpose()?.unwrap_or(first);
        if bounds.next().is_some() || first.0 > last.0 || first.1 > last.1 {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid format range",
            ));
        }
        if crate::calculation_policy::preserving(&tx, parent)? {
            crate::preservation::region(root, &tx, sheet, first, last, false)?;
        }
        tx.execute("INSERT INTO versions(parent) VALUES(?1)", [parent])
            .map_err(err)?;
        let version = tx.last_insert_rowid();
        for row in (first.0..=last.0).step_by(32) {
            for col in (first.1..=last.1).step_by(32) {
                let end_row = last.0.min(row + 31);
                let end_col = last.1.min(col + 31);
                let page = read_db(
                    &tx,
                    &json!({"version":parent,"sheet":sheet,"range":format!("{}:{}",address(row,col),address(end_row,end_col)),"limit":2048}),
                )?;
                let mut cells: std::collections::BTreeMap<String, Value> = page["cells"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .map(|c| (c[0].as_str().unwrap().into(), c[1].clone()))
                    .collect();
                for r in row..=end_row {
                    for c in col..=end_col {
                        let mut cell = cells.remove(&address(r, c)).unwrap_or_else(
                            || json!({"value":null,"dataType":"blank","styleId":0}),
                        );
                        let mut edit = inherited_edit(&tx, parent, sheet, r, c)?;
                        merge_edit(&mut edit, format)?;
                        if let Some(style) = format.get("styleId") {
                            for key in STYLE_FIELDS {
                                cell.as_object_mut().unwrap().remove(*key);
                            }
                            cell.as_object_mut().unwrap().remove("format");
                            cell["styleId"] = style.clone();
                        }
                        for key in STYLE_FIELDS {
                            if let Some(value) = format.get(*key) {
                                if *key == "numberFormat" {
                                    cell["format"] = display_format(value.as_str().unwrap());
                                } else {
                                    cell[*key] = value.clone();
                                }
                            }
                        }
                        // This path takes no caller-authored cells: only the
                        // validated style whitelist modifies the parent page.
                        // Retain complete calculated/spill fields unchanged;
                        // ordinary cell patches still cannot manufacture them.
                        if cell.get("$sourceFields").is_some() {
                            return Err(crate::coded(
                                "style-unsupported",
                                "Reserved calculation storage field",
                            ));
                        }
                        store_calculated_change(&tx, version, sheet, r, c, &cell, &edit)?;
                        changed += 1;
                    }
                }
            }
        }
        parent = version as u64;
    }
    tx.commit().map_err(err)?;
    Ok(json!({"version":parent,"changedCells":changed}))
}
