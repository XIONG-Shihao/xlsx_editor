//! Direct auditing reads use a readonly connection; no dependency plans, selection
//! scratch, calculation versions, cache proofs or result cells are published.
use crate::xml::{address, position};
use crate::{Result, err, field};
use rusqlite::{Connection, OpenFlags};
use serde_json::{Value, json};
use std::{collections::BTreeSet, path::Path};
const MAX_SCAN: u64 = 2048;
fn edge(sheet: &str, cell: &str, range: &Value, kind: &str) -> Value {
    json!({"from":{"sheet":range["sheet"],"range":range["range"]},"to":{"sheet":sheet,"range":cell},"kind":kind,"via":range["via"]})
}
fn contains(range: &str, cell: (u32, u32)) -> Result<bool> {
    Ok(crate::spill::Geometry::from_reference(range)?.contains(cell))
}
fn inspect(
    root: &Path,
    db: &Connection,
    version: u64,
    view: &Value,
    sheet: &Value,
    reference: &str,
    cell: &Value,
) -> Result<(Vec<Value>, Vec<Value>, bool)> {
    let name = field(sheet, "name")?;
    let path = field(sheet, "path")?;
    let at = position(reference)?;
    let mut edges = Vec::new();
    let mut markers = Vec::new();
    let mut unknown = false;
    let owned = if cell.get("spill").is_some() {
        crate::spill::accepted_cell(db, version, path, at, cell)?
    } else {
        false
    };
    if cell["formula"].is_string() {
        let attrs = crate::spill::source_attributes(root, db, path, at)?;
        let opaque = attrs.is_some_and(|a| {
            a.keys().any(|k| {
                !(["r", "s", "t"].contains(&k.as_str())
                    || owned && ["cm", "vm"].contains(&k.as_str()))
            })
        });
        if opaque {
            unknown = true;
            markers.push(json!({"sheet":name,"cell":reference,"kind":"special","reason":"Opaque source cell metadata is outside the auditing profile"}));
        }
        if let Err(reason) = crate::spill::source_content_proof(root, db, path, at, true) {
            if reason.contains("Opaque")
                || reason.contains("Source cell")
                || reason.contains("source cell")
            {
                unknown = true;
                markers.push(json!({"sheet":name,"cell":reference,"kind":"special","reason":reason.chars().take(240).collect::<String>()}));
            } else {
                return Err(crate::coded("formula-refused", reason));
            }
        }
    }
    if cell.get("spill").is_some() && !owned {
        unknown = true;
        markers.push(json!({"sheet":name,"cell":reference,"kind":"special","reason":"Spill marker has no complete accepted ownership proof"}));
    }
    if owned && cell["spill"]["role"] == "follower" {
        edges.push(json!({"from":{"sheet":name,"range":cell["spill"]["master"]},"to":{"sheet":name,"range":reference},"kind":"spill-origin","via":"formula"}));
        return Ok((edges, markers, unknown));
    }
    if let Some(formula) = cell["formula"].as_str() {
        if let Ok(Some(consumer)) = crate::formula::spill_consumer(formula) {
            if cell.get("formulaType").is_some() && !owned || cell["scenarioResult"] == true {
                unknown = true;
                markers.push(json!({"sheet":name,"cell":reference,"kind":"special","reason":"Special scalar source ownership is outside the complete auditing profile"}));
            }
            let anchor = address(consumer.anchor.0, consumer.anchor.1);
            let groups = crate::spill::effective(db, version, path)?;
            let owner = groups.iter().find(|g| g.master == consumer.anchor);
            if let Some(owner) = owner {
                let read = crate::range::read_db(
                    db,
                    &json!({"version":version,"sheet":path,"range":anchor,"limit":1}),
                )?;
                if owner.proof["status"] != "blocked"
                    && crate::spill::accepted_cell(
                        db,
                        version,
                        path,
                        consumer.anchor,
                        &read["cells"][0][1],
                    )?
                {
                    edges.push(edge(
                        name,
                        reference,
                        &json!({"sheet":name,"range":owner.geometry.reference(),"via":"formula"}),
                        "spill-range",
                    ));
                    return Ok((edges, markers, unknown));
                }
            }
            edges.push(edge(
                name,
                reference,
                &json!({"sheet":name,"range":anchor,"via":"formula"}),
                "reference",
            ));
            markers.push(json!({"sheet":name,"cell":reference,"kind":"special","reason":"Spill reference lacks a complete calculated output geometry"}));
            return Ok((edges, markers, true));
        }
    }
    let refs = crate::formula::auditing::references(
        view,
        name,
        reference,
        cell,
        owned && cell["spill"]["role"] == "master",
    )?;
    unknown |= refs.unknown;
    markers.extend(refs.markers);
    for range in refs.ranges {
        edges.push(edge(name, reference, &range, "reference"));
    }
    Ok((edges, markers, unknown))
}
pub(crate) fn command(root: &Path, input: &Value) -> Result<Value> {
    let version = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing trace version"))?;
    let selected_sheet = field(input, "sheet")?;
    let selected_ref = field(input, "cell")?;
    let selected = position(selected_ref)?;
    if address(selected.0, selected.1) != selected_ref {
        return Err(crate::coded(
            "formula-bad-input",
            "Trace target requires a canonical single cell",
        ));
    }
    let direction = field(input, "direction")?;
    if !["precedents", "dependents"].contains(&direction) {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown trace direction",
        ));
    }
    let limit = input["limit"]
        .as_u64()
        .filter(|n| *n >= 1 && *n <= 100)
        .ok_or_else(|| crate::coded("formula-limit", "Trace limit must be1 through100"))?
        as usize;
    let view = &input["view"];
    let sheets = view["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing trace worksheets"))?;
    if sheets.is_empty() || sheets.len() > 1024 {
        return Err(crate::coded(
            "formula-limit",
            "Trace worksheet inventory exceeds its bound",
        ));
    }
    let index: Value =
        serde_json::from_slice(&std::fs::read(root.join("index.json")).map_err(err)?)
            .map_err(err)?;
    let indexed = index["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing indexed trace worksheets"))?
        .iter()
        .map(|s| field(s, "path").map(str::to_owned))
        .collect::<Result<BTreeSet<_>>>()?;
    let mut paths = BTreeSet::new();
    let mut names = BTreeSet::new();
    for s in sheets {
        if !paths.insert(field(s, "path")?.to_owned())
            || !names.insert(field(s, "name")?.to_lowercase())
        {
            return Err(crate::coded(
                "formula-ambiguous",
                "Ambiguous trace worksheet identity",
            ));
        }
    }
    if paths != indexed {
        return Err(crate::coded(
            "internal-mismatch",
            "Trace worksheet inventory disagrees with the accepted index",
        ));
    }
    let selected_index = sheets
        .iter()
        .position(|s| s["name"] == selected_sheet)
        .ok_or_else(|| crate::coded("internal-unknown-reference", "Unknown trace worksheet"))?;
    let db = Connection::open_with_flags(
        root.join("cells.sqlite"),
        OpenFlags::SQLITE_OPEN_READ_ONLY | OpenFlags::SQLITE_OPEN_NO_MUTEX,
    )
    .map_err(err)?;
    db.execute_batch("PRAGMA cache_size=-4096; PRAGMA temp_store=MEMORY; PRAGMA mmap_size=0;")
        .map_err(err)?;
    db.progress_handler(10000, Some(crate::cancel::interrupted));
    if !db
        .prepare_cached("SELECT EXISTS(SELECT 1 FROM versions WHERE id=?1)")
        .map_err(crate::err)?
        .query_row([version], |r| r.get::<_, bool>(0))
        .map_err(err)?
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown trace version",
        ));
    }
    let metadata = crate::package_edit::auditing_metadata(root)?;
    let supplied_sheets = sheets
        .iter()
        .map(|s| json!({"name":s["name"],"path":s["path"]}))
        .collect::<Vec<_>>();
    if metadata["sheets"] != json!(supplied_sheets)
        || !crate::package_edit::same_records(&metadata["namedRanges"], &view["namedRanges"])
        || !crate::package_edit::same_records(&metadata["tables"], &view["tables"])
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Trace reference metadata disagrees with the accepted source package",
        ));
    }
    let cursor = &input["cursor"];
    let (mut sheet_index, mut after, mut offset) = if cursor.is_null() {
        (
            if direction == "precedents" {
                selected_index
            } else {
                0
            },
            None,
            0,
        )
    } else {
        let map = cursor
            .as_object()
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid trace cursor shape"))?;
        if map.len() != 3
            || map
                .keys()
                .any(|k| !["sheetIndex", "after", "offset"].contains(&k.as_str()))
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid trace cursor fields",
            ));
        }
        let index = cursor["sheetIndex"]
            .as_u64()
            .filter(|n| *n < 1024)
            .ok_or_else(|| {
                crate::coded("internal-invalid-field", "Invalid trace cursor worksheet")
            })? as usize;
        let offset = cursor["offset"]
            .as_u64()
            .filter(|n| *n <= 1024)
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid trace cursor offset"))?
            as usize;
        let after = if cursor["after"].is_null() {
            None
        } else {
            Some(field(cursor, "after")?.to_owned())
        };
        (index, after, offset)
    };
    if sheet_index >= sheets.len()
        || offset > 1024
        || after
            .as_deref()
            .map(position)
            .transpose()?
            .is_some_and(|p| address(p.0, p.1) != after.as_deref().unwrap())
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid trace continuation",
        ));
    }
    let mut edges = Vec::new();
    let mut markers = Vec::new();
    let mut scanned = 0;
    let mut unknown = false;
    let mut next = Value::Null;
    if direction == "precedents" {
        if sheet_index != selected_index || after.is_some() {
            return Err(crate::coded(
                "formula-stale",
                "Precedent cursor changed its target",
            ));
        }
        let cells = crate::range::read_db(
            &db,
            &json!({"version":version,"sheet":sheets[selected_index]["path"],"range":selected_ref,"limit":1}),
        )?;
        let cell = cells["cells"]
            .as_array()
            .and_then(|c| c.first())
            .map(|p| &p[1])
            .unwrap_or(&Value::Null);
        let (e, m, u) = inspect(
            root,
            &db,
            version,
            view,
            &sheets[selected_index],
            selected_ref,
            cell,
        )?;
        unknown = u;
        scanned = u64::from(cell["formula"].is_string());
        let count = e.len() + m.len();
        if offset > count {
            return Err(crate::coded(
                "formula-limit",
                "Trace cursor exceeds its formula",
            ));
        }
        for item in offset..count.min(offset + limit) {
            if item < e.len() {
                edges.push(e[item].clone());
            } else {
                markers.push(m[item - e.len()].clone());
            }
        }
        if offset + limit < count {
            next = json!({"sheetIndex":selected_index,"after":null,"offset":offset+limit});
        }
    } else {
        // Accepted owned outputs are a separate direct ownership relation, not formula caches.
        if input["cursor"].is_null() {
            for owner in
                crate::spill::effective(&db, version, field(&sheets[selected_index], "path")?)?
            {
                if owner.master == selected
                    && owner.geometry.first != owner.geometry.last
                    && owner.proof["status"] != "blocked"
                {
                    let read = crate::range::read_db(
                        &db,
                        &json!({"version":version,"sheet":owner.sheet,"range":selected_ref,"limit":1}),
                    )?;
                    if crate::spill::accepted_cell(
                        &db,
                        version,
                        &owner.sheet,
                        selected,
                        &read["cells"][0][1],
                    )? {
                        edges.push(json!({"from":{"sheet":selected_sheet,"range":selected_ref},"to":{"sheet":selected_sheet,"range":owner.geometry.reference()},"kind":"spill-range","via":"formula"}));
                    }
                }
            }
        }
        if edges.len() + markers.len() == limit {
            return Ok(
                json!({"version":version,"sheet":selected_sheet,"cell":selected_ref,"direction":direction,"edges":edges,"markers":markers,"next":{"sheetIndex":sheet_index,"after":after,"offset":offset},"scannedFormulas":0,"unknown":unknown,"limited":false}),
            );
        }
        'sheets: while sheet_index < sheets.len() {
            crate::cancel::check()?;
            let sheet = &sheets[sheet_index];
            let page = crate::range::read_db(
                &db,
                &json!({"version":version,"sheet":sheet["path"],"range":"A1:XFD1048576","after":after,"formulas":true,"limit":MAX_SCAN-scanned}),
            )?;
            let cells = page["cells"].as_array().ok_or_else(|| {
                crate::coded("internal-invalid-field", "Invalid formula scan receipt")
            })?;
            for pair in cells {
                crate::cancel::check()?;
                let reference = pair[0].as_str().ok_or_else(|| {
                    crate::coded("internal-invalid-field", "Invalid formula scan reference")
                })?;
                let (e, m, u) = inspect(root, &db, version, view, sheet, reference, &pair[1])?;
                unknown |= u;
                let mut relevant = Vec::new();
                for value in e {
                    if value["from"]["sheet"] == selected_sheet
                        && contains(field(&value["from"], "range")?, selected)?
                    {
                        relevant.push(value);
                    }
                }
                // A resumed partial formula is still physically inspected on this call.
                let count = relevant.len() + m.len();
                if offset > count {
                    return Err(crate::coded(
                        "formula-limit",
                        "Dependent cursor exceeds its formula",
                    ));
                }
                scanned += 1;
                let capacity = limit - edges.len() - markers.len();
                let stop = count.min(offset + capacity);
                for item in offset..stop {
                    if item < relevant.len() {
                        edges.push(relevant[item].clone());
                    } else {
                        markers.push(m[item - relevant.len()].clone());
                    }
                }
                if stop < count {
                    next = json!({"sheetIndex":sheet_index,"after":after,"offset":stop});
                    break 'sheets;
                }
                after = Some(reference.into());
                offset = 0;
                if edges.len() + markers.len() == limit || scanned == MAX_SCAN {
                    next = json!({"sheetIndex":sheet_index,"after":after,"offset":0});
                    break 'sheets;
                }
            }
            if page["nextCursor"].is_null() {
                sheet_index += 1;
                after = None;
                offset = 0;
            } else {
                next = json!({"sheetIndex":sheet_index,"after":after,"offset":0});
                break;
            }
        }
    }
    crate::cancel::check()?;
    Ok(
        json!({"version":version,"sheet":selected_sheet,"cell":selected_ref,"direction":direction,"edges":edges,"markers":markers,"next":next,"scannedFormulas":scanned,"unknown":unknown,"limited":false}),
    )
}
