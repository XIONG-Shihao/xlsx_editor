//! Stream reference-planned insertion. No literal-cell graph is materialized.
use crate::xml::{NS, address, attrs, position};
use crate::{Result, database, err, field, local};
use quick_xml::{
    NsReader, Writer,
    events::{BytesEnd, BytesStart, BytesText, Event},
};
use rusqlite::{Connection, OptionalExtension, params};
use serde_json::{Value, json};
use std::{
    collections::{BTreeMap, HashSet},
    fs::{self, File},
    io::{BufRead, BufReader, BufWriter, Write},
    path::Path,
};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

/// Plan qualified formulas beside their versioned storage. A semantic miss may
/// use the existing coordinator fallback; I/O, quota and cancellation errors may
/// not. Publish the spool only after every formula has been rewritten.
pub fn prepare(root: &Path, input: &Value) -> Result<Value> {
    let partial = local(root, "insertion-formulas.partial")?;
    let result = prepare_inner(root, input, &partial);
    if result.is_err() {
        let _ = fs::remove_file(&partial);
    }
    crate::cancel::check()?;
    match result {
        Err(reason) if crate::sentence(&reason).starts_with("Unsupported structural formula: ") => {
            Ok(json!({"supported":false,"reason":reason}))
        }
        other => other,
    }
}

fn prepare_inner(root: &Path, input: &Value, partial: &Path) -> Result<Value> {
    let command = &input["command"];
    let operation = field(command, "operation")?;
    if !["insertRow", "deleteRow", "insertColumn", "deleteColumn"].contains(&operation) {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid structural planning operation",
        ));
    }
    let columns = operation.ends_with("Column");
    let remove = operation.starts_with("delete");
    let limit = if columns { 16384 } else { 1048576 };
    let index = command["index"]
        .as_u64()
        .ok_or_else(|| crate::coded("workbook-malformed", "Invalid structural index"))?;
    let count = match command.get("count") {
        Some(value) => value
            .as_u64()
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid structural count"))?,
        None => 1,
    };
    if count == 0 || index >= limit || count > limit - index {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid structural planning bounds",
        ));
    }
    let (index, count) = (index as u32, count as u32);
    let target = field(command, "sheet")?;
    let db = database(root)?;
    // Everything beside the worksheets (names, tables, notes, pictures, chart
    // references) is planned here from the package; the caller names only the
    // sheet, the rows or columns and how many.
    let structure = crate::package_edit::structural_plan(root, &db, input["version"].as_u64().unwrap_or(0), command)?;
    let sheets = structure
        .sheets
        .iter()
        .map(|(name, path)| json!({"name": name, "path": path}))
        .collect::<Vec<_>>();
    let target_path = structure
        .sheets
        .iter()
        .find(|(name, _)| name == target)
        .map(|(_, path)| path.clone())
        .unwrap_or_default();
    let metadata = json!({
        "axis": if columns { "column" } else { "row" },
        "remove": remove,
        "target": target_path,
        "index": index,
        "count": count,
        "sheets": structure.sheets.iter().map(|(_, path)| path).collect::<Vec<_>>(),
        "tableChanges": structure.table_changes,
        "replacements": structure.replacements,
        "removed": structure.removed,
    });
    crate::quota::write(
        local(root, "insertion-metadata.json")?,
        serde_json::to_vec(&metadata).map_err(err)?,
    )
    .map_err(err)?;
    let rewrite = crate::used_range_transfer::admit_writer(
        root,
        &db,
        input,
        crate::used_range_transfer::WriterKind::Insertion,
    )?;
    let binding = if rewrite.is_some() {
        let inventory = crate::used_range_transfer::worksheet_inventory(&local(
            root,
            field(input, "source")?,
        )?)?;
        if sheets.len() != inventory.len()
            || sheets
                .iter()
                .zip(&inventory)
                .any(|(a, b)| a["name"] != b.name || a["path"] != b.path)
        {
            return Err(crate::coded(
                "internal-mismatch",
                "Insertion planning inventory differs from native input",
            ));
        }
        Some(
            json!({"axis":if columns {"column"} else {"row"},"remove":remove,"target":target_path,"index":index,"count":count,"sheets":inventory.iter().map(|sheet|&sheet.path).collect::<Vec<_>>(),"tableChanges":structure.table_changes}),
        )
    } else {
        None
    };
    crate::calculation::snapshot(
        &db,
        input["version"]
            .as_u64()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing version"))?,
    )?;
    let mut output = BufWriter::with_capacity(65536, crate::quota::create(partial).map_err(err)?);
    let mut formula_count = 0_u64;
    let mut request = json!({"target":target,"index":index,"count":count,"axis":if columns {"column"} else {"row"},"remove":remove,"tableChanges":structure.table_changes,"items":[]});
    let rewrite = |request: &Value| -> Result<String> {
        let result = crate::formula::rewrite(request).map_err(|reason| {
            crate::rephrase("formula-unsupported", reason, |said| {
                format!("Unsupported structural formula: {said}")
            })
        })?;
        Ok(result["formulas"][0]
            .as_str()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing rewritten formula"))?
            .into())
    };
    // Excel leaves an inserted row empty, except inside a table, where the
    // new row continues the formulas of the columns the table declares as
    // calculated: first body row, last row, first and last column.
    let bodies = structure.calculated.clone();
    for sheet in &sheets {
        let name = field(sheet, "name")?;
        let path = field(sheet, "path")?;
        let moving = name == target;
        let mut after = None;
        loop {
            let page = crate::calculation::page(&db, path, (0, 0), (1048575, 16383), after, true)?;
            for (reference, cell) in &page {
                crate::cancel::check()?;
                if !cell["formulaType"].is_null() {
                    return Err(crate::coded(
                        "formula-unsupported",
                        "Unsupported structural formula: special formula",
                    ));
                }
                let (r, c) = position(reference)?;
                let coordinate = if columns { c } else { r };
                if moving && remove && coordinate >= index && coordinate < index + count {
                    continue;
                }
                let shifted = if moving && coordinate >= index {
                    if remove {
                        coordinate - count
                    } else {
                        coordinate + count
                    }
                } else {
                    coordinate
                };
                if u64::from(shifted) >= limit {
                    return Err(crate::coded(
                        "formula-limit",
                        "Moved formula exceeds worksheet boundary",
                    ));
                }
                request["items"] = json!([{"formula":field(cell,"formula")?,"owner":name,"reference":reference,"boundary":!columns&&moving&&coordinate==index}]);
                let formula = rewrite(&request)?;
                let destination = if columns {
                    address(r, shifted)
                } else {
                    address(shifted, c)
                };
                let record = json!({"sheet":path,"ref":destination,"formula":formula});
                serde_json::to_writer(&mut output, &record).map_err(err)?;
                output.write_all(b"\n").map_err(err)?;
                formula_count += 1;
                // Rewrite the row above first, then copy its relative
                // references into each inserted table row.
                if moving
                    && !columns
                    && !remove
                    && index > 0
                    && r == index - 1
                    && bodies.iter().any(|(top, bottom, first, last)| {
                        r >= *top && index <= *bottom && c >= *first && c <= *last
                    })
                {
                    for offset in 1..=count {
                        crate::cancel::check()?;
                        let copied =
                            rewrite(&json!({"items":[{"formula":formula,"copyRows":offset}]}))?;
                        serde_json::to_writer(
                            &mut output,
                            &json!({"sheet":path,"ref":address(r+offset,c),"formula":copied}),
                        )
                        .map_err(err)?;
                        output.write_all(b"\n").map_err(err)?;
                        formula_count += 1;
                    }
                }
            }
            if page.len() < 512 {
                break;
            }
            after = Some(position(&page.last().unwrap().0)?);
        }
    }
    output.flush().map_err(err)?;
    drop(output);
    crate::cancel::check()?;
    fs::rename(partial, local(root, "insertion-formulas.jsonl")?).map_err(err)?;
    if let Some(binding) = binding {
        let binding = serde_json::to_string(&binding).map_err(err)?;
        let formula_sha = crate::hash(&local(root, "insertion-formulas.jsonl")?)?;
        let input_id = field(input, "rewriteInputId")?;
        crate::quota::reserve_artifact((binding.len() + formula_sha.len()) as u64 + 4096)?;
        db.execute_batch("CREATE TABLE IF NOT EXISTS export_rewrite_insertion_plans(input_id TEXT PRIMARY KEY,binding TEXT NOT NULL,formula_sha TEXT NOT NULL) WITHOUT ROWID;").map_err(err)?;
        // An input authorizes one immutable native plan, never a replaceable
        // JavaScript claim about its row/column movement.
        db.execute(
            "INSERT INTO export_rewrite_insertion_plans VALUES(?1,?2,?3)",
            params![input_id, binding, formula_sha],
        )
        .map_err(err)?;
    }
    Ok(
        json!({"supported":true,"formulaCount":formula_count,"inputCells":0,"nativeStructuralPlan":true,"headers":structure.headers}),
    )
}

fn changed(e: &BytesStart<'_>, key: &str, value: String) -> Result<BytesStart<'static>> {
    let mut attributes = attrs(e)?;
    attributes.insert(key.into(), value);
    let mut element = BytesStart::new(String::from_utf8_lossy(e.name().as_ref()).into_owned());
    for (k, v) in &attributes {
        element.push_attribute((k.as_str(), v.as_str()));
    }
    Ok(element)
}
fn formula(db: &Connection, sheet: &str, reference: &str) -> Result<Option<String>> {
    db.prepare_cached("SELECT formula FROM insertion_formulas WHERE sheet=?1 AND ref=?2")
        .map_err(err)?
        .query_row(params![sheet, reference], |r| r.get(0))
        .optional()
        .map_err(err)
}
fn inserted<W: Write>(
    writer: &mut Writer<W>,
    db: &Connection,
    sheet: &str,
    index: u32,
    count: u32,
    styles: &BTreeMap<u32, Option<String>>,
    above: &[(String, String)],
) -> Result<()> {
    if styles.is_empty() && above.is_empty() {
        return Ok(());
    }
    for row in index..index + count {
        crate::cancel::check()?;
        let mut element = BytesStart::new("row");
        element.push_attribute(("xmlns", std::str::from_utf8(NS).unwrap()));
        element.push_attribute(("r", (row + 1).to_string().as_str()));
        // A new row is set up like the row above it, as in Excel: its height,
        // its format and its place in an outline group.
        for (key, value) in above {
            element.push_attribute((key.as_str(), value.as_str()));
        }
        writer.write_event(Event::Start(element)).map_err(err)?;
        for (col, style) in styles {
            let reference = address(row, *col);
            let mut cell = BytesStart::new("c");
            cell.push_attribute(("r", reference.as_str()));
            if let Some(style) = style {
                cell.push_attribute(("s", style.as_str()));
            }
            if let Some(f) = formula(db, sheet, &reference)? {
                writer.write_event(Event::Start(cell)).map_err(err)?;
                writer
                    .write_event(Event::Start(BytesStart::new("f")))
                    .map_err(err)?;
                writer
                    .write_event(Event::Text(BytesText::new(&f)))
                    .map_err(err)?;
                writer
                    .write_event(Event::End(BytesEnd::new("f")))
                    .map_err(err)?;
                writer
                    .write_event(Event::End(BytesEnd::new("c")))
                    .map_err(err)?;
            } else {
                writer.write_event(Event::Empty(cell)).map_err(err)?;
            }
        }
        writer
            .write_event(Event::End(BytesEnd::new("row")))
            .map_err(err)?;
    }
    Ok(())
}
pub fn apply(root: &Path, input: &Value) -> Result<Value> {
    let plan: Value =
        serde_json::from_slice(&fs::read(local(root, "insertion-metadata.json")?).map_err(err)?)
            .map_err(err)?;
    let target = field(&plan, "target")?;
    let remove = plan["remove"] == true;
    let columns = plan["axis"] == "column";
    let index = plan["index"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing insertion index"))?;
    let count = plan["count"]
        .as_u64()
        .filter(|n| *n > 0)
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid insertion count"))?;
    if index
        .checked_add(count)
        .is_none_or(|end| end > if columns { 16384 } else { 1048576 })
    {
        return Err(crate::coded(
            "workbook-limit",
            "Insertion exceeds row boundary",
        ));
    }
    let (index, count) = (index as u32, count as u32);
    let db = database(root)?;
    let rewrite = crate::used_range_transfer::admit_writer(
        root,
        &db,
        input,
        crate::used_range_transfer::WriterKind::Insertion,
    )?;
    if rewrite.is_some() {
        let (binding, formula_sha): (String, String) = db
            .query_row(
                "SELECT binding,formula_sha FROM export_rewrite_insertion_plans WHERE input_id=?1",
                [field(input, "rewriteInputId")?],
                |row| Ok((row.get(0)?, row.get(1)?)),
            )
            .map_err(err)?;
        let binding: Value = serde_json::from_str(&binding).map_err(err)?;
        if ["axis", "remove", "target", "index", "count", "sheets"]
            .iter()
            .any(|key| plan[*key] != binding[*key])
            || crate::hash(&local(root, "insertion-formulas.jsonl")?)? != formula_sha
        {
            return Err(crate::coded(
                "internal-mismatch",
                "Insertion writer differs from its immutable native plan",
            ));
        }
    }
    // The engine works out every moved worksheet record itself.
    let fragments = crate::package_edit::insertion_records(root, input, &plan)?;
    db.execute_batch("CREATE TEMP TABLE insertion_formulas(sheet TEXT,ref TEXT,formula TEXT,PRIMARY KEY(sheet,ref));").map_err(err)?;
    for line in
        BufReader::new(File::open(local(root, "insertion-formulas.jsonl")?).map_err(err)?).lines()
    {
        crate::cancel::check()?;
        let record: Value = serde_json::from_str(&line.map_err(err)?).map_err(err)?;
        db.execute(
            "INSERT INTO insertion_formulas VALUES(?1,?2,?3)",
            params![
                field(&record, "sheet")?,
                field(&record, "ref")?,
                field(&record, "formula")?
            ],
        )
        .map_err(err)?;
    }
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, field(input, "source")?)?).map_err(err)?,
    ))
    .map_err(err)?;
    let partial = local(root, "insertion.partial")?;
    let output = local(root, "insertion.xlsx")?;
    let mut writer = ZipWriter::new(crate::quota::create(&partial).map_err(err)?);
    for i in 0..zip.len() {
        crate::cancel::check()?;
        let mut member = zip.by_index(i).map_err(err)?;
        let name = member.name().to_owned();
        if plan["removed"]
            .as_array()
            .unwrap()
            .iter()
            .any(|v| v == &name)
        {
            continue;
        }
        if let Some(replacement) = plan["replacements"][&name].as_str() {
            writer
                .start_file(
                    &name,
                    SimpleFileOptions::default()
                        .compression_method(zip::CompressionMethod::Deflated),
                )
                .map_err(err)?;
            writer.write_all(replacement.as_bytes()).map_err(err)?;
            continue;
        }
        if !plan["sheets"]
            .as_array()
            .unwrap()
            .iter()
            .any(|v| v == &name)
        {
            if name.contains("pivotCache/") && name.ends_with(".xml") && member.size() < 1 << 28 {
                let mut bytes = Vec::new();
                std::io::Read::read_to_end(&mut member, &mut bytes).map_err(err)?;
                let moved =
                    crate::package_edit::insertion_pivot_source(root, input, &plan, &bytes)?;
                writer
                    .start_file(
                        &name,
                        SimpleFileOptions::default()
                            .compression_method(zip::CompressionMethod::Deflated),
                    )
                    .map_err(err)?;
                writer
                    .write_all(moved.as_deref().unwrap_or(&bytes))
                    .map_err(err)?;
                continue;
            }
            writer.raw_copy_file(member).map_err(err)?;
            continue;
        }
        if name != target {
            let referenced: bool = db
                .query_row(
                    "SELECT EXISTS(SELECT 1 FROM insertion_formulas WHERE sheet=?1)",
                    [&name],
                    |r| r.get(0),
                )
                .map_err(err)?;
            if !referenced && fragments[&name].as_object().is_none_or(|m| m.is_empty()) {
                writer.raw_copy_file(member).map_err(err)?;
                continue;
            }
        }
        writer
            .start_file(
                &name,
                SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated),
            )
            .map_err(err)?;
        let moving = name == target;
        let mut reader = NsReader::from_reader(BufReader::new(member));
        let mut xml = Writer::new(&mut writer);
        let mut buffer = Vec::new();
        let (mut depth, mut skip) = (0, 0);
        let mut styles = BTreeMap::new();
        let mut above: Vec<(String, String)> = Vec::new();
        let mut wrote = remove || columns;
        let mut current_formula = None;
        let mut written_fragments = HashSet::new();
        let mut row = 0;
        // A new column takes the cell formats of the column on its left, as in
        // Excel: the name and format of the cell to copy, and whether it has
        // just been written.
        let mut left: Option<(String, String)> = None;
        let mut copy_left = false;
        loop {
            crate::cancel::check()?;
            let (ns, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
            let main = ns == quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(NS));
            let mut replacement = None;
            match &event {
                Event::Start(e) | Event::Empty(e) => {
                    let empty = matches!(event, Event::Empty(_));
                    let tag = e.local_name();
                    let tag = tag.as_ref();
                    if skip > 0 {
                        if !empty {
                            skip += 1
                        }
                        buffer.clear();
                        continue;
                    }
                    if main && depth == 1 {
                        let key = std::str::from_utf8(tag).map_err(err)?;
                        if let Some(fragment) = fragments[&name][key].as_str() {
                            if written_fragments.insert(key.to_string()) {
                                xml.get_mut().write_all(fragment.as_bytes()).map_err(err)?;
                            }
                            if !empty {
                                skip = 1;
                            }
                            buffer.clear();
                            continue;
                        }
                    }
                    if main && depth == 1 && tag == b"dimension" && moving {
                        if !empty {
                            skip = 1
                        }
                        buffer.clear();
                        continue;
                    }
                    if main && depth == 2 && tag == b"row" {
                        row = attrs(e)?
                            .get("r")
                            .ok_or_else(|| {
                                crate::coded("internal-missing-field", "Missing row index")
                            })?
                            .parse::<u32>()
                            .map_err(err)?
                            .checked_sub(1)
                            .ok_or_else(|| {
                                crate::coded("workbook-malformed", "Invalid row index")
                            })?;
                        if moving && !columns && !remove && index > 0 && row == index - 1 {
                            above = attrs(e)?
                                .into_iter()
                                .filter(|(key, _)| {
                                    !["r", "spans", "hidden", "collapsed"].contains(&key.as_str())
                                        && !key.starts_with("xmlns")
                                        && !key.contains(':')
                                })
                                .collect();
                        }
                        if moving && columns {
                            let mut element = BytesStart::new(
                                String::from_utf8_lossy(e.name().as_ref()).into_owned(),
                            );
                            for (k, v) in attrs(e)? {
                                if k != "spans" {
                                    element.push_attribute((k.as_str(), v.as_str()));
                                }
                            }
                            replacement = Some(if empty {
                                Event::Empty(element)
                            } else {
                                Event::Start(element)
                            });
                        }
                        if moving && !columns && remove && row >= index && row < index + count {
                            if !empty {
                                skip = 1;
                            }
                            buffer.clear();
                            continue;
                        }
                        if moving && !columns && row >= index {
                            if !wrote {
                                inserted(&mut xml, &db, &name, index, count, &styles, &above)?;
                                wrote = true;
                            }
                            if !remove && row + count >= 1048576 {
                                return Err(crate::coded(
                                    "workbook-limit",
                                    "Moved row exceeds worksheet boundary",
                                ));
                            }
                            let moved = if remove { row - count } else { row + count };
                            let element = changed(e, "r", (moved + 1).to_string())?;
                            replacement = Some(if empty {
                                Event::Empty(element)
                            } else {
                                Event::Start(element)
                            });
                        }
                    }
                    if main && depth == 3 && tag == b"c" {
                        let attributes = attrs(e)?;
                        let (r, c) = position(attributes.get("r").ok_or_else(|| {
                            crate::coded("internal-missing-field", "Missing cell reference")
                        })?)?;
                        if r != row {
                            return Err(crate::coded(
                                "workbook-unsupported",
                                "Cell outside its row",
                            ));
                        }
                        if moving && !columns && !remove && index > 0 && row == index - 1 {
                            styles.insert(c, attributes.get("s").cloned());
                        }
                        if moving && columns && !remove && index > 0 && c == index - 1 {
                            if let Some(style) = attributes.get("s").filter(|style| *style != "0") {
                                left = Some((
                                    String::from_utf8_lossy(e.name().as_ref()).into_owned(),
                                    style.clone(),
                                ));
                                copy_left = empty;
                            }
                        }
                        let coordinate = if columns { c } else { r };
                        if moving
                            && columns
                            && remove
                            && coordinate >= index
                            && coordinate < index + count
                        {
                            if !empty {
                                skip = 1;
                            }
                            buffer.clear();
                            continue;
                        }
                        let shifted = if moving && coordinate >= index {
                            if remove {
                                coordinate - count
                            } else {
                                coordinate + count
                            }
                        } else {
                            coordinate
                        };
                        if shifted >= if columns { 16384 } else { 1048576 } {
                            return Err(crate::coded(
                                "workbook-limit",
                                "Moved cell exceeds worksheet boundary",
                            ));
                        }
                        let reference = if columns {
                            address(r, shifted)
                        } else {
                            address(shifted, c)
                        };
                        current_formula = formula(&db, &name, &reference)?;
                        let mut element = changed(e, "r", reference)?;
                        if current_formula.is_some() {
                            let attributes = attrs(&element)?;
                            element = BytesStart::new(
                                String::from_utf8_lossy(e.name().as_ref()).into_owned(),
                            );
                            for (k, v) in attributes {
                                if k != "t" {
                                    element.push_attribute((k.as_str(), v.as_str()));
                                }
                            }
                        }
                        replacement = Some(if empty {
                            Event::Empty(element)
                        } else {
                            Event::Start(element)
                        });
                    }
                    if main && depth == 4 && tag == b"f" {
                        if attrs(e)?.contains_key("t") {
                            return Err(crate::coded(
                                "needs-structural-writer",
                                "Special formula insertion requires the shared structural writer",
                            ));
                        }
                        let f = current_formula.as_ref().ok_or_else(|| {
                            crate::coded("internal-missing-field", "Missing transformed formula")
                        })?;
                        xml.write_event(Event::Start(e.clone())).map_err(err)?;
                        xml.write_event(Event::Text(BytesText::new(f)))
                            .map_err(err)?;
                        xml.write_event(Event::End(BytesEnd::new(
                            String::from_utf8_lossy(e.name().as_ref()).into_owned(),
                        )))
                        .map_err(err)?;
                        if !empty {
                            skip = 1
                        }
                        buffer.clear();
                        continue;
                    }
                    if main && depth == 4 && tag == b"v" && current_formula.is_some() {
                        if !empty {
                            skip = 1
                        }
                        buffer.clear();
                        continue;
                    }
                    if !empty {
                        depth += 1;
                    }
                }
                Event::End(e) => {
                    if skip > 0 {
                        skip -= 1;
                        buffer.clear();
                        continue;
                    }
                    if main && depth == 4 && e.local_name().as_ref() == b"c" && left.is_some() {
                        copy_left = true;
                    }
                    if main
                        && depth == 2
                        && e.local_name().as_ref() == b"sheetData"
                        && moving
                        && !wrote
                    {
                        inserted(&mut xml, &db, &name, index, count, &styles, &above)?;
                        wrote = true;
                    }
                    depth -= 1;
                }
                Event::Eof => break,
                _ => {
                    if skip > 0 {
                        buffer.clear();
                        continue;
                    }
                }
            }
            xml.write_event(replacement.unwrap_or(event)).map_err(err)?;
            if copy_left {
                let (cell, style) = left.take().unwrap();
                for offset in 0..count {
                    let mut element = BytesStart::new(cell.as_str());
                    element.push_attribute(("r", address(row, index + offset).as_str()));
                    element.push_attribute(("s", style.as_str()));
                    xml.write_event(Event::Empty(element)).map_err(err)?;
                }
                copy_left = false;
            }
            buffer.clear();
        }
    }
    writer.finish().map_err(err)?.sync_all().map_err(err)?;
    fs::rename(partial, output).map_err(err)?;
    let prepared = rewrite
        .map(|rewrite| {
            crate::used_range_transfer::prepare_output(
                root,
                &db,
                rewrite,
                crate::used_range_transfer::NativeMapping::Axis {
                    sheet: target.to_owned(),
                    axis: if columns {
                        crate::used_range_transfer::Axis::Column
                    } else {
                        crate::used_range_transfer::Axis::Row
                    },
                    index,
                    count,
                    remove,
                },
                "insertion.xlsx",
            )
        })
        .transpose()?;
    drop(db);
    let mut result = json!({"file":"insertion.xlsx"});
    if let Some(prepared) = prepared {
        result["rewriteId"] = json!(crate::used_range_transfer::commit_output(root, prepared)?);
    }
    Ok(result)
}
