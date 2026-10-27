use crate::xml::{NS, attrs, borrowed_attrs, node_name, position, spreadsheet_text, xml_text};
use crate::{Result, database, err, hash, local};
use quick_xml::{NsReader, events::Event, name::ResolveResult};
use rusqlite::{Connection, params, params_from_iter, types::Value as SqlValue};
use serde_json::{Value, json};
use std::borrow::Cow;
use std::fs::File;
use std::io::{self, BufReader, Read, Seek, SeekFrom, Write};
use std::path::Path;
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

fn shared_strings(source: impl io::BufRead, db: &Connection) -> Result<()> {
    let _span = crate::diagnostics::span("import.shared_strings");
    let mut reader = NsReader::from_reader(source);
    let mut buffer = Vec::new();
    let mut stack = Vec::new();
    let mut value = String::new();
    let mut text_value = String::new();
    let mut index = 0u64;
    let mut protected_value = false;
    let mut insert = db
        .prepare_cached("INSERT INTO strings(id,value) VALUES (?1,?2)")
        .map_err(err)?;
    loop {
        crate::cancel::check()?;
        let (ns, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        match event {
            Event::DocType(_) => {
                return Err(crate::coded(
                    "package-malformed",
                    "XML entities are unsupported",
                ));
            }
            Event::Start(element) => {
                let name = node_name(
                    ns == ResolveResult::Bound(quick_xml::name::Namespace(NS)),
                    &element,
                );
                if stack.is_empty() && name != "sst" {
                    return Err(crate::coded(
                        "package-unsupported",
                        "Unsupported shared strings root",
                    ));
                }
                if stack.len() == 1 && name == "si" {
                    value.clear();
                    protected_value = attrs(&element)?
                        .keys()
                        .any(|key| !crate::preservation::safe_text_attribute(key));
                } else if stack.get(1).is_some_and(|s| s == "si")
                    && (name != "t"
                        || attrs(&element)?
                            .keys()
                            .any(|key| !crate::preservation::safe_text_attribute(key)))
                {
                    protected_value = true;
                }
                if name == "t" {
                    text_value.clear();
                }
                stack.push(name);
            }
            Event::Empty(element)
                if ns == ResolveResult::Bound(quick_xml::name::Namespace(NS))
                    && element.local_name().as_ref() == b"si"
                    && stack.len() == 1 =>
            {
                insert.execute(params![index, ""]).map_err(err)?;
                if attrs(&element)?
                    .keys()
                    .any(|key| !crate::preservation::safe_text_attribute(key))
                {
                    db.prepare_cached("INSERT INTO shared_value_features(id,reason) VALUES(?1,?2)")
                        .map_err(err)?
                        .execute(params![index, crate::preservation::REASON])
                        .map_err(err)?;
                }
                index += 1;
            }
            Event::Text(text)
                if stack.last().is_some_and(|s| s == "t") && !stack.iter().any(|s| s == "rPh") =>
            {
                text_value.push_str(&xml_text(text.as_ref())?)
            }
            Event::CData(text)
                if stack.last().is_some_and(|s| s == "t") && !stack.iter().any(|s| s == "rPh") =>
            {
                text_value.push_str(
                    &text
                        .decode()
                        .map_err(err)?
                        .replace("\r\n", "\n")
                        .replace('\r', "\n"),
                )
            }
            Event::Empty(element) if stack.get(1).is_some_and(|s| s == "si") => {
                if node_name(
                    ns == ResolveResult::Bound(quick_xml::name::Namespace(NS)),
                    &element,
                ) != "t"
                    || attrs(&element)?
                        .keys()
                        .any(|key| !crate::preservation::safe_text_attribute(key))
                {
                    protected_value = true;
                }
            }
            Event::Comment(_) | Event::PI(_) if stack.get(1).is_some_and(|s| s == "si") => {
                protected_value = true;
            }
            Event::End(_) => {
                if stack.last().is_some_and(|s| s == "t") && !stack.iter().any(|s| s == "rPh") {
                    value.push_str(&spreadsheet_text(&text_value)?);
                    text_value.clear();
                }
                if stack.len() == 2 && stack.last().is_some_and(|s| s == "si") {
                    insert.execute(params![index, value]).map_err(err)?;
                    if protected_value {
                        db.prepare_cached(
                            "INSERT INTO shared_value_features(id,reason) VALUES(?1,?2)",
                        )
                        .map_err(err)?
                        .execute(params![index, crate::preservation::REASON])
                        .map_err(err)?;
                    }
                    index += 1;
                }
                stack.pop();
            }
            Event::Eof => break,
            _ => (),
        }
        buffer.clear();
    }
    crate::diagnostics::count("import.strings", index);
    Ok(())
}

#[derive(Default)]
struct Cell {
    row: u32,
    col: u32,
    kind: String,
    extent_children: u8,
    raw: String,
    value_present: bool,
    inline: String,
    inline_piece: String,
    protected_value: bool,
    preservation_mask: u8,
    formula: Option<String>,
    formula_attrs: Option<String>,
    style: u32,
    style_explicit: bool,
    start: u64,
    open: u64,
    spans: Vec<(u64, u64)>,
    child_start: Option<u64>,
}

const CELL_COLUMNS: usize = 15;
const CELL_BATCH: usize = 32;
const CELL_BATCH_BYTES: usize = 64 * 1024;
const CELL_INSERT: &str = "INSERT INTO cells(sheet,r,c,kind,raw,inline,formula,formula_attrs,style,start,open,close,end,spans,style_explicit) VALUES ";
fn cell_insert_sql(rows: usize) -> String {
    format!(
        "{CELL_INSERT}{}",
        vec!["(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"; rows].join(",")
    )
}
fn flush_cells(db: &Connection, sql: &[String], values: &mut Vec<SqlValue>) -> Result<()> {
    if values.is_empty() {
        return Ok(());
    }
    crate::cancel::check()?;
    let _span = crate::diagnostics::span("import.batch_insert");
    let rows = values.len() / CELL_COLUMNS;
    db.prepare_cached(&sql[rows - 1])
        .map_err(err)?
        .execute(params_from_iter(values.iter()))
        .map_err(err)?;
    crate::diagnostics::count("import.insert_statements", 1);
    crate::diagnostics::count("import.insert_rows", rows as u64);
    values.clear();
    Ok(())
}

/// The legacy worksheet protection flags restrict edits but add no cell or
/// axis geometry. Qualifying this exact profile never removes its source XML;
/// the workbook owner still handles it as read-only and copies the source.
fn qualified_sheet_protection(attributes: &std::collections::BTreeMap<&str, Cow<'_, str>>) -> bool {
    attributes.iter().all(|(name, value)| {
        if *name == "xmlns" || name.starts_with("xmlns:") {
            return true;
        }
        if *name == "password" {
            return value.len() == 4 && value.bytes().all(|byte| byte.is_ascii_hexdigit());
        }
        [
            "sheet",
            "objects",
            "scenarios",
            "formatCells",
            "formatColumns",
            "formatRows",
            "insertColumns",
            "insertRows",
            "insertHyperlinks",
            "deleteColumns",
            "deleteRows",
            "selectLockedCells",
            "sort",
            "autoFilter",
            "pivotTables",
            "selectUnlockedCells",
        ]
        .contains(name)
            && matches!(value.as_ref(), "0" | "1" | "false" | "true")
    })
}

fn worksheet(file: &Path, db: &Connection, sheet: &str) -> Result<Value> {
    let _span = crate::diagnostics::span("import.worksheet");
    let diagnostic = crate::diagnostics::enabled();
    let mut row_count = 0;
    let mut reader = NsReader::from_reader(BufReader::new(File::open(file).map_err(err)?));
    let mut buffer = Vec::new();
    let mut stack: Vec<Cow<'static, str>> = Vec::new();
    let mut root_seen = false;
    let mut current: Option<Cell> = None;
    let (mut rows, mut columns, mut count, mut formulas) = (20u32, 8u32, 0u64, 0u64);
    let (mut sheet_data_start, mut sheet_data_end) = (None, None);
    let (mut data_open, mut data_close) = (None, None);
    let mut special_formulas = 0u64;
    let mut protected_formula_caches = 0u64;
    let mut extended = false;
    let mut dimension = Value::Null;
    let mut dimension_insert = None;
    let mut used_range_facts = crate::used_range::SourceFacts::new();
    let mut extent_nodes = std::collections::BTreeSet::new();
    let (mut row_number, mut next_column) = (0u32, 0u32);
    let (mut row_start, mut row_open) = (0u64, 0u64);
    let mut heights = Vec::new();
    let mut hidden = Vec::new();
    let mut visible = Vec::new();
    let mut hidden_columns = Vec::new();
    let insert_sql: Vec<_> = (1..=CELL_BATCH).map(cell_insert_sql).collect();
    let mut insert_values = Vec::with_capacity(CELL_COLUMNS * CELL_BATCH);
    let mut insert_bytes = 0;
    let mut insert_row = db
        .prepare_cached("INSERT INTO rows(sheet,r,start,open,close,end) VALUES (?1,?2,?3,?4,?5,?6)")
        .map_err(err)?;
    loop {
        crate::cancel::check()?;
        let start = reader.buffer_position();
        let (ns, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        let bound = ns == ResolveResult::Bound(quick_xml::name::Namespace(NS));
        let empty = matches!(event, Event::Empty(_));
        match event {
            Event::DocType(_) => {
                return Err(crate::coded(
                    "package-malformed",
                    "XML entities are unsupported",
                ));
            }
            Event::Start(ref element) | Event::Empty(ref element) => {
                let name = node_name(bound, element);
                if stack.is_empty() && name != "worksheet" {
                    return Err(crate::coded(
                        "package-unsupported",
                        "Unsupported worksheet namespace or root",
                    ));
                }
                if stack.is_empty() {
                    if root_seen {
                        return Err(crate::coded(
                            "package-malformed",
                            "Multiple worksheet XML roots",
                        ));
                    }
                    root_seen = true;
                }
                if stack.is_empty() && empty {
                    used_range_facts.issue(
                        "unqualified-worksheet-root",
                        "A self-closing worksheet needs a structural writer before normalization.",
                    );
                }
                if name == "extLst" || !bound && !stack.is_empty() {
                    extended = true;
                }
                let attributes = borrowed_attrs(element)?;
                if (element.local_name().as_ref() == b"sheetProtection"
                    && (!bound || stack.len() != 1 || !qualified_sheet_protection(&attributes)))
                    || stack
                        .get(1)
                        .is_some_and(|ancestor| ancestor == "sheetProtection")
                {
                    used_range_facts.issue(
                        "unqualified-sheet-protection",
                        "Unknown, misplaced or extended worksheet protection remains protected.",
                    );
                }
                if name == "sheetFormatPr" && stack.len() == 1 {
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(err)?;
                        if attribute.key.as_ref() == b"x14ac:dyDescent"
                            && reader.resolve_attribute(attribute.key).0
                                != ResolveResult::Bound(quick_xml::name::Namespace(
                                    b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
                                ))
                        {
                            used_range_facts.issue(
                                "unqualified-sheet-defaults",
                                "Sheet descent uses an unqualified namespace.",
                            );
                        }
                    }
                }
                if stack.len() == 1
                    && [
                        "sheetPr",
                        "sheetFormatPr",
                        "cols",
                        "sheetData",
                        "mergeCells",
                        "sheetProtection",
                    ]
                    .contains(&name.as_ref())
                    && !extent_nodes.insert(name.to_string())
                {
                    used_range_facts.issue(
                        "duplicate-worksheet-structure",
                        "Repeated worksheet extent structures remain unqualified.",
                    );
                }
                if stack.len() == 2 && stack[1] == "dimension" {
                    used_range_facts.issue(
                        "opaque-dimension",
                        "Dimension child markup remains protected.",
                    );
                }
                if stack.len() == 1
                    && ![
                        "sheetPr",
                        "dimension",
                        "sheetViews",
                        "sheetFormatPr",
                        "cols",
                        "sheetData",
                        "mergeCells",
                        "sheetProtection",
                        "pageMargins",
                        "pageSetup",
                        "printOptions",
                        "headerFooter",
                        "picture",
                    ]
                    .contains(&name.as_ref())
                {
                    used_range_facts.issue("unqualified-worksheet-metadata", "Worksheet metadata outside the qualified imported extent model remains protected.");
                }
                if bound && stack.len() == 2 && stack[1] == "sheetData" && name != "row" {
                    extended = true;
                    used_range_facts.issue(
                        "opaque-sheet-data-child",
                        "Only row elements are qualified direct sheetData children; misplaced worksheet metadata remains protected.",
                    );
                }
                if stack.len() == 3 && stack[1] == "sheetData" && stack[2] == "row" && name != "c" {
                    used_range_facts.issue(
                        "opaque-row-child",
                        "Unknown row children cannot authorize normalized row removal.",
                    );
                }
                if current.is_some()
                    && stack.len() == 4
                    && !["v", "f", "is"].contains(&name.as_ref())
                {
                    used_range_facts.issue(
                        "opaque-cell-child",
                        "Unknown cell children require preservation without normalized removal.",
                    );
                }
                if stack.len() == 1 && !["sheetPr", "dimension"].contains(&name.as_ref()) {
                    dimension_insert.get_or_insert(start);
                }
                if name == "dimension"
                    && stack.len() == 1
                    && attributes
                        .keys()
                        .any(|key| *key != "ref" && *key != "xmlns" && !key.starts_with("xmlns:"))
                {
                    used_range_facts.issue(
                        "opaque-dimension",
                        "Extended dimension attributes remain protected.",
                    );
                }
                if name == "sheetData" && stack.len() == 1 {
                    sheet_data_start = Some(start);
                    data_open = Some(reader.buffer_position());
                    if empty {
                        sheet_data_end = Some(reader.buffer_position());
                        data_close = data_open;
                    }
                }
                if name == "col"
                    && stack.len() == 2
                    && stack[1] == "cols"
                    && attributes
                        .get("hidden")
                        .is_some_and(|s| s == "1" || s == "true")
                {
                    hidden_columns.push(json!([
                        attributes
                            .get("min")
                            .and_then(|s| s.parse::<u32>().ok())
                            .unwrap_or(0),
                        attributes
                            .get("max")
                            .and_then(|s| s.parse::<u32>().ok())
                            .unwrap_or(0)
                    ]));
                }
                if name == "dimension" && stack.len() == 1 {
                    used_range_facts.dimension(attributes.get("ref").map(|s| s.as_ref()));
                    dimension = json!({"start":start,"end":reader.buffer_position(),"attributes":attributes,"name":String::from_utf8_lossy(element.name().as_ref())});
                }
                if name == "row" && stack.len() == 2 && stack[1] == "sheetData" {
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(err)?;
                        if attribute.key.as_ref() == b"x14ac:dyDescent"
                            && reader.resolve_attribute(attribute.key).0
                                != ResolveResult::Bound(quick_xml::name::Namespace(
                                    b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
                                ))
                        {
                            used_range_facts.issue(
                                "opaque-row-metadata",
                                "Row descent uses an unqualified namespace.",
                            );
                        }
                    }
                    row_number = attributes
                        .get("r")
                        .map(|s| s.parse())
                        .transpose()
                        .map_err(err)?
                        .unwrap_or(row_number + 1);
                    next_column = 0;
                    row_start = start;
                    row_open = reader.buffer_position();
                    if row_number == 0 || row_number > 1048576 {
                        return Err(crate::coded(
                            "internal-invalid-field",
                            "Invalid row coordinate",
                        ));
                    }
                    used_range_facts.capture_axis(
                        db,
                        sheet,
                        "row",
                        &attributes,
                        Some(row_number - 1),
                    )?;
                    if attributes
                        .get("hidden")
                        .is_some_and(|s| s == "1" || s == "true")
                    {
                        hidden.push(row_number - 1);
                    }
                    if attributes
                        .get("hidden")
                        .is_some_and(|s| s == "0" || s == "false")
                    {
                        visible.push(row_number - 1);
                    }
                    if let Some(height) = attributes.get("ht") {
                        heights.push(json!({"row":row_number - 1,"height":height.parse::<f64>().map_err(err)?}));
                    }
                }
                if name == "col" && stack.len() == 2 && stack[1] == "cols"
                    || name == "mergeCell" && stack.len() == 2 && stack[1] == "mergeCells"
                    || name == "sheetFormatPr" && stack.len() == 1
                {
                    if name == "col" {
                        used_range_facts.capture_axis(db, sheet, "col", &attributes, None)?;
                    } else {
                        used_range_facts.element(&name, &attributes, None);
                    }
                }
                if name == "c" && stack.len() == 3 && stack[2] == "row" && stack[1] == "sheetData" {
                    used_range_facts.cell_attributes(&attributes);
                    let (row, col) = match attributes.get("r") {
                        Some(reference) => position(reference)?,
                        None => (row_number - 1, next_column),
                    };
                    if row + 1 != row_number {
                        return Err(crate::coded(
                            "internal-mismatch",
                            "Cell coordinate disagrees with its row",
                        ));
                    }
                    next_column = col + 1;
                    rows = rows.max(row + 1);
                    columns = columns.max(col + 1);
                    current = Some(Cell {
                        row,
                        col,
                        kind: attributes
                            .get("t")
                            .map(|s| s.to_string())
                            .unwrap_or("n".into()),
                        style: attributes
                            .get("s")
                            .map(|s| s.parse())
                            .transpose()
                            .map_err(err)?
                            .unwrap_or(0),
                        style_explicit: attributes.contains_key("s"),
                        start,
                        open: reader.buffer_position(),
                        ..Cell::default()
                    });
                } else if let Some(cell) = current.as_mut() {
                    if stack.get(4).is_some_and(|parent| parent == "v") {
                        cell.preservation_mask |= 3;
                    }
                    if stack.get(4).is_some_and(|parent| parent == "f") {
                        cell.preservation_mask |= 1;
                    }
                    if stack.len() == 4
                        && name == "v"
                        && attributes
                            .keys()
                            .any(|key| !crate::preservation::safe_text_attribute(key))
                    {
                        cell.preservation_mask |= 3;
                    }
                    if stack.len() == 4
                        && name == "f"
                        && attributes.keys().any(|key| {
                            !(if attributes.get("t").is_some_and(|kind| kind == "dataTable") {
                                ["t", "ref", "r1", "r2", "dt2D", "dtr", "del1", "del2"]
                                    .contains(key)
                            } else {
                                ["t", "si", "ref"].contains(key)
                            }) && !crate::preservation::safe_text_attribute(key)
                        })
                    {
                        cell.preservation_mask |= 1;
                    }
                    if stack.len() == 4
                        && name == "is"
                        && attributes
                            .keys()
                            .any(|key| !crate::preservation::safe_text_attribute(key))
                        || stack.get(4).is_some_and(|parent| parent == "is")
                            && (name != "t"
                                || attributes
                                    .keys()
                                    .any(|key| !crate::preservation::safe_text_attribute(key)))
                    {
                        cell.protected_value = true;
                    }
                    if stack.len() == 4 && ["v", "f", "is"].contains(&name.as_ref()) {
                        let bit = match name.as_ref() {
                            "v" => 1,
                            "f" => 2,
                            _ => 4,
                        };
                        if cell.extent_children & bit != 0 {
                            used_range_facts.issue(
                                "duplicated-cell-field",
                                "Repeated source value/formula fields remain unqualified.",
                            );
                        }
                        cell.extent_children |= bit;
                        if name == "is" && cell.kind != "inlineStr"
                            || name == "v" && cell.kind == "inlineStr"
                        {
                            used_range_facts.issue(
                                "invalid-cell-representation",
                                "Cell text representation disagrees with its OOXML type.",
                            );
                        }
                        if empty {
                            cell.spans.push((start, reader.buffer_position()));
                        } else {
                            cell.child_start = Some(start);
                        }
                    }
                    if stack.len() == 4 && name == "v" {
                        cell.value_present = true;
                    }
                    if stack.len() == 4 && name == "f" {
                        cell.formula = Some(String::new());
                        let ordered = element
                            .attributes()
                            .map(|a| {
                                let a = a.map_err(err)?;
                                Ok(format!(
                                    "{}:{}",
                                    serde_json::to_string(
                                        std::str::from_utf8(a.key.as_ref()).map_err(err)?
                                    )
                                    .map_err(err)?,
                                    serde_json::to_string(&a.unescape_value().map_err(err)?)
                                        .map_err(err)?
                                ))
                            })
                            .collect::<Result<Vec<_>>>()?;
                        cell.formula_attrs = Some(format!("{{{}}}", ordered.join(",")));
                    }
                }
                if !empty {
                    stack.push(name);
                }
            }
            Event::Text(ref text) => {
                if stack
                    .get(1)
                    .is_some_and(|ancestor| ancestor == "sheetProtection")
                    && !xml_text(text.as_ref())?.trim().is_empty()
                {
                    used_range_facts.issue(
                        "unqualified-sheet-protection",
                        "Worksheet protection text remains protected.",
                    );
                }
                if stack.len() == 2
                    && stack[1] == "dimension"
                    && !xml_text(text.as_ref())?.trim().is_empty()
                {
                    used_range_facts.issue("opaque-dimension", "Dimension text remains protected.");
                }
                if let Some(cell) = current.as_mut() {
                    cell_text(cell, &stack, &xml_text(text.as_ref())?);
                }
            }
            Event::CData(ref text) => {
                if stack
                    .get(1)
                    .is_some_and(|ancestor| ancestor == "sheetProtection")
                {
                    used_range_facts.issue(
                        "unqualified-sheet-protection",
                        "Worksheet protection CDATA remains protected.",
                    );
                }
                if stack.len() == 2 && stack[1] == "dimension" {
                    used_range_facts
                        .issue("opaque-dimension", "Dimension CDATA remains protected.");
                }
                if let Some(cell) = current.as_mut() {
                    cell_text(
                        cell,
                        &stack,
                        &text
                            .decode()
                            .map_err(err)?
                            .replace("\r\n", "\n")
                            .replace('\r', "\n"),
                    );
                }
            }
            Event::End(ref element) => {
                if bound && element.local_name().as_ref() == b"dimension" && stack.len() == 2 {
                    dimension["end"] = json!(reader.buffer_position());
                }
                if bound && element.local_name().as_ref() == b"worksheet" && stack.len() == 1 {
                    dimension_insert.get_or_insert(start);
                }
                if bound
                    && element.local_name().as_ref() == b"t"
                    && stack.get(4).is_some_and(|s| s == "is")
                    && !stack.iter().any(|s| s == "rPh")
                    && let Some(cell) = current.as_mut()
                {
                    cell.inline.push_str(&spreadsheet_text(&cell.inline_piece)?);
                    cell.inline_piece.clear();
                }
                if bound && element.local_name().as_ref() == b"sheetData" && stack.len() == 2 {
                    sheet_data_end = Some(reader.buffer_position());
                    data_close = Some(start);
                }
                if stack.len() == 5
                    && bound
                    && [b"v".as_slice(), b"f".as_slice(), b"is".as_slice()]
                        .contains(&element.local_name().as_ref())
                    && let Some(cell) = current.as_mut()
                    && let Some(begin) = cell.child_start.take()
                {
                    cell.spans.push((begin, reader.buffer_position()));
                }
            }
            Event::Comment(_) | Event::PI(_) => {
                if let Some(cell) = current.as_mut() {
                    match stack.get(4).map(|name| name.as_ref()) {
                        Some("is") => cell.protected_value = true,
                        Some("v") => cell.preservation_mask |= 3,
                        Some("f") => cell.preservation_mask |= 1,
                        _ => {}
                    }
                }
                if stack
                    .get(1)
                    .is_some_and(|ancestor| ancestor == "sheetProtection")
                {
                    used_range_facts.issue(
                        "unqualified-sheet-protection",
                        "Opaque worksheet protection markup remains protected.",
                    );
                }
                if stack.len() >= 3 && stack[1] == "sheetData" {
                    used_range_facts.issue(
                        "opaque-row-child",
                        "Opaque row or cell XML cannot authorize normalized removal.",
                    );
                }
            }
            Event::Eof => break,
            _ => (),
        }
        let ended_cell = match &event {
            Event::End(e) => bound && e.local_name().as_ref() == b"c" && stack.len() == 4,
            Event::Empty(e) => bound && e.local_name().as_ref() == b"c" && stack.len() == 3,
            _ => false,
        };
        if ended_cell && let Some(cell) = current.take() {
            if cell.formula.is_some() && cell.preservation_mask & 2 != 0 {
                protected_formula_caches += 1;
            }
            if cell.protected_value || cell.preservation_mask != 0 {
                db.prepare_cached(
                    "INSERT INTO value_features(sheet,r,c,reason,mask) VALUES(?1,?2,?3,?4,?5)",
                )
                .map_err(err)?
                .execute(params![
                    sheet,
                    cell.row,
                    cell.col,
                    crate::preservation::REASON,
                    cell.preservation_mask | u8::from(cell.protected_value)
                ])
                .map_err(err)?;
            }
            if cell
                .formula_attrs
                .as_ref()
                .and_then(|s| serde_json::from_str::<Value>(s).ok())
                .is_some_and(|attrs| {
                    matches!(attrs["t"].as_str(), Some("array" | "dataTable"))
                        && attrs.get("ref").is_some()
                })
            {
                used_range_facts.issue("unqualified-formula-footprint", "An imported array or data-table rectangle needs independent native footprint ownership before extent normalization.");
            }
            if cell.formula.as_ref().is_some_and(|f| f.is_empty())
                && cell
                    .formula_attrs
                    .as_ref()
                    .and_then(|s| serde_json::from_str::<Value>(s).ok())
                    .is_none_or(|a| a.get("t").is_none())
            {
                used_range_facts.issue("empty-ordinary-formula", "An empty ordinary formula declaration remains preserved without guessed normalization.");
            }
            used_range_facts.cell(
                cell.row,
                cell.col,
                &cell.kind,
                cell.value_present.then_some(cell.raw.as_str()),
                &cell.inline,
                cell.formula.is_some(),
                cell.style,
                cell.style_explicit,
            );
            if cell.formula.is_some() {
                formulas += 1;
                if cell
                    .formula_attrs
                    .as_ref()
                    .and_then(|a| serde_json::from_str::<Value>(a).ok())
                    .is_some_and(|a| a.get("t").is_some())
                {
                    special_formulas += 1;
                }
            }
            // Packing samples are nested inside worksheet time. SQL batches
            // have their own span; one sample no longer stands in for SQL time.
            let sample = (diagnostic && count % 1024 == 0)
                .then(|| crate::diagnostics::span("import.cell_pack_sample"));
            let spans = serde_json::to_string(&cell.spans).map_err(err)?;
            let bytes = sheet.len()
                + cell.kind.capacity()
                + cell.raw.capacity()
                + cell.inline.capacity()
                + cell.formula.as_ref().map_or(0, String::capacity)
                + cell.formula_attrs.as_ref().map_or(0, String::capacity)
                + spans.capacity()
                + CELL_COLUMNS * std::mem::size_of::<SqlValue>();
            if insert_bytes + bytes > CELL_BATCH_BYTES {
                flush_cells(db, &insert_sql, &mut insert_values)?;
                insert_bytes = 0;
            }
            insert_values.extend([
                SqlValue::Text(sheet.to_owned()),
                SqlValue::Integer(cell.row.into()),
                SqlValue::Integer(cell.col.into()),
                SqlValue::Text(cell.kind),
                if cell.value_present {
                    SqlValue::Text(cell.raw)
                } else {
                    SqlValue::Null
                },
                SqlValue::Text(cell.inline),
                cell.formula.map_or(SqlValue::Null, SqlValue::Text),
                cell.formula_attrs.map_or(SqlValue::Null, SqlValue::Text),
                SqlValue::Integer(cell.style.into()),
                SqlValue::Integer(cell.start as i64),
                SqlValue::Integer(cell.open as i64),
                SqlValue::Integer(if empty { cell.open } else { start } as i64),
                SqlValue::Integer(reader.buffer_position() as i64),
                SqlValue::Text(spans),
                SqlValue::Integer(i64::from(cell.style_explicit)),
            ]);
            insert_bytes += bytes;
            if diagnostic {
                crate::diagnostics::maximum("import.staged_value_bytes", insert_bytes as u64);
                if bytes >= CELL_BATCH_BYTES {
                    crate::diagnostics::count("import.large_cell_flushes", 1);
                }
            }
            drop(sample);
            if insert_values.len() == CELL_COLUMNS * CELL_BATCH || insert_bytes >= CELL_BATCH_BYTES
            {
                flush_cells(db, &insert_sql, &mut insert_values)?;
                insert_bytes = 0;
            }
            count += 1;
        }
        let ended_row = match &event {
            Event::End(e) => bound && e.local_name().as_ref() == b"row" && stack.len() == 3,
            Event::Empty(e) => bound && e.local_name().as_ref() == b"row" && stack.len() == 2,
            _ => false,
        };
        if ended_row {
            row_count += 1;
            insert_row
                .execute(params![
                    sheet,
                    row_number - 1,
                    row_start,
                    row_open,
                    if empty { row_open } else { start },
                    reader.buffer_position()
                ])
                .map_err(err)?;
        }
        if matches!(event, Event::End(_)) {
            stack.pop();
        }
        buffer.clear();
    }
    if !root_seen || !stack.is_empty() || current.is_some() {
        return Err(crate::coded(
            "package-malformed",
            "Unclosed or missing worksheet XML root",
        ));
    }
    flush_cells(db, &insert_sql, &mut insert_values)?;
    if sheet_data_start.is_none() {
        used_range_facts.issue(
            "missing-sheet-data",
            "A worksheet without sheetData remains unqualified.",
        );
    }
    used_range_facts.resolve_shared_strings(db, sheet)?;
    crate::diagnostics::count("import.cells", count);
    crate::diagnostics::count("import.formulas", formulas);
    crate::diagnostics::count("import.rows", row_count);
    hidden.sort_unstable();
    Ok(
        json!({"path":sheet,"rows":rows,"columns":columns,"storedCells":count,"formulaCount":formulas,"specialFormulaCount":special_formulas,"protectedFormulaCaches":protected_formula_caches,"extended":extended,"rowHeights":heights,"hiddenRows":hidden,"visibleRows":visible,"hiddenColumns":hidden_columns,"dataStart":sheet_data_start,"dataEnd":sheet_data_end,"dataOpen":data_open,"dataClose":data_close,"dimension":dimension,"dimensionInsert":dimension_insert,"usedRangeFacts":used_range_facts.value()?}),
    )
}
fn cell_text(cell: &mut Cell, stack: &[Cow<'static, str>], text: &str) {
    if stack.len() == 5 && stack[4] == "v" {
        cell.raw.push_str(text);
    }
    if stack.len() == 5
        && stack[4] == "f"
        && let Some(formula) = cell.formula.as_mut()
    {
        formula.push_str(text);
    }
    if stack.len() >= 6
        && stack[4] == "is"
        && stack.last().is_some_and(|s| s == "t")
        && !stack.iter().any(|s| s == "rPh")
    {
        cell.inline_piece.push_str(text);
    }
}

fn validate_directory(source: &Path, zip: &ZipArchive<crate::cancel::Reader<File>>) -> Result<()> {
    // ZipArchive's name map may collapse duplicates. Check the central records
    // themselves before any import so ambiguous packages cannot be rewritten.
    let mut file = File::open(source).map_err(err)?;
    file.seek(SeekFrom::Start(zip.central_directory_start()))
        .map_err(err)?;
    let mut seen = std::collections::HashSet::new();
    for _ in 0..zip.len() {
        let mut header = [0u8; 46];
        file.read_exact(&mut header).map_err(err)?;
        if header[..4] != [0x50, 0x4b, 0x01, 0x02] {
            return Err(crate::coded("package-malformed", "Invalid ZIP directory"));
        }
        let size = |offset| u16::from_le_bytes([header[offset], header[offset + 1]]) as usize;
        let mut name = vec![0; size(28)];
        file.read_exact(&mut name).map_err(err)?;
        if !seen.insert(name) {
            return Err(crate::coded("package-duplicate", "Duplicate ZIP member"));
        }
        file.seek(SeekFrom::Current((size(30) + size(32)) as i64))
            .map_err(err)?;
    }
    let mut next = [0u8; 4];
    file.read_exact(&mut next).map_err(err)?;
    if next == [0x50, 0x4b, 0x01, 0x02] {
        return Err(crate::coded("package-duplicate", "Duplicate ZIP member"));
    }
    Ok(())
}
fn worksheet_parts(
    zip: &mut ZipArchive<crate::cancel::Reader<File>>,
) -> Result<std::collections::HashSet<String>> {
    let source = BufReader::new(zip.by_name("[Content_Types].xml").map_err(err)?);
    let mut reader = NsReader::from_reader(source);
    let mut buffer = Vec::new();
    let mut sheets = std::collections::HashSet::new();
    loop {
        crate::cancel::check()?;
        let (ns, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        match event {
            Event::Start(ref e) | Event::Empty(ref e)
                if ns
                    == ResolveResult::Bound(quick_xml::name::Namespace(
                        b"http://schemas.openxmlformats.org/package/2006/content-types",
                    ))
                    && e.local_name().as_ref() == b"Override" =>
            {
                let a = attrs(e)?;
                if a.get("ContentType").is_some_and(|v| {
                    v == "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
                }) {
                    let name = a.get("PartName").ok_or_else(|| {
                        crate::coded(
                            "internal-missing-field",
                            "Worksheet content type is missing its part",
                        )
                    })?;
                    sheets.insert(name.trim_start_matches('/').to_owned());
                }
            }
            Event::DocType(_) => {
                return Err(crate::coded(
                    "package-malformed",
                    "XML entities are unsupported",
                ));
            }
            Event::Eof => break,
            _ => (),
        }
        buffer.clear();
    }
    Ok(sheets)
}

pub fn index(root: &Path) -> Result<Value> {
    index_inner(root, None, None)
}
pub fn raw_index(root: &Path, input: &Value) -> Result<Value> {
    index_inner(
        root,
        Some(crate::field(input, "sheet")?),
        Some(input["strings"].as_str()),
    )
    .map(|_| json!({"indexed":true}))
}
fn index_inner(
    root: &Path,
    selected: Option<&str>,
    strings: Option<Option<&str>>,
) -> Result<Value> {
    let source = local(root, "source.xlsx")?;
    let revision = {
        let _span = crate::diagnostics::span("import.hash");
        hash(&source)?
    };
    if crate::diagnostics::enabled() {
        crate::diagnostics::count("import.source_bytes", source.metadata().map_err(err)?.len());
    }
    if local(root, "cells.sqlite")?.exists() {
        return Err(crate::coded(
            "package-duplicate",
            "Workbook index already exists",
        ));
    }
    let schema_span = crate::diagnostics::span("import.schema");
    let db = database(root)?;
    db.execute_batch("PRAGMA journal_mode=DELETE; BEGIN; CREATE TABLE versions(id INTEGER PRIMARY KEY,parent INTEGER); INSERT INTO versions(id,parent) VALUES(0,NULL); CREATE TABLE changes(version INTEGER NOT NULL,sheet TEXT NOT NULL,r INTEGER NOT NULL,c INTEGER NOT NULL,data TEXT NOT NULL,edit TEXT NOT NULL DEFAULT '{}',PRIMARY KEY(version,sheet,r,c)) WITHOUT ROWID; CREATE INDEX changes_coordinates ON changes(sheet,r,c,version); CREATE TABLE strings(id INTEGER PRIMARY KEY,value TEXT NOT NULL); CREATE TABLE shared_value_features(id INTEGER PRIMARY KEY,reason TEXT NOT NULL); CREATE TABLE value_features(sheet TEXT NOT NULL,r INTEGER NOT NULL,c INTEGER NOT NULL,reason TEXT NOT NULL,mask INTEGER NOT NULL DEFAULT 1,PRIMARY KEY(sheet,r,c)) WITHOUT ROWID; CREATE TABLE cells(sheet TEXT NOT NULL,r INTEGER NOT NULL,c INTEGER NOT NULL,kind TEXT NOT NULL,raw TEXT,inline TEXT NOT NULL,formula TEXT,formula_attrs TEXT,style INTEGER NOT NULL,start INTEGER NOT NULL,open INTEGER NOT NULL,close INTEGER NOT NULL,end INTEGER NOT NULL,spans TEXT NOT NULL,style_explicit INTEGER NOT NULL DEFAULT 0,PRIMARY KEY(sheet,r,c)) WITHOUT ROWID; CREATE INDEX cells_formulas ON cells(sheet,r,c) WHERE formula IS NOT NULL; CREATE TABLE rows(sheet TEXT NOT NULL,r INTEGER NOT NULL,start INTEGER NOT NULL,open INTEGER NOT NULL,close INTEGER NOT NULL,end INTEGER NOT NULL,PRIMARY KEY(sheet,r)) WITHOUT ROWID;").map_err(err)?;
    crate::used_range::schema(&db)?;
    crate::used_range_transfer::bind_import(root, &db, &revision)?;
    drop(schema_span);
    let directory_span = crate::diagnostics::span("import.directory");
    let mut zip =
        ZipArchive::new(crate::cancel::Reader(File::open(&source).map_err(err)?)).map_err(err)?;
    validate_directory(&source, &zip)?;
    let worksheet_names = worksheet_parts(&mut zip)?;
    let names: Vec<String> = zip.file_names().map(str::to_owned).collect();
    if names.iter().any(|name| {
        name.starts_with('/') || name.contains('\\') || name.split('/').any(|s| s == "..")
    }) {
        return Err(crate::coded("package-unsupported", "Unsafe ZIP member"));
    }
    if names
        .iter()
        .filter(|name| *name == "xl/workbook.xml")
        .count()
        != 1
    {
        return Err(crate::coded(
            "package-unsupported",
            "Workbook part is unavailable",
        ));
    }
    drop(directory_span);
    if let Some(strings) = strings.unwrap_or(Some("xl/sharedStrings.xml"))
        && names.iter().any(|name| name == strings)
    {
        shared_strings(BufReader::new(zip.by_name(strings).map_err(err)?), &db)?;
    }
    let mut sheets = Vec::new();
    // Worksheet relationships are interpreted by the common metadata adapter;
    // all actual worksheet parts are indexed, including nonstandard filenames.
    for name in &names {
        if !worksheet_names.contains(name) || selected.is_some_and(|target| target != name) {
            continue;
        }
        let file = root.join(format!("sheet-{}.xml", sheets.len()));
        let inflate_span = crate::diagnostics::span("import.inflate");
        // Some writers put a byte-order mark before the XML declaration. It
        // carries no content, and every offset recorded for this sheet is
        // measured in the staged file, so the staged file is written without it.
        let mut member = zip.by_name(name).map_err(err)?;
        let mut staged = crate::quota::create(&file).map_err(err)?;
        let (mut lead, mut got) = ([0u8; 3], 0);
        while got < lead.len() {
            let read = member.read(&mut lead[got..]).map_err(err)?;
            if read == 0 {
                break;
            }
            got += read;
        }
        if lead[..got] != [0xEF, 0xBB, 0xBF] {
            staged.write_all(&lead[..got]).map_err(err)?;
        }
        let inflated = io::copy(&mut member, &mut staged).map_err(err)? + got as u64;
        drop(staged);
        drop(member);
        drop(inflate_span);
        crate::diagnostics::count("import.worksheet_bytes", inflated);
        let mut sheet = worksheet(&file, &db, name)?;
        sheet["artifact"] = json!(file.file_name().unwrap().to_string_lossy());
        sheets.push(sheet);
    }
    {
        let _span = crate::diagnostics::span("import.commit");
        db.execute_batch("COMMIT; PRAGMA optimize;").map_err(err)?;
    }
    if selected.is_some() {
        let result = json!({"revision":revision,"sheets":sheets,"storage":"rust-raw-index-v1","sourceFactSchema":crate::used_range::FACTS_SCHEMA});
        crate::quota::write(
            local(root, "index.json")?,
            serde_json::to_vec(&result).map_err(err)?,
        )
        .map_err(err)?;
        return Ok(result);
    }
    let _metadata_span = crate::diagnostics::span("import.metadata");
    let mut metadata =
        ZipWriter::new(crate::quota::create(local(root, "metadata.xlsx")?).map_err(err)?);
    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    for name in &names {
        if let Some(sheet) = sheets
            .iter()
            .find(|sheet| sheet["path"].as_str() == Some(name))
        {
            let mut file =
                File::open(root.join(sheet["artifact"].as_str().unwrap())).map_err(err)?;
            metadata.start_file(name, options).map_err(err)?;
            if let (Some(start), Some(end)) =
                (sheet["dataStart"].as_u64(), sheet["dataEnd"].as_u64())
            {
                io::copy(&mut Read::by_ref(&mut file).take(start), &mut metadata).map_err(err)?;
                // Use an explicit namespace so prefixed worksheets retain their meaning.
                metadata.write_all(b"<sheetData xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>").map_err(err)?;
                file.seek(SeekFrom::Start(end)).map_err(err)?;
                io::copy(&mut file, &mut metadata).map_err(err)?;
            } else {
                io::copy(&mut file, &mut metadata).map_err(err)?;
            }
        } else if name == "xl/sharedStrings.xml" {
            metadata.start_file(name, options).map_err(err)?;
            metadata
                .write_all(
                    b"<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>",
                )
                .map_err(err)?;
        } else if name.ends_with(".xml") || name.ends_with(".rels") || name.ends_with(".vml") {
            // Note drawings (.vml) are kept: a note's box records its cell, and
            // that has to move when rows or columns do.
            metadata
                .raw_copy_file(zip.by_name(name).map_err(err)?)
                .map_err(err)?;
        } else {
            metadata.start_file(name, options).map_err(err)?;
        }
    }
    metadata.finish().map_err(err)?;
    for sheet in &mut sheets {
        let background = crate::package_edit::inspect_sheet_background(
            root,
            sheet["path"].as_str().ok_or_else(|| {
                crate::coded(
                    "internal-missing-field",
                    "Missing background worksheet path",
                )
            })?,
        );
        if !background.is_null() {
            sheet["background"] = background;
        }
    }
    let result = json!({"revision":revision,"sheets":sheets,"storage":"rust-sqlite-v1","sourceFactSchema":crate::used_range::FACTS_SCHEMA});
    crate::quota::write(
        local(root, "index.json")?,
        serde_json::to_vec(&result).map_err(err)?,
    )
    .map_err(err)?;
    Ok(result)
}
