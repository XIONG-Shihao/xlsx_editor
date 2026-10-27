//! Source-preserving duplicate record compaction; the accepted mapping is recomputed.
//! Only selected cell spans change. Physical row tags, outside cells, worksheet
//! metadata, header and unrelated ZIP members retain their original bytes.
use super::{Config, config, prepare};
use crate::xml::{NS, address, attrs, position, xml_text};
use crate::{Result, database, err, field, local};
use quick_xml::{
    NsReader,
    events::Event,
    name::{Namespace, ResolveResult},
};
use rusqlite::{Connection, OptionalExtension, params};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use std::{
    collections::{BTreeMap, BTreeSet},
    fs::{self, File},
    io::{BufReader, Read, Seek, SeekFrom, Write},
    path::{Path, PathBuf},
};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

const CELL_LIMIT: u64 = 1_048_576;
const FILTER_LIMIT: u64 = 16 * 1024 * 1024;
const OWN_FILTER: &str = workbook_format!(URN_OWNED_FILTER_V1);

struct Scratch {
    worksheet: PathBuf,
    partial: PathBuf,
}
impl Drop for Scratch {
    fn drop(&mut self) {
        let _ = fs::remove_file(&self.worksheet);
        let _ = fs::remove_file(&self.partial);
    }
}
fn sha(bytes: &[u8]) -> String {
    format!("{:x}", Sha256::digest(bytes))
}
fn inside(c: &Config, row: u32, col: u32) -> bool {
    row >= c.body && row <= c.last.0 && col >= c.first.1 && col <= c.last.1
}
fn plain(a: &BTreeMap<String, String>) -> BTreeMap<String, String> {
    a.iter()
        .filter(|(k, _)| *k != "xmlns" && !k.starts_with("xmlns:"))
        .map(|(k, v)| (k.clone(), v.clone()))
        .collect()
}
fn row_number(a: &BTreeMap<String, String>) -> Result<u32> {
    a.get("r")
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing physical row"))?
        .parse::<u32>()
        .map_err(err)?
        .checked_sub(1)
        .filter(|r| *r < 1048576)
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid physical row"))
}
fn fragment(file: &mut File, first: u64, last: u64, limit: u64) -> Result<Vec<u8>> {
    let length = last
        .checked_sub(first)
        .filter(|n| *n <= limit)
        .ok_or_else(|| crate::coded("duplicates-limit", "XML fragment exceeds its bound"))?;
    file.seek(SeekFrom::Start(first)).map_err(err)?;
    let mut bytes = vec![0; length as usize];
    file.read_exact(&mut bytes).map_err(err)?;
    Ok(bytes)
}
fn copy_span(file: &mut File, first: u64, last: u64, output: &mut impl Write) -> Result<()> {
    let mut left = last
        .checked_sub(first)
        .ok_or_else(|| crate::coded("duplicates-conflict", "Overlapping duplicate cell spans"))?;
    file.seek(SeekFrom::Start(first)).map_err(err)?;
    let mut buffer = [0u8; 65536];
    while left > 0 {
        crate::cancel::check()?;
        let size = left.min(buffer.len() as u64) as usize;
        let count = file.read(&mut buffer[..size]).map_err(err)?;
        if count == 0 {
            return Err(crate::coded("duplicates-malformed", "Truncated source XML"));
        }
        output.write_all(&buffer[..count]).map_err(err)?;
        left -= count as u64;
    }
    Ok(())
}
fn source_target(target: &str) -> Result<String> {
    if target.contains(['\\', ':', '?', '#']) {
        return Err(crate::coded(
            "duplicates-unsupported",
            "External/opaque workbook relationship",
        ));
    }
    let full = if target.starts_with('/') {
        target[1..].to_owned()
    } else {
        format!("xl/{target}")
    };
    let mut result = Vec::new();
    for segment in full.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                result.pop().ok_or_else(|| {
                    crate::coded("duplicates-malformed", "Escaping workbook relationship")
                })?;
            }
            _ => result.push(segment),
        }
    }
    Ok(result.join("/"))
}

/// The resolver projection is checked against the supplied, version-proved ZIP,
/// rather than accepting a renderer view or a mutable projected range inventory.
fn source_workbook(
    zip: &mut ZipArchive<crate::cancel::Reader<File>>,
    metadata: &Value,
) -> Result<()> {
    let mut sheets = Vec::new();
    let mut ids = BTreeSet::new();
    let mut names = BTreeSet::new();
    let mut reader =
        NsReader::from_reader(BufReader::new(zip.by_name("xl/workbook.xml").map_err(err)?));
    let mut buffer = Vec::new();
    let mut depth = 0;
    loop {
        crate::cancel::check()?;
        let (ns, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        let main = ns == ResolveResult::Bound(Namespace(NS));
        match &event {
            Event::Start(e) | Event::Empty(e) => {
                let empty = matches!(event, Event::Empty(_));
                let name = e.local_name();
                if main
                    && depth == 1
                    && [
                        b"workbookProtection".as_slice(),
                        b"externalReferences",
                        b"pivotCaches",
                    ]
                    .contains(&name.as_ref())
                {
                    return Err(crate::coded(
                        "duplicates-protected",
                        "Protected/dependent workbook graph is not qualified for duplicate compaction",
                    ));
                }
                if main && depth == 2 && name.as_ref() == b"definedName" {
                    return Err(crate::coded(
                        "duplicates-unsupported",
                        "Named dependency graphs require duplicate-compaction qualification",
                    ));
                }
                if main && depth == 2 && name.as_ref() == b"sheet" {
                    let a = attrs(e)?;
                    let name = a
                        .get("name")
                        .ok_or_else(|| {
                            crate::coded("internal-missing-field", "Missing source worksheet name")
                        })?
                        .clone();
                    let id = a
                        .get("r:id")
                        .ok_or_else(|| {
                            crate::coded("internal-missing-field", "Missing source worksheet edge")
                        })?
                        .clone();
                    if !ids.insert(id.clone()) || !names.insert(name.to_ascii_lowercase()) {
                        return Err(crate::coded(
                            "duplicates-ambiguous",
                            "Ambiguous source worksheet inventory",
                        ));
                    }
                    sheets.push((name, id));
                }
                if !empty {
                    depth += 1
                }
            }
            Event::End(_) => depth -= 1,
            Event::DocType(_) => {
                return Err(crate::coded(
                    "duplicates-malformed",
                    "Source XML entities are protected",
                ));
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    drop(reader);
    let mut edges = BTreeMap::new();
    let mut reader = NsReader::from_reader(BufReader::new(
        zip.by_name("xl/_rels/workbook.xml.rels").map_err(err)?,
    ));
    let mut buffer = Vec::new();
    loop {
        let (ns, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        match &event {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"Relationship" => {
                if ns
                    != ResolveResult::Bound(Namespace(
                        b"http://schemas.openxmlformats.org/package/2006/relationships",
                    ))
                {
                    return Err(crate::coded(
                        "duplicates-unsupported",
                        "Foreign source relationship",
                    ));
                }
                let a = attrs(e)?;
                let id = a
                    .get("Id")
                    .ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing relationship identity")
                    })?
                    .clone();
                if edges.contains_key(&id) {
                    return Err(crate::coded(
                        "duplicates-duplicate",
                        "Duplicate workbook relationship identity",
                    ));
                }
                if a.get("Type").map(String::as_str)
                    == Some(
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                    )
                {
                    if a.get("TargetMode").is_some_and(|m| m != "Internal") {
                        return Err(crate::coded(
                            "duplicates-protected",
                            "External worksheets are protected",
                        ));
                    }
                    edges.insert(
                        id,
                        source_target(a.get("Target").ok_or_else(|| {
                            crate::coded("internal-missing-field", "Missing worksheet target")
                        })?)?,
                    );
                } else {
                    edges.insert(id, String::new());
                }
            }
            Event::DocType(_) => {
                return Err(crate::coded(
                    "duplicates-malformed",
                    "Relationship entities are protected",
                ));
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    let actual = sheets
        .into_iter()
        .map(|(name, id)| {
            let path = edges.get(&id).filter(|p| !p.is_empty()).ok_or_else(|| {
                crate::coded(
                    "duplicates-bad-input",
                    "Source sheet has no ordinary worksheet edge",
                )
            })?;
            Ok(json!({"name":name,"path":path}))
        })
        .collect::<Result<Vec<_>>>()?;
    if metadata["sheets"] != json!(actual)
        || metadata["namedRanges"]
            .as_array()
            .is_none_or(|v| !v.is_empty())
        || metadata["tables"].as_array().is_none_or(|v| !v.is_empty())
    {
        return Err(crate::coded(
            "duplicates-duplicate",
            "Duplicate compaction source metadata is unqualified or inconsistent",
        ));
    }
    Ok(())
}

/// Visit the actual accepted formula inventory using the native SQL snapshot.
/// sorted_formula supplies qualified lexical rewrites; a separate range check
/// forbids a moved formula from following a nonmoving same-row outside cell.
fn formulas(db: &Connection, c: &Config, sheet: &str, metadata: &Value) -> Result<()> {
    crate::calculation::snapshot(db, c.version)?;
    let target = metadata["sheets"]
        .as_array()
        .unwrap()
        .iter()
        .find(|s| s["path"] == sheet)
        .ok_or_else(|| crate::coded("internal-unknown-reference", "Unknown duplicate worksheet"))?;
    let target_name = field(target, "name")?;
    // The first row of the range whose content moves or goes.
    let first_changed: Option<u32> = db
        .query_row(
            "SELECT MIN(source) FROM duplicate_rows WHERE destination IS NULL OR destination<>source",
            [],
            |r| r.get(0),
        )
        .map_err(err)?;
    for owner in metadata["sheets"].as_array().unwrap() {
        let path = field(owner, "path")?;
        let name = field(owner, "name")?;
        let mut after = None;
        loop {
            crate::cancel::check()?;
            let page = crate::calculation::page(db, path, (0, 0), (1048575, 16383), after, true)?;
            for (reference, cell) in &page {
                if !cell["formulaType"].is_null() || cell["scenarioResult"] == true {
                    return Err(crate::coded(
                        "duplicates-protected",
                        "Special/scenario formula graphs are protected",
                    ));
                }
                let (row, col) = position(reference)?;
                let selected = path == sheet && inside(c, row, col);
                let destination = if selected {
                    db.prepare_cached("SELECT destination FROM duplicate_rows WHERE source=?1")
                        .map_err(crate::err)?
                        .query_row([row], |r| r.get::<_, Option<u32>>(0))
                        .map_err(err)?
                        .or(Some(row))
                } else {
                    None
                };
                let old = field(cell, "formula")?;
                let rewritten = crate::formula::sorted_formula(
                    old,
                    name,
                    target_name,
                    row,
                    (c.body, c.last.0),
                    destination,
                )?;
                let deps = crate::formula::cell_dependencies(
                    metadata,
                    name,
                    reference,
                    &json!({"formula":old}),
                )?;
                if crate::formula::auditing::references(
                    metadata,
                    name,
                    reference,
                    &json!({"formula":old}),
                    false,
                )?
                .unknown
                {
                    return Err(crate::coded(
                        "duplicates-protected",
                        "Incomplete/dynamic duplicate dependency graph is protected",
                    ));
                }
                for dep in deps["ranges"].as_array().ok_or_else(|| {
                    crate::coded(
                        "internal-missing-field",
                        "Missing formula dependency ranges",
                    )
                })? {
                    if dep["sheet"].as_str() != Some(target_name) {
                        continue;
                    }
                    let (a, b) = dependency_bounds(dep)?;
                    if selected && a.0 <= row && b.0 >= row && (a.1 < c.first.1 || b.1 > c.last.1) {
                        return Err(crate::coded(
                            "duplicates-unsupported",
                            "A compacted formula reads nonmoving cells outside its record",
                        ));
                    }
                    // A formula outside the range keeps its text, as after a sort, but
                    // here cells it reads move or go: its result would be wrong.
                    if !selected
                        && let Some(changed) = first_changed
                        && b.0 >= changed
                        && a.0 <= c.last.0
                        && b.1 >= c.first.1
                        && a.1 <= c.last.1
                    {
                        return Err(crate::coded(
                            "duplicates-protected",
                            "A formula outside the range reads rows that duplicate removal moves",
                        ));
                    }
                }
                db.execute(
                    "INSERT INTO dedup_formulas VALUES(?1,?2,?3,?4)",
                    params![path, reference, old, rewritten],
                )
                .map_err(err)?;
            }
            if page.len() < 512 {
                break;
            }
            after = Some(position(&page.last().unwrap().0)?);
        }
    }
    Ok(())
}

/// A dependency's first and last cell, as (row, column).
fn dependency_bounds(dep: &Value) -> Result<((u32, u32), (u32, u32))> {
    if let Some(r) = dep["ref"].as_str() {
        let p = position(r)?;
        return Ok((p, p));
    }
    let number = |key: &str, what: &str| -> Result<u32> {
        dep[key]
            .as_u64()
            .map(|n| n as u32)
            .ok_or_else(|| crate::coded("internal-invalid-field", what))
    };
    Ok((
        (
            number("startRow", "Invalid dependency row")?,
            number("startColumn", "Invalid dependency column")?,
        ),
        (
            number("endRow", "Invalid dependency row")?,
            number("endColumn", "Invalid dependency column")?,
        ),
    ))
}

struct Row {
    r: u32,
    start: u64,
    open: u64,
    name: String,
    a: BTreeMap<String, String>,
    insert: Option<u64>,
    previous: Option<u32>,
}
struct Cell {
    r: u32,
    c: u32,
    start: u64,
    depth: usize,
    formula_seen: bool,
    payload: Option<String>,
    children: BTreeSet<String>,
}
struct Scan {
    sheet_close: u64,
    row_name: String,
    filter: Option<Vec<u8>>,
    format: BTreeMap<String, String>,
    namespaces: BTreeMap<String, String>,
}

fn store_row(db: &Connection, row: Row, close: u64, end: u64, empty: bool) -> Result<()> {
    db.execute(
        "INSERT INTO dedup_physical_rows VALUES(?1,?2,?3,?4,?5,?6,?7,?8,?9)",
        params![
            row.r,
            row.start,
            row.open,
            close,
            end,
            empty,
            row.name,
            serde_json::to_string(&row.a).map_err(err)?,
            row.insert.unwrap_or(close)
        ],
    )
    .map_err(err)?;
    Ok(())
}
fn scan(db: &Connection, file: &Path, c: &Config) -> Result<Scan> {
    let mut raw = File::open(file).map_err(err)?;
    let mut reader = NsReader::from_reader(BufReader::new(File::open(file).map_err(err)?));
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut row: Option<Row> = None;
    let mut cell: Option<Cell> = None;
    let mut filter_start = None;
    let mut filter = None;
    let mut format = None;
    let mut close = None;
    let mut row_name = None;
    let mut previous_row = None;
    let mut namespaces = BTreeMap::new();
    loop {
        crate::cancel::check()?;
        let before = reader.buffer_position();
        let (ns, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        let main = ns == ResolveResult::Bound(Namespace(NS));
        let end = reader.buffer_position();
        match &event {
            Event::Start(e) | Event::Empty(e) => {
                let empty = matches!(event, Event::Empty(_));
                let name = e.local_name();
                if depth == 0 {
                    if !main || name.as_ref() != b"worksheet" {
                        return Err(crate::coded(
                            "duplicates-unsupported",
                            "Unsupported source worksheet root",
                        ));
                    }
                    namespaces = attrs(e)?
                        .into_iter()
                        .filter(|(k, _)| k == "xmlns" || k.starts_with("xmlns:"))
                        .collect();
                    let q = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    row_name = Some(
                        q.rsplit_once(':')
                            .map(|(p, _)| format!("{p}:row"))
                            .unwrap_or("row".into()),
                    );
                }
                if main && depth == 1 {
                    match name.as_ref() {
                        b"sheetData" => {
                            if empty {
                                close = Some(before)
                            }
                        }
                        b"sheetFormatPr" => {
                            if format.is_some() {
                                return Err(crate::coded(
                                    "duplicates-ambiguous",
                                    "Ambiguous worksheet defaults",
                                ));
                            }
                            format = Some(plain(&attrs(e)?));
                        }
                        b"autoFilter" => {
                            if filter.is_some() || filter_start.is_some() {
                                return Err(crate::coded(
                                    "duplicates-protected",
                                    "Multiple source filters are protected",
                                ));
                            }
                            if empty {
                                filter = Some(fragment(&mut raw, before, end, FILTER_LIMIT)?)
                            } else {
                                filter_start = Some(before)
                            }
                        }
                        b"sheetPr" => {
                            if attrs(e)?
                                .get("filterMode")
                                .is_some_and(|v| v != "0" && v != "false")
                            {
                                return Err(crate::coded(
                                    "duplicates-protected",
                                    "Active/imported filter visibility is protected",
                                ));
                            }
                        }
                        b"dimension" | b"sheetViews" | b"cols" | b"printOptions"
                        | b"pageMargins" | b"pageSetup" | b"headerFooter" | b"rowBreaks"
                        | b"colBreaks" => {}
                        _ => {
                            return Err(crate::coded(
                                "duplicates-protected",
                                "Worksheet dependency/protection/object metadata requires compaction qualification",
                            ));
                        }
                    }
                }
                if !main && depth == 1 {
                    return Err(crate::coded(
                        "duplicates-protected",
                        "Foreign worksheet metadata is protected",
                    ));
                }
                if main && depth == 2 && name.as_ref() == b"row" {
                    let a = attrs(e)?;
                    let r = row_number(&a)?;
                    if previous_row.is_some_and(|old| r <= old) {
                        return Err(crate::coded(
                            "duplicates-bad-input",
                            "Source physical rows are not uniquely ordered",
                        ));
                    }
                    previous_row = Some(r);
                    let new = Row {
                        r,
                        start: before,
                        open: end,
                        name: String::from_utf8_lossy(e.name().as_ref()).into_owned(),
                        a,
                        insert: None,
                        previous: None,
                    };
                    if empty {
                        store_row(db, new, before, end, true)?
                    } else {
                        if row.is_some() {
                            return Err(crate::coded(
                                "duplicates-unsupported",
                                "Nested worksheet row",
                            ));
                        }
                        row = Some(new)
                    }
                }
                if main && depth == 3 && name.as_ref() == b"c" {
                    let current = row.as_mut().ok_or_else(|| {
                        crate::coded("duplicates-unsupported", "Cell outside physical row")
                    })?;
                    let a = attrs(e)?;
                    let (r, col) = position(a.get("r").ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing source cell address")
                    })?)?;
                    if r != current.r || current.previous.is_some_and(|p| col <= p) {
                        return Err(crate::coded(
                            "duplicates-bad-input",
                            "Source cells are not uniquely ordered in their physical row",
                        ));
                    }
                    current.previous = Some(col);
                    if r >= c.body && r <= c.last.0 && col >= c.first.1 && current.insert.is_none()
                    {
                        current.insert = Some(before)
                    }
                    if inside(c, r, col) {
                        if a.keys().any(|k| {
                            !["r", "s", "t", "xmlns"].contains(&k.as_str())
                                && !k.starts_with("xmlns:")
                        }) {
                            return Err(crate::coded(
                                "duplicates-unsupported",
                                "Special source cell metadata cannot be compacted",
                            ));
                        }
                        if a.get("t").is_some_and(|t| {
                            !["n", "b", "e", "str", "s", "inlineStr"].contains(&t.as_str())
                        }) {
                            return Err(crate::coded(
                                "duplicates-unsupported",
                                "Unsupported source cell type",
                            ));
                        }
                        if empty {
                            let bytes = fragment(&mut raw, before, end, CELL_LIMIT)?;
                            db.execute(
                                "INSERT INTO dedup_source_cells VALUES(?1,?2,?3,?4,?5)",
                                params![r, col, before, end, bytes],
                            )
                            .map_err(err)?;
                        } else {
                            cell = Some(Cell {
                                r,
                                c: col,
                                start: before,
                                depth,
                                formula_seen: false,
                                payload: None,
                                children: BTreeSet::new(),
                            });
                        }
                    }
                }
                if depth == 3 && row.is_some() && (!main || name.as_ref() != b"c") {
                    return Err(crate::coded(
                        "internal-unknown-reference",
                        "Unknown physical-row sidecars require compaction qualification",
                    ));
                }
                if cell.is_some() && !main {
                    return Err(crate::coded(
                        "duplicates-protected",
                        "Foreign selected-cell payload is protected",
                    ));
                }
                if let Some(selected) = cell.as_mut() {
                    if depth == selected.depth + 1 {
                        if ![b"f".as_slice(), b"v", b"is"].contains(&name.as_ref()) {
                            return Err(crate::coded(
                                "internal-unknown-reference",
                                "Unknown selected-cell payload is protected",
                            ));
                        }
                        let child = String::from_utf8_lossy(name.as_ref()).into_owned();
                        if !selected.children.insert(child.clone()) {
                            return Err(crate::coded(
                                "duplicates-protected",
                                "Repeated selected-cell payload is protected",
                            ));
                        }
                        selected.payload = Some(child);
                        if name.as_ref() == b"f" {
                            if selected.formula_seen || !plain(&attrs(e)?).is_empty() || empty {
                                return Err(crate::coded(
                                    "duplicates-protected",
                                    "Special/multiple source formulas are protected",
                                ));
                            }
                            selected.formula_seen = true;
                        }
                    } else if depth > selected.depth + 1
                        && !(depth == selected.depth + 2
                            && selected.payload.as_deref() == Some("is")
                            && name.as_ref() == b"t")
                    {
                        return Err(crate::coded(
                            "duplicates-unsupported",
                            "Rich/unknown selected-cell subgraphs require compaction qualification",
                        ));
                    }
                }
                if !empty {
                    depth += 1
                }
            }
            Event::End(e) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    crate::coded("duplicates-malformed", "Malformed worksheet XML")
                })?;
                if cell.as_ref().is_some_and(|p| p.depth == depth) {
                    let p = cell.take().unwrap();
                    let bytes = fragment(&mut raw, p.start, end, CELL_LIMIT)?;
                    db.execute(
                        "INSERT INTO dedup_source_cells VALUES(?1,?2,?3,?4,?5)",
                        params![p.r, p.c, p.start, end, bytes],
                    )
                    .map_err(err)?;
                }
                if main && depth == 2 && e.local_name().as_ref() == b"row" {
                    store_row(
                        db,
                        row.take().ok_or_else(|| {
                            crate::coded("duplicates-malformed", "Unexpected row close")
                        })?,
                        before,
                        end,
                        false,
                    )?
                }
                if main && depth == 1 && e.local_name().as_ref() == b"sheetData" {
                    if close.replace(before).is_some() {
                        return Err(crate::coded(
                            "duplicates-refused",
                            "Multiple worksheet data containers",
                        ));
                    }
                }
                if main && depth == 1 && e.local_name().as_ref() == b"autoFilter" {
                    filter = Some(fragment(
                        &mut raw,
                        filter_start.take().ok_or_else(|| {
                            crate::coded("duplicates-malformed", "Unexpected filter close")
                        })?,
                        end,
                        FILTER_LIMIT,
                    )?)
                }
            }
            Event::DocType(_) => {
                return Err(crate::coded(
                    "duplicates-malformed",
                    "Worksheet XML entities are protected",
                ));
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if depth != 0 || row.is_some() || cell.is_some() {
        return Err(crate::coded(
            "duplicates-unsupported",
            "Incomplete worksheet XML",
        ));
    }
    Ok(Scan {
        sheet_close: close.ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing source worksheet data")
        })?,
        row_name: row_name
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing worksheet namespace"))?,
        filter,
        format: format.unwrap_or_default(),
        namespaces,
    })
}

/// Bare dropdowns and complete inactive owned records are allowed. Active or
/// imported criteria refuse. The owned baseline must reproduce current source
/// physical-row visibility, so arbitrary revision hashes confer no authority.
fn inactive_filter(db: &Connection, scan: &Scan) -> Result<()> {
    let Some(original) = &scan.filter else {
        return Ok(());
    };
    // Add only inherited namespace bindings to the parsing copy; the original
    // filter fragment remains byte-for-byte untouched in the output sheet.
    let mut first = quick_xml::Reader::from_reader(original.as_slice());
    let event = first.read_event().map_err(err)?;
    let mut normalized = quick_xml::Writer::new(Vec::new());
    match &event {
        Event::Start(e) | Event::Empty(e) => {
            let own = attrs(e)?;
            let mut tag = e.clone().into_owned();
            for (k, v) in &scan.namespaces {
                if !own.contains_key(k) {
                    tag.push_attribute((k.as_str(), v.as_str()))
                }
            }
            normalized
                .write_event(if matches!(event, Event::Empty(_)) {
                    Event::Empty(tag)
                } else {
                    Event::Start(tag)
                })
                .map_err(err)?;
        }
        _ => {
            return Err(crate::coded(
                "duplicates-malformed",
                "Invalid source filter root",
            ));
        }
    }
    let mut bytes = normalized.into_inner();
    bytes.extend_from_slice(&original[first.buffer_position() as usize..]);
    let mut reader = NsReader::from_reader(bytes.as_slice());
    let mut buffer = Vec::new();
    let mut path = Vec::<String>::new();
    let mut reference = None;
    let mut revision = None;
    let mut text = String::new();
    let mut ext_count = 0;
    loop {
        let (ns, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        let main = ns == ResolveResult::Bound(Namespace(NS));
        let owned = ns == ResolveResult::Bound(Namespace(OWN_FILTER.as_bytes()))
            || ns
                == ResolveResult::Bound(Namespace(
                    workbook_format!(URN_OWNED_FILTER_V1).as_bytes(),
                ));
        match &event {
            Event::Start(e) | Event::Empty(e) => {
                let name = String::from_utf8_lossy(e.local_name().as_ref()).into_owned();
                let a = plain(&attrs(e)?);
                let empty = matches!(event, Event::Empty(_));
                match (path.len(), name.as_str()) {
                    (0, "autoFilter") if main && a.len() == 1 && a.contains_key("ref") => {
                        reference = a.get("ref").cloned();
                    }
                    (1, "extLst") if main && a.is_empty() => ext_count += 1,
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
                            "duplicates-protected",
                            "Active/imported/unknown filtering is protected during compaction",
                        ));
                    }
                }
                if !empty {
                    path.push(name)
                }
            }
            Event::End(_) => {
                path.pop();
            }
            Event::Text(t) => {
                let v = xml_text(t.as_ref())?;
                if path.last().is_some_and(|p| p == "filterState") {
                    text.push_str(&v)
                } else if !v.trim().is_empty() {
                    return Err(crate::coded(
                        "duplicates-malformed",
                        "Mixed filter metadata text",
                    ));
                }
            }
            Event::CData(t) if path.last().is_some_and(|p| p == "filterState") => {
                text.push_str(&t.decode().map_err(err)?)
            }
            Event::Comment(_) => {}
            Event::Eof => break,
            _ => {
                return Err(crate::coded(
                    "duplicates-unsupported",
                    "Unsupported filter metadata",
                ));
            }
        }
        buffer.clear();
    }
    if revision.is_none() {
        if ext_count != 0 || !text.is_empty() {
            return Err(crate::coded(
                "internal-unknown-reference",
                "Unknown filter extension",
            ));
        }
        return Ok(());
    }
    if ext_count != 1 || revision.as_deref() != Some(sha(text.as_bytes()).as_str()) {
        return Err(crate::coded(
            "duplicates-bad-input",
            "Inactive filter revision is not proved",
        ));
    }
    let model: Value = serde_json::from_str(&text).map_err(err)?;
    let object = model
        .as_object()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid inactive filter model"))?;
    if object.len() != 5
        || !["schema", "range", "column", "values", "baseline"]
            .iter()
            .all(|k| object.contains_key(*k))
        || model["schema"] != 1
        || model["values"].as_array().is_none_or(|v| !v.is_empty())
        || model["range"].as_str() != reference.as_deref()
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown or active owned filter state",
        ));
    }
    let (first, last) = crate::bulk::bounds(field(&model, "range")?)?;
    let column = position(&format!("{}1", field(&model, "column")?))?.1;
    if first.0 >= last.0 || column < first.1 || column > last.1 {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid inactive filter range",
        ));
    }
    if model["baseline"]
        .as_object()
        .is_none_or(|m| m.len() != 2 || !m.contains_key("format") || !m.contains_key("rows"))
        || model["baseline"]["format"] != json!(scan.format)
        || scan
            .format
            .get("zeroHeight")
            .is_some_and(|v| !["0", "1", "true", "false"].contains(&v.as_str()))
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Inactive filter source defaults differ",
        ));
    }
    let mut statement = db
        .prepare_cached("SELECT r,attrs FROM dedup_physical_rows WHERE r>?1 AND r<=?2 ORDER BY r")
        .map_err(err)?;
    let mut records = statement.query(params![first.0, last.0]).map_err(err)?;
    let mut observed = Vec::new();
    while let Some(row) = records.next().map_err(err)? {
        let r: u32 = row.get(0).map_err(err)?;
        let a: BTreeMap<String, String> =
            serde_json::from_str(&row.get::<_, String>(1).map_err(err)?).map_err(err)?;
        if a.get("hidden")
            .is_some_and(|v| !["0", "1", "true", "false"].contains(&v.as_str()))
        {
            return Err(crate::coded(
                "internal-unknown-reference",
                "Unknown inactive filter row visibility",
            ));
        }
        observed.push(json!([r, a.get("hidden")]));
    }
    if model["baseline"]["rows"] != json!(observed) {
        return Err(crate::coded(
            "duplicates-bad-input",
            "Inactive filter lexical visibility is not the current source baseline",
        ));
    }
    Ok(())
}

/// Locate a raw attribute value without reserializing any other attribute.
fn attribute_span(open: &[u8], name: &[u8]) -> Result<(usize, usize)> {
    let mut p = 1;
    while p < open.len() && !open[p].is_ascii_whitespace() && !matches!(open[p], b'>' | b'/') {
        p += 1
    }
    while p < open.len() {
        while p < open.len() && open[p].is_ascii_whitespace() {
            p += 1
        }
        if p == open.len() || matches!(open[p], b'>' | b'/') {
            break;
        }
        let start = p;
        while p < open.len() && !open[p].is_ascii_whitespace() && open[p] != b'=' {
            p += 1
        }
        let key = &open[start..p];
        while p < open.len() && open[p].is_ascii_whitespace() {
            p += 1
        }
        if open.get(p) != Some(&b'=') {
            return Err(crate::coded(
                "duplicates-malformed",
                "Malformed raw cell attribute",
            ));
        }
        p += 1;
        while p < open.len() && open[p].is_ascii_whitespace() {
            p += 1
        }
        let quote = *open
            .get(p)
            .filter(|q| matches!(q, b'\'' | b'"'))
            .ok_or_else(|| crate::coded("duplicates-malformed", "Malformed attribute quote"))?;
        p += 1;
        let value = p;
        while p < open.len() && open[p] != quote {
            p += 1
        }
        if p == open.len() {
            return Err(crate::coded(
                "duplicates-malformed",
                "Unclosed raw cell attribute",
            ));
        }
        if key == name {
            return Ok((value, p));
        }
        p += 1;
    }
    Err(crate::coded(
        "internal-missing-field",
        "Missing raw cell address attribute",
    ))
}
fn rewrite_cell(
    raw: &[u8],
    from: u32,
    to: u32,
    col: u32,
    formula: Option<(&str, &str)>,
) -> Result<Vec<u8>> {
    let mut reader = NsReader::from_reader(raw);
    let mut buffer = Vec::new();
    let mut root_open = None;
    let mut f_open = None;
    let mut f_close = None;
    let mut depth = 0usize;
    let mut actual_formula = String::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (_, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        let after = reader.buffer_position() as usize;
        match &event {
            Event::Start(e) | Event::Empty(e) => {
                let empty = matches!(event, Event::Empty(_));
                if depth == 0 {
                    root_open = Some(after)
                }
                if depth == 1 && e.local_name().as_ref() == b"f" {
                    if f_open.is_some() || plain(&attrs(e)?).len() != 0 || empty {
                        return Err(crate::coded(
                            "duplicates-unsupported",
                            "Special formula metadata cannot be compacted",
                        ));
                    }
                    f_open = Some(after)
                }
                if !empty {
                    depth += 1
                }
            }
            Event::End(e) => {
                depth -= 1;
                if depth == 1 && e.local_name().as_ref() == b"f" {
                    f_close = Some(before)
                }
            }
            Event::Text(t) if f_open.is_some() && f_close.is_none() => {
                actual_formula.push_str(&xml_text(t.as_ref())?)
            }
            Event::CData(t) if f_open.is_some() && f_close.is_none() => {
                actual_formula.push_str(&t.decode().map_err(err)?)
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    let address_span = attribute_span(
        &raw[..root_open
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing raw source cell"))?],
        b"r",
    )?;
    let mut patches = Vec::new();
    if from != to {
        patches.push((
            address_span.0,
            address_span.1,
            address(to, col).into_bytes(),
        ));
    }
    match (formula, f_open, f_close) {
        (Some((old, new)), Some(first), Some(last)) => {
            if old != actual_formula {
                return Err(crate::coded(
                    "internal-mismatch",
                    "Source formula differs from the accepted native formula",
                ));
            }
            if old != new {
                patches.push((
                    first,
                    last,
                    quick_xml::escape::escape(new).into_owned().into_bytes(),
                ))
            }
        }
        (None, None, None) => {}
        _ => {
            return Err(crate::coded(
                "internal-mismatch",
                "Source and accepted formula inventories disagree",
            ));
        }
    }
    patches.sort_by_key(|p| p.0);
    let mut output = Vec::with_capacity(raw.len() + 64);
    let mut cursor = 0;
    for (first, last, value) in patches {
        if first < cursor || last > raw.len() {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid raw source cell patch",
            ));
        }
        output.extend_from_slice(&raw[cursor..first]);
        output.extend_from_slice(&value);
        cursor = last;
    }
    output.extend_from_slice(&raw[cursor..]);
    Ok(output)
}
fn row_attrs(db: &Connection, row: u32) -> Result<BTreeMap<String, String>> {
    let value: Option<String> = db
        .query_row(
            "SELECT attrs FROM dedup_physical_rows WHERE r=?1",
            [row],
            |r| r.get(0),
        )
        .optional()
        .map_err(err)?;
    value
        .map(|s| serde_json::from_str(&s).map_err(err))
        .transpose()
        .map(|v| v.unwrap_or_default())
}
fn spool_outputs(db: &Connection, c: &Config, sheet: &str) -> Result<()> {
    let mut statement=db.prepare_cached("SELECT c.r,c.c,c.raw,m.destination FROM dedup_source_cells c JOIN duplicate_rows m ON m.source=c.r WHERE m.destination IS NOT NULL ORDER BY m.destination,c.c").map_err(err)?;
    let mut cells = statement.query([]).map_err(err)?;
    while let Some(cell) = cells.next().map_err(err)? {
        crate::cancel::check()?;
        let from: u32 = cell.get(0).map_err(err)?;
        let col: u32 = cell.get(1).map_err(err)?;
        let raw: Vec<u8> = cell.get(2).map_err(err)?;
        let to: u32 = cell.get(3).map_err(err)?;
        if !inside(c, to, col) {
            return Err(crate::coded(
                "duplicates-refused",
                "Kept duplicate destination escaped its rectangle",
            ));
        }
        let source = row_attrs(db, from)?;
        let destination = row_attrs(db, to)?;
        let mut reader = quick_xml::Reader::from_reader(raw.as_slice());
        let event = reader.read_event().map_err(err)?;
        let a = match event {
            Event::Start(e) | Event::Empty(e) => attrs(&e)?,
            _ => {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid raw selected cell",
                ));
            }
        };
        if !a.contains_key("s")
            && ["s", "customFormat"]
                .iter()
                .any(|key| source.get(*key) != destination.get(*key))
        {
            return Err(crate::coded(
                "internal-mismatch",
                "Implicit row formatting differs at the compacted destination",
            ));
        }
        let namespaces = |map: &BTreeMap<String, String>| {
            map.iter()
                .filter(|(key, _)| *key == "xmlns" || key.starts_with("xmlns:"))
                .map(|(k, v)| (k.clone(), v.clone()))
                .collect::<BTreeMap<_, _>>()
        };
        if namespaces(&source) != namespaces(&destination) {
            return Err(crate::coded(
                "internal-mismatch",
                "Row-local namespace environments differ at the compacted destination",
            ));
        }
        let formula: Option<(String, String)> = db
            .query_row(
                "SELECT old,new FROM dedup_formulas WHERE sheet=?1 AND ref=?2",
                params![sheet, address(from, col)],
                |r| Ok((r.get(0)?, r.get(1)?)),
            )
            .optional()
            .map_err(err)?;
        let output = rewrite_cell(
            &raw,
            from,
            to,
            col,
            formula.as_ref().map(|(a, b)| (a.as_str(), b.as_str())),
        )?;
        db.execute(
            "INSERT INTO dedup_output_cells VALUES(?1,?2,?3)",
            params![to, col, output],
        )
        .map_err(err)?;
    }
    Ok(())
}
fn cells_to(db: &Connection, row: u32, out: &mut impl Write) -> Result<()> {
    let mut statement = db
        .prepare_cached("SELECT raw FROM dedup_output_cells WHERE destination=?1 ORDER BY c")
        .map_err(err)?;
    let mut cells = statement.query([row]).map_err(err)?;
    while let Some(cell) = cells.next().map_err(err)? {
        crate::cancel::check()?;
        out.write_all(&cell.get::<_, Vec<u8>>(0).map_err(err)?)
            .map_err(err)?;
    }
    Ok(())
}
fn patches(db: &Connection, c: &Config, scan: &Scan) -> Result<()> {
    db.execute(
        "INSERT INTO dedup_patches SELECT start,end,'remove',r FROM dedup_source_cells",
        [],
    )
    .map_err(err)?;
    let mut targets = db
        .prepare_cached("SELECT DISTINCT destination FROM dedup_output_cells ORDER BY destination")
        .map_err(err)?;
    let mut rows = targets.query([]).map_err(err)?;
    while let Some(row) = rows.next().map_err(err)? {
        let r: u32 = row.get(0).map_err(err)?;
        let existing: Option<(u64, u64, bool, u64)> = db
            .query_row(
                "SELECT start,end,empty,insert_at FROM dedup_physical_rows WHERE r=?1",
                [r],
                |row| Ok((row.get(0)?, row.get(1)?, row.get(2)?, row.get(3)?)),
            )
            .optional()
            .map_err(err)?;
        if let Some((start, end, empty, at)) = existing {
            db.execute(
                "INSERT INTO dedup_patches VALUES(?1,?2,?3,?4)",
                params![
                    if empty { start } else { at },
                    if empty { end } else { at },
                    if empty { "expand" } else { "insert" },
                    r
                ],
            )
            .map_err(err)?;
        } else {
            let next: Option<u64> = db
                .query_row(
                    "SELECT start FROM dedup_physical_rows WHERE r>?1 ORDER BY r LIMIT 1",
                    [r],
                    |row| row.get(0),
                )
                .optional()
                .map_err(err)?;
            db.execute(
                "INSERT INTO dedup_patches VALUES(?1,?1,'row',?2)",
                params![next.unwrap_or(scan.sheet_close), r],
            )
            .map_err(err)?;
        }
    }
    let outside:u64=db.prepare_cached("SELECT count(*) FROM dedup_output_cells WHERE destination<?1 OR destination>?2 OR c<?3 OR c>?4").map_err(crate::err)?.query_row(params![c.body,c.last.0,c.first.1,c.last.1],|row|row.get(0)).map_err(err)?;
    if outside != 0 {
        return Err(crate::coded(
            "duplicates-duplicate",
            "Duplicate output exceeded its selected body",
        ));
    }
    Ok(())
}
fn worksheet(db: &Connection, source: &Path, scan: &Scan, out: &mut impl Write) -> Result<()> {
    let mut file = File::open(source).map_err(err)?;
    let length = file.metadata().map_err(err)?.len();
    let mut cursor = 0;
    let mut statement=db.prepare_cached("SELECT start,end,kind,r FROM dedup_patches ORDER BY start,CASE kind WHEN 'row' THEN 0 WHEN 'insert' THEN 1 WHEN 'expand' THEN 1 ELSE 2 END,r").map_err(err)?;
    let mut rows = statement.query([]).map_err(err)?;
    while let Some(row) = rows.next().map_err(err)? {
        crate::cancel::check()?;
        let first: u64 = row.get(0).map_err(err)?;
        let last: u64 = row.get(1).map_err(err)?;
        let kind: String = row.get(2).map_err(err)?;
        let r: u32 = row.get(3).map_err(err)?;
        copy_span(&mut file, cursor, first, out)?;
        match kind.as_str() {
            "remove" => {}
            "insert" => cells_to(db, r, out)?,
            "row" => {
                out.write_all(format!("<{} r=\"{}\">", scan.row_name, r + 1).as_bytes())
                    .map_err(err)?;
                cells_to(db, r, out)?;
                out.write_all(format!("</{}>", scan.row_name).as_bytes())
                    .map_err(err)?;
            }
            "expand" => {
                let mut raw = fragment(&mut file, first, last, 65536)?;
                let slash = raw.iter().rposition(|b| *b == b'/').ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing empty-row slash")
                })?;
                raw.remove(slash);
                out.write_all(&raw).map_err(err)?;
                cells_to(db, r, out)?;
                let name: String = db
                    .query_row(
                        "SELECT name FROM dedup_physical_rows WHERE r=?1",
                        [r],
                        |row| row.get(0),
                    )
                    .map_err(err)?;
                out.write_all(format!("</{name}>").as_bytes())
                    .map_err(err)?;
            }
            _ => {
                return Err(crate::coded(
                    "internal-unknown-reference",
                    "Unknown duplicate output patch",
                ));
            }
        }
        cursor = last;
    }
    copy_span(&mut file, cursor, length, out)
}

/// Create only a private candidate; owner adoption is a separate accepted boundary.
pub(super) fn apply(root: &Path, input: &Value) -> Result<Value> {
    let object = input
        .as_object()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid duplicate apply request"))?;
    let keys = [
        "action",
        "operation",
        "root",
        "source",
        "version",
        "sheet",
        "range",
        "hasHeaders",
        "columns",
        "expectedFingerprint",
    ];
    if keys.iter().any(|key| !object.contains_key(*key))
        || object
            .keys()
            .any(|key| !keys.contains(&key.as_str()) && key != "rewriteInputId")
        || input["action"] != "workbook"
        || input["operation"] != "dedup-apply"
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid duplicate apply fields",
        ));
    }
    let expected = field(input, "expectedFingerprint")?;
    if expected.len() != 64
        || !expected
            .bytes()
            .all(|b| b.is_ascii_digit() || (b'a'..=b'f').contains(&b))
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid duplicate preview fingerprint",
        ));
    }
    let c = config(input)?;
    let source = local(root, field(input, "source")?)?;
    let source_hash = crate::hash(&source)?;
    let base = local(root, "source.xlsx")?;
    let base_hash = crate::hash(&base)?;
    let index_path = local(root, "index.json")?;
    let index_hash = crate::hash(&index_path)?;
    let index: Value = serde_json::from_slice(&fs::read(&index_path).map_err(err)?).map_err(err)?;
    if index["revision"] != base_hash {
        return Err(crate::coded(
            "duplicates-duplicate",
            "Duplicate source changed since indexing",
        ));
    }
    let metadata_path = local(root, "metadata.xlsx")?;
    let metadata_hash = crate::hash(&metadata_path)?;
    let metadata = crate::package_edit::auditing_metadata(root)?;
    if crate::package_edit::source_calculation_policy(root)?["unsupportedMode"] == true {
        return Err(crate::coded(
            "duplicates-protected",
            "Protected source calculation policy",
        ));
    }
    if field(input, "source")? == "source.xlsx" {
        if c.version != 0 || source_hash != base_hash {
            return Err(crate::coded(
                "duplicates-duplicate",
                "Duplicate source does not bind the accepted version",
            ));
        }
    } else {
        let regenerated = crate::export::write(
            root,
            &json!({"version":c.version,"revision":index["revision"],"calculationMetadata":true}),
        )?;
        if regenerated["revision"] != source_hash {
            return Err(crate::coded(
                "duplicates-duplicate",
                "Duplicate source does not match independent accepted-version export",
            ));
        }
    }
    let db = database(root)?;
    let rewrite = crate::used_range_transfer::admit_writer(
        root,
        &db,
        input,
        crate::used_range_transfer::WriterKind::Dedup,
    )?;
    db.execute_batch("BEGIN; CREATE TEMP TABLE dedup_formulas(sheet TEXT,ref TEXT,old TEXT,new TEXT,PRIMARY KEY(sheet,ref)) WITHOUT ROWID; CREATE TEMP TABLE dedup_physical_rows(r INTEGER PRIMARY KEY,start INTEGER,open INTEGER,close INTEGER,end INTEGER,empty INTEGER,name TEXT,attrs TEXT,insert_at INTEGER); CREATE TEMP TABLE dedup_source_cells(r INTEGER,c INTEGER,start INTEGER,end INTEGER,raw BLOB,PRIMARY KEY(r,c)) WITHOUT ROWID; CREATE TEMP TABLE dedup_output_cells(destination INTEGER,c INTEGER,raw BLOB,PRIMARY KEY(destination,c)) WITHOUT ROWID; CREATE TEMP TABLE dedup_patches(start INTEGER,end INTEGER,kind TEXT,r INTEGER);").map_err(err)?;
    if db
        .query_row(
            "SELECT count(*) FROM versions WHERE id=?1",
            [c.version],
            |r| r.get::<_, u64>(0),
        )
        .map_err(err)?
        != 1
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown duplicate apply version",
        ));
    }
    let receipt = prepare(root, &db, input)?;
    if receipt["fingerprint"].as_str() != Some(expected) {
        return Err(crate::coded(
            "duplicates-duplicate",
            "Duplicate preview fingerprint is stale",
        ));
    }
    if receipt["removedRows"] == 0 {
        return Err(crate::coded(
            "duplicates-bad-input",
            "No duplicate records to remove",
        ));
    }
    let total = receipt["totalRows"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing duplicate record count"))?;
    let kept = receipt["keptRows"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing kept record count"))?;
    let(count,first,last,bad):(u64,u32,u32,u64)=db.prepare_cached("SELECT count(*),min(source),max(source),sum(CASE WHEN (source=first_row AND destination IS NULL) OR (source!=first_row AND destination IS NOT NULL) OR first_row>source THEN 1 ELSE 0 END) FROM duplicate_rows").map_err(crate::err)?.query_row([],|r|Ok((r.get(0)?,r.get(1)?,r.get(2)?,r.get(3)?))).map_err(err)?;
    if count != total || first != c.body || last != c.last.0 || bad != 0 {
        return Err(crate::coded(
            "internal-invalid-field",
            "Recomputed duplicate row map is invalid",
        ));
    }
    let (destinations, low, high): (u64, u32, u32) = db
        .query_row(
            "SELECT count(destination),min(destination),max(destination) FROM duplicate_rows",
            [],
            |r| Ok((r.get(0)?, r.get(1)?, r.get(2)?)),
        )
        .map_err(err)?;
    if destinations != kept || low != c.body || u64::from(high) != u64::from(c.body) + kept - 1 {
        return Err(crate::coded(
            "duplicates-duplicate",
            "Duplicate destinations are not contiguous",
        ));
    }
    if crate::spill::any_active(&db, c.version)? {
        return Err(crate::coded(
            "duplicates-unsupported",
            "Active spill ownership requires dedicated compaction qualification",
        ));
    }
    let mut zip =
        ZipArchive::new(crate::cancel::Reader(File::open(&source).map_err(err)?)).map_err(err)?;
    let names = zip.file_names().map(str::to_owned).collect::<BTreeSet<_>>();
    if names.len() != zip.len() {
        return Err(crate::coded(
            "duplicates-duplicate",
            "Duplicate package members",
        ));
    }
    if names.iter().any(|name| {
        name.starts_with("_xmlsignatures/")
            || name.contains("vbaProject")
            || name.contains("externalLinks/")
            || name.contains("connections.xml")
            || [
                "/drawings/",
                "/charts/",
                "/pivotTables/",
                "/pivotCache/",
                "/tables/",
                "/queryTables/",
                "/comments",
                "/threadedComments",
                "/activeX/",
                "/embeddings/",
            ]
            .iter()
            .any(|part| name.contains(part))
    }) {
        return Err(crate::coded(
            "duplicates-unsupported",
            "Foreign/signed/connected workbooks require compaction qualification",
        ));
    }
    {
        let mut content = String::new();
        zip.by_name("[Content_Types].xml")
            .map_err(err)?
            .take(1048577)
            .read_to_string(&mut content)
            .map_err(err)?;
        if content.len() > 1048576
            || content.contains("macroEnabled")
            || content.contains("vbaProject")
        {
            return Err(crate::coded(
                "duplicates-protected",
                "Foreign workbook package is protected",
            ));
        }
    }
    source_workbook(&mut zip, &metadata)?;
    let sheet = field(input, "sheet")?;
    formulas(&db, &c, sheet, &metadata)?;
    let scratch = Scratch {
        worksheet: local(root, "dedup-worksheet.xml")?,
        partial: local(root, "dedup.partial")?,
    };
    let output = local(root, "dedup.xlsx")?;
    let worksheet_hash = {
        let mut member = zip.by_name(sheet).map_err(err)?;
        let mut file = crate::quota::create(&scratch.worksheet).map_err(err)?;
        let mut buffer = [0u8; 65536];
        let mut digest = Sha256::new();
        loop {
            crate::cancel::check()?;
            let count = member.read(&mut buffer).map_err(err)?;
            if count == 0 {
                break;
            }
            file.write_all(&buffer[..count]).map_err(err)?;
            digest.update(&buffer[..count]);
        }
        file.sync_all().map_err(err)?;
        format!("{:x}", digest.finalize())
    };
    let scan = scan(&db, &scratch.worksheet, &c)?;
    inactive_filter(&db, &scan)?;
    spool_outputs(&db, &c, sheet)?;
    patches(&db, &c, &scan)?;
    let mut writer = ZipWriter::new(crate::quota::create(&scratch.partial).map_err(err)?);
    for i in 0..zip.len() {
        crate::cancel::check()?;
        let member = zip.by_index(i).map_err(err)?;
        if member.name() != sheet {
            writer.raw_copy_file(member).map_err(err)?;
        } else {
            writer
                .start_file(
                    sheet,
                    SimpleFileOptions::default()
                        .compression_method(zip::CompressionMethod::Deflated),
                )
                .map_err(err)?;
            worksheet(&db, &scratch.worksheet, &scan, &mut writer)?;
        }
    }
    writer.finish().map_err(err)?.sync_all().map_err(err)?;
    if crate::hash(&source)? != source_hash
        || crate::hash(&base)? != base_hash
        || crate::hash(&index_path)? != index_hash
        || crate::hash(&metadata_path)? != metadata_hash
        || crate::hash(&scratch.worksheet)? != worksheet_hash
    {
        return Err(crate::coded(
            "duplicates-duplicate",
            "Duplicate source/metadata retired while building its private candidate",
        ));
    }
    crate::cancel::check()?;
    fs::rename(&scratch.partial, &output).map_err(err)?;
    let prepared = if let Some(rewrite) = rewrite {
        let rows = db
            .prepare_cached("SELECT source,destination FROM duplicate_rows ORDER BY source")
            .map_err(err)?
            .query_map([], |row| {
                Ok((row.get::<_, u32>(0)?, row.get::<_, Option<u32>>(1)?))
            })
            .map_err(err)?
            .collect::<std::result::Result<Vec<_>, _>>()
            .map_err(err)?;
        Some(crate::used_range_transfer::prepare_output(
            root,
            &db,
            rewrite,
            crate::used_range_transfer::NativeMapping::Compact {
                sheet: sheet.to_owned(),
                first: c.first,
                last: c.last,
                body: c.body,
                rows,
            },
            "dedup.xlsx",
        )?)
    } else {
        None
    };
    // This writer's BEGIN owns only scratch TEMP state and rolls back on drop.
    // A rewrite receipt must commit on a fresh connection after that rollback.
    drop(db);
    let mut result = json!({"file":"dedup.xlsx","fingerprint":expected,"totalRows":total,"keptRows":kept,"removedRows":total-kept,"comparison":"exact-typed-values-v1"});
    if let Some(prepared) = prepared {
        result["rewriteId"] = json!(crate::used_range_transfer::commit_output(root, prepared)?);
    }
    Ok(result)
}
