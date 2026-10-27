//! Bounded one-root draft calculation. Main SQLite and accepted artifacts are read-only.
use crate::xml::{address, position};
use crate::{Result, err, field};
use rusqlite::{Connection, OpenFlags, params};
use serde_json::{Value, json};
use std::{
    collections::{HashMap, HashSet},
    io::Read,
    path::Path,
    time::Instant,
};
const POSITIONS: usize = 10_000;
const FORMULAS: usize = 256;
pub(crate) fn error_code(code: &str) -> bool {
    [
        "#DIV/0!", "#VALUE!", "#NUM!", "#REF!", "#N/A", "#NAME?", "#NULL!",
    ]
    .contains(&code)
}
fn request_id(id: &str) -> bool {
    let bytes = id.as_bytes();
    bytes.len() == 36
        && bytes.iter().enumerate().all(|(i, b)| {
            if [8, 13, 18, 23].contains(&i) {
                *b == b'-'
            } else {
                b.is_ascii_digit() || (b'a'..=b'f').contains(b)
            }
        })
        && [b'4', b'7'].contains(&bytes[14])
        && (b"89ab").contains(&bytes[19])
}
fn numeric_spelling(formula: &str) -> Result<()> {
    let b = formula.as_bytes();
    let mut at = 0;
    while at < b.len() {
        crate::cancel::check()?;
        let c = b[at];
        if c == b'"' || c == b'\'' {
            let q = c;
            at += 1;
            while at < b.len() {
                if b[at] == q {
                    at += 1;
                    if b.get(at) == Some(&q) {
                        at += 1;
                    } else {
                        break;
                    }
                } else {
                    at += 1;
                }
            }
            continue;
        }
        if c.is_ascii_alphabetic() || c == b'_' || c == b'$' {
            at += 1;
            while b
                .get(at)
                .is_some_and(|c| c.is_ascii_alphanumeric() || [b'_', b'.', b'$'].contains(c))
            {
                at += 1;
            }
            continue;
        }
        if c == b'#' {
            at += 1;
            while b.get(at).is_some_and(|c| !b" ,()+-*/<>=&".contains(c)) {
                at += 1;
            }
            continue;
        }
        if c.is_ascii_digit() || c == b'.' {
            let start = at;
            while b.get(at).is_some_and(u8::is_ascii_digit) {
                at += 1;
            }
            if b.get(at) == Some(&b'.') {
                at += 1;
                while b.get(at).is_some_and(u8::is_ascii_digit) {
                    at += 1;
                }
            }
            if b.get(at).is_some_and(|c| [b'e', b'E'].contains(c)) {
                at += 1;
                if b.get(at).is_some_and(|c| [b'+', b'-'].contains(c)) {
                    at += 1;
                }
                while b.get(at).is_some_and(u8::is_ascii_digit) {
                    at += 1;
                }
            }
            let raw = &formula[start..at];
            let n = raw.parse::<f64>().map_err(|_| {
                crate::coded("internal-invalid-field", "Invalid draft numeric literal")
            })?;
            if !n.is_finite()
                || n.abs() > 9007199254740991.0
                || crate::bulk::meaning(raw) != crate::bulk::meaning(&n.to_string())
            {
                return Err(crate::coded(
                    "workbook-unsupported",
                    "Draft numeric literal precision is outside the scalar profile",
                ));
            }
            continue;
        }
        at += 1;
    }
    Ok(())
}
struct Metadata {
    view: Value,
    merges: HashMap<String, Vec<((u32, u32), (u32, u32))>>,
    protected: HashSet<String>,
}
fn metadata(root: &Path, supplied: &Value) -> Result<Metadata> {
    let actual = crate::package_edit::auditing_metadata(root)?;
    let sheets = supplied["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing preview worksheets"))?;
    if actual["sheets"]
        != json!(
            sheets
                .iter()
                .map(|s| json!({"name":s["name"],"path":s["path"]}))
                .collect::<Vec<_>>()
        )
        || !crate::package_edit::same_records(&actual["namedRanges"], &supplied["namedRanges"])
        || !crate::package_edit::same_records(&actual["tables"], &supplied["tables"])
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Preview metadata differs from accepted source",
        ));
    }
    let index: Value =
        serde_json::from_slice(&std::fs::read(crate::local(root, "index.json")?).map_err(err)?)
            .map_err(err)?;
    let indexed = index["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing preview index"))?;
    if indexed.len() != sheets.len()
        || indexed
            .iter()
            .any(|i| !sheets.iter().any(|s| s["path"] == i["path"]))
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Preview worksheet index differs from source",
        ));
    }
    if supplied["readOnly"] == true
        || crate::package_edit::source_calculation_policy(root)?["unsupportedMode"] == true
    {
        return Err(crate::coded(
            "workbook-protected",
            "Draft preview target uses protected workbook settings",
        ));
    }
    let mut zip = zip::ZipArchive::new(crate::cancel::Reader(
        std::fs::File::open(crate::local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    if zip
        .file_names()
        .any(|name| name.starts_with("_xmlsignatures/") || name.ends_with("vbaProject.bin"))
    {
        return Err(crate::coded(
            "workbook-unsupported",
            "Signed or macro workbook draft preview is unavailable",
        ));
    }
    let members = zip.file_names().map(str::to_owned).collect::<HashSet<_>>();
    let mut total = 0;
    let mut part = |name: &str| -> Result<Vec<u8>> {
        let mut file = zip.by_name(name).map_err(err)?;
        let size = file.size();
        if size > 16 * 1024 * 1024 || total + size > 64 * 1024 * 1024 {
            return Err(crate::coded(
                "workbook-limit",
                "Preview source metadata byte bound",
            ));
        }
        total += size;
        let mut b = Vec::new();
        file.read_to_end(&mut b).map_err(err)?;
        Ok(b)
    };
    let types = part("[Content_Types].xml")?;
    let types = std::str::from_utf8(&types).map_err(err)?;
    if types.contains("macroEnabled")
        || types.contains("vbaProject")
        || types.contains("digital-signature")
    {
        return Err(crate::coded(
            "workbook-protected",
            "Macro or signed source settings are protected",
        ));
    }
    let wb = part("xl/workbook.xml")?;
    let mut date1904 = false;
    let mut count = 0;
    {
        let mut reader = quick_xml::Reader::from_reader(wb.as_slice());
        loop {
            match reader.read_event().map_err(err)? {
                quick_xml::events::Event::Start(e) | quick_xml::events::Event::Empty(e)
                    if e.local_name().as_ref() == b"workbookPr" =>
                {
                    count += 1;
                    if count > 1 {
                        return Err(crate::coded(
                            "workbook-ambiguous",
                            "Ambiguous preview calendar",
                        ));
                    }
                    let attrs = crate::xml::attrs(&e)?;
                    date1904 = match attrs.get("date1904").map(String::as_str) {
                        None | Some("0" | "false") => false,
                        Some("1" | "true") => true,
                        _ => {
                            return Err(crate::coded(
                                "internal-invalid-field",
                                "Invalid preview calendar",
                            ));
                        }
                    };
                }
                quick_xml::events::Event::Eof => break,
                _ => (),
            }
        }
    }
    let mut view = actual;
    let mut merges = HashMap::new();
    let mut protected = HashSet::new();
    for sheet in view["sheets"].as_array_mut().unwrap() {
        let name = field(sheet, "name")?.to_owned();
        let path = field(sheet, "path")?.to_owned();
        let supplied = sheets.iter().find(|s| s["path"] == path).unwrap();
        if supplied["date1904"].as_bool().unwrap_or(false) != date1904 {
            return Err(crate::coded(
                "internal-mismatch",
                "Preview calendar differs from accepted source",
            ));
        }
        sheet["date1904"] = json!(date1904);
        let bytes = part(&path)?;
        let mut reader = quick_xml::Reader::from_reader(bytes.as_slice());
        let mut ranges = Vec::new();
        loop {
            crate::cancel::check()?;
            match reader.read_event().map_err(err)? {
                quick_xml::events::Event::Start(e) | quick_xml::events::Event::Empty(e) => {
                    match e.local_name().as_ref() {
                        b"sheetProtection" => {
                            protected.insert(name.clone());
                        }
                        b"mergeCell" => {
                            if ranges.len() >= 10_000 {
                                return Err(crate::coded(
                                    "merge-limit",
                                    "Preview merged metadata bound",
                                ));
                            }
                            ranges.push(crate::bulk::bounds(
                                crate::xml::attrs(&e)?.get("ref").ok_or_else(|| {
                                    crate::coded(
                                        "internal-invalid-field",
                                        "Invalid merged preview metadata",
                                    )
                                })?,
                            )?);
                        }
                        _ => (),
                    }
                }
                quick_xml::events::Event::Eof => break,
                _ => (),
            }
        }
        let rel = format!(
            "{}/_rels/{}.rels",
            path.rsplit_once('/')
                .ok_or_else(|| crate::coded(
                    "workbook-malformed",
                    "Invalid preview worksheet path"
                ))?
                .0,
            path.rsplit_once('/').unwrap().1
        );
        let rel_bytes = if members.contains(&rel) {
            Some(part(&rel)?)
        } else {
            None
        };
        if let Some(bytes) = rel_bytes {
            let mut reader = quick_xml::Reader::from_reader(bytes.as_slice());
            let mut ids = HashSet::new();
            loop {
                match reader.read_event().map_err(err)? {
                    quick_xml::events::Event::Start(e) | quick_xml::events::Event::Empty(e)
                        if e.local_name().as_ref() == b"Relationship" =>
                    {
                        let attrs = crate::xml::attrs(&e)?;
                        let id = attrs.get("Id").ok_or_else(|| {
                            crate::coded(
                                "internal-missing-field",
                                "Missing preview relationship identity",
                            )
                        })?;
                        if !ids.insert(id.clone()) {
                            return Err(crate::coded(
                                "package-ambiguous",
                                "Ambiguous preview relationships",
                            ));
                        }
                        if attrs
                            .get("Type")
                            .is_some_and(|kind| kind.ends_with("/pivotTable"))
                        {
                            if attrs
                                .get("TargetMode")
                                .is_some_and(|mode| mode != "Internal")
                            {
                                return Err(crate::coded(
                                    "pivot-protected",
                                    "External pivot preview source is protected",
                                ));
                            }
                            let target = attrs.get("Target").ok_or_else(|| {
                                crate::coded(
                                    "internal-missing-field",
                                    "Missing pivot preview target",
                                )
                            })?;
                            if target.starts_with('/') || target.contains('\\') {
                                return Err(crate::coded(
                                    "pivot-unsupported",
                                    "Unconfined pivot preview target",
                                ));
                            }
                            let mut components = path
                                .rsplit_once('/')
                                .unwrap()
                                .0
                                .split('/')
                                .map(str::to_owned)
                                .collect::<Vec<_>>();
                            for component in target.split('/') {
                                match component {
                                    ".." => {
                                        components.pop().ok_or_else(|| {
                                            crate::coded(
                                                "pivot-unsupported",
                                                "Unconfined pivot preview target",
                                            )
                                        })?;
                                    }
                                    "." => (),
                                    "" => {
                                        return Err(crate::coded(
                                            "pivot-malformed",
                                            "Invalid pivot preview path",
                                        ));
                                    }
                                    _ => components.push(component.into()),
                                }
                            }
                            if components.first().map(String::as_str) != Some("xl") {
                                return Err(crate::coded(
                                    "pivot-unsupported",
                                    "Unconfined pivot preview source",
                                ));
                            }
                            let pivot = part(&components.join("/"))?;
                            let mut pivot = quick_xml::Reader::from_reader(pivot.as_slice());
                            let mut location = None;
                            loop {
                                match pivot.read_event().map_err(err)? {
                                    quick_xml::events::Event::Start(e)
                                    | quick_xml::events::Event::Empty(e)
                                        if e.local_name().as_ref() == b"location" =>
                                    {
                                        if location.is_some() {
                                            return Err(crate::coded(
                                                "pivot-ambiguous",
                                                "Ambiguous pivot preview location",
                                            ));
                                        }
                                        location = Some(crate::bulk::bounds(
                                            crate::xml::attrs(&e)?.get("ref").ok_or_else(|| {
                                                crate::coded(
                                                    "internal-missing-field",
                                                    "Missing pivot preview location",
                                                )
                                            })?,
                                        )?);
                                    }
                                    quick_xml::events::Event::Eof => break,
                                    _ => (),
                                }
                            }
                            ranges.push(location.ok_or_else(|| {
                                crate::coded(
                                    "internal-missing-field",
                                    "Missing pivot preview output geometry",
                                )
                            })?);
                        }
                    }
                    quick_xml::events::Event::Eof => break,
                    _ => (),
                }
            }
        }
        merges.insert(name, ranges);
    }
    Ok(Metadata {
        view,
        merges,
        protected,
    })
}
struct Closure<'a> {
    root: &'a Path,
    db: &'a Connection,
    version: u64,
    metadata: &'a Metadata,
    seen: HashSet<(String, String)>,
    heights: HashMap<(String, String), usize>,
    active: HashSet<(String, String)>,
    formulas: usize,
    bytes: usize,
    special_sheets: HashSet<String>,
}
impl Closure<'_> {
    fn path(&self, sheet: &str) -> Result<String> {
        Ok(field(
            self.metadata.view["sheets"]
                .as_array()
                .unwrap()
                .iter()
                .find(|s| s["name"] == sheet)
                .ok_or_else(|| {
                    crate::coded("internal-unknown-reference", "Unknown preview worksheet")
                })?,
            "path",
        )?
        .into())
    }
    fn overlay(&self, path: &str, at: (u32, u32)) -> Result<()> {
        self.db.execute("WITH RECURSIVE a(id,depth)AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL) INSERT OR IGNORE INTO calc_overlay SELECT c.sheet,c.r,c.c,c.data,'{}' FROM a JOIN changes c ON c.version=a.id WHERE c.sheet=?2 AND c.r=?3 AND c.c=?4 ORDER BY a.depth LIMIT 1",params![self.version,path,at.0,at.1]).map_err(err)?;
        Ok(())
    }
    fn specials(&mut self, path: &str) -> Result<()> {
        if !self.special_sheets.insert(path.into()) {
            return Ok(());
        }
        let mut query=self.db.prepare_cached("SELECT r,c,formula_attrs FROM cells WHERE sheet=?1 AND formula_attrs IS NOT NULL LIMIT 4097").map_err(err)?;
        let rows = query
            .query_map([path], |r| {
                Ok((
                    r.get::<_, u32>(0)?,
                    r.get::<_, u32>(1)?,
                    r.get::<_, String>(2)?,
                ))
            })
            .map_err(err)?;
        let mut count = 0;
        for row in rows {
            crate::cancel::check()?;
            count += 1;
            if count > 4096 {
                return Err(crate::coded(
                    "workbook-limit",
                    "Preview special source metadata bound",
                ));
            }
            let (r, c, attrs) = row.map_err(err)?;
            let attrs: Value = serde_json::from_str(&attrs).map_err(err)?;
            if matches!(attrs["t"].as_str(), Some("array" | "dataTable" | "shared")) {
                let (a, b) = crate::bulk::bounds(attrs["ref"].as_str().unwrap_or(&address(r, c)))?;
                self.db
                    .execute(
                        "INSERT INTO financial_protected VALUES(?1,?2,?3,?4,?5)",
                        params![path, a.0, a.1, b.0, b.1],
                    )
                    .map_err(err)?;
            }
        }
        for group in crate::spill::effective(self.db, self.version, path)? {
            let g = group.geometry;
            self.db
                .execute(
                    "INSERT INTO financial_protected VALUES(?1,?2,?3,?4,?5)",
                    params![path, g.first.0, g.first.1, g.last.0, g.last.1],
                )
                .map_err(err)?;
        }
        // Persistent scheduler metadata records derived output geometry; it supplies
        // protection, never a value or calculation certificate for this preview.
        if self.db.prepare_cached("SELECT EXISTS(SELECT 1 FROM sqlite_master WHERE name='calculation_object_outputs')").map_err(crate::err)?.query_row([],|r|r.get::<_,bool>(0)).map_err(err)? {self.db.execute("WITH RECURSIVE a(id,depth)AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL),p AS(SELECT b.plan FROM a JOIN calculation_bindings b ON b.version=a.id ORDER BY a.depth LIMIT 1) INSERT INTO financial_protected SELECT o.sheet,o.r0,o.c0,o.r1,o.c1 FROM p JOIN calculation_object_outputs o ON o.plan=p.plan WHERE o.sheet=?2",params![self.version,path]).map_err(err)?;}
        Ok(())
    }
    fn visit(&mut self, sheet: &str, reference: &str, depth: usize) -> Result<usize> {
        crate::cancel::check()?;
        let key = (sheet.to_owned(), reference.to_owned());
        if self.active.contains(&key) {
            return Err(crate::coded(
                "workbook-refused",
                "Circular draft prerequisite",
            ));
        }
        if let Some(height) = self.heights.get(&key) {
            if depth + height - 1 > 32 {
                return Err(crate::coded(
                    "workbook-limit",
                    "Draft prerequisite depth bound",
                ));
            }
            return Ok(*height);
        }
        if depth > 32 || self.seen.len() >= POSITIONS {
            return Err(crate::coded(
                "workbook-limit",
                "Draft prerequisite depth or position bound",
            ));
        }
        self.seen.insert(key.clone());
        let at = position(reference)?;
        let path = self.path(sheet)?;
        self.specials(&path)?;
        if self.metadata.protected.contains(sheet) {
            return Err(crate::coded(
                "workbook-protected",
                "Draft prerequisite uses protected worksheet",
            ));
        }
        if self.metadata.merges.get(sheet).is_some_and(|ranges| {
            ranges
                .iter()
                .any(|(a, b)| at.0 >= a.0 && at.0 <= b.0 && at.1 >= a.1 && at.1 <= b.1)
        }) {
            return Err(crate::coded(
                "merge-refused",
                "Draft preview uses a merged or derived output cell",
            ));
        }
        if self.db.prepare_cached("SELECT EXISTS(SELECT 1 FROM financial_protected WHERE sheet=?1 AND ?2 BETWEEN r0 AND r1 AND ?3 BETWEEN c0 AND c1)").map_err(crate::err)?.query_row(params![path,at.0,at.1],|r|r.get::<_,bool>(0)).map_err(err)?{return Err(crate::coded("spill-unsupported", "Draft preview reads a shared, array, spill or derived output"));}
        self.overlay(&path, at)?;
        let cell = crate::calculation::input_cell(self.db, &path, reference)?;
        let size = serde_json::to_vec(&cell).map_err(err)?.len();
        self.bytes += size;
        if size > 65536 || self.bytes > 2 * 1024 * 1024 {
            return Err(crate::coded(
                "workbook-limit",
                "Draft prerequisite byte bound",
            ));
        }
        if cell.get("formulaType").is_some()
            || cell.get("spill").is_some()
            || cell["scenarioResult"] == true
        {
            return Err(crate::coded(
                "workbook-unsupported",
                "Draft prerequisite is a special output",
            ));
        }
        if crate::spill::source_attributes(self.root, self.db, &path, at)?.is_some_and(|attrs| {
            attrs
                .keys()
                .any(|key| !["r", "s", "t", "xmlns"].contains(&key.as_str()))
                || attrs
                    .get("xmlns")
                    .is_some_and(|ns| ns.as_bytes() != crate::xml::NS)
        }) {
            return Err(crate::coded(
                "workbook-unsupported",
                "Draft source cell has opaque attributes",
            ));
        }
        crate::spill::source_content_proof(self.root, self.db, &path, at, true)?;
        let Some(formula) = cell["formula"].as_str() else {
            if cell["value"].is_number() {
                crate::bulk::literal(&cell, false)?;
            }
            self.heights.insert(key, 1);
            return Ok(1);
        };
        self.formulas += 1;
        if self.formulas > FORMULAS {
            return Err(crate::coded(
                "formula-limit",
                "Draft formula prerequisite count bound",
            ));
        }
        numeric_spelling(formula)?;
        let deps =
            crate::formula::preview_dependencies(&self.metadata.view, sheet, reference, formula)?;
        self.active.insert(key.clone());
        let mut height = 1;
        for dep in deps["ranges"].as_array().ok_or_else(|| {
            crate::coded(
                "internal-invalid-field",
                "Invalid preview static dependencies",
            )
        })? {
            let owner = field(dep, "sheet")?;
            let (a, b) = if let Some(reference) = dep["ref"].as_str() {
                let a = position(reference)?;
                (a, a)
            } else {
                (
                    (
                        dep["startRow"].as_u64().ok_or_else(|| {
                            crate::coded("internal-invalid-field", "Invalid reference")
                        })? as u32,
                        dep["startColumn"].as_u64().ok_or_else(|| {
                            crate::coded("internal-invalid-field", "Invalid reference")
                        })? as u32,
                    ),
                    (
                        dep["endRow"].as_u64().ok_or_else(|| {
                            crate::coded("internal-invalid-field", "Invalid reference")
                        })? as u32,
                        dep["endColumn"].as_u64().ok_or_else(|| {
                            crate::coded("internal-invalid-field", "Invalid reference")
                        })? as u32,
                    ),
                )
            };
            let n = u64::from(b.0 - a.0 + 1) * u64::from(b.1 - a.1 + 1);
            if n > POSITIONS as u64 {
                return Err(crate::coded(
                    "workbook-limit",
                    "Draft reference position bound",
                ));
            }
            for r in a.0..=b.0 {
                for c in a.1..=b.1 {
                    height = height.max(1 + self.visit(owner, &address(r, c), depth + 1)?);
                }
            }
        }
        self.active.remove(&key);
        self.heights.insert(key, height);
        Ok(height)
    }
}
pub(crate) fn command(root: &Path, input: &Value) -> Result<Value> {
    let allowed = [
        "action",
        "operation",
        "root",
        "version",
        "view",
        "sheet",
        "cell",
        "formula",
        "requestId",
    ];
    if input
        .as_object()
        .is_none_or(|o| o.keys().any(|key| !allowed.contains(&key.as_str())))
    {
        return Err(crate::coded(
            "formula-malformed",
            "Unexpected formula preview field",
        ));
    }
    let id = field(input, "requestId")?;
    if !request_id(id) {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid formula preview request ID",
        ));
    }
    let sheet = field(input, "sheet")?;
    let cell = field(input, "cell")?;
    if address(position(cell)?.0, position(cell)?.1) != cell {
        return Err(crate::coded(
            "workbook-malformed",
            "Invalid canonical preview target",
        ));
    }
    let formula = field(input, "formula")?;
    if formula.is_empty()
        || formula.starts_with('=')
        || formula.len() > 8192
        || formula.chars().any(|c| c == '\0')
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid formula preview draft",
        ));
    }
    crate::cancel::preview(root, id);
    let evaluated = (|| -> Result<Value> {
        crate::cancel::check()?;
        let version = input["version"]
            .as_u64()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing preview version"))?;
        let metadata = metadata(root, &input["view"])?;
        let path = field(
            metadata.view["sheets"]
                .as_array()
                .unwrap()
                .iter()
                .find(|s| s["name"] == sheet)
                .ok_or_else(|| {
                    crate::coded(
                        "internal-unknown-reference",
                        "Unknown preview target worksheet",
                    )
                })?,
            "path",
        )?;
        let db = Connection::open_with_flags(
            crate::local(root, "cells.sqlite")?,
            OpenFlags::SQLITE_OPEN_READ_ONLY | OpenFlags::SQLITE_OPEN_NO_MUTEX,
        )
        .map_err(err)?;
        db.execute_batch(
            "PRAGMA cache_size=-4096;PRAGMA temp_store=MEMORY;PRAGMA mmap_size=0;BEGIN;",
        )
        .map_err(err)?;
        let mut counter = 0usize;
        let start = Instant::now();
        db.progress_handler(
            1000,
            Some(move || {
                counter += 1000;
                crate::cancel::interrupted()
                    || counter > 50_000_000
                    || start.elapsed().as_secs() > 10
            }),
        );
        // Shared producers are outside this preview profile. Install an empty
        // TEMP shared context for read helpers rather than materializing every
        // accepted normalization into preview scratch before reaching a cell.
        db.execute_batch("CREATE TEMP TABLE shared_context(version INTEGER);CREATE TEMP TABLE shared_active(sheet TEXT,group_id TEXT,master_r INTEGER,master_c INTEGER,formula TEXT,invariant INTEGER,PRIMARY KEY(sheet,group_id)) WITHOUT ROWID;").map_err(err)?;
        db.execute("INSERT INTO shared_context VALUES(?1)", [version])
            .map_err(err)?;
        crate::calculation::snapshot_tables(&db, version)?;
        db.execute_batch("CREATE TEMP TABLE calc_scenarios(sheet TEXT,r INTEGER,c INTEGER,model TEXT,input1 TEXT,value1 TEXT,input2 TEXT,value2 TEXT,PRIMARY KEY(sheet,r,c)) WITHOUT ROWID;CREATE TEMP TABLE financial_protected(sheet TEXT,r0 INTEGER,c0 INTEGER,r1 INTEGER,c1 INTEGER);").map_err(err)?;
        // Prove the accepted target before replacing it. Its old formula/result never
        // becomes a prerequisite of the candidate; self references see this override.
        let at = position(cell)?;
        let mut closure = Closure {
            root,
            db: &db,
            version,
            metadata: &metadata,
            seen: HashSet::new(),
            heights: HashMap::new(),
            active: HashSet::new(),
            formulas: 0,
            bytes: 0,
            special_sheets: HashSet::new(),
        };
        closure.specials(path)?;
        closure.overlay(path, at)?;
        let prior = crate::calculation::input_cell(&db, path, cell)?;
        if prior.get("formulaType").is_some()
            || prior.get("spill").is_some()
            || prior["scenarioResult"] == true
        {
            return Err(crate::coded(
                "workbook-unsupported",
                "Draft target is a special output",
            ));
        }
        db.execute(
            "INSERT OR REPLACE INTO calc_overlay VALUES(?1,?2,?3,?4,'{}')",
            params![
                path,
                at.0,
                at.1,
                json!({"formula":formula,"value":null,"dataType":"blank"}).to_string()
            ],
        )
        .map_err(err)?;
        closure.visit(sheet, cell, 1)?;
        let mut eval = crate::formula::Eval::from_sqlite(&metadata.view, &db, false);
        eval.preview_limits();
        let value = eval.result(sheet, cell)?;
        if let Some(n) = value.as_f64() {
            if !n.is_finite() {
                return Err(crate::coded(
                    "workbook-unsupported",
                    "Nonfinite draft result",
                ));
            }
            Ok(
                json!({"status":"value","type":"number","value":if n==0.0&&n.is_sign_negative(){json!({"$workbookNumber":"-0"})}else{value}}),
            )
        } else if value.is_boolean() {
            Ok(json!({"status":"value","type":"boolean","value":value}))
        } else if value.is_string() {
            Ok(json!({"status":"value","type":"text","value":value}))
        } else {
            return Err(crate::coded(
                "workbook-unsupported",
                "Draft result is not a qualified typed scalar",
            ));
        }
    })();
    let result = match evaluated {
        Ok(result) => result,
        Err(reason) => {
            if crate::cancel::interrupted() {
                json!({"status":"cancelled","reason":"Draft preview request cancelled"})
            } else if let Some(code) = crate::sentence(&reason)
                .strip_prefix("Excel error: ")
                .filter(|code| error_code(code))
            {
                json!({"status":"value","type":"error","value":code})
            } else {
                json!({"status":if reason.to_lowercase().contains("circular")||crate::sentence(&reason).starts_with("Invalid ")||reason.contains("Unterminated")||reason.contains("Unsupported formula syntax"){"invalid"}else{"unavailable"},"reason":reason})
            }
        }
    };
    Ok(
        json!({"schema":1,"requestId":id,"sheet":sheet,"cell":cell,"formula":formula,"authority":"draft-preview","result":result}),
    )
}
#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn request_and_precision_bounds() {
        assert!(request_id("12345678-1234-4123-8123-123456789abc"));
        assert!(!request_id("../x"));
        for f in [
            "SUM(A1:A3)+0.25",
            "PMT(B1/12,B2*12,B3)",
            "LEN(\"9007199254740993\")",
            "'Sheet12'!A1+1",
        ] {
            numeric_spelling(f).unwrap();
        }
        for f in ["9007199254740993", "0.1234567890123456789012345", "1e999"] {
            assert!(numeric_spelling(f).is_err(), "{f}");
        }
    }
}
