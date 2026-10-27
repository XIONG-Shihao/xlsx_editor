//! What follows inserted or deleted rows and columns beside the worksheets:
//! defined names, tables, notes (and their drawings), pictures and shapes,
//! chart references, and the formulas stored inside table parts.
//!
//! The worksheets' own records and cells are moved by `insertion.rs` and
//! `rewrite_admission.rs`. This module plans everything else from the package
//! itself, so that the caller only has to say which sheet, which rows or
//! columns, and how many. It returns whole replacement parts, the parts to
//! remove and the facts the cell rewriter needs (which table columns went).
use super::xml_edit::{Dom, Edit, apply, escape};
use super::*;
use rusqlite::Connection;

const DRAW: &[u8] = DRAW_NS;
const C15_NS: &[u8] = b"http://schemas.microsoft.com/office/drawing/2012/chart";
const CHARTEX_NS: &[u8] = b"http://schemas.microsoft.com/office/drawing/2014/chartex";

pub(crate) struct Plan {
    /// Whole parts that change, as text.
    pub replacements: BTreeMap<String, String>,
    /// Parts that go.
    pub removed: Vec<String>,
    /// For each table that lost columns or all of itself, what was lost.
    pub table_changes: Vec<Value>,
    /// New column names to write into header cells: sheet name, cell, text.
    pub headers: Vec<Value>,
    /// Table columns that calculate: first body row, last row, column, column.
    pub calculated: Vec<(u32, u32, u32, u32)>,
    /// Every sheet in workbook order: name and part.
    pub sheets: Vec<(String, String)>,
}

struct Rel {
    id: String,
    kind: String,
    target: String,
    external: bool,
}
struct Parts {
    zip: ZipArchive<crate::cancel::Reader<File>>,
    changed: BTreeMap<String, Vec<u8>>,
    removed: BTreeSet<String>,
}
impl Parts {
    fn exists(&self, name: &str) -> bool {
        !self.removed.contains(name)
            && (self.changed.contains_key(name) || self.zip.file_names().any(|part| part == name))
    }
    fn read(&mut self, name: &str) -> Result<Vec<u8>> {
        if let Some(bytes) = self.changed.get(name) {
            return Ok(bytes.clone());
        }
        if self.removed.contains(name) || !self.exists(name) {
            return Err(crate::coded(
                "internal-missing-field",
                format!("Missing package part {name}"),
            ));
        }
        let mut member = self.zip.by_name(name).map_err(err)?;
        if member.size() > XML_LIMIT as u64 * 4 {
            return Err(crate::coded("package-limit", "Package part exceeds its bound"));
        }
        let mut data = Vec::new();
        member.read_to_end(&mut data).map_err(err)?;
        Ok(data)
    }
    fn names(&self) -> Vec<String> {
        let mut names: BTreeSet<String> = self.zip.file_names().map(str::to_owned).collect();
        names.extend(self.changed.keys().cloned());
        names.into_iter().filter(|n| !self.removed.contains(n)).collect()
    }
    fn put(&mut self, name: &str, bytes: Vec<u8>) {
        self.changed.insert(name.to_owned(), bytes);
    }
    fn remove(&mut self, name: &str) {
        self.changed.remove(name);
        self.removed.insert(name.to_owned());
    }
    fn rels(&mut self, owner: &str) -> Result<Vec<Rel>> {
        let (directory, file) = owner.rsplit_once('/').unwrap_or(("", owner));
        let part = if directory.is_empty() {
            format!("_rels/{file}.rels")
        } else {
            format!("{directory}/_rels/{file}.rels")
        };
        if !self.exists(&part) {
            return Ok(Vec::new());
        }
        let dom = Dom::parse(&self.read(&part)?)?;
        let mut found = Vec::new();
        for node in dom.all(PACKAGE_REL_NS, "Relationship") {
            found.push(Rel {
                id: dom.attr(node, "Id").unwrap_or("").to_owned(),
                kind: dom.attr(node, "Type").unwrap_or("").to_owned(),
                target: dom.attr(node, "Target").unwrap_or("").to_owned(),
                external: dom.attr(node, "TargetMode") == Some("External"),
            });
        }
        Ok(found)
    }
    /// The parts of one kind that a part points at ("comments", "drawing", ...).
    fn related(&mut self, owner: &str, kind: &str) -> Result<Vec<String>> {
        let mut found = Vec::new();
        for rel in self.rels(owner)? {
            if rel.external || !(rel.kind.ends_with(&format!("/{kind}"))) {
                continue;
            }
            found.push(resolve(owner, &rel.target)?);
        }
        Ok(found)
    }
    /// Run `edits` on a part and keep the result, if there is anything to change.
    fn edit(&mut self, name: &str, bytes: &[u8], edits: BTreeMap<usize, Edit>) -> Result<()> {
        if !edits.is_empty() {
            self.put(name, apply(bytes, &edits)?);
        }
        Ok(())
    }
}

struct Sheet {
    name: String,
    path: String,
}
struct Name {
    node: usize,
    scope: String,
    text: String,
}
struct Table {
    sheet: String,
    part: String,
    name: String,
    first: (u32, u32),
    last: (u32, u32),
    reference: String,
    header: u32,
    columns: Vec<String>,
    calculated: Vec<u32>,
    /// What the part holds that this edit would have to move too.
    extended: bool,
    linked: bool,
}
enum Shape {
    Same,
    Deleted,
    /// All data rows went; Excel keeps one blank row.
    EmptyBody,
    Columns {
        at: usize,
        remove: usize,
        insert: Vec<String>,
    },
}

struct Step {
    columns: bool,
    remove: bool,
    index: u32,
    count: u32,
    target: String,
}
impl Step {
    fn limit(&self) -> u32 {
        if self.columns { 16384 } else { 1048576 }
    }
    fn axis(&self) -> &'static str {
        if self.columns { "column" } else { "row" }
    }
}

/// Rewrites references the way the cell rewriter does, one at a time.
struct Rewriter<'a> {
    edit: &'a Step,
    changes: Value,
}
impl Rewriter<'_> {
    fn one(&self, item: Value) -> Result<String> {
        let result = crate::formula::rewrite(&json!({
            "target": self.edit.target,
            "axis": self.edit.axis(),
            "index": self.edit.index,
            "count": self.edit.count,
            "remove": self.edit.remove,
            "tableChanges": self.changes,
            "items": [item],
        }))?;
        result["formulas"][0]
            .as_str()
            .map(str::to_owned)
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing rewritten formula"))
    }
    /// A formula that the reader cannot parse must not stop the edit unless it
    /// could be about the edited sheet. `None`: leave it as it is.
    fn formula(&self, formula: &str, owner: &str, flags: &[(&str, Value)]) -> Result<Option<String>> {
        let item = |text: &str| {
            let mut item = json!({"formula": text, "owner": owner});
            for (key, value) in flags {
                item[*key] = value.clone();
            }
            item
        };
        match self.one(item(formula)) {
            Ok(done) => return Ok(Some(done)),
            Err(error) => {
                crate::cancel::check()?;
                // Print titles are two ranges joined by a comma ("Sheet!$A:$A,Sheet!$1:$3");
                // the reader takes one reference at a time, so each part is moved on its own.
                let parts = union_parts(formula);
                if parts.len() > 1 {
                    let moved: Result<Vec<String>> =
                        parts.iter().map(|part| self.one(item(part))).collect();
                    if let Ok(moved) = moved {
                        return Ok(Some(moved.join(",")));
                    }
                }
                if owner == self.edit.target
                    || formula.contains(&self.edit.target)
                    // A workbook-wide name with no sheet in it could mean any sheet.
                    || (owner.is_empty() && !formula.contains('!'))
                {
                    let _ = error;
                    return Err(crate::coded(
                        "name-unsupported",
                        format!(
                            "A formula or name that refers to this sheet cannot be adjusted for the change: {}",
                            formula.chars().take(80).collect::<String>()
                        ),
                    ));
                }
            }
        }
        Ok(None)
    }
}

/// A formula split at the commas that join whole ranges: not those inside quotes or brackets.
fn union_parts(formula: &str) -> Vec<String> {
    let (mut parts, mut depth, mut quote, mut start) = (Vec::new(), 0i32, None::<char>, 0usize);
    for (at, character) in formula.char_indices() {
        if let Some(open) = quote {
            if character == open {
                quote = None;
            }
        } else if character == '\'' || character == '"' {
            quote = Some(character);
        } else if character == '(' || character == '[' {
            depth += 1;
        } else if character == ')' || character == ']' {
            depth -= 1;
        } else if character == ',' && depth == 0 {
            parts.push(formula[start..at].to_owned());
            start = at + 1;
        }
    }
    parts.push(formula[start..].to_owned());
    if parts.iter().all(|part| !part.trim().is_empty()) {
        parts
    } else {
        vec![formula.to_owned()]
    }
}

fn number(text: &str) -> Result<u32> {
    text.trim()
        .parse::<u32>()
        .map_err(|_| crate::coded("internal-invalid-field", "Invalid object position"))
}

/// Where a row or column goes, or `None` when it is deleted.
fn moved(value: u32, edit: &Step) -> Option<u32> {
    if edit.remove {
        if value >= edit.index + edit.count {
            Some(value - edit.count)
        } else if value >= edit.index {
            None
        } else {
            Some(value)
        }
    } else if value >= edit.index {
        Some(value + edit.count)
    } else {
        Some(value)
    }
}

fn read_sheets(parts: &mut Parts) -> Result<(Dom, Vec<Sheet>)> {
    let bytes = parts.read("xl/workbook.xml")?;
    let dom = Dom::parse(&bytes)?;
    let rels = parts.rels("xl/workbook.xml")?;
    let mut sheets = Vec::new();
    for node in dom.all(NS, "sheet") {
        let id = dom.attr_local(node, "id").unwrap_or("");
        let name = dom.attr(node, "name").unwrap_or("").to_owned();
        let path = rels
            .iter()
            .find(|rel| rel.id == id && !rel.external)
            .map(|rel| resolve("xl/workbook.xml", &rel.target))
            .transpose()?
            .unwrap_or_default();
        sheets.push(Sheet { name, path });
    }
    Ok((dom, sheets))
}

fn read_tables(parts: &mut Parts, sheets: &[Sheet]) -> Result<Vec<Table>> {
    let mut tables = Vec::new();
    for sheet in sheets {
        if sheet.path.is_empty() {
            continue;
        }
        for part in parts.related(&sheet.path, "table")? {
            if !parts.exists(&part) {
                return Err(crate::coded(
                    "package-refused",
                    "Workbook contains a missing table",
                ));
            }
            let dom = Dom::parse(&parts.read(&part)?)?;
            let reference = dom.attr(0, "ref").unwrap_or("").to_owned();
            let (low, high) = reference
                .split_once(':')
                .map(|(a, b)| (a.to_owned(), b.to_owned()))
                .unwrap_or((reference.clone(), reference.clone()));
            let (first, last) = (position(&low)?, position(&high)?);
            let columns: Vec<usize> = dom.all(NS, "tableColumn");
            let mut calculated = Vec::new();
            for (offset, column) in columns.iter().enumerate() {
                if !dom.kids_named(*column, NS, "calculatedColumnFormula").is_empty() {
                    calculated.push(offset as u32);
                }
            }
            tables.push(Table {
                sheet: sheet.name.clone(),
                name: dom
                    .attr(0, "displayName")
                    .filter(|name| !name.is_empty())
                    .or(dom.attr(0, "name"))
                    .unwrap_or("")
                    .to_owned(),
                first,
                last,
                reference,
                header: dom
                    .attr(0, "headerRowCount")
                    .and_then(|n| n.parse().ok())
                    .unwrap_or(1),
                columns: columns
                    .iter()
                    .map(|c| dom.attr(*c, "name").unwrap_or("").to_owned())
                    .collect(),
                calculated,
                extended: !dom.all(NS, "extLst").is_empty() || !dom.all(NS, "xmlColumnPr").is_empty(),
                linked: !parts.rels(&part)?.is_empty(),
                part,
            });
        }
    }
    Ok(tables)
}

/// How each table on the edited sheet changes.
fn table_shapes(tables: &[Table], edit: &Step, db: &Connection, version: u64, sheets: &[Sheet]) -> Result<Vec<Shape>> {
    let (index, count) = (edit.index, edit.count);
    let mut shapes = Vec::new();
    for table in tables {
        if table.sheet != edit.target {
            shapes.push(Shape::Same);
            continue;
        }
        let (a, b) = (table.first, table.last);
        if !edit.columns {
            if !edit.remove {
                shapes.push(Shape::Same);
            } else if index <= a.0 && index + count > b.0 {
                shapes.push(Shape::Deleted);
            } else if index <= a.0 && index + count > a.0 {
                return Err(crate::coded(
                    "formula-unsupported",
                    "A table header cannot be deleted on its own; keep the header and delete data rows, or delete the complete table range",
                ));
            } else if index == a.0 + 1 && index + count > b.0 {
                // Excel retains a blank insertion row when the last data row is
                // deleted. Cells shifted up from below must not be absorbed.
                let row = index + count;
                if row < 1048576 {
                    let path = sheets
                        .iter()
                        .find(|s| s.name == edit.target)
                        .map(|s| s.path.as_str())
                        .unwrap_or("");
                    // The edited version's own cells, read as the range reader
                    // reads them; the calculation snapshot is not built here.
                    let range = format!("{}{}:{}{}", column_name(a.1), row + 1, column_name(b.1), row + 1);
                    let mut after = Value::Null;
                    loop {
                        let page = crate::range::read_db(
                            db,
                            &json!({"version": version, "sheet": path, "range": range, "limit": 512, "after": after}),
                        )?;
                        let cells = page["cells"].as_array().cloned().unwrap_or_default();
                        if cells.iter().any(|pair| {
                            let cell = &pair[1];
                            !cell["value"].is_null()
                                || !cell["formula"].is_null()
                                || !cell["formulaType"].is_null()
                                || cell["dataType"] == "error"
                        }) {
                            return Err(crate::coded(
                                "table-bad-input",
                                "Empty table insertion row would absorb cells from below the table",
                            ));
                        }
                        after = page["nextCursor"].clone();
                        if after.is_null() {
                            break;
                        }
                    }
                }
                shapes.push(Shape::EmptyBody);
            } else {
                shapes.push(Shape::Same);
            }
            continue;
        }
        let internal = if edit.remove {
            index <= b.1 && index + count > a.1
        } else {
            index > a.1 && index <= b.1
        };
        if !internal {
            shapes.push(Shape::Same);
            continue;
        }
        let old = &table.columns;
        if old.len() as u32 != b.1 - a.1 + 1 {
            return Err(crate::coded(
                "formula-unsupported",
                "Inconsistent table columns",
            ));
        }
        if table.extended || table.linked {
            return Err(crate::coded(
                "object-unsupported",
                "Extended table-column metadata requires qualification",
            ));
        }
        if edit.remove {
            let start = index.saturating_sub(a.1) as usize;
            let end = ((index + count - a.1) as usize).min(old.len());
            if end - start == old.len() {
                shapes.push(Shape::Deleted);
            } else {
                shapes.push(Shape::Columns {
                    at: start,
                    remove: end - start,
                    insert: Vec::new(),
                });
            }
        } else {
            let mut names: BTreeSet<String> = old.iter().map(|c| c.to_lowercase()).collect();
            let (mut serial, mut insert) = (1, Vec::new());
            for _ in 0..count {
                while names.contains(&format!("column{serial}")) {
                    serial += 1;
                }
                let name = format!("Column{serial}");
                serial += 1;
                names.insert(name.to_lowercase());
                insert.push(name);
            }
            shapes.push(Shape::Columns {
                at: (index - a.1) as usize,
                remove: 0,
                insert,
            });
        }
    }
    Ok(shapes)
}

/// What the cell rewriter needs to know of tables that lost columns.
fn table_changes(tables: &[Table], shapes: &[Shape]) -> Vec<Value> {
    let mut changes = Vec::new();
    for (table, shape) in tables.iter().zip(shapes) {
        let (deleted, removed): (bool, Vec<String>) = match shape {
            Shape::Deleted => (true, table.columns.clone()),
            Shape::Columns { at, remove, .. } if *remove > 0 => {
                (false, table.columns[*at..*at + *remove].to_vec())
            }
            _ => continue,
        };
        changes.push(json!({
            "name": table.name,
            "sheet": table.sheet,
            "ref": table.reference,
            "deleted": deleted,
            "removed": removed,
        }));
    }
    changes
}

fn column_name(column: u32) -> String {
    crate::xml::address(0, column)
        .trim_end_matches(|c: char| c.is_ascii_digit())
        .to_owned()
}

/// Move the text of one table part: its range, its columns, its filter.
fn edit_table(parts: &mut Parts, table: &Table, shape: &Shape, reference: &str) -> Result<()> {
    let bytes = parts.read(&table.part)?;
    let dom = Dom::parse(&bytes)?;
    let mut edits: BTreeMap<usize, Edit> = BTreeMap::new();
    let mut root = vec![("ref", reference.to_owned())];
    if matches!(shape, Shape::EmptyBody) {
        root.push(("insertRow", "1".into()));
    }
    edits.insert(0, Edit::attrs(root));
    for filter in dom.all(NS, "autoFilter") {
        edits.insert(filter, Edit::attrs([("ref", reference.to_owned())]));
    }
    if let Shape::Columns { at, remove, insert } = shape {
        let container = dom
            .all(NS, "tableColumns")
            .first()
            .copied()
            .ok_or_else(|| crate::coded("table-malformed", "The table has no columns"))?;
        let columns = dom.kids_named(container, NS, "tableColumn");
        let mut id = columns
            .iter()
            .filter_map(|c| dom.attr(*c, "id").and_then(|n| n.parse::<u64>().ok()))
            .max()
            .unwrap_or(0);
        let tag = dom.nodes[columns[0]].qname.clone();
        let mut fresh = String::new();
        for name in insert {
            id += 1;
            fresh.push_str(&format!("<{tag} id=\"{id}\" name=\"{}\"/>", escape(name, true)));
        }
        if !fresh.is_empty() {
            match columns.get(*at) {
                Some(anchor) => edits.entry(*anchor).or_default().before = fresh,
                None => edits.entry(*columns.last().unwrap()).or_default().after = fresh,
            }
        }
        for column in columns.iter().skip(*at).take(*remove) {
            edits.entry(*column).or_default().change = super::xml_edit::Change::Drop;
        }
        edits.insert(
            container,
            Edit::attrs([("count", (columns.len() + insert.len() - remove).to_string())]),
        );
        for filter in dom.all(NS, "filterColumn") {
            let column = number(dom.attr(filter, "colId").unwrap_or("x"))? as usize;
            if column >= *at && column < at + remove {
                edits.insert(filter, Edit::drop());
            } else if column >= *at {
                edits.insert(
                    filter,
                    Edit::attrs([("colId", (column + insert.len() - remove).to_string())]),
                );
            }
        }
    }
    parts.edit(&table.part, &bytes, edits)
}

fn drop_table(parts: &mut Parts, table: &Table, sheets: &[Sheet]) -> Result<()> {
    if table.linked || table.extended {
        return Err(crate::coded(
            "package-unsupported",
            "Deleting a linked or extended table is not qualified",
        ));
    }
    let owner = sheets
        .iter()
        .find(|s| s.name == table.sheet)
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing table owner"))?;
    let (directory, file) = owner.path.rsplit_once('/').unwrap_or(("", &owner.path));
    let rels = format!("{directory}/_rels/{file}.rels");
    let bytes = parts.read(&rels)?;
    let dom = Dom::parse(&bytes)?;
    let mut edits = BTreeMap::new();
    for node in dom.all(PACKAGE_REL_NS, "Relationship") {
        if dom.attr(node, "TargetMode") == Some("External") {
            continue;
        }
        if resolve(&owner.path, dom.attr(node, "Target").unwrap_or("")).ok().as_deref() == Some(&table.part) {
            if !dom.attr(node, "Type").unwrap_or("").ends_with("/table") {
                return Err(crate::coded("package-malformed", "Invalid table relationship"));
            }
            edits.insert(node, Edit::drop());
        }
    }
    if edits.len() != 1 {
        return Err(crate::coded(
            "internal-missing-field",
            "Missing table relationship",
        ));
    }
    parts.edit(&rels, &bytes, edits)?;
    let types = parts.read("[Content_Types].xml")?;
    let dom = Dom::parse(&types)?;
    let mut edits = BTreeMap::new();
    for node in dom.all(TYPE_NS, "Override") {
        if dom.attr(node, "PartName") == Some(&format!("/{}", table.part)) {
            edits.insert(node, Edit::drop());
        }
    }
    parts.edit("[Content_Types].xml", &types, edits)?;
    parts.remove(&table.part);
    Ok(())
}

/// Defined names follow the rows. (The workbook's calculation settings and
/// its cached calculation order are the exporter's to keep.)
fn edit_workbook(parts: &mut Parts, rewrite: &Rewriter, names: &[Name]) -> Result<()> {
    let bytes = parts.read("xl/workbook.xml")?;
    let mut edits: BTreeMap<usize, Edit> = BTreeMap::new();
    for name in names {
        if let Some(next) = rewrite.formula(&name.text, &name.scope, &[("name", json!(true))])?
            && next != name.text
        {
            edits.insert(name.node, Edit::text(next));
        }
    }
    parts.edit("xl/workbook.xml", &bytes, edits)
}

/// One cell reference such as "B7", moved. `None`: its row or column is deleted.
fn move_cell(reference: &str, edit: &Step) -> Result<Option<String>> {
    let (row, column) = position(reference)
        .map_err(|_| crate::coded("object-bad-input", "Object has no valid cell"))?;
    let (axis, other) = if edit.columns { (column, row) } else { (row, column) };
    let Some(next) = moved(axis, edit) else {
        return Ok(None);
    };
    if next >= edit.limit() {
        return Err(crate::coded("object-limit", "Object exceeds the worksheet"));
    }
    Ok(Some(if edit.columns {
        crate::xml::address(other, next)
    } else {
        crate::xml::address(next, other)
    }))
}

/// Notes and threaded comments name their cell; a deleted cell takes them along.
fn move_comments(parts: &mut Parts, part: &str, tag: &str, ns: Option<&[u8]>, edit: &Step) -> Result<()> {
    let bytes = parts.read(part)?;
    let dom = Dom::parse(&bytes)?;
    let mut edits = BTreeMap::new();
    for node in 0..dom.nodes.len() {
        if dom.nodes[node].name != tag || ns.is_some_and(|ns| dom.nodes[node].ns != ns) {
            continue;
        }
        let Some(reference) = dom.attr(node, "ref") else {
            continue;
        };
        match move_cell(reference, edit)? {
            None => {
                edits.insert(node, Edit::drop());
            }
            Some(next) if next != reference => {
                edits.insert(node, Edit::attrs([("ref", next)]));
            }
            Some(_) => {}
        }
    }
    parts.edit(part, &bytes, edits)
}

/// The tag and its text for `<x:Row>3</x:Row>`: where the number is, and its value.
fn tag_value(block: &str, local: &str) -> Option<(std::ops::Range<usize>, String)> {
    let mut from = 0;
    while let Some(found) = block[from..].find('<') {
        let open = from + found + 1;
        let rest = &block[open..];
        let name_end = rest.find(|c: char| c == '>' || c.is_whitespace() || c == '/').unwrap_or(rest.len());
        let name = &rest[..name_end];
        if name.rsplit(':').next() == Some(local) && rest[name_end..].starts_with('>') {
            let start = open + name_end + 1;
            let end = block[start..].find("</")? + start;
            return Some((start..end, block[start..end].trim().to_owned()));
        }
        from = open;
    }
    None
}
fn count_tags(text: &str, local: &str) -> usize {
    let mut count = 0;
    let mut from = 0;
    while let Some(found) = text[from..].find('<') {
        let open = from + found + 1;
        let rest = &text[open..];
        let name_end = rest.find(|c: char| c == '>' || c.is_whitespace() || c == '/').unwrap_or(rest.len());
        let name = &rest[..name_end];
        let wanted = name.rsplit(':').next() == Some(local) && !name.starts_with('/');
        if wanted && (rest[name_end..].starts_with('>') || rest[name_end..].starts_with(char::is_whitespace)) {
            count += 1;
        }
        from = open;
    }
    count
}

/// The note boxes of a sheet. This markup is not always well-formed XML (its
/// root is named "xml"), so it is changed as text. A note's box records its
/// cell as a row and a column and its place as a start and an end cell.
fn move_note_boxes(text: &str, edit: &Step) -> Result<String> {
    let axis = if edit.columns { "Column" } else { "Row" };
    if text.trim().is_empty() {
        return Err(crate::coded(
            "object-not-movable",
            "Close and reopen the workbook to move its notes",
        ));
    }
    if count_tags(text, axis) != count_tags(text, "ClientData") {
        return Err(crate::coded(
            "object-not-movable",
            "The note drawing has an unexpected layout",
        ));
    }
    let mut out = String::with_capacity(text.len());
    let mut at = 0;
    while let Some(found) = text[at..].find('<') {
        let open = at + found;
        let rest = &text[open + 1..];
        let name_end = rest.find(|c: char| c == '>' || c.is_whitespace() || c == '/').unwrap_or(rest.len());
        let name = &rest[..name_end];
        if name.rsplit(':').next() != Some("shape") || name.starts_with('/') {
            out.push_str(&text[at..open + 1]);
            at = open + 1;
            continue;
        }
        let close = format!("</{name}>");
        let end = text[open..]
            .find(&close)
            .map(|n| open + n + close.len())
            .ok_or_else(|| crate::coded("object-not-movable", "The note drawing has an unexpected layout"))?;
        let block = &text[open..end];
        if block[name.len() + 1..block.len() - close.len()].contains(&format!("<{name}")) {
            return Err(crate::coded(
                "object-not-movable",
                "The note drawing has an unexpected layout",
            ));
        }
        out.push_str(&text[at..open]);
        if let Some(changed) = move_note_box(block, axis, edit)? {
            out.push_str(&changed);
        }
        at = end;
    }
    out.push_str(&text[at..]);
    Ok(out)
}
fn move_note_box(block: &str, axis: &str, edit: &Step) -> Result<Option<String>> {
    let Some((span, value)) = tag_value(block, axis) else {
        return Ok(Some(block.to_owned()));
    };
    let value = value
        .parse::<u32>()
        .map_err(|_| crate::coded("internal-invalid-field", "Invalid note position"))?;
    let Some(next) = moved(value, edit) else {
        return Ok(None);
    };
    let mut changes = vec![(span, next.to_string())];
    if let Some((span, body)) = tag_value(block, "Anchor") {
        // left column, offset, top row, offset, right column, offset, bottom row, offset
        let mut numbers = Vec::new();
        for part in body.split(',') {
            numbers.push(
                part.trim()
                    .parse::<u32>()
                    .map_err(|_| crate::coded("internal-invalid-field", "Invalid note box"))?,
            );
        }
        if numbers.len() != 8 {
            return Err(crate::coded("internal-invalid-field", "Invalid note box"));
        }
        for at in if edit.columns { [0, 4] } else { [2, 6] } {
            // A box edge inside the deleted rows lands where they were.
            numbers[at] = if edit.remove {
                if numbers[at] >= edit.index + edit.count {
                    numbers[at] - edit.count
                } else if numbers[at] >= edit.index {
                    edit.index
                } else {
                    numbers[at]
                }
            } else if numbers[at] >= edit.index {
                numbers[at] + edit.count
            } else {
                numbers[at]
            };
        }
        changes.push((
            span,
            numbers.iter().map(u32::to_string).collect::<Vec<_>>().join(", "),
        ));
    }
    changes.sort_by_key(|(span, _)| span.start);
    let mut result = block.to_owned();
    for (span, text) in changes.into_iter().rev() {
        result.replace_range(span, &text);
    }
    Ok(Some(result))
}

/// Pictures, charts and shapes follow their cells the way Excel's placement
/// setting says: "move and size" (the default) moves each corner on its own;
/// "move but don't size" keeps the size; "don't move" stays put.
fn move_drawing(parts: &mut Parts, part: &str, rewrite: &Rewriter, edit: &Step) -> Result<()> {
    let bytes = parts.read(part)?;
    let dom = Dom::parse(&bytes)?;
    if !dom.is(0, DRAW, "wsDr") {
        return Err(crate::coded("object-malformed", "Invalid drawing relationship target"));
    }
    let axis = if edit.columns { "col" } else { "row" };
    let offset = if edit.columns { "colOff" } else { "rowOff" };
    let mut edits: BTreeMap<usize, Edit> = BTreeMap::new();
    // (element holding the number, the number, the offset into that cell)
    let corner = |marker: usize| -> Result<(usize, u32, u64)> {
        let place = dom
            .kids_named(marker, DRAW, axis)
            .first()
            .copied()
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid drawing marker"))?;
        let inside = dom
            .kids_named(marker, DRAW, offset)
            .first()
            .map(|n| dom.nodes[*n].text.trim().parse::<i64>().unwrap_or(0).max(0) as u64)
            .unwrap_or(0);
        Ok((place, number(&dom.nodes[place].text)?, inside))
    };
    let set = |edits: &mut BTreeMap<usize, Edit>, place: usize, value: u32| -> Result<()> {
        if value >= edit.limit() {
            return Err(crate::coded("object-limit", "Drawing marker exceeds worksheet"));
        }
        edits.insert(place, Edit::text(value.to_string()));
        Ok(())
    };
    let deleted_through = || {
        crate::coded(
            "object-unsupported",
            "Deleting through an image anchor is not qualified",
        )
    };
    for anchor in 0..dom.nodes.len() {
        let two = dom.is(anchor, DRAW, "twoCellAnchor");
        if !two && !dom.is(anchor, DRAW, "oneCellAnchor") {
            continue;
        }
        let from = dom.kids_named(anchor, DRAW, "from").first().copied();
        let to = dom.kids_named(anchor, DRAW, "to").first().copied();
        let Some(from) = from else {
            return Err(crate::coded("internal-invalid-field", "Invalid drawing marker"));
        };
        let mode = if two { dom.attr(anchor, "editAs").unwrap_or("twoCell") } else { "oneCell" };
        match to {
            Some(to) if matches!(mode, "oneCell" | "absolute") => {
                // The object keeps its size or its place, so it is either wholly
                // before the change, wholly after it, or it needs the heights of
                // the rows it covers to be placed, which is not done yet.
                let (first, first_at, _) = corner(from)?;
                let (last, last_at, last_inside) = corner(to)?;
                let before = last_at < edit.index || (last_at == edit.index && last_inside == 0);
                let after = first_at >= edit.index + if edit.remove { edit.count } else { 0 };
                if before {
                } else if after && mode == "oneCell" {
                    for (place, value) in [(first, first_at), (last, last_at)] {
                        set(&mut edits, place, moved(value, edit).ok_or_else(deleted_through)?)?;
                    }
                } else {
                    return Err(crate::coded(
                        "object-unsupported",
                        "A picture or shape that keeps its size or position cannot yet follow rows or columns inserted or deleted beside or through it",
                    ));
                }
            }
            _ => {
                // Each corner moves with its own cell.
                for marker in [Some(from), to].into_iter().flatten() {
                    let (place, value, _) = corner(marker)?;
                    let next = moved(value, edit).ok_or_else(deleted_through)?;
                    if next != value {
                        set(&mut edits, place, next)?;
                    }
                }
            }
        }
    }
    // A text box can show the value of a cell.
    for node in 0..dom.nodes.len() {
        if let Some(link) = dom.attr(node, "textlink").filter(|link| !link.trim().is_empty()) {
            match rewrite.formula(link, &edit.target, &[])? {
                Some(next) if next.contains("#REF!") => {
                    return Err(crate::coded(
                        "object-range-deleted",
                        "Deleting an entire object range is not qualified",
                    ));
                }
                Some(next) if next != link => {
                    edits.insert(node, Edit::attrs([("textlink", next)]));
                }
                _ => {}
            }
        }
    }
    parts.edit(part, &bytes, edits)
}

/// A chart names its data by sheet and range; those follow the rows too. The
/// numbers the chart has stored are refreshed by Excel when it opens the file.
fn move_chart(parts: &mut Parts, part: &str, rewrite: &Rewriter, edit: &Step) -> Result<()> {
    let bytes = parts.read(part)?;
    if !String::from_utf8_lossy(&bytes).contains("<") {
        return Ok(());
    }
    let dom = Dom::parse(&bytes)?;
    let mut edits = BTreeMap::new();
    for node in 0..dom.nodes.len() {
        let n = &dom.nodes[node];
        if n.name != "f" || ![CHART_NS, C15_NS, CHARTEX_NS].contains(&n.ns.as_slice()) || n.text.trim().is_empty() {
            continue;
        }
        let Some(next) = rewrite.formula(&n.text, &edit.target, &[]).map_err(|error| {
            crate::rephrase("object-unsupported", error, |said| {
                format!("Object reference is not qualified: {said}")
            })
        })?
        else {
            continue;
        };
        if next.contains("#REF!") {
            return Err(crate::coded(
                "object-range-deleted",
                "Deleting all of a chart's data is not supported; delete the chart first",
            ));
        }
        if next != n.text {
            edits.insert(node, Edit::text(next));
        }
    }
    parts.edit(part, &bytes, edits)
}

/// Formulas and the sort inside a table part follow the rows too.
fn move_table_formulas(parts: &mut Parts, table: &Table, rewrite: &Rewriter, edit: &Step) -> Result<()> {
    let bytes = parts.read(&table.part)?;
    let dom = Dom::parse(&bytes)?;
    let mut edits = BTreeMap::new();
    let (top, _) = crate::bulk::bounds(&table.reference)?;
    let data_cell = crate::xml::address(top.0 + 1, top.1);
    for node in 0..dom.nodes.len() {
        let n = &dom.nodes[node];
        if n.ns != NS {
            continue;
        }
        if ["calculatedColumnFormula", "totalsRowFormula"].contains(&n.name.as_str()) {
            if n.text.trim().is_empty() {
                continue;
            }
            // Shorthand such as [@Price] means this table; the rewriter reads it
            // from a data cell's point of view and turns a deleted column into #REF!.
            let Some(next) = rewrite.formula(&n.text, &table.sheet, &[("reference", json!(data_cell))])? else {
                continue;
            };
            if next != n.text {
                edits.insert(node, Edit::text(next));
            }
        } else if ["sortState", "sortCondition"].contains(&n.name.as_str()) {
            let Some(reference) = dom.attr(node, "ref") else {
                continue;
            };
            let mut next = Vec::new();
            for part in reference.split_whitespace() {
                let moved = rewrite.formula(part, &table.sheet, &[])?.unwrap_or_else(|| part.to_owned());
                if moved.contains("#REF!") {
                    return Err(crate::coded(
                        "object-range-deleted",
                        "Deleting an entire object range is not qualified",
                    ));
                }
                next.push(moved);
            }
            if next.join(" ") != reference {
                edits.insert(node, Edit::attrs([("ref", next.join(" "))]));
            }
        }
    }
    let _ = edit;
    parts.edit(&table.part, &bytes, edits)
}

pub(crate) fn plan(root: &Path, db: &Connection, version: u64, command: &Value) -> Result<Plan> {
    let operation = field(command, "operation")?;
    if !["insertRow", "deleteRow", "insertColumn", "deleteColumn"].contains(&operation) {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid structural planning operation",
        ));
    }
    let columns = operation.ends_with("Column");
    let limit: u64 = if columns { 16384 } else { 1048576 };
    let index = command["index"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid inserted row range"))?;
    let count = match command.get("count") {
        Some(value) => value
            .as_u64()
            .filter(|n| *n >= 1)
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid inserted row range"))?,
        None => 1,
    };
    if index + count > limit {
        return Err(crate::coded(
            "formula-limit",
            "Structural edit exceeds the XLSX boundary",
        ));
    }
    let edit = Step {
        columns,
        remove: operation.starts_with("delete"),
        index: index as u32,
        count: count as u32,
        target: field(command, "sheet")?.to_owned(),
    };
    let metadata = local(root, "metadata.xlsx")?;
    let mut parts = Parts {
        zip: ZipArchive::new(crate::cancel::Reader(File::open(&metadata).map_err(err)?))
            .map_err(err)?,
        changed: BTreeMap::new(),
        removed: BTreeSet::new(),
    };
    let (book, sheets) = read_sheets(&mut parts)?;
    if !sheets.iter().any(|s| s.name == edit.target) {
        return Err(crate::coded("internal-unknown-reference", "Unknown worksheet"));
    }
    let mut names = Vec::new();
    for node in book.all(NS, "definedName") {
        let scope = match book.attr(node, "localSheetId") {
            None => String::new(),
            Some(raw) => raw
                .parse::<usize>()
                .ok()
                .filter(|_| raw.bytes().all(|b| b.is_ascii_digit()))
                .and_then(|at| sheets.get(at))
                .map(|s| s.name.clone())
                .ok_or_else(|| {
                    crate::coded(
                        "workbook-unsupported",
                        "Structural reference syntax is not qualified for native preservation: metadata binding",
                    )
                })?,
        };
        names.push(Name {
            node,
            scope,
            text: book.nodes[node].text.clone(),
        });
    }
    let tables = read_tables(&mut parts, &sheets)?;
    let shapes = table_shapes(&tables, &edit, db, version, &sheets)?;
    let changes = table_changes(&tables, &shapes);
    for change in &changes {
        if change["deleted"] == true
            && names.iter().any(|name| {
                book.attr(name.node, "name")
                    .is_some_and(|n| n.eq_ignore_ascii_case(change["name"].as_str().unwrap_or("")))
            })
        {
            return Err(crate::coded(
                "formula-unsupported",
                "Deleted table name conflicts with a defined name",
            ));
        }
    }
    let rewrite = Rewriter {
        edit: &edit,
        changes: Value::Array(changes.clone()),
    };
    // Tables: the range moves with its rows and columns.
    let mut headers = Vec::new();
    let mut calculated = Vec::new();
    for (table, shape) in tables.iter().zip(&shapes) {
        match shape {
            Shape::Deleted => drop_table(&mut parts, table, &sheets)?,
            shape => {
                let reference = match shape {
                    Shape::EmptyBody => format!(
                        "{}{}:{}{}",
                        column_name(table.first.1),
                        table.first.0 + 1,
                        column_name(table.last.1),
                        table.first.0 + 2
                    ),
                    _ if table.sheet == edit.target => rewrite
                        .formula(&table.reference, &table.sheet, &[("nameBoundary", json!(true))])?
                        .unwrap_or_else(|| table.reference.clone()),
                    _ => table.reference.clone(),
                };
                if reference != table.reference || !matches!(shape, Shape::Same) {
                    edit_table(&mut parts, table, shape, &reference)?;
                }
                if let Shape::Columns { at, insert, .. } = shape {
                    for (i, name) in insert.iter().enumerate() {
                        headers.push(json!({
                            "sheet": table.sheet,
                            "cell": format!("{}{}", column_name(table.first.1 + *at as u32 + i as u32), table.first.0 + 1),
                            "value": name,
                        }));
                    }
                }
            }
        }
        // Only a column the table declares as calculated continues in a new row.
        if table.sheet == edit.target {
            for offset in &table.calculated {
                let column = table.first.1 + offset;
                if column <= table.last.1 {
                    calculated.push((table.first.0 + table.header, table.last.0, column, column));
                }
            }
        }
    }
    edit_workbook(&mut parts, &rewrite, &names)?;
    // Objects on the edited sheet.
    let target_path = sheets
        .iter()
        .find(|s| s.name == edit.target)
        .map(|s| s.path.clone())
        .unwrap_or_default();
    for part in parts.related(&target_path, "comments")? {
        move_comments(&mut parts, &part, "comment", Some(NS), &edit)?;
    }
    for part in parts.related(&target_path, "threadedComment")? {
        move_comments(&mut parts, &part, "threadedComment", None, &edit)?;
    }
    for part in parts.related(&target_path, "vmlDrawing")? {
        let text = String::from_utf8(parts.read(&part)?)
            .map_err(|_| crate::coded("object-not-movable", "The note drawing has an unexpected layout"))?;
        let moved = move_note_boxes(&text, &edit)?;
        if moved != text {
            parts.put(&part, moved.into_bytes());
        }
    }
    for part in parts.related(&target_path, "drawing")? {
        move_drawing(&mut parts, &part, &rewrite, &edit)?;
    }
    for part in parts.names() {
        if part.starts_with("xl/charts/") && part.ends_with(".xml") && part.matches('/').count() == 2 {
            move_chart(&mut parts, &part, &rewrite, &edit)?;
        }
    }
    for (table, shape) in tables.iter().zip(&shapes) {
        if !matches!(shape, Shape::Deleted) {
            move_table_formulas(&mut parts, table, &rewrite, &edit)?;
        }
    }
    let mut replacements = BTreeMap::new();
    for (name, bytes) in &parts.changed {
        // A part that is not text cannot have been planned here.
        replacements.insert(
            name.clone(),
            String::from_utf8(bytes.clone())
                .map_err(|_| crate::coded("package-malformed", "Package part is not text"))?,
        );
    }
    Ok(Plan {
        replacements,
        removed: parts.removed.into_iter().collect(),
        table_changes: changes,
        headers,
        calculated,
        sheets: sheets.into_iter().map(|s| (s.name, s.path)).collect(),
    })
}
