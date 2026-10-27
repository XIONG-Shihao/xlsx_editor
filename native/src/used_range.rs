//! Source facts and selected-version worksheet extent. Display dimensions are unrelated.
use crate::{Result, err, field};
use quick_xml::{NsReader, events::Event, name::ResolveResult};
use rusqlite::{Connection, OpenFlags, OptionalExtension, params};
use serde::{Deserialize, Serialize};
use serde_json::{Value, json};
use std::{
    borrow::Cow,
    collections::BTreeMap,
    fs::File,
    io::{BufReader, Read},
    path::Path,
};

#[path = "used_range_styles.rs"]
mod source_styles;
pub(crate) const FACTS_SCHEMA: &str = workbook_format!(SCHEMA_USED_RANGE_FACTS_V2);
#[cfg(test)]
const LEGACY_FACTS_SCHEMA: &str = workbook_format!(SCHEMA_USED_RANGE_FACTS_V1);
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(crate) struct Bounds {
    pub min_row: u32,
    pub min_column: u32,
    pub max_row: u32,
    pub max_column: u32,
}

pub(crate) fn schema(db: &Connection) -> Result<()> {
    db.execute_batch("CREATE TABLE used_range_history(version INTEGER NOT NULL,sheet TEXT NOT NULL,min_r INTEGER NOT NULL,min_c INTEGER NOT NULL,max_r INTEGER NOT NULL,max_c INTEGER NOT NULL,PRIMARY KEY(version,sheet)) WITHOUT ROWID; CREATE TABLE used_range_axes(id INTEGER PRIMARY KEY,sheet TEXT NOT NULL,axis TEXT NOT NULL,first INTEGER NOT NULL,last INTEGER NOT NULL,style INTEGER NOT NULL,extent INTEGER NOT NULL,removable INTEGER NOT NULL); CREATE INDEX used_range_axes_coordinates ON used_range_axes(sheet,axis,first,last);").map_err(err)?;
    guard_schema(db)
}
pub(crate) fn guard_schema(db: &Connection) -> Result<()> {
    db.execute_batch("CREATE TABLE IF NOT EXISTS used_range_guards(version INTEGER NOT NULL,sheet TEXT NOT NULL,reasons TEXT NOT NULL,PRIMARY KEY(version,sheet)) WITHOUT ROWID;").map_err(err)
}
pub(crate) fn guard_reasons(db: &Connection, version: u64, sheet: &str) -> Result<Vec<Value>> {
    let present: bool = db
        .query_row(
            "SELECT EXISTS(SELECT 1 FROM sqlite_master WHERE name='used_range_guards')",
            [],
            |r| r.get(0),
        )
        .map_err(err)?;
    if !present {
        return Ok(Vec::new());
    }
    let reasons:Option<String>=db.prepare_cached("WITH RECURSIVE a(id,depth) AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL) SELECT reasons FROM used_range_guards h JOIN a ON h.version=a.id WHERE h.sheet=?2 ORDER BY depth LIMIT 1").map_err(crate::err)?.query_row(params![version,sheet],|r|r.get(0)).optional().map_err(err)?;
    let reasons: Vec<Value> = reasons
        .map(|text| serde_json::from_str(&text).map_err(err))
        .transpose()?
        .unwrap_or_default();
    if reasons.len() > 16
        || reasons.iter().any(|r| {
            r.as_object().is_none_or(|o| o.len() != 2)
                || r["code"]
                    .as_str()
                    .is_none_or(|s| s.is_empty() || s.len() > 64)
                || r["message"]
                    .as_str()
                    .is_none_or(|s| s.is_empty() || s.len() > 1024)
        })
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid inherited UsedRange guard",
        ));
    }
    Ok(reasons)
}
pub(crate) fn store_guards(
    db: &Connection,
    version: u64,
    sheet: &str,
    reasons: &[Value],
) -> Result<()> {
    if reasons.is_empty() {
        return Ok(());
    }
    guard_schema(db)?;
    let mut unique = BTreeMap::new();
    for reason in reasons {
        let code = field(reason, "code")?;
        let message = field(reason, "message")?;
        if code.is_empty() || code.len() > 64 || message.is_empty() || message.len() > 1024 {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid inherited UsedRange guard",
            ));
        }
        unique
            .entry(code.to_owned())
            .or_insert_with(|| reason.clone());
    }
    unique.entry("private-rewrite-source-incomplete".into()).or_insert_with(||json!({"code":"private-rewrite-source-incomplete","message":"An accepted predecessor has an unqualified extent profile; a private rewrite cannot grant complete geometry."}));
    let reasons = unique.into_values().take(16).collect::<Vec<_>>();
    db.execute(
        "INSERT OR REPLACE INTO used_range_guards VALUES(?1,?2,?3)",
        params![
            version,
            sheet,
            serde_json::to_string(&reasons).map_err(err)?
        ],
    )
    .map_err(err)?;
    Ok(())
}
fn history_present(db: &Connection) -> Result<bool> {
    db.prepare_cached("SELECT EXISTS(SELECT 1 FROM sqlite_master WHERE name='used_range_history')")
        .map_err(crate::err)?
        .query_row([], |r| r.get(0))
        .map_err(err)
}
pub(crate) fn history(db: &Connection, version: u64, sheet: &str) -> Result<Option<Bounds>> {
    if !history_present(db)? {
        return Ok(None);
    }
    db.prepare_cached("WITH RECURSIVE a(id,depth) AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL) SELECT min_r,min_c,max_r,max_c FROM used_range_history h JOIN a ON h.version=a.id WHERE h.sheet=?2 ORDER BY depth LIMIT 1").map_err(crate::err)?.query_row(params![version,sheet],|r|Ok(Bounds{min_row:r.get(0)?,min_column:r.get(1)?,max_row:r.get(2)?,max_column:r.get(3)?})).optional().map_err(err)
}
fn content(cell: &Value) -> bool {
    cell.get("formula").is_some_and(|v| !v.is_null())
        || cell.get("formulaType").is_some_and(|v| !v.is_null())
        || !cell["value"].is_null()
}
fn has_spill_metadata(cell: &Value) -> bool {
    cell.get("spill").is_some_and(|value| !value.is_null())
}
fn formatted(cell: &Value) -> bool {
    cell["styleId"].as_u64().is_some_and(|s| s != 0)
        || crate::range::STYLE_FIELDS.iter().any(|key| {
            cell.get(*key).is_some_and(|v| match *key {
                "bold" | "italic" | "underline" | "wrapText" => v == true,
                "numberFormat" => v.as_str().is_some_and(|s| s != "general"),
                _ => !v.is_null(),
            })
        })
}
/// A cumulative binding follows only this accepted ancestor chain. Clears do
/// not extend it, and its own row survives collapsed private intermediates.
pub(crate) fn record(
    db: &Connection,
    version: i64,
    sheet: &str,
    row: u32,
    col: u32,
    cell: &Value,
) -> Result<()> {
    if !history_present(db)? {
        return Ok(());
    }
    let parent: Option<u64> = db
        .query_row("SELECT parent FROM versions WHERE id=?1", [version], |r| {
            r.get(0)
        })
        .map_err(err)?;
    let mut bounds = history(db, version as u64, sheet)?;
    if bounds.is_none()
        && let Some(parent) = parent
    {
        bounds = history(db, parent, sheet)?;
    }
    if content(cell) || formatted(cell) {
        add(&mut bounds, Bounds::point(row, col));
    }
    if let Some(b) = bounds {
        db.execute(
            "INSERT OR REPLACE INTO used_range_history VALUES(?1,?2,?3,?4,?5,?6)",
            params![
                version,
                sheet,
                b.min_row,
                b.min_column,
                b.max_row,
                b.max_column
            ],
        )
        .map_err(err)?;
    }
    Ok(())
}
pub(crate) fn retain(db: &Connection, keep: &std::collections::BTreeSet<u64>) -> Result<()> {
    if !history_present(db)? {
        return Ok(());
    }
    let sheets: Vec<String> = db
        .prepare_cached("SELECT DISTINCT sheet FROM used_range_history")
        .map_err(err)?
        .query_map([], |r| r.get(0))
        .map_err(err)?
        .collect::<std::result::Result<_, _>>()
        .map_err(err)?;
    for version in keep {
        for sheet in &sheets {
            if let Some(b) = history(db, *version, sheet)? {
                db.execute(
                    "INSERT OR REPLACE INTO used_range_history VALUES(?1,?2,?3,?4,?5,?6)",
                    params![
                        version,
                        sheet,
                        b.min_row,
                        b.min_column,
                        b.max_row,
                        b.max_column
                    ],
                )
                .map_err(err)?;
            }
        }
    }
    let guards: bool = db
        .query_row(
            "SELECT EXISTS(SELECT 1 FROM sqlite_master WHERE name='used_range_guards')",
            [],
            |r| r.get(0),
        )
        .map_err(err)?;
    if guards {
        let sheets = db
            .prepare_cached("SELECT DISTINCT sheet FROM used_range_guards")
            .map_err(err)?
            .query_map([], |r| r.get::<_, String>(0))
            .map_err(err)?
            .collect::<std::result::Result<Vec<_>, _>>()
            .map_err(err)?;
        for version in keep {
            for sheet in &sheets {
                let reasons = guard_reasons(db, *version, sheet)?;
                if !reasons.is_empty() {
                    store_guards(db, *version, sheet, &reasons)?;
                }
            }
        }
    }
    Ok(())
}
pub(crate) fn prune(db: &Connection) -> Result<()> {
    if history_present(db)? {
        db.execute(
            "DELETE FROM used_range_history WHERE version NOT IN(SELECT id FROM versions)",
            [],
        )
        .map_err(err)?;
    }
    let guards: bool = db
        .query_row(
            "SELECT EXISTS(SELECT 1 FROM sqlite_master WHERE name='used_range_guards')",
            [],
            |r| r.get(0),
        )
        .map_err(err)?;
    if guards {
        db.execute(
            "DELETE FROM used_range_guards WHERE version NOT IN(SELECT id FROM versions)",
            [],
        )
        .map_err(err)?;
    }
    Ok(())
}

fn bounded_part(zip: &mut zip::ZipArchive<File>, name: &str) -> Result<Vec<u8>> {
    let mut member = zip.by_name(name).map_err(err)?;
    if member.size() > 4 * 1024 * 1024 {
        return Err(crate::coded(
            "layout-limit",
            "UsedRange workbook metadata exceeds its bound",
        ));
    }
    let mut bytes = Vec::new();
    member
        .by_ref()
        .take(4 * 1024 * 1024 + 1)
        .read_to_end(&mut bytes)
        .map_err(err)?;
    if bytes.len() > 4 * 1024 * 1024 {
        return Err(crate::coded(
            "layout-limit",
            "UsedRange workbook metadata exceeds its bound",
        ));
    }
    Ok(bytes)
}
fn logical_sheet(root: &Path, target: &str) -> Result<String> {
    let mut zip =
        zip::ZipArchive::new(File::open(crate::local(root, "metadata.xlsx")?).map_err(err)?)
            .map_err(err)?;
    let relationships = bounded_part(&mut zip, "xl/_rels/workbook.xml.rels")?;
    let workbook = bounded_part(&mut zip, "xl/workbook.xml")?;
    let mut reader = NsReader::from_reader(relationships.as_slice());
    let mut buffer = Vec::new();
    let mut ids = std::collections::BTreeSet::new();
    loop {
        crate::cancel::check()?;
        let (ns, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        match event {
            Event::Start(ref e) | Event::Empty(ref e)
                if ns
                    == ResolveResult::Bound(quick_xml::name::Namespace(
                        b"http://schemas.openxmlformats.org/package/2006/relationships",
                    ))
                    && e.local_name().as_ref() == b"Relationship" =>
            {
                let a = crate::xml::attrs(e)?;
                if a.get("Type").is_some_and(|s| s.ends_with("/worksheet"))
                    && a.get("TargetMode").is_none_or(|s| s == "Internal")
                {
                    let raw = a.get("Target").ok_or_else(|| {
                        crate::coded(
                            "internal-missing-field",
                            "Missing worksheet relationship target",
                        )
                    })?;
                    let mut parts = if raw.starts_with('/') {
                        Vec::new()
                    } else {
                        vec!["xl"]
                    };
                    for p in raw.trim_start_matches('/').split('/') {
                        match p {
                            "." => {}
                            ".." => {
                                parts.pop().ok_or_else(|| {
                                    crate::coded(
                                        "layout-unsupported",
                                        "Worksheet relationship escapes package",
                                    )
                                })?;
                            }
                            "" => {
                                return Err(crate::coded(
                                    "layout-malformed",
                                    "Invalid worksheet relationship target",
                                ));
                            }
                            _ => parts.push(p),
                        }
                    }
                    if parts.join("/") == target {
                        ids.insert(
                            a.get("Id")
                                .ok_or_else(|| {
                                    crate::coded(
                                        "internal-missing-field",
                                        "Missing worksheet relationship id",
                                    )
                                })?
                                .clone(),
                        );
                    }
                }
            }
            Event::DocType(_) => {
                return Err(crate::coded(
                    "layout-malformed",
                    "Worksheet XML entities are unsupported",
                ));
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if ids.len() != 1 {
        return Err(crate::coded(
            "layout-ambiguous",
            "Ambiguous UsedRange worksheet relationship",
        ));
    }
    let mut reader = NsReader::from_reader(workbook.as_slice());
    let mut buffer = Vec::new();
    let mut result = None;
    loop {
        crate::cancel::check()?;
        let (ns, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        match event {
            Event::Start(ref e) | Event::Empty(ref e)
                if ns == ResolveResult::Bound(quick_xml::name::Namespace(crate::xml::NS))
                    && e.local_name().as_ref() == b"sheet" =>
            {
                let a = crate::xml::attrs(e)?;
                let mut id = None;
                for attribute in e.attributes() {
                    let attribute = attribute.map_err(err)?;
                    let (binding, name) = reader.resolve_attribute(attribute.key);
                    if binding
                        == ResolveResult::Bound(quick_xml::name::Namespace(
                            b"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                        ))
                        && name.as_ref() == b"id"
                    {
                        id = Some(attribute.unescape_value().map_err(err)?.into_owned());
                    }
                }
                if id.as_ref().is_some_and(|id| ids.contains(id)) {
                    if result.is_some() {
                        return Err(crate::coded(
                            "layout-ambiguous",
                            "Ambiguous UsedRange logical worksheet",
                        ));
                    }
                    let name = a.get("name").ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing logical worksheet name")
                    })?;
                    if name.is_empty() || name.len() > 124 || name.encode_utf16().count() > 31 {
                        return Err(crate::coded(
                            "internal-invalid-field",
                            "Invalid UsedRange worksheet name",
                        ));
                    }
                    result = Some(name.clone());
                }
            }
            Event::DocType(_) => {
                return Err(crate::coded(
                    "layout-malformed",
                    "Worksheet XML entities are unsupported",
                ));
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    result.ok_or_else(|| {
        crate::coded(
            "layout-unsupported",
            "UsedRange logical worksheet is unavailable",
        )
    })
}

fn validate_cell_styles(db: &Connection, version: u64, sheet: &str, styles: u64) -> Result<()> {
    let max:u64=db.prepare_cached("WITH RECURSIVE a(id,depth) AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL), ranked AS(SELECT p.*,row_number() OVER(PARTITION BY r,c ORDER BY depth) priority FROM changes p JOIN a ON p.version=a.id WHERE p.sheet=?2), o AS(SELECT * FROM ranked WHERE priority=1) SELECT max(style) FROM(SELECT coalesce(max(style),0) style FROM cells WHERE sheet=?2 UNION ALL SELECT coalesce(max(coalesce(json_extract(o.data,'$.styleId'),c.style,0)),0) FROM o LEFT JOIN cells c ON c.sheet=?2 AND c.r=o.r AND c.c=o.c)").map_err(crate::err)?.query_row(params![version,sheet],|r|r.get(0)).map_err(err)?;
    if max >= styles {
        return Err(crate::coded(
            "layout-unsupported",
            "UsedRange cell style index is outside its source style table",
        ));
    }
    Ok(())
}

const SELECTED_CELLS: &str = "WITH RECURSIVE a(id,depth) AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL), ranked AS(SELECT p.*,row_number() OVER(PARTITION BY r,c ORDER BY depth) priority FROM changes p JOIN a ON p.version=a.id WHERE p.sheet=?2), o AS(SELECT * FROM ranked WHERE priority=1), k AS(SELECT r,c FROM cells WHERE sheet=?2 UNION SELECT r,c FROM o) SELECT k.r,k.c,c.kind,c.raw,c.inline,c.formula,c.formula_attrs,c.style,s.value,o.data FROM k LEFT JOIN cells c ON c.sheet=?2 AND c.r=k.r AND c.c=k.c LEFT JOIN strings s ON c.kind='s' AND s.id=CAST(c.raw AS INTEGER) LEFT JOIN o ON o.r=k.r AND o.c=k.c ORDER BY k.r,k.c";
fn visit_cells(
    db: &Connection,
    version: u64,
    sheet: &str,
    mut visit: impl FnMut(u32, u32, Value) -> Result<()>,
) -> Result<()> {
    let mut statement = db.prepare_cached(SELECTED_CELLS).map_err(err)?;
    let mut rows = statement.query(params![version, sheet]).map_err(err)?;
    while let Some(row) = rows.next().map_err(err)? {
        crate::cancel::check()?;
        visit(
            row.get(0).map_err(err)?,
            row.get(1).map_err(err)?,
            crate::range::decode_cell(row)?.1,
        )?;
    }
    Ok(())
}
fn facts_issues(facts: &SourceFacts) -> Vec<Value> {
    let mut reasons = facts.issues.clone();
    if format_matches!(facts.schema, SCHEMA_USED_RANGE_FACTS_V1) {
        if facts.unstyled_blanks.count > 0
            || facts.explicit_default_blanks.count > 0
            || facts.empty_text.count > 0
        {
            reasons.push(json!({"code":"blank-source-policy","message":"Legacy source facts cannot authorize imported blank normalization; reimport the workbook."}));
        }
        if facts.row_styles.count > 0 || facts.column_styles.count > 0 {
            reasons.push(json!({"code":"axis-format-policy","message":"Legacy source facts lack qualified axis declarations; reimport the workbook."}));
        }
        let mut observed = facts.content.bounds;
        for b in [facts.styled_blanks.bounds, facts.merges.bounds]
            .into_iter()
            .flatten()
        {
            add(&mut observed, b);
        }
        if let (Some(declared), Some(observed)) = (facts.dimension_bounds, observed)
            && declared != observed
        {
            reasons.push(json!({"code":"dimension-policy","message":"Legacy declaration differs from its complete source facts; reimport the workbook."}));
        }
    }
    if facts.dimension_count > 1 || facts.dimension_count == 1 && facts.dimension_bounds.is_none() {
        reasons.push(json!({"code":"dimension-policy","message":"Duplicated or invalid worksheet dimensions remain unqualified."}));
    }
    reasons.into_iter().take(16).collect()
}
#[derive(Default)]
struct Axes {
    rows: Option<Bounds>,
    columns: Option<Bounds>,
}
impl Axes {
    fn project(&self, base: Option<Bounds>) -> Option<Bounds> {
        let row_span = base
            .map(|b| (b.min_row, b.max_row))
            .or(self.rows.map(|b| (b.min_row, b.max_row)))
            .unwrap_or((0, 0));
        let col_span = base
            .map(|b| (b.min_column, b.max_column))
            .or(self.columns.map(|b| (b.min_column, b.max_column)))
            .unwrap_or((0, 0));
        let mut result = None;
        if let Some(rows) = self.rows {
            add(
                &mut result,
                Bounds {
                    min_row: rows.min_row,
                    max_row: rows.max_row,
                    min_column: col_span.0,
                    max_column: col_span.1,
                },
            );
        }
        if let Some(cols) = self.columns {
            add(
                &mut result,
                Bounds {
                    min_row: row_span.0,
                    max_row: row_span.1,
                    min_column: cols.min_column,
                    max_column: cols.max_column,
                },
            );
        }
        result
    }
}
pub(crate) struct Normalization {
    pub bounds: Bounds,
    sheet: String,
    styles: source_styles::Inventory,
    modern: bool,
    live: bool,
    exporting: bool,
    output_styles: Option<source_styles::Inventory>,
}
impl Normalization {
    fn shielded(&self, db: &Connection, row: u32, col: u32) -> Result<bool> {
        if !self.modern {
            return Ok(false);
        }
        let mut statement=db.prepare_cached("SELECT style FROM used_range_axes WHERE sheet=?1 AND ((axis='row' AND first=?2) OR (axis='col' AND first<=?3 AND last>=?3))").map_err(err)?;
        let mut rows = statement
            .query(params![self.sheet, row, col])
            .map_err(err)?;
        while let Some(record) = rows.next().map_err(err)? {
            if !self
                .axis_inventory()
                .is_default(record.get(0).map_err(err)?)?
            {
                return Ok(true);
            }
        }
        Ok(false)
    }
    fn axis_inventory(&self) -> &source_styles::Inventory {
        self.output_styles.as_ref().unwrap_or(&self.styles)
    }
    fn persistent(&self, db: &Connection, row: u32, col: u32, cell: &Value) -> Result<bool> {
        let source_id = cell
            .get("styleId")
            .map(|s| {
                s.as_u64().ok_or_else(|| {
                    crate::coded("layout-malformed", "Invalid selected-version style index")
                })
            })
            .transpose()?
            .unwrap_or(0);
        self.styles.is_default(source_id)?;
        // A cell with content always persists. Deciding that first keeps export
        // from running style and shield lookups for every ordinary cell; the
        // export itself separately refuses an authored style without its
        // resolved graph.
        if content(cell) {
            return Ok(true);
        }
        let custom = crate::range::STYLE_FIELDS
            .iter()
            .filter(|k| **k != "styleId")
            .any(|key| cell.get(*key).is_some_and(|v| !v.is_null()));
        let shielded = self.shielded(db, row, col)?;
        if custom
            && !self.live
            && !self.exporting
            && !content(cell)
            && !shielded
            && !has_spill_metadata(cell)
            && cell["scenarioResult"] != true
        {
            return Err(crate::coded(
                "layout-bad-input",
                "Blank authored formatting requires the actual resolved export style graph",
            ));
        }
        let id = if self.exporting {
            let style:Option<(Option<u64>,Option<u64>)>=db.prepare_cached("SELECT e.style,s.style FROM export_changes e LEFT JOIN resolved_styles s ON s.id=e.style WHERE e.sheet=?1 AND e.r=?2 AND e.c=?3").map_err(crate::err)?.query_row(params![self.sheet,row,col],|r|Ok((r.get(0)?,r.get(1)?))).optional().map_err(err)?;
            if let Some((Some(_), resolved)) = style {
                resolved.ok_or_else(|| {
                    crate::coded(
                        "layout-refused",
                        "Native-prepared authored style lacks its resolved export graph",
                    )
                })?
            } else {
                // A value/calculation-only cell may carry display decorations.
                // The writer preserves the immutable source style unless the
                // native export plan contains an actual authored style slot.
                db.prepare_cached("SELECT style FROM cells WHERE sheet=?1 AND r=?2 AND c=?3")
                    .map_err(crate::err)?
                    .query_row(params![self.sheet, row, col], |r| r.get::<_, u64>(0))
                    .optional()
                    .map_err(err)?
                    .unwrap_or(0)
            }
        } else {
            source_id
        };
        let default = if self.modern {
            self.axis_inventory().is_default(id)?
        } else {
            id == 0
        };
        Ok(content(cell)
            || !default
            || self.live && custom
            || has_spill_metadata(cell)
            || cell["scenarioResult"] == true
            || shielded)
    }
    pub(crate) fn removable_cells(
        &self,
        db: &Connection,
        version: u64,
        mut visit: impl FnMut(u32, u32) -> Result<()>,
    ) -> Result<()> {
        visit_cells(db, version, &self.sheet, |r, c, cell| {
            if !self.persistent(db, r, c, &cell)? {
                visit(r, c)?;
            }
            Ok(())
        })
    }
    pub(crate) fn removable_row(&self, db: &Connection, row: u32) -> Result<bool> {
        if !self.modern {
            return Ok(false);
        }
        db.prepare_cached(
            "SELECT removable FROM used_range_axes WHERE sheet=?1 AND axis='row' AND first=?2",
        )
        .map_err(crate::err)?
        .query_row(params![self.sheet, row], |r| r.get(0))
        .map_err(err)
    }
    fn axes(&self, db: &Connection, physical_rows: bool) -> Result<Axes> {
        let mut result = Axes::default();
        if !self.modern {
            return Ok(result);
        }
        let mut statement=db.prepare_cached("SELECT axis,first,last,style,extent FROM used_range_axes WHERE sheet=?1 ORDER BY axis,first,last").map_err(err)?;
        let mut rows = statement.query([&self.sheet]).map_err(err)?;
        while let Some(record) = rows.next().map_err(err)? {
            crate::cancel::check()?;
            let axis: String = record.get(0).map_err(err)?;
            let first: u32 = record.get(1).map_err(err)?;
            let last: u32 = record.get(2).map_err(err)?;
            let style: u64 = record.get(3).map_err(err)?;
            let default = self.axis_inventory().is_default(style)?;
            let styled = !default;
            if axis == "row" && (physical_rows || record.get::<_, bool>(4).map_err(err)? || styled)
            {
                add(
                    &mut result.rows,
                    Bounds {
                        min_row: first,
                        max_row: last,
                        min_column: 0,
                        max_column: 0,
                    },
                );
            }
            if axis == "col" && styled {
                add(
                    &mut result.columns,
                    Bounds {
                        min_row: 0,
                        max_row: 0,
                        min_column: first,
                        max_column: last,
                    },
                );
            }
        }
        Ok(result)
    }
    fn effective(&self, db: &Connection, version: u64) -> Result<Option<Bounds>> {
        let mut bounds = None;
        visit_cells(db, version, &self.sheet, |r, c, cell| {
            if self.persistent(db, r, c, &cell)? {
                add(&mut bounds, Bounds::point(r, c));
            }
            Ok(())
        })?;
        Ok(bounds)
    }
}
fn accepted_schema(schema: &str) -> bool {
    format_matches!(schema, SCHEMA_USED_RANGE_FACTS_V2)
        || format_matches!(schema, SCHEMA_USED_RANGE_FACTS_V1)
}
fn normalization(
    root: &Path,
    db: &Connection,
    version: u64,
    meta: &Value,
    facts: &SourceFacts,
    live: bool,
    export_styles: Option<Option<&str>>,
) -> Result<Normalization> {
    let styles = source_styles::Inventory::load(root)?;
    let sheet = field(meta, "path")?.to_owned();
    validate_cell_styles(db, version, &sheet, styles.count)?;
    let output_styles = export_styles
        .flatten()
        .map(|xml| source_styles::Inventory::from_xml(xml.as_bytes()))
        .transpose()?;
    let mut result = Normalization {
        bounds: Bounds::point(0, 0),
        sheet,
        styles,
        modern: format_matches!(facts.schema, SCHEMA_USED_RANGE_FACTS_V2),
        live,
        exporting: export_styles.is_some(),
        output_styles,
    };
    if result.modern {
        let present: bool = db
            .query_row(
                "SELECT EXISTS(SELECT 1 FROM sqlite_master WHERE name='used_range_axes')",
                [],
                |r| r.get(0),
            )
            .map_err(err)?;
        if !present {
            return Err(crate::coded(
                "layout-unsupported",
                "Imported v2 axis facts are unavailable",
            ));
        }
    }
    let mut bounds = result.effective(db, version)?;
    if let Some(merge) = facts.merges.bounds {
        add(&mut bounds, merge);
    }
    if let Some(axis) = result.axes(db, false)?.project(bounds) {
        add(&mut bounds, axis);
    }
    result.bounds = bounds.unwrap_or(Bounds::point(0, 0));
    Ok(result)
}
/// A qualified export plan uses persistent literals/formulas/formatting and
/// finite axis projections. Physical source tags and accepted peaks are live
/// floors only; legacy v1 profiles retain their original conservative gate.
pub(crate) fn saved_normalization(
    root: &Path,
    db: &Connection,
    version: u64,
    meta: &Value,
) -> Result<Option<Normalization>> {
    qualified_normalization(root, db, version, meta, None)
}
pub(crate) fn saved_export_normalization(
    root: &Path,
    db: &Connection,
    version: u64,
    meta: &Value,
    styles: Option<&str>,
) -> Result<Option<Normalization>> {
    qualified_normalization(root, db, version, meta, Some(styles))
}
fn qualified_normalization(
    root: &Path,
    db: &Connection,
    version: u64,
    meta: &Value,
    styles: Option<Option<&str>>,
) -> Result<Option<Normalization>> {
    let Ok(mut facts) = serde_json::from_value::<SourceFacts>(meta["usedRangeFacts"].clone())
    else {
        return Ok(None);
    };
    if !format_matches!(facts.schema, SCHEMA_USED_RANGE_FACTS_V2)
        || facts
            .resolve_shared_strings(db, field(meta, "path")?)
            .is_err()
    {
        crate::cancel::check()?;
        return Ok(None);
    }
    if !facts_issues(&facts).is_empty() || meta["extended"] == true {
        return Ok(None);
    }
    match normalization(root, db, version, meta, &facts, false, styles) {
        Ok(plan) => Ok(Some(plan)),
        Err(_) => {
            crate::cancel::check()?;
            Ok(None)
        }
    }
}
#[cfg(test)]
pub(crate) fn saved_bounds(
    root: &Path,
    db: &Connection,
    version: u64,
    meta: &Value,
) -> Result<Option<Bounds>> {
    Ok(saved_normalization(root, db, version, meta)?.map(|plan| plan.bounds))
}
pub fn command(root: &Path, input: &Value) -> Result<Value> {
    if input
        .as_object()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid UsedRange request"))?
        .keys()
        .any(|k| {
            ![
                "action",
                "operation",
                "root",
                "version",
                "sheet",
                "sourceRevision",
                "diagnostics",
            ]
            .contains(&k.as_str())
        })
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid UsedRange request fields",
        ));
    }
    let version = input["version"]
        .as_u64()
        .filter(|v| *v <= 9_007_199_254_740_991)
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid UsedRange version"))?;
    let sheet_path = field(input, "sheet")?;
    let index: Value = serde_json::from_reader(BufReader::new(
        File::open(crate::local(root, "index.json")?).map_err(err)?,
    ))
    .map_err(err)?;
    let source_revision = field(&index, "revision")?;
    if source_revision.len() != 64
        || !source_revision
            .bytes()
            .all(|b| b.is_ascii_digit() || (b'a'..=b'f').contains(&b))
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid UsedRange source revision",
        ));
    }
    if field(input, "sourceRevision")? != source_revision {
        return Err(crate::coded(
            "layout-stale",
            "UsedRange source revision is stale",
        ));
    }
    let meta = index["sheets"]
        .as_array()
        .ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing UsedRange source worksheets",
            )
        })?
        .iter()
        .find(|s| s["path"] == sheet_path)
        .ok_or_else(|| crate::coded("internal-unknown-reference", "Unknown UsedRange worksheet"))?;
    let sheet = logical_sheet(root, sheet_path)?;
    let db = Connection::open_with_flags(
        crate::local(root, "cells.sqlite")?,
        OpenFlags::SQLITE_OPEN_READ_ONLY | OpenFlags::SQLITE_OPEN_NO_MUTEX,
    )
    .map_err(err)?;
    db.execute_batch("PRAGMA cache_size=-4096;PRAGMA temp_store=FILE;PRAGMA mmap_size=0;BEGIN;")
        .map_err(err)?;
    db.progress_handler(10000, Some(crate::cancel::interrupted));
    if db
        .query_row(
            "SELECT count(*) FROM versions WHERE id=?1",
            [version],
            |r| r.get::<_, u32>(0),
        )
        .map_err(err)?
        != 1
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown UsedRange version",
        ));
    }
    let mut facts: Option<SourceFacts> = serde_json::from_value(meta["usedRangeFacts"].clone())
        .ok()
        .filter(|f: &SourceFacts| {
            accepted_schema(&f.schema) && index["sourceFactSchema"] == f.schema
        });
    let mut reasons = guard_reasons(&db, version, sheet_path)?;
    if let Some(facts) = facts.as_mut()
        && facts.resolve_shared_strings(&db, sheet_path).is_err()
    {
        crate::cancel::check()?;
        facts.issue("invalid-source-facts","Complete shared-string source facts could not be resolved against the imported cell index.");
    }
    let default = SourceFacts::new();
    let facts = match facts.as_ref() {
        Some(f) => f,
        None => {
            reasons.push(json!({"code":"source-facts-unavailable","message":"This imported index lacks complete UsedRange source facts; reimport the workbook."}));
            &default
        }
    };
    reasons.extend(facts_issues(facts));
    let plan = match normalization(root, &db, version, meta, facts, true, None) {
        Ok(plan) => Some(plan),
        Err(_) => {
            crate::cancel::check()?;
            reasons.push(json!({"code":"invalid-cell-style","message":"Source cells, axis declarations or style graphs lack a qualified canonical extent policy."}));
            None
        }
    };
    if format_matches!(facts.schema, SCHEMA_USED_RANGE_FACTS_V1) {
        let bare_rows:bool=db.prepare_cached("SELECT EXISTS(SELECT 1 FROM rows r WHERE r.sheet=?1 AND NOT EXISTS(SELECT 1 FROM cells c WHERE c.sheet=r.sheet AND c.r=r.r))").unwrap().query_row([sheet_path],|r|r.get(0)).map_err(err)?;
        if bare_rows || meta["rowHeights"].as_array().is_some_and(|v| !v.is_empty()) {
            reasons.push(json!({"code":"legacy-row-facts","message":"Legacy facts lack complete physical row and height policy; reimport the workbook."}));
        }
    }
    if meta["extended"] == true {
        reasons.push(json!({"code":"extended-source","message":"Extended worksheet markup requires qualified complete extent interpretation."}));
    }
    let effective = if let Some(plan) = &plan {
        match plan.effective(&db, version) {
            Ok(b) => b,
            Err(_) => {
                crate::cancel::check()?;
                reasons.push(json!({"code":"invalid-source","message":"A complete selected-version cell scan could not validate source content."}));
                None
            }
        }
    } else {
        None
    };
    let mut accepted = history(&db, version, sheet_path)?;
    let source_floor = if format_matches!(facts.schema, SCHEMA_USED_RANGE_FACTS_V2) {
        facts.cell_tags.bounds
    } else {
        let mut floor = facts.content.bounds;
        if let Some(b) = facts.styled_blanks.bounds {
            add(&mut floor, b);
        }
        floor
    };
    if let Some(b) = source_floor {
        add(&mut accepted, b);
    }
    let mut scope = effective;
    for b in [facts.merges.bounds, accepted].into_iter().flatten() {
        add(&mut scope, b);
    }
    let mut axis = None;
    if let Some(plan) = &plan {
        if let Some(projected) = plan.axes(&db, true)?.project(scope) {
            add(&mut accepted, projected);
            add(&mut scope, projected);
        }
        axis = plan.axes(&db, false)?.project(scope);
    }
    let empty = scope.is_none();
    let b = scope.unwrap_or(Bounds::point(0, 0));
    let complete = reasons.is_empty();
    let reply = json!({"schema":workbook_format!(SCHEMA_USED_RANGE_V1),"policy":"xlsx-used-range-v1","status":if complete{"complete"}else{"incomplete"},"version":version,"sourceRevision":source_revision,"sheet":sheet,"sheetPath":sheet_path,"scopeMinMax":if complete{Some(b)}else{None},"range":if complete{Some(b.range())}else{None},"lastCell":if complete{Some(crate::xml::address(b.max_row,b.max_column))}else{None},"empty":if complete{Some(empty)}else{None},"provenance":{"factsSchema":facts.schema,"declaration":facts.declaration(),"sourceCells":facts.cell_tags.bounds,"sourceMerges":facts.merges.bounds,"effectiveCells":effective,"sourceAxisFormatting":axis,"acceptedHistory":accepted,"sourceDefaultFormatting":false},"reasons":reasons.into_iter().take(16).collect::<Vec<_>>()});
    if serde_json::to_vec(&reply).map_err(err)?.len() > 16384 {
        return Err(crate::coded(
            "layout-limit",
            "UsedRange receipt exceeds its bound",
        ));
    }
    Ok(reply)
}

/// Read-only admission for the unchanged-source Copy fast path. This is a
/// native geometry/export decision, never a caller assertion about XML bounds.
pub fn export_needed(root: &Path, input: &Value) -> Result<Value> {
    let map = input.as_object().ok_or_else(|| {
        crate::coded(
            "internal-invalid-field",
            "Invalid UsedRange export admission",
        )
    })?;
    if map.keys().any(|k| {
        ![
            "action",
            "operation",
            "root",
            "version",
            "sourceRevision",
            "diagnostics",
        ]
        .contains(&k.as_str())
    }) || input.get("diagnostics").is_some_and(|v| !v.is_boolean())
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid UsedRange export admission fields",
        ));
    }
    let version = input["version"]
        .as_u64()
        .filter(|v| *v <= 9_007_199_254_740_991)
        .ok_or_else(|| {
            crate::coded("internal-invalid-field", "Invalid UsedRange export version")
        })?;
    let index: Value = serde_json::from_reader(BufReader::new(
        File::open(crate::local(root, "index.json")?).map_err(err)?,
    ))
    .map_err(err)?;
    let revision = field(&index, "revision")?;
    if revision.len() != 64
        || !revision
            .bytes()
            .all(|b| b.is_ascii_digit() || (b'a'..=b'f').contains(&b))
        || field(input, "sourceRevision")? != revision
        || crate::hash(&crate::local(root, "source.xlsx")?)? != revision
    {
        return Err(crate::coded(
            "layout-stale",
            "UsedRange export source revision is stale",
        ));
    }
    let db = Connection::open_with_flags(
        crate::local(root, "cells.sqlite")?,
        OpenFlags::SQLITE_OPEN_READ_ONLY | OpenFlags::SQLITE_OPEN_NO_MUTEX,
    )
    .map_err(err)?;
    db.execute_batch("PRAGMA cache_size=-4096;PRAGMA temp_store=FILE;PRAGMA mmap_size=0;BEGIN;")
        .map_err(err)?;
    db.progress_handler(10000, Some(crate::cancel::interrupted));
    if db
        .query_row(
            "SELECT count(*) FROM versions WHERE id=?1",
            [version],
            |r| r.get::<_, u32>(0),
        )
        .map_err(err)?
        != 1
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown UsedRange export version",
        ));
    }
    let mut required = false;
    if format_matches!(index["sourceFactSchema"], SCHEMA_USED_RANGE_FACTS_V2) {
        for meta in index["sheets"]
            .as_array()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing source worksheets"))?
        {
            crate::cancel::check()?;
            let Some(plan) = saved_normalization(root, &db, version, meta)? else {
                continue;
            };
            if meta["dimension"]["attributes"]["ref"].as_str() != Some(plan.bounds.range().as_str())
            {
                required = true;
                break;
            }
            plan.removable_cells(&db, version, |_, _| {
                required = true;
                Ok(())
            })?;
            if required {
                break;
            }
            let mut rows=db.prepare_cached("WITH RECURSIVE a(id) AS(SELECT ?1 UNION ALL SELECT v.parent FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL) SELECT r.r FROM rows r WHERE r.sheet=?2 AND NOT EXISTS(SELECT 1 FROM cells c WHERE c.sheet=r.sheet AND c.r=r.r) AND NOT EXISTS(SELECT 1 FROM changes c JOIN a ON c.version=a.id WHERE c.sheet=r.sheet AND c.r=r.r)").map_err(err)?;
            let mut rows = rows
                .query(params![version, field(meta, "path")?])
                .map_err(err)?;
            while let Some(row) = rows.next().map_err(err)? {
                if plan.removable_row(&db, row.get(0).map_err(err)?)? {
                    required = true;
                    break;
                }
            }
            if required {
                break;
            }
        }
    }
    Ok(
        json!({"schema":workbook_format!(SCHEMA_USED_RANGE_EXPORT_NEEDED_V1),"sourceRevision":revision,"version":version,"required":required}),
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;
    fn workbook(xml: &str) -> tempfile::TempDir {
        workbook_with_strings(xml, None)
    }
    fn workbook_with_strings(xml: &str, shared_strings: Option<&str>) -> tempfile::TempDir {
        workbook_parts(xml, shared_strings, None)
    }
    fn workbook_parts(
        xml: &str,
        shared_strings: Option<&str>,
        styles: Option<&str>,
    ) -> tempfile::TempDir {
        let root = tempfile::tempdir().unwrap();
        let mut zip = zip::ZipWriter::new(File::create(root.path().join("source.xlsx")).unwrap());
        let opts = zip::write::SimpleFileOptions::default();
        for (name, text) in [
            (
                "[Content_Types].xml",
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/worksheets/cases.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#,
            ),
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Cases" sheetId="1" r:id="sheet1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="sheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/cases.xml"/></Relationships>"#,
            ),
            (
                "xl/styles.xml",
                r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><name val="Calibri"/><sz val="11"/></font><font><name val="Calibri"/><sz val="11"/><b/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>"#,
            ),
            ("xl/worksheets/cases.xml", xml),
        ] {
            zip.start_file(name, opts).unwrap();
            zip.write_all(
                if name == "xl/styles.xml" {
                    styles.unwrap_or(text)
                } else {
                    text
                }
                .as_bytes(),
            )
            .unwrap();
        }
        if let Some(strings) = shared_strings {
            zip.start_file("xl/sharedStrings.xml", opts).unwrap();
            zip.write_all(strings.as_bytes()).unwrap();
        }
        zip.finish().unwrap();
        crate::import::index(root.path()).unwrap();
        root
    }
    fn query(root: &Path, version: u64) -> Value {
        let index: Value =
            serde_json::from_reader(File::open(root.join("index.json")).unwrap()).unwrap();
        command(root,&json!({"action":"workbook","operation":"used-range","root":root,"version":version,"sourceRevision":index["revision"],"sheet":"xl/worksheets/cases.xml"})).unwrap()
    }
    fn basic() -> tempfile::TempDir {
        workbook(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C6"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="6"><c r="C6"><v>6</v></c></row></sheetData></worksheet>"#,
        )
    }
    #[test]
    fn source_capture_keeps_duplicate_declarations_and_blank_categories_bounded() {
        let mut f = SourceFacts::new();
        f.dimension(Some("A1"));
        f.dimension(Some("Z80"));
        f.cell(0, 0, "n", Some("1"), "", false, 0, false);
        f.cell(79, 25, "n", None, "", false, 1, true);
        f.cell(4, 3, "n", None, "", false, 0, true);
        f.cell(5, 4, "n", None, "", false, 0, false);
        assert_eq!(f.dimension_count, 2);
        assert_eq!(f.declaration()["status"], "duplicate");
        assert_eq!(f.declaration()["ref"], "A1");
        assert_eq!(f.declaration()["bounds"], Value::Null);
        assert_eq!(f.styled_blanks.bounds, Some(Bounds::point(79, 25)));
        assert_eq!(f.explicit_default_blanks.count, 1);
        assert_eq!(f.unstyled_blanks.count, 1);
        for i in 0..100_000 {
            f.cell(i % 1048576, 0, "n", Some("0"), "", false, 0, false);
        }
        assert!(
            serde_json::to_vec(&f).unwrap().len() < 4096,
            "aggregate size does not grow with source cells"
        );
    }
    fn resolved_facts(root: &Path) -> SourceFacts {
        let index: Value =
            serde_json::from_reader(File::open(root.join("index.json")).unwrap()).unwrap();
        let mut facts: SourceFacts =
            serde_json::from_value(index["sheets"][0]["usedRangeFacts"].clone()).unwrap();
        let db = Connection::open_with_flags(
            root.join("cells.sqlite"),
            OpenFlags::SQLITE_OPEN_READ_ONLY,
        )
        .unwrap();
        facts
            .resolve_shared_strings(&db, "xl/worksheets/cases.xml")
            .unwrap();
        facts
    }
    #[test]
    fn empty_shared_literals_keep_raw_categories_and_qualified_persistent_extent() {
        let strings = r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t/></si></sst>"#;
        for style in ["", " s=\"0\"", " s=\"1\""] {
            let xml = format!(
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="Z80"/><sheetData><row r="80"><c r="Z80" t="s"{style}><v>0</v></c></row></sheetData></worksheet>"#
            );
            let root = workbook_with_strings(&xml, Some(strings));
            let before_db = crate::hash(&root.path().join("cells.sqlite")).unwrap();
            let before_index = crate::hash(&root.path().join("index.json")).unwrap();
            let facts = resolved_facts(root.path());
            assert_eq!(facts.cell_tags.count, 1);
            assert_eq!(facts.content.count, 0);
            assert_eq!(facts.empty_text.count, 1);
            let stored: Value =
                serde_json::from_reader(File::open(root.path().join("index.json")).unwrap())
                    .unwrap();
            assert_eq!(stored["sheets"][0]["usedRangeFacts"]["content"]["count"], 0);
            assert_eq!(
                stored["sheets"][0]["usedRangeFacts"]["emptyText"]["count"],
                1
            );
            if style.is_empty() {
                assert_eq!(facts.unstyled_blanks.count, 1);
            } else if style.contains('1') {
                assert_eq!(facts.styled_blanks.bounds, Some(Bounds::point(79, 25)));
            } else {
                assert_eq!(facts.explicit_default_blanks.count, 1);
            }
            let receipt = query(root.path(), 0);
            assert_eq!(receipt["status"], "complete");
            assert_eq!(receipt["range"], "Z80");
            let index: Value =
                serde_json::from_reader(File::open(root.path().join("index.json")).unwrap())
                    .unwrap();
            let db = Connection::open_with_flags(
                root.path().join("cells.sqlite"),
                OpenFlags::SQLITE_OPEN_READ_ONLY,
            )
            .unwrap();
            assert_eq!(
                saved_bounds(root.path(), &db, 0, &index["sheets"][0]).unwrap(),
                Some(Bounds::point(79, 25))
            );
            assert_eq!(
                crate::hash(&root.path().join("cells.sqlite")).unwrap(),
                before_db
            );
            assert_eq!(
                crate::hash(&root.path().join("index.json")).unwrap(),
                before_index
            );
        }
    }
    #[test]
    fn shared_text_whitespace_and_empty_formula_caches_keep_canonical_presence() {
        let strings = r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>payload</t></si><si><t xml:space="preserve"> </t></si><si><t/></si></sst>"#;
        let root = workbook_with_strings(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:Z80"/><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row><row r="5"><c r="C5" t="s"><v>1</v></c></row><row r="6"><c r="J6" t="str"><f>""</f><v/></c></row><row r="80"><c r="Z80" t="s"><f>""</f><v>2</v></c></row></sheetData></worksheet>"#,
            Some(strings),
        );
        let facts = resolved_facts(root.path());
        assert_eq!(facts.content.count, 4);
        assert_eq!(facts.content.bounds, Some(Bounds::parse("A1:Z80").unwrap()));
        assert_eq!(facts.empty_text.count, 0);
        assert!(facts.issues.is_empty());
        let receipt = query(root.path(), 0);
        assert_eq!(receipt["status"], "complete");
        assert_eq!(receipt["range"], "A1:Z80");
        let db = Connection::open_with_flags(
            root.path().join("cells.sqlite"),
            OpenFlags::SQLITE_OPEN_READ_ONLY,
        )
        .unwrap();
        assert_eq!(
            db.prepare_cached("SELECT value FROM strings WHERE id=1")
                .unwrap()
                .query_row([], |row| row.get::<_, String>(0))
                .unwrap(),
            " "
        );
        assert_eq!(
            db.prepare_cached("SELECT value FROM strings WHERE id=2")
                .unwrap()
                .query_row([], |row| row.get::<_, String>(0))
                .unwrap(),
            ""
        );
        let index: Value =
            serde_json::from_reader(File::open(root.path().join("index.json")).unwrap()).unwrap();
        assert_eq!(
            saved_bounds(root.path(), &db, 0, &index["sheets"][0]).unwrap(),
            Some(Bounds::parse("A1:Z80").unwrap())
        );
    }
    #[test]
    fn malformed_or_dangling_shared_references_remain_incomplete_and_unmodified() {
        let strings = r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>payload</t></si></sst>"#;
        for ordinal in ["bad", "-1", "999", ""] {
            let xml = format!(
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="Z80"/><sheetData><row r="80"><c r="Z80" t="s"><v>{ordinal}</v></c></row></sheetData></worksheet>"#
            );
            let root = workbook_with_strings(&xml, Some(strings));
            let before_db = crate::hash(&root.path().join("cells.sqlite")).unwrap();
            let before_index = crate::hash(&root.path().join("index.json")).unwrap();
            let facts = resolved_facts(root.path());
            assert!(
                facts
                    .issues
                    .iter()
                    .any(|issue| issue["code"] == "invalid-shared-string")
            );
            let receipt = query(root.path(), 0);
            assert_eq!(receipt["status"], "incomplete");
            for key in ["scopeMinMax", "range", "lastCell", "empty"] {
                assert_eq!(receipt[key], Value::Null);
            }
            let stored: Value =
                serde_json::from_reader(File::open(root.path().join("index.json")).unwrap())
                    .unwrap();
            assert!(
                stored["sheets"][0]["usedRangeFacts"]["issues"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .any(|issue| issue["code"] == "invalid-shared-string")
            );
            assert!(
                receipt["reasons"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .any(|reason| reason["code"] == "invalid-shared-string")
            );
            let index: Value =
                serde_json::from_reader(File::open(root.path().join("index.json")).unwrap())
                    .unwrap();
            let db = Connection::open_with_flags(
                root.path().join("cells.sqlite"),
                OpenFlags::SQLITE_OPEN_READ_ONLY,
            )
            .unwrap();
            assert_eq!(
                saved_bounds(root.path(), &db, 0, &index["sheets"][0]).unwrap(),
                None
            );
            assert_eq!(
                crate::hash(&root.path().join("cells.sqlite")).unwrap(),
                before_db
            );
            assert_eq!(
                crate::hash(&root.path().join("index.json")).unwrap(),
                before_index
            );
        }
    }
    #[test]
    fn cached_v1_sst_ordinal_aggregates_are_resolved_without_index_or_sql_mutation() {
        let root = workbook_with_strings(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="Z80"/><sheetData><row r="80"><c r="Z80" t="s"><v>0</v></c></row></sheetData></worksheet>"#,
            Some(
                r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t/></si></sst>"#,
            ),
        );
        let mut index: Value =
            serde_json::from_reader(File::open(root.path().join("index.json")).unwrap()).unwrap();
        index["sourceFactSchema"] = json!(LEGACY_FACTS_SCHEMA);
        let facts = &mut index["sheets"][0]["usedRangeFacts"];
        facts["schema"] = json!(LEGACY_FACTS_SCHEMA);
        facts.as_object_mut().unwrap().remove("rowTags");
        facts.as_object_mut().unwrap().remove("rowEffects");
        facts["content"] = json!({"count":1,"bounds":Bounds::point(79,25)});
        facts["unstyledBlanks"] = json!({"count":0,"bounds":null});
        facts["emptyText"] = json!({"count":0,"bounds":null});
        std::fs::write(
            root.path().join("index.json"),
            serde_json::to_vec(&index).unwrap(),
        )
        .unwrap();
        let before_index = crate::hash(&root.path().join("index.json")).unwrap();
        let before_db = crate::hash(&root.path().join("cells.sqlite")).unwrap();
        let receipt = query(root.path(), 0);
        assert_eq!(receipt["status"], "incomplete");
        assert!(
            receipt["reasons"]
                .as_array()
                .unwrap()
                .iter()
                .any(|reason| reason["code"] == "blank-source-policy")
        );
        let db = Connection::open_with_flags(
            root.path().join("cells.sqlite"),
            OpenFlags::SQLITE_OPEN_READ_ONLY,
        )
        .unwrap();
        assert_eq!(
            saved_bounds(root.path(), &db, 0, &index["sheets"][0]).unwrap(),
            None
        );
        assert_eq!(
            crate::hash(&root.path().join("index.json")).unwrap(),
            before_index
        );
        assert_eq!(
            crate::hash(&root.path().join("cells.sqlite")).unwrap(),
            before_db
        );
    }
    #[test]
    fn source_axis_styles_and_merge_bounds_stay_distinct_from_geometry_defaults() {
        let mut f = SourceFacts::new();
        let attrs = |pairs: Vec<(&'static str, &'static str)>| {
            pairs
                .into_iter()
                .map(|(k, v)| (k, Cow::Borrowed(v)))
                .collect()
        };
        f.element(
            "sheetFormatPr",
            &attrs(vec![("defaultRowHeight", "24")]),
            None,
        );
        f.element(
            "col",
            &attrs(vec![("min", "7"), ("max", "7"), ("width", "18")]),
            None,
        );
        assert!(f.geometry_defaults);
        assert_eq!(f.column_styles.count, 0);
        f.element(
            "col",
            &attrs(vec![("min", "7"), ("max", "7"), ("style", "1")]),
            None,
        );
        f.element(
            "row",
            &attrs(vec![("s", "1"), ("customFormat", "1")]),
            Some(99),
        );
        f.element("mergeCell", &attrs(vec![("ref", "D4:G6")]), None);
        assert_eq!(f.column_styles.bounds.unwrap().max_row, 1048575);
        assert_eq!(f.row_styles.bounds.unwrap().max_column, 16383);
        assert_eq!(f.merges.bounds, Some(Bounds::parse("D4:G6").unwrap()));
        assert!(
            facts_issues(&f).is_empty(),
            "v2 raw axis coverage is distinct from qualified projected geometry"
        );
    }
    #[test]
    fn pure_native_query_uses_source_bounds_and_preserves_database_and_index_bytes() {
        let root = basic();
        let db_before = crate::hash(&root.path().join("cells.sqlite")).unwrap();
        let index_before = crate::hash(&root.path().join("index.json")).unwrap();
        let reply = query(root.path(), 0);
        assert_eq!(reply["status"], "complete");
        assert_eq!(reply["range"], "A1:C6");
        assert_eq!(reply["lastCell"], "C6");
        assert_eq!(reply["sheet"], "Cases");
        assert_eq!(reply["provenance"]["sourceAxisFormatting"], Value::Null);
        assert_eq!(
            crate::hash(&root.path().join("cells.sqlite")).unwrap(),
            db_before
        );
        assert_eq!(
            crate::hash(&root.path().join("index.json")).unwrap(),
            index_before
        );
        assert!(
            command(
                root.path(),
                &json!({"version":1,"sheet":"xl/worksheets/cases.xml"})
            )
            .is_err()
        );
        assert!(command(root.path(), &json!({"version":0,"sheet":"unknown"})).is_err());
        assert!(command(root.path(), &json!({"version":0,"sheet":"xl/worksheets/cases.xml","sourceRevision":"0".repeat(64)})).unwrap_err().contains("source revision is stale"));
    }
    #[test]
    fn accepted_far_cells_clear_floor_undo_sibling_and_compaction_use_selected_ancestry() {
        let root = basic();
        let sheet = "xl/worksheets/cases.xml";
        let patch = |version, items: Vec<(&str, Value)>| {
            crate::range::patch(root.path(),&json!({"version":version,"changes":items.into_iter().map(|(cell,value)|json!({"sheet":sheet,"cell":cell,"value":{"value":value,"dataType":"number","styleId":0},"edits":[{"value":value}]})).collect::<Vec<_>>()})).unwrap()["version"].as_u64().unwrap()
        };
        let far = patch(0, vec![("A100", json!(9)), ("Z1", json!(11))]);
        assert_eq!(query(root.path(), far)["lastCell"], "Z100");
        assert_eq!(query(root.path(), 0)["lastCell"], "C6");
        let clear=crate::bulk::apply(root.path(),&json!({"version":far,"commands":[{"type":"clearRange","sheet":sheet,"range":"Z1"},{"type":"clearRange","sheet":sheet,"range":"A100"}]})).unwrap()["version"].as_u64().unwrap();
        assert_eq!(query(root.path(), clear)["range"], "A1:Z100");
        assert_eq!(
            query(root.path(), clear)["provenance"]["effectiveCells"],
            serde_json::to_value(Bounds::parse("A1:C6").unwrap()).unwrap()
        );
        let sibling = patch(0, vec![("D7", json!(7))]);
        assert_eq!(query(root.path(), sibling)["range"], "A1:D7");
        crate::compact::compact(root.path(), &json!({"keep":[clear,sibling],"vacuum":false}))
            .unwrap();
        assert_eq!(query(root.path(), clear)["lastCell"], "Z100");
        assert_eq!(query(root.path(), sibling)["lastCell"], "D7");
        assert_eq!(query(root.path(), 0)["lastCell"], "C6");
        let noop=crate::bulk::apply(root.path(),&json!({"version":0,"commands":[{"type":"clearRange","sheet":sheet,"range":"XFD1048576"}]})).unwrap()["version"].as_u64().unwrap();
        assert_eq!(query(root.path(), noop)["range"], "A1:C6");
    }
    #[test]
    fn ambiguous_imported_extent_has_facts_but_no_authorized_geometry() {
        let root = workbook(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><dimension ref="Z80"/><sheetData><row r="80"><c r="Z80"/></row></sheetData></worksheet>"#,
        );
        let reply = query(root.path(), 0);
        assert_eq!(reply["status"], "incomplete");
        for key in ["scopeMinMax", "range", "lastCell", "empty"] {
            assert_eq!(reply[key], Value::Null);
        }
        assert_eq!(reply["provenance"]["declaration"]["status"], "duplicate");
        assert_eq!(
            reply["provenance"]["sourceCells"],
            serde_json::to_value(Bounds::point(79, 25)).unwrap()
        );
        assert!(!reply["reasons"].as_array().unwrap().is_empty());
    }
    fn export_saved(root: &Path, version: u64) -> String {
        let view = json!({"sheets":[{"name":"Cases","path":"xl/worksheets/cases.xml"}],"namedRanges":[],"tables":[]});
        let plan =
            crate::export_state::prepare(root, &json!({"version":version,"view":view})).unwrap();
        let db = crate::database(root).unwrap();
        let mut statement = db
            .prepare_cached("SELECT id,signature FROM export_styles ORDER BY id")
            .unwrap();
        let style_ids: Vec<Value> = statement
            .query_map([], |r| Ok((r.get::<_, u64>(0)?, r.get::<_, String>(1)?)))
            .unwrap()
            .map(|item| {
                let (id, signature) = item.unwrap();
                let value: Value = serde_json::from_str(&signature).unwrap();
                assert!(
                    value
                        .as_object()
                        .unwrap()
                        .keys()
                        .all(|key| key == "styleId")
                );
                json!({"id":id,"styleId":value["styleId"]})
            })
            .collect();
        drop(statement);
        drop(db);
        let mut source =
            zip::ZipArchive::new(File::open(root.join("source.xlsx")).unwrap()).unwrap();
        let mut workbook = String::new();
        source
            .by_name("xl/workbook.xml")
            .unwrap()
            .read_to_string(&mut workbook)
            .unwrap();
        let refresh = if plan["incomplete"] == true { "1" } else { "0" };
        let calculation = format!(
            "<calcPr calcMode=\"auto\" calcOnSave=\"1\" fullCalcOnLoad=\"{refresh}\" forceFullCalc=\"{refresh}\"/>"
        );
        // Re-exporting this helper's own saved fixtures replaces their known
        // generated calcPr, rather than manufacturing a duplicate element.
        let replacement = if let Some(start) = workbook.find("<calcPr ") {
            let end = start + workbook[start..].find("/>").unwrap() + 2;
            let mut replacement = workbook.clone();
            replacement.replace_range(start..end, &calculation);
            replacement
        } else {
            workbook.replace("</workbook>", &format!("{calculation}</workbook>"))
        };
        std::fs::write(
            root.join("export-metadata.json"),
            json!({"styleIds":style_ids,"replacements":{"xl/workbook.xml":replacement}})
                .to_string(),
        )
        .unwrap();
        let index: Value =
            serde_json::from_reader(File::open(root.join("index.json")).unwrap()).unwrap();
        crate::export::write(
            root,
            &json!({"revision":index["revision"],"version":version,"calculationMetadata":true}),
        )
        .unwrap();
        let mut saved =
            zip::ZipArchive::new(File::open(root.join("export.xlsx")).unwrap()).unwrap();
        let mut xml = String::new();
        saved
            .by_name("xl/worksheets/cases.xml")
            .unwrap()
            .read_to_string(&mut xml)
            .unwrap();
        xml
    }
    #[test]
    fn imported_cleared_content_and_nondefault_blank_styles_normalize_only_after_save() {
        let root = workbook(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:Z1000"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="6"><c r="C6"><v>6</v></c></row><row r="1000"><c r="Z1000" s="1"/></row></sheetData></worksheet>"#,
        );
        let sheet = "xl/worksheets/cases.xml";
        assert_eq!(query(root.path(), 0)["lastCell"], "Z1000");
        let contents = crate::bulk::apply(
            root.path(),
            &json!({"version":0,"commands":[{"type":"clearRange","sheet":sheet,"range":"Z1000"}]}),
        )
        .unwrap()["version"]
            .as_u64()
            .unwrap();
        let retained = export_saved(root.path(), contents);
        assert!(retained.contains("r=\"Z1000\"") && retained.contains("s=\"1\""));
        assert_eq!(query(root.path(), contents)["lastCell"], "Z1000");
        let formats=crate::format::apply(root.path(),&json!({"version":contents,"styles":2,"commands":[{"sheet":sheet,"range":"Z1000","format":{"styleId":0}}]})).unwrap()["version"].as_u64().unwrap();
        assert_eq!(query(root.path(), formats)["lastCell"], "Z1000");
        let cleared = export_saved(root.path(), formats);
        assert!(!cleared.contains("r=\"Z1000\""));
        assert!(cleared.contains("dimension ref=\"A1:C6\""));
        let cold = tempfile::tempdir().unwrap();
        std::fs::copy(
            root.path().join("export.xlsx"),
            cold.path().join("source.xlsx"),
        )
        .unwrap();
        crate::import::index(cold.path()).unwrap();
        assert_eq!(query(cold.path(), 0)["lastCell"], "C6");
    }
    fn saved_dimension(xml: &str) -> String {
        let mut reader = NsReader::from_reader(xml.as_bytes());
        let mut buffer = Vec::new();
        let mut refs = Vec::new();
        loop {
            let (ns, event) = reader.read_resolved_event_into(&mut buffer).unwrap();
            match event {
                Event::Start(e) | Event::Empty(e)
                    if ns == ResolveResult::Bound(quick_xml::name::Namespace(crate::xml::NS))
                        && e.local_name().as_ref() == b"dimension" =>
                {
                    refs.push(crate::xml::attrs(&e).unwrap()["ref"].clone())
                }
                Event::Eof => break,
                _ => {}
            }
            buffer.clear();
        }
        assert_eq!(refs.len(), 1);
        refs.pop().unwrap()
    }
    fn verify_save_and_cold(root: &Path, live: &str, saved: &str) -> String {
        assert_eq!(query(root, 0)["status"], "complete");
        assert_eq!(query(root, 0)["range"], live);
        let index_hash = crate::hash(&root.join("index.json")).unwrap();
        let xml = export_saved(root, 0);
        assert_eq!(saved_dimension(&xml), saved);
        assert_eq!(
            query(root, 0)["range"],
            live,
            "Copy/export cannot reset the accepted live source floor"
        );
        assert_eq!(crate::hash(&root.join("index.json")).unwrap(), index_hash);
        let cold = tempfile::tempdir().unwrap();
        std::fs::copy(root.join("export.xlsx"), cold.path().join("source.xlsx")).unwrap();
        crate::import::index(cold.path()).unwrap();
        assert_eq!(query(cold.path(), 0)["status"], "complete");
        assert_eq!(query(cold.path(), 0)["range"], saved);
        xml
    }
    const BASE_CELLS: &str =
        "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row><row r=\"6\"><c r=\"C6\"><v>6</v></c></row>";
    #[test]
    fn physical_default_blank_tags_and_bare_rows_drop_only_from_saved_geometry() {
        let styles = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts><font/></fonts><fills><fill/></fills><borders><border/></borders><cellStyleXfs><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0"/><xf xfId="0" borderId="0" fillId="0" fontId="0" numFmtId="0"></xf></cellXfs></styleSheet>"#;
        for attrs in ["", " s=\"0\"", " s=\"2\""] {
            let root = workbook_parts(
                &format!(
                    "<worksheet xmlns=\"{}\"><dimension ref=\"A1:C6\"/><sheetData>{BASE_CELLS}<row r=\"80\"><c r=\"Z80\"{attrs}/></row></sheetData></worksheet>",
                    String::from_utf8_lossy(crate::xml::NS)
                ),
                None,
                Some(styles),
            );
            let xml = verify_save_and_cold(root.path(), "A1:Z80", "A1:C6");
            assert!(!xml.contains("r=\"Z80\""));
            assert!(!xml.contains("r=\"80\""));
        }
        for (columns, cells, expected) in [
            (r#"<cols><col min="7" max="7" style="2"/></cols>"#, "", "A1"),
            (
                r#"<cols><col min="7" max="7" style="2"/></cols>"#,
                BASE_CELLS,
                "A1:C6",
            ),
            ("", r#"<row r="100" s="2" customFormat="1"/>"#, "A100"),
        ] {
            let root = workbook_parts(
                &format!(
                    "<worksheet xmlns=\"{}\">{columns}<sheetData>{cells}</sheetData></worksheet>",
                    String::from_utf8_lossy(crate::xml::NS)
                ),
                None,
                Some(styles),
            );
            verify_save_and_cold(root.path(), expected, expected);
        }
        let root = workbook(&format!(
            "<worksheet xmlns=\"{}\"><sheetData>{BASE_CELLS}<row r=\"100\"/></sheetData></worksheet>",
            String::from_utf8_lossy(crate::xml::NS)
        ));
        assert!(!verify_save_and_cold(root.path(), "A1:C100", "A1:C6").contains("r=\"100\""));
    }
    #[test]
    fn empty_literals_whitespace_and_formula_empty_caches_survive_saved_normalization() {
        for cell in [
            r#"<c r="Z80" t="inlineStr"><is><t/></is></c>"#,
            r#"<c r="Z80" t="str"><v/></c>"#,
            r#"<c r="Z80" t="s"><v>0</v></c>"#,
            r#"<c r="Z80" t="s"><v>1</v></c>"#,
            r#"<c r="Z80" t="str"><f>""</f><v/></c>"#,
        ] {
            let root = workbook_with_strings(
                &format!(
                    "<worksheet xmlns=\"{}\"><dimension ref=\"A1:C6\"/><sheetData>{BASE_CELLS}<row r=\"80\">{cell}</row></sheetData></worksheet>",
                    String::from_utf8_lossy(crate::xml::NS)
                ),
                Some(
                    r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t/></si><si><t xml:space="preserve"> </t></si></sst>"#,
                ),
            );
            let xml = verify_save_and_cold(root.path(), "A1:Z80", "A1:Z80");
            assert!(xml.contains("r=\"Z80\""));
        }
    }
    #[test]
    fn finite_axis_projection_matches_observed_empty_and_populated_controls() {
        let cases = [
            (
                r#"<cols><col min="7" max="7" style="1"/></cols>"#,
                BASE_CELLS,
                "A1:G6",
                "A1:G6",
            ),
            (
                r#"<cols><col min="7" max="7" width="18" customWidth="1"/></cols>"#,
                BASE_CELLS,
                "A1:C6",
                "A1:C6",
            ),
            (
                "",
                "<row r=\"100\" s=\"1\" customFormat=\"1\"/>",
                "A100",
                "A100",
            ),
            (
                r#"<cols><col min="7" max="7" style="1"/></cols>"#,
                "",
                "G1",
                "G1",
            ),
            (
                r#"<cols><col min="7" max="7" style="1"/></cols>"#,
                "<row r=\"100\" s=\"1\" customFormat=\"1\"/>",
                "G100",
                "G100",
            ),
            (
                "",
                "<row r=\"100\" ht=\"30\" customHeight=\"1\"/>",
                "A100",
                "A100",
            ),
            (
                "",
                "<row r=\"100\" s=\"0\" customFormat=\"1\"/>",
                "A100",
                "A100",
            ),
            ("", "<row r=\"100\"/>", "A100", "A1"),
            (
                r#"<cols><col min="7" max="7" width="18" customWidth="1"/></cols>"#,
                "",
                "A1",
                "A1",
            ),
            (
                r#"<cols><col min="7" max="7" style="0"/></cols>"#,
                "",
                "A1",
                "A1",
            ),
        ];
        for (columns, cells, live, saved) in cases {
            let root = workbook(&format!(
                "<worksheet xmlns=\"{}\"><dimension ref=\"A1\"/>{columns}<sheetData>{cells}</sheetData></worksheet>",
                String::from_utf8_lossy(crate::xml::NS)
            ));
            verify_save_and_cold(root.path(), live, saved);
            assert!(
                query(root.path(), 0)["range"].as_str().unwrap().len() < 20,
                "axis coverage is projected, never materialized as full row/column geometry"
            );
        }
        for row in [
            r#"<row r="100" s="1" customFormat="1"/>"#,
            r#"<row r="100" ht="30" customHeight="1"/>"#,
        ] {
            let root = workbook(&format!(
                "<worksheet xmlns=\"{}\"><sheetData>{BASE_CELLS}{row}</sheetData></worksheet>",
                String::from_utf8_lossy(crate::xml::NS)
            ));
            verify_save_and_cold(root.path(), "A1:C100", "A1:C100");
        }
    }
    #[test]
    fn explicit_default_cell_tags_shield_axis_styles_through_save_and_clear_formats() {
        for style in ["", " s=\"0\""] {
            for (columns, row, cell, expected) in [
                (
                    r#"<cols><col min="7" max="7" style="1"/></cols>"#,
                    r#"r="80""#,
                    "G80",
                    "G80",
                ),
                ("", r#"r="100" s="1" customFormat="1""#, "Z100", "Z100"),
            ] {
                let root = workbook(&format!(
                    "<worksheet xmlns=\"{}\">{columns}<sheetData><row {row}><c r=\"{cell}\"{style}/></row></sheetData></worksheet>",
                    String::from_utf8_lossy(crate::xml::NS)
                ));
                assert!(
                    verify_save_and_cold(root.path(), expected, expected)
                        .contains(&format!("r=\"{cell}\""))
                );
                let version=crate::format::apply(root.path(),&json!({"version":0,"styles":2,"commands":[{"sheet":"xl/worksheets/cases.xml","range":cell,"format":{"styleId":0}}]})).unwrap()["version"].as_u64().unwrap();
                let xml = export_saved(root.path(), version);
                assert_eq!(saved_dimension(&xml), expected);
                assert!(
                    xml.contains(&format!("r=\"{cell}\"")),
                    "default cell must keep shielding the source axis style"
                );
            }
        }
    }
    #[test]
    fn populated_empty_sheet_data_has_one_root_dimension_and_cold_sort_remains_qualified() {
        let sheet = "xl/worksheets/cases.xml";
        for data in ["<sheetData/>", "<sheetData></sheetData>", "<m:sheetData/>"] {
            let root = workbook(&format!(
                "<worksheet xmlns=\"{}\" xmlns:m=\"{}\">{data}</worksheet>",
                String::from_utf8_lossy(crate::xml::NS),
                String::from_utf8_lossy(crate::xml::NS)
            ));
            let changes=[("A1",1),("A2",20),("B2",9),("A3",30),("B3",2),("B4",7)].into_iter().map(|(cell,value)|json!({"sheet":sheet,"cell":cell,"value":{"value":value,"dataType":"number","rawValue":value.to_string(),"cachedValue":null},"edit":{"value":value}})).collect::<Vec<_>>();
            let seed = crate::range::patch(root.path(), &json!({"version":0,"changes":changes}))
                .unwrap()["version"]
                .as_u64()
                .unwrap();
            let filled=crate::bulk::apply(root.path(),&json!({"version":seed,"commands":[{"type":"fillRange","sheet":sheet,"source":"B4","range":"B4:B1200"}]})).unwrap()["version"].as_u64().unwrap();
            let version=crate::range::patch(root.path(),&json!({"version":filled,"changes":[{"sheet":sheet,"cell":"D1","value":{"formula":"SUM(B4:B1200)","value":null,"dataType":"blank","rawValue":null,"cachedValue":null},"edit":{"formula":"SUM(B4:B1200)"}}]})).unwrap()["version"].as_u64().unwrap();
            let xml = export_saved(root.path(), version);
            let mut reader = NsReader::from_reader(xml.as_bytes());
            let mut buffer = Vec::new();
            let (mut depth, mut dimensions, mut cells) = (0usize, 0usize, 0usize);
            loop {
                let (ns, event) = reader.read_resolved_event_into(&mut buffer).unwrap();
                let main = ns == ResolveResult::Bound(quick_xml::name::Namespace(crate::xml::NS));
                match event {
                    Event::Start(ref e) | Event::Empty(ref e) => {
                        if main && e.local_name().as_ref() == b"dimension" {
                            assert_eq!(
                                depth, 1,
                                "dimension must be a direct worksheet child: {data}"
                            );
                            assert_eq!(crate::xml::attrs(e).unwrap()["ref"], "A1:D1200");
                            dimensions += 1;
                        }
                        if main && e.local_name().as_ref() == b"sheetData" {
                            assert_eq!(depth, 1);
                        }
                        if main && e.local_name().as_ref() == b"c" {
                            cells += 1;
                        }
                        if matches!(event, Event::Start(_)) {
                            depth += 1;
                        }
                    }
                    Event::End(_) => depth -= 1,
                    Event::Eof => break,
                    _ => {}
                }
                buffer.clear();
            }
            assert_eq!(dimensions, 1);
            assert_eq!(cells, 1203);
            assert!(xml.find("dimension").unwrap() < xml.find("sheetData").unwrap());
            let cold = tempfile::tempdir().unwrap();
            std::fs::copy(
                root.path().join("export.xlsx"),
                cold.path().join("source.xlsx"),
            )
            .unwrap();
            let indexed = crate::import::index(cold.path()).unwrap();
            assert_eq!(indexed["sheets"][0]["usedRangeFacts"]["dimensionCount"], 1);
            assert_eq!(indexed["sheets"][0]["extended"], false);
            assert_eq!(indexed["sheets"][0]["storedCells"], 1203);
            assert_eq!(query(cold.path(), 0)["range"], "A1:D1200");
            let sorted=crate::transform::apply(cold.path(),&json!({"action":"workbook","operation":"transform","root":cold.path(),"source":"source.xlsx","version":0,"kind":"sortRows","sheet":sheet,"range":"A2:B3","keys":[{"column":"B","direction":"asc"}],"view":{"sheets":[{"name":"Cases","path":sheet}],"namedRanges":[],"tables":[]}})).unwrap();
            let sorted_root = tempfile::tempdir().unwrap();
            std::fs::copy(
                cold.path().join(sorted["file"].as_str().unwrap()),
                sorted_root.path().join("source.xlsx"),
            )
            .unwrap();
            let sorted_index = crate::import::index(sorted_root.path()).unwrap();
            assert_eq!(
                sorted_index["sheets"][0]["usedRangeFacts"]["dimensionCount"],
                0
            );
            // Private transforms omit dimension hints. Their raw source facts
            // remain truthful; public export must restore a canonical hint.
            assert_eq!(sorted_index["sheets"][0]["extended"], false);
            assert_eq!(sorted_index["sheets"][0]["storedCells"], 1203);
            assert_eq!(query(sorted_root.path(), 0)["range"], "A1:D1200");
            // This independent range does not disturb D1's SUM inputs.
            let db = crate::database(sorted_root.path()).unwrap();
            let rows = crate::range::read_db(
                &db,
                &json!({"version":0,"sheet":sheet,"range":"A2:B3","limit":4}),
            )
            .unwrap();
            for (reference, value) in [("A2", 30.0), ("B2", 2.0), ("A3", 20.0), ("B3", 9.0)] {
                let cell = rows["cells"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .find(|pair| pair[0] == reference)
                    .unwrap();
                assert_eq!(cell[1]["value"].as_f64(), Some(value));
            }
            let fills = crate::range::read_db(
                &db,
                &json!({"version":0,"sheet":sheet,"range":"B4:B1200","limit":2000}),
            )
            .unwrap();
            assert_eq!(fills["cells"].as_array().unwrap().len(), 1197);
            assert!(
                fills["cells"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .all(|pair| pair[1]["value"].as_f64() == Some(7.0))
            );
            let formula = crate::range::read_db(
                &db,
                &json!({"version":0,"sheet":sheet,"range":"D1","limit":1}),
            )
            .unwrap();
            assert_eq!(formula["cells"][0][1]["formula"], "SUM(B4:B1200)");
            drop(db);
            let final_xml = export_saved(sorted_root.path(), 0);
            let mut reader = NsReader::from_reader(final_xml.as_bytes());
            let mut buffer = Vec::new();
            let (mut depth, mut dimensions) = (0usize, 0usize);
            loop {
                let (ns, event) = reader.read_resolved_event_into(&mut buffer).unwrap();
                let main = ns == ResolveResult::Bound(quick_xml::name::Namespace(crate::xml::NS));
                match event {
                    Event::Start(ref e) | Event::Empty(ref e) => {
                        if main && e.local_name().as_ref() == b"dimension" {
                            assert_eq!(depth, 1);
                            assert_eq!(crate::xml::attrs(e).unwrap()["ref"], "A1:D1200");
                            dimensions += 1;
                        }
                        if matches!(event, Event::Start(_)) {
                            depth += 1;
                        }
                    }
                    Event::End(_) => depth -= 1,
                    Event::Eof => break,
                    _ => {}
                }
                buffer.clear();
            }
            assert_eq!(dimensions, 1);
            let final_cold = tempfile::tempdir().unwrap();
            std::fs::copy(
                sorted_root.path().join("export.xlsx"),
                final_cold.path().join("source.xlsx"),
            )
            .unwrap();
            let final_index = crate::import::index(final_cold.path()).unwrap();
            assert_eq!(
                final_index["sheets"][0]["usedRangeFacts"]["dimensionCount"],
                1
            );
            assert_eq!(final_index["sheets"][0]["extended"], false);
            assert_eq!(final_index["sheets"][0]["storedCells"], 1203);
            assert_eq!(query(final_cold.path(), 0)["range"], "A1:D1200");
        }
    }
    #[test]
    fn dimension_nested_under_sheet_data_is_guarded_and_preserved_without_repair() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><dimension ref="A1:D1200"/></sheetData></worksheet>"#;
        let root = workbook(xml);
        let before = crate::hash(&root.path().join("source.xlsx")).unwrap();
        let stored: Value =
            serde_json::from_reader(File::open(root.path().join("index.json")).unwrap()).unwrap();
        assert_eq!(stored["sheets"][0]["extended"], true);
        assert_eq!(stored["sheets"][0]["usedRangeFacts"]["dimensionCount"], 0);
        let receipt = query(root.path(), 0);
        assert_eq!(receipt["status"], "incomplete");
        assert!(
            receipt["reasons"]
                .as_array()
                .unwrap()
                .iter()
                .any(|reason| reason["code"] == "opaque-sheet-data-child")
        );
        assert_eq!(export_saved(root.path(), 0), xml);
        let error=crate::transform::apply(root.path(),&json!({"action":"workbook","operation":"transform","root":root.path(),"source":"source.xlsx","version":0,"kind":"sortRows","sheet":"xl/worksheets/cases.xml","range":"A1:A2","keys":[{"column":"A","direction":"asc"}],"view":{"sheets":[{"name":"Cases","path":"xl/worksheets/cases.xml"}],"namedRanges":[],"tables":[]}})).unwrap_err();
        assert!(error.contains("Extended sheetData"), "{error}");
        assert_eq!(
            crate::hash(&root.path().join("source.xlsx")).unwrap(),
            before
        );
        assert_eq!(query(root.path(), 0)["status"], "incomplete");
    }
    #[test]
    fn dimension_hints_are_replaced_or_inserted_without_live_geometry_authority() {
        for declaration in [
            "",
            r#"<dimension ref="A1"/>"#,
            r#"<dimension ref="A1:Z100"/>"#,
        ] {
            let root = workbook(&format!(
                "<worksheet xmlns=\"{}\"><sheetPr/>{declaration}<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><sheetData>{BASE_CELLS}</sheetData></worksheet>",
                String::from_utf8_lossy(crate::xml::NS)
            ));
            let xml = verify_save_and_cold(root.path(), "A1:C6", "A1:C6");
            assert!(xml.find("dimension").unwrap() < xml.find("sheetViews").unwrap());
        }
        for declaration in ["", r#"<dimension ref="A1:Z100"/>"#] {
            let root = workbook(&format!(
                "<worksheet xmlns=\"{}\">{declaration}<sheetData/></worksheet>",
                String::from_utf8_lossy(crate::xml::NS)
            ));
            verify_save_and_cold(root.path(), "A1", "A1");
        }
        let root = workbook(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="C10:C10"/><sheetData><row r="10"><c r="C10"><v>7</v></c></row></sheetData></worksheet>"#,
        );
        verify_save_and_cold(root.path(), "C10", "C10");
        let merge = workbook(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><mergeCells><mergeCell ref="G8:I10"/></mergeCells></worksheet>"#,
        );
        verify_save_and_cold(merge.path(), "G8:I10", "G8:I10");
    }
    #[test]
    fn invalid_axis_styles_unknown_properties_and_opaque_blanks_remain_guarded() {
        for (columns, cells) in [
            (r#"<cols><col min="7" max="7" style="99"/></cols>"#, ""),
            ("", r#"<row r="100" s="99" customFormat="1"/>"#),
            (
                r#"<cols><col min="7" max="7" style="1"/><col min="7" max="8" style="0"/></cols>"#,
                "",
            ),
            ("", r#"<row r="100" hidden="1"/>"#),
            ("", r#"<row r="100" s="1"/>"#),
            ("", r#"<row r="100" ht="0" customHeight="1"/>"#),
            ("", r#"<row r="80"><c r="Z80" cm="1"/></row>"#),
            ("", r#"<row r="80"><c r="Z80"><unknown/></c></row>"#),
            ("", r#"<row r="80"><c r="Z80"><f/></c></row>"#),
            (
                "",
                r#"<row r="100" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" x14ac:dyDescent="NaN"/>"#,
            ),
        ] {
            let root = workbook(&format!(
                "<worksheet xmlns=\"{}\"><dimension ref=\"A1\"/>{columns}<sheetData>{cells}</sheetData></worksheet>",
                String::from_utf8_lossy(crate::xml::NS)
            ));
            assert_eq!(query(root.path(), 0)["status"], "incomplete");
            assert_eq!(query(root.path(), 0)["range"], Value::Null);
            let index: Value =
                serde_json::from_reader(File::open(root.path().join("index.json")).unwrap())
                    .unwrap();
            let db = crate::database(root.path()).unwrap();
            assert_eq!(
                saved_bounds(root.path(), &db, 0, &index["sheets"][0]).unwrap(),
                None
            );
        }
    }
    #[test]
    fn legacy_v1_facts_never_grant_new_saved_normalization_or_missing_row_policy() {
        for (extra, complete) in [
            ("", true),
            (r#"<row r="100"/>"#, false),
            (r#"<row r="100" ht="30" customHeight="1"/>"#, false),
        ] {
            let root = workbook(&format!(
                "<worksheet xmlns=\"{}\"><dimension ref=\"A1:C6\"/><sheetData>{BASE_CELLS}{extra}</sheetData></worksheet>",
                String::from_utf8_lossy(crate::xml::NS)
            ));
            let mut index: Value =
                serde_json::from_reader(File::open(root.path().join("index.json")).unwrap())
                    .unwrap();
            index["sourceFactSchema"] = json!(LEGACY_FACTS_SCHEMA);
            index["sheets"][0]["usedRangeFacts"]["schema"] = json!(LEGACY_FACTS_SCHEMA);
            index["sheets"][0]["usedRangeFacts"]
                .as_object_mut()
                .unwrap()
                .remove("rowTags");
            index["sheets"][0]["usedRangeFacts"]
                .as_object_mut()
                .unwrap()
                .remove("rowEffects");
            std::fs::write(root.path().join("index.json"), index.to_string()).unwrap();
            assert_eq!(
                query(root.path(), 0)["status"],
                if complete { "complete" } else { "incomplete" }
            );
            let db = crate::database(root.path()).unwrap();
            assert_eq!(
                saved_bounds(root.path(), &db, 0, &index["sheets"][0]).unwrap(),
                None
            );
        }
    }
    fn admission(root: &Path, version: u64) -> Value {
        let index: Value =
            serde_json::from_reader(File::open(root.join("index.json")).unwrap()).unwrap();
        export_needed(root,&json!({"action":"workbook","operation":"used-range-export-needed","version":version,"sourceRevision":index["revision"],"diagnostics":true})).unwrap()
    }
    #[test]
    fn export_admission_derives_required_normalization_without_native_mutation() {
        for (dimension, extra, required) in [
            (r#"<dimension ref="A1:C6"/>"#, "", false),
            ("", "", true),
            (r#"<dimension ref="A1:Z100"/>"#, "", true),
            (
                r#"<dimension ref="A1:C6"/>"#,
                r#"<row r="80"><c r="Z80"/></row>"#,
                true,
            ),
            (r#"<dimension ref="A1:C6"/>"#, r#"<row r="100"/>"#, true),
            (
                r#"<dimension ref="A1:Z80"/>"#,
                r#"<row r="80"><c r="Z80" t="inlineStr"><is><t/></is></c></row>"#,
                false,
            ),
            (r#"<dimension ref="A1"/><dimension ref="Z80"/>"#, "", false),
            (
                r#"<dimension ref="A1:Z100"/>"#,
                r#"<row r="100" hidden="1"/>"#,
                false,
            ),
        ] {
            let root = workbook(&format!(
                "<worksheet xmlns=\"{}\">{dimension}<sheetData>{BASE_CELLS}{extra}</sheetData></worksheet>",
                String::from_utf8_lossy(crate::xml::NS)
            ));
            let db = crate::hash(&root.path().join("cells.sqlite")).unwrap();
            let index = crate::hash(&root.path().join("index.json")).unwrap();
            let source = crate::hash(&root.path().join("source.xlsx")).unwrap();
            let reply = admission(root.path(), 0);
            assert_eq!(reply["required"], required);
            assert_eq!(reply.as_object().unwrap().len(), 4);
            assert_eq!(
                reply["schema"],
                workbook_format!(SCHEMA_USED_RANGE_EXPORT_NEEDED_V1)
            );
            assert_eq!(crate::hash(&root.path().join("cells.sqlite")).unwrap(), db);
            assert_eq!(crate::hash(&root.path().join("index.json")).unwrap(), index);
            assert_eq!(
                crate::hash(&root.path().join("source.xlsx")).unwrap(),
                source
            );
        }
    }
    #[test]
    fn export_admission_refuses_stale_physical_source_versions_and_untrusted_fields() {
        let root = basic();
        let index: Value =
            serde_json::from_reader(File::open(root.path().join("index.json")).unwrap()).unwrap();
        let input = json!({"action":"workbook","operation":"used-range-export-needed","version":0,"sourceRevision":index["revision"]});
        for (key, value) in [
            ("bounds", json!("A1")),
            ("required", json!(true)),
            ("diagnostics", json!("true")),
            ("version", json!(99)),
            ("sourceRevision", json!("0".repeat(64))),
        ] {
            let mut invalid = input.clone();
            invalid[key] = value;
            assert!(export_needed(root.path(), &invalid).is_err());
        }
        let mut bytes = std::fs::read(root.path().join("source.xlsx")).unwrap();
        bytes.push(0);
        std::fs::write(root.path().join("source.xlsx"), bytes).unwrap();
        assert!(
            export_needed(root.path(), &input)
                .unwrap_err()
                .contains("source revision is stale")
        );
    }
    #[test]
    fn qualified_sheet_can_require_export_while_another_sheet_remains_guarded() {
        let source = basic();
        let root = tempfile::tempdir().unwrap();
        let mut input =
            zip::ZipArchive::new(File::open(source.path().join("source.xlsx")).unwrap()).unwrap();
        let mut output =
            zip::ZipWriter::new(File::create(root.path().join("source.xlsx")).unwrap());
        let options = zip::write::SimpleFileOptions::default();
        let extra = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><dimension ref="Z80"/><sheetData><row r="80"><c r="Z80"/></row></sheetData></worksheet>"#;
        for i in 0..input.len() {
            let mut member = input.by_index(i).unwrap();
            let name = member.name().to_owned();
            let mut bytes = Vec::new();
            member.read_to_end(&mut bytes).unwrap();
            let mut text = String::from_utf8(bytes).unwrap();
            if name == "[Content_Types].xml" {
                text=text.replace("</Types>",r#"<Override PartName="/xl/worksheets/guarded.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#);
            }
            if name == "xl/workbook.xml" {
                text = text.replace(
                    "</sheets>",
                    r#"<sheet name="Guarded" sheetId="2" r:id="sheet2"/></sheets>"#,
                );
            }
            if name == "xl/_rels/workbook.xml.rels" {
                text=text.replace("</Relationships>",r#"<Relationship Id="sheet2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/guarded.xml"/></Relationships>"#);
            }
            if name == "xl/worksheets/cases.xml" {
                text = text.replace("A1:C6", "A1:Z100");
            }
            output.start_file(name, options).unwrap();
            output.write_all(text.as_bytes()).unwrap();
        }
        output
            .start_file("xl/worksheets/guarded.xml", options)
            .unwrap();
        output.write_all(extra.as_bytes()).unwrap();
        output.finish().unwrap();
        crate::import::index(root.path()).unwrap();
        assert_eq!(admission(root.path(), 0)["required"], true);
        export_saved(root.path(), 0);
        let mut saved =
            zip::ZipArchive::new(File::open(root.path().join("export.xlsx")).unwrap()).unwrap();
        let mut kept = String::new();
        saved
            .by_name("xl/worksheets/guarded.xml")
            .unwrap()
            .read_to_string(&mut kept)
            .unwrap();
        assert_eq!(
            kept, extra,
            "unqualified sheet XML must remain preserved independently"
        );
    }
    #[test]
    fn removing_an_empty_source_row_keeps_new_row_insertions_at_its_boundary() {
        let root = workbook(&format!(
            "<worksheet xmlns=\"{}\"><dimension ref=\"A1:C6\"/><sheetData>{BASE_CELLS}<row r=\"80\"/></sheetData></worksheet>",
            String::from_utf8_lossy(crate::xml::NS)
        ));
        let version=crate::range::patch(root.path(),&json!({"version":0,"changes":[{"sheet":"xl/worksheets/cases.xml","cell":"A60","value":{"value":9,"styleId":0},"edits":[{"value":9}]}]})).unwrap()["version"].as_u64().unwrap();
        assert_eq!(query(root.path(), version)["range"], "A1:C80");
        let xml = export_saved(root.path(), version);
        assert_eq!(saved_dimension(&xml), "A1:C60");
        assert!(xml.contains("r=\"A60\""));
        assert!(!xml.contains("r=\"80\""));
        let cold = tempfile::tempdir().unwrap();
        std::fs::copy(
            root.path().join("export.xlsx"),
            cold.path().join("source.xlsx"),
        )
        .unwrap();
        crate::import::index(cold.path()).unwrap();
        assert_eq!(query(cold.path(), 0)["range"], "A1:C60");
        let db = crate::database(cold.path()).unwrap();
        assert_eq!(
            crate::range::read_db(
                &db,
                &json!({"version":0,"sheet":"xl/worksheets/cases.xml","range":"A60"})
            )
            .unwrap()["cells"][0][1]["value"]
                .as_f64(),
            Some(9.0)
        );
    }
    #[test]
    fn null_formula_clear_fields_do_not_invent_content_or_extend_history() {
        assert!(!content(
            &json!({"value":null,"formula":null,"formulaType":null,"styleId":0})
        ));
        assert!(content(&json!({"value":""})));
        assert!(content(&json!({"value":null,"formula":""})));
        assert!(!has_spill_metadata(&json!({"spill":null})));
        assert!(!has_spill_metadata(&json!({})));
        let root = basic();
        let counts = || {
            let db = crate::database(root.path()).unwrap();
            ["versions", "changes", "used_range_history"].map(|table| {
                db.query_row(&format!("SELECT count(*) FROM {table}"), [], |row| {
                    row.get::<_, u64>(0)
                })
                .unwrap()
            })
        };
        let before = counts();
        let rejected=crate::range::patch(root.path(),&json!({"version":0,"changes":[{"sheet":"xl/worksheets/cases.xml","cell":"Z80","value":{"value":null,"dataType":"blank","styleId":0,"spill":null},"edits":[{"value":null}]}]})).unwrap_err();
        assert!(
            rejected
                .contains("Spill output fields belong to the complete native calculation owner")
        );
        assert_eq!(
            counts(),
            before,
            "unowned null spill metadata cannot acquire a version, change or extent history"
        );
        let version=crate::range::patch(root.path(),&json!({"version":0,"changes":[{"sheet":"xl/worksheets/cases.xml","cell":"Z80","value":{"value":null,"dataType":"blank","rawValue":null,"cachedValue":null,"styleId":0},"edits":[{"value":null}]}]})).unwrap()["version"].as_u64().unwrap();
        assert_eq!(query(root.path(), version)["range"], "A1:C6");
        let xml = export_saved(root.path(), version);
        assert_eq!(saved_dimension(&xml), "A1:C6");
        assert!(!xml.contains("r=\"Z80\""));
        let formatted=crate::range::patch(root.path(),&json!({"version":0,"changes":[{"sheet":"xl/worksheets/cases.xml","cell":"Z80","value":{"value":null,"styleId":0,"fontSize":11},"edits":[{"fontSize":11}]}]})).unwrap()["version"].as_u64().unwrap();
        let index: Value =
            serde_json::from_reader(File::open(root.path().join("index.json")).unwrap()).unwrap();
        assert!(
            saved_bounds(
                root.path(),
                &crate::database(root.path()).unwrap(),
                formatted,
                &index["sheets"][0]
            )
            .unwrap()
            .is_none(),
            "authored blank formatting still needs its resolved export graph"
        );
    }
    #[test]
    fn value_only_aptos_display_fields_do_not_disable_default_blank_clear_normalization() {
        let styles = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts><font><sz val="12"/><color theme="1"/><name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font></fonts><fills><fill><patternFill patternType="none"/></fill></fills><borders><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>"#;
        let root = workbook_parts(
            &format!(
                "<worksheet xmlns=\"{}\"><dimension ref=\"A1:C6\"/><sheetData>{BASE_CELLS}<row r=\"80\"><c r=\"Z80\" s=\"0\"/></row></sheetData></worksheet>",
                String::from_utf8_lossy(crate::xml::NS)
            ),
            None,
            Some(styles),
        );
        let sheet = "xl/worksheets/cases.xml";
        let edited=crate::range::patch(root.path(),&json!({"version":0,"changes":[{"sheet":sheet,"cell":"A1","value":{"format":0,"bold":false,"italic":false,"underline":false,"fontFamily":"Aptos Narrow","fontSize":12,"border":"none","wrapText":false,"value":8,"dataType":"number","rawValue":"8","cachedValue":null,"styleId":0},"edits":[{"value":8}]}]})).unwrap()["version"].as_u64().unwrap();
        let cleared=crate::bulk::apply(root.path(),&json!({"version":edited,"commands":[{"type":"clearRange","sheet":sheet,"range":"Z80"}]})).unwrap()["version"].as_u64().unwrap();
        assert_eq!(query(root.path(), cleared)["range"], "A1:Z80");
        let saved = export_saved(root.path(), cleared);
        assert_eq!(saved_dimension(&saved), "A1:C6");
        assert!(!saved.contains("r=\"Z80\""));
        assert_eq!(query(root.path(), cleared)["range"], "A1:Z80");
        let db = crate::database(root.path()).unwrap();
        assert_eq!(
            db.prepare_cached("SELECT count(*) FROM export_styles")
                .unwrap()
                .query_row([], |r| r.get::<_, u64>(0))
                .unwrap(),
            0,
            "display properties with value-only intent do not manufacture authored style slots"
        );
        drop(db);
        let cold = tempfile::tempdir().unwrap();
        std::fs::copy(
            root.path().join("export.xlsx"),
            cold.path().join("source.xlsx"),
        )
        .unwrap();
        crate::import::index(cold.path()).unwrap();
        assert_eq!(query(cold.path(), 0)["range"], "A1:C6");
        assert_eq!(
            crate::range::read_db(
                &crate::database(cold.path()).unwrap(),
                &json!({"version":0,"sheet":sheet,"range":"A1"})
            )
            .unwrap()["cells"][0][1]["value"]
                .as_f64(),
            Some(8.0)
        );
    }
    #[test]
    fn saved_blank_formatting_uses_actual_resolved_export_graph_instead_of_authored_fields() {
        for (format, style, saved) in [
            (json!({"fontSize":11}), 0, "A1:C6"),
            (json!({"bold":true}), 1, "A1:Z80"),
        ] {
            let root = basic();
            let sheet = "xl/worksheets/cases.xml";
            let version=crate::format::apply(root.path(),&json!({"version":0,"styles":2,"commands":[{"sheet":sheet,"range":"Z80","format":format}]})).unwrap()["version"].as_u64().unwrap();
            assert_eq!(query(root.path(), version)["range"], "A1:Z80");
            let index: Value =
                serde_json::from_reader(File::open(root.path().join("index.json")).unwrap())
                    .unwrap();
            let db = crate::database(root.path()).unwrap();
            assert_eq!(
                saved_bounds(root.path(), &db, version, &index["sheets"][0]).unwrap(),
                None,
                "no resolved output graph is available before export preparation"
            );
            drop(db);
            crate::export_state::prepare(root.path(),&json!({"version":version,"view":{"sheets":[{"name":"Cases","path":sheet}],"namedRanges":[],"tables":[]}})).unwrap();
            let db = crate::database(root.path()).unwrap();
            db.execute_batch(
                "CREATE TEMP TABLE resolved_styles(id INTEGER PRIMARY KEY,style INTEGER NOT NULL)",
            )
            .unwrap();
            assert!(
                saved_export_normalization(root.path(), &db, version, &index["sheets"][0], None)
                    .unwrap()
                    .is_none(),
                "a real authored style slot without a resolved id cannot authorize normalization"
            );
            let mut statement = db
                .prepare_cached("SELECT id FROM export_styles ORDER BY id")
                .unwrap();
            let ids: Vec<Value> = statement
                .query_map([], |r| r.get::<_, u64>(0))
                .unwrap()
                .map(|r| json!({"id":r.unwrap(),"styleId":style}))
                .collect();
            drop(statement);
            drop(db);
            let mut zip =
                zip::ZipArchive::new(File::open(root.path().join("source.xlsx")).unwrap()).unwrap();
            let mut workbook = String::new();
            zip.by_name("xl/workbook.xml")
                .unwrap()
                .read_to_string(&mut workbook)
                .unwrap();
            let replacement=workbook.replace("</workbook>","<calcPr calcMode=\"auto\" calcOnSave=\"1\" fullCalcOnLoad=\"0\" forceFullCalc=\"0\"/></workbook>");
            std::fs::write(
                root.path().join("export-metadata.json"),
                json!({"styleIds":ids,"replacements":{"xl/workbook.xml":replacement}}).to_string(),
            )
            .unwrap();
            crate::export::write(
                root.path(),
                &json!({"revision":index["revision"],"version":version,"calculationMetadata":true}),
            )
            .unwrap();
            let mut zip =
                zip::ZipArchive::new(File::open(root.path().join("export.xlsx")).unwrap()).unwrap();
            let mut xml = String::new();
            zip.by_name(sheet)
                .unwrap()
                .read_to_string(&mut xml)
                .unwrap();
            assert_eq!(saved_dimension(&xml), saved);
            assert_eq!(xml.contains("r=\"Z80\""), style != 0);
            let cold = tempfile::tempdir().unwrap();
            std::fs::copy(
                root.path().join("export.xlsx"),
                cold.path().join("source.xlsx"),
            )
            .unwrap();
            crate::import::index(cold.path()).unwrap();
            assert_eq!(query(cold.path(), 0)["range"], saved);
        }
    }
    #[test]
    fn unowned_special_rectangles_without_native_footprint_proof_remain_guarded() {
        for kind in ["array", "dataTable"] {
            let root = workbook(&format!(
                "<worksheet xmlns=\"{}\"><dimension ref=\"A1:AA100\"/><sheetData><row r=\"80\"><c r=\"Z80\"><f t=\"{kind}\" ref=\"Z80:AA100\">1+1</f><v>2</v></c></row></sheetData></worksheet>",
                String::from_utf8_lossy(crate::xml::NS)
            ));
            let before = crate::hash(&root.path().join("index.json")).unwrap();
            let receipt = query(root.path(), 0);
            assert_eq!(receipt["status"], "incomplete");
            assert_eq!(receipt["range"], Value::Null);
            assert!(
                receipt["reasons"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .any(|r| r["code"] == "unqualified-formula-footprint")
            );
            let index: Value =
                serde_json::from_reader(File::open(root.path().join("index.json")).unwrap())
                    .unwrap();
            let db = crate::database(root.path()).unwrap();
            assert_eq!(
                saved_bounds(root.path(), &db, 0, &index["sheets"][0]).unwrap(),
                None
            );
            assert_eq!(admission(root.path(), 0)["required"], false);
            assert_eq!(
                crate::hash(&root.path().join("index.json")).unwrap(),
                before
            );
        }
        let shared = workbook(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B1"/><sheetData><row r="1"><c r="A1"><f t="shared" ref="A1:B1" si="0">1+1</f><v>2</v></c><c r="B1"><f t="shared" si="0"/><v>2</v></c></row></sheetData></worksheet>"#,
        );
        assert_eq!(query(shared.path(), 0)["status"], "complete");
        assert_eq!(query(shared.path(), 0)["range"], "A1:B1");
        let index: Value =
            serde_json::from_reader(File::open(shared.path().join("index.json")).unwrap()).unwrap();
        assert_eq!(
            saved_bounds(
                shared.path(),
                &crate::database(shared.path()).unwrap(),
                0,
                &index["sheets"][0]
            )
            .unwrap(),
            Some(Bounds::parse("A1:B1").unwrap())
        );
    }
    #[test]
    fn malformed_source_scalar_does_not_authorize_a_guessed_complete_extent() {
        let root = workbook(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="B2"/><sheetData><row r="2"><c r="B2" t="b"><v>not-a-boolean</v></c></row></sheetData></worksheet>"#,
        );
        let reply = query(root.path(), 0);
        assert_eq!(reply["status"], "incomplete");
        assert_eq!(reply["range"], Value::Null);
        assert!(
            reply["reasons"]
                .as_array()
                .unwrap()
                .iter()
                .any(|r| r["code"] == "invalid-scalar")
        );
    }
    #[test]
    fn source_style_index_outside_real_style_table_is_explicitly_incomplete() {
        let root = workbook(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="Z80"/><sheetData><row r="80"><c r="Z80" s="99"/></row></sheetData></worksheet>"#,
        );
        let reply = query(root.path(), 0);
        assert_eq!(reply["status"], "incomplete");
        assert_eq!(reply["range"], Value::Null);
        assert!(
            reply["reasons"]
                .as_array()
                .unwrap()
                .iter()
                .any(|r| r["code"] == "invalid-cell-style")
        );
    }
    #[test]
    fn saved_export_drops_new_cleared_default_cells_and_cold_source_resets_floor() {
        let root = basic();
        let sheet = "xl/worksheets/cases.xml";
        let far=crate::range::patch(root.path(),&json!({"version":0,"changes":[{"sheet":sheet,"cell":"A100","value":{"value":9,"styleId":0},"edits":[{"value":9}]},{"sheet":sheet,"cell":"Z1","value":{"value":11,"styleId":0},"edits":[{"value":11}]}]})).unwrap()["version"].as_u64().unwrap();
        let clear=crate::bulk::apply(root.path(),&json!({"version":far,"commands":[{"type":"clearRange","sheet":sheet,"range":"Z1"},{"type":"clearRange","sheet":sheet,"range":"A100"}]})).unwrap()["version"].as_u64().unwrap();
        assert_eq!(query(root.path(), clear)["lastCell"], "Z100");
        crate::export_state::prepare(root.path(),&json!({"version":clear,"view":{"sheets":[{"name":"Cases","path":sheet}],"namedRanges":[],"tables":[]}})).unwrap();
        let mut source =
            zip::ZipArchive::new(File::open(root.path().join("source.xlsx")).unwrap()).unwrap();
        let mut workbook = String::new();
        source
            .by_name("xl/workbook.xml")
            .unwrap()
            .read_to_string(&mut workbook)
            .unwrap();
        let replacement=workbook.replace("</workbook>","<calcPr calcMode=\"auto\" calcOnSave=\"1\" fullCalcOnLoad=\"0\" forceFullCalc=\"0\"/></workbook>");
        std::fs::write(
            root.path().join("export-metadata.json"),
            json!({"styleIds":[],"replacements":{"xl/workbook.xml":replacement}}).to_string(),
        )
        .unwrap();
        let index: Value =
            serde_json::from_reader(File::open(root.path().join("index.json")).unwrap()).unwrap();
        crate::export::write(
            root.path(),
            &json!({"revision":index["revision"],"version":clear,"calculationMetadata":true}),
        )
        .unwrap();
        let mut saved =
            zip::ZipArchive::new(File::open(root.path().join("export.xlsx")).unwrap()).unwrap();
        let mut xml = String::new();
        saved
            .by_name(sheet)
            .unwrap()
            .read_to_string(&mut xml)
            .unwrap();
        assert!(xml.contains("dimension ref=\"A1:C6\""));
        assert!(!xml.contains("r=\"A100\"") && !xml.contains("r=\"Z1\""));
        assert!(xml.contains("<v>1</v>") && xml.contains("<v>6</v>"));
        let cold = tempfile::tempdir().unwrap();
        std::fs::copy(
            root.path().join("export.xlsx"),
            cold.path().join("source.xlsx"),
        )
        .unwrap();
        crate::import::index(cold.path()).unwrap();
        assert_eq!(query(cold.path(), 0)["range"], "A1:C6");
        assert_eq!(
            query(root.path(), clear)["lastCell"],
            "Z100",
            "export does not rewrite the live pre-Save owner"
        );
    }
    #[test]
    fn known_worksheet_protection_is_non_geometric_and_used_range_reads_are_byte_pure() {
        for protection in [
            r#"<sheetProtection sheet="1"/>"#,
            r#"<sheetProtection sheet="true" objects="0" scenarios="false" password="aB12" selectLockedCells="1" selectUnlockedCells="0"></sheetProtection>"#,
            r#"<p:sheetProtection xmlns:p="http://schemas.openxmlformats.org/spreadsheetml/2006/main" sheet="1"/>"#,
        ] {
            let root = workbook(&format!(
                "<worksheet xmlns=\"{}\"><dimension ref=\"B2:D4\"/><sheetData><row r=\"2\"><c r=\"B2\"><v>2</v></c></row><row r=\"4\"><c r=\"D4\"><v>4</v></c></row></sheetData>{protection}</worksheet>",
                String::from_utf8_lossy(crate::xml::NS)
            ));
            let pins = ["source.xlsx", "metadata.xlsx", "index.json", "cells.sqlite"]
                .map(|name| crate::hash(&root.path().join(name)).unwrap());
            let reply = query(root.path(), 0);
            assert_eq!(reply["status"], "complete");
            assert_eq!(reply["reasons"], json!([]));
            assert_eq!(reply["range"], "B2:D4");
            assert_eq!(reply["lastCell"], "D4");
            assert_eq!(reply["empty"], false);
            assert_eq!(admission(root.path(), 0)["required"], false);
            assert_eq!(
                ["source.xlsx", "metadata.xlsx", "index.json", "cells.sqlite"]
                    .map(|name| crate::hash(&root.path().join(name)).unwrap()),
                pins
            );
        }
        let empty = workbook(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData/><sheetProtection sheet="1"/></worksheet>"#,
        );
        let reply = query(empty.path(), 0);
        assert_eq!(reply["status"], "complete");
        assert_eq!(reply["range"], "A1");
        assert_eq!(reply["lastCell"], "A1");
        assert_eq!(
            reply["empty"], true,
            "Protection never adds a used source position"
        );
    }
    #[test]
    fn opaque_worksheet_protection_keeps_geometry_and_normalization_guarded() {
        for protection in [
            r#"<sheetProtection sheet="maybe"/>"#,
            r#"<sheetProtection sheet="1" password="ZZZZ"/>"#,
            r#"<sheetProtection sheet="1" password="ABC"/>"#,
            r#"<sheetProtection sheet="1" extra="1"/>"#,
            r#"<sheetProtection sheet="1" algorithmName="SHA-512" hashValue="AAAA"/>"#,
            r#"<sheetProtection xmlns:x="urn:opaque" sheet="1" x:sheet="1"/>"#,
            r#"<x:sheetProtection xmlns:x="urn:opaque" sheet="1"/>"#,
            r#"<sheetProtection sheet="1"/><sheetProtection sheet="1"/>"#,
            r#"<sheetViews><sheetProtection sheet="1"/></sheetViews>"#,
            r#"<sheetProtection sheet="1"><sheetProtection sheet="1"/></sheetProtection>"#,
            r#"<sheetProtection sheet="1"><unknown/></sheetProtection>"#,
            r#"<sheetProtection sheet="1">opaque</sheetProtection>"#,
            r#"<sheetProtection sheet="1"><![CDATA[opaque]]></sheetProtection>"#,
            r#"<sheetProtection sheet="1"><!--opaque--></sheetProtection>"#,
            r#"<sheetProtection sheet="1"><?opaque retain?></sheetProtection>"#,
            r#"<sheetProtection sheet="1"/><conditionalFormatting sqref="B2"><cfRule type="expression" priority="1"><formula>1=1</formula></cfRule></conditionalFormatting>"#,
            r#"<sheetProtection sheet="1"/><extLst><ext uri="opaque"/></extLst>"#,
        ] {
            let xml = format!(
                "<worksheet xmlns=\"{}\"><dimension ref=\"A1:Z80\"/><sheetData><row r=\"2\"><c r=\"B2\"><v>2</v></c></row><row r=\"80\"><c r=\"Z80\"/></row></sheetData>{protection}</worksheet>",
                String::from_utf8_lossy(crate::xml::NS)
            );
            let root = workbook(&xml);
            let reply = query(root.path(), 0);
            assert_eq!(reply["status"], "incomplete", "{protection}");
            for key in ["scopeMinMax", "range", "lastCell", "empty"] {
                assert_eq!(reply[key], Value::Null, "{protection}: {key}");
            }
            assert!(reply["reasons"].as_array().unwrap().iter().any(|reason| {
                [
                    "unqualified-sheet-protection",
                    "duplicate-worksheet-structure",
                    "unqualified-worksheet-metadata",
                    "extended-source",
                ]
                .contains(&reason["code"].as_str().unwrap())
            }));
            let index: Value =
                serde_json::from_reader(File::open(root.path().join("index.json")).unwrap())
                    .unwrap();
            assert_eq!(
                saved_bounds(
                    root.path(),
                    &crate::database(root.path()).unwrap(),
                    0,
                    &index["sheets"][0]
                )
                .unwrap(),
                None,
                "{protection}"
            );
            assert_eq!(admission(root.path(), 0)["required"], false);
            let source = std::fs::read(root.path().join("source.xlsx")).unwrap();
            assert_eq!(export_saved(root.path(), 0), xml, "{protection}");
            assert_eq!(
                std::fs::read(root.path().join("source.xlsx")).unwrap(),
                source
            );
        }
    }
    #[test]
    fn explicit_native_normalization_preserves_known_protection_metadata_and_source_bytes() {
        // Ordinary protected-owner Copy/Save returns source.xlsx before this
        // writer is called. An explicit native writer may normalize geometry,
        // but must never rewrite or remove the recognized protection record.
        let protection = r#"<sheetProtection password="A1B2" sheet="1" objects="1" scenarios="0" formatCells="1" formatColumns="1" formatRows="1" insertColumns="1" insertRows="1" insertHyperlinks="1" deleteColumns="1" deleteRows="1" selectLockedCells="0" sort="1" autoFilter="1" pivotTables="1" selectUnlockedCells="0"/>"#;
        let root = workbook(&format!(
            "<worksheet xmlns=\"{}\"><dimension ref=\"A1:Z80\"/><sheetData>{BASE_CELLS}<row r=\"80\"><c r=\"Z80\"/></row></sheetData>{protection}</worksheet>",
            String::from_utf8_lossy(crate::xml::NS)
        ));
        let source = std::fs::read(root.path().join("source.xlsx")).unwrap();
        let metadata = std::fs::read(root.path().join("metadata.xlsx")).unwrap();
        assert_eq!(query(root.path(), 0)["range"], "A1:Z80");
        assert_eq!(admission(root.path(), 0)["required"], true);
        let saved = export_saved(root.path(), 0);
        assert!(saved.contains("dimension ref=\"A1:C6\""));
        assert!(!saved.contains("r=\"Z80\""));
        assert_eq!(saved.matches(protection).count(), 1);
        assert_eq!(
            std::fs::read(root.path().join("source.xlsx")).unwrap(),
            source
        );
        assert_eq!(
            std::fs::read(root.path().join("metadata.xlsx")).unwrap(),
            metadata
        );
        assert_eq!(query(root.path(), 0)["range"], "A1:Z80");
        let cold = tempfile::tempdir().unwrap();
        std::fs::copy(
            root.path().join("export.xlsx"),
            cold.path().join("source.xlsx"),
        )
        .unwrap();
        crate::import::index(cold.path()).unwrap();
        assert_eq!(query(cold.path(), 0)["range"], "A1:C6");
    }
    #[test]
    fn applied_default_alignment_blanks_normalize_but_real_alignment_is_retained() {
        for (flag, alignment, persistent) in [
            ("1", "vertical=\"bottom\"", false),
            ("true", "vertical=\"bottom\" wrapText=\"false\"", false),
            ("1", "horizontal=\"general\" wrapText=\"0\"", false),
            ("1", "vertical=\"top\"", true),
            ("1", "vertical=\"center\"", true),
            ("1", "horizontal=\"center\"", true),
            ("true", "wrapText=\"true\"", true),
        ] {
            let styles = format!(
                "<styleSheet xmlns=\"{}\"><fonts><font><name val=\"Calibri\"/><sz val=\"11\"/></font></fonts><fills><fill><patternFill patternType=\"none\"/></fill></fills><borders><border/></borders><cellStyleXfs><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"{flag}\"><alignment {alignment}/></xf></cellXfs></styleSheet>",
                String::from_utf8_lossy(crate::xml::NS)
            );
            let root = workbook_parts(
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="B2:Z80"/><sheetData><row r="2"><c r="B2"><v>2</v></c></row><row r="80"><c r="Z80" s="1"/></row></sheetData></worksheet>"#,
                None,
                Some(&styles),
            );
            let pins = ["source.xlsx", "metadata.xlsx", "index.json", "cells.sqlite"]
                .map(|name| crate::hash(&root.path().join(name)).unwrap());
            let live = query(root.path(), 0);
            assert_eq!(live["status"], "complete", "{alignment}");
            assert_eq!(live["reasons"], json!([]));
            assert_eq!(live["range"], "B2:Z80");
            assert_eq!(live["lastCell"], "Z80");
            assert_eq!(admission(root.path(), 0)["required"], !persistent);
            assert_eq!(
                ["source.xlsx", "metadata.xlsx", "index.json", "cells.sqlite"]
                    .map(|name| crate::hash(&root.path().join(name)).unwrap()),
                pins,
                "Style classification and export admission are read-only"
            );
            let saved = export_saved(root.path(), 0);
            let expected = if persistent { "B2:Z80" } else { "B2" };
            assert_eq!(saved_dimension(&saved), expected, "{alignment}");
            assert_eq!(saved.contains("r=\"Z80\""), persistent, "{alignment}");
            assert!(saved.contains("r=\"B2\"") && saved.contains("<v>2</v>"));
            assert_eq!(
                ["source.xlsx", "metadata.xlsx", "index.json"]
                    .map(|name| crate::hash(&root.path().join(name)).unwrap()),
                [pins[0].clone(), pins[1].clone(), pins[2].clone()],
                "Writing a normalized package leaves accepted source metadata intact"
            );
            assert_eq!(query(root.path(), 0)["range"], "B2:Z80");
            let cold = tempfile::tempdir().unwrap();
            std::fs::copy(
                root.path().join("export.xlsx"),
                cold.path().join("source.xlsx"),
            )
            .unwrap();
            crate::import::index(cold.path()).unwrap();
            assert_eq!(query(cold.path(), 0)["status"], "complete");
            assert_eq!(query(cold.path(), 0)["range"], expected);
        }
    }
    #[test]
    fn applied_general_and_closed_home_numeric_styles_normalize_unrelated_blanks() {
        for code in ["0.0", "\"$\"#,##0.00;(\"$\"#,##0.00)"] {
            for flag in ["1", "true"] {
                let styles = format!(
                    "<styleSheet xmlns=\"{}\"><numFmts><numFmt numFmtId=\"164\" formatCode='{code}'/></numFmts><fonts><font><name val=\"Calibri\"/><sz val=\"11\"/></font></fonts><fills><fill><patternFill patternType=\"none\"/></fill></fills><borders><border/></borders><cellStyleXfs><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"{flag}\"/><xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"{flag}\"/></cellXfs></styleSheet>",
                    String::from_utf8_lossy(crate::xml::NS)
                );
                let root = workbook_parts(
                    r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:Z80"/><sheetData><row r="2"><c r="B2"><v>2</v></c></row><row r="3"><c r="C3" s="2"><v>3</v></c></row><row r="4"><c r="D4" s="2"/></row><row r="6"><c r="A6" s="1"/></row><row r="80"><c r="Z80"/></row></sheetData></worksheet>"#,
                    None,
                    Some(&styles),
                );
                let pins = ["source.xlsx", "metadata.xlsx", "index.json", "cells.sqlite"]
                    .map(|name| crate::hash(&root.path().join(name)).unwrap());
                let live = query(root.path(), 0);
                assert_eq!(live["status"], "complete", "{code}: {flag}");
                assert_eq!(live["reasons"], json!([]));
                assert_eq!(live["range"], "A2:Z80");
                assert_eq!(admission(root.path(), 0)["required"], true);
                assert_eq!(
                    ["source.xlsx", "metadata.xlsx", "index.json", "cells.sqlite"]
                        .map(|name| crate::hash(&root.path().join(name)).unwrap()),
                    pins,
                    "Query and admission cannot mutate accepted source or its index"
                );
                let saved = export_saved(root.path(), 0);
                assert_eq!(saved_dimension(&saved), "B2:D4", "{code}: {flag}");
                assert!(!saved.contains("r=\"A6\"") && !saved.contains("r=\"Z80\""));
                assert!(saved.contains("r=\"B2\"") && saved.contains("<v>2</v>"));
                assert!(saved.contains("r=\"C3\" s=\"2\"") && saved.contains("<v>3</v>"));
                assert!(
                    saved.contains("r=\"D4\" s=\"2\""),
                    "Numeric-format blanks remain persistent"
                );
                let mut zip =
                    zip::ZipArchive::new(File::open(root.path().join("export.xlsx")).unwrap())
                        .unwrap();
                let mut saved_styles = String::new();
                zip.by_name("xl/styles.xml")
                    .unwrap()
                    .read_to_string(&mut saved_styles)
                    .unwrap();
                assert_eq!(
                    saved_styles, styles,
                    "Normalization preserves the exact numeric style graph"
                );
                assert_eq!(
                    ["source.xlsx", "metadata.xlsx", "index.json"]
                        .map(|name| crate::hash(&root.path().join(name)).unwrap()),
                    [pins[0].clone(), pins[1].clone(), pins[2].clone()],
                    "Writing a normalized package leaves accepted source metadata intact"
                );
                assert_eq!(query(root.path(), 0)["range"], "A2:Z80");
                let cold = tempfile::tempdir().unwrap();
                std::fs::copy(
                    root.path().join("export.xlsx"),
                    cold.path().join("source.xlsx"),
                )
                .unwrap();
                crate::import::index(cold.path()).unwrap();
                assert_eq!(query(cold.path(), 0)["status"], "complete");
                assert_eq!(query(cold.path(), 0)["range"], "B2:D4");
                assert_eq!(admission(cold.path(), 0)["required"], false);
            }
        }
    }
}
impl Bounds {
    pub fn point(row: u32, col: u32) -> Self {
        Self {
            min_row: row,
            min_column: col,
            max_row: row,
            max_column: col,
        }
    }
    pub fn union(self, other: Self) -> Self {
        Self {
            min_row: self.min_row.min(other.min_row),
            min_column: self.min_column.min(other.min_column),
            max_row: self.max_row.max(other.max_row),
            max_column: self.max_column.max(other.max_column),
        }
    }
    pub fn range(self) -> String {
        let a = crate::xml::address(self.min_row, self.min_column);
        let b = crate::xml::address(self.max_row, self.max_column);
        if a == b { a } else { format!("{a}:{b}") }
    }
    pub fn parse(reference: &str) -> Result<Self> {
        let (a, b) = crate::bulk::bounds(reference)?;
        let result = Self {
            min_row: a.0,
            min_column: a.1,
            max_row: b.0,
            max_column: b.1,
        };
        if reference != result.range()
            && !(a == b && reference == format!("{}:{}", result.range(), result.range()))
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Noncanonical worksheet extent reference",
            ));
        }
        Ok(result)
    }
}
pub(crate) fn add(bounds: &mut Option<Bounds>, value: Bounds) {
    *bounds = Some(bounds.map_or(value, |old| old.union(value)));
}

#[derive(Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(crate) struct Fact {
    pub count: u64,
    pub bounds: Option<Bounds>,
}
impl Fact {
    fn include(&mut self, bounds: Bounds) {
        self.count += 1;
        add(&mut self.bounds, bounds);
    }
}

/// Fixed-size aggregates; no per-cell or per-axis list is retained.
#[derive(Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(crate) struct SourceFacts {
    pub schema: String,
    pub dimension_count: u64,
    pub dimension_ref: Option<String>,
    pub dimension_bounds: Option<Bounds>,
    pub cell_tags: Fact,
    pub content: Fact,
    pub styled_blanks: Fact,
    pub unstyled_blanks: Fact,
    pub explicit_default_blanks: Fact,
    pub empty_text: Fact,
    pub row_styles: Fact,
    pub column_styles: Fact,
    #[serde(default)]
    pub row_tags: Fact,
    #[serde(default)]
    pub row_effects: Fact,
    pub merges: Fact,
    pub geometry_defaults: bool,
    pub issues: Vec<Value>,
}
impl SourceFacts {
    pub fn new() -> Self {
        Self {
            schema: FACTS_SCHEMA.into(),
            ..Self::default()
        }
    }
    /// Older v1 import aggregates used the SST ordinal as if it were literal
    /// text. Resolve through the same indexed decoder used for cell reads, then
    /// rebuild only fixed-size source aggregates when that distinction matters.
    /// Public scalar values and immutable XML/SQL remain unchanged.
    pub(crate) fn resolve_shared_strings(&mut self, db: &Connection, sheet: &str) -> Result<()> {
        let select = "SELECT c.r,c.c,c.kind,c.raw,c.inline,c.formula,c.formula_attrs,c.style,s.value,NULL,c.style_explicit FROM cells c LEFT JOIN strings s ON c.kind='s' AND s.id=CAST(c.raw AS INTEGER) WHERE c.sheet=?1";
        let mut shared = db
            .prepare_cached(&format!("{select} AND c.kind='s' ORDER BY c.r,c.c"))
            .map_err(err)?;
        let mut rows = shared.query([sheet]).map_err(err)?;
        let mut rebuild = false;
        while let Some(row) = rows.next().map_err(err)? {
            crate::cancel::check()?;
            match crate::range::decode_cell(row) {
                Ok((_, cell)) => {
                    if cell["value"].as_str() == Some("") {
                        rebuild = true;
                    }
                }
                Err(_) => {
                    rebuild = true;
                    self.issue("invalid-shared-string","A source shared-string reference cannot be resolved by the canonical imported cell decoder.");
                }
            }
        }
        if !rebuild {
            return Ok(());
        }
        self.cell_tags = Fact::default();
        self.content = Fact::default();
        self.styled_blanks = Fact::default();
        self.unstyled_blanks = Fact::default();
        self.explicit_default_blanks = Fact::default();
        self.empty_text = Fact::default();
        let mut all = db
            .prepare_cached(&format!("{select} ORDER BY c.r,c.c"))
            .map_err(err)?;
        let mut rows = all.query([sheet]).map_err(err)?;
        while let Some(row) = rows.next().map_err(err)? {
            crate::cancel::check()?;
            let r: u32 = row.get(0).map_err(err)?;
            let c: u32 = row.get(1).map_err(err)?;
            let kind: String = row.get(2).map_err(err)?;
            let raw: Option<String> = row.get(3).map_err(err)?;
            let inline: String = row.get(4).map_err(err)?;
            let formula = row.get::<_, Option<String>>(5).map_err(err)?.is_some();
            let style: u32 = row.get(7).map_err(err)?;
            let explicit: bool = row.get(10).map_err(err)?;
            let resolved = if kind == "s" {
                match crate::range::decode_cell(row) {
                    Ok((_, cell)) => cell["value"].as_str().map(str::to_owned),
                    Err(_) => {
                        self.issue("invalid-shared-string","A source shared-string reference cannot be resolved by the canonical imported cell decoder.");
                        None
                    }
                }
            } else {
                None
            };
            self.cell(
                r,
                c,
                &kind,
                if kind == "s" {
                    resolved.as_deref().or(raw.as_deref())
                } else {
                    raw.as_deref()
                },
                &inline,
                formula,
                style,
                explicit,
            );
        }
        Ok(())
    }
    pub fn issue(&mut self, code: &str, message: &str) {
        if self.issues.len() < 16 && !self.issues.iter().any(|x| x["code"] == code) {
            self.issues.push(json!({"code":code,"message":message}));
        }
    }
    pub fn dimension(&mut self, reference: Option<&str>) {
        self.dimension_count += 1;
        if self.dimension_count == 1 {
            self.dimension_ref = reference.filter(|r| r.len() <= 128).map(str::to_owned);
            self.dimension_bounds = reference.and_then(|r| Bounds::parse(r).ok());
        }
    }
    pub fn declaration(&self) -> Value {
        let status = if self.dimension_count == 0 {
            "absent"
        } else if self.dimension_count > 1 {
            "duplicate"
        } else if self.dimension_bounds.is_some() {
            "valid"
        } else {
            "invalid"
        };
        json!({"status":status,"ref":self.dimension_ref,"bounds":if self.dimension_count==1{self.dimension_bounds}else{None}})
    }
    pub fn cell(
        &mut self,
        row: u32,
        col: u32,
        kind: &str,
        raw: Option<&str>,
        inline: &str,
        formula: bool,
        style: u32,
        explicit: bool,
    ) {
        let p = Bounds::point(row, col);
        self.cell_tags.include(p);
        if !["n", "s", "inlineStr", "str", "b", "e", "d"].contains(&kind) {
            self.issue(
                "unsupported-cell-type",
                "A source cell has an unrecognized OOXML type.",
            );
        }
        if kind == "b"
            && (raw.is_some_and(|r| !r.is_empty() && !matches!(r, "0" | "1"))
                || !formula && raw.is_none_or(str::is_empty))
            || kind == "n"
                && raw.is_some_and(|r| {
                    !r.is_empty() && r.parse::<f64>().ok().is_none_or(|v| !v.is_finite())
                })
        {
            self.issue(
                "invalid-scalar",
                "A source numeric or logical cell has invalid scalar text.",
            );
        }
        let text = matches!(kind, "s" | "inlineStr" | "str");
        let authored =
            formula || raw.is_some_and(|value| !value.is_empty()) || text && !inline.is_empty();
        if authored {
            self.content.include(p);
        } else if style != 0 {
            self.styled_blanks.include(p);
        } else if explicit {
            self.explicit_default_blanks.include(p);
        } else {
            self.unstyled_blanks.include(p);
        }
        if !formula && text && raw.is_none_or(str::is_empty) && inline.is_empty() {
            self.empty_text.include(p);
        }
    }
    pub(crate) fn cell_attributes(&mut self, attrs: &BTreeMap<&str, Cow<'_, str>>) {
        if attrs
            .keys()
            .any(|key| !["r", "s", "t", "xmlns"].contains(key) && !key.starts_with("xmlns:"))
        {
            self.issue(
                "opaque-cell-metadata",
                "Imported cell attributes require preservation without normalized removal.",
            );
        }
    }
    pub(crate) fn capture_axis(
        &mut self,
        db: &Connection,
        sheet: &str,
        name: &str,
        attrs: &BTreeMap<&str, Cow<'_, str>>,
        row: Option<u32>,
    ) -> Result<()> {
        self.element(name, attrs, row);
        let boolean = |key: &str| -> Result<Option<bool>> {
            attrs
                .get(key)
                .map(|v| match v.as_ref() {
                    "1" | "true" => Ok(true),
                    "0" | "false" => Ok(false),
                    _ => Err(crate::coded(
                        "internal-invalid-field",
                        "Invalid axis boolean",
                    )),
                })
                .transpose()
        };
        let (axis, first, last, style, extent, removable) = if name == "row" {
            let r = row.ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing source row coordinate")
            })?;
            self.row_tags.include(Bounds::point(r, 0));
            let style = attrs
                .get("s")
                .and_then(|v| v.parse::<u32>().ok())
                .unwrap_or(0);
            let cf = boolean("customFormat");
            let ch = boolean("customHeight");
            if cf.is_err() || ch.is_err() {
                self.issue(
                    "invalid-row-policy",
                    "Row custom-format or height flag is invalid.",
                );
            }
            let cf = cf.ok().flatten();
            let ch = ch.ok().flatten();
            if attrs.get("x14ac:dyDescent").is_some_and(|v| {
                v.parse::<f64>()
                    .ok()
                    .is_none_or(|n| !n.is_finite() || n < 0.0)
            }) {
                self.issue(
                    "invalid-row-policy",
                    "Row descent is not finite nonnegative geometry.",
                );
            }
            let height = attrs
                .get("ht")
                .and_then(|v| v.parse::<f64>().ok())
                .filter(|h| h.is_finite() && *h > 0.0);
            if attrs.contains_key("ht") && (height.is_none() || ch != Some(true))
                || ch == Some(true) && height.is_none()
                || style != 0 && cf != Some(true)
                || cf == Some(false)
            {
                self.issue(
                    "unqualified-row-policy",
                    "Row style/height/default flags differ from the qualified imported profiles.",
                );
            }
            let known = attrs.keys().all(|key| {
                [
                    "r",
                    "spans",
                    "s",
                    "ht",
                    "customHeight",
                    "customFormat",
                    "x14ac:dyDescent",
                    "xmlns",
                ]
                .contains(key)
                    || key.starts_with("xmlns:")
            });
            if !known {
                self.issue(
                    "opaque-row-metadata",
                    "Unknown row properties cannot authorize normalized row removal.",
                );
            }
            let effect = cf == Some(true) || height.is_some();
            if effect {
                self.row_effects.include(Bounds::point(r, 0));
            }
            ("row", r, r, style, effect, known && !effect && style == 0)
        } else if name == "col" {
            let first = attrs.get("min").and_then(|v| v.parse::<u32>().ok());
            let last = attrs.get("max").and_then(|v| v.parse::<u32>().ok());
            let (Some(first), Some(last)) = (first, last) else {
                self.issue(
                    "invalid-column-style",
                    "Column declarations lack valid bounds.",
                );
                return Ok(());
            };
            if first == 0 || first > last || last > 16384 {
                self.issue(
                    "invalid-column-style",
                    "Column declarations have invalid bounds.",
                );
                return Ok(());
            }
            let style = attrs
                .get("style")
                .and_then(|v| v.parse::<u32>().ok())
                .unwrap_or(0);
            if attrs.get("width").is_some_and(|v| {
                v.parse::<f64>()
                    .ok()
                    .is_none_or(|w| !w.is_finite() || w <= 0.0)
            }) || boolean("customWidth").is_err()
            {
                self.issue("invalid-column-policy", "Column width metadata is invalid.");
            }
            if attrs.keys().any(|key| {
                !["min", "max", "style", "width", "customWidth", "xmlns"].contains(key)
                    && !key.starts_with("xmlns:")
            }) {
                self.issue(
                    "opaque-column-metadata",
                    "Unknown column properties have no qualified extent interpretation.",
                );
            }
            let overlap:bool=db.prepare_cached("SELECT EXISTS(SELECT 1 FROM used_range_axes WHERE sheet=?1 AND axis='col' AND first<=?3 AND last>=?2)").unwrap().query_row(params![sheet,first-1,last-1],|r|r.get(0)).map_err(err)?;
            if overlap {
                self.issue(
                    "overlapping-column-policy",
                    "Overlapping column declarations require independently qualified precedence.",
                );
            }
            ("col", first - 1, last - 1, style, false, false)
        } else {
            return Ok(());
        };
        db.execute("INSERT INTO used_range_axes(sheet,axis,first,last,style,extent,removable) VALUES(?1,?2,?3,?4,?5,?6,?7)",params![sheet,axis,first,last,style,extent,removable]).map_err(err)?;
        Ok(())
    }
    pub fn element(&mut self, name: &str, attrs: &BTreeMap<&str, Cow<'_, str>>, row: Option<u32>) {
        if name == "sheetFormatPr" {
            self.geometry_defaults = true;
            if attrs.keys().any(|key| {
                ![
                    "baseColWidth",
                    "defaultColWidth",
                    "defaultRowHeight",
                    "x14ac:dyDescent",
                    "xmlns",
                ]
                .contains(key)
                    && !key.starts_with("xmlns:")
            }) {
                self.issue(
                    "unqualified-sheet-defaults",
                    "Sheet defaults outside the observed geometry profile remain protected.",
                );
            }
            if attrs.get("x14ac:dyDescent").is_some_and(|v| {
                v.parse::<f64>()
                    .ok()
                    .is_none_or(|n| !n.is_finite() || n < 0.0)
            }) {
                self.issue(
                    "invalid-sheet-defaults",
                    "Sheet descent is not finite nonnegative geometry.",
                );
            }
            if attrs
                .get("baseColWidth")
                .is_some_and(|v| v.parse::<u32>().ok().is_none_or(|n| n > 255))
            {
                self.issue(
                    "invalid-sheet-defaults",
                    "Sheet base column width is invalid.",
                );
            }
            for key in ["defaultColWidth", "defaultRowHeight"] {
                if attrs.get(key).is_some_and(|v| {
                    v.parse::<f64>()
                        .ok()
                        .is_none_or(|n| !n.is_finite() || n <= 0.0)
                }) {
                    self.issue(
                        "invalid-sheet-defaults",
                        "Sheet defaults have invalid numeric geometry.",
                    );
                }
            }
        }
        if name == "row"
            && let Some(style) = attrs.get("s")
        {
            let valid = style.parse::<u32>().ok();
            if valid.is_none() {
                self.issue(
                    "invalid-row-style",
                    "A source row has an invalid style index.",
                );
            }
            if valid.is_some_and(|s| s != 0)
                && let Some(r) = row
            {
                self.row_styles.include(Bounds {
                    min_row: r,
                    min_column: 0,
                    max_row: r,
                    max_column: 16383,
                });
            }
        }
        if name == "col"
            && let Some(style) = attrs.get("style")
        {
            let parsed = (
                attrs.get("min").and_then(|s| s.parse::<u32>().ok()),
                attrs.get("max").and_then(|s| s.parse::<u32>().ok()),
                style.parse::<u32>().ok(),
            );
            if let (Some(a), Some(b), Some(s)) = parsed
                && a > 0
                && a <= b
                && b <= 16384
            {
                if s != 0 {
                    self.column_styles.include(Bounds {
                        min_row: 0,
                        min_column: a - 1,
                        max_row: 1048575,
                        max_column: b - 1,
                    });
                }
            } else {
                self.issue(
                    "invalid-column-style",
                    "A source column style has invalid bounds or style index.",
                );
            }
        }
        if name == "mergeCell" {
            if attrs
                .keys()
                .any(|key| !["ref", "xmlns"].contains(key) && !key.starts_with("xmlns:"))
            {
                self.issue(
                    "opaque-merge",
                    "Extended merge attributes remain protected.",
                );
            }
            if let Some(b) = attrs.get("ref").and_then(|s| Bounds::parse(s).ok()) {
                self.merges.include(b);
            } else {
                self.issue("invalid-merge", "A source merge has invalid bounds.");
            }
        }
    }
    pub fn value(&self) -> Result<Value> {
        serde_json::to_value(self).map_err(err)
    }
}
