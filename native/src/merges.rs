//! Merged ranges are accepted-version metadata, never renderer projections.
//! Whole collections are small and inherited like cell/calculation overlays.
use crate::{
    Result,
    bulk::bounds,
    bulk::fill_guard::{part, walk},
    err, local,
    xml::{NS, address},
};
use rusqlite::{Connection, OptionalExtension, params};
use serde::{Deserialize, Serialize};
use serde_json::{Value, json};
use std::{
    collections::{BTreeMap, BTreeSet},
    fs::File,
    path::Path,
};
use zip::ZipArchive;
#[derive(Clone, Debug, Deserialize, Serialize)]
#[serde(tag = "operation", rename_all = "camelCase", deny_unknown_fields)]
pub enum MergeCommand {
    MergeCells { sheet: String, range: String },
    UnmergeCells { sheet: String, range: String },
}
impl MergeCommand {
    pub fn target(&self) -> (&str, &str) {
        match self {
            Self::MergeCells { sheet, range } | Self::UnmergeCells { sheet, range } => {
                (sheet, range)
            }
        }
    }
    pub fn map_sheet(&mut self, path: String) {
        let (Self::MergeCells { sheet, .. } | Self::UnmergeCells { sheet, .. }) = self;
        *sheet = path;
    }
}
fn present(db: &Connection) -> Result<bool> {
    db.prepare_cached(
        "SELECT EXISTS(SELECT 1 FROM sqlite_master WHERE type='table' AND name='version_merges')",
    )
    .map_err(crate::err)?
    .query_row([], |row| row.get(0))
    .map_err(err)
}
pub(crate) fn selected_override(
    db: &Connection,
    version: u64,
    sheet: &str,
) -> Result<Option<Vec<String>>> {
    if !present(db)? {
        return Ok(None);
    }
    let encoded:Option<String>=db.prepare_cached("WITH RECURSIVE a(id,depth) AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL) SELECT m.ranges FROM a JOIN version_merges m ON m.version=a.id WHERE m.sheet=?2 ORDER BY a.depth LIMIT 1").map_err(crate::err)?.query_row(params![version,sheet],|row|row.get(0)).optional().map_err(err)?;
    encoded
        .map(|encoded| {
            if encoded.len() > 1024 * 1024 {
                return Err(crate::coded(
                    "merge-limit",
                    "Merged-range history exceeds its byte bound",
                ));
            }
            let ranges: Vec<String> = serde_json::from_str(&encoded).map_err(err)?;
            if ranges.len() > 10000 {
                return Err(crate::coded(
                    "merge-limit",
                    "Merged-range history exceeds its count bound",
                ));
            }
            for range in &ranges {
                bounds(range)?;
            }
            Ok(ranges)
        })
        .transpose()
}
fn overlaps(a: crate::bulk::Bounds, b: crate::bulk::Bounds) -> bool {
    a.0.0 <= b.1.0 && a.1.0 >= b.0.0 && a.0.1 <= b.1.1 && a.1.1 >= b.0.1
}
fn canonical(range: &str) -> Result<String> {
    let (first, last) = bounds(range)?;
    Ok(format!(
        "{}:{}",
        address(first.0, first.1),
        address(last.0, last.1)
    ))
}
struct Source {
    ranges: Vec<String>,
    protected: bool,
}
fn source(root: &Path, sheet: &str) -> Result<Source> {
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let mut read = 0;
    let bytes = part(&mut zip, sheet, &mut read)?;
    let mut source = Source {
        ranges: Vec::new(),
        protected: false,
    };
    let (mut containers, mut count) = (0usize, None);
    walk(&bytes, NS, "worksheet", |stack, node| {
        // Merging empty cells moves nothing, so rules, links, notes and drawings
        // are unaffected. Excel itself refuses merges inside a table, and a
        // protected sheet is not edited.
        if ["tableParts", "sheetProtection"].contains(&node.name.as_str()) {
            source.protected = true;
        }
        match node.name.as_str() {
            "mergeCells" => {
                containers += 1;
                if containers != 1 || stack.len() != 1 || node.namespace != NS {
                    return Err(crate::coded(
                        "merge-ambiguous",
                        "Ambiguous merged-range container",
                    ));
                }
                for key in node.attrs.keys() {
                    if key != "count" && key != "xmlns" && !key.starts_with("xmlns:") {
                        return Err(crate::coded(
                            "merge-unsupported",
                            "Unqualified merged-range attributes",
                        ));
                    }
                }
                count = node
                    .attrs
                    .get("count")
                    .map(|value| value.parse::<usize>().map_err(err))
                    .transpose()?;
            }
            "mergeCell" => {
                if node.namespace != NS || stack.len() != 2 || stack[1].name != "mergeCells" {
                    return Err(crate::coded(
                        "merge-unsupported",
                        "Orphan merged-range metadata",
                    ));
                }
                for key in node.attrs.keys() {
                    if key != "ref" && key != "xmlns" && !key.starts_with("xmlns:") {
                        return Err(crate::coded(
                            "merge-unsupported",
                            "Unqualified merged-cell attributes",
                        ));
                    }
                }
                let range = canonical(node.attrs.get("ref").ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing merged range")
                })?)?;
                let area = bounds(&range)?;
                if area.0 == area.1
                    || source.ranges.len() >= 10000
                    || source
                        .ranges
                        .iter()
                        .any(|other| bounds(other).is_ok_and(|other| overlaps(area, other)))
                {
                    return Err(crate::coded(
                        "internal-invalid-field",
                        "Invalid or overlapping source merges",
                    ));
                }
                source.ranges.push(range);
            }
            _ => {}
        }
        Ok(())
    })?;
    if count.is_some_and(|count| count != source.ranges.len()) {
        return Err(crate::coded(
            "internal-mismatch",
            "Merged-range count differs from its entries",
        ));
    }
    Ok(source)
}
pub(crate) fn selected(
    root: &Path,
    db: &Connection,
    version: u64,
    sheet: &str,
) -> Result<Vec<String>> {
    if let Some(ranges) = selected_override(db, version, sheet)? {
        Ok(ranges)
    } else {
        Ok(source(root, sheet)?.ranges)
    }
}
fn check_cells(db: &Connection, version: u64, sheet: &str, range: &str) -> Result<()> {
    // Formulas elsewhere on the sheet are unaffected: merging moves nothing, and
    // every hidden cell must already be empty. A formula in a hidden cell counts
    // as content.
    let (first, _) = bounds(range)?;
    let anchor = address(first.0, first.1);
    let mut after: Option<String> = None;
    loop {
        crate::cancel::check()?;
        let page = crate::range::read_db(
            db,
            &json!({"version":version,"sheet":sheet,"range":range,"after":after,"limit":512,"maxBytes":2*1024*1024}),
        )?;
        for pair in page["cells"].as_array().ok_or_else(|| {
            crate::coded("internal-invalid-field", "Invalid merged-range cell page")
        })? {
            if pair[0] != anchor
                && (!pair[1]["value"].is_null()
                    || !pair[1]["formula"].is_null()
                    || pair[1]
                        .get("formulaType")
                        .is_some_and(|kind| !kind.is_null())
                    || !pair[1]["rawValue"].is_null()
                    || !pair[1]["formulaOutput"].is_null()
                    || !pair[1]["spill"].is_null())
            {
                return Err(crate::coded(
                    "merge-conflict",
                    format!(
                        "Merge would hide a nonempty cell: {}",
                        pair[0].as_str().unwrap_or("?")
                    ),
                ));
            }
        }
        let next = page["nextCursor"].as_str().map(str::to_owned);
        if next.is_none() {
            return Ok(());
        }
        if next == after {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid merged-range page cursor",
            ));
        }
        after = next;
    }
}
pub(crate) fn apply(root: &Path, input: &Value) -> Result<Value> {
    let commands: Vec<MergeCommand> =
        serde_json::from_value(input["commands"].clone()).map_err(err)?;
    if commands.is_empty() || commands.len() > 10000 {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid merged-range command count",
        ));
    }
    let parent = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing merge version"))?;
    let index: Value =
        serde_json::from_slice(&std::fs::read(local(root, "index.json")?).map_err(err)?)
            .map_err(err)?;
    let mut db = crate::database(root)?;
    let tx = db.transaction().map_err(err)?;
    if !tx
        .query_row(
            "SELECT EXISTS(SELECT 1 FROM versions WHERE id=?1)",
            [parent],
            |row| row.get::<_, bool>(0),
        )
        .map_err(err)?
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown merge version",
        ));
    }
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let workbook = part(&mut zip, "xl/workbook.xml", &mut 0)?;
    let mut names = false;
    walk(&workbook, NS, "workbook", |_, node| {
        if node.name == "definedName" || node.name == "workbookProtection" {
            names = true;
        }
        Ok(())
    })?;
    if names {
        return Err(crate::coded(
            "merge-protected",
            "Merging a workbook with names or protection is not qualified",
        ));
    }
    let mut sheets = BTreeMap::new();
    for command in commands {
        crate::cancel::check()?;
        let (sheet, range) = command.target();
        if !index["sheets"]
            .as_array()
            .is_some_and(|sheets| sheets.iter().any(|entry| entry["path"] == sheet))
        {
            return Err(crate::coded(
                "internal-unknown-reference",
                "Unknown merge worksheet",
            ));
        }
        let range = canonical(range)?;
        let area = bounds(&range)?;
        if area.0 == area.1
            || matches!(command, MergeCommand::MergeCells { .. })
                && u64::from(area.1.0 - area.0.0 + 1) * u64::from(area.1.1 - area.0.1 + 1) > 1048576
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid or oversized merged range",
            ));
        }
        if !sheets.contains_key(sheet) {
            let source = source(root, sheet)?;
            if source.protected {
                return Err(crate::coded(
                    "table-or-protection",
                    "Cells on a sheet that has a table or is protected cannot be merged here",
                ));
            }
            sheets.insert(
                sheet.to_owned(),
                selected_override(&tx, parent, sheet)?.unwrap_or(source.ranges),
            );
        }
        let ranges = sheets.get_mut(sheet).unwrap();
        match command {
            MergeCommand::MergeCells { .. } => {
                if ranges.len() >= 10000
                    || ranges
                        .iter()
                        .any(|other| bounds(other).is_ok_and(|other| overlaps(area, other)))
                {
                    return Err(crate::coded(
                        "merge-limit",
                        "Merge overlaps an existing range or exceeds its count bound",
                    ));
                }
                check_cells(&tx, parent, sheet, &range)?;
                ranges.push(range);
            }
            MergeCommand::UnmergeCells { .. } => {
                let at = ranges
                    .iter()
                    .position(|other| *other == range)
                    .ok_or_else(|| {
                        crate::coded(
                            "merge-bad-input",
                            "Select the exact merged range to unmerge",
                        )
                    })?;
                ranges.remove(at);
            }
        }
    }
    tx.execute_batch("CREATE TABLE IF NOT EXISTS version_merges(version INTEGER NOT NULL,sheet TEXT NOT NULL,ranges TEXT NOT NULL,PRIMARY KEY(version,sheet)) WITHOUT ROWID;").map_err(err)?;
    tx.execute("INSERT INTO versions(parent) VALUES(?1)", [parent])
        .map_err(err)?;
    let version = tx.last_insert_rowid();
    for (sheet, ranges) in sheets {
        tx.execute(
            "INSERT INTO version_merges VALUES(?1,?2,?3)",
            params![version, sheet, serde_json::to_string(&ranges).map_err(err)?],
        )
        .map_err(err)?;
    }
    crate::cancel::check()?;
    if root.join("cancelled").exists() {
        return Err(crate::coded(
            "operation-cancelled",
            "Workbook operation cancelled",
        ));
    }
    tx.commit().map_err(err)?;
    Ok(json!({"version":version}))
}
pub(crate) fn fragment(db: &Connection, version: u64, sheet: &str) -> Result<Vec<u8>> {
    let ranges = selected_override(db, version, sheet)?
        .ok_or_else(|| crate::coded("merge-bad-input", "No accepted merged-range overlay"))?;
    if ranges.is_empty() {
        return Ok(Vec::new());
    }
    let mut xml = format!(
        "<mergeCells xmlns=\"{}\" count=\"{}\">",
        String::from_utf8_lossy(NS),
        ranges.len()
    );
    for range in ranges {
        let range = canonical(&range)?;
        xml.push_str(&format!("<mergeCell ref=\"{range}\"/>"));
    }
    xml.push_str("</mergeCells>");
    Ok(xml.into_bytes())
}
pub(crate) fn retain(db: &Connection, keep: &BTreeSet<u64>) -> Result<()> {
    if !present(db)? {
        return Ok(());
    }
    for version in keep {
        db.execute("WITH RECURSIVE a(id,depth) AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL), ranked AS(SELECT m.*,row_number() OVER(PARTITION BY m.sheet ORDER BY a.depth) priority FROM a JOIN version_merges m ON m.version=a.id) INSERT OR REPLACE INTO version_merges SELECT ?1,sheet,ranges FROM ranked WHERE priority=1",[version]).map_err(err)?;
    }
    Ok(())
}
pub(crate) fn prune(db: &Connection) -> Result<()> {
    if present(db)? {
        db.execute(
            "DELETE FROM version_merges WHERE version NOT IN(SELECT id FROM versions)",
            [],
        )
        .map_err(err)?;
    }
    Ok(())
}
