use crate::xml::{address, position, spreadsheet_text};
use crate::{Result, database, err, field};
use rusqlite::{Connection, OptionalExtension};
use serde_json::{Value, json};
use std::borrow::Borrow;
use std::path::Path;

fn coordinate_index(db: &Connection) -> Result<bool> {
    db.prepare_cached("SELECT EXISTS(SELECT 1 FROM sqlite_master WHERE type='index' AND name='changes_coordinates')").map_err(crate::err)?.query_row([], |row| row.get(0)).map_err(err)
}
/// One connection/request-local ancestry snapshot. Immutable version parents
/// are never changed during range/patch/bulk/format evaluation.
fn edit_ancestry(db: &Connection, version: u64) -> Result<()> {
    // Create the per-connection tables only when this connection has none.
    // A failed probe covers a new connection and a rolled-back creation alike.
    let known = db
        .prepare_cached("SELECT version FROM range_edit_context")
        .and_then(|mut statement| {
            statement
                .query_row([], |row| row.get::<_, u64>(0))
                .optional()
        })
        .ok();
    if known.is_none() {
        db.execute_batch("CREATE TEMP TABLE IF NOT EXISTS range_edit_context(version INTEGER); CREATE TEMP TABLE IF NOT EXISTS range_edit_ancestry(id INTEGER PRIMARY KEY,depth INTEGER NOT NULL) WITHOUT ROWID;").map_err(err)?;
    }
    let current: Option<u64> = known.flatten();
    if current != Some(version) {
        db.execute_batch("DELETE FROM range_edit_context; DELETE FROM range_edit_ancestry;")
            .map_err(err)?;
        db.execute("WITH RECURSIVE a(id,depth) AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL) INSERT INTO range_edit_ancestry SELECT id,depth FROM a", [version]).map_err(err)?;
        db.execute("INSERT INTO range_edit_context VALUES(?1)", [version])
            .map_err(err)?;
    }
    Ok(())
}

/// Reads a sheet's cells in coordinate order, one page at a time, for callers
/// that visit ascending positions. `get` answers a position without a separate
/// query per cell; a position the stream has already passed must not be asked.
pub(crate) struct OrderedCells<'a> {
    db: &'a Connection,
    version: u64,
    sheet: String,
    formulas: bool,
    after: Option<String>,
    page: std::collections::VecDeque<((u32, u32), Value)>,
    finished: bool,
}
impl<'a> OrderedCells<'a> {
    pub(crate) fn new(db: &'a Connection, version: u64, sheet: &str, formulas: bool) -> Self {
        Self {
            db,
            version,
            sheet: sheet.into(),
            formulas,
            after: None,
            page: Default::default(),
            finished: false,
        }
    }
    pub(crate) fn get(&mut self, at: (u32, u32)) -> Result<Option<Value>> {
        loop {
            while self
                .page
                .front()
                .is_some_and(|(position, _)| *position < at)
            {
                self.page.pop_front();
            }
            if let Some((position, _)) = self.page.front() {
                return Ok(if *position == at {
                    self.page.pop_front().map(|(_, cell)| cell)
                } else {
                    None
                });
            }
            if self.finished {
                return Ok(None);
            }
            let page = read_db(
                self.db,
                &json!({"version":self.version,"sheet":self.sheet,"range":"A1:XFD1048576","after":self.after,"limit":2048,"formulas":self.formulas}),
            )?;
            let cells = page["cells"].as_array().ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing ordered cell page")
            })?;
            for pair in cells {
                let reference = pair[0].as_str().ok_or_else(|| {
                    crate::coded("internal-invalid-field", "Invalid ordered cell coordinate")
                })?;
                self.page.push_back((position(reference)?, pair[1].clone()));
            }
            match cells.last().and_then(|pair| pair[0].as_str()) {
                Some(last) if page["nextCursor"].is_string() => self.after = Some(last.into()),
                _ => self.finished = true,
            }
        }
    }
}

pub fn read(root: &Path, input: &Value) -> Result<Value> {
    let db = database(root)?;
    if input["calculationInput"] == true {
        read_db(&db, input)
    } else {
        read_public_db(&db, input)
    }
}
pub(crate) fn read_public_db(db: &Connection, input: &Value) -> Result<Value> {
    let mut page = read_db(db, input)?;
    let version = input["version"].as_u64().unwrap_or(0);
    for pair in page["cells"]
        .as_array_mut()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid cell page"))?
    {
        let reference = pair[0]
            .as_str()
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid cell address"))?
            .to_owned();
        crate::calculation_policy::mask(
            db,
            version,
            field(input, "sheet")?,
            &reference,
            &mut pair[1],
        )?;
    }
    Ok(page)
}

pub(crate) fn read_db(db: &Connection, input: &Value) -> Result<Value> {
    let range = field(input, "range")?;
    let mut bounds = range.split(':');
    let first = position(
        bounds
            .next()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing range"))?,
    )?;
    let last = bounds.next().map(position).transpose()?.unwrap_or(first);
    if bounds.next().is_some() || first.0 > last.0 || first.1 > last.1 {
        return Err(crate::coded("internal-invalid-field", "Invalid range"));
    }
    // Pagination is a transport bound, never a workbook or operation size limit.
    let limit = match input.get("limit") {
        None => 2048,
        Some(value) => value
            .as_u64()
            .filter(|value| *value > 0)
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid range page limit"))?
            .min(10000),
    };
    let offset = input["offset"].as_u64().unwrap_or(0);
    let version = input["version"].as_u64().unwrap_or(0);
    if db
        .prepare_cached("SELECT count(*) FROM versions WHERE id=?1")
        .map_err(err)?
        .query_row([version], |r| r.get::<_, u32>(0))
        .map_err(err)?
        != 1
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown workbook version",
        ));
    }
    let after = input["after"].as_str().map(position).transpose()?;
    crate::shared::prepare(db, version)?;
    let (after_row, after_col) = after
        .map(|(r, c)| (i64::from(r), i64::from(c)))
        .unwrap_or((-1, -1));
    let formulas = input["formulas"].as_bool().unwrap_or(false);
    let base_filter = if formulas {
        "formula IS NOT NULL AND NOT EXISTS (SELECT 1 FROM overlay o WHERE o.r=cells.r AND o.c=cells.c AND json_extract(o.data,'$.formula') IS NULL AND json_extract(o.data,'$.formulaType') IS NULL)"
    } else {
        "1"
    };
    let has_changes = !coordinate_index(db)? || db.prepare_cached("SELECT EXISTS(SELECT 1 FROM changes WHERE sheet=?1 AND r BETWEEN ?2 AND ?3 AND c BETWEEN ?4 AND ?5)").map_err(crate::err)?.query_row(rusqlite::params![field(input,"sheet")?,first.0,last.0,first.1,last.1],|row|row.get::<_,bool>(0)).map_err(err)?;
    let ancestry = if has_changes {
        edit_ancestry(db, version)?;
        "SELECT id,depth FROM range_edit_ancestry"
    } else {
        "SELECT ?8,0 WHERE 0"
    };
    // Stored cells alone fill a page by the row of their last entry, so edits
    // below that row cannot be on this page and need not be ranked. A formula
    // page filters stored cells through the edits, so it keeps the full range.
    let last_row = if has_changes && !formulas {
        db.prepare_cached("SELECT r FROM cells WHERE sheet=?1 AND r BETWEEN max(?2,?6) AND ?3 AND c BETWEEN ?4 AND ?5 AND (r,c)>(?6,?7) ORDER BY r,c LIMIT 1 OFFSET ?8")
            .map_err(err)?
            .query_row(rusqlite::params![field(input, "sheet")?, first.0, last.0, first.1, last.1, after_row, after_col, limit + offset - 1], |row| row.get::<_, u32>(0))
            .optional()
            .map_err(err)?
            .unwrap_or(last.0)
    } else {
        last.0
    };
    let sql = format!(
        "WITH ancestry(id,depth) AS ({ancestry}), ranked AS (SELECT p.*,row_number() OVER (PARTITION BY p.sheet,p.r,p.c ORDER BY a.depth) priority FROM changes p JOIN ancestry a ON p.version=a.id WHERE p.sheet=?1 AND p.r BETWEEN max(?2,?10) AND ?12 AND p.c BETWEEN ?4 AND ?5 AND (p.r,p.c)>(?10,?11)), overlay AS (SELECT * FROM ranked WHERE priority=1), base AS (SELECT r,c FROM cells WHERE sheet=?1 AND r BETWEEN max(?2,?10) AND ?3 AND c BETWEEN ?4 AND ?5 AND (r,c)>(?10,?11) AND ({base_filter}) ORDER BY r,c LIMIT ?6+?7), coordinates AS (SELECT r,c FROM base UNION SELECT r,c FROM overlay WHERE ?9=0 OR json_extract(data,'$.formula') IS NOT NULL OR json_extract(data,'$.formulaType') IS NOT NULL) SELECT k.r,k.c,c.kind,c.raw,c.inline,c.formula,c.formula_attrs,c.style,s.value,o.data FROM coordinates k LEFT JOIN cells c ON c.sheet=?1 AND c.r=k.r AND c.c=k.c LEFT JOIN strings s ON c.kind='s' AND s.id=CAST(c.raw AS INTEGER) LEFT JOIN overlay o ON o.r=k.r AND o.c=k.c WHERE (?9=0 OR CASE WHEN o.data IS NOT NULL THEN (json_extract(o.data,'$.formula') IS NOT NULL OR json_extract(o.data,'$.formulaType') IS NOT NULL) ELSE c.formula IS NOT NULL END) ORDER BY k.r,k.c LIMIT ?6 OFFSET ?7"
    );
    let mut statement = db.prepare_cached(&sql).map_err(err)?;
    let mut rows = statement
        .query(rusqlite::params![
            field(input, "sheet")?,
            first.0,
            last.0,
            first.1,
            last.1,
            limit,
            offset,
            version,
            u32::from(formulas),
            after_row,
            after_col,
            last_row
        ])
        .map_err(err)?;
    let mut cells = Vec::new();
    let budget = input
        .get("maxBytes")
        .map(|value| {
            value
                .as_u64()
                .filter(|value| (1024..=64 * 1024 * 1024).contains(value))
                .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid range byte budget"))
        })
        .transpose()?;
    let mut bytes = 256u64;
    let mut byte_page = false;
    while let Some(row) = rows.next().map_err(err)? {
        crate::cancel::check()?;
        let (reference, cell) = decode_cell(row)?;
        let cell = crate::shared::resolve(db, field(input, "sheet")?, &reference, cell)?;
        let cell = json!([reference, cell]);
        if let Some(budget) = budget {
            let weight = serialized_bytes(&cell)?.saturating_add(1);
            if bytes.saturating_add(weight) > budget {
                if cells.is_empty() {
                    return Err(crate::coded(
                        "workbook-limit",
                        "A cell exceeds the range byte budget",
                    ));
                }
                byte_page = true;
                break;
            }
            bytes += weight;
        }
        cells.push(cell);
    }
    Ok(
        json!({"cells":cells,"offset":offset,"nextCursor":if byte_page || cells.len() as u64==limit {cells.last().map(|c|c[0].clone())}else{None},"nextOffset":if byte_page || cells.len() as u64==limit {Some(offset+cells.len() as u64)} else {None}}),
    )
}
fn serialized_bytes(value: &Value) -> Result<u64> {
    struct Counter(u64);
    impl std::io::Write for Counter {
        fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
            self.0 = self.0.saturating_add(bytes.len() as u64);
            Ok(bytes.len())
        }
        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }
    let mut count = Counter(0);
    serde_json::to_writer(&mut count, value).map_err(err)?;
    Ok(count.0)
}

/// Shared decoding for range transport and storage-side calculation snapshots.
pub(crate) const SOURCE_FIELDS: [&str; 5] =
    ["value", "dataType", "rawValue", "cachedValue", "styleId"];

pub(crate) fn decode_cell(row: &rusqlite::Row<'_>) -> Result<(String, Value)> {
    decode_cell_impl(row, false)
}

/// Formula evaluation consumes explicit formula identity and calculated values.
/// Authored-value, structural and export readers must use the complete decoder.
pub(crate) fn decode_evaluation_cell(row: &rusqlite::Row<'_>) -> Result<(String, Value)> {
    decode_cell_impl(row, true)
}

fn decode_cell_impl(row: &rusqlite::Row<'_>, formula_only: bool) -> Result<(String, Value)> {
    let r: u32 = row.get(0).map_err(err)?;
    let c: u32 = row.get(1).map_err(err)?;
    if let Some(data) = row.get::<_, Option<String>>(9).map_err(err)? {
        let mut data: Value = serde_json::from_str(&data).map_err(err)?;
        if let Some(marker) = data.get("$sourceFields") {
            if marker != &Value::Bool(true)
                || SOURCE_FIELDS.iter().any(|key| data.get(*key).is_some())
            {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid source-field calculation record",
                ));
            }
            let formula = data.get("formulaType").is_none() && data["formula"].is_string();
            let shared = data["formulaType"] == "shared";
            if formula_only
                && (formula || shared)
                && data["scenarioResult"] != true
                && !matches!(
                    row.get_ref(2).map_err(err)?,
                    rusqlite::types::ValueRef::Null
                )
            {
                data.as_object_mut().unwrap().remove("$sourceFields");
                return Ok((address(r, c), data));
            }
            let mut source = decode_source_cell(row, true)?;
            let map = data.as_object_mut().ok_or_else(|| {
                crate::coded(
                    "internal-invalid-field",
                    "Invalid source-field calculation record",
                )
            })?;
            map.remove("$sourceFields");
            for key in SOURCE_FIELDS {
                map.insert(
                    key.into(),
                    source.as_object_mut().unwrap().remove(key).unwrap(),
                );
            }
        }
        return Ok((address(r, c), data));
    }
    Ok((address(r, c), decode_source_cell(row, false)?))
}

fn decode_source_cell(row: &rusqlite::Row<'_>, fields_only: bool) -> Result<Value> {
    let kind: String = row.get(2).map_err(err)?;
    let raw: Option<String> = row.get(3).map_err(err)?;
    let inline: String = row.get(4).map_err(err)?;
    let formula_present = !matches!(
        row.get_ref(5).map_err(err)?,
        rusqlite::types::ValueRef::Null
    );
    let style: u32 = row.get(7).map_err(err)?;
    let text = raw.as_deref().unwrap_or("");
    let (value, data_type) = match kind.as_str() {
        "s" => {
            if text.parse::<u64>().is_err() {
                return Err(crate::coded(
                    "workbook-malformed",
                    "Invalid shared-string reference",
                ));
            }
            (
                json!(row.get::<_, Option<String>>(8).map_err(err)?.ok_or_else(
                    || crate::coded("workbook-malformed", "Invalid shared-string reference")
                )?),
                "text",
            )
        }
        "inlineStr" => (json!(inline), "text"),
        "str" => (json!(spreadsheet_text(text)?), "text"),
        "b" => (json!(text == "1"), "boolean"),
        "e" => (json!(text), "error"),
        "n" if text.is_empty() => (Value::Null, "blank"),
        "n" => {
            let number = text.parse::<f64>().map_err(err)?;
            (
                if number.is_finite() {
                    json!(number)
                } else {
                    json!({"$workbookNumber":if number.is_nan(){"NaN"}else if number.is_sign_negative(){"-Infinity"}else{"Infinity"}})
                },
                "number",
            )
        }
        _ => (json!(text), "text"),
    };
    let mut cell = json!({"value":value,"dataType":data_type,"rawValue":if kind=="s" || kind=="inlineStr" {None} else {raw.clone()},"cachedValue":if formula_present&&raw.is_some() {value.clone()} else {Value::Null},"styleId":style});
    if fields_only {
        return Ok(cell);
    }
    let formula: Option<String> = row.get(5).map_err(err)?;
    let attrs: Option<String> = row.get(6).map_err(err)?;
    if let Some(formula) = formula
        && !formula.is_empty()
    {
        cell["formula"] = json!(formula);
    }
    if let Some(attrs) = attrs {
        let attrs: Value = serde_json::from_str(&attrs).map_err(err)?;
        if let Some(kind) = attrs.get("t") {
            cell["formulaType"] = kind.clone();
            cell["sharedFormulaId"] = attrs["si"].clone();
            cell["formulaReference"] = attrs["ref"].clone();
        }
    }
    if kind == "e" {
        cell["error"] = json!(text);
    }
    Ok(cell)
}

/// Exact immutable ordinary-formula identity, with independent source cache
/// fields. Formatting and calculation decorations are deliberately excluded.
/// A shared/array/data-table formula cannot enter the ordinary source path.
pub(crate) fn source_formula_identity(row: &rusqlite::Row<'_>, cell: &Value) -> Result<bool> {
    let formula: Option<String> = row.get(5).map_err(err)?;
    if formula.as_deref().is_none_or(str::is_empty)
        || cell.get("formulaType").is_some()
        || formula.as_deref() != cell["formula"].as_str()
    {
        return Ok(false);
    }
    let source = decode_source_cell(row, false)?;
    if source.get("formulaType").is_some() {
        return Ok(false);
    }
    fn number(value: &Value) -> Option<f64> {
        value.as_f64().or_else(|| {
            value
                .as_object()
                .filter(|fields| fields.len() == 1)
                .and_then(|fields| fields.get("$workbookNumber"))
                .and_then(Value::as_str)
                .filter(|marker| *marker == "-0")
                .map(|_| -0.0)
        })
    }
    for key in [
        "formula",
        "formulaReference",
        "sharedFormulaId",
        "value",
        "dataType",
        "rawValue",
        "cachedValue",
        "error",
    ] {
        let a = &source[key];
        let b = &cell[key];
        let equal = if let (Some(a), Some(b)) = (number(a), number(b)) {
            a.to_bits() == b.to_bits()
        } else {
            a == b
        };
        if !equal {
            return Ok(false);
        }
    }
    Ok(true)
}
pub fn patch(root: &Path, input: &Value) -> Result<Value> {
    let parent = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing parent version"))?;
    let changes = input["changes"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing cell patches"))?;
    let version = patch_changes(root, parent, changes.iter().map(Ok), true)?;
    Ok(json!({"version":version}))
}

/// Consume one patch at a time while its new version remains private. A failed
/// producer, validation, write or cancellation drops the whole transaction.
pub(crate) fn patch_changes<I, C>(
    root: &Path,
    parent: u64,
    changes: I,
    create_empty_version: bool,
) -> Result<u64>
where
    I: IntoIterator<Item = Result<C>>,
    C: Borrow<Value>,
{
    let mut db = database(root)?;
    let transaction = db.transaction().map_err(err)?;
    // Mutable legacy roots acquire the index atomically; pure readers never
    // change persistent schema and retain the correct legacy fallback.
    transaction
        .execute_batch(
            "CREATE INDEX IF NOT EXISTS changes_coordinates ON changes(sheet,r,c,version);",
        )
        .map_err(err)?;
    if transaction
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
    let mut changes = changes.into_iter();
    crate::cancel::check()?;
    let mut next = changes.next().transpose()?;
    crate::cancel::check()?;
    if next.is_none() && !create_empty_version {
        if root.join("cancelled").exists() {
            return Err(crate::coded(
                "operation-cancelled",
                "Workbook operation cancelled",
            ));
        }
        return Ok(parent);
    }
    transaction
        .execute("INSERT INTO versions(parent) VALUES(?1)", [parent])
        .map_err(err)?;
    let version = transaction.last_insert_rowid();
    let has_spills = crate::spill::present(&transaction)?;
    // Calculation-only patches never need index metadata. Load it lazily if
    // an authored content edit arrives, including after earlier streamed rows.
    let mut special_sheets: Option<Vec<String>> = None;
    let mut merged_sheets = std::collections::BTreeMap::new();
    while let Some(change) = next {
        crate::cancel::check()?;
        let change = change.borrow();
        let (row, col) = position(field(change, "cell")?)?;
        let cell = change
            .get("value")
            .filter(|v| v.is_object())
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid cell patch"))?;
        let sheet = field(change, "sheet")?;
        if crate::calculation_policy::preserving(&transaction, parent)? {
            let edits = change["edits"].as_array().ok_or_else(|| {
                crate::coded(
                    "workbook-unsupported",
                    "Preserved patch requires explicit edit intent",
                )
            })?;
            if edits.iter().any(|e| e.get("formula").is_some()) {
                return Err(crate::coded(
                    "formula-protected",
                    "Formula changes remain protected while source calculation is preserved",
                ));
            }
            crate::preservation::region(
                root,
                &transaction,
                sheet,
                (row, col),
                (row, col),
                edits.iter().any(|e| e.get("value").is_some()),
            )?;
        }
        if has_spills {
            crate::spill::guard(
                &transaction,
                parent,
                sheet,
                &crate::spill::Geometry {
                    first: (row, col),
                    last: (row, col),
                },
            )?;
        }
        let mut edit = inherited_edit(&transaction, parent, sheet, row, col)?;
        let authored = change["edits"].as_array().is_none_or(|edits| {
            edits
                .iter()
                .any(|edit| edit.get("value").is_some() || edit.get("formula").is_some())
        });
        if authored {
            if !merged_sheets.contains_key(sheet) {
                let ranges = crate::merges::selected(root, &transaction, parent, sheet)?;
                merged_sheets.insert(sheet.to_owned(), ranges);
            }
            for range in &merged_sheets[sheet] {
                let (first, last) = crate::bulk::bounds(range)?;
                if (row, col) != first
                    && row >= first.0
                    && row <= last.0
                    && col >= first.1
                    && col <= last.1
                {
                    return Err(crate::coded(
                        "merge-unsupported",
                        "A merged follower cannot be edited independently",
                    ));
                }
            }
        }
        if let Some(edits) = change["edits"].as_array() {
            let content_edit = edits
                .iter()
                .any(|e| e.get("value").is_some() || e.get("formula").is_some());
            if content_edit && special_sheets.is_none() {
                let index: Value =
                    serde_json::from_slice(&std::fs::read(root.join("index.json")).map_err(err)?)
                        .map_err(err)?;
                special_sheets = Some(
                    index["sheets"]
                        .as_array()
                        .ok_or_else(|| {
                            crate::coded("internal-missing-field", "Missing sheet metadata")
                        })?
                        .iter()
                        .filter(|s| s["specialFormulaCount"].as_u64().unwrap_or(0) > 0)
                        .map(|s| s["path"].as_str().unwrap_or("").to_owned())
                        .collect(),
                );
            }
            if content_edit
                && special_sheets
                    .as_ref()
                    .is_some_and(|sheets| sheets.iter().any(|s| s == sheet))
            {
                crate::bulk::special_groups(
                    &transaction,
                    parent,
                    sheet,
                    ((row, col), (row, col)),
                    false,
                )?;
            }
            for delta in edits {
                merge_edit(&mut edit, delta)?;
            }
        } else if cell.get("formula").is_some() || cell.get("formulaType").is_some() {
            merge_edit(&mut edit, &json!({"formula":cell["formula"]}))?;
        } else {
            merge_edit(&mut edit, &json!({"value":cell["value"]}))?;
        }
        if edit["valueType"] == "error"
            && (cell["dataType"] != "error"
                || cell["value"] != edit["value"]
                || cell.get("formula").is_some())
        {
            return Err(crate::coded(
                "internal-mismatch",
                "Declared error edit differs from its typed cell",
            ));
        }
        store_change(&transaction, version, sheet, row, col, cell, &edit)?;
        crate::cancel::check()?;
        next = changes.next().transpose()?;
    }
    crate::cancel::check()?;
    // The regular token poll is throttled. A marker written by the producer
    // at its final yield must still prevent publication of the private version.
    if root.join("cancelled").exists() {
        return Err(crate::coded(
            "operation-cancelled",
            "Workbook operation cancelled",
        ));
    }
    let accepted = u64::try_from(version).map_err(err)?;
    transaction.commit().map_err(err)?;
    Ok(accepted)
}

/// Each overlay carries its cumulative export intent. Formatting never forces a
/// precise source number or rich string through a decoded literal rewrite.
pub(crate) fn inherited_edit(
    db: &Connection,
    version: u64,
    sheet: &str,
    row: u32,
    col: u32,
) -> Result<Value> {
    if coordinate_index(db)?
        && !db
            .prepare_cached("SELECT EXISTS(SELECT 1 FROM changes WHERE sheet=?1 AND r=?2 AND c=?3)")
            .map_err(err)?
            .query_row(rusqlite::params![sheet, row, col], |row| {
                row.get::<_, bool>(0)
            })
            .map_err(err)?
    {
        return Ok(json!({}));
    }
    edit_ancestry(db, version)?;
    let data: Option<String> = db.prepare_cached("SELECT p.edit FROM changes p JOIN range_edit_ancestry a ON p.version=a.id WHERE p.sheet=?1 AND p.r=?2 AND p.c=?3 ORDER BY a.depth LIMIT 1").map_err(err)?
        .query_row(rusqlite::params![sheet,row,col], |r| r.get(0)).optional().map_err(err)?;
    data.map(|s| serde_json::from_str(&s).map_err(err))
        .unwrap_or_else(|| Ok(json!({})))
}

#[cfg(test)]
mod ancestry_tests {
    use super::*;
    use rusqlite::params;

    fn fixture(index: bool) -> Connection {
        let db = Connection::open_in_memory().unwrap();
        db.execute_batch("CREATE TABLE versions(id INTEGER PRIMARY KEY,parent INTEGER);INSERT INTO versions VALUES(0,NULL),(1,0),(2,1),(3,1);CREATE TABLE changes(version INTEGER,sheet TEXT,r INTEGER,c INTEGER,data TEXT,edit TEXT,PRIMARY KEY(version,sheet,r,c)) WITHOUT ROWID;CREATE TABLE strings(id INTEGER PRIMARY KEY,value TEXT);CREATE TABLE value_features(sheet TEXT,r INTEGER,c INTEGER,reason TEXT,mask INTEGER DEFAULT 1,PRIMARY KEY(sheet,r,c));CREATE TABLE shared_value_features(id INTEGER PRIMARY KEY,reason TEXT);CREATE TABLE cells(sheet TEXT,r INTEGER,c INTEGER,kind TEXT,raw TEXT,inline TEXT,formula TEXT,formula_attrs TEXT,style INTEGER,PRIMARY KEY(sheet,r,c)) WITHOUT ROWID;").unwrap();
        if index {
            db.execute_batch("CREATE INDEX changes_coordinates ON changes(sheet,r,c,version)")
                .unwrap();
        }
        crate::shared::schema(&db).unwrap();
        db.execute(
            "INSERT INTO cells VALUES('Sheet',0,0,'n','9007199254740993','',NULL,NULL,7)",
            [],
        )
        .unwrap();
        db.execute(
            "INSERT INTO cells VALUES('Sheet',0,1,'n','-0','',NULL,NULL,8)",
            [],
        )
        .unwrap();
        db.execute(
            "INSERT INTO cells VALUES('Sheet',1,0,'n','913','', 'A1+1',NULL,9)",
            [],
        )
        .unwrap();
        for (version, row, col, data, edit) in [
            (
                1,
                0,
                0,
                json!({"value":12,"dataType":"number","rawValue":"12","cachedValue":null,"styleId":7}),
                json!({"value":12,"fontColor":"#123456"}),
            ),
            (
                2,
                0,
                0,
                json!({"value":14,"dataType":"number","rawValue":"14","cachedValue":null,"styleId":7}),
                json!({"value":14,"fontColor":"#123456","bold":true}),
            ),
            (
                3,
                0,
                0,
                json!({"value":16,"dataType":"number","rawValue":"16","cachedValue":null,"styleId":7}),
                json!({"value":16,"fontColor":"#123456","italic":true}),
            ),
            (
                3,
                2,
                0,
                json!({"value":"sibling-only","dataType":"text","styleId":5}),
                json!({"value":"sibling-only","wrapText":true}),
            ),
            (
                2,
                1,
                0,
                json!({"value":null,"dataType":"blank","rawValue":null,"cachedValue":null,"styleId":9}),
                json!({"value":null}),
            ),
        ] {
            db.execute(
                "INSERT INTO changes VALUES(?1,'Sheet',?2,?3,?4,?5)",
                params![version, row, col, data.to_string(), edit.to_string()],
            )
            .unwrap();
        }
        db
    }
    fn main_schema(db: &Connection) -> Vec<(String, String)> {
        db.prepare_cached("SELECT name,sql FROM sqlite_master WHERE sql IS NOT NULL ORDER BY name")
            .unwrap()
            .query_map([], |r| Ok((r.get(0)?, r.get(1)?)))
            .unwrap()
            .collect::<std::result::Result<_, _>>()
            .unwrap()
    }
    fn patch_fixture() -> tempfile::TempDir {
        let root = tempfile::tempdir().unwrap();
        let db = fixture(false);
        crate::used_range::schema(&db).unwrap();
        db.execute(
            "INSERT INTO used_range_history VALUES(0,'Sheet',0,0,1,1)",
            [],
        )
        .unwrap();
        db.execute(
            "INSERT INTO cells VALUES('Sheet',0,2,'n','9007199254740993','',NULL,NULL,10)",
            [],
        )
        .unwrap();
        db.backup("main", root.path().join("cells.sqlite"), None)
            .unwrap();
        std::fs::write(
            root.path().join("index.json"),
            json!({"sheets":[{"path":"Sheet","specialFormulaCount":0}]}).to_string(),
        )
        .unwrap();
        use std::io::Write;
        let mut zip =
            zip::ZipWriter::new(std::fs::File::create(root.path().join("metadata.xlsx")).unwrap());
        zip.start_file("Sheet", zip::write::SimpleFileOptions::default())
            .unwrap();
        zip.write_all(b"<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/></worksheet>").unwrap();
        zip.finish().unwrap();
        root
    }
    fn calculation_patch(reference: &str, value: Value) -> Value {
        json!({"sheet":"Sheet","cell":reference,"value":value,"edits":[]})
    }
    fn new_calculation_patch(row: u32) -> Value {
        calculation_patch(
            &address(row, 0),
            json!({"value":null,"dataType":"blank","formula":"1+1","rawValue":null,"cachedValue":null,"styleId":0,"calculation":{"status":"calculated","value":2,"revision":"test"}}),
        )
    }
    fn persistent_state(root: &Path) -> (Vec<(String, String)>, [u64; 3]) {
        let db = Connection::open(root.join("cells.sqlite")).unwrap();
        let counts = ["versions", "changes", "used_range_history"].map(|table| {
            db.query_row(&format!("SELECT count(*) FROM {table}"), [], |row| {
                row.get(0)
            })
            .unwrap()
        });
        (main_schema(&db), counts)
    }
    #[test]
    fn streamed_calculation_patches_publish_one_version_and_preserve_source_and_edit_fields() {
        let root = patch_fixture();
        let before = persistent_state(root.path());
        let db = Connection::open(root.path().join("cells.sqlite")).unwrap();
        let _scope = crate::shared::begin();
        let page = read_db(&db, &json!({"version":1,"sheet":"Sheet","range":"A1:C2"})).unwrap();
        let originals = page["cells"].as_array().unwrap().clone();
        drop(db);
        let preserved = originals.iter().map(|pair| {
            let mut value = pair[1].clone();
            value["calculation"] = json!({"status":"calculated","value":2,"revision":"stream"});
            Ok(calculation_patch(pair[0].as_str().unwrap(), value))
        });
        let patches = preserved.chain((3..303).map(|row| {
            // A separate reader cannot observe an intermediate version, even
            // after crossing the old compatibility batch boundary.
            let visible = Connection::open(root.path().join("cells.sqlite")).unwrap();
            assert_eq!(
                visible
                    .prepare_cached("SELECT count(*) FROM versions")
                    .unwrap()
                    .query_row([], |row| row.get::<_, u64>(0))
                    .unwrap(),
                before.1[0]
            );
            Ok(new_calculation_patch(row))
        }));
        let version = crate::apply_calculation_patches(root.path(), 1, patches).unwrap();
        let db = Connection::open(root.path().join("cells.sqlite")).unwrap();
        assert_eq!(
            db.prepare_cached("SELECT parent FROM versions WHERE id=?1")
                .unwrap()
                .query_row([version], |row| row.get::<_, u64>(0))
                .unwrap(),
            1
        );
        assert_eq!(
            persistent_state(root.path()).1,
            [before.1[0] + 1, before.1[1] + 304, before.1[2] + 1]
        );
        for pair in originals {
            let page = read_db(
                &db,
                &json!({"version":version,"sheet":"Sheet","range":pair[0]}),
            )
            .unwrap();
            let stored = &page["cells"][0][1];
            for (key, value) in pair[1].as_object().unwrap() {
                assert_eq!(stored.get(key), Some(value), "{} {key}", pair[0]);
            }
        }
        assert_eq!(
            inherited_edit(&db, version, "Sheet", 0, 0).unwrap(),
            json!({"value":12,"fontColor":"#123456"})
        );
        assert_eq!(
            inherited_edit(&db, version, "Sheet", 1, 0).unwrap(),
            json!({})
        );
        assert_eq!(
            read_db(
                &db,
                &json!({"version":version,"sheet":"Sheet","range":"C1"})
            )
            .unwrap()["cells"][0][1]["rawValue"],
            "9007199254740993"
        );
        assert_eq!(
            read_db(
                &db,
                &json!({"version":version,"sheet":"Sheet","range":"B1"})
            )
            .unwrap()["cells"][0][1]["value"]
                .as_f64()
                .unwrap()
                .to_bits(),
            (-0.0f64).to_bits()
        );
        assert!(!root.path().join("native-active.json").exists());
    }
    #[test]
    fn streamed_late_producer_or_validation_errors_roll_back_every_patch_and_index() {
        for failure in [
            Err("Late calculator provenance failure".into()),
            Ok(calculation_patch("XFE1", json!({"value":2}))),
            Ok(calculation_patch(
                "A400",
                json!({"value":2,"$sourceFields":true}),
            )),
            Ok(calculation_patch("A400", json!({"value":2,"spill":{}}))),
            Ok(new_calculation_patch(3)),
        ] {
            let root = patch_fixture();
            let before = persistent_state(root.path());
            let patches = (3..260)
                .map(|row| Ok(new_calculation_patch(row)))
                .chain([failure]);
            assert!(crate::apply_calculation_patches(root.path(), 2, patches).is_err());
            assert_eq!(persistent_state(root.path()), before);
            assert!(!root.path().join("native-active.json").exists());
        }
    }
    #[test]
    fn streamed_patches_retain_source_reference_values_and_inherited_formula_intent() {
        let root = patch_fixture();
        let db = Connection::open(root.path().join("cells.sqlite")).unwrap();
        db.execute("INSERT INTO versions VALUES(4,1)", []).unwrap();
        db.execute("INSERT INTO changes VALUES(4,'Sheet',1,0,?1,?2)", params![
            json!({"$sourceFields":true,"formula":"A1+1","calculation":{"status":"calculated","value":13,"revision":"old"}}).to_string(),
            json!({"fontColor":"#abcdef"}).to_string(),
        ]).unwrap();
        let mut original = read_db(&db, &json!({"version":4,"sheet":"Sheet","range":"A2"}))
            .unwrap()["cells"][0][1]
            .clone();
        assert!(original.get("$sourceFields").is_none());
        assert_eq!(original["rawValue"], "913");
        let source = SOURCE_FIELDS.map(|key| original[key].clone());
        original["calculation"] = json!({"status":"calculated","value":14,"revision":"new"});
        drop(db);
        let version = crate::apply_calculation_patches(
            root.path(),
            4,
            [Ok(calculation_patch("A2", original))],
        )
        .unwrap();
        let db = Connection::open(root.path().join("cells.sqlite")).unwrap();
        let stored = read_db(
            &db,
            &json!({"version":version,"sheet":"Sheet","range":"A2"}),
        )
        .unwrap()["cells"][0][1]
            .clone();
        assert_eq!(SOURCE_FIELDS.map(|key| stored[key].clone()), source);
        assert_eq!(stored["formula"], "A1+1");
        assert_eq!(stored["calculation"]["revision"], "new");
        assert_eq!(
            inherited_edit(&db, version, "Sheet", 1, 0).unwrap(),
            json!({"fontColor":"#abcdef"})
        );
        assert_eq!(
            read_db(&db, &json!({"version":4,"sheet":"Sheet","range":"A2"})).unwrap()["cells"][0]
                [1]["calculation"]["revision"],
            "old"
        );
    }
    #[test]
    fn streamed_patches_keep_special_formula_and_spill_guards_atomic() {
        for spill in [false, true] {
            let root = patch_fixture();
            let mut db = Connection::open(root.path().join("cells.sqlite")).unwrap();
            if spill {
                crate::spill::schema(&db).unwrap();
                let group = crate::spill::Group {
                    version: 1,
                    sheet: "Sheet".into(),
                    master: (4, 3),
                    geometry: crate::spill::Geometry::from_reference("D5:E5").unwrap(),
                    formula: "SEQUENCE(1,2)".into(),
                    proof: json!({"constructor":"sequence-integer-v1","arguments":[1,2],"status":"calculated"}),
                };
                let tx = db.transaction().unwrap();
                for (index, at) in group.geometry.positions().enumerate() {
                    let value = (index + 1) as f64;
                    let mut cell = json!({"value":value,"dataType":"number","rawValue":value.to_string(),"spill":group.cell_metadata(at).unwrap(),"calculation":{"status":"calculated","value":value}});
                    if at == group.master {
                        cell["formula"] = json!(group.formula);
                        cell["formulaType"] = json!("array");
                        cell["formulaReference"] = json!(group.geometry.reference());
                    }
                    store_calculated_change(&tx, 1, "Sheet", at.0, at.1, &cell, &json!({}))
                        .unwrap();
                }
                crate::spill::record(&tx, &group, true).unwrap();
                tx.commit().unwrap();
            } else {
                db.execute(
                    "INSERT INTO cells VALUES('Sheet',4,3,'n','1','','1+1',?1,0)",
                    [json!({"t":"array","ref":"D5:E5"}).to_string()],
                )
                .unwrap();
                std::fs::write(
                    root.path().join("index.json"),
                    json!({"sheets":[{"path":"Sheet","specialFormulaCount":1}]}).to_string(),
                )
                .unwrap();
            }
            drop(db);
            let before = persistent_state(root.path());
            let forbidden = json!({"sheet":"Sheet","cell":"E5","value":{"value":9,"dataType":"number"},"edits":[{"value":9}]});
            let patches = (3..260)
                .map(|row| Ok(new_calculation_patch(row)))
                .chain([Ok(forbidden)]);
            let reason = crate::apply_calculation_patches(root.path(), 1, patches).unwrap_err();
            assert!(
                reason.contains(if spill {
                    "spill follower"
                } else {
                    "complete group"
                }),
                "{reason}"
            );
            assert_eq!(persistent_state(root.path()), before);
        }
    }
    #[test]
    fn cancellation_on_last_iterator_pull_rolls_back_without_publishing_a_version() {
        let root = patch_fixture();
        let before = persistent_state(root.path());
        let mut row = 3;
        let patches = std::iter::from_fn(|| {
            if row == 260 {
                std::fs::write(root.path().join("cancelled"), b"cancel").unwrap();
                return None;
            }
            let patch = new_calculation_patch(row);
            row += 1;
            Some(Ok(patch))
        });
        assert_eq!(
            crate::split_code(
                &crate::apply_calculation_patches(root.path(), 2, patches).unwrap_err()
            )
            .0,
            Some("operation-cancelled")
        );
        assert_eq!(persistent_state(root.path()), before);
        assert!(!root.path().join("native-active.json").exists());
    }
    #[test]
    fn streamed_database_quota_failure_rolls_back_and_releases_the_writer() {
        let root = patch_fixture();
        std::fs::write(
            root.path().join("owner.json"),
            json!({"maxSessionBytes":1024*1024}).to_string(),
        )
        .unwrap();
        let before = persistent_state(root.path());
        let too_large = calculation_patch(
            "A400",
            json!({"value":"x".repeat(2*1024*1024),"dataType":"text"}),
        );
        let patches = (3..260)
            .map(|row| Ok(new_calculation_patch(row)))
            .chain([Ok(too_large)]);
        let reason = crate::apply_calculation_patches(root.path(), 2, patches).unwrap_err();
        assert!(reason.contains("full"), "{reason}");
        assert_eq!(persistent_state(root.path()), before);
        assert!(!root.path().join("native-active.json").exists());
        assert_eq!(
            crate::apply_calculation_patches(root.path(), 2, std::iter::empty()).unwrap(),
            2
        );
    }
    #[test]
    fn legacy_json_patches_share_the_writer_and_keep_empty_version_behavior() {
        let root = patch_fixture();
        let before = persistent_state(root.path());
        let empty = crate::apply_calculation_patches(root.path(), 1, std::iter::empty()).unwrap();
        assert_eq!(empty, 1);
        assert_eq!(persistent_state(root.path()), before);
        let result = crate::command(
            json!({"root":root.path(),"operation":"patch","version":1,"changes":[]}),
        )
        .unwrap();
        assert_ne!(result["version"], 1);
        let result = crate::command(json!({"root":root.path(),"operation":"patch","version":1,"changes":[{"sheet":"Sheet","cell":"A1","value":{"value":19,"dataType":"number"},"edits":[{"value":19}]}]})).unwrap();
        let version = result["version"].as_u64().unwrap();
        let db = Connection::open(root.path().join("cells.sqlite")).unwrap();
        assert_eq!(
            inherited_edit(&db, version, "Sheet", 0, 0).unwrap(),
            json!({"value":19,"fontColor":"#123456"})
        );
        assert_eq!(
            read_db(
                &db,
                &json!({"version":version,"sheet":"Sheet","range":"A1"})
            )
            .unwrap()["cells"][0][1]["value"],
            19
        );
    }
    #[test]
    fn inherited_intent_absence_and_sibling_hits_do_not_borrow_other_branches() {
        for indexed in [false, true] {
            let db = fixture(indexed);
            let before = main_schema(&db);
            for (version, expected) in [
                (2, json!({"value":14,"fontColor":"#123456","bold":true})),
                (3, json!({"value":16,"fontColor":"#123456","italic":true})),
                (1, json!({"value":12,"fontColor":"#123456"})),
                (0, json!({})),
            ] {
                assert_eq!(
                    inherited_edit(&db, version, "Sheet", 0, 0).unwrap(),
                    expected
                );
                assert_eq!(
                    inherited_edit(&db, version, "Sheet", 2, 0).unwrap(),
                    if version == 3 {
                        json!({"value":"sibling-only","wrapText":true})
                    } else {
                        json!({})
                    }
                );
                assert_eq!(
                    inherited_edit(&db, version, "Sheet", 1000, 100).unwrap(),
                    json!({})
                );
            }
            assert_eq!(
                main_schema(&db),
                before,
                "pure reads may not acquire an index"
            );
        }
    }
    #[test]
    fn same_connection_changes_refresh_version_cache_without_changing_parent_intents() {
        let db = fixture(true);
        assert_eq!(inherited_edit(&db, 2, "Sheet", 0, 0).unwrap()["value"], 14);
        db.execute("INSERT INTO versions VALUES(4,2)", []).unwrap();
        db.execute("INSERT INTO changes VALUES(4,'Sheet',0,0,'{}','{\"value\":18,\"fontColor\":\"#123456\",\"bold\":true,\"numberFormat\":\"date\"}')",[]).unwrap();
        assert_eq!(
            inherited_edit(&db, 4, "Sheet", 0, 0).unwrap(),
            json!({"value":18,"fontColor":"#123456","bold":true,"numberFormat":"date"})
        );
        assert_eq!(
            inherited_edit(&db, 2, "Sheet", 0, 0).unwrap(),
            json!({"value":14,"fontColor":"#123456","bold":true})
        );
        assert_eq!(inherited_edit(&db, 3, "Sheet", 0, 0).unwrap()["value"], 16);
    }
    #[test]
    fn range_fast_paths_keep_precise_source_fields_formulas_pagination_and_negative_zero() {
        for indexed in [false, true] {
            let _scope = crate::shared::begin();
            let db = fixture(indexed);
            let before = main_schema(&db);
            let request =
                |version, range: &str| json!({"version":version,"sheet":"Sheet","range":range});
            let source = read_db(&db, &request(0, "A1:B2")).unwrap();
            assert_eq!(source["cells"][0][1]["rawValue"], "9007199254740993");
            assert_eq!(source["cells"][0][1]["styleId"], 7);
            assert_eq!(
                source["cells"][1][1]["value"].as_f64().unwrap().to_bits(),
                (-0.0f64).to_bits()
            );
            assert_eq!(source["cells"][1][1]["rawValue"], "-0");
            assert_eq!(source["cells"][2][1]["cachedValue"], 913.0);
            let branch = read_db(&db, &request(2, "A1:B3")).unwrap();
            assert_eq!(branch["cells"].as_array().unwrap().len(), 3);
            assert_eq!(branch["cells"][0][1]["value"], 14);
            assert!(branch["cells"][2][1].get("formula").is_none());
            let sibling = read_db(&db, &request(3, "A1:B3")).unwrap();
            assert_eq!(sibling["cells"].as_array().unwrap().len(), 4);
            assert_eq!(sibling["cells"][3][1]["value"], "sibling-only");
            let formulas = read_db(
                &db,
                &json!({"version":2,"sheet":"Sheet","range":"A1:B3","formulas":true}),
            )
            .unwrap();
            assert_eq!(formulas["cells"], json!([]));
            let formulas = read_db(
                &db,
                &json!({"version":3,"sheet":"Sheet","range":"A1:B3","formulas":true}),
            )
            .unwrap();
            assert_eq!(formulas["cells"][0][0], "A2");
            let page = read_db(
                &db,
                &json!({"version":3,"sheet":"Sheet","range":"A1:B3","limit":2}),
            )
            .unwrap();
            assert_eq!(page["nextCursor"], "B1");
            let page = read_db(
                &db,
                &json!({"version":3,"sheet":"Sheet","range":"A1:B3","limit":2,"after":"B1"}),
            )
            .unwrap();
            assert_eq!(page["cells"][0][0], "A2");
            assert_eq!(page["cells"][1][0], "A3");
            assert_eq!(
                read_db(&db, &request(3, "Z100:Z200")).unwrap()["cells"],
                json!([])
            );
            assert_eq!(main_schema(&db), before);
        }
    }
}
pub(crate) const STYLE_FIELDS: &[&str] = &[
    "bold",
    "italic",
    "underline",
    "fontFamily",
    "fontSize",
    "fontColor",
    "fillColor",
    "horizontalAlign",
    "verticalAlign",
    "wrapText",
    "border",
    "numberFormat",
];

pub(crate) fn merge_edit(edit: &mut Value, delta: &Value) -> Result<()> {
    if let Some(kind) = delta.get("valueType") {
        if kind != "error"
            || delta.get("formula").is_some()
            || !delta["value"]
                .as_str()
                .is_some_and(crate::bulk::values::error_code)
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid declared error value",
            ));
        }
    }
    let map = edit
        .as_object_mut()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid stored edit"))?;
    if delta.get("value").is_some() {
        map.remove("formula");
        map.remove("valueType");
    }
    if delta.get("formula").is_some() {
        map.remove("value");
        map.remove("valueType");
    }
    if delta.get("styleId").is_some() {
        for key in STYLE_FIELDS {
            map.remove(*key);
        }
    }
    for key in ["value", "valueType", "formula", "styleId"]
        .iter()
        .chain(STYLE_FIELDS.iter())
    {
        if let Some(value) = delta.get(*key) {
            map.insert((*key).into(), value.clone());
        }
    }
    Ok(())
}
pub(crate) fn store_change(
    db: &Connection,
    version: i64,
    sheet: &str,
    row: u32,
    col: u32,
    cell: &Value,
    edit: &Value,
) -> Result<()> {
    if (edit.get("value").is_some() || edit.get("formula").is_some())
        && let Some(reason) = crate::preservation::reason(db, sheet, row, col)?
    {
        return Err(crate::coded("workbook-refused", reason));
    }
    if cell.get("$sourceFields").is_some() {
        return Err(crate::coded(
            "formula-unsupported",
            "Reserved calculation storage field",
        ));
    }
    if cell.get("spill").is_some() {
        return Err(crate::coded(
            "spill-refused",
            "Spill output fields belong to the complete native calculation owner",
        ));
    }
    store_calculated_change(db, version, sheet, row, col, cell, edit)
}

pub(crate) fn store_calculated_change(
    db: &Connection,
    version: i64,
    sheet: &str,
    row: u32,
    col: u32,
    cell: &Value,
    edit: &Value,
) -> Result<()> {
    if edit.get("cache").is_some()
        && let Some(reason) = crate::preservation::cache_reason(db, sheet, row, col)?
    {
        return Err(crate::coded("workbook-refused", reason));
    }

    crate::used_range::record(db, version, sheet, row, col, cell)?;
    db.prepare_cached("INSERT INTO changes(version,sheet,r,c,data,edit) VALUES(?1,?2,?3,?4,?5,?6)")
        .map_err(err)?
        .execute(rusqlite::params![
            version,
            sheet,
            row,
            col,
            serde_json::to_string(cell).map_err(err)?,
            serde_json::to_string(edit).map_err(err)?
        ])
        .map_err(err)?;
    Ok(())
}
