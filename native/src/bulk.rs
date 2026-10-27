//! Range value operations retain only a target tile and a bounded source cache.
//! Sources are read from the parent version, including for overlapping copies.
use crate::range::{STYLE_FIELDS, inherited_edit, merge_edit, read_db, store_change};
use crate::xml::{address, position};
use crate::{Result, database, err, field, local};
use rusqlite::Connection;
use serde_json::{Value, json};
use std::{
    collections::{BTreeMap, VecDeque},
    path::Path,
};

mod across_sheets;
pub(crate) mod fill_guard;
pub(crate) mod values;

pub(crate) type Bounds = ((u32, u32), (u32, u32));
pub(crate) fn bounds(range: &str) -> Result<Bounds> {
    let mut refs = range.split(':');
    let first = position(
        refs.next()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing range"))?,
    )?;
    let last = refs.next().map(position).transpose()?.unwrap_or(first);
    if refs.next().is_some() || first.0 > last.0 || first.1 > last.1 {
        return Err(crate::coded("internal-invalid-field", "Invalid range"));
    }
    Ok((first, last))
}
fn tile(
    db: &Connection,
    version: u64,
    sheet: &str,
    row: u32,
    col: u32,
) -> Result<BTreeMap<String, Value>> {
    let page = read_db(
        db,
        &json!({"version":version,"sheet":sheet,"range":format!("{}:{}",address(row,col),address((row+31).min(1048575),(col+31).min(16383))),"limit":2048}),
    )?;
    Ok(page["cells"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid native tile"))?
        .iter()
        .map(|c| (c[0].as_str().unwrap().into(), c[1].clone()))
        .collect())
}
type CachedTile = ((u32, u32), BTreeMap<String, Value>);
fn source_cell(
    cache: &mut VecDeque<CachedTile>,
    db: &Connection,
    version: u64,
    sheet: &str,
    row: u32,
    col: u32,
) -> Result<Value> {
    let key = (row / 32 * 32, col / 32 * 32);
    if !cache.iter().any(|(at, _)| *at == key) {
        if cache.len() == 8 {
            cache.pop_front();
        }
        let mut cells = tile(db, version, sheet, key.0, key.1)?;
        for (reference, cell) in &mut cells {
            let (r, c) = position(reference)?;
            crate::calculation_policy::mask(db, version, sheet, reference, cell)?;
            cell["$edit"] = inherited_edit(db, version, sheet, r, c)?;
            if let Some(reason) = crate::preservation::reason(db, sheet, r, c)? {
                cell["$valuePreservation"] = json!(reason);
            }
        }
        cache.push_back((key, cells));
    }
    Ok(cache
        .iter()
        .find(|(at, _)| *at == key)
        .unwrap()
        .1
        .get(&address(row, col))
        .cloned()
        .unwrap_or(json!({"value":null,"styleId":0})))
}
pub(crate) fn meaning(raw: &str) -> Option<(bool, String, i64)> {
    let (number, exponent) = raw
        .split_once(['e', 'E'])
        .map_or((raw, "0"), |(a, b)| (a, b));
    let negative = number.starts_with('-');
    let number = number.trim_start_matches(['+', '-']);
    let (integer, fraction) = number.split_once('.').unwrap_or((number, ""));
    let mut digits = format!("{integer}{fraction}")
        .trim_start_matches('0')
        .to_owned();
    if !digits.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    let mut power = exponent
        .parse::<i64>()
        .ok()?
        .checked_sub(fraction.len() as i64)?;
    if digits.is_empty() {
        return Some((false, "0".into(), 0));
    }
    while digits.ends_with('0') {
        digits.pop();
        power = power.checked_add(1)?;
    }
    Some((negative, digits, power))
}
pub(crate) fn literal(cell: &Value, values: bool) -> Result<Value> {
    if !values && let Some(reason) = cell["$valuePreservation"].as_str() {
        return Err(crate::coded("fill-refused", reason));
    }
    if cell.get("formulaType").is_some() {
        return Err(crate::coded(
            "fill-unsupported",
            "Special formula copy is not qualified",
        ));
    }
    if cell.get("formula").is_some() || cell["calculation"]["pending"] == true {
        if !values {
            return Err(crate::coded(
                "needs-structural-writer",
                "Formula fill requires the shared structural writer",
            ));
        }
        if cell["calculation"]["status"] != "calculated" {
            return Err(crate::coded(
                "fill-bad-input",
                "Paste Values needs a calculated formula result",
            ));
        }
        return Ok(cell["calculation"]["value"].clone());
    }
    let value = cell["value"].clone();
    if let Some(n) = value.as_f64() {
        if !n.is_finite()
            || cell["rawValue"]
                .as_str()
                .is_some_and(|raw| meaning(raw) != meaning(&n.to_string()))
        {
            return Err(crate::coded(
                "fill-refused",
                "Copy would lose stored numeric precision",
            ));
        }
    } else if value.is_object() {
        return Err(crate::coded(
            "fill-refused",
            "Copy would lose stored numeric precision",
        ));
    }
    Ok(value)
}
fn set_value(cell: &mut Value, value: Value) {
    let map = cell.as_object_mut().unwrap();
    for key in [
        "formula",
        "formulaType",
        "sharedFormulaId",
        "formulaReference",
        "calculation",
        "cachedResult",
        "scenarioResult",
        "formulaOutput",
        "spill",
        "error",
    ] {
        map.remove(key);
    }
    cell["rawValue"] = if value.is_null() {
        Value::Null
    } else if let Some(s) = value.as_str() {
        json!(s)
    } else {
        json!(value.to_string())
    };
    cell["dataType"] = json!(if value.is_null() {
        "blank"
    } else if value.is_boolean() {
        "boolean"
    } else if value.is_number() {
        "number"
    } else {
        "text"
    });
    cell["cachedValue"] = Value::Null;
    cell["value"] = value;
}
// Build the immutable accepted overlay once per scan, rather than reranking
// every revision for every 512-cell page. Temporary SQLite stays disk-backed.
fn normalized_without_active_masters(db: &Connection, version: u64, sheet: &str) -> Result<bool> {
    let normalized: bool = db
        .query_row(
            "SELECT EXISTS(SELECT 1 FROM shared_active WHERE sheet=?1)",
            [sheet],
            |r| r.get(0),
        )
        .map_err(err)?;
    if !normalized {
        return Ok(false);
    }
    let _span = crate::diagnostics::span("shared.special_master_check");
    // Build these sparse indexes only once compact normalization has occurred.
    // Ordinary imports do not pay for indexes they will never need. Candidate
    // masters include every ancestor override, including newly introduced array
    // or table masters; the latest value at each coordinate decides its meaning.
    // The generic version/sheet key otherwise wins SQLite's cost estimate and
    // scans every calculated follower twice. Use the sparse master index for
    // discovery, then pin candidates before exact coordinate lookups. All
    // ancestral masters remain candidates; the nearest override still wins.
    db.execute_batch("CREATE INDEX IF NOT EXISTS cells_special_master ON cells(sheet,r,c) WHERE formula IS NOT NULL AND json_extract(formula_attrs,'$.t') IS NOT NULL AND json_type(formula_attrs,'$.ref')='text'; CREATE INDEX IF NOT EXISTS changes_special_master ON changes(version,sheet,r,c) WHERE json_extract(data,'$.formulaType') IS NOT NULL AND json_type(data,'$.formulaReference')='text';").map_err(err)?;
    let active: bool=db.prepare_cached("WITH RECURSIVE ancestry(id,depth) AS (SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN ancestry a ON v.id=a.id WHERE v.parent IS NOT NULL), candidates AS MATERIALIZED (SELECT r,c FROM cells WHERE sheet=?2 AND formula IS NOT NULL AND json_extract(formula_attrs,'$.t') IS NOT NULL AND json_type(formula_attrs,'$.ref')='text' UNION SELECT p.r,p.c FROM ancestry a JOIN changes p INDEXED BY changes_special_master ON p.version=a.id WHERE p.sheet=?2 AND json_extract(p.data,'$.formulaType') IS NOT NULL AND json_type(p.data,'$.formulaReference')='text'), ranked AS (SELECT k.r,k.c,p.data,row_number() OVER(PARTITION BY k.r,k.c ORDER BY a.depth) priority FROM ancestry a CROSS JOIN candidates k CROSS JOIN changes p ON p.version=a.id AND p.sheet=?2 AND p.r=k.r AND p.c=k.c), effective AS (SELECT CASE WHEN o.r IS NULL THEN json_extract(c.formula_attrs,'$.t') ELSE json_extract(o.data,'$.formulaType') END kind, CASE WHEN o.r IS NULL THEN json_type(c.formula_attrs,'$.ref') ELSE json_type(o.data,'$.formulaReference') END reference_kind, CASE WHEN o.r IS NULL THEN json_quote(json_extract(c.formula_attrs,'$.si')) ELSE json_quote(json_extract(o.data,'$.sharedFormulaId')) END group_id FROM candidates k LEFT JOIN cells c ON c.sheet=?2 AND c.r=k.r AND c.c=k.c LEFT JOIN ranked o ON o.r=k.r AND o.c=k.c AND o.priority=1) SELECT EXISTS(SELECT 1 FROM effective e WHERE e.kind IS NOT NULL AND e.reference_kind='text' AND NOT(e.kind='shared' AND EXISTS(SELECT 1 FROM shared_active g WHERE g.sheet=?2 AND g.group_id=e.group_id)))").map_err(crate::err)?.query_row(rusqlite::params![version,sheet], |r|r.get(0)).map_err(err)?;
    Ok(!active)
}
fn prepare_special_scan(
    db: &Connection,
    version: u64,
    sheet: &str,
    calculation: bool,
) -> Result<()> {
    crate::shared::prepare(db, version)?;
    db.execute_batch("CREATE TEMP TABLE IF NOT EXISTS special_overlay(r INTEGER,c INTEGER,data TEXT,edit TEXT,PRIMARY KEY(r,c)) WITHOUT ROWID; DELETE FROM special_overlay;").map_err(err)?;
    db.execute_batch("CREATE TEMP TABLE IF NOT EXISTS special_positions(r INTEGER,c INTEGER,master INTEGER NOT NULL,PRIMARY KEY(r,c)) WITHOUT ROWID; CREATE INDEX IF NOT EXISTS special_masters ON special_positions(r,c) WHERE master=1; DELETE FROM special_positions;").map_err(err)?;
    if normalized_without_active_masters(db, version, sheet)? {
        crate::diagnostics::count("shared.special_scan_skipped", 1);
        return Ok(());
    }
    if calculation {
        // calculate() already materialized this exact immutable version. Reuse
        // that overlay instead of ranking every ancestor a second time.
        let rows = db
            .execute(
                "INSERT INTO special_overlay SELECT r,c,data,edit FROM calc_overlay WHERE sheet=?1",
                [sheet],
            )
            .map_err(err)?;
        crate::diagnostics::count("calculation.shared_overlay_reused_rows", rows as u64);
    } else {
        db.execute("WITH RECURSIVE ancestry(id,depth) AS (SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN ancestry a ON v.id=a.id WHERE v.parent IS NOT NULL), ranked AS (SELECT p.r,p.c,p.data,p.edit,row_number() OVER (PARTITION BY p.r,p.c ORDER BY a.depth) priority FROM changes p JOIN ancestry a ON p.version=a.id WHERE p.sheet=?2) INSERT INTO special_overlay SELECT r,c,data,edit FROM ranked WHERE priority=1",rusqlite::params![version,sheet]).map_err(err)?;
    }
    db.execute("INSERT INTO special_positions SELECT c.r,c.c,COALESCE(CASE WHEN o.r IS NULL THEN json_type(c.formula_attrs,'$.ref') ELSE json_type(o.data,'$.formulaReference') END='text',0) FROM cells c LEFT JOIN special_overlay o ON o.r=c.r AND o.c=c.c WHERE c.sheet=?1 AND c.formula IS NOT NULL AND CASE WHEN o.r IS NULL THEN json_extract(c.formula_attrs,'$.t') ELSE json_extract(o.data,'$.formulaType') END IS NOT NULL",[sheet]).map_err(err)?;
    db.execute_batch("INSERT OR IGNORE INTO special_positions SELECT r,c,COALESCE(json_type(data,'$.formulaReference')='text',0) FROM special_overlay WHERE json_extract(data,'$.formulaType') IS NOT NULL;").map_err(err)?;
    let normalized: bool = db
        .query_row(
            "SELECT EXISTS(SELECT 1 FROM shared_active WHERE sheet=?1)",
            [sheet],
            |r| r.get(0),
        )
        .map_err(err)?;
    if normalized {
        db.execute("DELETE FROM special_positions WHERE EXISTS(SELECT 1 FROM shared_active g LEFT JOIN cells c ON c.sheet=?1 AND c.r=special_positions.r AND c.c=special_positions.c LEFT JOIN special_overlay o ON o.r=special_positions.r AND o.c=special_positions.c WHERE g.sheet=?1 AND g.group_id=CASE WHEN o.r IS NULL THEN json_quote(json_extract(c.formula_attrs,'$.si')) ELSE json_quote(json_extract(o.data,'$.sharedFormulaId')) END AND CASE WHEN o.r IS NULL THEN json_extract(c.formula_attrs,'$.t') ELSE json_extract(o.data,'$.formulaType') END='shared')",[sheet]).map_err(err)?;
    }
    Ok(())
}
// Page source coordinates before overlay decoding. Excluding every overridden
// source coordinate before LIMIT would rescan the remaining base on every page.
fn special_page(db: &Connection, sheet: &str, after: Option<&str>, masters: bool) -> Result<Value> {
    let cursor = after
        .map(position)
        .transpose()?
        .map(|(r, c)| (i64::from(r), i64::from(c)))
        .unwrap_or((-1, -1));
    let filter = if masters { "AND master=1" } else { "" };
    let sql = format!(
        "WITH coordinates AS (SELECT r,c FROM special_positions WHERE (r,c)>(?2,?3) {filter} ORDER BY r,c LIMIT 512) SELECT k.r,k.c,c.kind,c.raw,c.inline,c.formula,c.formula_attrs,c.style,s.value,o.data,o.edit FROM coordinates k LEFT JOIN cells c ON c.sheet=?1 AND c.r=k.r AND c.c=k.c LEFT JOIN strings s ON c.kind='s' AND s.id=CAST(c.raw AS INTEGER) LEFT JOIN special_overlay o ON o.r=k.r AND o.c=k.c ORDER BY k.r,k.c LIMIT 512"
    );
    let mut query = db.prepare_cached(&sql).map_err(err)?;
    let mut rows = query
        .query(rusqlite::params![sheet, cursor.0, cursor.1])
        .map_err(err)?;
    let mut cells = Vec::new();
    while let Some(row) = rows.next().map_err(err)? {
        crate::cancel::check()?;
        let (reference, cell) = crate::range::decode_cell(row)?;
        let attributes: Option<String> = row.get(6).map_err(err)?;
        let edit: Option<String> = row.get(10).map_err(err)?;
        cells.push(json!([reference, cell, attributes, edit]));
    }
    Ok(
        json!({"nextCursor":if cells.len()==512 {cells.last().map(|c|c[0].clone())}else{None},"cells":cells}),
    )
}
// Expand only intersected shared groups, inside the caller's transaction. The
// rectangle is validated against every follower before a version is published.
// Array/data-table formulas are not storage compression and never enter here.
pub(crate) fn expand_shared(
    db: &Connection,
    parent: u64,
    sheet: &str,
    target: Bounds,
    partial_only: bool,
) -> Result<u64> {
    expand_shared_into(db, parent, sheet, target, partial_only, false)
}
pub(crate) fn calculate_shared(db: &Connection, parent: u64, sheet: &str) -> Result<()> {
    let present:bool=db.prepare_cached("SELECT EXISTS(SELECT 1 FROM cells WHERE sheet=?1 AND formula IS NOT NULL AND json_extract(formula_attrs,'$.t')='shared')").map_err(crate::err)?.query_row([sheet],|row|row.get(0)).map_err(err)?;
    if !present {
        return Ok(());
    }

    expand_shared_into(db, parent, sheet, ((0, 0), (1048575, 16383)), false, true)?;
    Ok(())
}
fn expand_shared_into(
    db: &Connection,
    parent: u64,
    sheet: &str,
    target: Bounds,
    partial_only: bool,
    calculation: bool,
) -> Result<u64> {
    if !calculation
        && let Some(version) =
            crate::shared::normalize_validated(db, parent, sheet, target, partial_only)?
    {
        return Ok(version);
    }
    let inside =
        |p: (u32, u32), b: Bounds| p.0 >= b.0.0 && p.0 <= b.1.0 && p.1 >= b.0.1 && p.1 <= b.1.1;
    prepare_special_scan(db, parent, sheet, calculation)?;
    let mut groups = BTreeMap::new();
    let mut after = None;
    loop {
        crate::cancel::check()?;
        let page = special_page(db, sheet, after.as_deref(), true)?;
        for pair in page["cells"]
            .as_array()
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid formula page"))?
        {
            let cell = &pair[1];
            if cell["formulaType"] != "shared" {
                continue;
            }
            let Some(reference) = cell["formulaReference"].as_str() else {
                continue;
            };
            let b = bounds(reference)?;
            if target.0.0 > b.1.0 || target.1.0 < b.0.0 || target.0.1 > b.1.1 || target.1.1 < b.0.1
            {
                continue;
            }
            if partial_only && inside(b.0, target) && inside(b.1, target) {
                continue;
            }
            let at = position(pair[0].as_str().ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing shared address")
            })?)?;
            if !inside(at, b)
                || cell["sharedFormulaId"].is_null()
                || cell["formula"].as_str().is_none_or(str::is_empty)
            {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid shared formula master",
                ));
            }
            if groups.len() >= 1024
                || groups
                    .insert(
                        cell["sharedFormulaId"].to_string(),
                        (
                            b,
                            at,
                            cell["formula"].as_str().unwrap().to_string(),
                            0_u64,
                            crate::formula::copy_invariant(cell["formula"].as_str().unwrap())?,
                        ),
                    )
                    .is_some()
            {
                return Err(crate::coded(
                    "fill-duplicate",
                    "Duplicate or excessive shared groups",
                ));
            }
        }
        after = page["nextCursor"].as_str().map(str::to_owned);
        if after.is_none() {
            break;
        }
    }
    if groups.is_empty() {
        return Ok(parent);
    }
    let positions: u64 = groups
        .values()
        .map(|(b, _, _, _, _)| u64::from(b.1.0 - b.0.0 + 1) * u64::from(b.1.1 - b.0.1 + 1))
        .sum();
    if positions > 16 * 1024 * 1024 {
        return Err(crate::coded(
            "fill-limit",
            "Shared formula expansion minimum byte budget",
        ));
    }
    let version = if calculation {
        parent as i64
    } else {
        db.execute("INSERT INTO versions(parent) VALUES(?1)", [parent])
            .map_err(err)?;
        db.last_insert_rowid()
    };
    let mut bytes = 0;
    let mut after = None;
    loop {
        crate::cancel::check()?;
        let page = special_page(db, sheet, after.as_deref(), false)?;
        for pair in page["cells"]
            .as_array()
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid formula page"))?
        {
            let cell = &pair[1];
            if cell["formulaType"] != "shared" {
                continue;
            }
            let Some((b, master, formula, count, invariant)) =
                groups.get_mut(&cell["sharedFormulaId"].to_string())
            else {
                continue;
            };
            let at = position(pair[0].as_str().ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing shared address")
            })?)?;
            if !inside(at, *b) {
                return Err(crate::coded(
                    "fill-unsupported",
                    "Shared formula expansion would orphan a follower",
                ));
            }
            if at != *master
                && (cell["formulaReference"].is_string()
                    || cell["formula"].as_str().is_some_and(|s| !s.is_empty()))
            {
                return Err(crate::coded(
                    "fill-unsupported",
                    "Inconsistent shared formula follower",
                ));
            }
            // The page already read the immutable source attributes. Avoid a
            // second SQL lookup per follower while preserving extended-attribute guards.
            if let Some(attributes) = pair[2].as_str() {
                let attributes: Value = serde_json::from_str(attributes).map_err(err)?;
                if attributes
                    .as_object()
                    .is_none_or(|m| m.keys().any(|k| !["t", "si", "ref"].contains(&k.as_str())))
                {
                    return Err(crate::coded(
                        "fill-unsupported",
                        "Extended shared formula attributes are not qualified",
                    ));
                }
            }
            let shifted = if *invariant {
                json!(formula)
            } else {
                crate::formula::rewrite(&json!({"items":[{"formula":formula,"copyRows":i64::from(at.0)-i64::from(master.0),"copyColumns":i64::from(at.1)-i64::from(master.1)}]}))?["formulas"][0].clone()
            };
            bytes += shifted
                .as_str()
                .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid expanded formula"))?
                .len();
            if bytes > 16 * 1024 * 1024 {
                return Err(crate::coded(
                    "fill-limit",
                    "Shared formula expansion byte budget",
                ));
            }
            // Validation still visits every follower and enforces the original
            // expansion budget. Persistence below is one record per group.
            *count += 1;
        }
        after = page["nextCursor"].as_str().map(str::to_owned);
        if after.is_none() {
            break;
        }
    }
    // OOXML sharing belongs only to explicit t=shared/si members. Ordinary
    // formulas, literals and blanks inside ref override the shared definition.
    // A member names its group by id, so two groups whose ranges overlap (WPS
    // and Excel both write these) are still unambiguous.
    if calculation {
        for (id, (bounds, master, formula, _, invariant)) in &groups {
            db.prepare_cached("INSERT INTO calc_shared VALUES(?1,?2,?3,?4,?5,?6,?7,?8,?9,?10)")
                .map_err(err)?
                .execute(rusqlite::params![
                    sheet, id, master.0, master.1, formula, invariant, bounds.0.0, bounds.0.1,
                    bounds.1.0, bounds.1.1
                ])
                .map_err(err)?;
        }
        crate::diagnostics::count("calculation.shared_groups", groups.len() as u64);
        crate::diagnostics::count(
            "calculation.shared_members_validated",
            groups.values().map(|(_, _, _, count, _)| *count).sum(),
        );
    } else {
        for (id, (_, master, formula, _, invariant)) in &groups {
            db.prepare_cached("INSERT INTO shared_normalizations VALUES(?1,?2,?3,?4,?5,?6,?7)")
                .map_err(err)?
                .execute(rusqlite::params![
                    version, sheet, id, master.0, master.1, formula, invariant
                ])
                .map_err(err)?;
        }
        crate::diagnostics::count("shared.normalized_groups", groups.len() as u64);
        crate::diagnostics::count("shared.validated_members", positions);
    }
    Ok(version as u64)
}
/// How many formulas in the imported workbook are array or data-table formulas.
/// Shared formulas are only a compact spelling and can be written out in full;
/// these cannot.
pub(crate) fn special_kinds(root: &Path) -> Result<Value> {
    let db = database(root)?;
    let other: u64 = db
        .query_row(
            "SELECT count(*) FROM cells WHERE formula IS NOT NULL \
             AND json_type(formula_attrs,'$.t') IS NOT NULL \
             AND json_extract(formula_attrs,'$.t')!='shared'",
            [],
            |row| row.get(0),
        )
        .map_err(err)?;
    Ok(json!({"notShared":other}))
}
pub(crate) fn normalize_shared(root: &Path, input: &Value) -> Result<Value> {
    let mut db = database(root)?;
    let tx = db.transaction().map_err(err)?;
    let mut parent = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing version"))?;
    for item in input["ranges"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing shared ranges"))?
    {
        parent = expand_shared(
            &tx,
            parent,
            field(item, "sheet")?,
            bounds(field(item, "range")?)?,
            false,
        )?;
    }
    tx.commit().map_err(err)?;
    Ok(json!({"version":parent}))
}
// Whole-group clearing is safe for legacy shared and array formulas. Partial
// edits must also detect ordinary cached cells inside an array's master range.
pub(crate) fn special_groups(
    db: &Connection,
    version: u64,
    sheet: &str,
    target: Bounds,
    clear: bool,
) -> Result<()> {
    prepare_special_scan(db, version, sheet, false)?;
    db.execute_batch("CREATE TEMP TABLE IF NOT EXISTS clear_shared(id TEXT PRIMARY KEY); DELETE FROM clear_shared;").map_err(err)?;
    let mut after = None;
    loop {
        crate::cancel::check()?;
        let page = special_page(db, sheet, after.as_deref(), true)?;
        for pair in page["cells"]
            .as_array()
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid formula page"))?
        {
            let cell = &pair[1];
            let Some(kind) = cell["formulaType"].as_str() else {
                continue;
            };
            let Some(reference) = cell["formulaReference"].as_str() else {
                continue;
            };
            let group = bounds(reference)?;
            let at = position(pair[0].as_str().ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing formula address")
            })?)?;
            if at.0 < group.0.0 || at.0 > group.1.0 || at.1 < group.0.1 || at.1 > group.1.1 {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid special formula master range",
                ));
            }
            let intersects = target.0.0 <= group.1.0
                && target.1.0 >= group.0.0
                && target.0.1 <= group.1.1
                && target.1.1 >= group.0.1;
            if !intersects {
                continue;
            }
            if kind == "array"
                && group.0 == group.1
                && cell["formula"].as_str().is_some_and(|formula| {
                    crate::formula::spill_consumer(formula)
                        .ok()
                        .flatten()
                        .is_some()
                        || crate::formula::scalar_array_root(formula)
                })
            {
                return Err(crate::coded(
                    "fill-unsupported",
                    "Imported scalar spill consumers retain source array metadata; content replacement is not qualified",
                ));
            }
            let contains = target.0.0 <= group.0.0
                && target.1.0 >= group.1.0
                && target.0.1 <= group.0.1
                && target.1.1 >= group.1.1;
            if !clear || !contains || !["shared", "array", "dataTable"].contains(&kind) {
                return Err(crate::coded(
                    "fill-unsupported",
                    "A shared, array or data-table formula must be replaced as a complete group; partial edits are not qualified",
                ));
            }
            if kind == "shared" {
                if cell["sharedFormulaId"].is_null()
                    || cell["formula"].as_str().is_none_or(str::is_empty)
                {
                    return Err(crate::coded(
                        "internal-invalid-field",
                        "Invalid shared formula master",
                    ));
                }
                db.execute(
                    "INSERT INTO clear_shared VALUES(?1)",
                    [cell["sharedFormulaId"].to_string()],
                )
                .map_err(err)?;
            }
        }
        after = page["nextCursor"].as_str().map(str::to_owned);
        if after.is_none() {
            break;
        }
    }
    // Shared followers can omit their own range. Refuse malformed groups whose
    // remaining followers would be orphaned by clearing the master.
    if clear {
        let mut after = None;
        loop {
            crate::cancel::check()?;
            let page = special_page(db, sheet, after.as_deref(), false)?;
            for pair in page["cells"]
                .as_array()
                .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid formula page"))?
            {
                let cell = &pair[1];
                if cell["formulaType"] != "shared" {
                    continue;
                }
                let cleared = db
                    .query_row(
                        "SELECT count(*) FROM clear_shared WHERE id=?1",
                        [cell["sharedFormulaId"].to_string()],
                        |r| r.get::<_, u32>(0),
                    )
                    .map_err(err)?
                    == 1;
                let at = position(pair[0].as_str().ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing formula address")
                })?)?;
                if cleared
                    && (at.0 < target.0.0
                        || at.0 > target.1.0
                        || at.1 < target.0.1
                        || at.1 > target.1.1)
                {
                    return Err(crate::coded(
                        "fill-unsupported",
                        "Clearing would orphan a shared formula follower",
                    ));
                }
            }
            after = page["nextCursor"].as_str().map(str::to_owned);
            if after.is_none() {
                break;
            }
        }
    }
    Ok(())
}

pub fn apply(root: &Path, input: &Value) -> Result<Value> {
    let index: Value =
        serde_json::from_slice(&std::fs::read(local(root, "index.json")?).map_err(err)?)
            .map_err(err)?;
    let commands = input["commands"]
        .as_array()
        .filter(|c| !c.is_empty())
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing bulk commands"))?;
    // Derive protected merge/table geometry from the accepted package, never
    // from caller projections, before a candidate transaction can write.
    across_sheets::validate_runs(commands)?;
    let mut db = database(root)?;
    fill_guard::guard_version(
        root,
        &index,
        commands,
        &db,
        input["version"]
            .as_u64()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing version"))?,
    )?;
    let tx = db.transaction().map_err(err)?;
    let mut parent = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing version"))?;
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
    if input["formulaOverrides"] == true {
        return Err(crate::coded(
            "fill-bad-input",
            "External formula overrides are not accepted",
        ));
    }
    let mut formula_touched = false;
    let has_across = commands
        .iter()
        .any(|command| command["type"] == "fillAcrossSheets");
    let mut content_changed = false;
    let has_spills = crate::spill::present(&tx)?;
    let mut formula_delta: BTreeMap<String, i64> = BTreeMap::new();
    for command in commands {
        let kind = field(command, "type")?;
        if kind == "fillAcrossSheets" {
            let copied = across_sheets::apply(root, &tx, parent, command, &index)?;
            parent = copied.version;
            formula_touched |= copied.formula_touched;
            content_changed |= copied.content_changed;
            if copied.content_changed {
                *formula_delta
                    .entry(field(command, "sheet")?.into())
                    .or_default() += copied.formula_delta;
            }
            continue;
        }
        if ![
            "clearRange",
            "fillRange",
            "pasteValues",
            "pasteFormulas",
            "fillSeries",
        ]
        .contains(&kind)
        {
            return Err(crate::coded("fill-unsupported", "Unsupported bulk command"));
        }
        if command
            .as_object()
            .unwrap()
            .keys()
            .any(|k| !["type", "sheet", "sourceSheet", "source", "range"].contains(&k.as_str()))
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid bulk command field",
            ));
        }
        content_changed = true;
        let sheet = field(command, "sheet")?;
        let source_sheet = command["sourceSheet"].as_str().unwrap_or(sheet);
        for name in [sheet, source_sheet] {
            if !index["sheets"]
                .as_array()
                .unwrap()
                .iter()
                .any(|s| s["path"] == name)
            {
                return Err(crate::coded(
                    "internal-unknown-reference",
                    "Unknown worksheet",
                ));
            }
        }
        let (first, last) = bounds(field(command, "range")?)?;
        if crate::calculation_policy::preserving(&tx, parent)? {
            if !["clearRange", "pasteValues"].contains(&kind) {
                return Err(crate::coded(
                    "fill-unsupported",
                    "Only literal clearing and Paste Values are qualified while calculation is preserved",
                ));
            }
            crate::preservation::region(root, &tx, sheet, first, last, true)?;
        }
        if has_spills {
            crate::spill::guard(&tx, parent, sheet, &crate::spill::Geometry { first, last })?;
        }
        if index["sheets"]
            .as_array()
            .unwrap()
            .iter()
            .any(|s| s["path"] == sheet && s["specialFormulaCount"].as_u64().unwrap_or(0) > 0)
        {
            let normalized = expand_shared(&tx, parent, sheet, (first, last), true)?;
            formula_touched |= normalized != parent;
            parent = normalized;
            special_groups(
                &tx,
                parent,
                sheet,
                (first, last),
                ["clearRange", "pasteValues", "fillRange"].contains(&kind),
            )?;
        }
        let source = if kind == "clearRange" || kind == "fillSeries" {
            None
        } else {
            Some(bounds(field(command, "source")?)?)
        };
        if has_spills
            && kind != "pasteValues"
            && let Some((first, last)) = source
        {
            crate::spill::guard(
                &tx,
                parent,
                source_sheet,
                &crate::spill::Geometry { first, last },
            )?;
        }
        let mut values_context = values::Context::new();
        let series = if kind == "fillSeries" {
            let vertical = first.1 == last.1 && last.0 >= first.0 + 2;
            let horizontal = first.0 == last.0 && last.1 >= first.1 + 2;
            if !vertical && !horizontal {
                return Err(crate::coded(
                    "fill-bad-input",
                    "Fill Series needs one row or column with two seeds and a destination",
                ));
            }
            let mut cache = VecDeque::new();
            let a = source_cell(&mut cache, &tx, parent, sheet, first.0, first.1)?;
            let b = source_cell(
                &mut cache,
                &tx,
                parent,
                sheet,
                first.0 + u32::from(vertical),
                first.1 + u32::from(horizontal),
            )?;
            let a = literal(&a, false)?.as_f64().ok_or_else(|| {
                crate::coded("fill-bad-input", "Fill Series needs exact numeric seeds")
            })?;
            let b = literal(&b, false)?.as_f64().ok_or_else(|| {
                crate::coded("fill-bad-input", "Fill Series needs exact numeric seeds")
            })?;
            Some((vertical, a, b - a))
        } else {
            None
        };
        if kind == "pasteFormulas" {
            let (a, b) = source.unwrap();
            if last.0 - first.0 != b.0 - a.0
                || last.1 - first.1 != b.1 - a.1
                || u64::from(last.0 - first.0 + 1) * u64::from(last.1 - first.1 + 1) > 1000
            {
                return Err(crate::coded(
                    "fill-limit",
                    "Paste Formulas requires matching source and target rectangles of at most 1,000 cells",
                ));
            }
        }
        let mut copied_formulas = 0_u64;
        tx.execute("INSERT INTO versions(parent) VALUES(?1)", [parent])
            .map_err(err)?;
        let version = tx.last_insert_rowid();
        let mut cache = VecDeque::new();
        for row in (first.0 / 32 * 32..=last.0).step_by(32) {
            crate::cancel::check()?;
            for col in (first.1 / 32 * 32..=last.1).step_by(32) {
                let mut cells = tile(&tx, parent, sheet, row, col)?;
                for r in row.max(first.0)..=(row + 31).min(last.0) {
                    for c in col.max(first.1)..=(col + 31).min(last.1) {
                        let mut cell = cells
                            .remove(&address(r, c))
                            .unwrap_or(json!({"value":null,"styleId":0}));
                        if let Some((vertical, _, _)) = series {
                            let offset = if vertical { r - first.0 } else { c - first.1 };
                            if offset < 2 {
                                continue;
                            }
                            if !cell["value"].is_null()
                                || cell.get("formula").is_some()
                                || cell.get("formulaType").is_some()
                            {
                                return Err(crate::coded(
                                    "fill-conflict",
                                    format!("Fill Series would overwrite {}", address(r, c)),
                                ));
                            }
                        }
                        if let Some(kind) = cell["formulaType"].as_str() {
                            let allowed =
                                ["clearRange", "pasteValues", "fillRange"]
                                    .contains(&command["type"].as_str().unwrap_or(""))
                                    && match kind {
                                        "array" | "dataTable" => {
                                            cell["formulaReference"].as_str().is_some()
                                        }
                                        "shared" => !cell["sharedFormulaId"].is_null()
                                            && tx
                                                .query_row(
                                                    "SELECT count(*) FROM clear_shared WHERE id=?1",
                                                    [cell["sharedFormulaId"].to_string()],
                                                    |r| r.get::<_, u32>(0),
                                                )
                                                .map_err(err)?
                                                == 1,
                                        _ => false,
                                    };
                            if !allowed {
                                return Err(crate::coded(
                                    "fill-unsupported",
                                    "Editing a shared, array or data-table formula cell is not qualified",
                                ));
                            }
                        }
                        let mut edit = inherited_edit(&tx, parent, sheet, r, c)?;
                        let from = if let Some((a, b)) = source {
                            let sr = a.0 + (r - first.0) % (b.0 - a.0 + 1);
                            let sc = a.1 + (c - first.1) % (b.1 - a.1 + 1);
                            Some((
                                source_cell(&mut cache, &tx, parent, source_sheet, sr, sc)?,
                                sr,
                                sc,
                            ))
                        } else {
                            None
                        };
                        // Cached source followers are not authored literals.
                        // Check before Formula Paste can skip a non-formula cell.
                        if ["fillRange", "pasteFormulas"].contains(&kind)
                            && from.as_ref().is_some_and(|(cell, _, _)| {
                                cell.get("formulaOutput").is_some()
                                    || cell["scenarioResult"] == true
                            })
                        {
                            return Err(crate::coded(
                                "fill-unsupported",
                                "Source formula output copy is not qualified",
                            ));
                        }
                        if kind == "pasteFormulas"
                            && from
                                .as_ref()
                                .is_none_or(|(cell, _, _)| cell.get("formula").is_none())
                        {
                            continue;
                        }
                        let was_formula =
                            cell.get("formula").is_some() || cell.get("formulaType").is_some();
                        let formula = if kind != "pasteValues" {
                            if let Some((from, sr, sc)) = &from {
                                if from.get("formulaType").is_some() {
                                    return Err(crate::coded(
                                        "fill-unsupported",
                                        "Special formula copy is not qualified",
                                    ));
                                }
                                if let Some(formula) = from["formula"].as_str() {
                                    match crate::formula::rewrite(
                                        &json!({"items":[{"formula":formula,"copyRows":i64::from(r)-i64::from(*sr),"copyColumns":i64::from(c)-i64::from(*sc)}]}),
                                    ) {
                                        Ok(shifted) => Some(shifted["formulas"][0].clone()),
                                        Err(reason) => {
                                            return Err(crate::coded(
                                                "fill-unsupported",
                                                format!(
                                                    "Formula copy syntax is not qualified for native preservation: {reason}"
                                                ),
                                            ));
                                        }
                                    }
                                } else {
                                    None
                                }
                            } else {
                                None
                            }
                        } else {
                            None
                        };
                        let mut delta;
                        if let Some(formula) = formula {
                            set_value(&mut cell, Value::Null);
                            cell["formula"] = formula.clone();
                            delta = json!({"formula":formula});
                            copied_formulas += 1;
                        } else {
                            let (value, typed_error) = if let Some((vertical, start, step)) = series
                            {
                                let offset = if vertical { r - first.0 } else { c - first.1 };
                                let value = start + step * f64::from(offset);
                                if !value.is_finite() {
                                    return Err(crate::coded(
                                        "fill-limit",
                                        "Fill Series produced a number outside workbook limits",
                                    ));
                                }
                                // Match the editor's established Number(value.toPrecision(15)).
                                (json!(crate::decimal::series_precision(value)?), false)
                            } else if let Some((from, sr, sc)) = &from {
                                if kind == "pasteValues" {
                                    let snapshot = values_context.read(
                                        root,
                                        &tx,
                                        parent,
                                        source_sheet,
                                        (*sr, *sc),
                                        from,
                                    )?;
                                    (snapshot.value, snapshot.error)
                                } else {
                                    let snapshot = values::literal_snapshot(from)?;
                                    (snapshot.value, snapshot.error)
                                }
                            } else {
                                (Value::Null, false)
                            };
                            delta = json!({"value":value});
                            set_value(&mut cell, value);
                            if typed_error {
                                delta["valueType"] = json!("error");
                                cell["dataType"] = json!("error");
                                cell["error"] = cell["value"].clone();
                            }
                        }
                        let is_formula = cell.get("formula").is_some();
                        formula_touched |= was_formula || is_formula;
                        *formula_delta.entry(sheet.into()).or_default() +=
                            i64::from(is_formula) - i64::from(was_formula);
                        let from = from.map(|(cell, _, _)| cell);
                        if kind == "fillRange"
                            && let Some(from) = from
                        {
                            let style = from["styleId"].as_u64().unwrap_or(0);
                            delta["styleId"] = json!(style);
                            for key in STYLE_FIELDS {
                                cell.as_object_mut().unwrap().remove(*key);
                            }
                            cell.as_object_mut().unwrap().remove("format");
                            cell["styleId"] = json!(style);
                            for key in STYLE_FIELDS {
                                if let Some(v) = from["$edit"].get(*key) {
                                    delta[*key] = v.clone();
                                }
                                if let Some(v) = from.get(*key) {
                                    cell[*key] = v.clone();
                                }
                            }
                            if let Some(v) = from.get("format") {
                                cell["format"] = v.clone();
                            }
                        }
                        merge_edit(&mut edit, &delta)?;
                        store_change(&tx, version, sheet, r, c, &cell, &edit)?;
                    }
                }
            }
        }
        if kind == "pasteFormulas" && copied_formulas == 0 {
            return Err(crate::coded(
                "fill-bad-input",
                "Copied range contains no formulas",
            ));
        }
        parent = version as u64;
    }
    tx.commit().map_err(err)?;
    let mut result =
        json!({"version":parent,"formulaDelta":formula_delta,"formulaTouched":formula_touched});
    if has_across {
        result["contentChanged"] = json!(content_changed);
    }
    Ok(result)
}

#[cfg(test)]
mod master_lookup_tests {
    use super::*;
    use rusqlite::params;

    #[test]
    fn sparse_master_lookup_obeys_nearest_override_and_branch_ancestry() {
        let db = Connection::open_in_memory().unwrap();
        db.execute_batch("CREATE TABLE versions(id INTEGER PRIMARY KEY,parent INTEGER);
            INSERT INTO versions VALUES(0,NULL),(1,0),(2,1),(3,0),(4,2);
            CREATE TABLE cells(sheet TEXT,r INTEGER,c INTEGER,formula TEXT,formula_attrs TEXT,PRIMARY KEY(sheet,r,c)) WITHOUT ROWID;
            CREATE TABLE changes(version INTEGER,sheet TEXT,r INTEGER,c INTEGER,data TEXT,PRIMARY KEY(version,sheet,r,c)) WITHOUT ROWID;
            CREATE TEMP TABLE shared_active(sheet TEXT,group_id TEXT,PRIMARY KEY(sheet,group_id)) WITHOUT ROWID;").unwrap();
        db.execute(
            "INSERT INTO cells VALUES('s',0,0,'1+1',?1)",
            [json!({"t":"shared","si":"0","ref":"A1:A10000"}).to_string()],
        )
        .unwrap();
        db.execute(
            "INSERT INTO shared_active VALUES('s',?1)",
            [json!("0").to_string()],
        )
        .unwrap();
        let safe = |version| normalized_without_active_masters(&db, version, "s").unwrap();
        assert!(
            safe(0),
            "Normalized shared masters are not active special groups"
        );
        db.execute(
            "INSERT INTO cells VALUES('s',0,3,'1+1',?1)",
            [json!({"t":"array","ref":"D1:D2"}).to_string()],
        )
        .unwrap();
        assert!(!safe(0), "An original array master remains protected");
        let patch = |version, col, value: Value| {
            db.execute(
                "INSERT INTO changes VALUES(?1,'s',0,?2,?3)",
                params![version, col, value.to_string()],
            )
            .unwrap();
        };
        patch(1, 3, json!({"value":9}));
        assert!(
            safe(1),
            "An ordinary override removes the original array master"
        );
        patch(
            1,
            2,
            json!({"formulaType":"array","formulaReference":"C1:C2","formula":"1+1"}),
        );
        assert!(
            !safe(1),
            "New masters in ancestor overlays must be discovered"
        );
        patch(2, 2, json!({"value":7}));
        assert!(
            safe(2),
            "The nearest ordinary override defeats an ancestral master"
        );
        assert!(!safe(1), "Older revisions keep their own protected group");
        patch(3, 3, json!({"value":8}));
        assert!(
            safe(3),
            "A sibling branch must not inherit another branch's master"
        );
        patch(
            2,
            0,
            json!({"formulaType":"dataTable","formulaReference":"A1:A2","formula":""}),
        );
        assert!(
            !safe(2),
            "Replacing a normalized shared master with a data table stays protected"
        );
        patch(4, 0, json!({"value":6}));
        assert!(safe(4));
        assert!(!safe(0), "Original source semantics remain unchanged");
    }
}
