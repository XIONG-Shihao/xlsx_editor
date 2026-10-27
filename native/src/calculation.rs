//! Version-fenced calculation: input pages and staged results stay beside SQLite.
use crate::xml::{address, position};
use crate::{Result, database, field};
fn err(error: impl std::fmt::Display) -> String {
    format!("SQLite calculation: {error}")
}
fn unsupported(reason: String) -> String {
    let said = crate::sentence(&reason);
    if said.starts_with("SQLite calculation:")
        || said.to_ascii_lowercase().contains("cancel")
        || said.contains("What-If scenario calculation cannot publish")
    {
        reason
    } else {
        crate::rephrase("formula-unsupported", &reason, |said| {
            format!("Unsupported native calculation: {said}")
        })
    }
}
use rusqlite::{Connection, OptionalExtension, params};
use serde_json::{Value, json};
use std::path::Path;
mod spill_inputs;

pub(crate) fn snapshot_tables(db: &Connection, version: u64) -> Result<()> {
    crate::shared::prepare(db, version)?;
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
            "Unknown calculation version",
        ));
    }
    // Bulk jobs can rebuild a snapshot on the same connection. Its old table
    // cleanup does not own this evaluation-only projection; retire it here.
    db.execute_batch("DROP TABLE IF EXISTS temp.calc_owned_spills; CREATE TEMP TABLE calc_overlay(sheet TEXT,r INTEGER,c INTEGER,data TEXT,edit TEXT,PRIMARY KEY(sheet,r,c)) WITHOUT ROWID; CREATE TEMP TABLE calc_shared(sheet TEXT,group_id TEXT,master_r INTEGER,master_c INTEGER,formula TEXT,invariant INTEGER,r0 INTEGER,c0 INTEGER,r1 INTEGER,c1 INTEGER,PRIMARY KEY(sheet,group_id)) WITHOUT ROWID; CREATE TEMP TABLE calc_targets(sheet TEXT,r INTEGER,c INTEGER,PRIMARY KEY(sheet,r,c)) WITHOUT ROWID; CREATE TEMP TABLE calc_values(sheet TEXT,ref TEXT,r INTEGER,c INTEGER,value TEXT NOT NULL,PRIMARY KEY(sheet,ref)) WITHOUT ROWID; CREATE TEMP TABLE calc_owned_spills(sheet TEXT,r INTEGER,c INTEGER,literal TEXT NOT NULL,PRIMARY KEY(sheet,r,c)) WITHOUT ROWID; CREATE TEMP TABLE calc_traces(owner_sheet TEXT,owner_ref TEXT,sheet TEXT,r0 INTEGER,r1 INTEGER,c0 INTEGER,c1 INTEGER); CREATE INDEX calc_trace_owner ON calc_traces(owner_sheet,owner_ref);").map_err(err)?;
    Ok(())
}
pub(crate) fn snapshot(db: &Connection, version: u64) -> Result<()> {
    let _span = crate::diagnostics::span("calculation.snapshot");
    snapshot_tables(db, version)?;
    // The snapshot key selects one immutable value per coordinate. Visiting
    // nearest ancestors first lets that key reject older values directly,
    // without window-sorting every cumulative JSON cell and export edit.
    let mut ancestry = db.prepare_cached("WITH RECURSIVE ancestry(id,depth) AS (SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN ancestry a ON v.id=a.id WHERE v.parent IS NOT NULL) SELECT id FROM ancestry ORDER BY depth").map_err(err)?;
    let mut versions = ancestry.query([version]).map_err(err)?;
    let mut insert = db.prepare_cached("INSERT OR IGNORE INTO calc_overlay SELECT sheet,r,c,data,edit FROM changes WHERE version=?1").map_err(err)?;
    let mut rows = 0;
    while let Some(ancestor) = versions.next().map_err(err)? {
        crate::cancel::check()?;
        rows += insert
            .execute([ancestor.get::<_, u64>(0).map_err(err)?])
            .map_err(err)?;
        crate::diagnostics::count("calculation.snapshot_ancestor_queries", 1);
    }
    crate::diagnostics::count("calculation.snapshot_rows", rows as u64);
    Ok(())
}
/// Formula-edge replacement inspects only the changed formula positions. It
/// never evaluates their inputs, so unrelated overlays need not be materialized.
/// Evaluation and structural callers still require a complete snapshot.
pub(crate) fn snapshot_changes(db: &Connection, version: u64, input: &Value) -> Result<()> {
    let _span = crate::diagnostics::span("calculation.snapshot_changes");
    snapshot_tables(db, version)?;
    for change in input["changedRanges"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing changes"))?
    {
        crate::cancel::check()?;
        let (a, b) = crate::bulk::bounds(field(change, "range")?)?;
        let name = field(change, "sheet")?;
        let sheet = input["view"]["sheets"]
            .as_array()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheets"))?
            .iter()
            .find(|sheet| sheet["name"] == name)
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheet"))?;
        let rows = db.execute("WITH RECURSIVE ancestry(id,depth) AS (SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN ancestry a ON v.id=a.id WHERE v.parent IS NOT NULL), ranked AS (SELECT p.*,row_number() OVER (PARTITION BY p.sheet,p.r,p.c ORDER BY a.depth) priority FROM changes p JOIN ancestry a ON p.version=a.id WHERE p.sheet=?2 AND p.r BETWEEN ?3 AND ?4 AND p.c BETWEEN ?5 AND ?6) INSERT OR IGNORE INTO calc_overlay SELECT sheet,r,c,data,edit FROM ranked WHERE priority=1",params![version,field(sheet,"path")?,a.0,b.0,a.1,b.1]).map_err(err)?;
        crate::diagnostics::count("calculation.snapshot_rows", rows as u64);
    }
    Ok(())
}

pub(crate) fn page(
    db: &Connection,
    sheet: &str,
    first: (u32, u32),
    last: (u32, u32),
    after: Option<(u32, u32)>,
    formulas: bool,
) -> Result<Vec<(String, Value)>> {
    page_impl(db, sheet, first, last, after, formulas, true)
}
/// Evaluation can accept a proved, unaffected result without translating its
/// compressed shared formula. Structural/dependency readers still use page().
pub(crate) fn input_page(
    db: &Connection,
    sheet: &str,
    first: (u32, u32),
    last: (u32, u32),
    after: Option<(u32, u32)>,
    formulas: bool,
) -> Result<Vec<(String, Value)>> {
    page_impl(db, sheet, first, last, after, formulas, false)
}
fn page_impl(
    db: &Connection,
    sheet: &str,
    first: (u32, u32),
    last: (u32, u32),
    after: Option<(u32, u32)>,
    formulas: bool,
    resolve: bool,
) -> Result<Vec<(String, Value)>> {
    let cursor = after
        .map(|(r, c)| (i64::from(r), i64::from(c)))
        .unwrap_or((-1, -1));
    // A literal formula predicate lets SQLite use the sparse formula index.
    // The parameterized OR in ordinary range reads otherwise scans all cells.
    let base_filter = if formulas {
        "formula IS NOT NULL AND NOT EXISTS(SELECT 1 FROM calc_overlay o WHERE o.sheet=cells.sheet AND o.r=cells.r AND o.c=cells.c AND json_extract(o.data,'$.formula') IS NULL AND json_extract(o.data,'$.formulaType') IS NULL)"
    } else {
        "1"
    };
    let sql = format!(
        "WITH base AS (SELECT r,c FROM cells WHERE sheet=?1 AND r BETWEEN ?2 AND ?3 AND c BETWEEN ?4 AND ?5 AND (r,c)>(?6,?7) AND ({base_filter}) ORDER BY r,c LIMIT 512), added AS (SELECT r,c FROM calc_overlay WHERE sheet=?1 AND r BETWEEN ?2 AND ?3 AND c BETWEEN ?4 AND ?5 AND (r,c)>(?6,?7) AND (?8=0 OR json_extract(data,'$.formula') IS NOT NULL OR json_extract(data,'$.formulaType') IS NOT NULL) ORDER BY r,c LIMIT 512), coordinates AS (SELECT r,c FROM base UNION SELECT r,c FROM added) SELECT k.r,k.c,c.kind,c.raw,c.inline,c.formula,c.formula_attrs,c.style,s.value,o.data,p.literal FROM coordinates k LEFT JOIN cells c ON c.sheet=?1 AND c.r=k.r AND c.c=k.c LEFT JOIN strings s ON c.kind='s' AND s.id=CAST(c.raw AS INTEGER) LEFT JOIN calc_overlay o ON o.sheet=?1 AND o.r=k.r AND o.c=k.c LEFT JOIN calc_owned_spills p ON p.sheet=?1 AND p.r=k.r AND p.c=k.c ORDER BY k.r,k.c LIMIT 512"
    );
    let mut statement = db.prepare_cached(&sql).map_err(err)?;
    let mut rows = statement
        .query(params![
            sheet,
            after.map_or(first.0, |p| first.0.max(p.0)),
            last.0,
            first.1,
            last.1,
            cursor.0,
            cursor.1,
            u32::from(formulas)
        ])
        .map_err(err)?;
    let mut result = Vec::new();
    while let Some(row) = rows.next().map_err(err)? {
        crate::cancel::check()?;
        if !resolve && let Some(literal) = row.get::<_, Option<String>>(10).map_err(err)? {
            result.push((
                address(row.get(0).map_err(err)?, row.get(1).map_err(err)?),
                serde_json::from_str(&literal).map_err(err)?,
            ));
            continue;
        }
        let (reference, cell) = if resolve {
            crate::range::decode_cell(row)?
        } else {
            crate::range::decode_evaluation_cell(row)?
        };
        let cell = if resolve {
            crate::shared::resolve(db, sheet, &reference, cell)?
        } else {
            cell
        };
        result.push((reference, cell))
    }
    Ok(result)
}
pub(crate) fn cell(db: &Connection, sheet: &str, reference: &str) -> Result<Value> {
    crate::shared::resolve(db, sheet, reference, input_cell(db, sheet, reference)?)
}
pub(crate) fn input_cell(db: &Connection, sheet: &str, reference: &str) -> Result<Value> {
    read_cell(db, sheet, reference, false)
}
pub(crate) fn evaluation_cell(db: &Connection, sheet: &str, reference: &str) -> Result<Value> {
    read_cell(db, sheet, reference, true)
}
fn read_cell(db: &Connection, sheet: &str, reference: &str, formula_only: bool) -> Result<Value> {
    let (r, c) = position(reference)?;
    let mut query=db.prepare_cached("SELECT ?2,?3,c.kind,c.raw,c.inline,c.formula,c.formula_attrs,c.style,s.value,o.data,p.literal FROM (SELECT 1) LEFT JOIN cells c ON c.sheet=?1 AND c.r=?2 AND c.c=?3 LEFT JOIN strings s ON c.kind='s' AND s.id=CAST(c.raw AS INTEGER) LEFT JOIN calc_overlay o ON o.sheet=?1 AND o.r=?2 AND o.c=?3 LEFT JOIN calc_owned_spills p ON p.sheet=?1 AND p.r=?2 AND p.c=?3 WHERE c.r IS NOT NULL OR o.r IS NOT NULL").map_err(err)?;
    let mut rows = query.query(params![sheet, r, c]).map_err(err)?;
    match rows.next().map_err(err)? {
        Some(row) => Ok(
            if formula_only && let Some(literal) = row.get::<_, Option<String>>(10).map_err(err)? {
                serde_json::from_str(&literal).map_err(err)?
            } else if formula_only {
                crate::range::decode_evaluation_cell(row)?.1
            } else {
                crate::range::decode_cell(row)?.1
            },
        ),
        None => Ok(Value::Null),
    }
}
pub(crate) fn resolved_cell(
    db: &Connection,
    sheet: &str,
    reference: &str,
    cell: &Value,
) -> Result<Value> {
    if cell["formulaType"] != "shared" {
        return Ok(cell.clone());
    }
    if crate::shared::normalized(db, sheet, cell)? {
        return crate::shared::resolve(db, sheet, reference, cell.clone());
    }
    let (r, c) = position(reference)?;
    let group: Option<(u32, u32, String, bool)> = db
        .prepare_cached("SELECT master_r,master_c,formula,invariant FROM calc_shared WHERE sheet=?1 AND group_id=?2")
        .map_err(err)?
        .query_row(params![sheet, cell["sharedFormulaId"].to_string()], |row| Ok((row.get(0)?, row.get(1)?, row.get(2)?, row.get(3)?)))
        .optional().map_err(err)?;
    let mut cell = cell.clone();
    if let Some((mr, mc, formula, invariant)) = group {
        let shifted = if invariant {
            json!(formula)
        } else {
            crate::formula::rewrite(&json!({"items":[{"formula":formula,"copyRows":i64::from(r)-i64::from(mr),"copyColumns":i64::from(c)-i64::from(mc)}]}))?["formulas"][0].clone()
        };
        cell.as_object_mut().unwrap().remove("formulaType");
        cell["formula"] = shifted;
    }
    Ok(cell)
}

pub(crate) fn proved_shared(db: &Connection, sheet: &str, cell: &Value) -> Result<bool> {
    if crate::shared::normalized(db, sheet, cell)? {
        return Ok(true);
    }
    db.prepare_cached("SELECT EXISTS(SELECT 1 FROM calc_shared WHERE sheet=?1 AND group_id=?2)")
        .map_err(err)?
        .query_row(params![sheet, cell["sharedFormulaId"].to_string()], |r| {
            r.get(0)
        })
        .map_err(err)
}

pub(crate) fn memo(db: &Connection, sheet: &str, reference: &str) -> Result<Option<Value>> {
    let raw: Option<String> = db
        .prepare_cached("SELECT value FROM calc_values WHERE sheet=?1 AND ref=?2")
        .map_err(err)?
        .query_row(params![sheet, reference], |r| r.get(0))
        .optional()
        .map_err(err)?;
    raw.map(|s| serde_json::from_str(&s).map_err(err))
        .transpose()
}

pub(crate) fn remember(db: &Connection, sheet: &str, reference: &str, value: &Value) -> Result<()> {
    let (r, c) = position(reference)?;
    db.prepare_cached("INSERT OR REPLACE INTO calc_values VALUES(?1,?2,?3,?4,?5)")
        .map_err(err)?
        .execute(params![sheet, reference, r, c, value.to_string()])
        .map_err(err)?;
    Ok(())
}
pub(crate) fn selected(db: &Connection, sheet: &str, reference: &str) -> Result<bool> {
    let (r, c) = position(reference)?;
    Ok(db
        .prepare_cached("SELECT 1 FROM calc_targets WHERE sheet=?1 AND r=?2 AND c=?3")
        .map_err(err)?
        .query_row(params![sheet, r, c], |_| Ok(()))
        .optional()
        .map_err(err)?
        .is_some())
}
pub(crate) fn append(db: &Connection, plan: u64, item: &Value) -> Result<()> {
    let sheet = field(item, "sheet")?;
    let reference = field(item, "ref")?;
    let (r, c) = position(reference)?;
    db.prepare_cached("INSERT INTO dependency_formulas VALUES(?1,?2,?3,?4,?5)")
        .map_err(err)?
        .execute(params![plan, sheet, reference, r, c])
        .map_err(err)?;
    for dep in item["ranges"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing dependencies"))?
    {
        let (a, b) = if let Some(r) = dep["ref"].as_str() {
            crate::bulk::bounds(r)?
        } else {
            (
                (
                    dep["startRow"].as_u64().unwrap() as u32,
                    dep["startColumn"].as_u64().unwrap() as u32,
                ),
                (
                    dep["endRow"].as_u64().unwrap() as u32,
                    dep["endColumn"].as_u64().unwrap() as u32,
                ),
            )
        };
        db.prepare_cached("INSERT INTO dependency_ranges VALUES(?1,?2,?3,?4,?5,?6,?7,?8)")
            .map_err(err)?
            .execute(params![
                plan,
                sheet,
                reference,
                field(dep, "sheet")?,
                a.0,
                b.0,
                a.1,
                b.1
            ])
            .map_err(err)?;
    }
    Ok(())
}
fn calculate(db: &mut Connection, input: &Value) -> Result<Value> {
    let _syntax_cache = crate::formula::syntax_cache_scope();
    crate::dependencies::schema(db).map_err(err)?;
    let tx = db.transaction().map_err(err)?;
    let version = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing calculation version"))?;
    snapshot(&tx, version)?;
    let view = &input["view"];
    let sheets = view["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheets"))?;
    spill_inputs::prepare(&tx, version, view)?;
    let shared_span = crate::diagnostics::span("calculation.shared_prepare");
    for sheet in sheets {
        crate::bulk::calculate_shared(&tx, version, field(sheet, "path")?).map_err(unsupported)?;
    }
    drop(shared_span);
    {
        let _span = crate::diagnostics::span("calculation.scenarios");
        crate::scenarios::prepare(&tx, view).map_err(unsupported)?;
    }
    let plan_span = crate::diagnostics::span("calculation.plan");
    let partial = input["partial"] == true;
    let existing = input["plan"].as_u64();
    let mut plan;
    if partial && existing.is_none() {
        plan = 0;
        for sheet in sheets {
            let mut after = None;
            loop {
                let cells = page(
                    &tx,
                    field(sheet, "path")?,
                    (0, 0),
                    (1048575, 16383),
                    after,
                    true,
                )?;
                for (reference, _) in &cells {
                    let (r, c) = position(reference)?;
                    tx.prepare_cached("INSERT INTO calc_targets VALUES(?1,?2,?3)")
                        .map_err(err)?
                        .execute(params![field(sheet, "path")?, r, c])
                        .map_err(err)?;
                }
                if cells.len() < 512 {
                    break;
                }
                after = Some(position(&cells.last().unwrap().0)?);
            }
        }
    } else if let Some(id) = existing {
        plan = id;
        if tx
            .query_row(
                "SELECT count(*) FROM dependency_plans WHERE id=?1 AND complete=1 AND metadata=?2",
                params![id, field(input, "metadata")?],
                |r| r.get::<_, u32>(0),
            )
            .map_err(err)?
            != 1
        {
            return Err(crate::coded("formula-stale", "Stale calculation plan"));
        }
        for sheet in sheets {
            tx.execute("INSERT INTO calc_targets SELECT ?1,f.r,f.c FROM dependency_selected s JOIN dependency_formulas f ON f.plan=s.plan AND f.sheet=s.sheet AND f.ref=s.ref WHERE s.plan=?2 AND s.sheet=?3",params![field(sheet,"path")?,plan,field(sheet,"name")?]).map_err(err)?;
        }
    } else {
        tx.execute(
            "INSERT INTO dependency_plans(metadata) VALUES(?1)",
            [field(input, "metadata")?],
        )
        .map_err(err)?;
        plan = tx.last_insert_rowid() as u64;
        for sheet in sheets {
            let mut after = None;
            loop {
                let cells = page(
                    &tx,
                    field(sheet, "path")?,
                    (0, 0),
                    (1048575, 16383),
                    after,
                    true,
                )?;
                if cells.is_empty() {
                    break;
                }
                for (reference, cell) in &cells {
                    crate::cancel::check()?;
                    if cell["formulaType"] == "dataTable" {
                        continue;
                    }
                    let owned = cell.get("spill").is_some()
                        && spill_inputs::owned(&tx, field(sheet, "path")?, reference)?;
                    let deps = if owned {
                        // Keep constructor edges for future invalidation without
                        // dispatching an accepted array master as a scalar.
                        crate::formula::sequence_dependencies(
                            view,
                            field(sheet, "name")?,
                            reference,
                            field(cell, "formula")?,
                        )
                        .map_err(unsupported)?
                        .ok_or_else(|| {
                            crate::coded(
                                "formula-bad-input",
                                "Accepted spill has no constructor dependencies",
                            )
                        })?
                    } else {
                        crate::formula::cell_dependencies(
                            view,
                            field(sheet, "name")?,
                            reference,
                            &resolved_cell(&tx, field(sheet, "path")?, reference, cell)?,
                        )
                        .map_err(unsupported)?
                    };
                    append(&tx, plan, &deps)?;
                    if owned {
                        continue;
                    }
                    let (r, c) = position(reference)?;
                    tx.prepare_cached("INSERT INTO calc_targets VALUES(?1,?2,?3)")
                        .map_err(err)?
                        .execute(params![field(sheet, "path")?, r, c])
                        .map_err(err)?;
                }
                if cells.len() < 512 {
                    break;
                }
                after = Some(position(&cells.last().unwrap().0)?);
            }
        }
    }
    // Existing or partial plans may still select an accepted array master.
    // Complete spill output ownership, never the scalar writer, owns its cache.
    tx.execute("DELETE FROM calc_targets WHERE EXISTS(SELECT 1 FROM calc_owned_spills p WHERE p.sheet=calc_targets.sheet AND p.r=calc_targets.r AND p.c=calc_targets.c)", []).map_err(err)?;
    // Data Table output cells have no individual formula records. Include all
    // outputs with explicit model and parameter edges; model dependencies are transitive.
    for sheet in sheets {
        let mut query = tx
            .prepare_cached("SELECT r,c FROM calc_scenarios WHERE sheet=?1")
            .map_err(err)?;
        let mut rows = query.query([field(sheet, "path")?]).map_err(err)?;
        while let Some(row) = rows.next().map_err(err)? {
            let r: u32 = row.get(0).map_err(err)?;
            let c: u32 = row.get(1).map_err(err)?;
            if existing.is_none()
                || crate::calculation::selected(&tx, field(sheet, "path")?, &address(r, c))?
            {
                tx.execute(
                    "INSERT OR IGNORE INTO calc_targets VALUES(?1,?2,?3)",
                    params![field(sheet, "path")?, r, c],
                )
                .map_err(err)?;
            }
            if existing.is_none() && plan > 0 {
                append(
                    &tx,
                    plan,
                    &crate::scenarios::dependencies(
                        &tx,
                        view,
                        field(sheet, "name")?,
                        &address(r, c),
                    )?,
                )?;
            }
        }
    }
    drop(plan_span);
    let evaluate_span = crate::diagnostics::span("calculation.evaluate");
    let mut engine = crate::formula::Eval::from_sqlite(view, &tx, existing.is_some());
    let mut count = 0;
    let mut misses = 0;
    // Targets are paged too: dense formula books never allocate all ASTs/results.
    for sheet in sheets {
        let mut cursor = (-1_i64, -1_i64);
        loop {
            let targets = {
                let mut s=tx.prepare_cached("SELECT r,c FROM calc_targets WHERE sheet=?1 AND (r,c)>(?2,?3) ORDER BY r,c LIMIT 512").map_err(err)?;
                s.query_map(params![field(sheet, "path")?, cursor.0, cursor.1], |r| {
                    Ok((r.get::<_, u32>(0)?, r.get::<_, u32>(1)?))
                })
                .map_err(err)?
                .collect::<std::result::Result<Vec<_>, _>>()
                .map_err(err)?
            };
            if targets.is_empty() {
                break;
            }
            for (r, c) in &targets {
                match engine.result(field(sheet, "name")?, &address(*r, *c)) {
                    Ok(_) => count += 1,
                    Err(reason) => {
                        let reason = unsupported(reason);
                        if !partial
                            || !crate::sentence(&reason)
                                .starts_with("Unsupported native calculation: ")
                        {
                            return Err(crate::coded("formula-refused", reason));
                        }
                        misses += 1;
                        // A formerly valid target can become unsupported after
                        // a selector edit. Never rebind its stale precise edges.
                        let reference = address(*r, *c);
                        let failed = cell(&tx, field(sheet, "path")?, &reference)?;
                        if crate::formula::dynamic_cell(
                            view,
                            field(sheet, "name")?,
                            &reference,
                            &failed,
                        ) {
                            for target in sheets {
                                tx.execute(
                                    "INSERT INTO calc_traces VALUES(?1,?2,?3,0,1048575,0,16383)",
                                    params![
                                        field(sheet, "name")?,
                                        reference,
                                        field(target, "name")?
                                    ],
                                )
                                .map_err(err)?;
                            }
                        }
                    }
                }
            }
            let last = targets.last().unwrap();
            cursor = (i64::from(last.0), i64::from(last.1));
        }
    }
    let visits = engine.steps;
    let (entries, bytes) = engine.constant_cache_size();
    crate::diagnostics::count("calculation.constant_cache_entries", entries as u64);
    crate::diagnostics::count("calculation.constant_cache_bytes", bytes as u64);
    crate::diagnostics::count("calculation.constant_hits", engine.constant_hits);
    drop(engine);
    drop(evaluate_span);
    crate::diagnostics::count("calculation.cell_visits", visits as u64);
    crate::diagnostics::count("calculation.evaluated", count);
    crate::diagnostics::count("calculation.unsupported", misses);
    let persist_span = crate::diagnostics::span("calculation.persist");
    tx.execute("INSERT INTO versions(parent) VALUES(?1)", [version])
        .map_err(err)?;
    let accepted = tx.last_insert_rowid();
    {
        let mut s = tx
            .prepare_cached("SELECT v.r,v.c,c.kind,c.raw,c.inline,c.formula,c.formula_attrs,c.style,s.value,o.data,v.sheet,v.ref,v.value,o.edit,(c.r IS NOT NULL OR o.r IS NOT NULL),(z.r IS NOT NULL),(o.data IS NULL OR coalesce(json_extract(o.data,'$.$sourceFields'),0)=1) FROM calc_values v LEFT JOIN cells c ON c.sheet=v.sheet AND c.r=v.r AND c.c=v.c LEFT JOIN strings s ON c.kind='s' AND s.id=CAST(c.raw AS INTEGER) LEFT JOIN calc_overlay o ON o.sheet=v.sheet AND o.r=v.r AND o.c=v.c LEFT JOIN calc_scenarios z ON z.sheet=v.sheet AND z.r=v.r AND z.c=v.c ORDER BY v.sheet,v.ref")
            .map_err(err)?;
        let mut rows = s.query([]).map_err(err)?;
        while let Some(row) = rows.next().map_err(err)? {
            crate::cancel::check()?;
            let sheet: String = row.get(10).map_err(err)?;
            let reference: String = row.get(11).map_err(err)?;
            let value: Value =
                serde_json::from_str(&row.get::<_, String>(12).map_err(err)?).map_err(err)?;
            // Use the same decoder and shared-resolution semantics as cell(),
            // while SQLite streams the cell, scenario and inherited edit joins.
            let source_cell = if row.get::<_, bool>(14).map_err(err)? {
                crate::range::decode_cell(row)?.1
            } else {
                Value::Null
            };
            // Only these five fields may reference the immutable source. All
            // formula identity and calculation provenance remain explicit.
            let source_fields = (row.get::<_, bool>(14).map_err(err)?
                && row.get::<_, bool>(16).map_err(err)?)
            .then(|| crate::range::SOURCE_FIELDS.map(|key| source_cell[key].clone()));
            let mut cell = crate::shared::resolve(&tx, &sheet, &reference, source_cell)?;
            let scenario: bool = row.get(15).map_err(err)?;
            if scenario {
                cell["scenarioResult"] = json!(true);
                cell["value"] = value.clone();
            }
            cell["calculation"] =
                json!({"status":"calculated","value":value,"revision":input["revision"]});
            if partial {
                cell["calculation"]["nativePartial"] = json!(true);
            }
            let r: u32 = row.get(0).map_err(err)?;
            let c: u32 = row.get(1).map_err(err)?;
            let inherited: Option<String> = row.get(13).map_err(err)?;
            let mut edit = inherited
                .map(|s| serde_json::from_str(&s).map_err(err))
                .transpose()?
                .unwrap_or_else(|| json!({}));
            if cell.get("formulaType").is_none()
                && !scenario
                && !crate::range::source_formula_identity(row, &cell)?
            {
                crate::range::merge_edit(&mut edit, &json!({"formula":cell["formula"]}))?;
            }
            if let Some(source) = source_fields
                && !scenario
                && crate::range::SOURCE_FIELDS
                    .iter()
                    .zip(&source)
                    .all(|(key, value)| cell.get(*key) == Some(value))
            {
                let map = cell.as_object_mut().ok_or_else(|| {
                    crate::coded("internal-invalid-field", "Invalid calculated cell")
                })?;
                for key in crate::range::SOURCE_FIELDS {
                    map.remove(key);
                }
                map.insert("$sourceFields".into(), Value::Bool(true));
            }
            crate::range::store_calculated_change(&tx, accepted, &sheet, r, c, &cell, &edit)?;
        }
    }
    let traced_formulas: u64 = tx
        .query_row(
            "SELECT count(*) FROM (SELECT DISTINCT owner_sheet,owner_ref FROM calc_traces)",
            [],
            |r| r.get(0),
        )
        .map_err(err)?;
    let trace_source = input["tracePlan"].as_u64().unwrap_or(plan);
    if traced_formulas > 0 && trace_source > 0 {
        // Never mutate a plan referenced by an undo/older workbook version.
        let previous = trace_source;
        tx.execute(
            "INSERT INTO dependency_plans(metadata) VALUES(?1)",
            [field(input, "metadata")?],
        )
        .map_err(err)?;
        plan = tx.last_insert_rowid() as u64;
        tx.execute("INSERT INTO dependency_formulas SELECT ?1,sheet,ref,r,c FROM dependency_formulas WHERE plan=?2",params![plan,previous]).map_err(err)?;
        tx.execute("INSERT INTO dependency_ranges SELECT ?1,owner_sheet,owner_ref,sheet,r0,r1,c0,c1 FROM dependency_ranges d WHERE plan=?2 AND NOT EXISTS(SELECT 1 FROM calc_traces t WHERE t.owner_sheet=d.owner_sheet AND t.owner_ref=d.owner_ref)",params![plan,previous]).map_err(err)?;
        tx.execute("INSERT INTO dependency_ranges SELECT DISTINCT ?1,owner_sheet,owner_ref,sheet,r0,r1,c0,c1 FROM calc_traces",[plan]).map_err(err)?;
    }
    if !partial {
        tx.execute("UPDATE dependency_plans SET complete=1 WHERE id=?1", [plan])
            .map_err(err)?;
        tx.execute(
            "INSERT OR REPLACE INTO dependency_bindings VALUES(?1,?2)",
            params![accepted, plan],
        )
        .map_err(err)?;
    }
    crate::shared::store_validations(&tx, accepted as u64)?;
    drop(persist_span);
    {
        let _span = crate::diagnostics::span("calculation.commit");
        tx.commit().map_err(err)?;
    }
    Ok(
        json!({"supported":true,"tracedFormulaCells":traced_formulas,"runtimePlan":if traced_formulas>0 && trace_source>0 {Some(plan)} else {None},"partial":partial,"unsupportedFormulaCells":misses,"version":accepted,"nativeSqlite":true,"nativeParser":existing.is_none(),"nativeEvaluator":true,"nativePlanHit":existing.is_some(),"inputCells":0,"rangeReads":0,"nativeCellVisits":visits,"evaluatedFormulaCells":count,"mode":if existing.is_some(){"affected-native-sqlite"}else{"native-sqlite"}}),
    )
}
pub fn command(root: &Path, input: &Value) -> Result<Value> {
    if crate::package_edit::source_calculation_policy(root)?["unsupportedMode"] == true {
        return Err(crate::coded(
            "formula-protected",
            "Workbook calculation source settings are protected",
        ));
    }
    let mut db = database(root)?;
    let result = calculate(&mut db, input);
    crate::cancel::check()?;
    match result {
        Ok(value) => Ok(value),
        Err(reason) => {
            // Only an explicit semantic miss may enter the compatibility engine.
            // Database, request, quota and cancellation errors stay operational.
            let Some(said) =
                crate::sentence(&reason).strip_prefix("Unsupported native calculation: ")
            else {
                return Err(crate::coded("formula-refused", reason));
            };
            Ok(json!({"supported":false,"reason":said}))
        }
    }
}

#[cfg(test)]
mod tests;
