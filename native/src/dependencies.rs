//! Version-bound dependency plans and transitive invalidation live beside cells.
//! JavaScript supplies parser-proved static references; Rust stores and queries
//! them. Incomplete/uncertain plans can never authorize selective calculation.
use crate::bulk::bounds;
use crate::xml::{address, position};
use crate::{Result, database, err, field};
use rusqlite::{Connection, OptionalExtension, params};
use serde_json::{Value, json};
use std::path::Path;
// Keep an unselected formula outside both joins. ANALYZE can otherwise choose
// dirty positions as the outer loop after a large fan-out, revisiting all
// already-selected formula dependencies for every position in that frontier.
const SELECT_DEPENDENTS: &str = "INSERT OR IGNORE INTO frontier SELECT f.sheet,f.ref,f.r,f.c FROM dependency_formulas f CROSS JOIN dependency_ranges d ON d.plan=f.plan AND d.owner_sheet=f.sheet AND d.owner_ref=f.ref CROSS JOIN dirty q ON q.sheet=d.sheet AND q.r0<=d.r1 AND q.r1>=d.r0 AND q.c0<=d.c1 AND q.c1>=d.c0 WHERE f.plan=?1 AND NOT EXISTS(SELECT 1 FROM dependency_selected s WHERE s.plan=f.plan AND s.sheet=f.sheet AND s.ref=f.ref)";
pub(crate) fn select_plan(db: &Connection, plan: u64, changes: &Value) -> Result<()> {
    db.execute("DELETE FROM dependency_selected", [])
        .map_err(err)?;
    db.execute_batch("CREATE TEMP TABLE dirty(sheet TEXT,r0 INTEGER,r1 INTEGER,c0 INTEGER,c1 INTEGER);CREATE TEMP TABLE frontier(sheet TEXT,ref TEXT,r INTEGER,c INTEGER,PRIMARY KEY(sheet,ref));").map_err(err)?;
    for change in changes
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing dirty input ranges"))?
    {
        let (a, b) = bounds(field(change, "range")?)?;
        db.execute(
            "INSERT INTO dirty VALUES(?1,?2,?3,?4,?5)",
            params![field(change, "sheet")?, a.0, b.0, a.1, b.1],
        )
        .map_err(err)?;
    }
    db.execute("INSERT OR IGNORE INTO frontier SELECT f.sheet,f.ref,f.r,f.c FROM dependency_formulas f JOIN dirty q ON q.sheet=f.sheet AND f.r BETWEEN q.r0 AND q.r1 AND f.c BETWEEN q.c0 AND q.c1 WHERE f.plan=?1",[plan]).map_err(err)?;
    loop {
        crate::cancel::check()?;
        db.execute(SELECT_DEPENDENTS, [plan]).map_err(err)?;
        let added = db
            .execute(
                "INSERT OR IGNORE INTO dependency_selected SELECT ?1,sheet,ref FROM frontier",
                [plan],
            )
            .map_err(err)?;
        if added == 0 {
            break;
        }
        db.execute_batch("DELETE FROM dirty;INSERT INTO dirty SELECT sheet,r,r,c,c FROM frontier;DELETE FROM frontier;").map_err(err)?;
    }
    Ok(())
}
pub fn schema(db: &Connection) -> Result<()> {
    db.execute_batch("CREATE TABLE IF NOT EXISTS dependency_plans(id INTEGER PRIMARY KEY,metadata TEXT NOT NULL,complete INTEGER NOT NULL DEFAULT 0); CREATE TABLE IF NOT EXISTS dependency_bindings(version INTEGER PRIMARY KEY,plan INTEGER NOT NULL); CREATE TABLE IF NOT EXISTS dependency_formulas(plan INTEGER NOT NULL,sheet TEXT NOT NULL,ref TEXT NOT NULL,r INTEGER NOT NULL,c INTEGER NOT NULL,PRIMARY KEY(plan,sheet,ref)); CREATE TABLE IF NOT EXISTS dependency_ranges(plan INTEGER NOT NULL,owner_sheet TEXT NOT NULL,owner_ref TEXT NOT NULL,sheet TEXT NOT NULL,r0 INTEGER NOT NULL,r1 INTEGER NOT NULL,c0 INTEGER NOT NULL,c1 INTEGER NOT NULL); CREATE INDEX IF NOT EXISTS dependency_lookup ON dependency_ranges(plan,sheet,r0); CREATE INDEX IF NOT EXISTS dependency_owner ON dependency_ranges(plan,owner_sheet,owner_ref); CREATE TABLE IF NOT EXISTS dependency_selected(plan INTEGER NOT NULL,sheet TEXT NOT NULL,ref TEXT NOT NULL,PRIMARY KEY(plan,sheet,ref));").map_err(err)
}
fn closest(db: &Connection, version: u64, metadata: Option<&str>) -> Result<Option<u64>> {
    let row:Option<(u64,Option<u32>,Option<String>)>=db.prepare_cached("WITH RECURSIVE ancestry(id,depth) AS (SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN ancestry a ON v.id=a.id WHERE v.parent IS NOT NULL) SELECT b.plan,p.complete,p.metadata FROM ancestry a JOIN dependency_bindings b ON b.version=a.id LEFT JOIN dependency_plans p ON p.id=b.plan ORDER BY a.depth LIMIT 1").map_err(crate::err)?.query_row([version],|r|Ok((r.get(0)?,r.get(1)?,r.get(2)?))).optional().map_err(err)?;
    Ok(row.and_then(|(plan, ready, key)| {
        (ready == Some(1) && metadata.is_none_or(|m| key.as_deref() == Some(m))).then_some(plan)
    }))
}
pub fn retain(db: &Connection, keep: &[u64]) -> Result<()> {
    schema(db)?;
    for version in keep {
        let plan = closest(db, *version, None)?.unwrap_or(0);
        db.execute(
            "INSERT OR REPLACE INTO dependency_bindings VALUES(?1,?2)",
            params![version, plan],
        )
        .map_err(err)?;
    }
    Ok(())
}
pub fn prune(db: &Connection) -> Result<()> {
    let calculation = db
        .prepare_cached(
            "SELECT EXISTS(SELECT 1 FROM sqlite_master WHERE name='calculation_bindings')",
        )
        .map_err(crate::err)?
        .query_row([], |row| row.get::<_, bool>(0))
        .map_err(err)?;
    db.execute(
        "DELETE FROM dependency_bindings WHERE version NOT IN (SELECT id FROM versions)",
        [],
    )
    .map_err(err)?;
    if calculation {
        db.execute(
            "DELETE FROM calculation_bindings WHERE version NOT IN (SELECT id FROM versions)",
            [],
        )
        .map_err(err)?;
        db.execute("DELETE FROM dependency_plans WHERE id NOT IN (SELECT plan FROM dependency_bindings UNION SELECT plan FROM calculation_bindings)", []).map_err(err)?;
    } else {
        db.execute(
            "DELETE FROM dependency_plans WHERE id NOT IN (SELECT plan FROM dependency_bindings)",
            [],
        )
        .map_err(err)?;
    }
    db.execute_batch("DELETE FROM dependency_formulas WHERE plan NOT IN (SELECT id FROM dependency_plans); DELETE FROM dependency_ranges WHERE plan NOT IN (SELECT id FROM dependency_plans); DELETE FROM dependency_selected;").map_err(err)?;
    crate::shared::prune_validations(db)
}
pub fn command(root: &Path, input: &Value) -> Result<Value> {
    let mut db = database(root)?;
    schema(&db)?;
    match field(input, "method")? {
        "cache-status" => {
            let _syntax_cache = crate::formula::syntax_cache_scope();
            let version = input["version"]
                .as_u64()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing version"))?;
            let view = &input["view"];
            let plan = closest(&db, version, Some(field(input, "metadata")?))?;
            let tx = db.transaction().map_err(err)?;
            crate::calculation::snapshot(&tx, version)?;
            let index: Value =
                serde_json::from_slice(&std::fs::read(root.join("index.json")).map_err(err)?)
                    .map_err(err)?;
            let exported = input["exported"] == true;
            if exported
                && tx
                    .query_row("SELECT version FROM export_state", [], |r| {
                        r.get::<_, u64>(0)
                    })
                    .map_err(err)?
                    != version
            {
                return Ok(
                    json!({"qualified":false,"reason":"Exported checkpoint version does not match its package"}),
                );
            }
            let mut count = 0u64;
            let mut sheets = Vec::new();
            for sheet in view["sheets"]
                .as_array()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheets"))?
            {
                let path = field(sheet, "path")?;
                let name = field(sheet, "name")?;
                let before = count;
                crate::bulk::calculate_shared(&tx, version, path)?;
                let mut after = None;
                loop {
                    let page =
                        crate::calculation::page(&tx, path, (0, 0), (1048575, 16383), after, true)?;
                    for (reference, cell) in &page {
                        crate::cancel::check()?;
                        count += 1;
                        if plan.is_none() {
                            return Ok(
                                json!({"qualified":false,"reason":"No complete version-bound dependency plan"}),
                            );
                        }
                        let cell = crate::calculation::resolved_cell(&tx, path, reference, cell)?;
                        if let Err(reason) =
                            crate::formula::cache_eligible(view, name, reference, &cell)
                        {
                            crate::cancel::check()?;
                            return Ok(json!({"qualified":false,"reason":reason}));
                        }
                    }
                    if page.len() < 512 {
                        break;
                    }
                    after = Some(position(&page.last().unwrap().0)?);
                }
                let base = index["sheets"]
                    .as_array()
                    .ok_or_else(|| crate::coded("internal-missing-field", "Missing index sheets"))?
                    .iter()
                    .find(|s| s["path"] == path)
                    .ok_or_else(|| crate::coded("internal-missing-field", "Missing index sheet"))?;
                let (added, rows, columns): (u64, u64, u64) = if exported {
                    tx.prepare_cached("SELECT count(*),coalesce(max(e.r)+1,0),coalesce(max(e.c)+1,0) FROM export_changes e WHERE e.sheet=?1 AND (e.data!='{}' OR e.style IS NOT NULL) AND NOT EXISTS(SELECT 1 FROM cells c WHERE c.sheet=e.sheet AND c.r=e.r AND c.c=e.c)").map_err(crate::err)?.query_row([path], |r|Ok((r.get(0)?,r.get(1)?,r.get(2)?))).map_err(err)?
                } else {
                    (0, 0, 0)
                };
                // Export overrides are authoritative here: calculation may
                // resolve a shared follower to ordinary syntax while its XML
                // intentionally retains the original shared formula group.
                let special: u64 = if exported {
                    tx.prepare_cached("SELECT count(*) FROM cells c LEFT JOIN export_changes e ON e.sheet=c.sheet AND e.r=c.r AND e.c=c.c WHERE c.sheet=?1 AND c.formula IS NOT NULL AND json_type(c.formula_attrs,'$.t') IS NOT NULL AND json_type(e.data,'$.formula') IS NULL AND json_type(e.data,'$.value') IS NULL").map_err(crate::err)?.query_row([path], |r|r.get(0)).map_err(err)?
                } else {
                    base["specialFormulaCount"].as_u64().unwrap_or(0)
                };
                sheets.push(json!({"path":path,"storedCells":base["storedCells"].as_u64()
                    .ok_or_else(|| crate::coded("formula-malformed", "Invalid indexed cell count"))?+added,"storedFormulaCount":count-before,"specialFormulaCount":special,"extended":base["extended"],"rows":base["rows"].as_u64().unwrap_or(20).max(rows),"columns":base["columns"].as_u64().unwrap_or(8).max(columns)}));
            }
            Ok(json!({"qualified":true,"formulaCells":count,"plan":plan,"sheets":sheets}))
        }
        "invalidate" => {
            db.execute(
                "INSERT OR REPLACE INTO dependency_bindings VALUES(?1,0)",
                [input["version"]
                    .as_u64()
                    .ok_or_else(|| crate::coded("internal-missing-field", "Missing version"))?],
            )
            .map_err(err)?;
            Ok(json!({"invalidated":true}))
        }
        "create" => {
            db.execute(
                "INSERT INTO dependency_plans(metadata) VALUES(?1)",
                [field(input, "metadata")?],
            )
            .map_err(err)?;
            Ok(json!({"plan":db.last_insert_rowid()}))
        }
        "append" => {
            let plan = input["plan"]
                .as_u64()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing dependency plan"))?;
            let tx = db.transaction().map_err(err)?;
            for item in input["items"]
                .as_array()
                .filter(|a| a.len() <= 1024)
                .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid dependency page"))?
            {
                crate::cancel::check()?;
                let sheet = field(item, "sheet")?;
                let reference = field(item, "ref")?;
                let (r, c) = position(reference)?;
                tx.execute(
                    "INSERT OR IGNORE INTO dependency_formulas VALUES(?1,?2,?3,?4,?5)",
                    params![plan, sheet, reference, r, c],
                )
                .map_err(err)?;
                for dep in item["ranges"].as_array().ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing formula dependencies")
                })? {
                    let (first, last) = if let Some(reference) = dep["ref"].as_str() {
                        let p = position(reference)?;
                        (p, p)
                    } else {
                        let n = |key: &str, cap: u32| -> Result<u32> {
                            let value = dep[key].as_u64().unwrap_or(u64::from(cap));
                            if value > u64::from(cap) {
                                return Err(crate::coded(
                                    "internal-invalid-field",
                                    "Invalid dependency bounds",
                                ));
                            }
                            Ok(value as u32)
                        };
                        (
                            (n("startRow", 1048575)?, n("startColumn", 16383)?),
                            (n("endRow", 1048575)?, n("endColumn", 16383)?),
                        )
                    };
                    if first.0 > last.0 || first.1 > last.1 {
                        return Err(crate::coded(
                            "internal-invalid-field",
                            "Invalid dependency range",
                        ));
                    }
                    tx.execute(
                        "INSERT INTO dependency_ranges VALUES(?1,?2,?3,?4,?5,?6,?7,?8)",
                        params![
                            plan,
                            sheet,
                            reference,
                            field(dep, "sheet")?,
                            first.0,
                            last.0,
                            first.1,
                            last.1
                        ],
                    )
                    .map_err(err)?;
                }
            }
            tx.commit().map_err(err)?;
            Ok(json!({"stored":true}))
        }
        "bind" => {
            let plan = input["plan"]
                .as_u64()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing plan"))?;
            let version = input["version"]
                .as_u64()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing version"))?;
            let tx = db.transaction().map_err(err)?;
            tx.execute("UPDATE dependency_plans SET complete=1 WHERE id=?1", [plan])
                .map_err(err)?;
            tx.execute(
                "INSERT OR REPLACE INTO dependency_bindings VALUES(?1,?2)",
                params![version, plan],
            )
            .map_err(err)?;
            tx.commit().map_err(err)?;
            Ok(json!({"bound":true}))
        }
        "remove-formulas" | "update-formulas" => {
            let version = input["version"]
                .as_u64()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing version"))?;
            let Some(previous) = closest(&db, version, Some(field(input, "metadata")?))? else {
                return Ok(json!({"hit":false}));
            };
            let tx = db.transaction().map_err(err)?;
            crate::calculation::snapshot_changes(&tx, version, input)?;
            tx.execute_batch("CREATE TEMP TABLE updated_formulas(sheet TEXT,path TEXT,ref TEXT,r INTEGER,c INTEGER,deps TEXT,PRIMARY KEY(sheet,ref)) WITHOUT ROWID;").map_err(err)?;
            for change in input["changedRanges"]
                .as_array()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing changes"))?
            {
                crate::cancel::check()?;
                let (a, b) = bounds(field(change, "range")?)?;
                let name = field(change, "sheet")?;
                let sheet = input["view"]["sheets"]
                    .as_array()
                    .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheets"))?
                    .iter()
                    .find(|s| s["name"] == name)
                    .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheet"))?;
                let path = field(sheet, "path")?;
                tx.execute("INSERT OR IGNORE INTO updated_formulas SELECT sheet,?2,ref,r,c,NULL FROM dependency_formulas WHERE plan=?1 AND sheet=?3 AND r BETWEEN ?4 AND ?5 AND c BETWEEN ?6 AND ?7",params![previous,path,name,a.0,b.0,a.1,b.1]).map_err(err)?;
                let mut after = None;
                loop {
                    let page = crate::calculation::page(&tx, path, a, b, after, true)?;
                    for (reference, _) in &page {
                        let (r, c) = position(reference)?;
                        tx.prepare_cached(
                            "INSERT OR IGNORE INTO updated_formulas VALUES(?1,?2,?3,?4,?5,NULL)",
                        )
                        .map_err(err)?
                        .execute(params![name, path, reference, r, c])
                        .map_err(err)?;
                    }
                    if page.len() < 512 {
                        break;
                    }
                    after = Some(position(&page.last().unwrap().0)?);
                }
            }
            let mut count = 0u64;
            {
                let mut query = tx
                    .prepare_cached(
                        "SELECT sheet,path,ref FROM updated_formulas ORDER BY sheet,ref",
                    )
                    .map_err(err)?;
                let mut rows = query.query([]).map_err(err)?;
                while let Some(row) = rows.next().map_err(err)? {
                    crate::cancel::check()?;
                    count += 1;
                    let name: String = row.get(0).map_err(err)?;
                    let path: String = row.get(1).map_err(err)?;
                    let reference: String = row.get(2).map_err(err)?;
                    let cell = crate::calculation::cell(&tx, &path, &reference)?;
                    if cell.get("formulaType").is_some() || cell["scenarioResult"] == true {
                        return Ok(json!({"hit":false}));
                    }
                    if cell["formula"].is_string() {
                        if input["method"] == "remove-formulas" {
                            return Ok(json!({"hit":false}));
                        }
                        let deps = match crate::formula::cell_dependencies(
                            &input["view"],
                            &name,
                            &reference,
                            &cell,
                        ) {
                            Ok(v) => v,
                            Err(_) => return Ok(json!({"hit":false})),
                        };
                        tx.prepare_cached(
                            "UPDATE updated_formulas SET deps=?3 WHERE sheet=?1 AND ref=?2",
                        )
                        .map_err(err)?
                        .execute(params![name, reference, deps.to_string()])
                        .map_err(err)?;
                    }
                }
            }
            tx.execute(
                "INSERT INTO dependency_plans(metadata,complete) VALUES(?1,1)",
                [field(input, "metadata")?],
            )
            .map_err(err)?;
            let plan = tx.last_insert_rowid() as u64;
            tx.execute("INSERT INTO dependency_formulas SELECT ?1,sheet,ref,r,c FROM dependency_formulas f WHERE plan=?2 AND NOT EXISTS(SELECT 1 FROM updated_formulas u WHERE u.sheet=f.sheet AND u.ref=f.ref)",params![plan,previous]).map_err(err)?;
            tx.execute("INSERT INTO dependency_ranges SELECT ?1,owner_sheet,owner_ref,sheet,r0,r1,c0,c1 FROM dependency_ranges d WHERE plan=?2 AND NOT EXISTS(SELECT 1 FROM updated_formulas u WHERE u.sheet=d.owner_sheet AND u.ref=d.owner_ref)",params![plan,previous]).map_err(err)?;
            {
                let mut query = tx
                    .prepare_cached("SELECT deps FROM updated_formulas WHERE deps IS NOT NULL")
                    .map_err(err)?;
                let mut rows = query.query([]).map_err(err)?;
                while let Some(row) = rows.next().map_err(err)? {
                    crate::cancel::check()?;
                    let deps: Value = serde_json::from_str(&row.get::<_, String>(0).map_err(err)?)
                        .map_err(err)?;
                    crate::calculation::append(&tx, plan, &deps)?;
                }
            }
            tx.execute(
                "INSERT OR REPLACE INTO dependency_bindings VALUES(?1,?2)",
                params![version, plan],
            )
            .map_err(err)?;
            tx.commit().map_err(err)?;
            Ok(json!({"hit":true,"plan":plan,"changedFormulaPositions":count}))
        }
        "select" => {
            let version = input["version"]
                .as_u64()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing version"))?;
            let Some(plan) = closest(&db, version, Some(field(input, "metadata")?))? else {
                return Ok(json!({"hit":false}));
            };
            let tx = db.transaction().map_err(err)?;
            tx.execute("DELETE FROM dependency_selected", [])
                .map_err(err)?;
            tx.execute_batch("CREATE TEMP TABLE dirty(sheet TEXT,r0 INTEGER,r1 INTEGER,c0 INTEGER,c1 INTEGER); CREATE TEMP TABLE frontier(sheet TEXT,ref TEXT,r INTEGER,c INTEGER,PRIMARY KEY(sheet,ref));").map_err(err)?;
            for change in input["changedRanges"]
                .as_array()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing changed ranges"))?
            {
                let (a, b) = bounds(field(change, "range")?)?;
                tx.execute(
                    "INSERT INTO dirty VALUES(?1,?2,?3,?4,?5)",
                    params![field(change, "sheet")?, a.0, b.0, a.1, b.1],
                )
                .map_err(err)?;
            }
            tx.execute("INSERT OR IGNORE INTO frontier SELECT f.sheet,f.ref,f.r,f.c FROM dependency_formulas f JOIN dirty q ON q.sheet=f.sheet AND f.r BETWEEN q.r0 AND q.r1 AND f.c BETWEEN q.c0 AND q.c1 WHERE f.plan=?1",[plan]).map_err(err)?;
            loop {
                crate::cancel::check()?;
                tx.execute(SELECT_DEPENDENTS, [plan]).map_err(err)?;
                let added=tx.execute("INSERT OR IGNORE INTO dependency_selected SELECT ?1,sheet,ref FROM frontier",[plan]).map_err(err)?;
                if added == 0 {
                    break;
                }
                tx.execute_batch("DELETE FROM dirty; INSERT INTO dirty SELECT sheet,r,r,c,c FROM frontier; DELETE FROM frontier;").map_err(err)?;
            }
            let count: u64 = tx
                .query_row(
                    "SELECT count(*) FROM dependency_selected WHERE plan=?1",
                    [plan],
                    |r| r.get(0),
                )
                .map_err(err)?;
            tx.commit().map_err(err)?;
            Ok(json!({"hit":true,"plan":plan,"count":count}))
        }
        "page" => {
            let plan = input["plan"]
                .as_u64()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing plan"))?;
            let offset = input["offset"].as_u64().unwrap_or(0);
            let kind = field(input, "kind")?;
            let sql = if kind == "formulas" {
                "SELECT f.sheet,f.ref,f.r,f.r,f.c,f.c FROM dependency_formulas f JOIN dependency_selected s ON s.plan=f.plan AND s.sheet=f.sheet AND s.ref=f.ref WHERE f.plan=?1 ORDER BY f.sheet,f.r,f.c LIMIT 1024 OFFSET ?2"
            } else if kind == "ranges" {
                "SELECT DISTINCT d.sheet,'',d.r0,d.r1,d.c0,d.c1 FROM dependency_ranges d JOIN dependency_selected s ON s.plan=d.plan AND s.sheet=d.owner_sheet AND s.ref=d.owner_ref WHERE d.plan=?1 ORDER BY d.sheet,d.r0,d.c0,d.r1,d.c1 LIMIT 1024 OFFSET ?2"
            } else {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid dependency page kind",
                ));
            };
            let mut statement = db.prepare_cached(sql).map_err(err)?;
            let items = statement
                .query_map(params![plan, offset], |row| {
                    let sheet: String = row.get(0)?;
                    let reference: String = row.get(1)?;
                    let a: u32 = row.get(2)?;
                    let b: u32 = row.get(3)?;
                    let c: u32 = row.get(4)?;
                    let d: u32 = row.get(5)?;
                    Ok(if kind == "formulas" {
                        json!({"sheet":sheet,"ref":reference})
                    } else if a == b && c == d {
                        json!({"sheet":sheet,"ref":address(a,c)})
                    } else {
                        json!({"sheet":sheet,"startRow":a,"endRow":b,"startColumn":c,"endColumn":d})
                    })
                })
                .map_err(err)?
                .collect::<std::result::Result<Vec<_>, _>>()
                .map_err(err)?;
            Ok(
                json!({"nextOffset":if items.len()==1024{Some(offset+1024)}else{None},"items":items}),
            )
        }
        _ => Err(crate::coded(
            "internal-unknown-reference",
            "Unknown dependency operation",
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::{
        Arc,
        atomic::{AtomicU64, Ordering},
    };

    fn graph(root: &Path) -> Connection {
        let db = Connection::open(root.join("cells.sqlite")).unwrap();
        db.execute_batch("CREATE TABLE versions(id INTEGER PRIMARY KEY,parent INTEGER); INSERT INTO versions VALUES(0,NULL),(1,0); ").unwrap();
        schema(&db).unwrap();
        db.execute_batch("INSERT INTO dependency_plans VALUES(1,'qualified',1); INSERT INTO dependency_bindings VALUES(0,1);").unwrap();
        db
    }
    fn formula(
        db: &Connection,
        sheet: &str,
        reference: &str,
        row: u32,
        col: u32,
        deps: &[(&str, u32, u32, u32, u32)],
    ) {
        db.execute(
            "INSERT INTO dependency_formulas VALUES(1,?1,?2,?3,?4)",
            params![sheet, reference, row, col],
        )
        .unwrap();
        for &(target, r0, r1, c0, c1) in deps {
            db.execute(
                "INSERT INTO dependency_ranges VALUES(1,?1,?2,?3,?4,?5,?6,?7)",
                params![sheet, reference, target, r0, r1, c0, c1],
            )
            .unwrap();
        }
    }
    fn selected(db: &Connection) -> Vec<(String, String)> {
        db.prepare_cached("SELECT sheet,ref FROM dependency_selected ORDER BY sheet,ref")
            .unwrap()
            .query_map([], |r| Ok((r.get(0)?, r.get(1)?)))
            .unwrap()
            .collect::<rusqlite::Result<_>>()
            .unwrap()
    }
    #[test]
    fn analyzed_select_preserves_transitive_cycles_duplicates_cross_sheet_and_range_overlap() {
        let dir = tempfile::tempdir().unwrap();
        let db = graph(dir.path());
        formula(
            &db,
            "First",
            "B1",
            0,
            1,
            &[("First", 0, 0, 0, 0), ("First", 0, 0, 0, 0)],
        );
        formula(
            &db,
            "First",
            "C1",
            0,
            2,
            &[("First", 0, 0, 1, 1), ("Second", 0, 0, 0, 0)],
        );
        formula(&db, "Second", "A1", 0, 0, &[("First", 0, 0, 2, 2)]);
        formula(&db, "Second", "B2", 1, 1, &[("Second", 0, 3, 0, 2)]);
        formula(&db, "Third", "A1", 0, 0, &[("First", 0, 0, 0, 16383)]);
        formula(&db, "Third", "A2", 1, 0, &[("First", 0, 1048575, 0, 0)]);
        formula(&db, "Isolated", "D9", 8, 3, &[("Isolated", 0, 0, 0, 0)]);
        db.execute_batch("ANALYZE;").unwrap();
        let result=command(dir.path(),&json!({"method":"select","version":1,"metadata":"qualified","changedRanges":[{"sheet":"First","range":"A1"},{"sheet":"First","range":"A1"}]})).unwrap();
        assert_eq!(result["count"], 6);
        assert_eq!(
            selected(&db),
            vec![
                ("First".into(), "B1".into()),
                ("First".into(), "C1".into()),
                ("Second".into(), "A1".into()),
                ("Second".into(), "B2".into()),
                ("Third".into(), "A1".into()),
                ("Third".into(), "A2".into())
            ]
        );
        let unchanged = selected(&db);
        assert_eq!(command(dir.path(),&json!({"method":"select","version":1,"metadata":"foreign","changedRanges":[{"sheet":"First","range":"A1"}]})).unwrap(),json!({"hit":false}));
        assert_eq!(selected(&db), unchanged);
        let result=command(dir.path(),&json!({"method":"select","version":1,"metadata":"qualified","changedRanges":[{"sheet":"Isolated","range":"D9"}]})).unwrap();
        assert_eq!(result["count"], 1);
        assert_eq!(selected(&db), vec![("Isolated".into(), "D9".into())]);
    }
    #[test]
    fn analyzed_fanout_does_not_rescan_selected_owners_per_frontier_position() {
        let dir = tempfile::tempdir().unwrap();
        let mut db = graph(dir.path());
        {
            let tx = db.transaction().unwrap();
            for row in 0..10000 {
                formula(
                    &tx,
                    "Large",
                    &format!("D{}", row + 1),
                    row,
                    3,
                    &[("Large", 0, 0, 0, 0)],
                );
            }
            tx.commit().unwrap();
        }
        db.execute_batch("CREATE TEMP TABLE dirty(sheet TEXT,r0 INTEGER,r1 INTEGER,c0 INTEGER,c1 INTEGER); CREATE TEMP TABLE frontier(sheet TEXT,ref TEXT,r INTEGER,c INTEGER,PRIMARY KEY(sheet,ref)); INSERT INTO dirty VALUES('Large',0,0,0,0); ANALYZE;").unwrap();
        let steps = Arc::new(AtomicU64::new(0));
        let observed = steps.clone();
        db.progress_handler(
            1000,
            Some(move || observed.fetch_add(1000, Ordering::Relaxed) > 2_000_000),
        );
        db.execute(SELECT_DEPENDENTS, [1]).unwrap();
        assert_eq!(
            db.query_row::<u64, _, _>("SELECT count(*) FROM frontier", [], |r| r.get(0))
                .unwrap(),
            10000
        );
        db.execute_batch("INSERT INTO dependency_selected SELECT 1,sheet,ref FROM frontier; DELETE FROM dirty; INSERT INTO dirty SELECT sheet,r,r,c,c FROM frontier; DELETE FROM frontier;").unwrap();
        db.execute(SELECT_DEPENDENTS, [1]).unwrap();
        assert_eq!(
            db.query_row::<u64, _, _>("SELECT count(*) FROM frontier", [], |r| r.get(0))
                .unwrap(),
            0
        );
        assert!(
            steps.load(Ordering::Relaxed) < 2_000_000,
            "Instruction bound proves completed owners are skipped before frontier fan-out."
        );
    }
}
