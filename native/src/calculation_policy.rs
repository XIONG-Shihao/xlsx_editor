//! Workbook-owned scheduling policy and immutable pending masks. Graphs stay SQL/paged.
use crate::xml::{address, position};
use crate::{Result, err, field};
use rusqlite::{Connection, OptionalExtension, params};
use serde_json::{Value, json};
use std::path::Path;
fn present(db: &Connection) -> Result<bool> {
    db.prepare_cached(
        "SELECT EXISTS(SELECT 1 FROM sqlite_master WHERE name='calculation_policies')",
    )
    .map_err(crate::err)?
    .query_row([], |r| r.get(0))
    .map_err(err)
}
fn schema(db: &Connection) -> Result<()> {
    db.execute_batch("CREATE TABLE IF NOT EXISTS calculation_policies(version INTEGER PRIMARY KEY,mode TEXT NOT NULL,on_save INTEGER NOT NULL);CREATE TABLE IF NOT EXISTS calculation_pending(version INTEGER NOT NULL,sheet TEXT NOT NULL,r INTEGER NOT NULL,c INTEGER NOT NULL,pending INTEGER NOT NULL,PRIMARY KEY(version,sheet,r,c)) WITHOUT ROWID;CREATE TABLE IF NOT EXISTS calculation_graphs(plan INTEGER PRIMARY KEY,scoped INTEGER NOT NULL);CREATE TABLE IF NOT EXISTS calculation_bindings(version INTEGER PRIMARY KEY,plan INTEGER NOT NULL);CREATE TABLE IF NOT EXISTS calculation_formula_guards(plan INTEGER NOT NULL,sheet TEXT NOT NULL,ref TEXT NOT NULL,reason TEXT NOT NULL,PRIMARY KEY(plan,sheet,ref)) WITHOUT ROWID;CREATE TABLE IF NOT EXISTS calculation_outputs(plan INTEGER NOT NULL,sheet TEXT NOT NULL,r0 INTEGER,c0 INTEGER,r1 INTEGER,c1 INTEGER,mr INTEGER,mc INTEGER);CREATE INDEX IF NOT EXISTS calculation_output_sheet ON calculation_outputs(plan,sheet,r0,r1);CREATE TABLE IF NOT EXISTS calculation_object_inputs(plan INTEGER NOT NULL,object TEXT NOT NULL,sheet TEXT NOT NULL,path TEXT NOT NULL,r0 INTEGER,c0 INTEGER,r1 INTEGER,c1 INTEGER,target_sheet TEXT NOT NULL,target_path TEXT NOT NULL,PRIMARY KEY(plan,object));CREATE TABLE IF NOT EXISTS calculation_selected_objects(plan INTEGER NOT NULL,object TEXT NOT NULL,PRIMARY KEY(plan,object));CREATE TABLE IF NOT EXISTS calculation_after(plan INTEGER NOT NULL,sheet TEXT NOT NULL,ref TEXT NOT NULL,PRIMARY KEY(plan,sheet,ref));CREATE TABLE IF NOT EXISTS calculation_object_outputs(plan INTEGER NOT NULL,sheet TEXT NOT NULL,r0 INTEGER,c0 INTEGER,r1 INTEGER,c1 INTEGER,object TEXT NOT NULL);CREATE INDEX IF NOT EXISTS calculation_object_output_sheet ON calculation_object_outputs(plan,sheet,r0,r1);CREATE TABLE IF NOT EXISTS calculation_derived(version INTEGER NOT NULL,object TEXT NOT NULL,pending INTEGER NOT NULL,PRIMARY KEY(version,object)) WITHOUT ROWID;").map_err(err)
}
pub(crate) fn options(db: &Connection, version: u64) -> Result<Option<(String, bool)>> {
    if !present(db)? {
        return Ok(None);
    }
    db.prepare_cached("WITH RECURSIVE a(id,depth) AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL) SELECT p.mode,p.on_save FROM a JOIN calculation_policies p ON p.version=a.id ORDER BY a.depth LIMIT 1").map_err(crate::err)?.query_row([version],|r|Ok((r.get(0)?,r.get(1)?))).optional().map_err(err)
}
pub(crate) fn pending(db: &Connection, version: u64, sheet: &str, at: (u32, u32)) -> Result<bool> {
    if !present(db)? {
        return Ok(false);
    }
    Ok(db.prepare_cached("WITH RECURSIVE a(id,depth) AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL) SELECT p.pending FROM a JOIN calculation_pending p ON p.version=a.id WHERE p.sheet=?2 AND p.r=?3 AND p.c=?4 ORDER BY a.depth LIMIT 1").map_err(crate::err)?.query_row(params![version,sheet,at.0,at.1],|r|r.get::<_,bool>(0)).optional().map_err(err)?.unwrap_or(false))
}
pub(crate) fn derived_pending(db: &Connection, version: u64, object: &str) -> Result<bool> {
    if !present(db)? {
        return Ok(false);
    }
    Ok(db.prepare_cached("WITH RECURSIVE a(id,depth)AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL)SELECT p.pending FROM a JOIN calculation_derived p ON p.version=a.id WHERE p.object=?2 ORDER BY a.depth LIMIT 1").map_err(crate::err)?.query_row(params![version,object],|row|row.get(0)).optional().map_err(err)?.unwrap_or(false))
}
pub(crate) fn mask(
    db: &Connection,
    version: u64,
    sheet: &str,
    reference: &str,
    cell: &mut Value,
) -> Result<()> {
    let original = position(reference)?;
    let mut at = original;
    let mut output = false;
    let mut object_pending = false;
    if cell["spill"]["role"] == "follower" {
        at = position(field(&cell["spill"], "master")?)?;
        output = true;
    } else if present(db)? {
        if let Some((mr,mc,r0,c0,r1,c1))=db.prepare_cached("WITH RECURSIVE a(id,depth)AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL),p AS(SELECT b.plan FROM a JOIN calculation_bindings b ON b.version=a.id JOIN calculation_graphs g ON g.plan=b.plan ORDER BY a.depth LIMIT 1)SELECT o.mr,o.mc,o.r0,o.c0,o.r1,o.c1 FROM p JOIN calculation_outputs o ON o.plan=p.plan WHERE o.sheet=?2 AND ?3 BETWEEN o.r0 AND o.r1 AND ?4 BETWEEN o.c0 AND o.c1 LIMIT 1").map_err(crate::err)?.query_row(params![version,sheet,at.0,at.1],|row|Ok((row.get::<_,u32>(0)?,row.get::<_,u32>(1)?,row.get::<_,u32>(2)?,row.get::<_,u32>(3)?,row.get::<_,u32>(4)?,row.get::<_,u32>(5)?))).optional().map_err(err)?{
   at=(mr,mc);output=at!=original;
   // Source output identity is independent of cache scheduling. This is not
   // an owned spill proof or an authored formula on the follower. Keep all
   // stored/current fields intact; only the existing pending path masks them.
   if output&&cell.get("spill").is_none(){cell["formulaOutput"]=json!({"schema":1,"master":address(mr,mc),"range":format!("{}:{}",address(r0,c0),address(r1,c1))});}
  }
        if let Some(object)=db.prepare_cached("WITH RECURSIVE a(id,depth)AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL),p AS(SELECT b.plan FROM a JOIN calculation_bindings b ON b.version=a.id JOIN calculation_graphs g ON g.plan=b.plan ORDER BY a.depth LIMIT 1)SELECT o.object FROM p JOIN calculation_object_outputs o ON o.plan=p.plan WHERE o.sheet=?2 AND ?3 BETWEEN o.r0 AND o.r1 AND ?4 BETWEEN o.c0 AND o.c1 LIMIT 1").map_err(crate::err)?.query_row(params![version,sheet,original.0,original.1],|row|row.get::<_,String>(0)).optional().map_err(err)?{object_pending=derived_pending(db,version,&object)?;output=true;}
    }
    if !object_pending && !pending(db, version, sheet, at)? {
        return Ok(());
    }
    let mut cache = cell["calculation"].clone();
    if !cache.is_object() {
        if !cell["cachedValue"].is_null() {
            cache = if cell["dataType"] == "error" {
                json!({"status":"error","error":cell["cachedValue"]})
            } else {
                json!({"status":"calculated","value":cell["cachedValue"]})
            };
        } else if output && !cell["value"].is_null() {
            cache = if cell["dataType"] == "error" {
                json!({"status":"error","error":cell["value"]})
            } else {
                json!({"status":"calculated","value":cell["value"]})
            };
        }
    }
    if cache.is_object() {
        cache["authority"] = json!("stored");
        if cell["rawValue"].is_string() {
            cache["rawValue"] = cell["rawValue"].clone();
        }
        cell["cachedResult"] = cache;
    } else {
        cell.as_object_mut()
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid pending cell"))?
            .remove("cachedResult");
    }
    cell["calculation"] = json!({"status":"stored","pending":true,"error":"Formula or derived result is pending calculation"});
    Ok(())
}
#[cfg(test)]
mod formula_output_tests {
    use super::*;

    fn fixture() -> Connection {
        let db = Connection::open_in_memory().unwrap();
        db.execute_batch("CREATE TABLE versions(id INTEGER PRIMARY KEY,parent INTEGER);INSERT INTO versions VALUES(0,NULL),(1,0),(2,1),(3,2),(4,0);").unwrap();
        schema(&db).unwrap();
        db.execute_batch("INSERT INTO calculation_policies VALUES(0,'auto',1);INSERT INTO calculation_graphs VALUES(1,0),(2,1);INSERT INTO calculation_bindings VALUES(0,1),(3,2);INSERT INTO calculation_outputs VALUES(1,'s',0,0,1,0,0,0);INSERT INTO calculation_pending VALUES(0,'s',0,0,0),(1,'s',0,0,1),(2,'s',0,0,0);").unwrap();
        db
    }
    fn raw() -> Value {
        json!({"value":"wrld","dataType":"text","rawValue":"wrld","cachedValue":null,"styleId":7})
    }
    fn marker() -> Value {
        json!({"schema":1,"master":"A1","range":"A1:A2"})
    }

    #[test]
    fn nonpending_source_output_is_decorated_without_cache_or_storage_writes() {
        let db = fixture();
        let before = db.total_changes();
        let mut cell = raw();
        mask(&db, 0, "s", "A2", &mut cell).unwrap();
        let mut expected = raw();
        expected["formulaOutput"] = marker();
        assert_eq!(cell, expected);
        assert!(cell.get("formula").is_none());
        assert!(cell.get("formulaType").is_none());
        assert!(cell.get("spill").is_none());
        assert!(cell.get("calculation").is_none());
        assert_eq!(db.total_changes(), before);
        mask(&db, 0, "s", "A2", &mut cell).unwrap();
        assert_eq!(cell, expected);
    }

    #[test]
    fn pending_and_independently_current_results_keep_their_existing_authority() {
        let db = fixture();
        let mut current = raw();
        current["calculation"] =
            json!({"status":"calculated","value":"certified result","revision":"current"});
        let before = db.total_changes();
        let mut pending = current.clone();
        mask(&db, 1, "s", "A2", &mut pending).unwrap();
        assert_eq!(pending["formulaOutput"], marker());
        assert_eq!(pending["rawValue"], current["rawValue"]);
        assert_eq!(pending["value"], current["value"]);
        assert_eq!(pending["calculation"]["status"], "stored");
        assert_eq!(pending["calculation"]["pending"], true);
        assert!(pending["calculation"].get("value").is_none());
        assert_eq!(pending["cachedResult"]["authority"], "stored");
        assert_eq!(pending["cachedResult"]["value"], "certified result");
        let mut restored = current.clone();
        mask(&db, 2, "s", "A2", &mut restored).unwrap();
        current["formulaOutput"] = marker();
        assert_eq!(restored, current);
        assert_eq!(db.total_changes(), before);
    }

    #[test]
    fn nearest_graph_removal_branch_and_master_outside_boundaries_are_respected() {
        let db = fixture();
        for (version, reference) in [(0, "B2"), (3, "A2")] {
            let mut cell = raw();
            mask(&db, version, "s", reference, &mut cell).unwrap();
            assert_eq!(cell, raw());
        }
        let mut master = raw();
        master["formula"] = json!("IF(1,\"helo\",\"world\")");
        master["formulaType"] = json!("array");
        let expected = master.clone();
        mask(&db, 0, "s", "A1", &mut master).unwrap();
        assert_eq!(master, expected);
        for version in [0, 4] {
            let mut follower = raw();
            mask(&db, version, "s", "A2", &mut follower).unwrap();
            assert_eq!(follower["formulaOutput"], marker());
        }
        let legacy = Connection::open_in_memory().unwrap();
        let mut cell = raw();
        mask(&legacy, 0, "s", "A2", &mut cell).unwrap();
        assert_eq!(cell, raw());
    }

    #[test]
    fn owned_spill_followers_keep_complete_existing_metadata_without_duplicate_marker() {
        let db = fixture();
        let mut cell = json!({"value":2,"dataType":"number","rawValue":"2","cachedValue":null,"spill":{"schema":1,"master":"A1","range":"A1:A2","role":"follower"},"calculation":{"status":"calculated","value":2,"revision":"current"}});
        let before = cell.clone();
        mask(&db, 0, "s", "A2", &mut cell).unwrap();
        assert_eq!(cell, before);
        assert!(cell.get("formulaOutput").is_none());
        mask(&db, 1, "s", "A2", &mut cell).unwrap();
        assert_eq!(cell["spill"], before["spill"]);
        assert_eq!(cell["calculation"]["pending"], true);
        assert_eq!(cell["cachedResult"]["value"], 2);
        assert!(cell.get("formulaOutput").is_none());
    }
}

/// Stored caches are revalidated at an actual ancestor containing the same
/// formula/result/proof. A pending flag alone cannot authorize a forged cache.
pub(crate) fn cache_ancestor(
    db: &Connection,
    version: u64,
    sheet: &str,
    reference: &str,
    cell: &Value,
    marker: &str,
) -> Result<Option<u64>> {
    if !pending(db, version, sheet, position(reference)?)? {
        return Ok(None);
    }
    let at = position(reference)?;
    let mut query=db.prepare_cached("WITH RECURSIVE a(id,depth)AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL) SELECT c.version,c.data FROM a JOIN changes c ON c.version=a.id WHERE c.sheet=?2 AND c.r=?3 AND c.c=?4 ORDER BY a.depth DESC").map_err(err)?;
    let mut rows = query
        .query(params![version, sheet, at.0, at.1])
        .map_err(err)?;
    while let Some(row) = rows.next().map_err(err)? {
        let old: Value =
            serde_json::from_str(&row.get::<_, String>(1).map_err(err)?).map_err(err)?;
        let a = &old["calculation"];
        let b = &cell["calculation"];
        let value = if let (Some(a), Some(b)) = (a["value"].as_f64(), b["value"].as_f64()) {
            a.to_bits() == b.to_bits()
        } else {
            a["value"] == b["value"]
        };
        if old["formula"] == cell["formula"]
            && a.get(marker) == b.get(marker)
            && a["status"] == b["status"]
            && a["error"] == b["error"]
            && value
        {
            return Ok(Some(row.get(0).map_err(err)?));
        }
    }
    Err(crate::coded(
        "formula-bad-input",
        "Pending scalar cache has no matching accepted ancestry",
    ))
}
fn metadata(view: &Value) -> String {
    use sha2::{Digest, Sha256};
    let source = json!({"sheets":view["sheets"].as_array().unwrap().iter().map(|s|json!({"name":s["name"],"path":s["path"],"date1904":s["date1904"].as_bool().unwrap_or(false)})).collect::<Vec<_>>(),"names":view["namedRanges"],"tables":view["tables"]});
    format!("{:x}", Sha256::digest(serde_json::to_vec(&source).unwrap()))
}
pub(crate) fn graph(
    root: &Path,
    db: &Connection,
    version: u64,
    view: &Value,
) -> Result<(u64, bool)> {
    schema(db)?;
    crate::dependencies::schema(db)?;
    crate::calculation::snapshot(db, version)?;
    db.execute(
        "INSERT INTO dependency_plans(metadata,complete)VALUES(?1,1)",
        [metadata(view)],
    )
    .map_err(err)?;
    let plan = db.last_insert_rowid() as u64;
    let mut scoped = true;
    let objects = crate::package_edit::calculation_objects(root)?;
    for output in objects["outputs"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing derived outputs"))?
    {
        let (first, last) = crate::bulk::bounds(field(output, "range")?)?;
        db.execute(
            "INSERT INTO calculation_object_outputs VALUES(?1,?2,?3,?4,?5,?6,?7)",
            params![
                plan,
                field(output, "sheet")?,
                first.0,
                first.1,
                last.0,
                last.1,
                field(output, "object")?
            ],
        )
        .map_err(err)?;
    }
    for input in objects["inputs"].as_array().ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing owned object source bindings",
        )
    })? {
        let (first, last) = crate::bulk::bounds(field(input, "range")?)?;
        db.execute(
            "INSERT INTO calculation_object_inputs VALUES(?1,?2,?3,?4,?5,?6,?7,?8,?9,?10)",
            params![
                plan,
                field(input, "object")?,
                field(input, "sourceSheet")?,
                field(input, "sourcePath")?,
                first.0,
                first.1,
                last.0,
                last.1,
                field(input, "targetSheet")?,
                field(input, "targetPath")?
            ],
        )
        .map_err(err)?;
    }
    if let Ok(actual) = crate::package_edit::auditing_metadata(root) {
        let sheets = view["sheets"]
            .as_array()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing worksheets"))?
            .iter()
            .map(|s| json!({"name":s["name"],"path":s["path"]}))
            .collect::<Vec<_>>();
        if actual["sheets"] != json!(sheets)
            || !crate::package_edit::same_records(&actual["namedRanges"], &view["namedRanges"])
            || !crate::package_edit::same_records(&actual["tables"], &view["tables"])
        {
            return Err(crate::coded(
                "internal-mismatch",
                "Calculation graph metadata differs from accepted source",
            ));
        }
    } else {
        scoped = false;
    }
    for sheet in view["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing calculation worksheets"))?
    {
        let path = field(sheet, "path")?;
        let name = field(sheet, "name")?;
        crate::bulk::calculate_shared(db, version, path)?;
        let mut after = None;
        loop {
            let page = crate::calculation::page(db, path, (0, 0), (1048575, 16383), after, true)?;
            for (reference, cell) in &page {
                crate::cancel::check()?;
                let cell = crate::calculation::resolved_cell(db, path, reference, cell)?;
                let refs =
                    crate::formula::auditing::references(view, name, reference, &cell, false)?;
                if refs.unknown
                    || cell.get("formulaType").is_some_and(|kind| kind != "shared")
                    || cell.get("spill").is_some()
                    || cell["scenarioResult"] == true
                {
                    db.execute(
                        "INSERT OR REPLACE INTO calculation_formula_guards VALUES(?1,?2,?3,?4)",
                        params![
                            plan,
                            name,
                            reference,
                            "Required formula has dynamic, unknown, array or special dependencies"
                        ],
                    )
                    .map_err(err)?;
                }
                if matches!(cell["formulaType"].as_str(), Some("array" | "dataTable")) {
                    if let Some(range) = cell["formulaReference"].as_str() {
                        if let Ok((first, last)) = crate::bulk::bounds(range) {
                            let master = position(reference)?;
                            db.execute(
                                "INSERT INTO calculation_outputs VALUES(?1,?2,?3,?4,?5,?6,?7,?8)",
                                params![
                                    plan, path, first.0, first.1, last.0, last.1, master.0,
                                    master.1
                                ],
                            )
                            .map_err(err)?;
                        }
                    }
                }
                if cell["spill"]["role"] == "master" {
                    if let Some(range) = cell["spill"]["range"].as_str() {
                        let (first, last) = crate::bulk::bounds(range)?;
                        let master = position(reference)?;
                        db.execute(
                            "INSERT INTO calculation_outputs VALUES(?1,?2,?3,?4,?5,?6,?7,?8)",
                            params![
                                plan, path, first.0, first.1, last.0, last.1, master.0, master.1
                            ],
                        )
                        .map_err(err)?;
                    }
                }
                let item = match crate::formula::cell_dependencies(view, name, reference, &cell) {
                    Ok(item) => item,
                    Err(_) => {
                        db.execute(
                            "INSERT OR REPLACE INTO calculation_formula_guards VALUES(?1,?2,?3,?4)",
                            params![
                                plan,
                                name,
                                reference,
                                "Required formula dependency syntax is unqualified"
                            ],
                        )
                        .map_err(err)?;
                        json!({"sheet":name,"ref":reference,"ranges":view["sheets"].as_array().unwrap().iter().map(|s|json!({"sheet":s["name"],"startRow":0,"endRow":1048575,"startColumn":0,"endColumn":16383})).collect::<Vec<_>>()})
                    }
                };
                crate::calculation::append(db, plan, &item)?;
            }
            if page.len() < 512 {
                break;
            }
            after = Some(position(&page.last().unwrap().0)?);
        }
    }
    // A removed formula must not leave an inherited pending marker on its new
    // ordinary/blank cell. Tombstones belong to the accepted edit version.
    for sheet in view["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing worksheets"))?
    {
        db.execute("WITH RECURSIVE a(id,depth)AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL),p AS(SELECT x.*,row_number()OVER(PARTITION BY sheet,r,c ORDER BY a.depth)priority FROM a JOIN calculation_pending x ON x.version=a.id)INSERT OR REPLACE INTO calculation_pending SELECT ?1,p.sheet,p.r,p.c,0 FROM p WHERE p.priority=1 AND p.pending=1 AND p.sheet=?2 AND NOT EXISTS(SELECT 1 FROM dependency_formulas f WHERE f.plan=?3 AND f.sheet=?4 AND f.r=p.r AND f.c=p.c)",params![version,field(sheet,"path")?,plan,field(sheet,"name")?]).map_err(err)?;
    }
    db.execute(
        "INSERT OR REPLACE INTO calculation_bindings VALUES(?1,?2)",
        params![version, plan],
    )
    .map_err(err)?;
    db.execute(
        "INSERT INTO calculation_graphs VALUES(?1,?2)",
        params![plan, scoped],
    )
    .map_err(err)?;
    Ok((plan, scoped))
}
/// Expand exact scalar prerequisites in SQL, including special/derived output
/// masters. This is reused by independently qualified PivotTable source reads.
pub(crate) fn scalar_prerequisites(db: &Connection, plan: u64, view: &Value) -> Result<()> {
    loop {
        crate::cancel::check()?;
        let mut added=db.execute("INSERT OR IGNORE INTO dependency_selected SELECT ?1,f.sheet,f.ref FROM dependency_ranges d JOIN dependency_selected s ON s.plan=d.plan AND s.sheet=d.owner_sheet AND s.ref=d.owner_ref JOIN dependency_formulas f ON f.plan=d.plan AND f.sheet=d.sheet AND f.r BETWEEN d.r0 AND d.r1 AND f.c BETWEEN d.c0 AND d.c1 WHERE d.plan=?1",[plan]).map_err(err)?;
        for source in view["sheets"]
            .as_array()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing source sheets"))?
        {
            added+=db.execute("INSERT OR IGNORE INTO dependency_selected SELECT ?1,?2,f.ref FROM dependency_ranges d JOIN dependency_selected s ON s.plan=d.plan AND s.sheet=d.owner_sheet AND s.ref=d.owner_ref JOIN calculation_outputs o ON o.plan=d.plan AND o.sheet=?3 AND o.r0<=d.r1 AND o.r1>=d.r0 AND o.c0<=d.c1 AND o.c1>=d.c0 JOIN dependency_formulas f ON f.plan=o.plan AND f.sheet=?2 AND f.r=o.mr AND f.c=o.mc WHERE d.plan=?1 AND d.sheet=?2",params![plan,field(source,"name")?,field(source,"path")?]).map_err(err)?;
        }
        if added == 0 {
            break;
        }
    }
    let guard:Option<String>=db.prepare_cached("SELECT g.reason FROM calculation_formula_guards g JOIN dependency_selected s ON s.plan=g.plan AND s.sheet=g.sheet AND s.ref=g.ref WHERE g.plan=?1 LIMIT 1").unwrap().query_row([plan],|row|row.get(0)).optional().map_err(err)?;
    if let Some(reason) = guard {
        return Err(crate::coded(
            "formula-unsupported",
            format!("Required scalar prerequisite is not qualified: {reason}"),
        ));
    }
    for source in view["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing source sheets"))?
    {
        if db.prepare_cached("SELECT EXISTS(SELECT 1 FROM dependency_ranges d JOIN dependency_selected s ON s.plan=d.plan AND s.sheet=d.owner_sheet AND s.ref=d.owner_ref JOIN calculation_object_outputs o ON o.plan=d.plan AND o.sheet=?2 AND o.r0<=d.r1 AND o.r1>=d.r0 AND o.c0<=d.c1 AND o.c1>=d.c0 WHERE d.plan=?1 AND d.sheet=?3)").unwrap().query_row(params![plan,field(source,"path")?,field(source,"name")?],|row|row.get::<_,bool>(0)).map_err(err)?{return Err(crate::coded("formula-unsupported", "Required scalar input depends on an unqualified derived output"));}
    }
    Ok(())
}
fn scope_clone(db: &Connection, plan: u64, after: bool) -> Result<u64> {
    db.execute("INSERT INTO dependency_plans(metadata,complete)SELECT metadata,complete FROM dependency_plans WHERE id=?1",[plan]).map_err(err)?;
    let copied = db.last_insert_rowid() as u64;
    db.execute("INSERT INTO dependency_formulas SELECT ?1,sheet,ref,r,c FROM dependency_formulas WHERE plan=?2",params![copied,plan]).map_err(err)?;
    db.execute("INSERT INTO dependency_ranges SELECT ?1,owner_sheet,owner_ref,sheet,r0,r1,c0,c1 FROM dependency_ranges WHERE plan=?2",params![copied,plan]).map_err(err)?;
    db.execute(if after{"INSERT INTO dependency_selected SELECT ?1,s.sheet,s.ref FROM dependency_selected s JOIN calculation_after a ON a.plan=s.plan AND a.sheet=s.sheet AND a.ref=s.ref WHERE s.plan=?2"}else{"INSERT INTO dependency_selected SELECT ?1,s.sheet,s.ref FROM dependency_selected s WHERE s.plan=?2 AND NOT EXISTS(SELECT 1 FROM calculation_after a WHERE a.plan=s.plan AND a.sheet=s.sheet AND a.ref=s.ref)"},params![copied,plan]).map_err(err)?;
    Ok(copied)
}
fn selected_scope(
    db: &Connection,
    plan: u64,
    view: &Value,
    sheet: &str,
) -> Result<(u64, Option<u64>, Vec<String>)> {
    db.execute("DELETE FROM dependency_selected", [])
        .map_err(err)?;
    db.execute("INSERT INTO dependency_selected SELECT plan,sheet,ref FROM dependency_formulas WHERE plan=?1 AND sheet=?2",params![plan,sheet]).map_err(err)?;
    // Objects anchored on the selected sheet also contribute source prerequisites
    // when there is no worksheet formula reading their output.
    for source in view["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheets"))?
    {
        if source["name"] == sheet {
            db.execute("INSERT OR IGNORE INTO calculation_selected_objects SELECT plan,object FROM calculation_object_outputs WHERE plan=?1 AND sheet=?2",params![plan,field(source,"path")?]).map_err(err)?;
        }
    }
    loop {
        crate::cancel::check()?;
        let before:u64=db.prepare_cached("SELECT(SELECT count(*)FROM dependency_selected WHERE plan=?1)+(SELECT count(*)FROM calculation_selected_objects WHERE plan=?1)").unwrap().query_row([plan],|row|row.get(0)).map_err(err)?;
        db.execute("INSERT OR IGNORE INTO dependency_selected SELECT ?1,f.sheet,f.ref FROM dependency_ranges d JOIN dependency_selected s ON s.plan=d.plan AND s.sheet=d.owner_sheet AND s.ref=d.owner_ref JOIN dependency_formulas f ON f.plan=d.plan AND f.sheet=d.sheet AND f.r BETWEEN d.r0 AND d.r1 AND f.c BETWEEN d.c0 AND d.c1 WHERE d.plan=?1",[plan]).map_err(err)?;
        for source in view["sheets"]
            .as_array()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheets"))?
        {
            db.execute("INSERT OR IGNORE INTO dependency_selected SELECT ?1,?2,f.ref FROM dependency_ranges d JOIN dependency_selected s ON s.plan=d.plan AND s.sheet=d.owner_sheet AND s.ref=d.owner_ref JOIN calculation_outputs o ON o.plan=d.plan AND o.sheet=?3 AND o.r0<=d.r1 AND o.r1>=d.r0 AND o.c0<=d.c1 AND o.c1>=d.c0 JOIN dependency_formulas f ON f.plan=o.plan AND f.sheet=?2 AND f.r=o.mr AND f.c=o.mc WHERE d.plan=?1 AND d.sheet=?2",params![plan,field(source,"name")?,field(source,"path")?]).map_err(err)?;
            db.execute("INSERT OR IGNORE INTO calculation_selected_objects SELECT ?1,o.object FROM dependency_ranges d JOIN dependency_selected s ON s.plan=d.plan AND s.sheet=d.owner_sheet AND s.ref=d.owner_ref JOIN calculation_object_outputs o ON o.plan=d.plan AND o.sheet=?3 AND o.r0<=d.r1 AND o.r1>=d.r0 AND o.c0<=d.c1 AND o.c1>=d.c0 WHERE d.plan=?1 AND d.sheet=?2",params![plan,field(source,"name")?,field(source,"path")?]).map_err(err)?;
        }
        if db.prepare_cached("SELECT EXISTS(SELECT 1 FROM calculation_selected_objects o WHERE o.plan=?1 AND NOT EXISTS(SELECT 1 FROM calculation_object_inputs i WHERE i.plan=o.plan AND i.object=o.object))").unwrap().query_row([plan],|row|row.get::<_,bool>(0)).map_err(err)?{return Err(crate::coded("formula-unsupported", "Selected prerequisites require a foreign or unqualified pivot cache"));}
        db.execute("INSERT OR IGNORE INTO dependency_selected SELECT ?1,f.sheet,f.ref FROM calculation_selected_objects o JOIN calculation_object_inputs i ON i.plan=o.plan AND i.object=o.object JOIN dependency_formulas f ON f.plan=i.plan AND f.sheet=i.sheet AND f.r BETWEEN i.r0 AND i.r1 AND f.c BETWEEN i.c0 AND i.c1 WHERE o.plan=?1",[plan]).map_err(err)?;
        let after:u64=db.prepare_cached("SELECT(SELECT count(*)FROM dependency_selected WHERE plan=?1)+(SELECT count(*)FROM calculation_selected_objects WHERE plan=?1)").unwrap().query_row([plan],|row|row.get(0)).map_err(err)?;
        if before == after {
            break;
        }
    }
    let guard:Option<String>=db.prepare_cached("SELECT g.reason FROM calculation_formula_guards g JOIN dependency_selected s ON s.plan=g.plan AND s.sheet=g.sheet AND s.ref=g.ref WHERE g.plan=?1 LIMIT 1").unwrap().query_row([plan],|row|row.get(0)).optional().map_err(err)?;
    if let Some(reason) = guard {
        return Err(crate::coded(
            "formula-unsupported",
            format!("Selected prerequisite is not qualified: {reason}"),
        ));
    }
    let objects = db
        .prepare_cached(
            "SELECT object FROM calculation_selected_objects WHERE plan=?1 ORDER BY object",
        )
        .map_err(err)?
        .query_map([plan], |row| row.get(0))
        .map_err(err)?
        .collect::<std::result::Result<Vec<String>, _>>()
        .map_err(err)?;
    if objects.is_empty() {
        return Ok((plan, None, objects));
    }
    for source in view["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheets"))?
    {
        db.execute("INSERT OR IGNORE INTO calculation_after SELECT ?1,s.sheet,s.ref FROM dependency_selected s JOIN dependency_ranges d ON d.plan=s.plan AND d.owner_sheet=s.sheet AND d.owner_ref=s.ref JOIN calculation_selected_objects selected ON selected.plan=s.plan JOIN calculation_object_outputs o ON o.plan=selected.plan AND o.object=selected.object AND o.sheet=?3 AND o.r0<=d.r1 AND o.r1>=d.r0 AND o.c0<=d.c1 AND o.c1>=d.c0 WHERE s.plan=?1 AND d.sheet=?2",params![plan,field(source,"name")?,field(source,"path")?]).map_err(err)?;
    }
    loop {
        crate::cancel::check()?;
        let added=db.execute("INSERT OR IGNORE INTO calculation_after SELECT ?1,f.sheet,f.ref FROM dependency_ranges d JOIN dependency_selected f ON f.plan=d.plan AND f.sheet=d.owner_sheet AND f.ref=d.owner_ref JOIN calculation_after a ON a.plan=d.plan JOIN dependency_formulas source ON source.plan=a.plan AND source.sheet=a.sheet AND source.ref=a.ref AND source.sheet=d.sheet AND source.r BETWEEN d.r0 AND d.r1 AND source.c BETWEEN d.c0 AND d.c1 WHERE d.plan=?1",[plan]).map_err(err)?;
        if added == 0 {
            break;
        }
    }
    if db.prepare_cached("SELECT EXISTS(SELECT 1 FROM calculation_after a JOIN dependency_formulas f ON f.plan=a.plan AND f.sheet=a.sheet AND f.ref=a.ref JOIN calculation_selected_objects o ON o.plan=a.plan JOIN calculation_object_inputs i ON i.plan=o.plan AND i.object=o.object AND i.sheet=f.sheet AND f.r BETWEEN i.r0 AND i.r1 AND f.c BETWEEN i.c0 AND i.c1 WHERE a.plan=?1)").unwrap().query_row([plan],|row|row.get::<_,bool>(0)).map_err(err)?{return Err(crate::coded("formula-refused", "Pivot sources depend cyclically on selected pivot output"));}
    let source = scope_clone(db, plan, false)?;
    let after = scope_clone(db, plan, true)?;
    Ok((after, Some(source), objects))
}
fn mark(
    db: &Connection,
    version: u64,
    view: &Value,
    plan: u64,
    only_selected: bool,
    value: bool,
) -> Result<()> {
    for sheet in view["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing calculation worksheets"))?
    {
        let sql = if only_selected {
            "INSERT OR REPLACE INTO calculation_pending SELECT ?1,?2,f.r,f.c,?3 FROM dependency_formulas f JOIN dependency_selected s ON s.plan=f.plan AND s.sheet=f.sheet AND s.ref=f.ref WHERE f.plan=?4 AND f.sheet=?5"
        } else {
            "INSERT OR REPLACE INTO calculation_pending SELECT ?1,?2,r,c,?3 FROM dependency_formulas WHERE plan=?4 AND sheet=?5"
        };
        db.execute(
            sql,
            params![
                version,
                field(sheet, "path")?,
                value,
                plan,
                field(sheet, "name")?
            ],
        )
        .map_err(err)?;
    }
    Ok(())
}
fn mark_objects(root: &Path, db: &Connection, version: u64, value: bool) -> Result<()> {
    for object in crate::package_edit::calculation_objects(root)?["objects"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing calculation objects"))?
    {
        db.execute(
            "INSERT OR REPLACE INTO calculation_derived VALUES(?1,?2,?3)",
            params![version, field(object, "key")?, value],
        )
        .map_err(err)?;
    }
    Ok(())
}
pub(crate) fn derived_count(db: &Connection, version: u64) -> Result<u64> {
    if !present(db)? {
        return Ok(0);
    }
    db.prepare_cached("WITH RECURSIVE a(id,depth)AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL),p AS(SELECT x.*,row_number()OVER(PARTITION BY object ORDER BY a.depth)priority FROM a JOIN calculation_derived x ON x.version=a.id) SELECT count(*)FROM p WHERE priority=1 AND pending=1").unwrap().query_row([version],|r|r.get(0)).map_err(err)
}
pub(crate) fn pending_count(db: &Connection, version: u64) -> Result<u64> {
    if !present(db)? {
        return Ok(0);
    }
    db.prepare_cached("WITH RECURSIVE a(id,depth)AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL),p AS(SELECT x.*,row_number()OVER(PARTITION BY sheet,r,c ORDER BY a.depth)priority FROM a JOIN calculation_pending x ON x.version=a.id) SELECT count(*)FROM p WHERE priority=1 AND pending=1").unwrap().query_row([version],|r|r.get(0)).map_err(err)
}
pub(crate) fn preserving(db: &Connection, version: u64) -> Result<bool> {
    if !db
        .prepare_cached(
            "SELECT EXISTS(SELECT 1 FROM sqlite_master WHERE name='calculation_preservation')",
        )
        .unwrap()
        .query_row([], |r| r.get::<_, bool>(0))
        .map_err(err)?
    {
        return Ok(false);
    }
    db.prepare_cached("WITH RECURSIVE a(id)AS(SELECT ?1 UNION ALL SELECT v.parent FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL)SELECT EXISTS(SELECT 1 FROM a JOIN calculation_preservation p ON p.version=a.id)").unwrap().query_row([version],|r|r.get(0)).map_err(err)
}
fn status(db: &Connection, version: u64, view: &Value, default: &Value) -> Result<Value> {
    let (mode, on_save) = options(db, version)?.unwrap_or((
        field(default, "mode")?.into(),
        default["recalculateBeforeSave"].as_bool().unwrap_or(true),
    ));
    let count = pending_count(db, version)?;
    let objects = derived_count(db, version)?;
    let scoped = if present(db)? {
        db.prepare_cached("WITH RECURSIVE a(id,depth)AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL) SELECT g.scoped FROM a JOIN calculation_bindings b ON b.version=a.id JOIN calculation_graphs g ON g.plan=b.plan ORDER BY a.depth LIMIT 1").unwrap().query_row([version],|r|r.get::<_,bool>(0)).optional().map_err(err)?.unwrap_or(false)
    } else {
        false
    };
    let empty = view["sheets"].as_array().is_some_and(|sheets| {
        sheets
            .iter()
            .all(|sheet| sheet["storedFormulaCount"].as_u64() == Some(0))
    });
    let preserved = preserving(db, version)?;
    let mut result = json!({"schema":1,"mode":mode,"recalculateBeforeSave":on_save,"pendingFormulaCells":count,"pendingCacheObjects":objects,"protectedSource":default["unsupportedMode"]==true||preserved,"canRecalculate":view["readOnly"]!=true&&default["unsupportedMode"]!=true&&!preserved,"canChangeOptions":view["readOnly"]!=true&&default["unsupportedMode"]!=true&&!preserved,"canCalculateSheet":(scoped||empty)&&view["readOnly"]!=true&&default["unsupportedMode"]!=true&&!preserved,"reason":if default["unsupportedMode"]==true{json!("Workbook calculation settings use an unsupported source mode or opaque/iterative calcPr") }else if scoped||empty{Value::Null}else{json!("Selected-sheet calculation requires a complete qualified static scalar prerequisite graph")}});
    if preserved {
        result["preserveSource"] = json!(true);
        result["reason"] =
            json!("Formula results are stored; recalculate in Microsoft Excel after editing.");
    }
    Ok(result)
}
pub(crate) fn command(root: &Path, input: &Value) -> Result<Value> {
    let mut db = crate::database(root)?;
    let version = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing policy version"))?;
    if !db
        .prepare_cached("SELECT EXISTS(SELECT 1 FROM versions WHERE id=?1)")
        .unwrap()
        .query_row([version], |r| r.get::<_, bool>(0))
        .map_err(err)?
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown calculation policy version",
        ));
    }
    let view = &input["view"];
    let source = crate::package_edit::source_calculation_policy(root)?;
    let method = field(input, "method")?;
    if method == "preserve" {
        if view["readOnly"] == true {
            return Err(crate::coded(
                "formula-protected",
                "Read-only package cannot acquire preservation editing",
            ));
        }
        schema(&db)?;
        let tx = db.transaction().map_err(err)?;
        tx.execute_batch(
            "CREATE TABLE IF NOT EXISTS calculation_preservation(version INTEGER PRIMARY KEY);",
        )
        .map_err(err)?;
        tx.execute(
            "INSERT OR IGNORE INTO calculation_preservation VALUES(?1)",
            [version],
        )
        .map_err(err)?;
        tx.execute("INSERT OR REPLACE INTO calculation_pending SELECT ?1,sheet,r,c,1 FROM cells WHERE formula IS NOT NULL",[version]).map_err(err)?;
        tx.commit().map_err(err)?;
        return Ok(json!({"version":version,"policy":status(&db,version,view,&source)?}));
    }
    if method == "status" {
        return status(&db, version, view, &source);
    }
    if method == "chart-status" {
        let mut parts = Vec::new();
        for object in crate::package_edit::calculation_objects(root)?["objects"]
            .as_array()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing objects"))?
        {
            if object["kind"] == "chart"
                && object["owned"] == true
                && !derived_pending(&db, version, field(object, "key")?)?
            {
                parts.push(field(object, "key")?.to_owned());
            }
        }
        return Ok(json!({"version":version,"parts":parts}));
    }
    if method == "mask" {
        let sheet = field(input, "sheet")?;
        let mut cells = input["cells"]
            .as_array()
            .filter(|a| a.len() <= 10000)
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid pending read page"))?
            .clone();
        for pair in &mut cells {
            let reference = pair[0]
                .as_str()
                .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid pending cell"))?
                .to_owned();
            mask(&db, version, sheet, &reference, &mut pair[1])?;
        }
        return Ok(json!({"version":version,"cells":cells}));
    }
    let pivot_source = if method == "settle-derived" && input["kind"] == "pivot" {
        let before = crate::hash(&crate::local(root, "source.xlsx")?)?;
        for proof in input["proofs"]
            .as_array()
            .filter(|proofs| proofs.len() <= 1)
            .ok_or_else(|| {
                crate::coded("internal-invalid-field", "Invalid pivot calculation proofs")
            })?
        {
            crate::package_edit::qualify_calculated_pivot(root, &db, version, proof)?;
        }
        if crate::hash(&crate::local(root, "source.xlsx")?)? != before {
            return Err(crate::coded(
                "formula-stale",
                "Pivot calculation source changed during qualification",
            ));
        }
        Some(before)
    } else {
        None
    };
    if preserving(&db, version)? {
        if method == "initialize" {
            return Ok(json!({"version":version,"policy":status(&db,version,view,&source)?}));
        }
        return Err(crate::coded(
            "formula-protected",
            "Calculation is protected while source formulas are preserved",
        ));
    }
    if method == "initialize"
        && source["mode"] != "manual"
        && source["unsupportedMode"] != true
        && input.get("initialPolicy").is_none()
    {
        return status(&db, version, view, &source);
    }
    schema(&db)?;
    let tx = db.transaction().map_err(err)?;
    let inherited = options(&tx, version)?.unwrap_or((
        field(&source, "mode")?.into(),
        source["recalculateBeforeSave"].as_bool().unwrap_or(true),
    ));
    let mut accepted = version;
    let mut mode = inherited.0.clone();
    let mut on_save = inherited.1;
    if method == "initialize" {
        if let Some(p) = input.get("initialPolicy") {
            let keys = p
                .as_object()
                .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid initial policy"))?
                .keys()
                .map(String::as_str)
                .collect::<std::collections::BTreeSet<_>>();
            if keys != ["mode", "recalculateBeforeSave"].into() {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid initial policy fields",
                ));
            }
            let requested = field(p, "mode")?;
            if !["auto", "manual"].contains(&requested) {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid initial calculation mode",
                ));
            }
            let requested_save = p["recalculateBeforeSave"].as_bool().ok_or_else(|| {
                crate::coded(
                    "internal-invalid-field",
                    "Invalid initial Save calculation policy",
                )
            })?;
            if source["unsupportedMode"] != true {
                mode = requested.into();
                on_save = requested_save;
            }
        }
    }
    if source["unsupportedMode"] == true && !["initialize", "defer"].contains(&method) {
        return Err(crate::coded(
            "formula-protected",
            "Workbook calculation source settings are protected",
        ));
    }
    if method == "fork" {
        tx.execute("INSERT INTO versions(parent)VALUES(?1)", [version])
            .map_err(err)?;
        accepted = tx.last_insert_rowid() as u64;
    }
    if method == "options" {
        if view["readOnly"] == true || source["unsupportedMode"] == true {
            return Err(crate::coded(
                "formula-protected",
                "Calculation options are read-only or use a protected source mode",
            ));
        }
        mode = field(input, "mode")?.into();
        if !["auto", "manual"].contains(&mode.as_str()) {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid calculation mode",
            ));
        }
        on_save = if let Some(v) = input.get("recalculateBeforeSave") {
            v.as_bool().ok_or_else(|| {
                crate::coded(
                    "internal-invalid-field",
                    "Invalid recalculation before Save",
                )
            })?
        } else {
            if mode == "manual" && inherited.0 != "manual" {
                true
            } else {
                inherited.1
            }
        };
        tx.execute("INSERT INTO versions(parent)VALUES(?1)", [version])
            .map_err(err)?;
        accepted = tx.last_insert_rowid() as u64;
    }
    tx.execute(
        "INSERT OR REPLACE INTO calculation_policies VALUES(?1,?2,?3)",
        params![accepted, mode, on_save],
    )
    .map_err(err)?;
    let (plan, scoped) = graph(root, &tx, accepted, view)?;
    match method {
        "initialize" => {
            if mode == "manual" || source["unsupportedMode"] == true {
                mark(&tx, accepted, view, plan, false, true)?;
                mark_objects(root, &tx, accepted, true)?;
            }
        }
        "defer" => {
            mark_objects(root, &tx, accepted, true)?;
            if scoped && input["changedRanges"].is_array() {
                let mut dirty = input["changedRanges"].as_array().unwrap().clone();
                for worksheet in view["sheets"]
                    .as_array()
                    .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheets"))?
                {
                    let path = field(worksheet, "path")?;
                    let mut query=tx.prepare_cached("SELECT r0,c0,r1,c1 FROM calculation_object_outputs WHERE plan=?1 AND sheet=?2").map_err(err)?;
                    let outputs = query
                        .query_map(params![plan, path], |row| {
                            Ok((
                                row.get::<_, u32>(0)?,
                                row.get::<_, u32>(1)?,
                                row.get::<_, u32>(2)?,
                                row.get::<_, u32>(3)?,
                            ))
                        })
                        .map_err(err)?;
                    for output in outputs {
                        let (a, b, c, d) = output.map_err(err)?;
                        dirty.push(json!({"sheet":worksheet["name"],"range":format!("{}:{}",crate::xml::address(a,b),crate::xml::address(c,d))}));
                    }
                }
                crate::dependencies::select_plan(&tx, plan, &json!(dirty))?;
                mark(&tx, accepted, view, plan, true, true)?;
            } else {
                mark(&tx, accepted, view, plan, false, true)?;
            }
        }
        "options" | "fork" => {
            if method == "options" && mode == "manual" && inherited.0 != "manual" {
                mark_objects(root, &tx, accepted, true)?;
            }
        }
        "settle-derived" => {
            let kind = field(input, "kind")?;
            if !["chart", "pivot"].contains(&kind) {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid derived calculation kind",
                ));
            }
            let selected = input
                .get("sheet")
                .map(|sheet| {
                    sheet.as_str().ok_or_else(|| {
                        crate::coded("internal-invalid-field", "Invalid derived worksheet")
                    })
                })
                .transpose()?;
            if selected.is_some_and(|sheet| {
                !view["sheets"]
                    .as_array()
                    .is_some_and(|sheets| sheets.iter().any(|entry| entry["name"] == sheet))
            }) {
                return Err(crate::coded(
                    "internal-unknown-reference",
                    "Unknown derived worksheet",
                ));
            }
            let objects = crate::package_edit::calculation_objects(root)?;
            let expected = if kind == "pivot" && selected.is_some() {
                let (_, _, pivots) = selected_scope(&tx, plan, view, selected.unwrap())?;
                pivots
                    .into_iter()
                    .collect::<std::collections::BTreeSet<_>>()
            } else {
                objects["objects"]
                    .as_array()
                    .ok_or_else(|| crate::coded("internal-missing-field", "Missing objects"))?
                    .iter()
                    .filter(|object| {
                        object["kind"] == kind
                            && selected.is_none_or(|sheet| object["sheet"] == sheet)
                    })
                    .map(|object| field(object, "key").map(str::to_owned))
                    .collect::<Result<std::collections::BTreeSet<_>>>()?
            };
            if kind == "pivot" {
                let proofs = input["proofs"]
                    .as_array()
                    .filter(|proofs| proofs.len() <= 1)
                    .ok_or_else(|| {
                        crate::coded("internal-invalid-field", "Invalid pivot calculation proofs")
                    })?;
                let mut proved = std::collections::BTreeSet::new();
                for proof in proofs {
                    if !proved.insert(field(&proof["parts"], "table")?.to_owned()) {
                        return Err(crate::coded("formula-duplicate", "Duplicate pivot proof"));
                    }
                }
                if proved != expected {
                    return Err(crate::coded(
                        "formula-unsupported",
                        "Incomplete pivot cache calculation proof",
                    ));
                }
            } else {
                let proof = &input["proof"];
                let proved = crate::export::checked_chart_caches(root, &tx, accepted, proof)?
                    .into_keys()
                    .collect::<std::collections::BTreeSet<_>>();
                if proved != expected {
                    return Err(crate::coded(
                        "formula-unsupported",
                        "Incomplete chart cache calculation proof",
                    ));
                }
            }
            for object in expected {
                tx.execute(
                    "INSERT OR REPLACE INTO calculation_derived VALUES(?1,?2,0)",
                    params![accepted, object],
                )
                .map_err(err)?;
            }
        }
        "settle" => {
            if input["sheet"].is_string() {
                let sheet = field(input, "sheet")?;
                let selected = input["plan"]
                    .as_u64()
                    .ok_or_else(|| crate::coded("internal-missing-field", "Missing scoped plan"))?;
                if !scoped {
                    return Err(crate::coded(
                        "formula-unsupported",
                        "Scoped graph is incomplete",
                    ));
                }
                let _ = sheet;
                mark(&tx, accepted, view, selected, true, false)?;
                tx.execute(
                    "INSERT OR REPLACE INTO dependency_bindings VALUES(?1,0)",
                    [accepted],
                )
                .map_err(err)?;
            } else {
                mark(&tx, accepted, view, plan, false, false)?;
            }
        }
        "scope" => {
            if !scoped {
                return Err(crate::coded(
                    "formula-unsupported",
                    "Selected-sheet graph is incomplete",
                ));
            }
            let sheet = field(input, "sheet")?;
            if !view["sheets"]
                .as_array()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheets"))?
                .iter()
                .any(|s| s["name"] == sheet)
            {
                return Err(crate::coded(
                    "internal-unknown-reference",
                    "Unknown calculation worksheet",
                ));
            }
            let (selected, source_plan, pivots) = selected_scope(&tx, plan, view, sheet)?;
            tx.commit().map_err(err)?;
            return Ok(
                json!({"version":accepted,"plan":selected,"sourcePlan":source_plan,"pivots":pivots,"policy":status(&db,accepted,view,&source)?}),
            );
        }
        _ => {
            return Err(crate::coded(
                "internal-unknown-reference",
                "Unknown calculation policy method",
            ));
        }
    }
    if let Some(before) = pivot_source {
        if crate::hash(&crate::local(root, "source.xlsx")?)? != before {
            return Err(crate::coded(
                "formula-stale",
                "Pivot calculation source changed before settlement",
            ));
        }
    }
    tx.commit().map_err(err)?;
    Ok(json!({"version":accepted,"policy":status(&db,accepted,view,&source)?}))
}
pub(crate) fn prune(db: &Connection) -> Result<()> {
    if !present(db)? {
        return Ok(());
    }
    if db
        .prepare_cached(
            "SELECT EXISTS(SELECT 1 FROM sqlite_master WHERE name='calculation_preservation')",
        )
        .unwrap()
        .query_row([], |r| r.get::<_, bool>(0))
        .map_err(err)?
    {
        db.execute(
            "DELETE FROM calculation_preservation WHERE version NOT IN(SELECT id FROM versions)",
            [],
        )
        .map_err(err)?;
    }
    db.execute_batch("DELETE FROM calculation_policies WHERE version NOT IN(SELECT id FROM versions);DELETE FROM calculation_bindings WHERE version NOT IN(SELECT id FROM versions);DELETE FROM calculation_pending WHERE version NOT IN(SELECT id FROM versions);DELETE FROM calculation_derived WHERE version NOT IN(SELECT id FROM versions);DELETE FROM calculation_outputs WHERE plan NOT IN(SELECT id FROM dependency_plans);DELETE FROM calculation_object_outputs WHERE plan NOT IN(SELECT id FROM dependency_plans);DELETE FROM calculation_object_inputs WHERE plan NOT IN(SELECT id FROM dependency_plans);DELETE FROM calculation_selected_objects WHERE plan NOT IN(SELECT id FROM dependency_plans);DELETE FROM calculation_after WHERE plan NOT IN(SELECT id FROM dependency_plans);DELETE FROM calculation_formula_guards WHERE plan NOT IN(SELECT id FROM dependency_plans);DELETE FROM calculation_graphs WHERE plan NOT IN(SELECT id FROM dependency_plans);").map_err(err)
}
pub(crate) fn retain(db: &Connection, keep: &[u64]) -> Result<()> {
    if !present(db)? {
        return Ok(());
    }
    for version in keep {
        if preserving(db, *version)? {
            db.execute(
                "INSERT OR IGNORE INTO calculation_preservation VALUES(?1)",
                [version],
            )
            .map_err(err)?;
        }
        let binding:Option<u64>=db.prepare_cached("WITH RECURSIVE a(id,depth)AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL)SELECT b.plan FROM a JOIN calculation_bindings b ON b.version=a.id ORDER BY a.depth LIMIT 1").unwrap().query_row([version],|row|row.get(0)).optional().map_err(err)?;
        if let Some(plan) = binding {
            db.execute(
                "INSERT OR REPLACE INTO calculation_bindings VALUES(?1,?2)",
                params![version, plan],
            )
            .map_err(err)?;
        }
        if let Some((mode, on_save)) = options(db, *version)? {
            db.execute(
                "INSERT OR REPLACE INTO calculation_policies VALUES(?1,?2,?3)",
                params![version, mode, on_save],
            )
            .map_err(err)?;
        }
        db.execute("WITH RECURSIVE a(id,depth)AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL),p AS(SELECT x.*,row_number()OVER(PARTITION BY sheet,r,c ORDER BY a.depth)priority FROM a JOIN calculation_pending x ON x.version=a.id) INSERT OR REPLACE INTO calculation_pending SELECT ?1,sheet,r,c,pending FROM p WHERE priority=1",[version]).map_err(err)?;
        db.execute("WITH RECURSIVE a(id,depth)AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL),p AS(SELECT x.*,row_number()OVER(PARTITION BY object ORDER BY a.depth)priority FROM a JOIN calculation_derived x ON x.version=a.id)INSERT OR REPLACE INTO calculation_derived SELECT ?1,object,pending FROM p WHERE priority=1",[version]).map_err(err)?;
    }
    Ok(())
}
