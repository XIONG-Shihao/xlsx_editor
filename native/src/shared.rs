//! Versioned, compact normalization intent. Only actual edits occupy `changes`;
//! unchanged followers translate on read and materialize only for export.
use crate::{Result, err};
use rusqlite::{Connection, OptionalExtension, params};
use serde_json::{Value, json};
use std::{cell::RefCell, collections::HashMap};

type Group = (u32, u32, String, bool);
#[derive(Default)]
struct Context {
    connection: usize,
    version: u64,
    present: bool,
    bytes: usize,
    groups: HashMap<(String, String), Option<Group>>,
    proofs: HashMap<(String, String), bool>,
}
thread_local! { static CONTEXT: RefCell<Context> = RefCell::new(Context::default()); }
pub(crate) struct Scope;
pub(crate) fn begin() -> Scope {
    CONTEXT.with(|c| *c.borrow_mut() = Context::default());
    Scope
}
impl Drop for Scope {
    fn drop(&mut self) {
        CONTEXT.with(|c| *c.borrow_mut() = Context::default());
    }
}

pub(crate) fn schema(db: &Connection) -> Result<()> {
    db.execute_batch("CREATE TABLE IF NOT EXISTS shared_normalizations(version INTEGER,sheet TEXT,group_id TEXT,master_r INTEGER,master_c INTEGER,formula TEXT,invariant INTEGER,PRIMARY KEY(version,sheet,group_id)) WITHOUT ROWID;").map_err(err)
}
fn validation_table(db: &Connection) -> Result<bool> {
    db.prepare_cached("SELECT EXISTS(SELECT 1 FROM sqlite_master WHERE type='table' AND name='shared_validations')").map_err(crate::err)?.query_row([], |row| row.get(0)).map_err(err)
}
pub(crate) fn store_validations(db: &Connection, version: u64) -> Result<()> {
    let present: bool = db
        .query_row("SELECT EXISTS(SELECT 1 FROM calc_shared)", [], |row| {
            row.get(0)
        })
        .map_err(err)?;
    if !present {
        return Ok(());
    }
    // Literal/ordinary-formula books must not pay for an unused disk table,
    // especially when the caller deliberately chooses a small session quota.
    db.execute_batch("CREATE TABLE IF NOT EXISTS shared_validations(version INTEGER,sheet TEXT,group_id TEXT,r0 INTEGER,c0 INTEGER,r1 INTEGER,c1 INTEGER,master_r INTEGER,master_c INTEGER,formula TEXT,invariant INTEGER,PRIMARY KEY(version,sheet,group_id)) WITHOUT ROWID;").map_err(err)?;
    let proved = db.execute("INSERT INTO shared_validations SELECT ?1,sheet,group_id,r0,c0,r1,c1,master_r,master_c,formula,invariant FROM calc_shared", [version]).map_err(err)?;
    crate::diagnostics::count("shared.validation_proofs_stored", proved as u64);
    Ok(())
}
pub(crate) fn prune_validations(db: &Connection) -> Result<()> {
    if validation_table(db)? {
        db.execute(
            "DELETE FROM shared_validations WHERE version NOT IN(SELECT id FROM versions)",
            [],
        )
        .map_err(err)?;
    }
    Ok(())
}

/// A successful calculation validated every member of these groups. Its result
/// version only adds derived values, so the exact immutable version can reuse
/// the shape/translation-budget proof when normalizing a contained edit range.
/// Never infer that proof from an ancestor or from arbitrary cached cell values.
pub(crate) fn normalize_validated(
    db: &Connection,
    parent: u64,
    sheet: &str,
    target: crate::bulk::Bounds,
    partial_only: bool,
) -> Result<Option<u64>> {
    crate::cancel::check()?;
    if !validation_table(db)? {
        return Ok(None);
    }
    let proof: Option<(String, u32, u32, u32, u32)> = db
        .prepare_cached("SELECT group_id,r0,c0,r1,c1 FROM shared_validations WHERE version=?1 AND sheet=?2 AND r0<=?3 AND c0<=?4 AND r1>=?5 AND c1>=?6")
        .map_err(err)?
        .query_row(params![parent, sheet, target.0.0, target.0.1, target.1.0, target.1.1], |r| Ok((r.get(0)?, r.get(1)?, r.get(2)?, r.get(3)?, r.get(4)?)))
        .optional()
        .map_err(err)?;
    let Some((group, r0, c0, r1, c1)) = proof else {
        return Ok(None);
    };
    // Group ranges may overlap. The shortcut below settles one group only, so it
    // applies only when no other group reaches into the target.
    let reaching: u32 = db
        .prepare_cached("SELECT count(*) FROM shared_validations WHERE version=?1 AND sheet=?2 AND r0<=?5 AND c0<=?6 AND r1>=?3 AND c1>=?4")
        .map_err(err)?
        .query_row(params![parent, sheet, target.0.0, target.0.1, target.1.0, target.1.1], |r| r.get(0))
        .map_err(err)?;
    if reaching != 1 {
        return Ok(None);
    }
    if partial_only && target == ((r0, c0), (r1, c1)) {
        return Ok(Some(parent));
    }
    crate::cancel::check()?;
    db.execute("INSERT INTO versions(parent) VALUES(?1)", [parent])
        .map_err(err)?;
    let version = db.last_insert_rowid() as u64;
    db.execute("INSERT INTO shared_normalizations SELECT ?1,sheet,group_id,master_r,master_c,formula,invariant FROM shared_validations WHERE version=?2 AND sheet=?3 AND group_id=?4", params![version,parent,sheet,group]).map_err(err)?;
    crate::diagnostics::count("shared.validation_proof_hits", 1);
    crate::diagnostics::count(
        "shared.validation_reused_members",
        u64::from(r1 - r0 + 1) * u64::from(c1 - c0 + 1),
    );
    crate::diagnostics::count("shared.normalized_groups", 1);
    Ok(Some(version))
}

pub(crate) fn prepare(db: &Connection, version: u64) -> Result<()> {
    // Range reads call this once per request, and some callers read one cell at
    // a time. Ask the connection's own temporary table first: a connection that
    // has none is new, whatever address it happens to occupy.
    // The tables also vanish when the transaction that created them rolls back,
    // so a failed probe means "create them", not an error.
    let known = db
        .prepare_cached("SELECT version FROM shared_context")
        .and_then(|mut statement| statement.query_row([], |r| r.get::<_, u64>(0)).optional())
        .ok();
    let fresh = known.is_none();
    if fresh {
        db.execute_batch("CREATE TEMP TABLE IF NOT EXISTS shared_context(version INTEGER); CREATE TEMP TABLE IF NOT EXISTS shared_active(sheet TEXT,group_id TEXT,master_r INTEGER,master_c INTEGER,formula TEXT,invariant INTEGER,PRIMARY KEY(sheet,group_id)) WITHOUT ROWID;").map_err(err)?;
    }
    let current: Option<u64> = known.flatten();
    let refilled = current != Some(version);
    if current != Some(version) {
        db.execute_batch("DELETE FROM shared_active; DELETE FROM shared_context;")
            .map_err(err)?;
        db.execute("WITH RECURSIVE ancestry(id,depth) AS (SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN ancestry a ON v.id=a.id WHERE v.parent IS NOT NULL), ranked AS (SELECT g.*,row_number() OVER(PARTITION BY g.sheet,g.group_id ORDER BY a.depth) priority FROM shared_normalizations g JOIN ancestry a ON g.version=a.id) INSERT INTO shared_active SELECT sheet,group_id,master_r,master_c,formula,invariant FROM ranked WHERE priority=1", [version]).map_err(err)?;
        db.execute("INSERT INTO shared_context VALUES(?1)", [version])
            .map_err(err)?;
    }
    let connection = db as *const Connection as usize;
    let changed = fresh
        || refilled
        || CONTEXT.with(|c| {
            let c = c.borrow();
            c.connection != connection || c.version != version
        });
    if changed {
        let present = db
            .query_row("SELECT EXISTS(SELECT 1 FROM shared_active)", [], |r| {
                r.get(0)
            })
            .map_err(err)?;
        CONTEXT.with(|c| {
            *c.borrow_mut() = Context {
                connection,
                version,
                present,
                ..Context::default()
            }
        });
    }
    Ok(())
}

pub(crate) fn resolve(db: &Connection, sheet: &str, reference: &str, cell: Value) -> Result<Value> {
    if cell["formulaType"] != "shared" {
        return Ok(cell);
    }
    let connection = db as *const Connection as usize;
    if CONTEXT.with(|c| {
        let c = c.borrow();
        c.connection == connection && !c.present
    }) {
        return Ok(cell);
    }
    let key = (sheet.to_owned(), cell["sharedFormulaId"].to_string());
    let cached = CONTEXT.with(|c| {
        let c = c.borrow();
        (c.connection == connection)
            .then(|| c.groups.get(&key).cloned())
            .flatten()
    });
    let group = if let Some(group) = cached {
        group
    } else {
        let group:Option<Group>=db.prepare_cached("SELECT master_r,master_c,formula,invariant FROM shared_active WHERE sheet=?1 AND group_id=?2").map_err(err)?
            .query_row(params![key.0,key.1], |r| Ok((r.get(0)?,r.get(1)?,r.get(2)?,r.get(3)?))).optional().map_err(err)?;
        let cost = key.0.len() + key.1.len() + group.as_ref().map_or(0, |g| g.2.len()) + 192;
        CONTEXT.with(|c| {
            let mut c = c.borrow_mut();
            if c.connection == connection
                && c.groups.len() + c.proofs.len() < 64
                && c.bytes + cost <= 65536
            {
                c.groups.insert(key, group.clone());
                c.bytes += cost;
            }
        });
        group
    };
    let Some((mr, mc, formula, invariant)) = group else {
        return Ok(cell);
    };
    crate::cancel::check()?;
    let (r, c) = crate::xml::position(reference)?;
    let translated = if invariant {
        json!(formula)
    } else {
        crate::formula::rewrite(&json!({"items":[{"formula":formula,"copyRows":i64::from(r)-i64::from(mr),"copyColumns":i64::from(c)-i64::from(mc)}]}))?["formulas"][0].clone()
    };
    let mut cell = cell;
    let map = cell
        .as_object_mut()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid shared cell"))?;
    for key in [
        "formulaType",
        "formulaReference",
        "sharedFormulaId",
        "scenarioResult",
        "error",
    ] {
        map.remove(key);
    }
    if cell["calculation"]["status"] != "calculated" {
        cell.as_object_mut().unwrap().remove("calculation");
    }
    // Match the previous materialized normalization's authored value semantics.
    cell["value"] = Value::Null;
    cell["rawValue"] = Value::Null;
    cell["cachedValue"] = Value::Null;
    cell["dataType"] = json!("blank");
    cell["formula"] = translated;
    Ok(cell)
}

/// Compact normalization is a completed validation proof, never inferred from
/// a cached cell value. A caller may inspect this before translating text.
pub(crate) fn normalized(db: &Connection, sheet: &str, cell: &Value) -> Result<bool> {
    if cell["formulaType"] != "shared" {
        return Ok(false);
    }
    let connection = db as *const Connection as usize;
    if CONTEXT.with(|c| {
        let c = c.borrow();
        c.connection == connection && !c.present
    }) {
        return Ok(false);
    }
    let key = (sheet.to_owned(), cell["sharedFormulaId"].to_string());
    if let Some(proved) = CONTEXT.with(|c| {
        let c = c.borrow();
        if c.connection != connection {
            return None;
        }
        c.proofs
            .get(&key)
            .copied()
            .or_else(|| c.groups.get(&key).map(Option::is_some))
    }) {
        return Ok(proved);
    }
    let proved = db
        .prepare_cached("SELECT EXISTS(SELECT 1 FROM shared_active WHERE sheet=?1 AND group_id=?2)")
        .map_err(err)?
        .query_row(params![key.0, key.1], |r| r.get(0))
        .map_err(err)?;
    crate::diagnostics::count("shared.proof_queries", 1);
    let cost = key.0.len() + key.1.len() + 128;
    CONTEXT.with(|c| {
        let mut c = c.borrow_mut();
        if c.connection == connection
            && c.groups.len() + c.proofs.len() < 64
            && c.bytes + cost <= 65536
        {
            c.proofs.insert(key, proved);
            c.bytes += cost;
        }
    });
    Ok(proved)
}

/// Add untouched normalized members to export's disk-backed effective table.
/// Existing edits and calculation results always win. No cell map enters JS.
pub(crate) fn export_members(db: &Connection) -> Result<()> {
    let present: bool = db
        .query_row("SELECT EXISTS(SELECT 1 FROM shared_active)", [], |r| {
            r.get(0)
        })
        .map_err(err)?;
    if !present {
        return Ok(());
    }
    let mut query = db.prepare_cached("SELECT c.r,c.c,c.kind,c.raw,c.inline,c.formula,c.formula_attrs,c.style,s.value,NULL,c.sheet FROM cells c JOIN shared_active g ON g.sheet=c.sheet AND g.group_id=json_quote(json_extract(c.formula_attrs,'$.si')) LEFT JOIN strings s ON c.kind='s' AND s.id=CAST(c.raw AS INTEGER) WHERE c.formula IS NOT NULL AND json_extract(c.formula_attrs,'$.t')='shared' AND NOT EXISTS(SELECT 1 FROM effective e WHERE e.sheet=c.sheet AND e.r=c.r AND e.c=c.c)").map_err(err)?;
    let mut rows = query.query([]).map_err(err)?;
    while let Some(row) = rows.next().map_err(err)? {
        crate::cancel::check()?;
        let sheet: String = row.get(10).map_err(err)?;
        let r: u32 = row.get(0).map_err(err)?;
        let c: u32 = row.get(1).map_err(err)?;
        let (reference, cell) = crate::range::decode_cell(row)?;
        let cell = resolve(db, &sheet, &reference, cell)?;
        db.prepare_cached("INSERT INTO effective VALUES(?1,?2,?3,?4,'{}')")
            .map_err(err)?
            .execute(params![sheet, r, c, cell.to_string()])
            .map_err(err)?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    fn fixture(formula: &str) -> Connection {
        let db = Connection::open_in_memory().unwrap();
        schema(&db).unwrap();
        db.execute_batch("CREATE TABLE versions(id INTEGER PRIMARY KEY,parent INTEGER); INSERT INTO versions VALUES(0,NULL),(1,0),(2,1);").unwrap();
        db.execute(
            "INSERT INTO shared_normalizations VALUES(1,'sheet','\"0\"',0,0,?1,1)",
            [formula],
        )
        .unwrap();
        db
    }
    #[test]
    fn compact_group_cache_is_connection_version_and_command_scoped() {
        let a = fixture("1+1");
        let b = fixture("2+2");
        let cell = json!({"formulaType":"shared","sharedFormulaId":"0","value":99,"calculation":{"status":"calculated","value":2}});
        {
            let _scope = begin();
            prepare(&a, 1).unwrap();
            assert_eq!(
                resolve(&a, "sheet", "A5", cell.clone()).unwrap()["formula"],
                "1+1"
            );
            prepare(&a, 0).unwrap();
            assert_eq!(resolve(&a, "sheet", "A5", cell.clone()).unwrap(), cell);
            prepare(&a, 2).unwrap();
            assert_eq!(
                resolve(&a, "sheet", "A5", cell.clone()).unwrap()["formula"],
                "1+1"
            );
            prepare(&b, 1).unwrap();
            assert_eq!(
                resolve(&b, "sheet", "A5", cell.clone()).unwrap()["formula"],
                "2+2"
            );
            assert_eq!(
                resolve(&a, "sheet", "A5", cell.clone()).unwrap()["formula"],
                "1+1"
            );
        }
        CONTEXT.with(|c| assert!(c.borrow().groups.is_empty()));
    }
    #[test]
    fn compact_group_cache_stays_within_entry_and_byte_budgets() {
        let _scope = begin();
        let db = fixture("1+1");
        for i in 1..200 {
            db.execute(
                "INSERT INTO shared_normalizations VALUES(1,'sheet',?1,0,0,?2,1)",
                params![
                    json!(i.to_string()).to_string(),
                    format!("\"{}\"", "x".repeat(2048))
                ],
            )
            .unwrap();
        }
        prepare(&db, 1).unwrap();
        for i in 0..200 {
            resolve(
                &db,
                "sheet",
                "A5",
                json!({"formulaType":"shared","sharedFormulaId":i.to_string()}),
            )
            .unwrap();
        }
        CONTEXT.with(|c| {
            let c = c.borrow();
            assert!(c.groups.len() <= 64);
            assert!(c.bytes <= 65536);
            assert!(c.groups.len() < 64);
        });
    }
}
