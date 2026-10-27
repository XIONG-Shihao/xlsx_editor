//! Keep public undo versions; collapse private calculation/batch intermediates
//! and discard abandoned branches. All changes remain in one transaction.
use crate::{Result, database, err};
use rusqlite::params;
use serde_json::{Value, json};
use std::{
    collections::{BTreeMap, BTreeSet},
    path::Path,
};
fn validate_versions(keep: &BTreeSet<u64>, versions: &BTreeMap<u64, Option<u64>>) -> Result<()> {
    if keep.iter().any(|id| !versions.contains_key(id)) {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown retained version",
        ));
    }
    let mut complete = BTreeSet::new();
    for start in versions.keys() {
        let mut path = BTreeSet::new();
        let mut current = Some(*start);
        while let Some(id) = current {
            if complete.contains(&id) {
                break;
            }
            if !path.insert(id) {
                return Err(crate::coded(
                    "workbook-malformed",
                    "Cyclic workbook versions",
                ));
            }
            current = *versions.get(&id).ok_or_else(|| {
                crate::coded("internal-unknown-reference", "Unknown workbook ancestor")
            })?;
        }
        complete.extend(path);
    }
    Ok(())
}
/// Copy each skipped path directly into its retained child. Closest cumulative
/// data/export intent wins; no accumulated overlay travels through intermediates.
fn retain_nodes(
    tx: &rusqlite::Transaction<'_>,
    keep: &BTreeSet<u64>,
    versions: &BTreeMap<u64, Option<u64>>,
    spills: bool,
    mut check: impl FnMut() -> Result<()>,
) -> Result<()> {
    validate_versions(keep, versions)?;
    for child in keep {
        let mut parent = versions[child];
        let mut seen = BTreeSet::new();
        while let Some(id) = parent {
            if keep.contains(&id) {
                break;
            }
            if !seen.insert(id) {
                return Err(crate::coded(
                    "workbook-malformed",
                    "Cyclic workbook versions",
                ));
            }
            check()?;
            if spills {
                tx.execute("INSERT OR IGNORE INTO spill_groups SELECT ?1,sheet,mr,mc,active,r0,c0,r1,c1,formula,proof,schema FROM spill_groups WHERE version=?2",params![child,id]).map_err(err)?;
            }
            tx.execute("INSERT OR IGNORE INTO shared_normalizations SELECT ?1,sheet,group_id,master_r,master_c,formula,invariant FROM shared_normalizations WHERE version=?2",params![child,id]).map_err(err)?;
            tx.execute("INSERT OR IGNORE INTO changes(version,sheet,r,c,data,edit) SELECT ?1,sheet,r,c,data,edit FROM changes WHERE version=?2",params![child,id]).map_err(err)?;
            parent = *versions.get(&id).ok_or_else(|| {
                crate::coded("internal-unknown-reference", "Unknown workbook ancestor")
            })?;
        }
        check()?;
        tx.execute(
            "UPDATE versions SET parent=?1 WHERE id=?2",
            params![parent, child],
        )
        .map_err(err)?;
    }
    for id in versions.keys().filter(|id| !keep.contains(id)) {
        check()?;
        if spills {
            tx.execute("DELETE FROM spill_groups WHERE version=?1", [id])
                .map_err(err)?;
        }
        tx.execute("DELETE FROM shared_normalizations WHERE version=?1", [id])
            .map_err(err)?;
        tx.execute("DELETE FROM changes WHERE version=?1", [id])
            .map_err(err)?;
        tx.execute("DELETE FROM versions WHERE id=?1", [id])
            .map_err(err)?;
    }
    check()
}
pub fn compact(root: &Path, input: &Value) -> Result<Value> {
    let mut db = database(root)?;
    let mut keep = input["keep"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing retained versions"))?
        .iter()
        .map(|v| {
            v.as_u64()
                .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid retained version"))
        })
        .collect::<Result<BTreeSet<_>>>()?;
    keep.insert(0);
    let before: u64 = db
        .query_row("SELECT count(*) FROM versions", [], |r| r.get(0))
        .map_err(err)?;
    let tx = db.transaction().map_err(err)?;
    let spills = crate::spill::present(&tx)?;
    let mut versions: BTreeMap<u64, Option<u64>> = tx
        .prepare_cached("SELECT id,parent FROM versions ORDER BY id")
        .map_err(err)?
        .query_map([], |r| Ok((r.get(0)?, r.get(1)?)))
        .map_err(err)?
        .collect::<std::result::Result<_, _>>()
        .map_err(err)?;
    validate_versions(&keep, &versions)?;
    crate::dependencies::retain(&tx, &keep.iter().copied().collect::<Vec<_>>())?;
    crate::calculation_policy::retain(&tx, &keep.iter().copied().collect::<Vec<_>>())?;
    crate::used_range::retain(&tx, &keep)?;
    crate::calculation_transfer::retain_publication(
        &tx,
        &keep.iter().copied().collect::<Vec<_>>(),
    )?;
    crate::merges::retain(&tx, &keep)?;
    retain_nodes(&tx, &keep, &versions, spills, crate::cancel::check)?;
    versions.retain(|id, _| keep.contains(id));
    crate::merges::prune(&tx)?;
    crate::dependencies::prune(&tx)?;
    crate::calculation_policy::prune(&tx)?;
    crate::used_range::prune(&tx)?;
    crate::used_range_transfer::prune(&tx)?;
    crate::calculation_transfer::prune_publication(&tx)?;
    tx.commit().map_err(err)?;
    // User-requested compaction returns freed pages to disk. A short-lived
    // immutable cache clone can reuse its free pages without rewriting a large
    // unchanged cell table; its full file length still counts toward cache quota.
    let vacuum = input["vacuum"].as_bool().unwrap_or(true);
    if vacuum {
        db.execute_batch("VACUUM;").map_err(err)?;
    }
    db.execute_batch("PRAGMA optimize;").map_err(err)?;
    let free_pages: u64 = db
        .query_row("PRAGMA freelist_count", [], |r| r.get(0))
        .map_err(err)?;
    let page_size: u64 = db
        .query_row("PRAGMA page_size", [], |r| r.get(0))
        .map_err(err)?;
    Ok(
        json!({"beforeVersions":before,"afterVersions":versions.len(),"indexBytes":std::fs::metadata(root.join("cells.sqlite")).map_err(err)?.len(),"vacuumed":vacuum,"retainedFreeBytes":free_pages*page_size}),
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use rusqlite::Connection;

    fn fixture() -> Connection {
        let db = Connection::open_in_memory().unwrap();
        db.execute_batch("CREATE TABLE versions(id INTEGER PRIMARY KEY,parent INTEGER);INSERT INTO versions VALUES(0,NULL);CREATE TABLE changes(version INTEGER,sheet TEXT,r INTEGER,c INTEGER,data TEXT,edit TEXT,PRIMARY KEY(version,sheet,r,c)) WITHOUT ROWID;").unwrap();
        crate::shared::schema(&db).unwrap();
        crate::spill::schema(&db).unwrap();
        db
    }
    fn versions(db: &Connection) -> BTreeMap<u64, Option<u64>> {
        db.prepare_cached("SELECT id,parent FROM versions ORDER BY id")
            .unwrap()
            .query_map([], |r| Ok((r.get(0)?, r.get(1)?)))
            .unwrap()
            .collect::<std::result::Result<_, _>>()
            .unwrap()
    }
    fn change(db: &Connection, version: u64, row: u32, data: Value, edit: Value) {
        db.execute(
            "INSERT INTO changes VALUES(?1,'Sheet',?2,0,?3,?4)",
            params![version, row, data.to_string(), edit.to_string()],
        )
        .unwrap();
    }
    // Independent closest-first oracle over the immutable parent map; preserve
    // JSON bytes, including raw source precision and cumulative export intent.
    fn effective(
        db: &Connection,
        version: u64,
        table: &str,
        key: &str,
        fields: &str,
    ) -> BTreeMap<String, String> {
        let parents = versions(db);
        let mut result = BTreeMap::new();
        let mut current = Some(version);
        while let Some(id) = current {
            let sql = format!(
                "SELECT {key},json_array({fields}) FROM {table} WHERE version=?1 ORDER BY {key}"
            );
            let rows = db
                .prepare_cached(&sql)
                .unwrap()
                .query_map([id], |row| {
                    Ok((row.get::<_, String>(0)?, row.get::<_, String>(1)?))
                })
                .unwrap()
                .collect::<std::result::Result<Vec<_>, _>>()
                .unwrap();
            for (key, value) in rows {
                result.entry(key).or_insert(value);
            }
            current = parents[&id];
        }
        result
    }
    fn snapshot(db: &Connection) -> Vec<String> {
        ["SELECT json_array(id,parent) FROM versions ORDER BY id", "SELECT json_array(version,sheet,r,c,data,edit) FROM changes ORDER BY version,sheet,r,c", "SELECT json_array(version,sheet,group_id,master_r,master_c,formula,invariant) FROM shared_normalizations ORDER BY version,sheet,group_id", "SELECT json_array(version,sheet,mr,mc,active,r0,c0,r1,c1,formula,proof,schema) FROM spill_groups ORDER BY version,sheet,mr,mc"].iter().flat_map(|sql| db.prepare_cached(sql).unwrap().query_map([], |r| r.get::<_,String>(0)).unwrap().collect::<std::result::Result<Vec<_>,_>>().unwrap()).collect()
    }

    #[test]
    fn deep_paths_keep_intermediate_two_branches_and_exact_cumulative_intent() {
        let mut db = fixture();
        let tx = db.transaction().unwrap();
        for id in 1..=4000u64 {
            tx.execute("INSERT INTO versions VALUES(?1,?2)", params![id, id - 1])
                .unwrap();
            change(
                &tx,
                id,
                id as u32,
                json!({"value":id,"rawValue":id.to_string(),"styleId":id%5}),
                json!({"value":id,"fontColor":"#123456"}),
            );
            if id % 17 == 0 {
                change(
                    &tx,
                    id,
                    0,
                    json!({"formula":"A1+1","rawValue":"9007199254740993","calculation":{"status":"calculated","value":{"$workbookNumber":"-0"}}}),
                    json!({"formula":"A1+1","numberFormat":"date","fontColor":"#123456","bold":id%2==0}),
                );
            }
        }
        for (id, parent, marker) in [
            (4001, 1500, "left"),
            (4002, 4001, "left-tip"),
            (4003, 2000, "abandoned"),
            (4004, 1500, "right"),
            (4005, 4004, "right-tip"),
        ] {
            tx.execute("INSERT INTO versions VALUES(?1,?2)", params![id, parent])
                .unwrap();
            change(
                &tx,
                id,
                0,
                json!({"value":marker,"styleId":8}),
                json!({"value":marker,"fontColor":"#abcdef","wrapText":true}),
            );
        }
        tx.commit().unwrap();
        let keep = BTreeSet::from([0, 1000, 3000, 4000, 4002, 4005]);
        let before: BTreeMap<_, _> = keep
            .iter()
            .map(|id| {
                (
                    *id,
                    effective(
                        &db,
                        *id,
                        "changes",
                        "printf('%s:%d:%d',sheet,r,c)",
                        "data,edit",
                    ),
                )
            })
            .collect();
        let parents = versions(&db);
        let tx = db.transaction().unwrap();
        retain_nodes(&tx, &keep, &parents, true, || Ok(())).unwrap();
        tx.commit().unwrap();
        assert_eq!(
            versions(&db),
            BTreeMap::from([
                (0, None),
                (1000, Some(0)),
                (3000, Some(1000)),
                (4000, Some(3000)),
                (4002, Some(1000)),
                (4005, Some(1000))
            ])
        );
        for id in keep {
            assert_eq!(
                effective(
                    &db,
                    id,
                    "changes",
                    "printf('%s:%d:%d',sheet,r,c)",
                    "data,edit"
                ),
                before[&id],
                "retained {id}"
            );
        }
        assert_eq!(
            db.prepare_cached("SELECT count(*) FROM changes WHERE version=4003")
                .unwrap()
                .query_row([], |r| r.get::<_, u32>(0))
                .unwrap(),
            0
        );
    }

    #[test]
    fn closest_shared_normalization_and_spill_release_win_on_each_branch() {
        let mut db = fixture();
        for (id, parent) in [(1, 0), (2, 1), (3, 2), (4, 1), (5, 4)] {
            db.execute("INSERT INTO versions VALUES(?1,?2)", params![id, parent])
                .unwrap();
        }
        for (id, formula, active) in [(1, "A1+1", 1), (2, "A1+2", 0), (4, "A1+4", 1)] {
            db.execute(
                "INSERT INTO shared_normalizations VALUES(?1,'Sheet','0',0,0,?2,0)",
                params![id, formula],
            )
            .unwrap();
            db.execute(
                "INSERT INTO spill_groups VALUES(?1,'Sheet',0,0,?2,0,0,1,1,?3,'exact-proof',1)",
                params![id, active, formula],
            )
            .unwrap();
        }
        let keep = BTreeSet::from([0, 3, 5]);
        let shared: BTreeMap<_, _> = keep
            .iter()
            .map(|id| {
                (
                    *id,
                    effective(
                        &db,
                        *id,
                        "shared_normalizations",
                        "sheet||':'||group_id",
                        "master_r,master_c,formula,invariant",
                    ),
                )
            })
            .collect();
        let spills: BTreeMap<_, _> = keep
            .iter()
            .map(|id| {
                (
                    *id,
                    effective(
                        &db,
                        *id,
                        "spill_groups",
                        "printf('%s:%d:%d',sheet,mr,mc)",
                        "active,r0,c0,r1,c1,formula,proof,schema",
                    ),
                )
            })
            .collect();
        let parents = versions(&db);
        let tx = db.transaction().unwrap();
        retain_nodes(&tx, &keep, &parents, true, || Ok(())).unwrap();
        tx.commit().unwrap();
        for id in keep {
            assert_eq!(
                effective(
                    &db,
                    id,
                    "shared_normalizations",
                    "sheet||':'||group_id",
                    "master_r,master_c,formula,invariant"
                ),
                shared[&id]
            );
            assert_eq!(
                effective(
                    &db,
                    id,
                    "spill_groups",
                    "printf('%s:%d:%d',sheet,mr,mc)",
                    "active,r0,c0,r1,c1,formula,proof,schema"
                ),
                spills[&id]
            );
        }
        assert_eq!(
            db.prepare_cached("SELECT active FROM spill_groups WHERE version=3")
                .unwrap()
                .query_row([], |r| r.get::<_, u32>(0))
                .unwrap(),
            0
        );
    }

    #[test]
    fn cancellation_after_copy_rolls_back_rows_and_parent_links() {
        let mut db = fixture();
        for id in 1..=6 {
            db.execute("INSERT INTO versions VALUES(?1,?2)", params![id, id - 1])
                .unwrap();
            change(
                &db,
                id,
                id as u32,
                json!({"value":id}),
                json!({"value":id,"bold":true}),
            );
        }
        let before = snapshot(&db);
        let parents = versions(&db);
        let mut checks = 0;
        {
            let tx = db.transaction().unwrap();
            let result = retain_nodes(&tx, &BTreeSet::from([0, 6]), &parents, true, || {
                checks += 1;
                if checks == 4 {
                    Err("Workbook operation cancelled".into())
                } else {
                    Ok(())
                }
            });
            assert_eq!(result.unwrap_err(), "Workbook operation cancelled");
            assert!(
                tx.prepare_cached("SELECT count(*) FROM changes WHERE version=6")
                    .unwrap()
                    .query_row([], |r| r.get::<_, u32>(0))
                    .unwrap()
                    > 1
            );
        }
        assert_eq!(snapshot(&db), before);
    }

    #[test]
    fn corrupt_parent_graph_and_unknown_keep_refuse_before_writes() {
        for parents in [
            BTreeMap::from([(0, None), (1, Some(9))]),
            BTreeMap::from([(0, None), (1, Some(2)), (2, Some(1))]),
        ] {
            let mut db = fixture();
            for (id, parent) in &parents {
                if *id != 0 {
                    db.execute("INSERT INTO versions VALUES(?1,?2)", params![id, parent])
                        .unwrap();
                }
            }
            let before = snapshot(&db);
            let tx = db.transaction().unwrap();
            assert!(
                retain_nodes(
                    &tx,
                    &parents.keys().copied().collect(),
                    &parents,
                    true,
                    || Ok(())
                )
                .is_err()
            );
            drop(tx);
            assert_eq!(snapshot(&db), before);
        }
        let mut db = fixture();
        let parents = versions(&db);
        let tx = db.transaction().unwrap();
        assert_eq!(
            crate::split_code(
                &retain_nodes(&tx, &BTreeSet::from([0, 9]), &parents, true, || Ok(())).unwrap_err()
            )
            .0,
            Some("internal-unknown-reference")
        );
    }
}
