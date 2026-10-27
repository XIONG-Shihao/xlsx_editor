//! Carry accepted formula authority across one independently validated pivot
//! package rewrite. XML caches alone never grant a current calculation result.
use crate::{Result, err, field, local};
use rusqlite::{Connection, OptionalExtension, params};
use serde_json::{Value, json};
use std::{fs, path::Path};

fn source_fields(cell: &Value) -> Value {
    let mut cell = cell.clone();
    if let Some(fields) = cell.as_object_mut() {
        for key in ["calculation", "cachedResult", "spill"] {
            fields.remove(key);
        }
        for key in ["value", "cachedValue"] {
            if let Some(number) = fields.get(key).and_then(Value::as_f64) {
                fields.insert(key.into(), json!(number));
            }
        }
    }
    cell
}
fn transfer_fields(input: &Value, required: &[&str], message: &str) -> Result<()> {
    let object = input.as_object().ok_or_else(|| message.to_string())?;
    if object
        .get("diagnostics")
        .is_some_and(|value| !value.is_boolean())
    {
        return Err(crate::coded(
            "formula-bad-input",
            "Calculation transfer diagnostics must be a boolean",
        ));
    }
    let fields = object
        .keys()
        .filter(|key| key.as_str() != "diagnostics")
        .map(String::as_str)
        .collect::<std::collections::BTreeSet<_>>();
    if fields
        != required
            .iter()
            .copied()
            .collect::<std::collections::BTreeSet<_>>()
    {
        return Err(crate::coded("formula-refused", message));
    }
    Ok(())
}
fn untouched_import(
    root: &Path,
    db: &rusqlite::Connection,
    version: u64,
    view: &Value,
) -> Result<()> {
    if !db
        .prepare_cached("SELECT EXISTS(SELECT 1 FROM versions WHERE id=?1)")
        .map_err(crate::err)?
        .query_row([version], |row| row.get::<_, bool>(0))
        .map_err(err)?
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown new calculation version",
        ));
    }
    if crate::spill::any_history(db, version)? {
        crate::spill::validate_current(root, db, &json!({"version":version,"view":view}))?;
    }
    // A stored spill adoption may add only ownership/provenance to unchanged
    // imported cells. Any effective literal, formula or style edit refuses.
    let mut after = (String::new(), -1i64, -1i64);
    loop {
        let mut query=db.prepare_cached("WITH RECURSIVE a(id)AS(SELECT ?1 UNION ALL SELECT v.parent FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL)SELECT DISTINCT sheet,r,c FROM changes JOIN a ON a.id=changes.version WHERE(sheet,r,c)>(?2,?3,?4) ORDER BY sheet,r,c LIMIT 512").map_err(err)?;
        let positions = query
            .query_map(params![version, after.0, after.1, after.2], |row| {
                Ok((
                    row.get::<_, String>(0)?,
                    row.get::<_, u32>(1)?,
                    row.get::<_, u32>(2)?,
                ))
            })
            .map_err(err)?
            .collect::<std::result::Result<Vec<_>, _>>()
            .map_err(err)?;
        for (sheet, r, c) in &positions {
            crate::cancel::check()?;
            let reference = crate::xml::address(*r, *c);
            let cell = |version| -> Result<Value> {
                let page = crate::range::read_db(
                    db,
                    &json!({"version":version,"sheet":sheet,"range":reference,"limit":2}),
                )?;
                Ok(page["cells"]
                    .as_array()
                    .and_then(|cells| cells.first())
                    .map(|pair| pair[1].clone())
                    .unwrap_or(Value::Null))
            };
            if source_fields(&cell(version)?) != source_fields(&cell(0)?) {
                return Err(crate::coded(
                    "internal-mismatch",
                    "Calculation transfer destination differs from its untouched imported SQL source",
                ));
            }
        }
        if positions.len() < 512 {
            break;
        }
        let (sheet, r, c) = positions.last().unwrap();
        after = (sheet.clone(), *r as i64, *c as i64);
    }
    Ok(())
}

pub(crate) fn command(root: &Path, input: &Value) -> Result<Value> {
    if input["purpose"] == "publication-restore" {
        return restore_publication(root, input);
    }
    if input.get("purpose").is_some() {
        return publication(root, input);
    }
    transfer_fields(
        input,
        &[
            "action",
            "operation",
            "root",
            "version",
            "previousRoot",
            "previousVersion",
            "sheet",
            "view",
        ],
        "Invalid calculation transfer fields",
    )?;
    let previous = Path::new(field(input, "previousRoot")?);
    if !previous.is_absolute()
        || fs::symlink_metadata(previous)
            .map_err(err)?
            .file_type()
            .is_symlink()
    {
        return Err(crate::coded(
            "formula-unsupported",
            "Unsafe previous calculation root",
        ));
    }
    let previous = fs::canonicalize(previous).map_err(err)?;
    if previous == root {
        return Err(crate::coded(
            "formula-bad-input",
            "Calculation transfer requires a new index",
        ));
    }
    let old_version = input["previousVersion"].as_u64().ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing previous calculation version",
        )
    })?;
    let version = input["version"].as_u64().ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing calculation transfer version",
        )
    })?;
    let sheet = field(input, "sheet")?;
    let old = crate::database(&previous)?;
    if !old
        .prepare_cached("SELECT EXISTS(SELECT 1 FROM versions WHERE id=?1)")
        .map_err(crate::err)?
        .query_row([old_version], |row| row.get::<_, bool>(0))
        .map_err(err)?
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown previous calculation version",
        ));
    }
    untouched_import(root, &crate::database(root)?, version, &input["view"])?;
    let plan: Value =
        serde_json::from_slice(&fs::read(local(&previous, "package-metadata.json")?).map_err(err)?)
            .map_err(err)?;
    if plan["kind"] != "refreshPivot" {
        return Err(crate::coded(
            "formula-bad-input",
            "Calculation transfer is not an owned pivot refresh",
        ));
    }
    let selected = crate::calculation_policy::command(
        &previous,
        &json!({"method":"scope","version":old_version,"view":input["view"],"sheet":sheet}),
    )?;
    if !selected["pivots"]
        .as_array()
        .is_some_and(|objects| objects.contains(&plan["parts"]["table"]))
    {
        return Err(crate::coded(
            "formula-unsupported",
            "Transferred pivot is outside selected prerequisites",
        ));
    }
    crate::package_edit::qualify_pivot_refresh(&previous, old_version, "export.xlsx", &plan)?;
    // Reproduce the native export and package edit from their accepted SQL
    // version. This also validates every opaque/unrelated package member.
    crate::export_state::prepare(
        &previous,
        &json!({"version":old_version,"view":input["view"]}),
    )?;
    let index: Value =
        serde_json::from_slice(&fs::read(local(&previous, "index.json")?).map_err(err)?)
            .map_err(err)?;
    crate::export::write(
        &previous,
        &json!({"version":old_version,"revision":index["revision"],"calculationMetadata":true}),
    )?;
    let rewritten = crate::package_edit::apply(
        &previous,
        &json!({"version":old_version,"source":"export.xlsx"}),
    )?;
    if crate::hash(&local(&previous, field(&rewritten, "file")?)?)?
        != crate::hash(&local(root, "source.xlsx")?)?
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Calculation transfer package differs from independent native rewrite",
        ));
    }
    let old_metadata = crate::package_edit::auditing_metadata(&previous)?;
    let new_metadata = crate::package_edit::auditing_metadata(root)?;
    if old_metadata != new_metadata {
        return Err(crate::coded(
            "formula-stale",
            "Calculation transfer changed worksheet/name/table bindings",
        ));
    }
    let old_policy = crate::calculation_policy::options(&old, old_version)?.ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing previous calculation policy",
        )
    })?;
    let mut db = crate::database(root)?;
    let tx = db.transaction().map_err(err)?;
    if !tx
        .prepare_cached("SELECT EXISTS(SELECT 1 FROM versions WHERE id=?1)")
        .map_err(crate::err)?
        .query_row([version], |row| row.get::<_, bool>(0))
        .map_err(err)?
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown new calculation version",
        ));
    }
    tx.execute("INSERT INTO versions(parent)VALUES(?1)", [version])
        .map_err(err)?;
    let accepted = tx.last_insert_rowid() as u64;
    let mut formulas = 0u64;
    for source in old_metadata["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing transfer worksheets"))?
    {
        let path = field(source, "path")?;
        let mut after: Option<String> = None;
        loop {
            let page = crate::range::read_db(
                &old,
                &json!({"version":old_version,"sheet":path,"range":"A1:XFD1048576","after":after,"limit":512,"formulas":true}),
            )?;
            for pair in page["cells"].as_array().ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing transfer formula page")
            })? {
                crate::cancel::check()?;
                let reference = pair[0].as_str().ok_or_else(|| {
                    crate::coded("internal-invalid-field", "Invalid transfer coordinate")
                })?;
                let cell = &pair[1];
                let next = crate::range::read_db(
                    &tx,
                    &json!({"version":version,"sheet":path,"range":reference,"limit":2}),
                )?;
                let next = next["cells"]
                    .as_array()
                    .and_then(|cells| cells.first())
                    .map(|pair| &pair[1])
                    .ok_or_else(|| {
                        crate::coded("internal-mismatch", "Transfer lost a formula cell")
                    })?;
                for field in [
                    "formula",
                    "formulaType",
                    "formulaReference",
                    "formulaSharedIndex",
                    "styleId",
                ] {
                    if cell[field] != next[field] {
                        return Err(crate::coded(
                            "formula-stale",
                            "Transfer changed formula source identity",
                        ));
                    }
                }
                let at = crate::xml::position(reference)?;
                let mut pending = crate::calculation_policy::pending(&old, old_version, path, at)?;
                // Version/source-bound special proofs are not portable tokens.
                // Retain their stored output, with explicit pending authority,
                // until their complete existing owner recalculates them.
                let special = cell.get("spill").is_some()
                    || cell.get("formulaType").is_some()
                    || cell["calculation"].get("scalarArray").is_some()
                    || cell["calculation"].get("spillConsumer").is_some();
                if special {
                    pending = true;
                } else if cell["calculation"].is_object() {
                    let mut carried = next.clone();
                    carried["calculation"] = cell["calculation"].clone();
                    crate::range::store_calculated_change(
                        &tx,
                        accepted as i64,
                        path,
                        at.0,
                        at.1,
                        &carried,
                        &json!({}),
                    )?;
                }
                tx.execute(
                    "INSERT OR REPLACE INTO calculation_pending VALUES(?1,?2,?3,?4,?5)",
                    params![accepted, path, at.0, at.1, pending],
                )
                .map_err(err)?;
                formulas += 1;
            }
            after = page["nextCursor"].as_str().map(str::to_owned);
            if after.is_none() {
                break;
            }
        }
    }
    tx.execute(
        "INSERT OR REPLACE INTO calculation_policies VALUES(?1,?2,?3)",
        params![accepted, old_policy.0, old_policy.1],
    )
    .map_err(err)?;
    for object in crate::package_edit::calculation_objects(root)?["objects"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing transfer objects"))?
    {
        let key = field(object, "key")?;
        let pending = if key == field(&plan["parts"], "table")? {
            false
        } else {
            crate::calculation_policy::derived_pending(&old, old_version, key)?
        };
        tx.execute(
            "INSERT OR REPLACE INTO calculation_derived VALUES(?1,?2,?3)",
            params![accepted, key, pending],
        )
        .map_err(err)?;
    }
    // A transferred scheduling graph is not an evaluator reuse proof.
    crate::dependencies::schema(&tx)?;
    tx.execute(
        "INSERT OR REPLACE INTO dependency_bindings VALUES(?1,0)",
        [accepted],
    )
    .map_err(err)?;
    tx.commit().map_err(err)?;
    let policy = crate::calculation_policy::command(
        root,
        &json!({"method":"fork","version":accepted,"view":input["view"]}),
    )?;
    Ok(
        json!({"version":policy["version"],"policy":policy["policy"],"transferredFormulaCells":formulas}),
    )
}

// Publication receipts are issued by the validated native export, never by a
// caller-supplied cache or hash. They authorize a source change only within the
// same IndexedWorkbook owner. The source package and all accepted data remain
// immutable; transfer writes exclusively into the new, unpublished index.
fn publication_schema(db: &Connection) -> Result<()> {
    db.execute_batch("CREATE TABLE IF NOT EXISTS publication_export_receipts(version INTEGER PRIMARY KEY,source_revision TEXT NOT NULL,export_revision TEXT NOT NULL,policy TEXT NOT NULL);CREATE TABLE IF NOT EXISTS publication_calculation_certificates(version INTEGER NOT NULL,sheet TEXT NOT NULL,r INTEGER NOT NULL,c INTEGER NOT NULL,source_revision TEXT NOT NULL,cell TEXT NOT NULL,PRIMARY KEY(version,sheet,r,c)) WITHOUT ROWID;CREATE TABLE IF NOT EXISTS publication_spill_certificates(version INTEGER NOT NULL,sheet TEXT NOT NULL,r INTEGER NOT NULL,c INTEGER NOT NULL,source_revision TEXT NOT NULL,proof TEXT NOT NULL,PRIMARY KEY(version,sheet,r,c)) WITHOUT ROWID;").map_err(err)
}
fn has_table(db: &Connection, name: &str) -> Result<bool> {
    db.prepare_cached("SELECT EXISTS(SELECT 1 FROM sqlite_master WHERE type='table' AND name=?1)")
        .map_err(crate::err)?
        .query_row([name], |row| row.get(0))
        .map_err(err)
}
fn source_index(root: &Path) -> Result<Value> {
    serde_json::from_slice(&fs::read(local(root, "index.json")?).map_err(err)?).map_err(err)
}
fn policy(root: &Path, db: &Connection, version: u64) -> Result<(String, bool)> {
    if let Some(value) = crate::calculation_policy::options(db, version)? {
        return Ok(value);
    }
    let source = crate::package_edit::source_calculation_policy(root)?;
    Ok((
        field(&source, "mode")?.into(),
        source["recalculateBeforeSave"]
            .as_bool()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing source Save policy"))?,
    ))
}
pub(crate) fn issue_publication_receipt(
    root: &Path,
    db: &Connection,
    version: u64,
    export_revision: &str,
) -> Result<()> {
    let index = source_index(root)?;
    let revision = field(&index, "revision")?;
    if crate::hash(&local(root, "source.xlsx")?)? != revision {
        return Err(crate::coded(
            "internal-mismatch",
            "Publication export source differs from accepted index",
        ));
    }
    if !db
        .prepare_cached("SELECT EXISTS(SELECT 1 FROM versions WHERE id=?1)")
        .map_err(crate::err)?
        .query_row([version], |row| row.get::<_, bool>(0))
        .map_err(err)?
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown publication export version",
        ));
    }
    publication_schema(db)?;
    db.execute(
        "INSERT OR REPLACE INTO publication_export_receipts VALUES(?1,?2,?3,?4)",
        params![
            version,
            revision,
            export_revision,
            serde_json::to_string(&policy(root, db, version)?).map_err(err)?
        ],
    )
    .map_err(err)?;
    Ok(())
}
fn certificate_subject(cell: &Value) -> Value {
    let mut subject = json!({});
    for name in [
        "formula",
        "formulaType",
        "formulaReference",
        "formulaSharedIndex",
        "value",
        "dataType",
        "rawValue",
        "cachedValue",
        "calculation",
        "spill",
        "scenarioResult",
    ] {
        if let Some(value) = cell.get(name) {
            subject[name] = value.clone();
        }
    }
    subject
}
/// A certificate preserves the authority already granted by the old native
/// owner. It does not grant current authority to a cache imported from XML.
pub(crate) fn publication_certificate(
    root: &Path,
    db: &Connection,
    version: u64,
    sheet: &str,
    reference: &str,
    cell: &Value,
) -> Result<bool> {
    if !has_table(db, "publication_calculation_certificates")? {
        return Ok(false);
    }
    let at = crate::xml::position(reference)?;
    let saved:Option<(String,String)>=db.prepare_cached("WITH RECURSIVE a(id,depth)AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL)SELECT c.source_revision,c.cell FROM a JOIN publication_calculation_certificates c ON c.version=a.id WHERE c.sheet=?2 AND c.r=?3 AND c.c=?4 ORDER BY a.depth LIMIT 1").map_err(crate::err)?.query_row(params![version,sheet,at.0,at.1],|row|Ok((row.get(0)?,row.get(1)?))).optional().map_err(err)?;
    let Some((revision, saved)) = saved else {
        return Ok(false);
    };
    if source_index(root)?["revision"] != revision {
        return Err(crate::coded(
            "formula-stale",
            "Publication calculation certificate source is stale",
        ));
    }
    let saved: Value = serde_json::from_str(&saved).map_err(err)?;
    Ok(saved == certificate_subject(cell))
}
fn group_subject(group: &crate::spill::Group) -> Value {
    json!({"formula":group.formula,"range":group.geometry.reference(),"proof":group.proof})
}
pub(crate) fn publication_spill_certificate(
    root: &Path,
    db: &Connection,
    version: u64,
    group: &crate::spill::Group,
) -> Result<bool> {
    if !has_table(db, "publication_spill_certificates")? {
        return Ok(false);
    }
    let saved:Option<(String,String)>=db.prepare_cached("WITH RECURSIVE a(id,depth)AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL)SELECT c.source_revision,c.proof FROM a JOIN publication_spill_certificates c ON c.version=a.id WHERE c.sheet=?2 AND c.r=?3 AND c.c=?4 ORDER BY a.depth LIMIT 1").map_err(crate::err)?.query_row(params![version,group.sheet,group.master.0,group.master.1],|row|Ok((row.get(0)?,row.get(1)?))).optional().map_err(err)?;
    let Some((revision, saved)) = saved else {
        return Ok(false);
    };
    if source_index(root)?["revision"] != revision {
        return Err(crate::coded(
            "formula-stale",
            "Publication spill certificate source is stale",
        ));
    }
    Ok(serde_json::from_str::<Value>(&saved).map_err(err)? == group_subject(group))
}
pub(crate) fn retain_publication(db: &Connection, keep: &[u64]) -> Result<()> {
    if !has_table(db, "publication_calculation_certificates")? {
        return Ok(());
    }
    for version in keep {
        for (table, column) in [
            ("publication_calculation_certificates", "cell"),
            ("publication_spill_certificates", "proof"),
        ] {
            db.execute(&format!("WITH RECURSIVE a(id,depth)AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL),c AS(SELECT x.*,row_number()OVER(PARTITION BY sheet,r,c ORDER BY a.depth)priority FROM a JOIN {table} x ON x.version=a.id)INSERT OR IGNORE INTO {table} SELECT ?1,sheet,r,c,source_revision,{column} FROM c WHERE priority=1"),[version]).map_err(err)?;
        }
    }
    Ok(())
}
pub(crate) fn prune_publication(db: &Connection) -> Result<()> {
    for table in [
        "publication_export_receipts",
        "publication_calculation_certificates",
        "publication_spill_certificates",
    ] {
        if has_table(db, table)? {
            db.execute(
                &format!("DELETE FROM {table} WHERE version NOT IN(SELECT id FROM versions)"),
                [],
            )
            .map_err(err)?;
        }
    }
    Ok(())
}
fn read_cell(db: &Connection, version: u64, sheet: &str, reference: &str) -> Result<Value> {
    let page = crate::range::read_db(
        db,
        &json!({"version":version,"sheet":sheet,"range":reference,"limit":2}),
    )?;
    Ok(page["cells"]
        .as_array()
        .and_then(|cells| cells.first())
        .map(|pair| pair[1].clone())
        .unwrap_or(Value::Null))
}
fn owned_publication_formula(
    db: &Connection,
    version: u64,
    sheet: &str,
    reference: &str,
    cell: &Value,
    next: &Value,
) -> Result<bool> {
    if cell["spill"]["role"] != "master"
        || cell["formulaType"] != "array"
        || next["formulaType"] != "array"
        || cell["formulaReference"] != next["formulaReference"]
    {
        return Ok(false);
    }
    let at = crate::xml::position(reference)?;
    let Some(group) = crate::spill::effective(db, version, sheet)?
        .into_iter()
        .find(|group| group.master == at)
    else {
        return Ok(false);
    };
    if group.proof["qualification"] != "retained-sequence-v1"
        || cell["formula"] != group.formula
        || cell["spill"] != group.cell_metadata(at)?
        || cell["formulaReference"] != group.geometry.reference()
    {
        return Ok(false);
    }
    if next["formula"] != crate::export::spill_storage_formula(&group.formula) {
        return Ok(false);
    }
    // This validates every accepted output position against the native group,
    // including a pending cache retained from an actual ancestor. A source XML
    // array tag or a merely similar formula cannot satisfy this ownership gate.
    crate::spill::accepted_cell(db, version, sheet, at, cell)
}
fn publication(root: &Path, input: &Value) -> Result<Value> {
    let notes = input["purpose"] == "note-authoring";
    let background = input["purpose"] == "sheet-background";
    let theme = input["purpose"] == "page-theme";
    let authoring = input["purpose"] == "hyperlink-authoring";
    let clear = input["purpose"] == "clear-hyperlinks" || authoring || theme || background || notes;
    if clear && (input["action"] != "workbook" || input["operation"] != "calculation-transfer") {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid Clear Hyperlinks calculation transfer operation",
        ));
    }
    if clear {
        transfer_fields(
            input,
            &[
                "action",
                "operation",
                "purpose",
                "root",
                "version",
                "previousRoot",
                "previousVersion",
                "sourceRevision",
                "rewriteId",
                "view",
            ],
            "Invalid Clear Hyperlinks calculation transfer fields",
        )?;
    } else {
        transfer_fields(
            input,
            &[
                "action",
                "operation",
                "purpose",
                "root",
                "version",
                "previousRoot",
                "previousVersion",
                "sourceRevision",
                "view",
            ],
            "Invalid publication calculation transfer fields",
        )?;
    }
    if !clear && input["purpose"] != "publication" {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid publication calculation transfer fields",
        ));
    }
    let previous = Path::new(field(input, "previousRoot")?);
    if !previous.is_absolute()
        || fs::symlink_metadata(previous)
            .map_err(err)?
            .file_type()
            .is_symlink()
    {
        return Err(crate::coded(
            "formula-unsupported",
            "Unsafe previous publication root",
        ));
    }
    let previous = fs::canonicalize(previous).map_err(err)?;
    let canonical_root = fs::canonicalize(root).map_err(err)?;
    if previous == canonical_root || previous.parent() != canonical_root.parent() {
        return Err(crate::coded(
            "formula-bad-input",
            "Publication transfer requires sibling indices from the same owner",
        ));
    }
    let old_version = input["previousVersion"].as_u64().ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing previous publication version",
        )
    })?;
    let version = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing publication version"))?;
    let revision = field(input, "sourceRevision")?;
    let old_index = source_index(&previous)?;
    let index = source_index(root)?;
    if index["revision"] != revision
        || crate::hash(&local(root, "source.xlsx")?)? != revision
        || crate::hash(&local(&previous, "source.xlsx")?)? != field(&old_index, "revision")?
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Publication transfer source revision mismatch",
        ));
    }
    let old = crate::database(&previous)?;
    if !old
        .prepare_cached("SELECT EXISTS(SELECT 1 FROM versions WHERE id=?1)")
        .map_err(crate::err)?
        .query_row([old_version], |row| row.get::<_, bool>(0))
        .map_err(err)?
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown previous publication version",
        ));
    }
    let old_policy = policy(&previous, &old, old_version)?;
    if clear {
        if notes {
            crate::used_range_transfer::validate_note_transfer(
                &previous,
                root,
                old_version,
                version,
                revision,
                field(input, "rewriteId")?,
            )?;
        } else if background {
            crate::used_range_transfer::validate_background_transfer(
                &previous,
                root,
                old_version,
                version,
                revision,
                field(input, "rewriteId")?,
            )?;
        } else if theme {
            crate::used_range_transfer::validate_page_theme_transfer(
                &previous,
                root,
                old_version,
                version,
                revision,
                field(input, "rewriteId")?,
            )?;
        } else if authoring {
            crate::used_range_transfer::validate_hyperlink_transfer(
                &previous,
                root,
                old_version,
                version,
                revision,
                field(input, "rewriteId")?,
            )?;
        } else {
            crate::used_range_transfer::validate_clear_hyperlinks_transfer(
                &previous,
                root,
                old_version,
                version,
                revision,
                field(input, "rewriteId")?,
            )?;
        }
    } else {
        if !has_table(&old, "publication_export_receipts")? {
            return Err(crate::coded(
                "formula-bad-input",
                "Publication transfer has no native export receipt",
            ));
        }
        let receipt:Option<(String,String,String)>=old.prepare_cached("SELECT source_revision,export_revision,policy FROM publication_export_receipts WHERE version=?1").map_err(crate::err)?.query_row([old_version],|row|Ok((row.get(0)?,row.get(1)?,row.get(2)?))).optional().map_err(err)?;
        let receipt = receipt.ok_or_else(|| {
            crate::coded(
                "formula-bad-input",
                "Publication transfer has no selected-version native export receipt",
            )
        })?;
        if receipt.0 != field(&old_index, "revision")?
            || receipt.1 != revision
            || receipt.2 != serde_json::to_string(&old_policy).map_err(err)?
        {
            return Err(crate::coded(
                "internal-mismatch",
                "Publication transfer differs from selected-version native export receipt",
            ));
        }
    }
    let metadata = crate::package_edit::auditing_metadata(&previous)?;
    if metadata != crate::package_edit::auditing_metadata(root)? {
        return Err(crate::coded(
            "formula-stale",
            "Publication transfer changed worksheet/name/table bindings",
        ));
    }
    let view = &input["view"];
    let sheets = view["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing publication worksheets"))?
        .iter()
        .map(|sheet| Ok(json!({"name":field(sheet,"name")?,"path":field(sheet,"path")?})))
        .collect::<Result<Vec<_>>>()?;
    if metadata["sheets"] != json!(sheets)
        || !crate::package_edit::same_records(&metadata["namedRanges"], &view["namedRanges"])
        || !crate::package_edit::same_records(&metadata["tables"], &view["tables"])
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Publication transfer metadata differs from native source",
        ));
    }
    let mut db = crate::database(root)?;
    untouched_import(root, &db, version, view)?;
    if policy(root, &db, version)? != old_policy {
        return Err(crate::coded(
            "internal-mismatch",
            "Publication transfer destination calculation policy mismatch",
        ));
    }
    let tx = db.transaction().map_err(err)?;
    publication_schema(&tx)?;
    tx.execute("INSERT INTO versions(parent)VALUES(?1)", [version])
        .map_err(err)?;
    let accepted = tx.last_insert_rowid() as u64;
    let (mut formulas, mut derived, mut groups) = (0u64, 0u64, 0u64);
    for source in metadata["sheets"].as_array().ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing publication worksheet metadata",
        )
    })? {
        let sheet = field(source, "path")?;
        let mut after: Option<String> = None;
        // The accepted cells are visited in the same order as the old ones.
        let mut accepted_cells = crate::range::OrderedCells::new(&tx, version, sheet, false);
        loop {
            let page = crate::range::read_db(
                &old,
                &json!({"version":old_version,"sheet":sheet,"range":"A1:XFD1048576","after":after,"limit":512}),
            )?;
            for pair in page["cells"].as_array().ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing publication cell page")
            })? {
                crate::cancel::check()?;
                let reference = pair[0].as_str().ok_or_else(|| {
                    crate::coded("internal-invalid-field", "Invalid publication coordinate")
                })?;
                let cell = &pair[1];
                let at = crate::xml::position(reference)?;
                let is_formula = cell["formula"].is_string() || cell.get("formulaType").is_some();
                if !is_formula
                    && cell.get("calculation").is_none()
                    && cell.get("spill").is_none()
                    && cell.get("scenarioResult").is_none()
                {
                    continue;
                }
                let mut carried = accepted_cells.get(at)?.unwrap_or(Value::Null);
                if !carried.is_object() {
                    return Err(crate::coded(
                        "internal-mismatch",
                        "Publication lost an accepted formula or derived result",
                    ));
                }
                if is_formula && crate::calculation_policy::preserving(&old, old_version)? {
                    if crate::preservation::content_fields(cell)
                        != crate::preservation::content_fields(&carried)
                    {
                        return Err(crate::coded(
                            "formula-stale",
                            "Preservation publication changed stored formula/cache identity",
                        ));
                    }
                    tx.execute(
                        "INSERT OR REPLACE INTO calculation_pending VALUES(?1,?2,?3,?4,1)",
                        params![accepted, sheet, at.0, at.1],
                    )
                    .map_err(err)?;
                    continue;
                }
                if is_formula {
                    let same = cell["formula"] == carried["formula"];
                    let canonical = ["scalarArray", "spillConsumer"].iter().any(|marker| {
                        cell["calculation"][marker]["canonicalFormula"] == carried["formula"]
                            && carried["formula"].is_string()
                    });
                    let owned_canonical = if !same && !canonical {
                        owned_publication_formula(
                            &old,
                            old_version,
                            sheet,
                            reference,
                            cell,
                            &carried,
                        )?
                    } else {
                        false
                    };
                    if !same && !canonical && !owned_canonical
                        || cell["formulaType"] != carried["formulaType"]
                        || cell["formulaReference"] != carried["formulaReference"]
                    {
                        return Err(crate::coded(
                            "formula-stale",
                            "Publication changed accepted formula identity",
                        ));
                    }
                    formulas += 1;
                } else {
                    derived += 1;
                }
                if let Some(calculation) = cell.get("calculation") {
                    let mut calculation = calculation.clone();
                    for marker in ["scalarArray", "spillConsumer"] {
                        if calculation[marker].is_object() {
                            calculation[marker]["formula"] = carried["formula"].clone();
                        }
                    }
                    carried["calculation"] = calculation;
                }
                if let Some(spill) = cell.get("spill") {
                    carried["spill"] = spill.clone();
                    // Complete spill output values are runtime ownership fields;
                    // imported masters otherwise have value:null and only a cache.
                    for field in ["value", "dataType", "rawValue", "error"] {
                        if let Some(value) = cell.get(field) {
                            carried[field] = value.clone();
                        } else if let Some(fields) = carried.as_object_mut() {
                            fields.remove(field);
                        }
                    }
                }
                if let Some(scenario) = cell.get("scenarioResult") {
                    carried["scenarioResult"] = scenario.clone();
                }
                crate::range::store_calculated_change(
                    &tx,
                    accepted as i64,
                    sheet,
                    at.0,
                    at.1,
                    &carried,
                    &json!({}),
                )?;
                if is_formula {
                    tx.execute(
                        "INSERT OR REPLACE INTO calculation_pending VALUES(?1,?2,?3,?4,?5)",
                        params![
                            accepted,
                            sheet,
                            at.0,
                            at.1,
                            crate::calculation_policy::pending(&old, old_version, sheet, at)?
                        ],
                    )
                    .map_err(err)?;
                }
                tx.execute(
                    "INSERT INTO publication_calculation_certificates VALUES(?1,?2,?3,?4,?5,?6)",
                    params![
                        accepted,
                        sheet,
                        at.0,
                        at.1,
                        revision,
                        certificate_subject(&carried).to_string()
                    ],
                )
                .map_err(err)?;
            }
            after = page["nextCursor"].as_str().map(str::to_owned);
            if after.is_none() {
                break;
            }
        }
        for mut group in crate::spill::effective(&old, old_version, sheet)? {
            group.version = accepted;
            group.formula = crate::export::spill_storage_formula(&group.formula);
            crate::spill::schema(&tx)?;
            crate::spill::record(&tx, &group, true)?;
            tx.execute(
                "INSERT INTO publication_spill_certificates VALUES(?1,?2,?3,?4,?5,?6)",
                params![
                    accepted,
                    sheet,
                    group.master.0,
                    group.master.1,
                    revision,
                    group_subject(&group).to_string()
                ],
            )
            .map_err(err)?;
            groups += 1;
        }
    }
    tx.execute(
        "INSERT OR REPLACE INTO calculation_policies VALUES(?1,?2,?3)",
        params![accepted, old_policy.0, old_policy.1],
    )
    .map_err(err)?;
    for object in crate::package_edit::calculation_objects(root)?["objects"]
        .as_array()
        .ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing publication derived objects",
            )
        })?
    {
        let key = field(object, "key")?;
        tx.execute(
            "INSERT OR REPLACE INTO calculation_derived VALUES(?1,?2,?3)",
            params![
                accepted,
                key,
                crate::calculation_policy::derived_pending(&old, old_version, key)?
            ],
        )
        .map_err(err)?;
    }
    crate::dependencies::schema(&tx)?;
    tx.execute(
        "INSERT OR REPLACE INTO dependency_bindings VALUES(?1,0)",
        [accepted],
    )
    .map_err(err)?;
    tx.commit().map_err(err)?;
    let status = crate::calculation_policy::command(
        root,
        &json!({"method":"status","version":accepted,"view":view}),
    )?;
    let witness = write_witness(root, &db, accepted, revision, &metadata, &old_policy)?;
    Ok(
        json!({"schema":if notes{workbook_format!(SCHEMA_NOTE_CALCULATION_TRANSFER_V1)}else if background{workbook_format!(SCHEMA_BACKGROUND_CALCULATION_TRANSFER_V1)}else if theme{workbook_format!(SCHEMA_PAGE_THEME_CALCULATION_TRANSFER_V1)}else if authoring{workbook_format!(SCHEMA_HYPERLINK_CALCULATION_TRANSFER_V1)}else if clear{workbook_format!(SCHEMA_CLEAR_HYPERLINKS_CALCULATION_TRANSFER_V1)}else{workbook_format!(SCHEMA_PUBLICATION_CALCULATION_TRANSFER_V1)},"version":accepted,"policy":status,"transferredFormulaCells":formulas,"transferredDerivedCells":derived,"transferredSpillGroups":groups,"witness":witness}),
    )
}

const WITNESS_APPLICATION: u32 = 0x52504658;
const WITNESS_LIMIT: u64 = 256 * 1024 * 1024;
const WITNESS_ROW_LIMIT: u64 = 1024 * 1024;
fn single_link(info: &fs::Metadata) -> bool {
    #[cfg(unix)]
    {
        use std::os::unix::fs::MetadataExt;
        info.nlink() == 1
    }
    #[cfg(not(unix))]
    {
        let _ = info;
        true
    }
}
fn write_witness(
    root: &Path,
    db: &Connection,
    version: u64,
    revision: &str,
    bindings: &Value,
    options: &(String, bool),
) -> Result<Value> {
    let file = local(root, "publication-calculation.sqlite")?;
    // Never overwrite another candidate's witness.
    let created = fs::OpenOptions::new()
        .create_new(true)
        .write(true)
        .open(&file)
        .map_err(err)?;
    drop(created);
    crate::quota::reserve_artifact(32 * 1024 + bindings.to_string().len() as u64 * 2)?;
    let mut witness = Connection::open(&file).map_err(err)?;
    crate::quota::database(&witness)?;
    let maximum = witness
        .prepare_cached("PRAGMA max_page_count")
        .map_err(crate::err)?
        .query_row([], |row| row.get::<_, u64>(0))
        .map_err(err)?
        .min(WITNESS_LIMIT / 4096);
    witness.execute_batch(&format!("PRAGMA application_id={WITNESS_APPLICATION};PRAGMA user_version=1;PRAGMA page_size=4096;PRAGMA max_page_count={maximum};CREATE TABLE metadata(schema INTEGER PRIMARY KEY,source_revision TEXT NOT NULL,bindings TEXT NOT NULL,policy TEXT NOT NULL);CREATE TABLE cells(sheet TEXT NOT NULL,r INTEGER NOT NULL,c INTEGER NOT NULL,source TEXT NOT NULL,accepted TEXT NOT NULL,pending INTEGER NOT NULL,PRIMARY KEY(sheet,r,c)) WITHOUT ROWID;CREATE TABLE derived(object TEXT PRIMARY KEY,pending INTEGER NOT NULL) WITHOUT ROWID;CREATE TABLE groups(sheet TEXT NOT NULL,r INTEGER NOT NULL,c INTEGER NOT NULL,data TEXT NOT NULL,PRIMARY KEY(sheet,r,c)) WITHOUT ROWID;")).map_err(err)?;
    let tx = witness.transaction().map_err(err)?;
    tx.execute(
        "INSERT INTO metadata VALUES(1,?1,?2,?3)",
        params![
            revision,
            bindings.to_string(),
            serde_json::to_string(options).map_err(err)?
        ],
    )
    .map_err(err)?;
    let mut query=db.prepare_cached("SELECT sheet,r,c FROM publication_calculation_certificates WHERE version=?1 ORDER BY sheet,r,c").map_err(err)?;
    let mut rows = query.query([version]).map_err(err)?;
    let mut positions = Vec::new();
    while let Some(row) = rows.next().map_err(err)? {
        crate::cancel::check()?;
        positions.push((
            row.get::<_, String>(0).map_err(err)?,
            row.get::<_, u32>(1).map_err(err)?,
            row.get::<_, u32>(2).map_err(err)?,
        ));
    }
    // Read every source cell, then every accepted cell. Alternating versions per
    // cell rebuilt the shared-formula context twice for each formula.
    let read_all = |at: u64| -> Result<Vec<String>> {
        let mut stream: Option<(String, crate::range::OrderedCells)> = None;
        positions
            .iter()
            .map(|(sheet, r, c)| {
                crate::cancel::check()?;
                // Certificates are formula cells, so stream formulas only; a position that
                // is not a formula at this version is read on its own.
                if stream.as_ref().is_none_or(|(name, _)| name != sheet) {
                    stream = Some((
                        sheet.clone(),
                        crate::range::OrderedCells::new(db, at, sheet, true),
                    ));
                }
                let text = match stream.as_mut().unwrap().1.get((*r, *c))? {
                    Some(cell) => cell,
                    None => read_cell(db, at, sheet, &crate::xml::address(*r, *c))?,
                }
                .to_string();
                if text.len() as u64 > WITNESS_ROW_LIMIT {
                    return Err(crate::coded(
                        "formula-limit",
                        "Publication witness cell exceeds its bound",
                    ));
                }
                Ok(text)
            })
            .collect()
    };
    let sources = read_all(0)?;
    let accepted_cells = read_all(version)?;
    {
        let mut insert = tx
            .prepare_cached("INSERT INTO cells VALUES(?1,?2,?3,?4,?5,?6)")
            .map_err(err)?;
        for (((sheet, r, c), source), accepted) in
            positions.iter().zip(&sources).zip(&accepted_cells)
        {
            crate::cancel::check()?;
            crate::quota::reserve_artifact(2048 + (source.len() + accepted.len()) as u64 * 2)?;
            insert
                .execute(params![
                    sheet,
                    r,
                    c,
                    source,
                    accepted,
                    crate::calculation_policy::pending(db, version, sheet, (*r, *c))?
                ])
                .map_err(err)?;
        }
    }
    for object in crate::package_edit::calculation_objects(root)?["objects"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing witness derived objects"))?
    {
        let key = field(object, "key")?;
        crate::quota::reserve_artifact(2048 + key.len() as u64 * 2)?;
        tx.execute(
            "INSERT INTO derived VALUES(?1,?2)",
            params![
                key,
                crate::calculation_policy::derived_pending(db, version, key)?
            ],
        )
        .map_err(err)?;
    }
    for sheet in bindings["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing witness worksheets"))?
    {
        let sheet = field(sheet, "path")?;
        for group in crate::spill::effective(db, version, sheet)? {
            let data=json!({"formula":group.formula,"range":group.geometry.reference(),"proof":group.proof}).to_string();
            if data.len() as u64 > WITNESS_ROW_LIMIT {
                return Err(crate::coded(
                    "formula-limit",
                    "Publication witness spill proof exceeds its bound",
                ));
            }
            crate::quota::reserve_artifact(2048 + data.len() as u64 * 2)?;
            tx.execute(
                "INSERT INTO groups VALUES(?1,?2,?3,?4)",
                params![sheet, group.master.0, group.master.1, data],
            )
            .map_err(err)?;
        }
    }
    tx.commit().map_err(err)?;
    drop(rows);
    drop(query);
    witness.close().map_err(|(_, error)| err(error))?;
    fs::File::open(&file)
        .map_err(err)?
        .sync_all()
        .map_err(err)?;
    let info = fs::symlink_metadata(&file).map_err(err)?;
    let bytes = info.len();
    if !info.is_file()
        || info.file_type().is_symlink()
        || !single_link(&info)
        || bytes > WITNESS_LIMIT
    {
        return Err(crate::coded(
            "formula-limit",
            "Publication witness exceeds its storage or file-identity bound",
        ));
    }
    Ok(
        json!({"schema":1,"file":"publication-calculation.sqlite","revision":crate::hash(&file)?,"bytes":bytes}),
    )
}
fn restore_publication(root: &Path, input: &Value) -> Result<Value> {
    transfer_fields(
        input,
        &[
            "action",
            "operation",
            "purpose",
            "root",
            "version",
            "sourceRevision",
            "witness",
            "witnessRevision",
            "view",
        ],
        "Invalid publication witness restoration fields",
    )?;
    let file = Path::new(field(input, "witness")?);
    let revision = field(input, "sourceRevision")?;
    let witness_revision = field(input, "witnessRevision")?;
    let info = fs::symlink_metadata(file).map_err(err)?;
    if !file.is_absolute()
        || !info.is_file()
        || info.file_type().is_symlink()
        || !single_link(&info)
        || info.len() > WITNESS_LIMIT
        || info.len() < 4096
    {
        return Err(crate::coded(
            "formula-limit",
            "Unsafe or oversized publication witness",
        ));
    }
    if crate::hash(file)? != witness_revision {
        return Err(crate::coded(
            "internal-mismatch",
            "Publication witness revision mismatch",
        ));
    }
    let index = source_index(root)?;
    if index["revision"] != revision || crate::hash(&local(root, "source.xlsx")?)? != revision {
        return Err(crate::coded(
            "internal-mismatch",
            "Publication witness source revision mismatch",
        ));
    }
    let witness = Connection::open_with_flags(
        file,
        rusqlite::OpenFlags::SQLITE_OPEN_READ_ONLY | rusqlite::OpenFlags::SQLITE_OPEN_NO_MUTEX,
    )
    .map_err(err)?;
    witness.execute_batch("PRAGMA query_only=ON;PRAGMA trusted_schema=OFF;PRAGMA cache_size=-4096;PRAGMA mmap_size=0;").map_err(err)?;
    if witness
        .prepare_cached("PRAGMA application_id")
        .map_err(crate::err)?
        .query_row([], |row| row.get::<_, u32>(0))
        .map_err(err)?
        != WITNESS_APPLICATION
        || witness
            .prepare_cached("PRAGMA user_version")
            .map_err(crate::err)?
            .query_row([], |row| row.get::<_, u32>(0))
            .map_err(err)?
            != 1
    {
        return Err(crate::coded(
            "formula-unsupported",
            "Unsupported publication witness schema",
        ));
    }
    if witness
        .prepare_cached("PRAGMA integrity_check(1)")
        .map_err(crate::err)?
        .query_row([], |row| row.get::<_, String>(0))
        .map_err(err)?
        != "ok"
    {
        return Err(crate::coded(
            "formula-malformed",
            "Damaged publication witness",
        ));
    }
    let tables=witness.prepare_cached("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name").map_err(err)?.query_map([],|row|row.get::<_,String>(0)).map_err(err)?.collect::<std::result::Result<Vec<_>,_>>().map_err(err)?;
    if tables != ["cells", "derived", "groups", "metadata"] {
        return Err(crate::coded(
            "internal-mismatch",
            "Publication witness table inventory differs",
        ));
    }
    if witness
        .prepare_cached("SELECT count(*) FROM sqlite_master WHERE type IN('trigger','view')")
        .map_err(crate::err)?
        .query_row([], |row| row.get::<_, u64>(0))
        .map_err(err)?
        != 0
    {
        return Err(crate::coded(
            "formula-refused",
            "Publication witness executable schema is refused",
        ));
    }
    if witness
        .prepare_cached("SELECT count(*) FROM metadata")
        .map_err(crate::err)?
        .query_row([], |row| row.get::<_, u64>(0))
        .map_err(err)?
        != 1
    {
        return Err(crate::coded(
            "formula-refused",
            "Publication witness metadata is ambiguous",
        ));
    }
    let (saved_revision, bindings, options): (String, String, String) = witness
        .prepare_cached("SELECT source_revision,bindings,policy FROM metadata WHERE schema=1")
        .map_err(crate::err)?
        .query_row([], |row| Ok((row.get(0)?, row.get(1)?, row.get(2)?)))
        .map_err(err)?;
    if saved_revision != revision {
        return Err(crate::coded(
            "formula-stale",
            "Publication witness belongs to another source",
        ));
    }
    let bindings: Value = serde_json::from_str(&bindings).map_err(err)?;
    let options: (String, bool) = serde_json::from_str(&options).map_err(err)?;
    let metadata = crate::package_edit::auditing_metadata(root)?;
    if bindings != metadata {
        return Err(crate::coded(
            "formula-stale",
            "Publication witness worksheet/name/table bindings changed",
        ));
    }
    let view = &input["view"];
    let source_sheets = view["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing witness worksheets"))?
        .iter()
        .map(|sheet| Ok(json!({"name":field(sheet,"name")?,"path":field(sheet,"path")?})))
        .collect::<Result<Vec<_>>>()?;
    if metadata["sheets"] != json!(source_sheets)
        || !crate::package_edit::same_records(&metadata["namedRanges"], &view["namedRanges"])
        || !crate::package_edit::same_records(&metadata["tables"], &view["tables"])
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Publication witness caller metadata differs",
        ));
    }
    let version = input["version"].as_u64().ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing witness destination version",
        )
    })?;
    let mut db = crate::database(root)?;
    untouched_import(root, &db, version, view)?;
    if policy(root, &db, version)? != options {
        return Err(crate::coded(
            "formula-stale",
            "Publication witness calculation options changed",
        ));
    }
    let tx = db.transaction().map_err(err)?;
    publication_schema(&tx)?;
    tx.execute("INSERT INTO versions(parent)VALUES(?1)", [version])
        .map_err(err)?;
    let accepted = tx.last_insert_rowid() as u64;
    let (mut formulas, mut derived, mut groups) = (0u64, 0u64, 0u64);
    let mut query=witness.prepare_cached("SELECT sheet,r,c,source,accepted,pending,length(source),length(accepted) FROM cells ORDER BY sheet,r,c").map_err(err)?;
    let mut rows = query.query([]).map_err(err)?;
    while let Some(row) = rows.next().map_err(err)? {
        crate::cancel::check()?;
        if row.get::<_, u64>(6).map_err(err)? > WITNESS_ROW_LIMIT
            || row.get::<_, u64>(7).map_err(err)? > WITNESS_ROW_LIMIT
        {
            return Err(crate::coded(
                "formula-limit",
                "Publication witness cell exceeds its bound",
            ));
        }
        let sheet: String = row.get(0).map_err(err)?;
        let (r, c) = (
            row.get::<_, u32>(1).map_err(err)?,
            row.get::<_, u32>(2).map_err(err)?,
        );
        let reference = crate::xml::address(r, c);
        if r >= 1_048_576
            || c >= 16_384
            || !metadata["sheets"]
                .as_array()
                .is_some_and(|sheets| sheets.iter().any(|source| source["path"] == sheet))
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid publication witness coordinate",
            ));
        }
        let source: Value =
            serde_json::from_str(&row.get::<_, String>(3).map_err(err)?).map_err(err)?;
        let cell: Value =
            serde_json::from_str(&row.get::<_, String>(4).map_err(err)?).map_err(err)?;
        let pending: bool = row.get(5).map_err(err)?;
        if !cell.is_object() || source != read_cell(&tx, 0, &sheet, &reference)? {
            return Err(crate::coded(
                "formula-stale",
                "Publication witness source cell/cache declaration changed",
            ));
        }
        if cell["formula"] != source["formula"]
            || cell["formulaType"] != source["formulaType"]
            || cell["formulaReference"] != source["formulaReference"]
            || cell["styleId"] != source["styleId"]
        {
            return Err(crate::coded(
                "internal-mismatch",
                "Publication witness changes authored formula or style",
            ));
        }
        let is_formula = cell["formula"].is_string() || cell.get("formulaType").is_some();
        if is_formula {
            formulas += 1;
        } else {
            derived += 1;
        }
        crate::range::store_calculated_change(
            &tx,
            accepted as i64,
            &sheet,
            r,
            c,
            &cell,
            &json!({}),
        )?;
        if is_formula {
            tx.execute(
                "INSERT OR REPLACE INTO calculation_pending VALUES(?1,?2,?3,?4,?5)",
                params![accepted, sheet, r, c, pending],
            )
            .map_err(err)?;
        }
        tx.execute(
            "INSERT INTO publication_calculation_certificates VALUES(?1,?2,?3,?4,?5,?6)",
            params![
                accepted,
                sheet,
                r,
                c,
                revision,
                certificate_subject(&cell).to_string()
            ],
        )
        .map_err(err)?;
    }
    let actual_objects = crate::package_edit::calculation_objects(root)?["objects"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing witness objects"))?
        .iter()
        .map(|object| field(object, "key").map(str::to_owned))
        .collect::<Result<std::collections::BTreeSet<_>>>()?;
    let recorded_objects = witness
        .prepare_cached("SELECT object FROM derived ORDER BY object")
        .map_err(err)?
        .query_map([], |row| row.get::<_, String>(0))
        .map_err(err)?
        .collect::<std::result::Result<std::collections::BTreeSet<_>, _>>()
        .map_err(err)?;
    if actual_objects != recorded_objects {
        return Err(crate::coded(
            "formula-stale",
            "Publication witness derived object inventory changed",
        ));
    }
    let mut query = witness
        .prepare_cached("SELECT object,pending FROM derived")
        .map_err(err)?;
    let mut rows = query.query([]).map_err(err)?;
    while let Some(row) = rows.next().map_err(err)? {
        tx.execute(
            "INSERT OR REPLACE INTO calculation_derived VALUES(?1,?2,?3)",
            params![
                accepted,
                row.get::<_, String>(0).map_err(err)?,
                row.get::<_, bool>(1).map_err(err)?
            ],
        )
        .map_err(err)?;
    }
    let mut query = witness
        .prepare_cached("SELECT sheet,r,c,data,length(data) FROM groups ORDER BY sheet,r,c")
        .map_err(err)?;
    let mut rows = query.query([]).map_err(err)?;
    while let Some(row) = rows.next().map_err(err)? {
        if row.get::<_, u64>(4).map_err(err)? > WITNESS_ROW_LIMIT {
            return Err(crate::coded(
                "formula-limit",
                "Publication witness spill exceeds its bound",
            ));
        }
        let sheet: String = row.get(0).map_err(err)?;
        let master = (
            row.get::<_, u32>(1).map_err(err)?,
            row.get::<_, u32>(2).map_err(err)?,
        );
        let data: Value =
            serde_json::from_str(&row.get::<_, String>(3).map_err(err)?).map_err(err)?;
        let group = crate::spill::Group {
            version: accepted,
            sheet: sheet.clone(),
            master,
            geometry: crate::spill::Geometry::from_reference(field(&data, "range")?)?,
            formula: field(&data, "formula")?.into(),
            proof: data["proof"].clone(),
        };
        crate::spill::schema(&tx)?;
        crate::spill::record(&tx, &group, true)?;
        tx.execute(
            "INSERT INTO publication_spill_certificates VALUES(?1,?2,?3,?4,?5,?6)",
            params![
                accepted,
                sheet,
                master.0,
                master.1,
                revision,
                group_subject(&group).to_string()
            ],
        )
        .map_err(err)?;
        groups += 1;
    }
    tx.execute(
        "INSERT OR REPLACE INTO calculation_policies VALUES(?1,?2,?3)",
        params![accepted, options.0, options.1],
    )
    .map_err(err)?;
    crate::dependencies::schema(&tx)?;
    tx.execute(
        "INSERT OR REPLACE INTO dependency_bindings VALUES(?1,0)",
        [accepted],
    )
    .map_err(err)?;
    tx.commit().map_err(err)?;
    let status = crate::calculation_policy::command(
        root,
        &json!({"method":"status","version":accepted,"view":view}),
    )?;
    Ok(
        json!({"schema":workbook_format!(SCHEMA_PUBLICATION_CALCULATION_RESTORE_V1),"version":accepted,"policy":status,"transferredFormulaCells":formulas,"transferredDerivedCells":derived,"transferredSpillGroups":groups}),
    )
}

#[cfg(test)]
mod publication_tests {
    use super::*;
    use std::io::Write;
    const SHEET: &str = "xl/worksheets/main.xml";
    fn source(root: &Path, on_save: bool) {
        fs::create_dir_all(root).unwrap();
        let mut zip = zip::ZipWriter::new(fs::File::create(root.join("source.xlsx")).unwrap());
        let workbook = format!(
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Main\" sheetId=\"1\" r:id=\"main\"/></sheets><calcPr calcId=\"0\" calcMode=\"manual\" calcOnSave=\"{}\" fullCalcOnLoad=\"0\" forceFullCalc=\"0\"/></workbook>",
            u8::from(on_save)
        );
        for (name, text) in [
            (
                "[Content_Types].xml",
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Override PartName=\"/xl/worksheets/main.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>",
            ),
            ("xl/workbook.xml", workbook.as_str()),
            (
                "xl/_rels/workbook.xml.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"main\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/main.xml\"/></Relationships>",
            ),
            (
                SHEET,
                "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"A1:D1\"/><sheetData><row r=\"1\"><c r=\"A1\"><v>2</v></c><c r=\"B1\"><f>A1+1</f><v>91</v></c><c r=\"C1\"><f>IF(A1=2,TRUE,FALSE)</f><v>92</v></c><c r=\"D1\"><f>1/0</f><v>93</v></c></row></sheetData></worksheet>",
            ),
            (
                "xl/styles.xml",
                "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts count=\"1\"><font/></fonts><fills count=\"1\"><fill/></fills><borders count=\"1\"><border/></borders><cellXfs count=\"1\"><xf numFmtId=\"0\"/></cellXfs></styleSheet>",
            ),
        ] {
            zip.start_file(name, zip::write::SimpleFileOptions::default())
                .unwrap();
            zip.write_all(text.as_bytes()).unwrap();
        }
        zip.finish().unwrap();
        crate::import::index(root).unwrap();
    }
    fn view() -> Value {
        json!({"sheets":[{"name":"Main","path":SHEET,"storedFormulaCount":3,"date1904":false,"merges":[]}],"namedRanges":[],"tables":[]})
    }
    fn initialize(root: &Path, on_save: bool) {
        crate::calculation_policy::command(root,&json!({"method":"initialize","version":0,"view":view(),"initialPolicy":{"mode":"manual","recalculateBeforeSave":on_save}})).unwrap();
    }
    fn calculate(root: &Path) -> u64 {
        let result = crate::calculation::command(
            root,
            &json!({"version":0,"view":view(),"revision":"accepted-calculation","partial":true}),
        )
        .unwrap();
        assert_eq!(result["supported"], true);
        let version = result["version"].as_u64().unwrap();
        // Model the trusted compatibility owner committing a typed error after
        // the native partial scalar pass; this tests transfer, not the evaluator.
        let mut error = read_cell(&crate::database(root).unwrap(), version, SHEET, "D1").unwrap();
        error["calculation"] =
            json!({"status":"error","error":"#DIV/0!","revision":"accepted-calculation"});
        let patched=crate::range::patch(root,&json!({"version":version,"changes":[{"sheet":SHEET,"cell":"D1","value":error,"edit":{}}]})).unwrap();
        let version = patched["version"].as_u64().unwrap();
        crate::calculation_policy::command(
            root,
            &json!({"method":"settle","version":version,"view":view()}),
        )
        .unwrap();
        version
    }
    fn exported(root: &Path, version: u64, publication: bool) -> Value {
        let db = crate::database(root).unwrap();
        let options = policy(root, &db, version).unwrap();
        let pending = crate::calculation_policy::pending_count(&db, version).unwrap() > 0;
        drop(db);
        crate::export_state::prepare(root, &json!({"version":version,"view":view()})).unwrap();
        let source = source_index(root).unwrap();
        let mut zip =
            zip::ZipArchive::new(fs::File::open(root.join("source.xlsx")).unwrap()).unwrap();
        let mut workbook = String::new();
        std::io::Read::read_to_string(&mut zip.by_name("xl/workbook.xml").unwrap(), &mut workbook)
            .unwrap();
        workbook = workbook
            .replace(
                "fullCalcOnLoad=\"0\"",
                &format!("fullCalcOnLoad=\"{}\"", u8::from(pending)),
            )
            .replace(
                "forceFullCalc=\"0\"",
                &format!("forceFullCalc=\"{}\"", u8::from(pending)),
            );
        assert_eq!(options.0, "manual");
        fs::write(
            root.join("export-metadata.json"),
            json!({"styleIds":[],"replacements":{"xl/workbook.xml":workbook}}).to_string(),
        )
        .unwrap();
        crate::export::write(root,&json!({"revision":source["revision"],"version":version,"calculationMetadata":true,"publication":publication})).unwrap()
    }
    fn target(old: &Path, new: &Path, on_save: bool) {
        fs::create_dir(new).unwrap();
        fs::copy(old.join("export.xlsx"), new.join("source.xlsx")).unwrap();
        crate::import::index(new).unwrap();
        initialize(new, on_save);
    }
    fn transfer(old: &Path, new: &Path, version: u64) -> Value {
        command(new,&json!({"action":"workbook","operation":"calculation-transfer","purpose":"publication","root":new,"version":0,"previousRoot":old,"previousVersion":version,"sourceRevision":source_index(new).unwrap()["revision"],"view":view()})).unwrap()
    }
    fn restore(new: &Path, witness: &Path, hash: &str) -> Value {
        command(new,&json!({"action":"workbook","operation":"calculation-transfer","purpose":"publication-restore","root":new,"version":0,"sourceRevision":source_index(new).unwrap()["revision"],"witness":witness,"witnessRevision":hash,"view":view()})).unwrap()
    }
    fn footprint(root: &Path) -> std::collections::BTreeMap<String, String> {
        fs::read_dir(root)
            .unwrap()
            .map(|entry| {
                let entry = entry.unwrap();
                (
                    entry.file_name().to_string_lossy().into_owned(),
                    crate::hash(&entry.path()).unwrap(),
                )
            })
            .collect()
    }
    #[test]
    fn publication_preserves_current_typed_values_errors_and_old_owner_bytes() {
        let _scope = crate::shared::begin();
        let owner = tempfile::tempdir().unwrap();
        let old = owner.path().join("base");
        let new = owner.path().join("publication");
        source(&old, true);
        initialize(&old, true);
        let version = calculate(&old);
        exported(&old, version, true);
        target(&old, &new, true);
        let before = footprint(&old);
        let result = transfer(&old, &new, version);
        assert_eq!(footprint(&old), before);
        assert_eq!(result["policy"]["pendingFormulaCells"], 0);
        assert_eq!(result["transferredFormulaCells"], 3);
        let db = crate::database(&new).unwrap();
        let version = result["version"].as_u64().unwrap();
        let b = read_cell(&db, version, SHEET, "B1").unwrap();
        assert_eq!(b["calculation"]["value"], 3.0);
        let c = read_cell(&db, version, SHEET, "C1").unwrap();
        assert_eq!(c["calculation"]["value"], true);
        let d = read_cell(&db, version, SHEET, "D1").unwrap();
        assert_eq!(d["calculation"]["error"], "#DIV/0!");
        assert_eq!(result["witness"]["schema"], 1);
        assert_eq!(
            result["witness"]["revision"],
            crate::hash(&new.join("publication-calculation.sqlite")).unwrap()
        );
    }
    #[test]
    fn publication_witness_restores_exact_authority_without_evaluation() {
        let _scope = crate::shared::begin();
        let owner = tempfile::tempdir().unwrap();
        let old = owner.path().join("base");
        let new = owner.path().join("publication");
        let cold = owner.path().join("cold");
        source(&old, true);
        initialize(&old, true);
        let version = calculate(&old);
        exported(&old, version, true);
        target(&old, &new, true);
        let accepted = transfer(&old, &new, version);
        target(&old, &cold, true);
        let result = restore(
            &cold,
            &new.join("publication-calculation.sqlite"),
            accepted["witness"]["revision"].as_str().unwrap(),
        );
        let db = crate::database(&cold).unwrap();
        let version = result["version"].as_u64().unwrap();
        assert_eq!(result["policy"]["pendingFormulaCells"], 0);
        assert_eq!(
            read_cell(&db, version, SHEET, "B1").unwrap()["calculation"]["revision"],
            "accepted-calculation"
        );
        assert_eq!(
            read_cell(&db, version, SHEET, "C1").unwrap()["calculation"]["value"],
            true
        );
    }
    #[test]
    fn manual_false_publication_and_witness_keep_stored_caches_pending() {
        let _scope = crate::shared::begin();
        let owner = tempfile::tempdir().unwrap();
        let old = owner.path().join("base");
        let new = owner.path().join("publication");
        let cold = owner.path().join("cold");
        source(&old, false);
        initialize(&old, false);
        exported(&old, 0, true);
        target(&old, &new, false);
        let accepted = transfer(&old, &new, 0);
        assert_eq!(accepted["policy"]["pendingFormulaCells"], 3);
        target(&old, &cold, false);
        let result = restore(
            &cold,
            &new.join("publication-calculation.sqlite"),
            accepted["witness"]["revision"].as_str().unwrap(),
        );
        let db = crate::database(&cold).unwrap();
        let page = crate::range::read_public_db(
            &db,
            &json!({"version":result["version"],"sheet":SHEET,"range":"B1","limit":1}),
        )
        .unwrap();
        assert_eq!(page["cells"][0][1]["calculation"]["status"], "stored");
        assert_eq!(page["cells"][0][1]["cachedResult"]["value"], 91.0);
        assert_eq!(result["policy"]["pendingFormulaCells"], 3);
    }
    #[test]
    fn publication_and_restore_preserve_partial_current_and_pending_authority() {
        let _scope = crate::shared::begin();
        let owner = tempfile::tempdir().unwrap();
        let old = owner.path().join("base");
        let new = owner.path().join("publication");
        let cold = owner.path().join("cold");
        source(&old, false);
        initialize(&old, false);
        let version = calculate(&old);
        let edited=crate::range::patch(&old,&json!({"version":version,"changes":[{"sheet":SHEET,"cell":"A1","value":{"value":5,"dataType":"number","rawValue":"5","cachedValue":null,"styleId":0}}]})).unwrap()["version"].as_u64().unwrap();
        crate::calculation_policy::command(&old,&json!({"method":"defer","version":edited,"view":view(),"changedRanges":[{"sheet":"Main","range":"A1"}]})).unwrap();
        exported(&old, edited, true);
        target(&old, &new, false);
        let accepted = transfer(&old, &new, edited);
        assert_eq!(accepted["policy"]["pendingFormulaCells"], 2);
        target(&old, &cold, false);
        let restored = restore(
            &cold,
            &new.join("publication-calculation.sqlite"),
            accepted["witness"]["revision"].as_str().unwrap(),
        );
        assert_eq!(restored["policy"]["pendingFormulaCells"], 2);
        let db = crate::database(&cold).unwrap();
        let page = crate::range::read_public_db(
            &db,
            &json!({"version":restored["version"],"sheet":SHEET,"range":"A1:D1","limit":4}),
        )
        .unwrap();
        let cells = page["cells"].as_array().unwrap();
        assert_eq!(cells[0][1]["value"], 5.0);
        assert_eq!(cells[1][1]["cachedResult"]["value"], 3.0);
        assert_eq!(cells[1][1]["calculation"]["pending"], true);
        assert_eq!(cells[2][1]["cachedResult"]["value"], true);
        assert_eq!(cells[3][1]["calculation"]["status"], "error");
        assert_eq!(cells[3][1]["calculation"]["error"], "#DIV/0!");
    }
    #[test]
    fn accepted_volatile_output_is_carried_without_an_evaluator_call() {
        let _scope = crate::shared::begin();
        let owner = tempfile::tempdir().unwrap();
        let old = owner.path().join("base");
        let new = owner.path().join("publication");
        let cold = owner.path().join("cold");
        source(&old, true);
        initialize(&old, true);
        let version = calculate(&old);
        // RAND is outside this evaluator's admitted grammar. A seeded accepted
        // trusted-owner result demonstrates that transfer never evaluates it.
        let mut cell = read_cell(&crate::database(&old).unwrap(), version, SHEET, "B1").unwrap();
        cell["formula"] = json!("RAND()");
        cell["calculation"] = json!({"status":"calculated","value":0.125,"revision":"single-accepted-volatile-result"});
        let version = crate::range::patch(
            &old,
            &json!({"version":version,"changes":[{"sheet":SHEET,"cell":"B1","value":cell}]}),
        )
        .unwrap()["version"]
            .as_u64()
            .unwrap();
        crate::calculation_policy::command(
            &old,
            &json!({"method":"settle","version":version,"view":view()}),
        )
        .unwrap();
        exported(&old, version, true);
        target(&old, &new, true);
        let accepted = transfer(&old, &new, version);
        target(&old, &cold, true);
        let result = restore(
            &cold,
            &new.join("publication-calculation.sqlite"),
            accepted["witness"]["revision"].as_str().unwrap(),
        );
        let db = crate::database(&cold).unwrap();
        let carried = read_cell(&db, result["version"].as_u64().unwrap(), SHEET, "B1").unwrap();
        assert_eq!(carried["formula"], "RAND()");
        assert_eq!(carried["calculation"]["value"], 0.125);
        assert_eq!(
            carried["calculation"]["revision"],
            "single-accepted-volatile-result"
        );
    }
    #[test]
    fn owned_spill_storage_spelling_requires_complete_native_ownership() {
        let _scope = crate::shared::begin();
        let owner = tempfile::tempdir().unwrap();
        let root = owner.path().join("base");
        source(&root, false);
        initialize(&root, false);
        let mut db = crate::database(&root).unwrap();
        let tx = db.transaction().unwrap();
        crate::calculation::snapshot(&tx, 0).unwrap();
        tx.execute("INSERT INTO versions(parent)VALUES(0)", [])
            .unwrap();
        let version = tx.last_insert_rowid() as u64;
        let group = crate::spill::Group {
            version,
            sheet: SHEET.into(),
            master: (3, 0),
            geometry: crate::spill::Geometry::from_reference("A4:B6").unwrap(),
            formula: " ( SeQuEnCe(3,2) )".into(),
            proof: json!({"constructor":"sequence-integer-v1","arguments":[3,2,1,1],"dependencies":[],"status":"calculated","qualification":"retained-sequence-v1"}),
        };
        crate::spill::schema(&tx).unwrap();
        crate::spill::publish(&tx, &group, None, "single-owned-cache").unwrap();
        tx.commit().unwrap();
        let cell = read_cell(&db, version, SHEET, "A4").unwrap();
        let mut exported = cell.clone();
        exported["formula"] = json!(crate::export::spill_storage_formula(&group.formula));
        assert_eq!(exported["formula"], " ( _xlfn.SeQuEnCe(3,2) )");
        let before = crate::hash(&root.join("cells.sqlite")).unwrap();
        assert!(owned_publication_formula(&db, version, SHEET, "A4", &cell, &exported).unwrap());
        assert_eq!(crate::hash(&root.join("cells.sqlite")).unwrap(), before);
        let mut unowned = cell.clone();
        unowned.as_object_mut().unwrap().remove("spill");
        assert!(
            !owned_publication_formula(&db, version, SHEET, "A4", &unowned, &exported).unwrap()
        );
        let mut wrong = exported.clone();
        wrong["formula"] = json!("_xlfn.SEQUENCE(2,2)");
        assert!(!owned_publication_formula(&db, version, SHEET, "A4", &cell, &wrong).unwrap());
        wrong = exported.clone();
        wrong["formulaReference"] = json!("A4:B5");
        assert!(!owned_publication_formula(&db, version, SHEET, "A4", &cell, &wrong).unwrap());
        // The carried group uses the same serialized spelling as its actual
        // imported master, while preserving all six exact owned cache values.
        let tx = db.transaction().unwrap();
        tx.execute("INSERT INTO versions(parent)VALUES(?1)", [version])
            .unwrap();
        let next = tx.last_insert_rowid() as u64;
        for at in group.geometry.positions() {
            let mut carried =
                read_cell(&tx, version, SHEET, &crate::xml::address(at.0, at.1)).unwrap();
            if at == group.master {
                carried["formula"] = exported["formula"].clone();
            }
            crate::range::store_calculated_change(
                &tx,
                next as i64,
                SHEET,
                at.0,
                at.1,
                &carried,
                &json!({}),
            )
            .unwrap();
        }
        let mut rebound = group.clone();
        rebound.version = next;
        rebound.formula = crate::export::spill_storage_formula(&group.formula);
        crate::spill::record(&tx, &rebound, true).unwrap();
        tx.commit().unwrap();
        assert!(
            crate::spill::accepted_cell(
                &db,
                next,
                SHEET,
                group.master,
                &read_cell(&db, next, SHEET, "A4").unwrap()
            )
            .unwrap()
        );
        assert_eq!(read_cell(&db, next, SHEET, "B6").unwrap()["value"], 6.0);
        assert_eq!(
            read_cell(&db, version, SHEET, "A4").unwrap()["formula"],
            group.formula
        );
    }
    #[test]
    fn copy_export_does_not_issue_publication_authority() {
        let _scope = crate::shared::begin();
        let owner = tempfile::tempdir().unwrap();
        let old = owner.path().join("base");
        let new = owner.path().join("publication");
        source(&old, true);
        initialize(&old, true);
        let version = calculate(&old);
        exported(&old, version, false);
        target(&old, &new, true);
        let before = footprint(&new);
        let error=command(&new,&json!({"action":"workbook","operation":"calculation-transfer","purpose":"publication","root":new,"version":0,"previousRoot":old,"previousVersion":version,"sourceRevision":source_index(&new).unwrap()["revision"],"view":view()})).unwrap_err();
        assert!(error.contains("no native export receipt"));
        assert_eq!(footprint(&new), before);
        assert!(
            !has_table(
                &crate::database(&old).unwrap(),
                "publication_export_receipts"
            )
            .unwrap()
        );
    }
    #[test]
    fn foreign_owner_wrong_selected_version_and_changed_witness_refuse_atomically() {
        let _scope = crate::shared::begin();
        let owner = tempfile::tempdir().unwrap();
        let foreign = tempfile::tempdir().unwrap();
        let old = owner.path().join("base");
        let new = owner.path().join("publication");
        source(&old, true);
        initialize(&old, true);
        let version = calculate(&old);
        exported(&old, version, true);
        target(&old, &new, true);
        let request = json!({"action":"workbook","operation":"calculation-transfer","purpose":"publication","root":new,"version":0,"previousRoot":old,"previousVersion":version,"sourceRevision":source_index(&new).unwrap()["revision"],"view":view()});
        let before = footprint(&new);
        for mutation in [
            json!({"previousVersion":0}),
            json!({"sourceRevision":"f".repeat(64)}),
            json!({"unexpected":true}),
        ] {
            let mut changed = request.clone();
            changed
                .as_object_mut()
                .unwrap()
                .extend(mutation.as_object().unwrap().clone());
            assert!(command(&new, &changed).is_err());
            assert_eq!(footprint(&new), before);
        }
        let abroad = foreign.path().join("publication");
        target(&old, &abroad, true);
        let mut changed = request.clone();
        changed["root"] = json!(abroad);
        assert!(
            command(&abroad, &changed)
                .unwrap_err()
                .contains("same owner")
        );
        let accepted = transfer(&old, &new, version);
        let cold = owner.path().join("cold");
        target(&old, &cold, true);
        let before = footprint(&cold);
        let witness = new.join("publication-calculation.sqlite");
        let good_hash = accepted["witness"]["revision"].as_str().unwrap();
        let error=command(&cold,&json!({"action":"workbook","operation":"calculation-transfer","purpose":"publication-restore","root":cold,"version":0,"sourceRevision":source_index(&cold).unwrap()["revision"],"witness":witness,"witnessRevision":"f".repeat(64),"view":view()})).unwrap_err();
        assert!(error.contains("revision mismatch"));
        assert_eq!(footprint(&cold), before);
        assert_eq!(crate::hash(&witness).unwrap(), good_hash);
    }
    #[test]
    fn traced_publication_and_restore_keep_partial_authority_and_diagnostics() {
        let _scope = crate::shared::begin();
        for diagnostics in [true, false] {
            let owner = tempfile::tempdir().unwrap();
            let old = owner.path().join("base");
            let new = owner.path().join("publication");
            let cold = owner.path().join("cold");
            source(&old, false);
            initialize(&old, false);
            let version = calculate(&old);
            let edited=crate::range::patch(&old,&json!({"version":version,"changes":[{"sheet":SHEET,"cell":"A1","value":{"value":5,"dataType":"number","rawValue":"5","cachedValue":null,"styleId":0}}]})).unwrap()["version"].as_u64().unwrap();
            crate::calculation_policy::command(&old,&json!({"method":"defer","version":edited,"view":view(),"changedRanges":[{"sheet":"Main","range":"A1"}]})).unwrap();
            exported(&old, edited, true);
            target(&old, &new, false);
            let before = footprint(&old);
            let accepted=crate::command(json!({"action":"workbook","operation":"calculation-transfer","purpose":"publication","root":new,"version":0,"previousRoot":old,"previousVersion":edited,"sourceRevision":source_index(&new).unwrap()["revision"],"view":view(),"diagnostics":diagnostics})).unwrap();
            assert_eq!(footprint(&old), before);
            assert_eq!(accepted["policy"]["pendingFormulaCells"], 2);
            assert_eq!(accepted["transferredFormulaCells"], 3);
            assert_eq!(accepted.get("diagnostics").is_some(), diagnostics);
            if diagnostics {
                assert_eq!(accepted["diagnostics"]["schema"], 1);
            }
            target(&old, &cold, false);
            let witness = new.join("publication-calculation.sqlite");
            let witness_before = crate::hash(&witness).unwrap();
            let restored=crate::command(json!({"action":"workbook","operation":"calculation-transfer","purpose":"publication-restore","root":cold,"version":0,"sourceRevision":source_index(&cold).unwrap()["revision"],"witness":witness,"witnessRevision":accepted["witness"]["revision"],"view":view(),"diagnostics":diagnostics})).unwrap();
            assert_eq!(crate::hash(&witness).unwrap(), witness_before);
            assert_eq!(footprint(&old), before);
            assert_eq!(restored["policy"]["pendingFormulaCells"], 2);
            assert_eq!(restored.get("diagnostics").is_some(), diagnostics);
            if diagnostics {
                assert_eq!(restored["diagnostics"]["schema"], 1);
            }
            let db = crate::database(&cold).unwrap();
            let page = crate::range::read_public_db(
                &db,
                &json!({"version":restored["version"],"sheet":SHEET,"range":"A1:D1","limit":4}),
            )
            .unwrap();
            let cells = page["cells"].as_array().unwrap();
            assert_eq!(cells[0][1]["value"], 5.0);
            assert_eq!(cells[1][1]["cachedResult"]["value"], 3.0);
            assert_eq!(cells[1][1]["calculation"]["pending"], true);
            assert_eq!(cells[2][1]["cachedResult"]["value"], true);
            assert_eq!(cells[3][1]["calculation"]["error"], "#DIV/0!");
        }
    }
    #[test]
    fn diagnostics_admission_is_boolean_only_and_keeps_all_transfer_fields_strict() {
        let _scope = crate::shared::begin();
        let owner = tempfile::tempdir().unwrap();
        let old = owner.path().join("base");
        let new = owner.path().join("publication");
        let cold = owner.path().join("cold");
        source(&old, true);
        initialize(&old, true);
        let version = calculate(&old);
        exported(&old, version, true);
        target(&old, &new, true);
        let publication = json!({"action":"workbook","operation":"calculation-transfer","purpose":"publication","root":new,"version":0,"previousRoot":old,"previousVersion":version,"sourceRevision":source_index(&new).unwrap()["revision"],"view":view()});
        let before = footprint(&new);
        for diagnostic in [Value::Null, json!(1), json!("true"), json!([]), json!({})] {
            let mut request = publication.clone();
            request["diagnostics"] = diagnostic;
            assert!(
                command(&new, &request)
                    .unwrap_err()
                    .contains("diagnostics must be a boolean")
            );
            assert_eq!(footprint(&new), before);
        }
        let mut unexpected = publication.clone();
        unexpected["diagnostics"] = json!(true);
        unexpected["sourcePath"] = json!("/foreign");
        assert!(
            command(&new, &unexpected)
                .unwrap_err()
                .contains("transfer fields")
        );
        assert_eq!(footprint(&new), before);
        let mut missing = publication.clone();
        missing["diagnostics"] = json!(true);
        missing.as_object_mut().unwrap().remove("previousVersion");
        assert!(
            command(&new, &missing)
                .unwrap_err()
                .contains("transfer fields")
        );
        assert_eq!(footprint(&new), before);
        let accepted = transfer(&old, &new, version);
        target(&old, &cold, true);
        let restoration = json!({"action":"workbook","operation":"calculation-transfer","purpose":"publication-restore","root":cold,"version":0,"sourceRevision":source_index(&cold).unwrap()["revision"],"witness":new.join("publication-calculation.sqlite"),"witnessRevision":accepted["witness"]["revision"],"view":view()});
        let before = footprint(&cold);
        for diagnostic in [Value::Null, json!(1), json!("false"), json!([]), json!({})] {
            let mut request = restoration.clone();
            request["diagnostics"] = diagnostic;
            assert!(
                command(&cold, &request)
                    .unwrap_err()
                    .contains("diagnostics must be a boolean")
            );
            assert_eq!(footprint(&cold), before);
        }
        let mut unexpected = restoration.clone();
        unexpected["diagnostics"] = json!(false);
        unexpected["current"] = json!(true);
        assert!(
            command(&cold, &unexpected)
                .unwrap_err()
                .contains("restoration fields")
        );
        assert_eq!(footprint(&cold), before);
        let pivot = json!({"action":"workbook","operation":"calculation-transfer","root":new,"version":0,"previousRoot":old,"previousVersion":version,"sheet":"Main","view":view()});
        let required = [
            "action",
            "operation",
            "root",
            "version",
            "previousRoot",
            "previousVersion",
            "sheet",
            "view",
        ];
        for diagnostics in [true, false] {
            let mut request = pivot.clone();
            request["diagnostics"] = json!(diagnostics);
            transfer_fields(&request, &required, "Invalid calculation transfer fields").unwrap();
            request["sourceRevision"] = json!("f".repeat(64));
            assert!(
                transfer_fields(&request, &required, "Invalid calculation transfer fields")
                    .is_err()
            );
        }
    }
}
