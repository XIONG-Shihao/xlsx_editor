//! Sparse source-text features fence destructive decoded value replacement.
use crate::{Result, err};
use rusqlite::{Connection, OptionalExtension};
pub(crate) const REASON: &str = "Rich text or annotated/opaque value metadata is preserved; content editing requires preservation-aware support";
pub(crate) fn safe_text_attribute(name: &str) -> bool {
    name == "xml:space" || name == "xmlns" || name.starts_with("xmlns:")
}
pub(crate) fn reason(db: &Connection, sheet: &str, row: u32, col: u32) -> Result<Option<String>> {
    reason_for(db, sheet, row, col, false)
}
pub(crate) fn cache_reason(
    db: &Connection,
    sheet: &str,
    row: u32,
    col: u32,
) -> Result<Option<String>> {
    reason_for(db, sheet, row, col, true)
}
fn reason_for(
    db: &Connection,
    sheet: &str,
    row: u32,
    col: u32,
    cache: bool,
) -> Result<Option<String>> {
    let kind: Option<String> = db
        .prepare_cached("SELECT kind FROM cells WHERE sheet=?1 AND r=?2 AND c=?3")
        .map_err(err)?
        .query_row(rusqlite::params![sheet, row, col], |r| r.get(0))
        .optional()
        .map_err(err)?;
    if kind.is_none() {
        return Ok(None);
    }
    let indexed: bool = db.prepare_cached("SELECT COUNT(*)=2 FROM sqlite_master WHERE type='table' AND name IN ('value_features','shared_value_features')").map_err(crate::err)?.query_row([],|r|r.get(0)).map_err(err)?;
    if !indexed {
        if !cache && !kind.is_some_and(|kind| kind == "s" || kind == "inlineStr") {
            return Ok(None);
        }
        return Err(crate::coded(
            "internal-missing-field",
            "Source text preservation metadata is missing; a fresh workbook index is required before content editing",
        ));
    }
    db.prepare_cached("SELECT CASE WHEN v.mask & ?4 != 0 THEN v.reason WHEN ?4=1 THEN s.reason ELSE NULL END FROM cells c LEFT JOIN value_features v ON v.sheet=c.sheet AND v.r=c.r AND v.c=c.c LEFT JOIN shared_value_features s ON c.kind='s' AND s.id=CAST(c.raw AS INTEGER) WHERE c.sheet=?1 AND c.r=?2 AND c.c=?3").map_err(err)?.query_row(rusqlite::params![sheet,row,col,if cache {2} else {1}],|r|r.get(0)).map_err(err)
}
#[cfg(test)]
mod tests {
    use super::*;
    fn fixture() -> Connection {
        let db = Connection::open_in_memory().unwrap();
        db.execute_batch("CREATE TABLE cells(sheet TEXT,r INTEGER,c INTEGER,kind TEXT,raw TEXT);INSERT INTO cells VALUES('S',0,0,'inlineStr',''),('S',0,1,'s','2'),('S',0,2,'n','7');CREATE TABLE value_features(sheet TEXT,r INTEGER,c INTEGER,reason TEXT,mask INTEGER DEFAULT 1,PRIMARY KEY(sheet,r,c));CREATE TABLE shared_value_features(id INTEGER PRIMARY KEY,reason TEXT);INSERT INTO value_features VALUES('S',0,0,'inline',1);INSERT INTO shared_value_features VALUES(2,'shared');").unwrap();
        db
    }
    #[test]
    fn source_features_are_selected_by_exact_source_coordinate_and_shared_identity() {
        let db = fixture();
        assert_eq!(reason(&db, "S", 0, 0).unwrap(), Some("inline".into()));
        assert_eq!(reason(&db, "S", 0, 1).unwrap(), Some("shared".into()));
        assert_eq!(reason(&db, "S", 0, 2).unwrap(), None);
        assert_eq!(reason(&db, "S", 1, 0).unwrap(), None);
    }
    #[test]
    fn missing_legacy_text_index_refuses_instead_of_inventing_safe_authority() {
        let db = fixture();
        db.execute_batch("DROP TABLE shared_value_features")
            .unwrap();
        assert!(
            reason(&db, "S", 0, 1)
                .unwrap_err()
                .contains("fresh workbook index")
        );
        assert_eq!(reason(&db, "S", 0, 2).unwrap(), None);
    }
    #[test]
    fn value_masks_distinguish_content_from_cache_writes() {
        let db = fixture();
        db.execute_batch("INSERT INTO value_features VALUES('S',0,2,'opaque cache',3)")
            .unwrap();
        assert_eq!(
            cache_reason(&db, "S", 0, 2).unwrap(),
            Some("opaque cache".into())
        );
        assert_eq!(cache_reason(&db, "S", 0, 0).unwrap(), None);
        assert_eq!(reason(&db, "S", 0, 2).unwrap(), Some("opaque cache".into()));
    }
    #[test]
    fn namespaces_and_space_are_not_format_authority() {
        assert!(safe_text_attribute("xml:space"));
        assert!(safe_text_attribute("xmlns:u"));
        assert!(!safe_text_attribute("xml:lang"));
        assert!(!safe_text_attribute("u:tag"));
    }
}

/// No shared expansion or formula evaluation is needed to protect source geometry.
pub(crate) fn region(
    root: &std::path::Path,
    db: &Connection,
    sheet: &str,
    first: (u32, u32),
    last: (u32, u32),
    content: bool,
) -> Result<()> {
    let index: serde_json::Value =
        serde_json::from_slice(&std::fs::read(root.join("index.json")).map_err(err)?)
            .map_err(err)?;
    if !index["sheets"].as_array().is_some_and(|s| {
        s.iter()
            .any(|s| s["path"] == sheet && s["dataStart"].is_number())
    }) {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown or protected preservation worksheet",
        ));
    }
    let positions = u64::from(last.0 - first.0 + 1) * u64::from(last.1 - first.1 + 1);
    let stored:u64=db.prepare_cached("SELECT count(*) FROM cells WHERE sheet=?1 AND r BETWEEN ?2 AND ?3 AND c BETWEEN ?4 AND ?5").unwrap().query_row(rusqlite::params![sheet,first.0,last.0,first.1,last.1],|r|r.get(0)).map_err(err)?;
    if stored < positions && !crate::used_range::guard_reasons(db, 0, sheet)?.is_empty() {
        return Err(crate::coded(
            "workbook-unsupported",
            "New cells require a qualified source worksheet extent",
        ));
    }
    if !content {
        return Ok(());
    }
    if db.prepare_cached("SELECT EXISTS(SELECT 1 FROM cells WHERE sheet=?1 AND r BETWEEN ?2 AND ?3 AND c BETWEEN ?4 AND ?5 AND (formula IS NOT NULL OR formula_attrs IS NOT NULL))").unwrap().query_row(rusqlite::params![sheet,first.0,last.0,first.1,last.1],|r|r.get::<_,bool>(0)).map_err(err)?{return Err(crate::coded("formula-protected", "Source formula cells remain protected while calculation is preserved"));}
    if db.prepare_cached("SELECT EXISTS(SELECT 1 FROM cells c LEFT JOIN value_features v ON v.sheet=c.sheet AND v.r=c.r AND v.c=c.c LEFT JOIN shared_value_features s ON c.kind='s' AND s.id=CAST(c.raw AS INTEGER) WHERE c.sheet=?1 AND c.r BETWEEN ?2 AND ?3 AND c.c BETWEEN ?4 AND ?5 AND (v.mask&1!=0 OR s.reason IS NOT NULL))").unwrap().query_row(rusqlite::params![sheet,first.0,last.0,first.1,last.1],|r|r.get::<_,bool>(0)).map_err(err)?{return Err(crate::coded("workbook-refused", REASON));}
    let mut q=db.prepare_cached("SELECT formula_attrs FROM cells WHERE sheet=?1 AND formula_attrs IS NOT NULL AND json_type(formula_attrs,'$.ref')='text'").map_err(err)?;
    let mut rows = q.query([sheet]).map_err(err)?;
    while let Some(record) = rows.next().map_err(err)? {
        crate::cancel::check()?;
        let value: serde_json::Value =
            serde_json::from_str(&record.get::<_, String>(0).map_err(err)?).map_err(err)?;
        let (a, b) = crate::bulk::bounds(value["ref"].as_str().ok_or_else(|| {
            crate::coded("internal-invalid-field", "Invalid preserved formula range")
        })?)?;
        if first.0 <= b.0 && last.0 >= a.0 && first.1 <= b.1 && last.1 >= a.1 {
            return Err(crate::coded(
                "table-unsupported",
                "A preserved shared, array or data-table formula region cannot be edited",
            ));
        }
    }
    Ok(())
}
pub(crate) fn check(
    root: &std::path::Path,
    input: &serde_json::Value,
) -> Result<serde_json::Value> {
    let db = crate::database(root)?;
    let version = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing preservation version"))?;
    if !db
        .prepare_cached("SELECT EXISTS(SELECT 1 FROM versions WHERE id=?1)")
        .unwrap()
        .query_row([version], |r| r.get::<_, bool>(0))
        .map_err(err)?
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown preservation version",
        ));
    }
    let bounds = crate::bulk::bounds(
        input["range"]
            .as_str()
            .or(input["cell"].as_str())
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing preservation range"))?,
    )?;
    if u64::from(bounds.1.0 - bounds.0.0 + 1) * u64::from(bounds.1.1 - bounds.0.1 + 1) > 10000 {
        return Err(crate::coded(
            "workbook-limit",
            "Preserved range exceeds 10000 positions",
        ));
    }
    if input.get("content").is_some_and(|v| !v.is_boolean()) {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid preservation content flag",
        ));
    }
    region(
        root,
        &db,
        crate::field(input, "sheet")?,
        bounds.0,
        bounds.1,
        input["content"] != false,
    )?;
    Ok(serde_json::json!({"allowed":true}))
}
pub(crate) fn source_content(
    db: &Connection,
    sheet: &str,
    r: u32,
    c: u32,
) -> Result<serde_json::Value> {
    let mut q=db.prepare_cached("SELECT c.r,c.c,c.kind,c.raw,c.inline,c.formula,c.formula_attrs,c.style,s.value,NULL FROM cells c LEFT JOIN strings s ON c.kind='s' AND s.id=CAST(c.raw AS INTEGER) WHERE c.sheet=?1 AND c.r=?2 AND c.c=?3").map_err(err)?;
    let mut rows = q.query(rusqlite::params![sheet, r, c]).map_err(err)?;
    let cell = if let Some(record) = rows.next().map_err(err)? {
        crate::range::decode_cell(record)?.1
    } else {
        serde_json::json!({"value":null,"dataType":"blank","cachedValue":null,"rawValue":null})
    };
    Ok(content_fields(&cell))
}
pub(crate) fn content_fields(cell: &serde_json::Value) -> serde_json::Value {
    let mut out = serde_json::json!({});
    for key in [
        "value",
        "dataType",
        "rawValue",
        "cachedValue",
        "formula",
        "formulaType",
        "formulaReference",
        "sharedFormulaId",
        "error",
    ] {
        if let Some(value) = cell.get(key) {
            out[key] = value.clone();
        }
    }
    for key in ["rawValue", "cachedValue"] {
        if out.get(key).is_none() {
            out[key] = serde_json::Value::Null;
        }
    }
    out
}
/// Export authority is derived from the complete accepted ancestry, not a UI flag.
pub(crate) fn export_edits(root: &std::path::Path, db: &Connection, version: u64) -> Result<()> {
    let mut q=db.prepare_cached("WITH RECURSIVE a(id,depth)AS(SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN a ON v.id=a.id WHERE v.parent IS NOT NULL),p AS(SELECT c.*,row_number()OVER(PARTITION BY sheet,r,c ORDER BY a.depth)priority FROM a JOIN changes c ON c.version=a.id) SELECT sheet,r,c,data,edit FROM p WHERE priority=1").map_err(err)?;
    let mut rows = q.query([version]).map_err(err)?;
    while let Some(record) = rows.next().map_err(err)? {
        crate::cancel::check()?;
        let sheet: String = record.get(0).map_err(err)?;
        let (r, c) = (
            record.get::<_, u32>(1).map_err(err)?,
            record.get::<_, u32>(2).map_err(err)?,
        );
        let data: serde_json::Value =
            serde_json::from_str(&record.get::<_, String>(3).map_err(err)?).map_err(err)?;
        let edit: serde_json::Value =
            serde_json::from_str(&record.get::<_, String>(4).map_err(err)?).map_err(err)?;
        let content = edit.get("value").is_some();
        if !edit.is_object()
            || edit.as_object().unwrap().keys().any(|key| {
                key != "value"
                    && key != "styleId"
                    && !crate::range::STYLE_FIELDS.contains(&key.as_str())
            })
            || data.get("spill").is_some()
        {
            return Err(crate::coded(
                "style-unsupported",
                "Preserved-calculation export admits only ordinary literal or style edits",
            ));
        }
        if content
            && (data["value"] != edit["value"]
                || data.get("formula").is_some()
                || data.get("formulaType").is_some()
                || data.get("calculation").is_some())
        {
            return Err(crate::coded(
                "formula-unsupported",
                "Preserved-calculation export admits only ordinary literal edits",
            ));
        }
        if !content && content_fields(&data) != source_content(db, &sheet, r, c)? {
            return Err(crate::coded(
                "formula-stale",
                "Preserved formatting changed source value, formula or cache fields",
            ));
        }
        region(root, db, &sheet, (r, c), (r, c), content)?;
    }
    Ok(())
}
