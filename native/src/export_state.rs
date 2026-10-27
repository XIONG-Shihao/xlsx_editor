//! Prepare export from the accepted native version. Scratch tables are disposable
//! and version-fenced; no cell collection crosses the JavaScript boundary.
use crate::range::STYLE_FIELDS;
use crate::{Result, database, err};
use rusqlite::params;
use serde_json::{Value, json};
use std::path::Path;

pub fn prepare(root: &Path, input: &Value) -> Result<Value> {
    let mut db = database(root)?;
    let version = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing export version"))?;
    let tx = db.transaction().map_err(err)?;
    if tx
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
            "Unknown workbook version",
        ));
    }
    let has_spills = crate::spill::any_history(&tx, version)?;
    let spills = if has_spills {
        crate::spill::validate_current(root, &tx, input)?;
        crate::spill::export_groups(&tx, version, &input["view"])?
    } else {
        Vec::new()
    };
    let mut scalar_context = crate::spill::scalars::Context::new(root, &tx, version)?;
    let consumer_context = if has_spills {
        Some(crate::spill::consumers::Context::new(
            root,
            &tx,
            version,
            &input["view"],
        )?)
    } else {
        None
    };
    tx.execute_batch("DROP TABLE IF EXISTS export_changes; DROP TABLE IF EXISTS export_styles; DROP TABLE IF EXISTS export_state; CREATE TABLE export_changes(sheet TEXT,r INTEGER,c INTEGER,data TEXT,style INTEGER,PRIMARY KEY(sheet,r,c)) WITHOUT ROWID; CREATE TABLE export_styles(id INTEGER PRIMARY KEY,signature TEXT UNIQUE NOT NULL); CREATE TABLE export_state(version INTEGER,spill_view TEXT,incomplete INTEGER NOT NULL); CREATE TEMP TABLE effective(sheet TEXT,r INTEGER,c INTEGER,data TEXT,edit TEXT,PRIMARY KEY(sheet,r,c)) WITHOUT ROWID;").map_err(err)?;
    tx.execute("WITH RECURSIVE ancestry(id,depth) AS (SELECT ?1,0 UNION ALL SELECT v.parent,a.depth+1 FROM versions v JOIN ancestry a ON v.id=a.id WHERE v.parent IS NOT NULL), ranked AS (SELECT p.*,row_number() OVER (PARTITION BY p.sheet,p.r,p.c ORDER BY a.depth) priority FROM changes p JOIN ancestry a ON p.version=a.id) INSERT INTO effective SELECT sheet,r,c,data,edit FROM ranked WHERE priority=1",[version]).map_err(err)?;
    crate::shared::prepare(&tx, version)?;
    crate::shared::export_members(&tx)?;
    // The first spill slice does not certify every remaining array consumer.
    // Excel must refresh those preserved source caches when it opens the file.
    let (special, mut incomplete) = (false, has_spills);
    // Cells whose saved content differs from the source, and the formulas left
    // uncalculated; together they decide which stored results are still true.
    let mut changed_cells = std::collections::HashSet::new();
    let mut uncalculated = Vec::new();
    {
        let mut statement = tx.prepare_cached("SELECT e.r,e.c,c.kind,c.raw,c.inline,c.formula,c.formula_attrs,c.style,s.value,e.data,e.sheet,e.edit,coalesce(c.style,0) FROM effective e LEFT JOIN cells c ON c.sheet=e.sheet AND c.r=e.r AND c.c=e.c LEFT JOIN strings s ON c.kind='s' AND s.id=CAST(c.raw AS INTEGER)").map_err(err)?;
        let mut rows = statement.query([]).map_err(err)?;
        while let Some(row) = rows.next().map_err(err)? {
            crate::cancel::check()?;
            let sheet: String = row.get(10).map_err(err)?;
            let r: u32 = row.get(0).map_err(err)?;
            let c: u32 = row.get(1).map_err(err)?;
            let cell = crate::range::decode_cell(row)?.1;
            let cell = crate::shared::resolve(&tx, &sheet, &crate::xml::address(r, c), cell)?;
            let edit: Value =
                serde_json::from_str(&row.get::<_, String>(11).map_err(err)?).map_err(err)?;
            let mut change = json!({});
            // Null transport fields are explicit clears, not identities. Keep
            // the original special-source comparison for a null type marker;
            // validated spill-release tombstones deliberately omit these keys.
            let source_special_null = if cell.get("formulaType").is_some_and(Value::is_null) {
                row.get::<_, Option<String>>(6)
                    .map_err(err)?
                    .map(|raw| serde_json::from_str::<Value>(&raw).map_err(err))
                    .transpose()?
                    .is_some_and(|attrs| attrs.get("t").is_some())
            } else {
                false
            };
            let consumer_formula = if let Some(context) = &consumer_context {
                context.validate(&sheet, &crate::xml::address(r, c), &cell)?
            } else {
                if cell["calculation"].get("spillConsumer").is_some() {
                    return Err(crate::coded(
                        "spill-bad-input",
                        "Scalar spill consumer has no accepted ownership history",
                    ));
                }
                None
            };
            let scalar_formula =
                scalar_context.validate(&sheet, &crate::xml::address(r, c), &cell)?;
            if crate::calculation_policy::preserving(&tx, version)?
                && edit.get("value").is_none()
                && edit.get("formula").is_none()
            {
                // Style-only changes retain the complete stored value/formula/cache XML.
                if crate::preservation::content_fields(&cell)
                    != crate::preservation::source_content(&tx, &sheet, r, c)?
                {
                    return Err(crate::coded(
                        "workbook-stale",
                        "Preserved format changed source content",
                    ));
                }
            } else if has_spills
                && crate::spill::accepted_cell(&tx, version, &sheet, (r, c), &cell)?
            {
                change["spill"] = cell["spill"].clone();
                if cell["spill"]["role"] == "master" {
                    change["formula"] = cell["formula"].clone();
                    change["calculation"] = cell["calculation"].clone();
                } else {
                    change["value"] = cell["value"].clone();
                }
            } else if cell.get("formulaType").is_some_and(|v| !v.is_null())
                || source_special_null
                || cell["scenarioResult"] == true
            {
                let mut statement=tx.prepare_cached("SELECT c.r,c.c,c.kind,c.raw,c.inline,c.formula,c.formula_attrs,c.style,s.value,NULL FROM cells c LEFT JOIN strings s ON c.kind='s' AND s.id=CAST(c.raw AS INTEGER) WHERE c.sheet=?1 AND c.r=?2 AND c.c=?3").map_err(err)?;
                let mut records = statement.query(params![sheet, r, c]).map_err(err)?;
                let mut original = match records.next().map_err(err)? {
                    Some(row) => crate::range::decode_cell(row)?.1,
                    None => Value::Null,
                };
                let mut current = cell.clone();
                if let Some(map) = original.as_object_mut() {
                    map.remove("calculation");
                    map.remove("scenarioResult");
                }
                if let Some(map) = current.as_object_mut() {
                    map.remove("calculation");
                    map.remove("scenarioResult");
                }
                // JSON transport may spell an integral f64 as an integer.
                // rawValue and all formula metadata still require exact equality.
                for value in [&mut original, &mut current] {
                    for key in ["value", "cachedValue"] {
                        if let Some(n) = value[key].as_f64() {
                            value[key] = json!(n);
                        }
                    }
                }
                if cell["scenarioResult"] == true {
                    current["value"] = original["value"].clone();
                }
                // Styles are independently validated; compare authored content only.
                for value in [&mut original, &mut current] {
                    if let Some(map) = value.as_object_mut() {
                        for key in STYLE_FIELDS {
                            map.remove(*key);
                        }
                        map.remove("format");
                        map.remove("styleId");
                    }
                }
                if original != current {
                    return Err(crate::coded(
                        "formula-unsupported",
                        "Changing special formula content is not qualified",
                    ));
                }
                // Calculation only marked this source cell unresolved. Preserve
                // its entire original XML and cached value, never publish it as fresh.
                if cell["calculation"]["status"] == "calculated" {
                    if cell["calculation"]["value"] != cell["cachedValue"] {
                        change["cache"] = cell["calculation"].clone();
                    }
                } else {
                    incomplete = true;
                }
            } else if let Some(formula) = cell.get("formula").filter(|v| !v.is_null()) {
                let unresolved_source = matches!(
                    cell["calculation"]["status"].as_str(),
                    Some("unresolved" | "circular")
                ) && edit.get("formula").is_none()
                    && edit.get("value").is_none()
                    && crate::range::source_formula_identity(row, &cell)?;
                if unresolved_source {
                    // Keep the exact imported XML/cache, including typed text,
                    // errors, absent values and signed zero. This is stored source
                    // evidence only, never a current calculation result.
                    incomplete = true;
                    if let Some(text) = formula.as_str() {
                        uncalculated.push((sheet.clone(), r, c, text.to_owned()));
                    }
                } else if !(crate::calculation_policy::pending(&tx, version, &sheet, (r, c))?
                    && cell["calculation"].is_null()
                    && edit.get("formula").is_none())
                {
                    change["formula"] = consumer_formula
                        .or(scalar_formula)
                        .map(Value::String)
                        .unwrap_or_else(|| formula.clone());
                    change["calculation"] = cell["calculation"].clone();
                    incomplete |= !matches!(
                        cell["calculation"]["status"].as_str(),
                        Some("calculated" | "error")
                    );
                }
                // A style-only Manual edit retains the source formula/cache
                // XML verbatim. Re-entering/replacing a formula is explicit
                // content intent and cannot borrow that previous cache.
            } else if edit.get("value").is_some() {
                change["value"] = cell["value"].clone();
                if let Some(kind) = edit.get("valueType") {
                    if kind != "error"
                        || cell["dataType"] != "error"
                        || cell["value"] != edit["value"]
                        || !cell["value"]
                            .as_str()
                            .is_some_and(crate::bulk::values::error_code)
                    {
                        return Err(crate::coded(
                            "internal-mismatch",
                            "Declared error export differs from its typed accepted cell",
                        ));
                    }
                    change["valueType"] = json!("error");
                }
            }
            let recalculated = cell["calculation"]["status"] == "calculated"
                && cell["calculation"]["value"] != cell["cachedValue"];
            if edit.get("value").is_some()
                || edit.get("formula").is_some()
                || change.get("cache").is_some()
                || change.get("spill").is_some()
                || change.get("formula").is_some()
                    && (recalculated || cell["calculation"]["status"] == "error")
            {
                changed_cells.insert((sheet.clone(), r, c));
            }
            if edit["$spillRelease"] == true && cell.get("spill").is_none() {
                change["$spillRelease"] = json!(true);
            }
            let mut style = None;
            if edit.get("styleId").is_some()
                || STYLE_FIELDS.iter().any(|key| edit.get(*key).is_some())
            {
                let mut signature = json!({"styleId":edit["styleId"].as_u64().unwrap_or(row.get::<_,u64>(12).map_err(err)?)});
                for key in STYLE_FIELDS {
                    if let Some(value) = edit.get(*key) {
                        signature[*key] = value.clone();
                    }
                }
                let signature = serde_json::to_string(&signature).map_err(err)?;
                tx.prepare_cached("INSERT OR IGNORE INTO export_styles(signature) VALUES(?1)")
                    .map_err(err)?
                    .execute([&signature])
                    .map_err(err)?;
                style = Some(
                    tx.prepare_cached("SELECT id FROM export_styles WHERE signature=?1")
                        .map_err(err)?
                        .query_row([&signature], |r| r.get::<_, u64>(0))
                        .map_err(err)?,
                );
            }
            tx.prepare_cached("INSERT INTO export_changes VALUES(?1,?2,?3,?4,?5)")
                .map_err(err)?
                .execute(params![
                    sheet,
                    r,
                    c,
                    serde_json::to_string(&change).map_err(err)?,
                    style
                ])
                .map_err(err)?;
        }
    }
    {
        // An uncalculated formula keeps its imported result only while nothing it
        // reads has changed. Otherwise the result is dropped, so a program that
        // does not recalculate shows an empty cell instead of a wrong number.
        let sheets: std::collections::HashMap<String, String> = input["view"]["sheets"]
            .as_array()
            .map(|sheets| {
                sheets
                    .iter()
                    .filter_map(|sheet| {
                        Some((
                            sheet["name"].as_str()?.to_lowercase(),
                            sheet["path"].as_str()?.to_owned(),
                        ))
                    })
                    .collect()
            })
            .unwrap_or_default();
        let mut drop_result = tx
            .prepare_cached("UPDATE export_changes SET data=json_set(data,'$.cache',json('{\"status\":\"unresolved\"}')) WHERE sheet=?1 AND r=?2 AND c=?3")
            .map_err(err)?;
        for (sheet, r, c) in crate::stale_cache::stale(&changed_cells, &uncalculated, &sheets) {
            drop_result.execute(params![sheet, r, c]).map_err(err)?;
        }
    }
    tx.execute(
        "INSERT INTO export_state VALUES(?1,?2,?3)",
        params![
            version,
            if has_spills {
                Some(input["view"].to_string())
            } else {
                None
            },
            incomplete
        ],
    )
    .map_err(err)?;
    let styles: u64 = tx
        .query_row("SELECT count(*) FROM export_styles", [], |r| r.get(0))
        .map_err(err)?;
    tx.commit().map_err(err)?;
    let mut result = json!({"special":special,"incomplete":incomplete,"styleCount":styles});
    if has_spills {
        result["spills"] = json!(spills);
    }
    Ok(result)
}

pub fn styles(root: &Path, input: &Value) -> Result<Value> {
    let db = database(root)?;
    let version = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing export version"))?;
    if db
        .query_row("SELECT version FROM export_state", [], |r| {
            r.get::<_, u64>(0)
        })
        .map_err(err)?
        != version
    {
        return Err(crate::coded("workbook-stale", "Stale export preparation"));
    }
    let after = input["after"].as_u64().unwrap_or(0);
    let mut statement = db
        .prepare_cached("SELECT id,signature FROM export_styles WHERE id>?1 ORDER BY id LIMIT 1024")
        .map_err(err)?;
    let mut rows = statement.query([after]).map_err(err)?;
    let mut styles = Vec::new();
    while let Some(row) = rows.next().map_err(err)? {
        crate::cancel::check()?;
        let id: u64 = row.get(0).map_err(err)?;
        let format: Value =
            serde_json::from_str(&row.get::<_, String>(1).map_err(err)?).map_err(err)?;
        styles.push(json!({"id":id,"format":format}));
    }
    Ok(json!({"next":styles.last().map(|s|s["id"].clone()),"styles":styles}))
}

#[cfg(test)]
mod unresolved_source_tests {
    use super::*;
    use std::{
        fs,
        io::{Read, Write},
    };
    const SHEET: &str = "xl/worksheets/main.xml";
    fn fixture() -> tempfile::TempDir {
        let root = tempfile::tempdir().unwrap();
        let mut zip =
            zip::ZipWriter::new(fs::File::create(root.path().join("source.xlsx")).unwrap());
        for (name, text) in [
            (
                "[Content_Types].xml",
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Override PartName=\"/xl/worksheets/main.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>",
            ),
            (
                "xl/workbook.xml",
                "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Main\" sheetId=\"1\" r:id=\"main\"/></sheets><calcPr calcId=\"0\" calcMode=\"auto\" calcOnSave=\"1\" fullCalcOnLoad=\"0\" forceFullCalc=\"0\"/></workbook>",
            ),
            (
                "xl/_rels/workbook.xml.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"main\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/main.xml\"/></Relationships>",
            ),
            (
                SHEET,
                "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"A1:A8\"/><sheetData><row r=\"1\"><c r=\"A1\"><f>UNKNOWN_NUM()</f><v>0.125</v></c></row><row r=\"2\"><c r=\"A2\" t=\"str\"><f>UNKNOWN_TEXT()</f><v>stored 雪</v></c></row><row r=\"3\"><c r=\"A3\" t=\"b\"><f>UNKNOWN_BOOL()</f><v>1</v></c></row><row r=\"4\"><c r=\"A4\" t=\"e\"><f>UNKNOWN_ERROR()</f><v>#N/A</v></c></row><row r=\"5\"><c r=\"A5\" t=\"n\"><f>UNKNOWN_ZERO()</f><v>-0</v></c></row><row r=\"6\"><c r=\"A6\"><f>UNKNOWN_ABSENT()</f></c></row><row r=\"7\"><c r=\"A7\"><f>A7+1</f><v>77</v></c></row><row r=\"8\"><c r=\"A8\"><f>1+1</f><v>99</v></c></row></sheetData></worksheet>",
            ),
            (
                "xl/styles.xml",
                "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts count=\"2\"><font/><font><b/></font></fonts><fills count=\"1\"><fill/></fills><borders count=\"1\"><border/></borders><cellXfs count=\"2\"><xf numFmtId=\"0\" fontId=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" applyFont=\"1\"/></cellXfs></styleSheet>",
            ),
        ] {
            zip.start_file(name, zip::write::SimpleFileOptions::default())
                .unwrap();
            zip.write_all(text.as_bytes()).unwrap();
        }
        zip.finish().unwrap();
        crate::import::index(root.path()).unwrap();
        root
    }
    fn view() -> Value {
        json!({"sheets":[{"name":"Main","path":SHEET,"storedFormulaCount":8,"merges":[]}],"namedRanges":[],"tables":[]})
    }
    fn cells(root: &Path, version: u64) -> Value {
        crate::range::read_db(
            &database(root).unwrap(),
            &json!({"version":version,"sheet":SHEET,"range":"A1:A8","limit":8}),
        )
        .unwrap()["cells"]
            .clone()
    }
    fn mark_unresolved(root: &Path) -> u64 {
        let changes=cells(root,0).as_array().unwrap().iter().take(7).map(|pair|{
            let mut cell=pair[1].clone();cell["calculation"]=json!({"status":if pair[0]=="A7"{"circular"}else{"unresolved"},"error":"Retained unavailable source result","revision":"unresolved-source"});
            // Exercise the production wire spelling of negative zero as well.
            if pair[0]=="A5"{cell["value"]=json!({"$workbookNumber":"-0"});cell["cachedValue"]=json!({"$workbookNumber":"-0"});}
            json!({"sheet":SHEET,"cell":pair[0],"value":cell,"edits":[]})
        }).collect::<Vec<_>>();
        crate::range::patch(root, &json!({"version":0,"changes":changes})).unwrap()["version"]
            .as_u64()
            .unwrap()
    }
    fn xml(root: &Path, file: &str) -> String {
        let mut zip = zip::ZipArchive::new(fs::File::open(root.join(file)).unwrap()).unwrap();
        let mut xml = String::new();
        zip.by_name(SHEET)
            .unwrap()
            .read_to_string(&mut xml)
            .unwrap();
        xml
    }
    fn element(xml: &str, reference: &str) -> String {
        let at = xml.find(&format!("r=\"{reference}\"")).unwrap();
        let start = xml[..at].rfind("<c").unwrap();
        let end = at + xml[at..].find("</c>").unwrap() + 4;
        xml[start..end].into()
    }
    fn export(root: &Path, version: u64) -> String {
        let plan = prepare(root, &json!({"version":version,"view":view()})).unwrap();
        assert_eq!(plan["incomplete"], true);
        let style_ids = styles(root, &json!({"version":version})).unwrap()["styles"]
            .as_array()
            .unwrap()
            .iter()
            .map(|style| json!({"id":style["id"],"styleId":1}))
            .collect::<Vec<_>>();
        let mut zip =
            zip::ZipArchive::new(fs::File::open(root.join("source.xlsx")).unwrap()).unwrap();
        let mut workbook = String::new();
        zip.by_name("xl/workbook.xml")
            .unwrap()
            .read_to_string(&mut workbook)
            .unwrap();
        workbook = workbook
            .replace("fullCalcOnLoad=\"0\"", "fullCalcOnLoad=\"1\"")
            .replace("forceFullCalc=\"0\"", "forceFullCalc=\"1\"");
        fs::write(
            root.join("export-metadata.json"),
            json!({"styleIds":style_ids,"replacements":{"xl/workbook.xml":workbook}}).to_string(),
        )
        .unwrap();
        let index: Value =
            serde_json::from_slice(&fs::read(root.join("index.json")).unwrap()).unwrap();
        crate::export::write(
            root,
            &json!({"version":version,"revision":index["revision"],"calculationMetadata":true}),
        )
        .unwrap();
        xml(root, "export.xlsx")
    }
    #[test]
    fn unresolved_and_circular_source_caches_keep_exact_typed_xml_with_style_edits() {
        let _scope = crate::shared::begin();
        let fixture = fixture();
        let root = fixture.path();
        let before = xml(root, "source.xlsx");
        let version = mark_unresolved(root);
        let after = export(root, version);
        for row in 1..=7 {
            assert_eq!(
                element(&after, &format!("A{row}")),
                element(&before, &format!("A{row}"))
            );
        }
        let mut cell = cells(root, version)[0][1].clone();
        cell["bold"] = json!(true);
        let styled=crate::range::patch(root,&json!({"version":version,"changes":[{"sheet":SHEET,"cell":"A1","value":cell,"edits":[{"bold":true}]}]})).unwrap()["version"].as_u64().unwrap();
        let after = export(root, styled);
        let a = element(&after, "A1");
        assert!(a.contains("s=\"1\""));
        assert!(a.contains("<f>UNKNOWN_NUM()</f><v>0.125</v>"));
        for row in 2..=7 {
            assert_eq!(
                element(&after, &format!("A{row}")),
                element(&before, &format!("A{row}"))
            );
        }
        let current = cells(root, styled);
        for pair in current.as_array().unwrap().iter().take(7) {
            assert!(matches!(
                pair[1]["calculation"]["status"].as_str(),
                Some("unresolved" | "circular")
            ));
        }
    }
    #[test]
    fn changed_and_reentered_formulas_cannot_borrow_an_imported_cache() {
        let _scope = crate::shared::begin();
        let fixture = fixture();
        let root = fixture.path();
        let version = mark_unresolved(root);
        let source = cells(root, version);
        let mut changed = source[1][1].clone();
        changed["formula"] = json!("UNKNOWN_NEW()");
        let changed=crate::range::patch(root,&json!({"version":version,"changes":[{"sheet":SHEET,"cell":"A1","value":source[0][1],"edits":[{"formula":"UNKNOWN_NUM()"}]},{"sheet":SHEET,"cell":"A2","value":changed,"edits":[{"formula":"UNKNOWN_NEW()"}]}]})).unwrap()["version"].as_u64().unwrap();
        let saved = export(root, changed);
        assert!(!element(&saved, "A1").contains("<v"));
        assert!(!element(&saved, "A2").contains("<v"));
        assert!(element(&saved, "A2").contains("UNKNOWN_NEW()"));
        assert!(element(&saved, "A3").contains("<v>1</v>"));
    }
    #[test]
    fn typed_null_formula_clears_export_as_blanks_without_losing_untouched_typed_caches() {
        let _scope = crate::shared::begin();
        let fixture = fixture();
        let root = fixture.path();
        let before = xml(root, "source.xlsx");
        let version = mark_unresolved(root);
        let cleared=crate::range::patch(root,&json!({"version":version,"changes":[{"sheet":SHEET,"cell":"A8","value":{"value":null,"formula":null,"formulaType":null,"dataType":"blank","styleId":0},"edits":[{"value":null}]},{"sheet":SHEET,"cell":"Z80","value":{"value":null,"formula":null,"formulaType":null,"dataType":"blank","styleId":0},"edits":[{"value":null}]}]})).unwrap()["version"].as_u64().unwrap();
        let saved = export(root, cleared);
        assert!(!saved.contains("r=\"A8\""));
        assert!(!saved.contains("r=\"Z80\""));
        for row in 1..=7 {
            assert_eq!(
                element(&saved, &format!("A{row}")),
                element(&before, &format!("A{row}"))
            );
        }
        let accepted = cells(root, cleared);
        assert!(accepted[7][1]["value"].is_null());
        assert!(accepted[7][1]["formula"].is_null());
        assert_eq!(
            xml(root, "source.xlsx"),
            before,
            "export cannot mutate the immutable source"
        );
    }
    #[test]
    fn null_special_formula_marker_cannot_bypass_opaque_source_identity_guard() {
        let _scope = crate::shared::begin();
        let original = fixture();
        let fixture = tempfile::tempdir().unwrap();
        let root = fixture.path();
        let mut source =
            zip::ZipArchive::new(fs::File::open(original.path().join("source.xlsx")).unwrap())
                .unwrap();
        let mut output = zip::ZipWriter::new(fs::File::create(root.join("source.xlsx")).unwrap());
        for i in 0..source.len() {
            let mut member = source.by_index(i).unwrap();
            if member.name() == SHEET {
                let mut worksheet = String::new();
                member.read_to_string(&mut worksheet).unwrap();
                worksheet = worksheet.replace("<f>1+1</f>", "<f t=\"array\" ref=\"A8:B8\">1+1</f>");
                output
                    .start_file(SHEET, zip::write::SimpleFileOptions::default())
                    .unwrap();
                output.write_all(worksheet.as_bytes()).unwrap();
            } else {
                output.raw_copy_file(member).unwrap();
            }
        }
        output.finish().unwrap();
        crate::import::index(root).unwrap();
        let source_hash = crate::hash(&root.join("source.xlsx")).unwrap();
        let index_hash = crate::hash(&root.join("index.json")).unwrap();
        prepare(root, &json!({"version":0,"view":view()})).unwrap();
        let rejected=crate::range::patch(root,&json!({"version":0,"changes":[{"sheet":SHEET,"cell":"A8","value":{"value":null,"formula":null,"formulaType":null,"dataType":"blank","styleId":0},"edits":[]}]})).unwrap()["version"].as_u64().unwrap();
        let before = cells(root, rejected);
        let error = prepare(root, &json!({"version":rejected,"view":view()})).unwrap_err();
        assert!(error.contains("Changing special formula content is not qualified"));
        assert_eq!(cells(root, rejected), before);
        assert_eq!(
            database(root)
                .unwrap()
                .query_row("SELECT version FROM export_state", [], |row| row
                    .get::<_, u64>(0))
                .unwrap(),
            0,
            "failed preparation retains its previous export plan"
        );
        assert_eq!(crate::hash(&root.join("source.xlsx")).unwrap(), source_hash);
        assert_eq!(crate::hash(&root.join("index.json")).unwrap(), index_hash);
        let mut styled = cells(root, 0)[7][1].clone();
        styled["bold"] = json!(true);
        let styled=crate::range::patch(root,&json!({"version":0,"changes":[{"sheet":SHEET,"cell":"A8","value":styled,"edits":[{"bold":true}]}]})).unwrap()["version"].as_u64().unwrap();
        let plan = prepare(root, &json!({"version":styled,"view":view()})).unwrap();
        assert_eq!(plan["incomplete"], true);
        let raw: String = database(root)
            .unwrap()
            .query_row(
                "SELECT data FROM export_changes WHERE sheet=?1 AND r=7 AND c=0",
                [SHEET],
                |row| row.get(0),
            )
            .unwrap();
        assert_eq!(
            serde_json::from_str::<Value>(&raw).unwrap(),
            json!({}),
            "style-only special source preserves its formula and typed cache XML"
        );
    }
    #[test]
    fn native_calculation_does_not_manufacture_source_formula_author_intent() {
        let _scope = crate::shared::begin();
        let fixture = fixture();
        let root = fixture.path();
        let calculated = crate::calculation::command(
            root,
            &json!({"version":0,"view":view(),"revision":"native-current","partial":true}),
        )
        .unwrap();
        assert_eq!(calculated["supported"], true);
        let version = calculated["version"].as_u64().unwrap();
        let db = database(root).unwrap();
        let edit: String = db
            .query_row(
                "SELECT edit FROM changes WHERE version=?1 AND sheet=?2 AND r=7 AND c=0",
                params![version, SHEET],
                |row| row.get(0),
            )
            .unwrap();
        assert_eq!(serde_json::from_str::<Value>(&edit).unwrap(), json!({}));
        assert_eq!(
            crate::range::read_db(
                &db,
                &json!({"version":version,"sheet":SHEET,"range":"A8","limit":1})
            )
            .unwrap()["cells"][0][1]["calculation"]["value"],
            2.0
        );
    }
}
