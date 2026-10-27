//! Values reads an immutable accepted version. Stored outputs never become
//! literals; current complete spill/scenario proofs permit scalar snapshots.
use crate::{Result, err, field, local};
use rusqlite::{Connection, params};
use serde_json::{Value, json};
use std::{fs::File, io::Read, path::Path};
pub(crate) fn error_code(code: &str) -> bool {
    [
        "#DIV/0!", "#VALUE!", "#NUM!", "#REF!", "#N/A", "#NAME?", "#NULL!", "#SPILL!", "#CALC!",
    ]
    .contains(&code)
}
pub(super) struct Snapshot {
    pub value: Value,
    pub error: bool,
}
pub(super) fn literal_snapshot(cell: &Value) -> Result<Snapshot> {
    if [
        "formula",
        "formulaType",
        "formulaOutput",
        "spill",
        "calculation",
    ]
    .iter()
    .any(|key| cell.get(*key).is_some_and(|value| !value.is_null()))
        || cell["scenarioResult"] == true
    {
        return Err(crate::coded(
            "fill-unsupported",
            "Source formula output copy is not qualified",
        ));
    }
    let value = super::literal(cell, false)?;
    let error = cell["dataType"] == "error";
    if error && !value.as_str().is_some_and(error_code) {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid declared error value",
        ));
    }
    Ok(Snapshot { value, error })
}
pub(super) fn snapshot(cell: &Value) -> Result<Snapshot> {
    if cell["calculation"]["pending"] == true {
        return Err(crate::coded(
            "fill-bad-input",
            "Paste Values needs a current calculated result; source is pending",
        ));
    }
    let derived = ["formula", "formulaType", "formulaOutput", "spill"]
        .iter()
        .any(|key| cell.get(*key).is_some())
        || cell["scenarioResult"] == true;
    let (value, error) = if derived {
        match cell["calculation"]["status"].as_str() {
            Some("calculated") if cell["calculation"].get("value").is_some() => {
                (cell["calculation"]["value"].clone(), false)
            }
            Some("error")
                if cell["calculation"]["error"]
                    .as_str()
                    .is_some_and(error_code) =>
            {
                (cell["calculation"]["error"].clone(), true)
            }
            _ => {
                return Err(crate::coded(
                    "fill-bad-input",
                    "Paste Values needs a current calculated result",
                ));
            }
        }
    } else {
        (cell["value"].clone(), cell["dataType"] == "error")
    };
    if error && !value.as_str().is_some_and(error_code) {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid typed Values error",
        ));
    }
    if !value.is_null() && !value.is_string() && !value.is_boolean() && !value.is_number() {
        return Err(crate::coded(
            "fill-bad-input",
            "Values source is not a scalar",
        ));
    }
    if let Some(number) = value.as_f64() {
        if !number.is_finite()
            || !derived
                && cell["rawValue"]
                    .as_str()
                    .is_some_and(|raw| super::meaning(raw) != super::meaning(&number.to_string()))
        {
            return Err(crate::coded(
                "fill-refused",
                "Paste Values would lose stored numeric precision",
            ));
        }
    }
    Ok(Snapshot { value, error })
}
fn metadata(root: &Path) -> Result<Value> {
    let mut view = crate::package_edit::auditing_metadata(root)?;
    let mut zip = zip::ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let mut member = zip.by_name("xl/workbook.xml").map_err(err)?;
    if member.size() > 16 * 1024 * 1024 {
        return Err(crate::coded(
            "fill-limit",
            "Values calendar metadata exceeds its bound",
        ));
    }
    let mut bytes = Vec::new();
    member.read_to_end(&mut bytes).map_err(err)?;
    let mut reader = quick_xml::Reader::from_reader(bytes.as_slice());
    let mut calendar = false;
    let mut count = 0;
    loop {
        match reader.read_event().map_err(err)? {
            quick_xml::events::Event::Start(e) | quick_xml::events::Event::Empty(e)
                if e.local_name().as_ref() == b"workbookPr" =>
            {
                count += 1;
                if count > 1 {
                    return Err(crate::coded("fill-ambiguous", "Ambiguous Values calendar"));
                }
                calendar = match crate::xml::attrs(&e)?.get("date1904").map(String::as_str) {
                    None | Some("0" | "false") => false,
                    Some("1" | "true") => true,
                    _ => {
                        return Err(crate::coded(
                            "internal-invalid-field",
                            "Invalid Values calendar",
                        ));
                    }
                };
            }
            quick_xml::events::Event::Eof => break,
            _ => (),
        }
    }
    for sheet in view["sheets"]
        .as_array_mut()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing Values sheets"))?
    {
        sheet["date1904"] = json!(calendar);
    }
    Ok(view)
}
pub(super) struct Context {
    view: Option<Value>,
    prepared: bool,
    spill_validated: bool,
}
impl Context {
    pub fn new() -> Self {
        Self {
            view: None,
            prepared: false,
            spill_validated: false,
        }
    }
    fn prepare(&mut self, root: &Path, db: &Connection, version: u64, spill: bool) -> Result<()> {
        if self.view.is_none() {
            self.view = Some(metadata(root)?);
        }
        if !self.prepared || spill && !self.spill_validated {
            // Every command pins a distinct parent; TEMP work is not accepted
            // history and cannot borrow another command's evaluator memo.
            db.execute_batch("DROP TABLE IF EXISTS temp.calc_overlay;DROP TABLE IF EXISTS temp.calc_shared;DROP TABLE IF EXISTS temp.calc_targets;DROP TABLE IF EXISTS temp.calc_values;DROP TABLE IF EXISTS temp.calc_traces;DROP TABLE IF EXISTS temp.calc_scenarios;DROP TABLE IF EXISTS temp.spill_protected;").map_err(err)?;
            if spill {
                crate::spill::validate_current(
                    root,
                    db,
                    &json!({"version":version,"view":self.view.as_ref().unwrap()}),
                )?;
                self.spill_validated = true;
            } else {
                crate::calculation::snapshot(db, version)?;
                crate::scenarios::prepare(db, self.view.as_ref().unwrap())?;
            }
            self.prepared = true;
        }
        Ok(())
    }
    pub fn read(
        &mut self,
        root: &Path,
        db: &Connection,
        version: u64,
        sheet: &str,
        at: (u32, u32),
        cell: &Value,
    ) -> Result<Snapshot> {
        let result = snapshot(cell)?;
        if cell.get("spill").is_some() {
            self.prepare(root, db, version, true)?;
            if !crate::spill::accepted_cell(db, version, sheet, at, cell)? {
                return Err(crate::coded(
                    "fill-bad-input",
                    "Paste Values spill source has no current complete owner",
                ));
            }
        } else if cell.get("formulaOutput").is_some()
            || cell["scenarioResult"] == true
            || cell["formulaType"] == "dataTable"
        {
            self.prepare(root, db, version, false)?;
            let known: bool = db
                .prepare_cached(
                    "SELECT EXISTS(SELECT 1 FROM calc_scenarios WHERE sheet=?1 AND r=?2 AND c=?3)",
                )
                .map_err(crate::err)?
                .query_row(params![sheet, at.0, at.1], |row| row.get(0))
                .map_err(err)?;
            if !known {
                return Err(crate::coded(
                    "fill-bad-input",
                    "Paste Values source output has no current scalar scenario proof",
                ));
            }
            let view = self.view.as_ref().unwrap();
            let name = field(
                view["sheets"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .find(|s| s["path"] == sheet)
                    .ok_or_else(|| {
                        crate::coded("internal-unknown-reference", "Unknown Values sheet")
                    })?,
                "name",
            )?;
            let mut engine = crate::formula::Eval::from_sqlite(view, db, false);
            let actual = engine.result(name, &crate::xml::address(at.0, at.1))?;
            if result.error
                || actual
                    .as_f64()
                    .zip(result.value.as_f64())
                    .is_none_or(|(a, b)| a.to_bits() != b.to_bits())
            {
                return Err(crate::coded(
                    "internal-mismatch",
                    "Paste Values scenario result differs from current native inputs",
                ));
            }
        } else if cell.get("formulaType").is_some() {
            return Err(crate::coded(
                "fill-unsupported",
                "Special formula Values source has no current scalar proof",
            ));
        }
        Ok(result)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;
    use zip::{ZipWriter, write::SimpleFileOptions};
    const SHEET: &str = "xl/worksheets/sheet1.xml";
    fn fixture(body: &str) -> tempfile::TempDir {
        let directory = tempfile::tempdir().unwrap();
        let mut zip = ZipWriter::new(File::create(directory.path().join("source.xlsx")).unwrap());
        let ns = String::from_utf8_lossy(crate::xml::NS);
        for(name,text)in [(SHEET,format!("<worksheet xmlns=\"{ns}\"><sheetData>{body}</sheetData></worksheet>")),("xl/workbook.xml",format!("<workbook xmlns=\"{ns}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><workbookPr date1904=\"0\"/><sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"s1\"/></sheets><calcPr calcMode=\"auto\"/></workbook>")),("xl/_rels/workbook.xml.rels","<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"s1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>".into()),("[Content_Types].xml",format!("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Override PartName=\"/{SHEET}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>"))]{zip.start_file(name,SimpleFileOptions::default()).unwrap();zip.write_all(text.as_bytes()).unwrap();}
        zip.finish().unwrap();
        crate::import::index(directory.path()).unwrap();
        directory
    }
    fn cmd(source: &str, target: &str) -> Value {
        json!({"type":"pasteValues","sheet":SHEET,"source":source,"range":target})
    }
    fn read(root: &Path, version: u64, range: &str) -> Value {
        crate::range::read_db(
            &crate::database(root).unwrap(),
            &json!({"version":version,"sheet":SHEET,"range":range}),
        )
        .unwrap()["cells"]
            .clone()
    }
    fn pins(root: &Path) -> Vec<Vec<u8>> {
        ["source.xlsx", "metadata.xlsx", "cells.sqlite", "index.json"]
            .iter()
            .map(|file| std::fs::read(root.join(file)).unwrap())
            .collect()
    }
    #[test]
    fn typed_snapshot_errors_text_pending_and_precision_are_distinct() {
        assert!(
            snapshot(&json!({"value":"#N/A","dataType":"error"}))
                .unwrap()
                .error
        );
        assert!(
            !snapshot(&json!({"value":"#N/A","dataType":"text"}))
                .unwrap()
                .error
        );
        let error =
            snapshot(&json!({"formula":"1/0","calculation":{"status":"error","error":"#DIV/0!"}}))
                .unwrap();
        assert!(error.error);
        assert_eq!(error.value, "#DIV/0!");
        for status in ["stored", "unresolved", "circular"] {
            assert!(snapshot(&json!({"formulaOutput":{"master":"A1"},"value":913,"calculation":{"status":status,"value":913,"error":"#N/A"}})).is_err());
        }
        assert!(snapshot(&json!({"formula":"1","calculation":{"status":"calculated","value":913,"pending":true}})).is_err());
        assert!(
            snapshot(&json!({"value":9007199254740992.0,"rawValue":"9007199254740993"})).is_err()
        );
    }
    #[test]
    fn actual_native_typed_error_values_and_general_edit_export_preserve_literal_type() {
        let directory = fixture(
            "<row r=\"1\"><c r=\"A1\" t=\"e\"><v>#N/A</v></c><c r=\"B1\" t=\"inlineStr\"><is><t>#N/A</t></is></c><c r=\"D1\" s=\"0\"><v>8</v></c></row>",
        );
        let root = directory.path();
        let original = std::fs::read(root.join("source.xlsx")).unwrap();
        let result = super::super::apply(
            root,
            &json!({"version":0,"commands":[cmd("A1:B1","D1:E1")]}),
        )
        .unwrap();
        let version = result["version"].as_u64().unwrap();
        let cells = read(root, version, "D1:E1");
        assert_eq!(cells[0][1]["dataType"], "error");
        assert_eq!(cells[1][1]["dataType"], "text");
        assert_eq!(cells[0][1]["styleId"], 0);
        assert_eq!(read(root, 0, "A1")[0][1]["dataType"], "error");
        let index: Value =
            serde_json::from_slice(&std::fs::read(root.join("index.json")).unwrap()).unwrap();
        crate::export::write(root,&json!({"revision":index["revision"],"changes":[{"sheet":SHEET,"cell":"D1","value":"#N/A","valueType":"error"},{"sheet":SHEET,"cell":"E1","value":"#N/A"}]})).unwrap();
        let mut zip = zip::ZipArchive::new(File::open(root.join("export.xlsx")).unwrap()).unwrap();
        let mut xml = String::new();
        zip.by_name(SHEET)
            .unwrap()
            .read_to_string(&mut xml)
            .unwrap();
        assert!(xml.contains("r=\"D1\" s=\"0\" t=\"e\""), "{xml}");
        assert!(xml.contains("r=\"E1\" t=\"inlineStr\""), "{xml}");
        assert_eq!(std::fs::read(root.join("source.xlsx")).unwrap(), original);
        crate::export_state::prepare(root, &json!({"version":version})).unwrap();
        drop(zip);
        let mut metadata_zip =
            zip::ZipArchive::new(File::open(root.join("metadata.xlsx")).unwrap()).unwrap();
        let mut workbook = String::new();
        metadata_zip
            .by_name("xl/workbook.xml")
            .unwrap()
            .read_to_string(&mut workbook)
            .unwrap();
        assert!(workbook.contains("<calcPr calcMode=\"auto\"/>"));
        let replacement = workbook.replace(
            "<calcPr calcMode=\"auto\"/>",
            "<calcPr calcMode=\"auto\" calcOnSave=\"1\" fullCalcOnLoad=\"0\" forceFullCalc=\"0\"/>",
        );
        std::fs::write(
            root.join("export-metadata.json"),
            json!({"styleIds":[],"replacements":{"xl/workbook.xml":replacement}}).to_string(),
        )
        .unwrap();
        drop(metadata_zip);
        crate::export::write(
            root,
            &json!({"revision":index["revision"],"version":version,"calculationMetadata":true}),
        )
        .unwrap();
        let mut prepared =
            zip::ZipArchive::new(File::open(root.join("export.xlsx")).unwrap()).unwrap();
        let mut prepared_xml = String::new();
        prepared
            .by_name(SHEET)
            .unwrap()
            .read_to_string(&mut prepared_xml)
            .unwrap();
        assert!(
            prepared_xml.contains("r=\"D1\" s=\"0\" t=\"e\""),
            "Prepared actual export lost typed error: {prepared_xml}"
        );
        assert!(prepared_xml.contains("r=\"E1\" t=\"inlineStr\""));
        drop(prepared);
        let cold = tempfile::tempdir().unwrap();
        std::fs::copy(root.join("export.xlsx"), cold.path().join("source.xlsx")).unwrap();
        crate::import::index(cold.path()).unwrap();
        let cold_cell = read(cold.path(), 0, "D1");
        assert_eq!(cold_cell[0][1]["dataType"], "error");
        assert_eq!(cold_cell[0][1]["error"], "#N/A");
        assert_eq!(std::fs::read(root.join("source.xlsx")).unwrap(), original);
        let mut styled = read(root, version, "D1")[0][1].clone();
        styled["bold"] = json!(true);
        let styled_version=crate::range::patch(root,&json!({"version":version,"changes":[{"sheet":SHEET,"cell":"D1","value":styled,"edits":[{"bold":true}]}]})).unwrap()["version"].as_u64().unwrap();
        assert_eq!(read(root, styled_version, "D1")[0][1]["dataType"], "error");
        assert_eq!(read(root, styled_version, "D1")[0][1]["bold"], true);
        crate::export_state::prepare(root, &json!({"version":styled_version})).unwrap();
        let db = crate::database(root).unwrap();
        let raw: String = db
            .prepare_cached("SELECT data FROM export_changes WHERE sheet=?1 AND r=0 AND c=3")
            .unwrap()
            .query_row([SHEET], |row| row.get(0))
            .unwrap();
        assert_eq!(
            serde_json::from_str::<Value>(&raw).unwrap()["valueType"],
            "error"
        );
        drop(db);
        for (value, kind) in [(json!("#N/A"), "text"), (json!(12.0), "number")] {
            let mut cell = read(root, styled_version, "D1")[0][1].clone();
            cell["value"] = value.clone();
            cell["dataType"] = json!(kind);
            cell["rawValue"] = if value.is_string() {
                value.clone()
            } else {
                json!("12")
            };
            cell.as_object_mut().unwrap().remove("error");
            let changed=crate::range::patch(root,&json!({"version":styled_version,"changes":[{"sheet":SHEET,"cell":"D1","value":cell,"edits":[{"value":value}]}]})).unwrap()["version"].as_u64().unwrap();
            let db = crate::database(root).unwrap();
            let edit = crate::range::inherited_edit(&db, changed, SHEET, 0, 3).unwrap();
            assert!(edit.get("valueType").is_none());
            assert_eq!(edit["bold"], true);
            assert!(read(root, changed, "D1")[0][1].get("error").is_none());
            drop(db);
            crate::export_state::prepare(root, &json!({"version":changed})).unwrap();
            let raw: String = crate::database(root)
                .unwrap()
                .query_row(
                    "SELECT data FROM export_changes WHERE sheet=?1 AND r=0 AND c=3",
                    [SHEET],
                    |row| row.get(0),
                )
                .unwrap();
            assert!(
                serde_json::from_str::<Value>(&raw)
                    .unwrap()
                    .get("valueType")
                    .is_none()
            );
        }
        let before = pins(root);
        for bad in [
            json!({"value":"word","valueType":"error"}),
            json!({"value":"#N/A","valueType":"text"}),
        ] {
            let mut input = bad;
            input["sheet"] = json!(SHEET);
            input["cell"] = json!("F1");
            let error=crate::range::patch(root,&json!({"version":version,"changes":[{"sheet":SHEET,"cell":"F1","value":{"value":"word","dataType":"text"},"edits":[input]}]})).unwrap_err();
            assert!(error.contains("declared error"));
            assert_eq!(pins(root), before);
        }
    }
    #[test]
    fn actual_native_unknown_follower_and_late_failure_ignore_forged_current_projection() {
        let directory = fixture(
            "<row r=\"1\"><c r=\"A1\"><v>5</v></c><c r=\"C1\"><f t=\"array\" ref=\"C1:C2\">UNKNOWN(1)</f><v>913</v></c></row><row r=\"2\"><c r=\"C2\"><v>914</v></c></row>",
        );
        let root = directory.path();
        let view = metadata(root).unwrap();
        crate::calculation_policy::command(root,&json!({"version":0,"method":"initialize","view":view,"initialPolicy":{"mode":"auto","recalculateBeforeSave":false}})).unwrap();
        let before = pins(root);
        for commands in [
            vec![cmd("C2", "G1")],
            vec![cmd("A1", "F1"), cmd("C2", "G1")],
        ] {
            let error=super::super::apply(root,&json!({"version":0,"commands":commands,"view":{"sheets":[{"path":SHEET,"cells":{"C2":{"value":914,"calculation":{"status":"calculated","value":914}}}}]}})).unwrap_err();
            assert!(error.contains("current calculated"), "{error}");
            assert_eq!(pins(root), before);
        }
        assert!(
            super::super::apply(root, &json!({"version":0,"commands":[cmd("A1","G1")]})).is_ok()
        );
    }
    #[test]
    fn actual_native_literal_fill_preserves_error_text_style_intent_and_late_protected_refusal() {
        let directory = fixture(
            "<row r=\"1\"><c r=\"A1\" t=\"e\"><v>#N/A</v></c><c r=\"C1\"><f t=\"array\" ref=\"C1:C2\">UNKNOWN(1)</f><v>913</v></c></row><row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>#N/A</t></is></c><c r=\"C2\"><v>914</v></c></row>",
        );
        let root = directory.path();
        let view = metadata(root).unwrap();
        crate::calculation_policy::command(root,&json!({"version":0,"method":"initialize","view":view,"initialPolicy":{"mode":"auto","recalculateBeforeSave":false}})).unwrap();
        let fill = |source: &str, target: &str| json!({"type":"fillRange","sheet":SHEET,"source":source,"range":target});
        let before = pins(root);
        let error=super::super::apply(root,&json!({"version":0,"commands":[fill("A1","E1"),fill("C2","E2")],"view":{"sheets":[{"path":SHEET,"cells":{"C2":{"value":914,"calculation":{"status":"calculated","value":914}}}}]}})).unwrap_err();
        assert!(error.contains("Source formula output"), "{error}");
        assert_eq!(pins(root), before);
        let version = super::super::apply(
            root,
            &json!({"version":0,"commands":[fill("A1:A2","E1:E2"),fill("A1","F1:H1")]}),
        )
        .unwrap()["version"]
            .as_u64()
            .unwrap();
        let cells = read(root, version, "E1:H2");
        for reference in ["E1", "F1", "G1", "H1"] {
            let cell = cells
                .as_array()
                .unwrap()
                .iter()
                .find(|item| item[0] == reference)
                .unwrap();
            assert_eq!(cell[1]["dataType"], "error");
            assert_eq!(cell[1]["error"], "#N/A");
            assert_eq!(cell[1]["styleId"], 0);
        }
        let text = cells
            .as_array()
            .unwrap()
            .iter()
            .find(|item| item[0] == "E2")
            .unwrap();
        assert_eq!(text[1]["dataType"], "text");
        assert!(text[1].get("error").is_none());
        assert_eq!(read(root, 0, "A1")[0][1]["dataType"], "error");
        assert_eq!(std::fs::read(root.join("source.xlsx")).unwrap(), before[0]);
        let db = crate::database(root).unwrap();
        assert_eq!(
            crate::range::inherited_edit(&db, version, SHEET, 0, 4).unwrap()["valueType"],
            "error"
        );
        assert!(
            literal_snapshot(
                &json!({"value":914,"calculation":{"status":"calculated","value":914}})
            )
            .is_err()
        );
        assert!(literal_snapshot(&json!({"value":"#INVALID!","dataType":"error"})).is_err());
    }
    #[test]
    fn actual_native_current_spill_snapshot_requires_unchanged_constructor_and_target_remains_protected()
     {
        let directory = fixture(
            "<row r=\"1\"><c r=\"A1\"><f>_xlfn.SEQUENCE(J1,2)</f></c><c r=\"J1\"><v>2</v></c></row>",
        );
        let root = directory.path();
        let view = metadata(root).unwrap();
        let version = crate::spill::stage_qualified(
            root,
            &json!({"version":0,"view":view,"revision":"values-spill-unit"}),
        )
        .unwrap()["version"]
            .as_u64()
            .unwrap();
        assert_eq!(
            read(root, version, "B1")[0][1]["calculation"]["value"].as_f64(),
            Some(2.0)
        );
        let copied = super::super::apply(
            root,
            &json!({"version":version,"commands":[cmd("B1","Z1"),cmd("A1:B2","Z3:AA4")]}),
        )
        .unwrap()["version"]
            .as_u64()
            .unwrap();
        assert_eq!(read(root, copied, "Z1")[0][1]["value"].as_f64(), Some(2.0));
        assert!(read(root, copied, "Z1")[0][1].get("spill").is_none());
        let before = pins(root);
        assert!(
            super::super::apply(
                root,
                &json!({"version":version,"commands":[cmd("J1","B1")]})
            )
            .is_err()
        );
        assert_eq!(pins(root), before);
        let changed=crate::range::patch(root,&json!({"version":version,"changes":[{"sheet":SHEET,"cell":"J1","value":{"value":3.0,"dataType":"number"},"edits":[{"value":3.0}]}]})).unwrap()["version"].as_u64().unwrap();
        let before = pins(root);
        assert!(
            super::super::apply(
                root,
                &json!({"version":changed,"commands":[cmd("B1","Z1")]})
            )
            .is_err()
        );
        assert_eq!(pins(root), before);
    }
    #[test]
    fn actual_native_current_scenario_scalar_is_reproved_but_stored_and_stale_outputs_refuse() {
        let directory = fixture(
            "<row r=\"1\"><c r=\"A1\"><v>9</v></c></row><row r=\"2\"><c r=\"B2\"><f>A1*2</f><v>913</v></c></row><row r=\"3\"><c r=\"A3\"><v>3</v></c><c r=\"B3\"><f t=\"dataTable\" ref=\"B3:B4\" r1=\"A1\"/><v>913</v></c></row><row r=\"4\"><c r=\"A4\"><v>4</v></c><c r=\"B4\" t=\"e\"><v>#N/A</v></c></row>",
        );
        let root = directory.path();
        let view = metadata(root).unwrap();
        crate::calculation_policy::command(root,&json!({"version":0,"method":"initialize","view":view,"initialPolicy":{"mode":"auto","recalculateBeforeSave":false}})).unwrap();
        let before = pins(root);
        assert!(
            super::super::apply(root, &json!({"version":0,"commands":[cmd("B4","Z1")]})).is_err()
        );
        assert_eq!(pins(root), before);
        let result=crate::calculation::command(root,&json!({"version":0,"view":view,"metadata":"values-scenario-unit","revision":"actual-scenario"})).unwrap();
        assert_ne!(result["supported"], false, "{result}");
        let version = result["version"].as_u64().unwrap();
        let copied = super::super::apply(
            root,
            &json!({"version":version,"commands":[cmd("B4","Z1")]}),
        )
        .unwrap()["version"]
            .as_u64()
            .unwrap();
        assert_eq!(read(root, copied, "Z1")[0][1]["value"].as_f64(), Some(8.0));
        let changed=crate::range::patch(root,&json!({"version":version,"changes":[{"sheet":SHEET,"cell":"A4","value":{"value":5.0,"dataType":"number"},"edits":[{"value":5.0}]}]})).unwrap()["version"].as_u64().unwrap();
        let before = pins(root);
        let error = super::super::apply(
            root,
            &json!({"version":changed,"commands":[cmd("B4","Z1")]}),
        )
        .unwrap_err();
        assert!(
            error.contains("differs from current native inputs"),
            "{error}"
        );
        assert_eq!(pins(root), before);
    }
}
