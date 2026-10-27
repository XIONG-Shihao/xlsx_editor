use super::*;
use std::{fs, io::Write};
use zip::{ZipWriter, write::SimpleFileOptions};
const SOURCE: &str = "xl/worksheets/sheet1.xml";
const FIRST: &str = "xl/worksheets/sheet2.xml";
const SECOND: &str = "xl/worksheets/sheet3.xml";
const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
fn source_data() -> String {
    "<row r=\"1\"><c r=\"A1\" s=\"1\"><v>42</v></c><c r=\"B1\"><f>A1 + $A$1</f><v>84</v></c><c r=\"C1\" t=\"e\"><v>#N/A</v></c><c r=\"D1\" t=\"inlineStr\"><is><t>#N/A</t></is></c><c r=\"E1\" t=\"b\"><v>1</v></c><c r=\"F1\" t=\"inlineStr\"><is><t></t></is></c></row>".into()
}
fn target_data() -> String {
    "<row r=\"1\"><c r=\"A1\"><v>7</v></c><c r=\"B1\" s=\"1\"><f>A1*3</f><v>21</v></c><c r=\"C1\" s=\"1\"><v>3</v></c><c r=\"D1\" s=\"1\"><v>4</v></c><c r=\"E1\" s=\"1\"><v>5</v></c><c r=\"F1\" s=\"1\"><v>6</v></c><c r=\"G1\" s=\"1\"><f>A1*2</f><v>14</v></c></row>".into()
}
fn fixture(
    source: &str,
    target: &str,
    second: &str,
    extra: &[(&str, String)],
) -> tempfile::TempDir {
    let directory = tempfile::tempdir().unwrap();
    let mut zip = ZipWriter::new(fs::File::create(directory.path().join("source.xlsx")).unwrap());
    let styles = format!(
        "<styleSheet xmlns=\"{NS}\"><numFmts count=\"1\"><numFmt numFmtId=\"165\" formatCode=\"$0.0000;[Red]($0.0000)\"/></numFmts><fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border/></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"2\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"165\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/></cellXfs></styleSheet>"
    );
    let workbook = format!(
        "<workbook xmlns=\"{NS}\" xmlns:r=\"{REL}\"><sheets><sheet name=\"Template\" sheetId=\"1\" r:id=\"s1\"/><sheet name=\"January\" sheetId=\"2\" r:id=\"s2\"/><sheet name=\"February\" sheetId=\"3\" r:id=\"s3\"/></sheets><calcPr calcMode=\"auto\"/></workbook>"
    );
    let relationships = format!(
        "<Relationships xmlns=\"{PACKAGE_REL}\"><Relationship Id=\"s1\" Type=\"{REL}/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"s2\" Type=\"{REL}/worksheet\" Target=\"worksheets/sheet2.xml\"/><Relationship Id=\"s3\" Type=\"{REL}/worksheet\" Target=\"worksheets/sheet3.xml\"/><Relationship Id=\"styles\" Type=\"{REL}/styles\" Target=\"styles.xml\"/></Relationships>"
    );
    let mut parts = vec![
        (
            SOURCE,
            format!("<worksheet xmlns=\"{NS}\"><sheetData>{source}</sheetData></worksheet>"),
        ),
        (
            FIRST,
            format!("<worksheet xmlns=\"{NS}\"><sheetData>{target}</sheetData></worksheet>"),
        ),
        (
            SECOND,
            format!("<worksheet xmlns=\"{NS}\"><sheetData>{second}</sheetData></worksheet>"),
        ),
        ("xl/styles.xml", styles),
        ("xl/workbook.xml", workbook),
        ("xl/_rels/workbook.xml.rels", relationships),
        (
            "[Content_Types].xml",
            format!(
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Override PartName=\"/{SOURCE}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/{FIRST}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/{SECOND}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>"
            ),
        ),
    ];
    for (name, content) in extra {
        if let Some(part) = parts.iter_mut().find(|part| part.0 == *name) {
            part.1 = content.clone();
        } else {
            parts.push((name, content.clone()));
        }
    }
    for (name, content) in parts {
        zip.start_file(name, SimpleFileOptions::default()).unwrap();
        zip.write_all(content.as_bytes()).unwrap();
    }
    zip.finish().unwrap();
    crate::import::index(directory.path()).unwrap();
    directory
}
fn ordinary_fixture() -> tempfile::TempDir {
    fixture(&source_data(), &target_data(), &target_data(), &[])
}
fn command(sheet: &str, range: &str, mode: &str) -> Value {
    json!({"type":"fillAcrossSheets","sheet":sheet,"sourceSheet":SOURCE,"range":range,"mode":mode})
}
fn apply(root: &Path, version: u64, commands: Vec<Value>) -> crate::Result<Value> {
    super::super::apply(root, &json!({"version":version,"commands":commands}))
}
fn read(root: &Path, version: u64, sheet: &str, range: &str) -> Value {
    crate::range::read_public_db(
        &crate::database(root).unwrap(),
        &json!({"version":version,"sheet":sheet,"range":range,"limit":2048}),
    )
    .unwrap()["cells"]
        .clone()
}
fn cell(cells: &Value, name: &str) -> Value {
    cells
        .as_array()
        .unwrap()
        .iter()
        .find(|pair| pair[0] == name)
        .map(|pair| pair[1].clone())
        .unwrap_or(Value::Null)
}
fn pins(root: &Path) -> Vec<Vec<u8>> {
    ["source.xlsx", "metadata.xlsx", "cells.sqlite", "index.json"]
        .iter()
        .map(|name| fs::read(root.join(name)).unwrap())
        .collect()
}
fn edits(root: &Path, version: u64, sheet: &str, range: &str) -> Value {
    let at = crate::xml::position(range).unwrap();
    inherited_edit(&crate::database(root).unwrap(), version, sheet, at.0, at.1).unwrap()
}
fn decorated(root: &Path) -> u64 {
    crate::format::apply(root, &json!({"version":0,"commands":[{"sheet":SOURCE,"range":"A1","format":{"bold":true,"wrapText":true,"numberFormat":"decimal1"}},{"sheet":FIRST,"range":"G1","format":{"bold":true,"fillColor":"AA0000"}},{"sheet":SECOND,"range":"G1","format":{"bold":true,"fillColor":"AA0000"}}]})).unwrap()["version"].as_u64().unwrap()
}
fn assert_number(cell: &Value, expected: f64) {
    assert_eq!(cell["dataType"], "number");
    assert!(cell["value"].is_number());
    let number = cell["value"]
        .as_f64()
        .expect("Copied numeric value must remain a JSON number");
    assert!(number.is_finite());
    assert_eq!(cell["value"].as_f64(), Some(expected));
}
fn content(cell: &Value) -> Value {
    let mut value = cell.clone();
    if let Some(map) = value.as_object_mut() {
        for key in STYLE_FIELDS {
            map.remove(*key);
        }
        map.remove("styleId");
        map.remove("format");
    }
    value
}

#[test]
fn across_all_copies_exact_formulas_types_blank_resets_and_authored_styles_to_two_destinations() {
    let fixture = ordinary_fixture();
    let root = fixture.path();
    let version = decorated(root);
    let source = read(root, version, SOURCE, "A1:G1");
    let old = read(root, version, FIRST, "A1:G1");
    let bytes = fs::read(root.join("source.xlsx")).unwrap();
    let result = apply(
        root,
        version,
        vec![
            command(FIRST, "A1:G1", "all"),
            command(SECOND, "A1:G1", "all"),
        ],
    )
    .unwrap();
    let next = result["version"].as_u64().unwrap();
    assert_eq!(result["contentChanged"], true);
    assert_eq!(result["formulaTouched"], true);
    assert_eq!(result["formulaDelta"][FIRST], -1);
    assert_eq!(result["formulaDelta"][SECOND], -1);
    for sheet in [FIRST, SECOND] {
        let cells = read(root, next, sheet, "A1:G1");
        assert_number(&cell(&cells, "A1"), 42.0);
        assert_eq!(cell(&cells, "A1")["styleId"], 1);
        assert_eq!(cell(&cells, "A1")["bold"], true);
        assert_eq!(cell(&cells, "A1")["wrapText"], true);
        assert_eq!(cell(&cells, "A1")["format"], "0.0");
        let edit = edits(root, next, sheet, "A1");
        assert_eq!(edit["numberFormat"], "decimal1");
        assert_eq!(edit["styleId"], 1);
        assert_eq!(cell(&cells, "B1")["formula"], "A1 + $A$1");
        assert_eq!(cell(&cells, "B1")["cachedValue"], Value::Null);
        assert_eq!(cell(&cells, "C1")["dataType"], "error");
        assert_eq!(edits(root, next, sheet, "C1")["valueType"], "error");
        assert_eq!(cell(&cells, "D1")["dataType"], "text");
        assert_eq!(cell(&cells, "D1")["error"], Value::Null);
        assert_eq!(cell(&cells, "E1")["value"], true);
        assert_eq!(cell(&cells, "F1")["value"], "");
        assert_eq!(cell(&cells, "F1")["dataType"], "text");
        let blank = cell(&cells, "G1");
        assert_eq!(blank["value"], Value::Null);
        assert_eq!(blank["styleId"], 0);
        assert!(blank.get("formula").is_none());
        assert!(blank.get("bold").is_none());
        let blank_edit = edits(root, next, sheet, "G1");
        assert_eq!(blank_edit["styleId"], 0);
        assert!(blank_edit.get("bold").is_none());
        assert!(blank_edit.get("fillColor").is_none());
    }
    assert_eq!(read(root, next, SOURCE, "A1:G1"), source);
    assert_eq!(read(root, version, FIRST, "A1:G1"), old);
    assert_eq!(fs::read(root.join("source.xlsx")).unwrap(), bytes);
}
#[test]
fn across_contents_copies_formulas_and_blanks_while_retaining_destination_styles() {
    let fixture = ordinary_fixture();
    let root = fixture.path();
    let version = decorated(root);
    let before = read(root, version, FIRST, "A1:G1");
    let result = apply(root, version, vec![command(FIRST, "A1:G1", "contents")]).unwrap();
    let next = result["version"].as_u64().unwrap();
    let cells = read(root, next, FIRST, "A1:G1");
    assert_eq!(result["contentChanged"], true);
    assert_eq!(cell(&cells, "B1")["formula"], "A1 + $A$1");
    for name in ["A1", "B1", "C1", "D1", "E1", "F1", "G1"] {
        assert_eq!(
            cell(&cells, name)["styleId"],
            cell(&before, name)["styleId"]
        );
    }
    let blank = cell(&cells, "G1");
    assert_eq!(blank["value"], Value::Null);
    assert_eq!(blank["bold"], true);
    assert_eq!(blank["fillColor"], "AA0000");
    assert_eq!(edits(root, next, FIRST, "G1")["bold"], true);
    assert!(edits(root, next, FIRST, "G1").get("styleId").is_none());
}
#[test]
fn across_formats_preserves_current_and_pending_formula_fields_and_pending_policy() {
    for pending in [false, true] {
        let fixture = ordinary_fixture();
        let root = fixture.path();
        let mut version = decorated(root);
        let db = crate::database(root).unwrap();
        let mut value = cell(&read(root, version, FIRST, "B1"), "B1");
        value["calculation"] = json!({"status":"calculated","value":21});
        db.execute("INSERT INTO versions(parent) VALUES(?1)", [version])
            .unwrap();
        let calculated = db.last_insert_rowid();
        let edit = inherited_edit(&db, version, FIRST, 0, 1).unwrap();
        crate::range::store_calculated_change(&db, calculated, FIRST, 0, 1, &value, &edit).unwrap();
        version = calculated as u64;
        drop(db);
        let metadata = crate::package_edit::auditing_metadata(root).unwrap();
        crate::calculation_policy::command(root, &json!({"method":"initialize","version":version,"view":metadata,"initialPolicy":{"mode":if pending {"manual"} else {"auto"},"recalculateBeforeSave":false}})).unwrap();
        if pending {
            let db = crate::database(root).unwrap();
            db.execute("INSERT OR REPLACE INTO calculation_pending(version,sheet,r,c,pending) VALUES(?1,?2,0,1,1)", params![version,FIRST]).unwrap();
        }
        let before = read(root, version, FIRST, "A1:G1");
        let db = crate::database(root).unwrap();
        let policy = crate::calculation_policy::options(&db, version).unwrap();
        let pending_count = crate::calculation_policy::pending_count(&db, version).unwrap();
        drop(db);
        let result = apply(root, version, vec![command(FIRST, "A1:G1", "formats")]).unwrap();
        let next = result["version"].as_u64().unwrap();
        let after = read(root, next, FIRST, "A1:G1");
        assert_eq!(result["contentChanged"], false);
        assert_eq!(result["formulaTouched"], false);
        assert_eq!(result["formulaDelta"], json!({}));
        for name in ["A1", "B1", "C1", "D1", "E1", "F1", "G1"] {
            assert_eq!(
                content(&cell(&after, name)),
                content(&cell(&before, name)),
                "{name}, pending={pending}"
            );
        }
        assert_eq!(cell(&after, "A1")["styleId"], 1);
        assert_eq!(cell(&after, "G1")["styleId"], 0);
        assert_eq!(cell(&after, "G1")["formula"], "A1*2");
        assert!(cell(&after, "G1").get("bold").is_none());
        let db = crate::database(root).unwrap();
        assert_eq!(
            crate::calculation_policy::options(&db, next).unwrap(),
            policy
        );
        assert_eq!(
            crate::calculation_policy::pending_count(&db, next).unwrap(),
            pending_count
        );
        if pending {
            assert_eq!(cell(&after, "B1")["calculation"]["pending"], true);
            assert_eq!(cell(&after, "B1")["calculation"]["value"], Value::Null);
        }
    }
}
#[test]
fn across_closed_schema_distinct_run_identity_limits_and_ordinary_output_compatibility() {
    let fixture = ordinary_fixture();
    let root = fixture.path();
    let before = pins(root);
    let mut extra = command(FIRST, "A1", "all");
    extra["source"] = json!("A1");
    let mut same = command(SOURCE, "A1", "all");
    same["sourceSheet"] = json!(SOURCE);
    for commands in [
        vec![extra],
        vec![same],
        vec![command(FIRST, "A1", "values")],
        vec![command(FIRST, "a1", "all")],
        vec![command(FIRST, "A1:A10001", "all")],
        vec![command(FIRST, "A1", "all"), command(FIRST, "A1", "all")],
        vec![command(FIRST, "A1", "all"), command(SECOND, "B1", "all")],
        vec![
            command(FIRST, "A1", "all"),
            command(SECOND, "A1", "contents"),
        ],
    ] {
        assert!(apply(root, 0, commands).is_err());
        assert_eq!(pins(root), before);
    }
    let accepted = apply(
        root,
        0,
        vec![
            command(FIRST, "A1", "all"),
            json!({"type":"clearRange","sheet":FIRST,"range":"A1"}),
            command(FIRST, "A1", "contents"),
        ],
    )
    .unwrap();
    assert_eq!(accepted["contentChanged"], true);
    assert_number(
        &cell(
            &read(root, accepted["version"].as_u64().unwrap(), FIRST, "A1"),
            "A1",
        ),
        42.0,
    );
    let ordinary = apply(
        root,
        0,
        vec![json!({"type":"clearRange","sheet":FIRST,"range":"A1"})],
    )
    .unwrap();
    assert!(ordinary.get("contentChanged").is_none());
}
#[test]
fn across_late_destination_refusal_rolls_back_first_native_candidate_and_history() {
    let second = target_data().replace("r=\"G1\" s=\"1\"", "r=\"G1\" s=\"1\" cm=\"1\"");
    let fixture = fixture(&source_data(), &target_data(), &second, &[]);
    let root = fixture.path();
    let before = pins(root);
    let error = apply(
        root,
        0,
        vec![
            command(FIRST, "A1:G1", "all"),
            command(SECOND, "A1:G1", "all"),
        ],
    )
    .unwrap_err();
    assert!(
        error.contains("source facts") || error.contains("attributes"),
        "{error}"
    );
    assert_eq!(pins(root), before);
    let db = crate::database(root).unwrap();
    assert_eq!(
        db.query_row("SELECT count(*) FROM versions", [], |r| r.get::<_, i64>(0))
            .unwrap(),
        1
    );
    assert_eq!(
        db.query_row("SELECT count(*) FROM changes", [], |r| r.get::<_, i64>(0))
            .unwrap(),
        0
    );
}
#[test]
fn across_metadata_dependency_precision_and_rich_source_profiles_fail_closed() {
    let cases = [
        (
            format!(
                "<worksheet xmlns=\"{NS}\"><sheetData>{}</sheetData><hyperlinks><hyperlink ref=\"Z1\" location=\"Template!A1\"/></hyperlinks></worksheet>",
                source_data()
            ),
            "",
            "metadata",
        ),
        (
            format!(
                "<worksheet xmlns=\"{NS}\"><sheetData>{}</sheetData><conditionalFormatting sqref=\"Z1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting></worksheet>",
                source_data()
            ),
            "",
            "metadata",
        ),
        (
            format!(
                "<worksheet xmlns=\"{NS}\"><cols><col min=\"1\" max=\"1\" style=\"1\"/></cols><sheetData>{}</sheetData></worksheet>",
                source_data()
            ),
            "",
            "inherited",
        ),
        (
            format!(
                "<worksheet xmlns=\"{NS}\"><sheetData>{}</sheetData></worksheet>",
                source_data().replace("A1 + $A$1", "INDIRECT(\"A1\")")
            ),
            "",
            "dependencies",
        ),
        (
            format!(
                "<worksheet xmlns=\"{NS}\"><sheetData>{}</sheetData></worksheet>",
                source_data().replace("<v>42</v>", "<v>9007199254740993</v>")
            ),
            "",
            "precision",
        ),
        (
            format!(
                "<worksheet xmlns=\"{NS}\"><sheetData>{}</sheetData></worksheet>",
                source_data().replace(
                    "<c r=\"A1\" s=\"1\"><v>42</v></c>",
                    "<c r=\"A1\" t=\"inlineStr\"><is><r><t>rich</t></r></is></c>"
                )
            ),
            "",
            "source cell",
        ),
        (
            format!(
                "<worksheet xmlns=\"{NS}\"><sheetData>{}</sheetData></worksheet>",
                source_data().replace(
                    "<c r=\"A1\" s=\"1\"><v>42</v></c>",
                    "<c r=\"A1\" t=\"s\"><v>0</v></c>"
                )
            ),
            "rich",
            "shared strings",
        ),
    ];
    for (worksheet, shared, reason) in cases {
        let mut extras = vec![(SOURCE, worksheet)];
        if shared == "rich" {
            extras.push((
                "xl/sharedStrings.xml",
                format!("<sst xmlns=\"{NS}\"><si><r><t>rich</t></r></si></sst>"),
            ));
        }
        let fixture = fixture(&source_data(), &target_data(), &target_data(), &extras);
        let root = fixture.path();
        let before = pins(root);
        let error = apply(root, 0, vec![command(FIRST, "A1:G1", "all")]).unwrap_err();
        assert!(error.contains(reason), "{reason}: {error}");
        assert_eq!(pins(root), before);
    }
}
#[test]
fn across_selected_plain_shared_empty_text_and_closed_error_formulas_are_contents_and_ten_thousand_positions_are_bounded()
 {
    let source = source_data()
        .replace(
            "<c r=\"A1\" s=\"1\"><v>42</v></c>",
            "<c r=\"A1\" s=\"1\" t=\"s\"><v>0</v></c>",
        )
        .replace("A1 + $A$1", "IFERROR(#REF!,99)");
    let fixture = fixture(
        &source,
        &target_data(),
        &target_data(),
        &[(
            "xl/sharedStrings.xml",
            format!("<sst xmlns=\"{NS}\"><si><t xml:space=\"preserve\"> 00123 </t></si></sst>"),
        )],
    );
    let root = fixture.path();
    let copied = apply(root, 0, vec![command(FIRST, "A1:G1", "contents")]).unwrap();
    let next = copied["version"].as_u64().unwrap();
    assert_eq!(
        cell(&read(root, next, FIRST, "A1:G1"), "A1")["value"],
        " 00123 "
    );
    assert_eq!(cell(&read(root, next, FIRST, "A1:G1"), "F1")["value"], "");
    assert_eq!(
        cell(&read(root, next, FIRST, "A1:G1"), "B1")["formula"],
        "IFERROR(#REF!,99)"
    );
    let maximum = apply(root, next, vec![command(FIRST, "J1:J10000", "all")]).unwrap();
    assert_eq!(maximum["contentChanged"], true);
    let last = cell(
        &read(root, maximum["version"].as_u64().unwrap(), FIRST, "J10000"),
        "J10000",
    );
    assert_eq!(last["value"], Value::Null);
    assert_eq!(last["styleId"], 0);
}
