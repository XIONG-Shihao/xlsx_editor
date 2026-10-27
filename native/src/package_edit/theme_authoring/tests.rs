use super::*;
const SHEET: &str = "xl/worksheets/main.xml";
pub(crate) fn fixture() -> (tempfile::TempDir, std::path::PathBuf) {
    let parent = tempfile::tempdir().unwrap();
    fs::write(parent.path().join("owner.json"),json!({"schema":workbook_format!(WORKBOOK_INDEX_OWNER_V1),"pid":std::process::id(),"sessionId":"native-page-theme-fixture","maxSessionBytes":4u64*1024*1024*1024}).to_string()).unwrap();
    let root = parent.path().join("original");
    fs::create_dir(&root).unwrap();
    let mut zip = ZipWriter::new(File::create(root.join("source.xlsx")).unwrap());
    for (part, bytes) in [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/main.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="workbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
        ),
        (
            "xl/workbook.xml",
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Main" sheetId="1" r:id="main"/></sheets><calcPr calcMode="auto" calcOnSave="0" fullCalcOnLoad="0" forceFullCalc="0"/></workbook>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/main.xml"/><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#,
        ),
        (
            "xl/styles.xml",
            r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><name val="Calibri"/><sz val="11"/></font><font><name val="Courier New"/><sz val="12"/><color rgb="FF119922"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf fontId="0" fillId="0" borderId="0" numFmtId="0"/></cellStyleXfs><cellXfs count="2"><xf fontId="0" fillId="0" borderId="0" numFmtId="0"/><xf fontId="1" fillId="0" borderId="0" numFmtId="0"/></cellXfs></styleSheet>"#,
        ),
        (
            SHEET,
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C1"/><sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>SUM(A1,1)</f><v>3</v></c><c r="C1" s="1" t="inlineStr"><is><t>Literal</t></is></c></row></sheetData></worksheet>"#,
        ),
        ("custom/opaque.bin", "unchanged original payload"),
    ] {
        zip.start_file(part, SimpleFileOptions::default()).unwrap();
        zip.write_all(bytes.as_bytes()).unwrap();
    }
    zip.finish().unwrap();
    crate::import::index(&root).unwrap();
    fs::copy(root.join("source.xlsx"), root.join("export.xlsx")).unwrap();
    (parent, root)
}
fn command(id: &str, prior: Option<String>) -> Value {
    json!({"type":"applyPageTheme","sample":id,"expectedThemeRevision":prior})
}
pub(crate) fn sealed(root: &Path, target: &Path) -> Value {
    fs::create_dir(target).unwrap();
    let view = json!({"sheets":[{"name":"Main","path":SHEET,"storedFormulaCount":1,"date1904":false,"merges":[]}],"namedRanges":[],"tables":[]});
    let exported = crate::export_state::prepare(root, &json!({"version":0,"view":view})).unwrap();
    assert_eq!(exported["styleCount"], 0);
    let mut workbook =
        String::from_utf8(read_part(root, "source.xlsx", "xl/workbook.xml")).unwrap();
    let db = crate::database(root).unwrap();
    let options = crate::calculation_policy::options(&db, 0)
        .unwrap()
        .unwrap_or(("auto".into(), false));
    let pending = exported["incomplete"] == true
        || crate::calculation_policy::pending_count(&db, 0).unwrap() > 0;
    drop(db);
    workbook = workbook
        .replace("calcMode=\"auto\"", &format!("calcMode=\"{}\"", options.0))
        .replace(
            "calcOnSave=\"0\"",
            &format!("calcOnSave=\"{}\"", u8::from(options.1)),
        )
        .replace(
            "fullCalcOnLoad=\"0\"",
            &format!("fullCalcOnLoad=\"{}\"", u8::from(pending)),
        )
        .replace(
            "forceFullCalc=\"0\"",
            &format!("forceFullCalc=\"{}\"", u8::from(pending)),
        );
    fs::write(
        root.join("export-metadata.json"),
        json!({"styleIds":[],"replacements":{"xl/workbook.xml":workbook}}).to_string(),
    )
    .unwrap();
    let index: Value = serde_json::from_slice(&fs::read(root.join("index.json")).unwrap()).unwrap();
    crate::export::write(root,&json!({"action":"workbook","operation":"export","root":root,"revision":index["revision"],"version":0,"calculationMetadata":true,"internalRewrite":true,"rewriteTargetRoot":target})).unwrap()
}
fn request(root: &Path, input: &Value, cmd: Value) -> Value {
    json!({"action":"workbook","operation":"package-edit","root":root,"source":"export.xlsx","version":0,"rewriteInputId":input["rewriteInputId"],"command":cmd})
}
pub(crate) fn read_part(root: &Path, file: &str, name: &str) -> Vec<u8> {
    let mut zip =
        ZipArchive::new(crate::cancel::Reader(File::open(root.join(file)).unwrap())).unwrap();
    part(&mut zip, name).unwrap()
}
pub(crate) fn mutate(root: &Path, file: &str, name: &str, bytes: &[u8]) {
    let mut input = ZipArchive::new(File::open(root.join(file)).unwrap()).unwrap();
    let replacement = root.join("mutated.xlsx");
    let mut output = ZipWriter::new(File::create(&replacement).unwrap());
    let mut found = false;
    for i in 0..input.len() {
        let entry = input.by_index(i).unwrap();
        if entry.name() == name {
            output
                .start_file(name, SimpleFileOptions::default())
                .unwrap();
            output.write_all(bytes).unwrap();
            found = true;
        } else {
            output.raw_copy_file(entry).unwrap();
        }
    }
    if !found {
        output
            .start_file(name, SimpleFileOptions::default())
            .unwrap();
        output.write_all(bytes).unwrap();
    }
    output.finish().unwrap();
    drop(input);
    fs::rename(replacement, root.join(file)).unwrap();
}
#[test]
fn native_samples_attach_normal_role_preserve_literals_and_unrelated_package() {
    let _scope = crate::shared::begin();
    for id in ["neutral", "forest", "plum"] {
        let (owner, root) = fixture();
        let original = crate::hash(&root.join("source.xlsx")).unwrap();
        let token = sealed(&root, &owner.path().join("candidate"));
        let cmd = command(id, None);
        let result = super::super::apply(&root, &request(&root, &token, cmd.clone())).unwrap();
        assert_eq!(result["sample"], id);
        assert_eq!(result["changed"], true);
        assert_eq!(result["contentChanged"], false);
        assert_eq!(
            result["themeRevision"],
            digest(sample(id).unwrap().as_bytes())
        );
        assert!(result["rewriteId"].is_string());
        verify_receipt_package(&root, &cmd, "package.xlsx").unwrap();
        assert_eq!(
            read_part(&root, "package.xlsx", THEME_PART),
            sample(id).unwrap().as_bytes()
        );
        assert_eq!(
            read_part(&root, "export.xlsx", SHEET),
            read_part(&root, "package.xlsx", SHEET)
        );
        assert_eq!(
            read_part(&root, "source.xlsx", "custom/opaque.bin"),
            read_part(&root, "package.xlsx", "custom/opaque.bin")
        );
        let styles = String::from_utf8(read_part(&root, "package.xlsx", "xl/styles.xml")).unwrap();
        assert!(styles.contains("val=\"minor\""));
        assert!(styles.contains("theme=\"1\""));
        assert!(
            styles.contains("<name val=\"Courier New\"/><sz val=\"12\"/><color rgb=\"FF119922\"/>")
        );
        assert_eq!(crate::hash(&root.join("source.xlsx")).unwrap(), original);
    }
}
#[test]
fn native_updates_and_identical_selection_are_revision_guarded() {
    let _scope = crate::shared::begin();
    let (_owner, root) = fixture();
    let first = prepare_from(&root, "export.xlsx", &command("forest", None)).unwrap();
    for (name, value) in first.plan["replacements"].as_object().unwrap() {
        mutate(
            &root,
            "export.xlsx",
            name,
            value.as_str().unwrap().as_bytes(),
        );
    }
    let prior = digest(sample("forest").unwrap().as_bytes());
    assert!(
        !prepare_from(
            &root,
            "export.xlsx",
            &command("forest", Some(prior.clone()))
        )
        .unwrap()
        .changed
    );
    assert!(
        prepare_from(&root, "export.xlsx", &command("plum", Some(prior)))
            .unwrap()
            .changed
    );
    assert!(prepare_from(&root, "export.xlsx", &command("plum", None)).is_err());
}
#[test]
fn source_unknowns_and_caller_metadata_plans_refuse_before_writer_output() {
    let _scope = crate::shared::begin();
    let (owner, root) = fixture();
    let token = sealed(&root, &owner.path().join("candidate"));
    for cmd in [
        json!({"type":"applyPageTheme","sample":"Office","expectedThemeRevision":null}),
        json!({"type":"applyPageTheme","sample":"forest","expectedThemeRevision":null,"xml":"caller XML"}),
        json!({"type":"applyPageTheme","sample":"forest","expectedThemeRevision":"bad"}),
    ] {
        assert!(super::super::apply(&root, &request(&root, &token, cmd)).is_err());
        assert!(!root.join("package.xlsx").exists());
    }
    let mut invalid = request(&root, &token, command("forest", None));
    invalid["version"] = json!(999);
    assert!(super::super::apply(&root, &invalid).is_err());
    fs::write(
        root.join("package-metadata.json"),
        prepare_from(&root, "export.xlsx", &command("forest", None))
            .unwrap()
            .plan
            .to_string(),
    )
    .unwrap();
    invalid.as_object_mut().unwrap().remove("command");
    invalid["version"] = json!(0);
    assert!(super::super::apply(&root, &invalid).is_err());
    let rels = String::from_utf8(read_part(
        &root,
        "export.xlsx",
        "xl/_rels/workbook.xml.rels",
    ))
    .unwrap();
    let foreign=rels.replace("</Relationships>",&format!("<Relationship Id=\"foreign\" Type=\"{THEME_EDGE}\" Target=\"theme/foreign.xml\"/></Relationships>"));
    mutate(
        &root,
        "export.xlsx",
        "xl/_rels/workbook.xml.rels",
        foreign.as_bytes(),
    );
    assert!(prepare_from(&root, "export.xlsx", &command("forest", None)).is_err());
}
#[test]
fn independent_output_verification_rejects_numfmt_cell_and_asset_mutants() {
    let _scope = crate::shared::begin();
    let (owner, root) = fixture();
    let token = sealed(&root, &owner.path().join("candidate"));
    let cmd = command("forest", None);
    super::super::apply(&root, &request(&root, &token, cmd.clone())).unwrap();
    let clean = fs::read(root.join("package.xlsx")).unwrap();
    for (name, bytes) in [
        ("custom/opaque.bin", b"wrong payload".to_vec()),
        (THEME_PART, sample("plum").unwrap().as_bytes().to_vec()),
        (
            "xl/styles.xml",
            String::from_utf8(read_part(&root, "package.xlsx", "xl/styles.xml"))
                .unwrap()
                .replace("numFmtId=\"0\"", "numFmtId=\"14\"")
                .into_bytes(),
        ),
        (
            SHEET,
            String::from_utf8(read_part(&root, "package.xlsx", SHEET))
                .unwrap()
                .replace("<v>2</v>", "<v>999</v>")
                .into_bytes(),
        ),
    ] {
        fs::write(root.join("package.xlsx"), &clean).unwrap();
        mutate(&root, "package.xlsx", name, &bytes);
        assert!(verify_receipt_package(&root, &cmd, "package.xlsx").is_err());
    }
}
