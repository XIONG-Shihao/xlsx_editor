use super::*;
use std::path::PathBuf;
const PART: &str = "xl/custom/sheet-main.xml";
const RELPART: &str = "xl/custom/_rels/sheet-main.xml.rels";
pub(crate) fn fixture(links: &str, edges: &str) -> (tempfile::TempDir, PathBuf) {
    let parent = tempfile::tempdir().unwrap();
    fs::write(parent.path().join("owner.json"),json!({"schema":workbook_format!(WORKBOOK_INDEX_OWNER_V1),"pid":std::process::id(),"sessionId":"native-clear-links-fixture","maxSessionBytes":4u64*1024*1024*1024}).to_string()).unwrap();
    let root = parent.path().join("base");
    fs::create_dir(&root).unwrap();
    let worksheet = format!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:Z80"/><sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1+1</f><v>99</v></c></row><row r="80"><c r="Z80"/></row></sheetData><mergeCells count="1"><mergeCell ref="D4:E4"/></mergeCells>{links}<dataValidations count="1"><dataValidation type="whole" sqref="A1" allowBlank="1"><formula1>0</formula1><formula2>99</formula2></dataValidation></dataValidations></worksheet>"#
    );
    let relationships = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{edges}</Relationships>"#
    );
    let mut zip = ZipWriter::new(File::create(root.join("source.xlsx")).unwrap());
    for (part, text) in [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/custom/sheet-main.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#,
        ),
        (
            "xl/workbook.xml",
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Main" sheetId="1" r:id="main"/></sheets><calcPr calcMode="auto" calcOnSave="0" fullCalcOnLoad="0" forceFullCalc="0"/></workbook>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="custom/sheet-main.xml"/></Relationships>"#,
        ),
        (PART, worksheet.as_str()),
        (RELPART, relationships.as_str()),
        ("custom/opaque.bin", "opaque original bytes"),
    ] {
        zip.start_file(part, SimpleFileOptions::default()).unwrap();
        zip.write_all(text.as_bytes()).unwrap();
    }
    zip.finish().unwrap();
    crate::import::index(&root).unwrap();
    fs::copy(root.join("source.xlsx"), root.join("export.xlsx")).unwrap();
    (parent, root)
}
fn selector(ranges: &[&str]) -> Value {
    json!({"type":"clearHyperlinks","sheet":"Main","ranges":ranges})
}
pub(crate) fn edge(id: &str) -> String {
    format!(
        r#"<Relationship Id="{id}" Type="{HYPERLINK_KIND}" Target="https://example.test/a?x=1&amp;y=2#fragment" TargetMode="External"/>"#
    )
}
pub(crate) fn view() -> Value {
    json!({"sheets":[{"name":"Main","path":PART,"storedFormulaCount":1,"date1904":false,"merges":["D4:E4"]}],"namedRanges":[],"tables":[]})
}
fn selected_output(_root: &Path, plan: &Prepared) -> String {
    plan.plan["fragments"][PART]["hyperlinks"]
        .as_str()
        .unwrap_or_default()
        .to_string()
}
pub(crate) fn sealed(root: &Path, target: &Path, version: u64) -> Value {
    fs::create_dir(target).unwrap();
    let exported =
        crate::export_state::prepare(root, &json!({"version":version,"view":view()})).unwrap();
    assert_eq!(exported["styleCount"], 0);
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(root.join("source.xlsx")).unwrap(),
    ))
    .unwrap();
    let mut workbook = String::new();
    zip.by_name("xl/workbook.xml")
        .unwrap()
        .read_to_string(&mut workbook)
        .unwrap();
    let db = crate::database(root).unwrap();
    let options = crate::calculation_policy::options(&db, version)
        .unwrap()
        .unwrap_or(("auto".into(), false));
    let pending = exported["incomplete"] == true
        || crate::calculation_policy::pending_count(&db, version).unwrap() > 0;
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
    crate::export::write(root,&json!({"action":"workbook","operation":"export","root":root,"revision":index["revision"],"version":version,"calculationMetadata":true,"internalRewrite":true,"rewriteTargetRoot":target})).unwrap()
}
fn clear_request(root: &Path, input: &Value, version: u64, ranges: &[&str]) -> Value {
    json!({"action":"workbook","operation":"package-edit","root":root,"source":"export.xlsx","version":version,"rewriteInputId":input["rewriteInputId"],"command":selector(ranges)})
}
#[test]
fn rectangle_subtraction_is_exact_and_disjoint_without_cell_expansion() {
    let source = rectangle("A1:E5", true).unwrap();
    let residual = source.subtract(rectangle("B2:D4", true).unwrap());
    assert_eq!(
        residual.iter().map(|r| r.reference()).collect::<Vec<_>>(),
        vec!["A1:E1", "A5:E5", "A2:A4", "E2:E4"]
    );
    disjoint(&residual).unwrap();
    let full = rectangle("A1:XFD1048576", true).unwrap();
    assert_eq!(
        full.subtract(rectangle("XFD1048576", true).unwrap()).len(),
        2
    );
    assert!(full.subtract(full).is_empty());
    assert!(command(&selector(&["A1:B2", "B2:C3"])).is_err());
    assert!(command(&selector(&["A1:A1"])).is_err());
    assert!(command(&selector(&["a1"])).is_err());
    assert!(command(&selector(&["A1:XFD1048576"])).is_ok());
    let mut wrong = selector(&["A1"]);
    wrong["mode"] = json!("all");
    assert!(command(&wrong).is_err());
}
#[test]
fn native_graph_preserves_partial_coverage_attributes_and_shared_edges() {
    let _scope = crate::shared::begin();
    let (_owner, root) = fixture(
        r#"<hyperlinks><hyperlink ref="$A$1:$C$3" q:id="shared" tooltip="quoted ref='kept'" display="keep"/><hyperlink ref="E1" q:id="shared"/><hyperlink ref="G1" location="'Other Sheet'!A1"/></hyperlinks>"#,
        &edge("shared"),
    );
    let plan = prepare_from(&root, "export.xlsx", &selector(&["B2", "G1"])).unwrap();
    let xml = selected_output(&root, &plan);
    assert_eq!(
        plan.counts,
        Counts {
            removed: 1,
            trimmed: 1,
            relationships: 0,
            remaining: 5
        }
    );
    for reference in ["A1:C1", "A3:C3", "A2", "C2", "E1"] {
        assert!(xml.contains(&format!("ref=\"{reference}\"")));
    }
    assert!(xml.contains("tooltip=\"quoted ref='kept'\""));
    assert!(xml.contains("display=\"keep\""));
    assert!(plan.plan["replacements"].as_object().unwrap().is_empty());
    let all = prepare_from(&root, "export.xlsx", &selector(&["A1:E3", "G1"])).unwrap();
    assert_eq!(
        all.counts,
        Counts {
            removed: 3,
            trimmed: 0,
            relationships: 1,
            remaining: 0
        }
    );
    assert_eq!(selected_output(&root, &all), "");
    assert!(
        !all.plan["replacements"][RELPART]
            .as_str()
            .unwrap()
            .contains("Id=\"shared\"")
    );
}
#[test]
fn owned_internal_targets_and_no_link_actions_preserve_unrelated_edges() {
    let _scope = crate::shared::begin();
    let internal = format!(
        r#"<Relationship Id="inside" Type="{HYPERLINK_KIND}" Target="sheet-main.xml" TargetMode="Internal"/>"#
    );
    let (_owner, root) = fixture(
        r#"<hyperlinks><hyperlink ref="A1" q:id="inside" location="A1"/><hyperlink ref="B1" location="NamedDestination"/></hyperlinks>"#,
        &(internal + &edge("orphan")),
    );
    let noop = prepare_from(&root, "export.xlsx", &selector(&["Z80"])).unwrap();
    assert_eq!(noop.result()["changed"], false);
    assert_eq!(noop.counts.remaining, 2);
    assert!(noop.plan["fragments"].as_object().unwrap().is_empty());
    let clear = prepare_from(&root, "export.xlsx", &selector(&["A1:B1"])).unwrap();
    assert_eq!(clear.counts.relationships, 1);
    let rels = clear.plan["replacements"][RELPART].as_str().unwrap();
    assert!(rels.contains("Id=\"orphan\""));
    assert!(!rels.contains("Id=\"inside\""));
    assert!(clear.plan["removed"].as_array().unwrap().is_empty());
}
#[test]
fn malformed_and_opaque_hyperlink_graphs_refuse_before_writer_publication() {
    let _scope = crate::shared::begin();
    for(links,edges)in [
        (r#"<hyperlinks><hyperlink ref="A1" q:id="missing"/></hyperlinks>"#.to_string(),String::new()),
        (r#"<hyperlinks><hyperlink ref="A1:B2" location="A1"/><hyperlink ref="B2" location="A2"/></hyperlinks>"#.into(),String::new()),
        (r#"<hyperlinks><hyperlink ref="A1" location="A1" hidden="1"/></hyperlinks>"#.into(),String::new()),
        (r#"<hyperlinks><hyperlink ref="A1" q:id="same"/></hyperlinks>"#.into(),edge("same")+&edge("same")),
        (r#"<hyperlinks><hyperlink ref="A1" q:id="wrong"/></hyperlinks>"#.into(),r#"<Relationship Id="wrong" Type="http://example.test/not-hyperlink" Target="https://example.test/" TargetMode="External"/>"#.into()),
        (r#"<hyperlinks><hyperlink ref="A1" q:id="inside"/></hyperlinks>"#.into(),format!(r#"<Relationship Id="inside" Type="{HYPERLINK_KIND}" Target="missing.xml"/>"#)),
        (r#"<hyperlinks/><hyperlinks/>"#.into(),String::new()),
    ]{let(_owner,root)=fixture(&links,&edges);let before=crate::hash(&root.join("source.xlsx")).unwrap();assert!(prepare_from(&root,"export.xlsx",&selector(&["A1"])).is_err());assert_eq!(crate::hash(&root.join("source.xlsx")).unwrap(),before);assert!(!root.join("package.xlsx").exists());}
}
#[test]
fn sealed_writer_proves_cell_style_rule_and_opaque_package_preservation() {
    let _scope = crate::shared::begin();
    let (parent, root) = fixture(
        r#"<hyperlinks><hyperlink ref="A1:C3" q:id="shared"/></hyperlinks>"#,
        &edge("shared"),
    );
    let target = parent.path().join("candidate");
    let input = sealed(&root, &target, 0);
    let request = clear_request(&root, &input, 0, &["B2"]);
    let before = crate::hash(&root.join("source.xlsx")).unwrap();
    let result = super::super::apply(&root, &request).unwrap();
    assert_eq!(result["contentChanged"], false);
    assert_eq!(result["trimmedHyperlinks"], 1);
    assert_eq!(result["remainingHyperlinks"], 4);
    assert!(result["rewriteId"].is_string());
    assert_eq!(crate::hash(&root.join("source.xlsx")).unwrap(), before);
    let plan = prepare_from(&root, "export.xlsx", &request["command"]).unwrap();
    verify_output(&root, "export.xlsx", "package.xlsx", &plan).unwrap();
    let mut forged = request.clone();
    forged["command"]["ranges"] = json!(["A1"]);
    let other = prepare_from(&root, "export.xlsx", &forged["command"]).unwrap();
    assert!(verify_output(&root, "export.xlsx", "package.xlsx", &other).is_err());
}
#[test]
fn clear_receipt_transfers_current_calculation_without_export_publication_authority() {
    let _scope = crate::shared::begin();
    let (parent, root) = fixture(
        r#"<hyperlinks><hyperlink ref="A1" q:id="only"/></hyperlinks>"#,
        &edge("only"),
    );
    crate::calculation_policy::command(&root,&json!({"method":"initialize","version":0,"view":view(),"initialPolicy":{"mode":"auto","recalculateBeforeSave":false}})).unwrap();
    let calculated = crate::calculation::command(
        &root,
        &json!({"version":0,"view":view(),"revision":"accepted-current","partial":true}),
    )
    .unwrap();
    assert_eq!(calculated["supported"], true);
    let selected = calculated["version"].as_u64().unwrap();
    crate::calculation_policy::command(
        &root,
        &json!({"method":"settle","version":selected,"view":view()}),
    )
    .unwrap();
    let target = parent.path().join("candidate");
    let input = sealed(&root, &target, selected);
    let writer =
        super::super::apply(&root, &clear_request(&root, &input, selected, &["A1"])).unwrap();
    fs::copy(root.join("package.xlsx"), target.join("source.xlsx")).unwrap();
    let index = crate::import::index(&target).unwrap();
    crate::calculation_policy::command(&target,&json!({"method":"initialize","version":0,"view":view(),"initialPolicy":{"mode":"auto","recalculateBeforeSave":false}})).unwrap();
    let request = json!({"action":"workbook","operation":"calculation-transfer","purpose":"clear-hyperlinks","root":target,"version":0,"previousRoot":root,"previousVersion":selected,"sourceRevision":index["revision"],"rewriteId":writer["rewriteId"],"view":view()});
    let mut forged = request.clone();
    forged["rewriteId"] = json!("00000000-0000-4000-8000-000000000000");
    assert!(crate::calculation_transfer::command(&target, &forged).is_err());
    let db = crate::database(&target).unwrap();
    assert_eq!(
        db.query_row("SELECT count(*) FROM versions", [], |r| r.get::<_, u64>(0))
            .unwrap(),
        1
    );
    drop(db);
    let result = crate::calculation_transfer::command(&target, &request).unwrap();
    assert_eq!(
        result["schema"],
        workbook_format!(SCHEMA_CLEAR_HYPERLINKS_CALCULATION_TRANSFER_V1)
    );
    assert_eq!(result["transferredFormulaCells"], 1);
    assert_eq!(result["policy"]["pendingFormulaCells"], 0);
    let accepted = result["version"].as_u64().unwrap();
    let db = crate::database(&target).unwrap();
    let read = crate::range::read_public_db(
        &db,
        &json!({"version":accepted,"sheet":PART,"range":"B1","limit":1}),
    )
    .unwrap();
    assert_eq!(read["cells"][0][1]["calculation"]["value"], 3.0);
    assert_eq!(
        read["cells"][0][1]["calculation"]["revision"],
        "accepted-current"
    );
    drop(db);
    let extent=crate::used_range_transfer::command(&target,&json!({"action":"workbook","operation":"used-range-transfer","root":target,"version":accepted,"sourceRevision":index["revision"],"previousRoot":root,"previousVersion":selected,"previousSourceRevision":serde_json::from_slice::<Value>(&fs::read(root.join("index.json")).unwrap()).unwrap()["revision"],"rewriteId":writer["rewriteId"]})).unwrap();
    assert!(extent["version"].as_u64().unwrap() > accepted);
}

#[test]
fn opaque_owner_reference_retains_edge_and_source_attributes() {
    let _scope = crate::shared::begin();
    let (_owner, root) = fixture(
        r#"<hyperlinks><hyperlink ref="A1" q:id="only"/></hyperlinks><extension xmlns:s="urn:opaque-extension" s:id="only"/>"#,
        &edge("only"),
    );
    let prepared = prepare_from(&root, "export.xlsx", &selector(&["A1"])).unwrap();
    assert_eq!(prepared.counts.removed, 1);
    assert_eq!(prepared.counts.relationships, 0);
    assert!(
        prepared.plan["replacements"]
            .as_object()
            .unwrap()
            .is_empty()
    );
}

#[test]
fn clear_receipt_keeps_manual_stored_cache_pending_and_rejects_changed_candidate() {
    let _scope = crate::shared::begin();
    let (parent, root) = fixture(
        r#"<hyperlinks><hyperlink ref="A1" q:id="only"/></hyperlinks>"#,
        &edge("only"),
    );
    crate::calculation_policy::command(&root,&json!({"method":"initialize","version":0,"view":view(),"initialPolicy":{"mode":"manual","recalculateBeforeSave":false}})).unwrap();
    let target = parent.path().join("candidate");
    let input = sealed(&root, &target, 0);
    let writer = super::super::apply(&root, &clear_request(&root, &input, 0, &["A1"])).unwrap();
    fs::copy(root.join("package.xlsx"), target.join("source.xlsx")).unwrap();
    let index = crate::import::index(&target).unwrap();
    crate::calculation_policy::command(&target,&json!({"method":"initialize","version":0,"view":view(),"initialPolicy":{"mode":"manual","recalculateBeforeSave":false}})).unwrap();
    let request = json!({"action":"workbook","operation":"calculation-transfer","purpose":"clear-hyperlinks","root":target,"version":0,"previousRoot":root,"previousVersion":0,"sourceRevision":index["revision"],"rewriteId":writer["rewriteId"],"view":view()});
    let original = fs::read(target.join("source.xlsx")).unwrap();
    fs::write(target.join("source.xlsx"), b"swapped output").unwrap();
    assert!(crate::calculation_transfer::command(&target, &request).is_err());
    fs::write(target.join("source.xlsx"), original).unwrap();
    let result = crate::calculation_transfer::command(&target, &request).unwrap();
    assert_eq!(result["policy"]["pendingFormulaCells"], 1);
    assert_eq!(result["policy"]["mode"], "manual");
    let db = crate::database(&target).unwrap();
    let read = crate::range::read_public_db(
        &db,
        &json!({"version":result["version"],"sheet":PART,"range":"B1","limit":1}),
    )
    .unwrap();
    let cell = &read["cells"][0][1];
    assert_eq!(cell["calculation"]["pending"], true);
    assert!(cell["calculation"]["value"].is_null());
    assert_eq!(cell["cachedResult"]["value"], 99.0);
}

#[test]
fn residual_metadata_is_bounded_before_retaining_large_replacement_batches() {
    let _scope = crate::shared::begin();
    let mut links = String::from("<hyperlinks>");
    let mut ranges = Vec::new();
    let tooltip = "x".repeat(25_000);
    for index in 0..200 {
        let row = index * 3 + 1;
        links.push_str(&format!(
            "<hyperlink ref=\"A{row}:C{}\" location=\"A1\" tooltip=\"{tooltip}\"/>",
            row + 2
        ));
        ranges.push(format!("B{}", row + 1));
    }
    links.push_str("</hyperlinks>");
    let (_owner, root) = fixture(&links, "");
    let before = crate::hash(&root.join("source.xlsx")).unwrap();
    let command = json!({"type":"clearHyperlinks","sheet":"Main","ranges":ranges});
    let result = prepare_from(&root, "export.xlsx", &command);
    assert!(matches!(result,Err(message) if message.contains("fragment byte bound")));
    assert_eq!(crate::hash(&root.join("source.xlsx")).unwrap(), before);
    assert!(!root.join("package.xlsx").exists());
}
