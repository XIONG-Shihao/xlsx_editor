use super::super::theme_authoring::tests::{fixture, mutate, read_part, sealed};
use super::*;
#[test]
fn portable_comments_omit_shape_ids_but_exact_legacy_records_remain_readable() {
    let notes = vec![note("A1", "Alice", "Text & 東京\nSecond").unwrap()];
    let portable = comments(&notes);
    assert!(!portable.contains("shapeId="));
    assert_eq!(
        comments(&parse_comments(portable.as_bytes()).unwrap()),
        portable
    );
    let legacy = comments_xml(&notes, true);
    assert!(legacy.contains("shapeId=\"1025\""));
    assert_eq!(
        comments(&parse_comments(legacy.as_bytes()).unwrap()),
        portable
    );
    assert!(
        parse_comments(
            legacy
                .replace("shapeId=\"1025\"", "shapeId=\"1026\"")
                .as_bytes()
        )
        .is_err()
    );
    assert!(
        parse_comments(
            portable
                .replace("authorId=\"0\"", "authorId=\"99\"")
                .as_bytes()
        )
        .is_err()
    );
}
const SHEET: &str = "xl/worksheets/main.xml";
fn command(cell: &str, body: Option<&str>, prior: Option<String>) -> Value {
    if let Some(body) = body {
        json!({"type":"setNote","sheet":"Main","cell":cell,"expectedNoteRevision":prior,"note":{"author":"Alice & Bob","text":body}})
    } else {
        json!({"type":"removeNote","sheet":"Main","cell":cell,"expectedNoteRevision":prior})
    }
}
fn request(root: &Path, token: &Value, cmd: Value) -> Value {
    json!({"action":"workbook","operation":"package-edit","root":root,"source":"export.xlsx","version":0,"rewriteInputId":token["rewriteInputId"],"command":cmd})
}
fn inspect(root: &Path) -> Value {
    let index: Value = serde_json::from_slice(&fs::read(root.join("index.json")).unwrap()).unwrap();
    read(root,&json!({"action":"workbook","operation":"notes","root":root,"version":0,"sourceRevision":index["revision"],"sheet":"Main"})).unwrap()
}
fn fresh(root: &Path, target: &Path) {
    fs::copy(root.join("package.xlsx"), target.join("source.xlsx")).unwrap();
    crate::import::index(target).unwrap();
    fs::copy(target.join("source.xlsx"), target.join("export.xlsx")).unwrap();
}
fn applied(root: &Path, target: &Path, cmd: Value) -> Value {
    let token = sealed(root, target);
    super::super::apply(root, &request(root, &token, cmd)).unwrap()
}
#[test]
fn real_comments_vml_and_escaped_text_preserve_cells_and_opaque_parts() {
    let _scope = crate::shared::begin();
    let (owner, root) = fixture();
    let before = crate::hash(&root.join("source.xlsx")).unwrap();
    let cmd = command("B1", Some("  Text & <tag> \r\n\t😀  "), None);
    let token = sealed(&root, &owner.path().join("candidate"));
    let result = super::super::apply(&root, &request(&root, &token, cmd.clone())).unwrap();
    assert_eq!(result["schema"], workbook_format!(SCHEMA_NOTE_WRITER_V1));
    assert_eq!(result["remainingNotes"], 1);
    assert_eq!(result["contentChanged"], false);
    assert_eq!(result["changed"], true);
    verify_receipt_package(&root, &cmd, "package.xlsx").unwrap();
    let (c, v) = parts(SHEET);
    let notes = parse_comments(&read_part(&root, "package.xlsx", &c)).unwrap();
    assert_eq!(notes[0].text, "  Text & <tag> \r\n\t😀  ");
    assert_eq!(notes[0].author, "Alice & Bob");
    let draw = String::from_utf8(read_part(&root, "package.xlsx", &v)).unwrap();
    assert!(draw.contains("ObjectType=\"Note\""));
    assert!(draw.contains("<x:Row>0</x:Row><x:Column>1</x:Column>"));
    let sheet = String::from_utf8(read_part(&root, "package.xlsx", SHEET)).unwrap();
    assert!(sheet.contains("<f>SUM(A1,1)</f><v>3</v>"));
    assert!(sheet.contains("legacyDrawing"));
    assert_eq!(
        read_part(&root, "package.xlsx", "custom/opaque.bin"),
        read_part(&root, "source.xlsx", "custom/opaque.bin")
    );
    assert_eq!(crate::hash(&root.join("source.xlsx")).unwrap(), before);
}
#[test]
fn add_edit_multiple_authors_remove_last_and_noop_reimport_as_real_note_models() {
    let _scope = crate::shared::begin();
    let (owner, root) = fixture();
    assert_eq!(inspect(&root)["notes"], json!([]));
    let first = owner.path().join("one");
    applied(&root, &first, command("C1", Some("First"), None));
    fresh(&root, &first);
    let state = inspect(&first);
    assert_eq!(state["canEdit"], true);
    assert_eq!(state["notes"][0]["cell"], "C1");
    let prior = state["noteRevision"].as_str().unwrap().to_owned();
    assert!(
        !prepare_from(
            &first,
            "export.xlsx",
            &command("C1", Some("First"), Some(prior.clone()))
        )
        .unwrap()
        .changed
    );
    assert!(prepare_from(&first, "export.xlsx", &command("C1", Some("Changed"), None)).is_err());
    let second = owner.path().join("two");
    let mut cmd = command("A1", Some("Added"), Some(prior));
    cmd["note"]["author"] = json!("Carol");
    applied(&first, &second, cmd);
    fresh(&first, &second);
    let state = inspect(&second);
    assert_eq!(state["notes"].as_array().unwrap().len(), 2);
    assert_eq!(state["notes"][0]["author"], "Carol");
    assert_eq!(state["notes"][1]["text"], "First");
    let third = owner.path().join("three");
    applied(
        &second,
        &third,
        command(
            "C1",
            Some("Edited"),
            Some(state["noteRevision"].as_str().unwrap().into()),
        ),
    );
    fresh(&second, &third);
    assert_eq!(inspect(&third)["notes"][1]["text"], "Edited");
    let fourth = owner.path().join("four");
    applied(
        &third,
        &fourth,
        command(
            "A1",
            None,
            Some(inspect(&third)["noteRevision"].as_str().unwrap().into()),
        ),
    );
    fresh(&third, &fourth);
    assert_eq!(inspect(&fourth)["notes"].as_array().unwrap().len(), 1);
    let fifth = owner.path().join("five");
    applied(
        &fourth,
        &fifth,
        command(
            "C1",
            None,
            Some(inspect(&fourth)["noteRevision"].as_str().unwrap().into()),
        ),
    );
    fresh(&fourth, &fifth);
    assert_eq!(inspect(&fifth)["noteRevision"], Value::Null);
    assert_eq!(inspect(&fifth)["notes"], json!([]));
    let names = clear_hyperlinks::package_names(
        &ZipArchive::new(crate::cancel::Reader(
            File::open(fifth.join("source.xlsx")).unwrap(),
        ))
        .unwrap(),
    )
    .unwrap();
    let (c, v) = parts(SHEET);
    assert!(!names.contains(&c) && !names.contains(&v));
    assert!(
        !String::from_utf8(read_part(&fifth, "source.xlsx", SHEET))
            .unwrap()
            .contains("legacyDrawing")
    );
    assert!(
        !prepare_from(&fifth, "export.xlsx", &command("C1", None, None))
            .unwrap()
            .changed
    );
}
#[test]
fn forged_output_cannot_change_note_model_vml_or_unrelated_worksheet_facts() {
    let _scope = crate::shared::begin();
    let (owner, root) = fixture();
    let cmd = command("A1", Some("Original note"), None);
    applied(&root, &owner.path().join("candidate"), cmd.clone());
    let clean = fs::read(root.join("package.xlsx")).unwrap();
    let (c, v) = parts(SHEET);
    for (name, bytes) in [
        (
            c,
            String::from_utf8(read_part(&root, "package.xlsx", &parts(SHEET).0))
                .unwrap()
                .replace("Original note", "Forged text")
                .into_bytes(),
        ),
        (
            v,
            String::from_utf8(read_part(&root, "package.xlsx", &parts(SHEET).1))
                .unwrap()
                .replace("<x:Column>0</x:Column>", "<x:Column>1</x:Column>")
                .into_bytes(),
        ),
        (
            SHEET.into(),
            String::from_utf8(read_part(&root, "package.xlsx", SHEET))
                .unwrap()
                .replace("<v>2</v>", "<v>77</v>")
                .into_bytes(),
        ),
        ("xl/styles.xml".into(), b"forged".to_vec()),
        ("custom/opaque.bin".into(), b"forged".to_vec()),
    ] {
        fs::write(root.join("package.xlsx"), &clean).unwrap();
        mutate(&root, "package.xlsx", &name, &bytes);
        assert!(
            verify_receipt_package(&root, &cmd, "package.xlsx").is_err(),
            "{name}"
        );
    }
}
#[test]
fn foreign_threaded_shared_modified_and_extended_graphs_have_no_edit_grant() {
    let _scope = crate::shared::begin();
    let (owner, root) = fixture();
    let first = owner.path().join("owned");
    applied(&root, &first, command("A1", Some("Original"), None));
    fresh(&root, &first);
    let original = fs::read(first.join("source.xlsx")).unwrap();
    let prior = inspect(&first)["noteRevision"].as_str().unwrap().to_owned();
    let (c, v) = parts(SHEET);
    for (index,(name,bytes))in [(c.clone(),String::from_utf8(read_part(&first,"source.xlsx",&c)).unwrap().replace("<text>","<text><r/>").into_bytes()),(v,String::from_utf8(read_part(&first,"source.xlsx",&parts(SHEET).1)).unwrap().replace("144pt","900pt").into_bytes()),("custom/_rels/other.xml.rels".into(),format!("<Relationships xmlns=\"{}\"><Relationship Id=\"shared\" Type=\"{COMMENTS_EDGE}\" Target=\"/{c}\"/></Relationships>",String::from_utf8_lossy(PACKAGE_REL_NS)).into_bytes())].into_iter().enumerate(){
  let target=owner.path().join(format!("mutant-{index}"));fs::create_dir(&target).unwrap();fs::write(target.join("source.xlsx"),&original).unwrap();mutate(&target,"source.xlsx",&name,&bytes);crate::import::index(&target).unwrap();fs::copy(target.join("source.xlsx"),target.join("export.xlsx")).unwrap();let raw=inspect(&target);assert_eq!(raw["canEdit"],false);assert_eq!(raw["notes"],Value::Null);assert!(prepare_from(&target,"export.xlsx",&command("A1",None,Some(prior.clone()))).is_err());
 }
    let target = owner.path().join("threaded");
    fs::create_dir(&target).unwrap();
    fs::write(target.join("source.xlsx"), &original).unwrap();
    let rels=String::from_utf8(read_part(&first,"source.xlsx",&rel_part(SHEET).unwrap())).unwrap().replace("</Relationships>","<Relationship Id=\"thread\" Type=\"http://schemas.microsoft.com/office/2017/10/relationships/threadedComment\" Target=\"/xl/threadedComments/thread1.xml\"/></Relationships>");
    mutate(
        &target,
        "source.xlsx",
        &rel_part(SHEET).unwrap(),
        rels.as_bytes(),
    );
    mutate(
        &target,
        "source.xlsx",
        "xl/threadedComments/thread1.xml",
        b"foreign thread",
    );
    crate::import::index(&target).unwrap();
    assert_eq!(inspect(&target)["canEdit"], false);
}
#[test]
fn bounds_whitespace_unicode_and_stale_or_caller_owned_requests_refuse() {
    let _scope = crate::shared::begin();
    let (owner, root) = fixture();
    let token = sealed(&root, &owner.path().join("candidate"));
    for cmd in [
        command("a1", Some("note"), None),
        command("XFE1", Some("note"), None),
        command("A0", Some("note"), None),
        command("A1", Some(""), None),
        command("A1", Some(&"n".repeat(4097)), None),
        command("A1", Some("invalid\u{0001}"), None),
    ] {
        assert!(super::super::apply(&root, &request(&root, &token, cmd)).is_err());
        assert!(!root.join("package.xlsx").exists());
    }
    let mut cmd = command("A1", Some("note"), None);
    cmd["xml"] = json!("caller bytes");
    assert!(super::super::apply(&root, &request(&root, &token, cmd)).is_err());
    let mut req = request(&root, &token, command("A1", Some("note"), None));
    req["version"] = json!(999);
    assert!(super::super::apply(&root, &req).is_err());
    fs::write(
        root.join("package-metadata.json"),
        prepare_from(&root, "export.xlsx", &command("A1", Some("note"), None))
            .unwrap()
            .plan
            .to_string(),
    )
    .unwrap();
    req.as_object_mut().unwrap().remove("command");
    req["version"] = json!(0);
    assert!(super::super::apply(&root, &req).is_err());
    let far = note("XFD1048576", "Author", "End note").unwrap();
    assert!(
        vml(&[far])
            .unwrap()
            .contains("<x:Row>1048575</x:Row><x:Column>16383</x:Column>")
    );
    let mut notes = (0..256)
        .map(|i| note(&format!("A{}", i + 1), "Author", "note").unwrap())
        .collect::<Vec<_>>();
    notes.sort_by_key(|n| position(&n.cell).unwrap());
    assert_eq!(
        parse_comments(comments(&notes).as_bytes()).unwrap().len(),
        256
    );
    notes.push(note("A257", "Author", "note").unwrap());
    assert!(parse_comments(comments(&notes).as_bytes()).is_err());
}

#[test]
fn note_receipt_preserves_current_formula_truth_and_refuses_a_different_metadata_purpose() {
    use super::super::clear_hyperlinks::tests::{
        fixture as links_fixture, sealed as links_sealed, view,
    };
    let _scope = crate::shared::begin();
    let (owner, root) = links_fixture("", "");
    crate::calculation_policy::command(&root,&json!({"method":"initialize","version":0,"view":view(),"initialPolicy":{"mode":"auto","recalculateBeforeSave":false}})).unwrap();
    let calculated = crate::calculation::command(
        &root,
        &json!({"version":0,"view":view(),"revision":"accepted-current","partial":true}),
    )
    .unwrap();
    let version = calculated["version"].as_u64().unwrap();
    crate::calculation_policy::command(
        &root,
        &json!({"method":"settle","version":version,"view":view()}),
    )
    .unwrap();
    let target = owner.path().join("candidate");
    let token = links_sealed(&root, &target, version);
    let mut req = request(&root, &token, command("B1", Some("Computed total"), None));
    req["version"] = json!(version);
    let receipt = super::super::apply(&root, &req).unwrap();
    fs::copy(root.join("package.xlsx"), target.join("source.xlsx")).unwrap();
    let index = crate::import::index(&target).unwrap();
    crate::calculation_policy::command(&target,&json!({"method":"initialize","version":0,"view":view(),"initialPolicy":{"mode":"auto","recalculateBeforeSave":false}})).unwrap();
    let mut transfer = json!({"action":"workbook","operation":"calculation-transfer","purpose":"hyperlink-authoring","root":target,"version":0,"previousRoot":root,"previousVersion":version,"sourceRevision":index["revision"],"rewriteId":receipt["rewriteId"],"view":view()});
    assert!(crate::calculation_transfer::command(&target, &transfer).is_err());
    transfer["purpose"] = json!("note-authoring");
    let result = crate::calculation_transfer::command(&target, &transfer).unwrap();
    assert_eq!(
        result["schema"],
        workbook_format!(SCHEMA_NOTE_CALCULATION_TRANSFER_V1)
    );
    assert_eq!(result["transferredFormulaCells"], 1);
    assert_eq!(result["policy"]["pendingFormulaCells"], 0);
    let db = crate::database(&target).unwrap();
    let read=crate::range::read_public_db(&db,&json!({"version":result["version"],"sheet":"xl/custom/sheet-main.xml","range":"B1","limit":1})).unwrap();
    assert_eq!(read["cells"][0][1]["calculation"]["value"], 3.0);
}

#[test]
fn reserved_thread_author_tokens_and_a_legacy_fragment_missing_from_metadata_refuse() {
    let _scope = crate::shared::begin();
    let (_owner, root) = fixture();
    let mut cmd = command("A1", Some("Note"), None);
    cmd["note"]["author"] = json!("tc={00000000-0000-4000-8000-000000000000}");
    assert!(prepare_from(&root, "export.xlsx", &cmd).is_err());
    let source=String::from_utf8(read_part(&root,"export.xlsx",SHEET)).unwrap().replace("</worksheet>","<legacyDrawing xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"foreign\"/></worksheet>");
    mutate(&root, "export.xlsx", SHEET, source.as_bytes());
    assert!(prepare_from(&root, "export.xlsx", &command("A1", Some("Note"), None)).is_err());
}

#[test]
fn protected_owned_notes_remain_readable_but_native_writes_refuse() {
    let _scope = crate::shared::begin();
    let (owner, root) = fixture();
    let first = owner.path().join("first");
    applied(&root, &first, command("A1", Some("Protected body"), None));
    let target = owner.path().join("protected");
    fs::create_dir(&target).unwrap();
    fs::copy(root.join("package.xlsx"), target.join("source.xlsx")).unwrap();
    let sheet = String::from_utf8(read_part(&target, "source.xlsx", SHEET))
        .unwrap()
        .replace(
            "<legacyDrawing",
            "<sheetProtection sheet=\"1\" objects=\"1\"/><legacyDrawing",
        );
    mutate(&target, "source.xlsx", SHEET, sheet.as_bytes());
    crate::import::index(&target).unwrap();
    fs::copy(target.join("source.xlsx"), target.join("export.xlsx")).unwrap();
    let raw = inspect(&target);
    assert_eq!(raw["canEdit"], false);
    assert_eq!(raw["notes"][0]["text"], "Protected body");
    assert!(
        prepare_from(
            &target,
            "export.xlsx",
            &command(
                "A1",
                Some("Wrong"),
                raw["noteRevision"].as_str().map(str::to_owned)
            )
        )
        .is_err()
    );
    assert!(!target.join("package.xlsx").exists());
}
