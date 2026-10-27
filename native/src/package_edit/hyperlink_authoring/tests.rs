use super::super::clear_hyperlinks::tests::{edge, fixture, sealed, view};
use super::*;
const PART: &str = "xl/custom/sheet-main.xml";
const RELPART: &str = "xl/custom/_rels/sheet-main.xml.rels";
fn set(link: Value, revision: Option<String>) -> Value {
    json!({"type":"setHyperlink","sheet":"Main","cell":"A1","expectedHyperlinkRevision":revision,"link":link,"tooltip":"Open this & inspect \"target\""})
}
fn web(revision: Option<String>) -> Value {
    set(
        json!({"kind":"url","target":"https://example.test/path?q=one%20two&x=3#section"}),
        revision,
    )
}
fn request(root: &Path, input: &Value, command: Value) -> Value {
    json!({"action":"workbook","operation":"package-edit","root":root,"source":"export.xlsx","version":0,"rewriteInputId":input["rewriteInputId"],"command":command})
}
fn revision(root: &Path, reference: &str) -> String {
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(root.join("metadata.xlsx")).unwrap(),
    ))
    .unwrap();
    let tree = Tree::parse(&part(&mut zip, PART).unwrap()).unwrap();
    let rels = Tree::parse(&part(&mut zip, RELPART).unwrap()).unwrap();
    let index = tree
        .nodes
        .iter()
        .position(|n| {
            n.name == "hyperlink" && n.attrs.get("ref").map(String::as_str) == Some(reference)
        })
        .unwrap();
    fingerprint(&tree, index, Some(&rels)).unwrap()
}
fn output(root: &Path, name: &str) -> String {
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(root.join("package.xlsx")).unwrap(),
    ))
    .unwrap();
    String::from_utf8(part(&mut zip, name).unwrap()).unwrap()
}
#[test]
fn creates_web_link_from_sealed_source_and_preserves_cells_and_unrelated_edges() {
    let _scope = crate::shared::begin();
    let (owner, root) = fixture(
        r#"<hyperlinks><hyperlink ref="E1" q:id="old"/></hyperlinks>"#,
        &edge("old"),
    );
    let before = crate::hash(&root.join("source.xlsx")).unwrap();
    let target = owner.path().join("candidate");
    let input = sealed(&root, &target, 0);
    let command = web(None);
    let result = super::super::apply(&root, &request(&root, &input, command.clone())).unwrap();
    assert_eq!(
        result["schema"],
        workbook_format!(SCHEMA_HYPERLINK_WRITER_V1)
    );
    assert_eq!(result["remainingHyperlinks"], 2);
    assert_eq!(result["contentChanged"], false);
    assert_eq!(result["priorHyperlinkRevision"], Value::Null);
    assert!(result["rewriteId"].is_string());
    verify_receipt_package(&root, &command, "package.xlsx").unwrap();
    let sheet = output(&root, PART);
    assert!(sheet.contains("ref=\"E1\" q:id=\"old\""));
    assert!(sheet.contains(workbook_format!(XML_B441F277FE)));
    assert!(sheet.contains("<f>A1+1</f>"));
    let rels = output(&root, RELPART);
    assert!(rels.contains("Id=\"old\""));
    assert!(rels.contains("q=one%20two&amp;x=3#section"));
    assert_eq!(crate::hash(&root.join("source.xlsx")).unwrap(), before);
}
#[test]
fn updates_web_to_internal_with_exact_revision_and_removes_only_exclusive_edge() {
    let _scope = crate::shared::begin();
    let (owner, root) = fixture(
        r#"<hyperlinks><hyperlink ref="A1" q:id="old"/><hyperlink ref="E1" q:id="old"/></hyperlinks>"#,
        &edge("old"),
    );
    let prior = revision(&root, "A1");
    let command = set(
        json!({"kind":"internal","sheet":"Main","cell":"Z80"}),
        Some(prior.clone()),
    );
    let plan = prepare_from(&root, "export.xlsx", &command).unwrap();
    assert_eq!(plan.revision, Some(prior));
    assert!(
        plan.plan["replacements"].as_object().unwrap().is_empty(),
        "Shared external edge is retained"
    );
    let target = owner.path().join("candidate");
    let input = sealed(&root, &target, 0);
    super::super::apply(&root, &request(&root, &input, command.clone())).unwrap();
    verify_receipt_package(&root, &command, "package.xlsx").unwrap();
    let sheet = output(&root, PART);
    assert!(sheet.contains("location=\"'Main'!Z80\""));
    assert!(sheet.contains("ref=\"E1\" q:id=\"old\""));
    let (_owner, root) = fixture(
        r#"<hyperlinks><hyperlink ref="A1" q:id="old"/></hyperlinks>"#,
        &edge("old"),
    );
    let command = set(
        json!({"kind":"internal","sheet":"Main","cell":"A2"}),
        Some(revision(&root, "A1")),
    );
    let plan = prepare_from(&root, "export.xlsx", &command).unwrap();
    assert!(
        !plan.plan["replacements"][RELPART]
            .as_str()
            .unwrap()
            .contains("Id=\"old\"")
    );
}
#[test]
fn removal_and_absent_noop_are_native_derived_and_revision_guarded() {
    let _scope = crate::shared::begin();
    let (owner, root) = fixture(
        r#"<hyperlinks><hyperlink ref="A1" q:id="old"/></hyperlinks>"#,
        &edge("old"),
    );
    let remove = json!({"type":"removeHyperlink","sheet":"Main","cell":"A1","expectedHyperlinkRevision":revision(&root,"A1")});
    let target = owner.path().join("candidate");
    let input = sealed(&root, &target, 0);
    let result = super::super::apply(&root, &request(&root, &input, remove.clone())).unwrap();
    assert_eq!(result["changed"], true);
    assert_eq!(result["remainingHyperlinks"], 0);
    verify_receipt_package(&root, &remove, "package.xlsx").unwrap();
    assert!(!output(&root, PART).contains("<hyperlinks"));
    assert!(!output(&root, RELPART).contains("Id=\"old\""));
    let (_owner, root) = fixture("", "");
    let none = json!({"type":"removeHyperlink","sheet":"Main","cell":"A1","expectedHyperlinkRevision":null});
    assert!(!prepare_from(&root, "export.xlsx", &none).unwrap().changed);
}
#[test]
fn closed_commands_refuse_stale_ranges_unknown_targets_and_script_or_file_urls() {
    let _scope = crate::shared::begin();
    let (_owner, root) = fixture("", "");
    for target in [
        "javascript:alert(1)",
        "file:///tmp/file",
        "mailto:user@example.test",
        "https://user:password@example.test/",
        "https://example.test\\bad",
        "https://example.test/%QQ",
        "https://example.test:99999/",
        "https://",
        "https://example..test/",
        "https://example.test/ space",
    ] {
        let c = set(json!({"kind":"url","target":target}), None);
        assert!(prepare_from(&root, "export.xlsx", &c).is_err(), "{target}");
    }
    for target in [
        "https://example.test/",
        "http://localhost:8080/a",
        "https://[::1]/a",
    ] {
        url(target).unwrap();
    }
    for command in [
        set(
            json!({"kind":"internal","sheet":"Missing","cell":"A1"}),
            None,
        ),
        set(
            json!({"kind":"internal","sheet":"Main","cell":"A1:B2"}),
            None,
        ),
        set(
            json!({"kind":"url","target":"https://example.test/","extra":true}),
            None,
        ),
        web(Some("a".repeat(64))),
    ] {
        assert!(prepare_from(&root, "export.xlsx", &command).is_err());
    }
    let mut wrong = web(None);
    wrong["cell"] = json!("$A$1");
    assert!(prepare_from(&root, "export.xlsx", &wrong).is_err());
    wrong = web(None);
    wrong["value"] = json!("must not rewrite cells");
    assert!(prepare_from(&root, "export.xlsx", &wrong).is_err());
    let (_owner, root) = fixture(
        r#"<hyperlinks><hyperlink ref="A1:B2" location="A1"/></hyperlinks>"#,
        "",
    );
    assert!(
        prepare_from(&root, "export.xlsx", &web(None))
            .unwrap_err_text()
            .contains("Range/shared")
    );
}
trait ErrorText {
    fn unwrap_err_text(self) -> String;
}
impl<T> ErrorText for Result<T> {
    fn unwrap_err_text(self) -> String {
        match self {
            Err(e) => e,
            Ok(_) => panic!("expected refusal"),
        }
    }
}
#[test]
fn malformed_foreign_graph_refuses_even_disjoint_creation() {
    let _scope = crate::shared::begin();
    for (links, rels) in [
        (
            r#"<hyperlinks><hyperlink ref="E1" q:id="missing"/></hyperlinks>"#,
            String::new(),
        ),
        (
            r#"<hyperlinks><hyperlink ref="E1" location="A1" foreign="1"/></hyperlinks>"#,
            String::new(),
        ),
        (
            r#"<hyperlinks><hyperlink ref="E1" q:id="old"/></hyperlinks>"#,
            edge("old") + &edge("old"),
        ),
    ] {
        let (_owner, root) = fixture(links, &rels);
        assert!(prepare_from(&root, "export.xlsx", &web(None)).is_err());
        assert!(!root.join("package.xlsx").exists());
    }
}
#[test]
fn authoring_receipt_preserves_current_calculation_and_cannot_impersonate_clear() {
    let _scope = crate::shared::begin();
    let (owner, root) = fixture("", "");
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
    let input = sealed(&root, &target, version);
    let mut req = request(&root, &input, web(None));
    req["version"] = json!(version);
    let receipt = super::super::apply(&root, &req).unwrap();
    fs::copy(root.join("package.xlsx"), target.join("source.xlsx")).unwrap();
    crate::import::index(&target).unwrap();
    let index: Value =
        serde_json::from_slice(&fs::read(target.join("index.json")).unwrap()).unwrap();
    crate::calculation_policy::command(&target,&json!({"method":"initialize","version":0,"view":view(),"initialPolicy":{"mode":"auto","recalculateBeforeSave":false}})).unwrap();
    let mut transfer = json!({"action":"workbook","operation":"calculation-transfer","purpose":"clear-hyperlinks","root":target,"version":0,"previousRoot":root,"previousVersion":version,"sourceRevision":index["revision"],"rewriteId":receipt["rewriteId"],"view":view()});
    assert!(crate::calculation_transfer::command(&target, &transfer).is_err());
    transfer["purpose"] = json!("hyperlink-authoring");
    let accepted = crate::calculation_transfer::command(&target, &transfer).unwrap();
    assert_eq!(
        accepted["schema"],
        workbook_format!(SCHEMA_HYPERLINK_CALCULATION_TRANSFER_V1)
    );
    assert_eq!(accepted["transferredFormulaCells"], 1);
    assert_eq!(accepted["policy"]["pendingFormulaCells"], 0);
    let db = crate::database(&target).unwrap();
    let read = crate::range::read_public_db(
        &db,
        &json!({"version":accepted["version"],"sheet":PART,"range":"B1","limit":1}),
    )
    .unwrap();
    assert_eq!(read["cells"][0][1]["calculation"]["value"], 3.0);
}
fn rewrite_member(
    root: &Path,
    source: &str,
    target: &str,
    member: &str,
    replacement: Option<&str>,
) {
    let mut before = ZipArchive::new(crate::cancel::Reader(
        File::open(root.join(source)).unwrap(),
    ))
    .unwrap();
    let mut after = ZipWriter::new(File::create(root.join(target)).unwrap());
    for i in 0..before.len() {
        let file = before.by_index(i).unwrap();
        if file.name() != member {
            after.raw_copy_file(file).unwrap();
        } else if let Some(text) = replacement {
            after
                .start_file(member, SimpleFileOptions::default())
                .unwrap();
            after.write_all(text.as_bytes()).unwrap();
        }
    }
    after.finish().unwrap();
}
#[test]
fn first_external_link_creates_only_missing_sheet_relationship_part() {
    let _scope = crate::shared::begin();
    let (owner, old) = fixture("", "");
    let root = owner.path().join("no-rel");
    fs::create_dir(&root).unwrap();
    rewrite_member(&old, "source.xlsx", "without-rels.xlsx", RELPART, None);
    fs::copy(old.join("without-rels.xlsx"), root.join("source.xlsx")).unwrap();
    crate::import::index(&root).unwrap();
    fs::copy(root.join("source.xlsx"), root.join("export.xlsx")).unwrap();
    let target = owner.path().join("candidate");
    let input = sealed(&root, &target, 0);
    let command = web(None);
    super::super::apply(&root, &request(&root, &input, command.clone())).unwrap();
    verify_receipt_package(&root, &command, "package.xlsx").unwrap();
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(root.join("package.xlsx")).unwrap(),
    ))
    .unwrap();
    let rels = Tree::parse(&part(&mut zip, RELPART).unwrap()).unwrap();
    rels.root(PACKAGE_REL_NS, "Relationships").unwrap();
    assert_eq!(rels.nodes[0].children.len(), 1);
    assert_eq!(rels.nodes[1].attrs["Id"], workbook_format!(HYPERLINK1));
}
#[test]
fn arbitrary_metadata_plans_unsealed_commands_and_tampered_outputs_do_not_get_authority() {
    let _scope = crate::shared::begin();
    let (owner, root) = fixture("", "");
    let command = web(None);
    let forged = json!({"action":"workbook","operation":"package-edit","root":root,"version":0,"source":"export.xlsx","command":command});
    assert!(super::super::apply(&root, &forged).is_err());
    assert!(!root.join("package.xlsx").exists());
    fs::write(
        root.join("package-metadata.json"),
        json!({"kind":"setHyperlink","replacements":{},"fragments":{},"copies":{},"removed":[]})
            .to_string(),
    )
    .unwrap();
    assert!(super::super::apply(&root, &json!({"version":0,"source":"export.xlsx"})).is_err());
    let target = owner.path().join("candidate");
    let input = sealed(&root, &target, 0);
    super::super::apply(&root, &request(&root, &input, command.clone())).unwrap();
    verify_receipt_package(&root, &command, "package.xlsx").unwrap();
    let sheet = output(&root, PART).replace("<v>2</v>", "<v>999</v>");
    rewrite_member(
        &root,
        "package.xlsx",
        "tampered-cell.xlsx",
        PART,
        Some(&sheet),
    );
    assert!(verify_receipt_package(&root, &command, "tampered-cell.xlsx").is_err());
    rewrite_member(
        &root,
        "package.xlsx",
        "tampered-opaque.xlsx",
        "custom/opaque.bin",
        Some("changed unrelated payload"),
    );
    assert!(verify_receipt_package(&root, &command, "tampered-opaque.xlsx").is_err());
    let rels = output(&root, RELPART).replace("example.test", "changed.test");
    rewrite_member(
        &root,
        "package.xlsx",
        "tampered-target.xlsx",
        RELPART,
        Some(&rels),
    );
    assert!(verify_receipt_package(&root, &command, "tampered-target.xlsx").is_err());
}
fn inspect(root: &Path, reference: &str) -> Value {
    let index: Value = serde_json::from_slice(&fs::read(root.join("index.json")).unwrap()).unwrap();
    read(root,&json!({"action":"workbook","operation":"hyperlink","root":root,"version":0,"sourceRevision":index["revision"],"sheet":"Main","cell":reference})).unwrap()
}
#[test]
fn inspection_is_bound_readonly_and_admits_only_the_editable_profile() {
    let _scope = crate::shared::begin();
    let (_owner, root) = fixture(
        r#"<hyperlinks><hyperlink ref="A1" q:id="old" tooltip="Visit &amp; inspect"/><hyperlink ref="C1" location="'Main'!$Z$80"/></hyperlinks>"#,
        &edge("old"),
    );
    let source = crate::hash(&root.join("source.xlsx")).unwrap();
    let a = inspect(&root, "A1");
    assert_eq!(a["link"]["kind"], "url");
    assert_eq!(a["hyperlinkRevision"], revision(&root, "A1"));
    assert_eq!(a["tooltip"], "Visit & inspect");
    assert_eq!(a["canEdit"], true);
    let c = inspect(&root, "C1");
    assert_eq!(
        c["link"],
        json!({"kind":"internal","sheet":"Main","cell":"Z80"})
    );
    let empty = inspect(&root, "B1");
    assert_eq!(empty["link"], Value::Null);
    assert_eq!(empty["hyperlinkRevision"], Value::Null);
    assert_eq!(empty["canEdit"], true);
    assert_eq!(crate::hash(&root.join("source.xlsx")).unwrap(), source);
    assert!(!root.join("package.xlsx").exists());
}
#[test]
fn inspection_explains_range_named_package_and_credentialed_destinations() {
    let _scope = crate::shared::begin();
    for (links, edges) in [
        (
            r#"<hyperlinks><hyperlink ref="A1:B2" location="A1"/></hyperlinks>"#.to_string(),
            String::new(),
        ),
        (
            r#"<hyperlinks><hyperlink ref="A1" location="DefinedTarget"/></hyperlinks>"#
                .to_string(),
            String::new(),
        ),
        (
            r#"<hyperlinks><hyperlink ref="A1" q:id="old"/></hyperlinks>"#.to_string(),
            edge("old").replace(
                "https://example.test/a?x=1&amp;y=2#fragment",
                "https://user:password@example.test/",
            ),
        ),
    ] {
        let (_owner, root) = fixture(&links, &edges);
        let value = inspect(&root, "A1");
        assert_eq!(value["canEdit"], false);
        assert_eq!(value["link"], Value::Null);
        assert!(
            value["protectedReason"]
                .as_str()
                .is_some_and(|s| !s.is_empty())
        );
        let command = web(value["hyperlinkRevision"].as_str().map(str::to_owned));
        assert!(prepare_from(&root, "export.xlsx", &command).is_err());
    }
}
#[test]
fn inspection_rejects_foreign_source_version_and_forged_request_fields() {
    let _scope = crate::shared::begin();
    let (_owner, root) = fixture("", "");
    let index: Value = serde_json::from_slice(&fs::read(root.join("index.json")).unwrap()).unwrap();
    let good = json!({"action":"workbook","operation":"hyperlink","root":root,"version":0,"sourceRevision":index["revision"],"sheet":"Main","cell":"A1"});
    for patch in [
        json!({"sourceRevision":"a".repeat(64)}),
        json!({"version":999}),
        json!({"sheet":"Missing"}),
        json!({"cell":"A1:B2"}),
        json!({"privatePath":"source.xlsx"}),
    ] {
        let mut input = good.clone();
        for (k, v) in patch.as_object().unwrap() {
            input[k] = v.clone();
        }
        assert!(read(&root, &input).is_err());
    }
}
