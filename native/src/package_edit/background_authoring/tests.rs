use super::super::theme_authoring::tests::{fixture, mutate, read_part, sealed};
use super::*;
const SHEET: &str = "xl/worksheets/main.xml";
fn command(id: Option<&str>, prior: Option<String>) -> Value {
    match id {
        Some(id) => {
            json!({"type":"setSheetBackground","sheet":"Main","sample":id,"expectedBackgroundRevision":prior})
        }
        None => {
            json!({"type":"removeSheetBackground","sheet":"Main","expectedBackgroundRevision":prior})
        }
    }
}
fn request(root: &Path, token: &Value, cmd: Value) -> Value {
    json!({"action":"workbook","operation":"package-edit","root":root,"source":"export.xlsx","version":0,"rewriteInputId":token["rewriteInputId"],"command":cmd})
}
#[test]
fn two_samples_are_saved_as_picture_relationships_with_real_pngs_and_unchanged_cells() {
    let _scope = crate::shared::begin();
    for id in ["graph-paper", "soft-dots"] {
        let (owner, root) = fixture();
        let source_revision = crate::hash(&root.join("source.xlsx")).unwrap();
        let token = sealed(&root, &owner.path().join("candidate"));
        let cmd = command(Some(id), None);
        let result = super::super::apply(&root, &request(&root, &token, cmd.clone())).unwrap();
        assert_eq!(
            result["schema"],
            workbook_format!(SCHEMA_BACKGROUND_WRITER_V1)
        );
        assert_eq!(result["sample"], id);
        assert_eq!(result["contentChanged"], false);
        assert_eq!(result["changed"], true);
        verify_receipt_package(&root, &cmd, "package.xlsx").unwrap();
        assert_eq!(
            read_part(&root, "package.xlsx", &media_part(SHEET, id)),
            sample(id).unwrap()
        );
        let output = String::from_utf8(read_part(&root, "package.xlsx", SHEET)).unwrap();
        assert!(output.contains("<picture"));
        assert!(!output.contains("<drawing"));
        assert!(output.contains("<f>SUM(A1,1)</f>"));
        assert_eq!(
            read_part(&root, "package.xlsx", "xl/styles.xml"),
            read_part(&root, "export.xlsx", "xl/styles.xml")
        );
        assert_eq!(
            read_part(&root, "package.xlsx", "custom/opaque.bin"),
            read_part(&root, "source.xlsx", "custom/opaque.bin")
        );
        assert_eq!(
            crate::hash(&root.join("source.xlsx")).unwrap(),
            source_revision
        );
    }
}
#[test]
fn imported_index_verifies_owned_background_then_switch_remove_and_noop_are_closed() {
    let _scope = crate::shared::begin();
    let (owner, root) = fixture();
    let target = owner.path().join("candidate");
    let token = sealed(&root, &target);
    super::super::apply(
        &root,
        &request(&root, &token, command(Some("graph-paper"), None)),
    )
    .unwrap();
    fs::copy(root.join("package.xlsx"), target.join("source.xlsx")).unwrap();
    let index = crate::import::index(&target).unwrap();
    let value = &index["sheets"][0]["background"];
    assert_eq!(value["sample"], "graph-paper");
    assert_eq!(value["canEdit"], true);
    let prior = value["revision"].as_str().unwrap().to_owned();
    fs::copy(target.join("source.xlsx"), target.join("export.xlsx")).unwrap();
    assert!(
        !prepare_from(
            &target,
            "export.xlsx",
            &command(Some("graph-paper"), Some(prior.clone()))
        )
        .unwrap()
        .changed
    );
    let switched = prepare_from(
        &target,
        "export.xlsx",
        &command(Some("soft-dots"), Some(prior.clone())),
    )
    .unwrap();
    assert_eq!(switched.additions[0].1, sample("soft-dots").unwrap());
    assert!(
        switched.plan["removed"]
            .as_array()
            .unwrap()
            .iter()
            .any(|v| v == &media_part(SHEET, "graph-paper"))
    );
    let removed = prepare_from(&target, "export.xlsx", &command(None, Some(prior))).unwrap();
    assert_eq!(removed.plan["fragments"][SHEET]["picture"], "");
    assert!(removed.additions.is_empty());
    assert!(prepare_from(&target, "export.xlsx", &command(None, None)).is_err());
    assert!(
        !prepare_from(&root, "export.xlsx", &command(None, None))
            .unwrap()
            .changed
    );
}
#[test]
fn native_output_proof_rejects_cell_image_style_and_unrelated_mutants() {
    let _scope = crate::shared::begin();
    let (owner, root) = fixture();
    let token = sealed(&root, &owner.path().join("candidate"));
    let cmd = command(Some("graph-paper"), None);
    super::super::apply(&root, &request(&root, &token, cmd.clone())).unwrap();
    let clean = fs::read(root.join("package.xlsx")).unwrap();
    for (name, bytes) in [
        (
            SHEET,
            String::from_utf8(read_part(&root, "package.xlsx", SHEET))
                .unwrap()
                .replace("<v>2</v>", "<v>900</v>")
                .into_bytes(),
        ),
        (
            media_part(SHEET, "graph-paper").as_str(),
            sample("soft-dots").unwrap().to_vec(),
        ),
        ("xl/styles.xml", b"changed styles".to_vec()),
        ("custom/opaque.bin", b"changed payload".to_vec()),
    ] {
        fs::write(root.join("package.xlsx"), &clean).unwrap();
        mutate(&root, "package.xlsx", name, &bytes);
        assert!(verify_receipt_package(&root, &cmd, "package.xlsx").is_err());
    }
}
#[test]
fn source_foreign_shared_and_modified_assets_refuse_and_do_not_authorize_read_projection() {
    let _scope = crate::shared::begin();
    let (owner, root) = fixture();
    let target = owner.path().join("candidate");
    let token = sealed(&root, &target);
    super::super::apply(
        &root,
        &request(&root, &token, command(Some("graph-paper"), None)),
    )
    .unwrap();
    fs::copy(root.join("package.xlsx"), target.join("source.xlsx")).unwrap();
    crate::import::index(&target).unwrap();
    fs::copy(target.join("source.xlsx"), target.join("export.xlsx")).unwrap();
    let original = inspect(&target, SHEET);
    let prior = original["revision"].as_str().unwrap().to_owned();
    let original_source = fs::read(target.join("source.xlsx")).unwrap();
    let original_export = fs::read(target.join("export.xlsx")).unwrap();
    let shared = format!(
        "<Relationships xmlns=\"{}\"><Relationship Id=\"other\" Type=\"{IMAGE_EDGE}\" Target=\"/{}\"/></Relationships>",
        String::from_utf8_lossy(PACKAGE_REL_NS),
        media_part(SHEET, "graph-paper")
    );
    mutate(
        &target,
        "source.xlsx",
        "custom/_rels/other.xml.rels",
        shared.as_bytes(),
    );
    mutate(
        &target,
        "export.xlsx",
        "custom/_rels/other.xml.rels",
        shared.as_bytes(),
    );
    assert_eq!(inspect(&target, SHEET)["canEdit"], false);
    assert!(prepare_from(&target, "export.xlsx", &command(None, Some(prior.clone()))).is_err());
    fs::write(target.join("source.xlsx"), original_source).unwrap();
    fs::write(target.join("export.xlsx"), original_export).unwrap();
    mutate(
        &target,
        "source.xlsx",
        &media_part(SHEET, "graph-paper"),
        sample("soft-dots").unwrap(),
    );
    assert_eq!(inspect(&target, SHEET)["canEdit"], false);
    mutate(
        &target,
        "export.xlsx",
        &media_part(SHEET, "graph-paper"),
        sample("soft-dots").unwrap(),
    );
    assert!(prepare_from(&target, "export.xlsx", &command(None, Some(prior))).is_err());
    for cmd in [
        json!({"type":"setSheetBackground","sheet":"Main","sample":"foreign","expectedBackgroundRevision":null}),
        json!({"type":"setSheetBackground","sheet":"Main","sample":"graph-paper","expectedBackgroundRevision":null,"bytes":"caller"}),
        json!({"type":"setSheetBackground","sheet":"Missing","sample":"graph-paper","expectedBackgroundRevision":null}),
    ] {
        assert!(prepare_from(&root, "export.xlsx", &cmd).is_err());
    }
}
