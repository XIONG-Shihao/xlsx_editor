//! Produce cache-publication metadata from the selected native version. The
//! Calculation flags remain native-owned. A bounded style planner may supply
//! the resolved style table; mappings are checked against this exact plan.
use super::{print_settings::Tree, *};
use crate::database;
fn escape(value: &str) -> String {
    quick_xml::escape::escape(value)
        .replace('\r', "&#13;")
        .replace('\n', "&#10;")
        .replace('\t', "&#9;")
}
pub(crate) fn prepare(root: &Path, input: &Value) -> Result<Value> {
    let requested_version = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing export metadata version"))?;
    let db = database(root)?;
    let (version, incomplete, spill_view): (u64, bool, Option<String>) = db
        .query_row(
            "SELECT version,incomplete,spill_view FROM export_state",
            [],
            |row| Ok((row.get(0)?, row.get(1)?, row.get(2)?)),
        )
        .map_err(err)?;
    if version != requested_version {
        return Err(crate::coded(
            "package-stale",
            "Stale export metadata preparation",
        ));
    }
    let style_count: u64 = db
        .query_row("SELECT count(*) FROM export_styles", [], |row| row.get(0))
        .map_err(err)?;
    if style_count != 0 && input.get("styles").is_none() {
        return Err(crate::coded(
            "package-bad-input",
            "Authored styles require the style publication service",
        ));
    }
    if style_count == 0 && input.get("styles").is_some() {
        return Err(crate::coded(
            "package-bad-input",
            "Style resolution requires authored style signatures",
        ));
    }
    let source = source_calculation_policy(root)?;
    if source["unsupportedMode"] == true {
        return Err(crate::coded(
            "package-protected",
            "Workbook calculation source settings are protected",
        ));
    }
    let (mode, on_save) = crate::calculation_policy::options(&db, version)?.unwrap_or((
        field(&source, "mode")?.into(),
        source["recalculateBeforeSave"].as_bool().ok_or_else(|| {
            crate::coded(
                "internal-invalid-field",
                "Invalid source calculation policy",
            )
        })?,
    ));
    let refresh = incomplete
        || crate::calculation_policy::pending_count(&db, version)? > 0
        || crate::calculation_policy::derived_count(&db, version)? > 0;
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let bytes = part_bytes(&mut zip, "xl/workbook.xml", &serde_json::Map::new())?;
    let tree = Tree::parse(&bytes)?;
    tree.root(NS, "workbook")?;
    let previous = tree.direct(0, "calcPr")?;
    let defaults = previous.is_none() && !refresh && mode == "auto" && on_save;
    let mut attributes = previous
        .map(|index| tree.nodes[index].attrs.clone())
        .unwrap_or_default();
    let qname = previous
        .map(|index| tree.nodes[index].qname.clone())
        .unwrap_or_else(|| "calcPr".into());
    if previous.is_none() {
        attributes.insert("xmlns".into(), String::from_utf8(NS.to_vec()).map_err(err)?);
    }
    attributes.insert("calcMode".into(), mode);
    attributes.insert("calcOnSave".into(), if on_save { "1" } else { "0" }.into());
    for key in ["fullCalcOnLoad", "forceFullCalc"] {
        attributes.insert(key.into(), if refresh { "1" } else { "0" }.into());
    }
    let element = format!(
        "<{qname}{}/>",
        attributes
            .iter()
            .map(|(key, value)| format!(" {key}=\"{}\"", escape(value)))
            .collect::<String>()
    );
    let mut replacement = bytes;
    if let Some(index) = previous {
        replacement.splice(
            tree.nodes[index].start..tree.nodes[index].end,
            element.bytes(),
        );
    } else if !defaults {
        // calcPr precedes these later workbook children in the OOXML sequence.
        // Appending it after extLst produces invalid workbook order.
        let tail = [
            "oleSize",
            "customWorkbookViews",
            "pivotCaches",
            "smartTagPr",
            "smartTagTypes",
            "webPublishing",
            "fileRecoveryPr",
            "webPublishObjects",
            "extLst",
        ];
        let at = tree.nodes[0]
            .children
            .iter()
            .copied()
            .find(|index| {
                tree.nodes[*index].ns == NS && tail.contains(&tree.nodes[*index].name.as_str())
            })
            .map(|index| tree.nodes[index].start)
            .unwrap_or_else(|| {
                let closing = format!("</{}", tree.nodes[0].qname);
                std::str::from_utf8(&tree.bytes)
                    .unwrap()
                    .rfind(&closing)
                    .unwrap_or(tree.nodes[0].end)
            });
        replacement.splice(at..at, element.bytes());
    }
    let mut replacements = serde_json::Map::new();
    replacements.insert(
        "xl/workbook.xml".into(),
        json!(String::from_utf8(replacement).map_err(err)?),
    );
    let relationships = part_bytes(
        &mut zip,
        "xl/_rels/workbook.xml.rels",
        &serde_json::Map::new(),
    )?;
    let rels = Tree::parse(&relationships)?;
    rels.root(PACKAGE_REL_NS, "Relationships")?;
    let mut removed = std::collections::BTreeSet::new();
    let mut spans = Vec::new();
    for index in &rels.nodes[0].children {
        let node = &rels.nodes[*index];
        if node.ns==PACKAGE_REL_NS && node.name=="Relationship" && node.attrs.get("Type").is_some_and(|kind|kind=="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain") {
            if node.attrs.get("TargetMode").is_some_and(|mode|mode=="External"){return Err(crate::coded("package-protected", "External calculation-chain relationship is protected"))}
            removed.insert(resolve("xl/workbook.xml",node.attrs.get("Target")
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing calculation chain target"))?)?);
            spans.push(node.start..node.end);
        }
    }
    if !removed.is_empty() {
        let types_bytes = part_bytes(&mut zip, "[Content_Types].xml", &serde_json::Map::new())?;
        let types = Tree::parse(&types_bytes)?;
        types.root(TYPE_NS, "Types")?;
        let type_spans = types.nodes[0]
            .children
            .iter()
            .filter_map(|index| {
                let node = &types.nodes[*index];
                (node.ns == TYPE_NS
                    && node.name == "Override"
                    && node
                        .attrs
                        .get("PartName")
                        .is_some_and(|part| removed.contains(part.trim_start_matches('/'))))
                .then_some(node.start..node.end)
            })
            .collect();
        replacements.insert(
            "xl/_rels/workbook.xml.rels".into(),
            json!(remove_spans(relationships, spans)?),
        );
        replacements.insert(
            "[Content_Types].xml".into(),
            json!(remove_spans(types_bytes, type_spans)?),
        );
    }
    let mut metadata = json!({"styleIds":[],"replacements":replacements,"removed":removed});
    if let Some(styles) = input.get("styles") {
        let xml = styles["stylesXml"]
            .as_str()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing resolved styles XML"))?;
        let tree = Tree::parse(xml.as_bytes())?;
        tree.root(NS, "styleSheet")?;
        let xfs = tree.direct(0, "cellXfs")?.ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing resolved cell styles")
        })?;
        let count = tree.nodes[xfs]
            .children
            .iter()
            .filter(|index| tree.nodes[**index].ns == NS && tree.nodes[**index].name == "xf")
            .count() as u64;
        let ids = styles["styleIds"].as_array().ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing resolved style mappings")
        })?;
        if ids.len() as u64 != style_count {
            return Err(crate::coded(
                "package-unsupported",
                "Incomplete resolved style mappings",
            ));
        }
        let mut seen = std::collections::BTreeSet::new();
        for mapping in ids {
            let id = mapping["id"].as_u64().ok_or_else(|| {
                crate::coded("internal-invalid-field", "Invalid style mapping id")
            })?;
            let target = mapping["styleId"].as_u64().ok_or_else(|| {
                crate::coded("internal-invalid-field", "Invalid resolved style id")
            })?;
            if target >= count
                || !seen.insert(id)
                || !db
                    .query_row(
                        "SELECT EXISTS(SELECT 1 FROM export_styles WHERE id=?1)",
                        [id],
                        |row| row.get::<_, bool>(0),
                    )
                    .map_err(err)?
            {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid resolved style mapping",
                ));
            }
        }
        metadata["styleIds"] = Value::Array(ids.clone());
        metadata["replacements"]["xl/styles.xml"] = json!(xml);
    }
    validate_calculation_metadata(root, &db, version, &metadata)?;
    if crate::spill::any_history(&db, version)? {
        let view: Value = serde_json::from_str(spill_view.as_deref().ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing stored spill export view")
        })?)
        .map_err(err)?;
        crate::spill::validate_current(root, &db, &json!({"version":version,"view":view}))?;
        let groups = crate::spill::export_groups(&db, version, &view)?;
        let before = crate::spill::metadata::calculation_base(
            &crate::spill::metadata::source(root)?,
            &metadata["replacements"],
            &metadata["removed"],
        )?;
        // The stream exporter admits this plan separately and applies its
        // append delta after the calculation owner's calcChain normalization.
        metadata["spillMetadata"] = crate::spill::metadata::build(&before, &groups)?;
    }
    crate::quota::write(
        local(root, "export-metadata.json")?,
        serde_json::to_vec(&metadata).map_err(err)?,
    )
    .map_err(err)?;
    Ok(json!({"version":version,"refreshOnLoad":refresh,"removedParts":removed.len()}))
}
fn remove_spans(mut bytes: Vec<u8>, mut spans: Vec<std::ops::Range<usize>>) -> Result<String> {
    spans.sort_by_key(|span| span.start);
    for span in spans.into_iter().rev() {
        bytes.drain(span);
    }
    String::from_utf8(bytes).map_err(err)
}
