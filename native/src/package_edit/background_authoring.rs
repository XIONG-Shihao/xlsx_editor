//! Two real worksheet background samples, owned picture/media relationships.
//! Cells and drawing objects remain streamed and unchanged. No arbitrary PNG.
use super::{clear_hyperlinks, print_settings::Tree, theme_authoring, *};
const IMAGE_EDGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
fn sample(id: &str) -> Result<&'static [u8]> {
    match id {
        "graph-paper" => Ok(include_bytes!(
            "../../assets/page-backgrounds/graph-paper.png"
        )),
        "soft-dots" => Ok(include_bytes!(
            "../../assets/page-backgrounds/soft-dots.png"
        )),
        _ => Err(crate::coded(
            "internal-unknown-reference",
            "Unknown worksheet background sample",
        )),
    }
}
fn digest(bytes: &[u8]) -> String {
    format!("{:x}", Sha256::digest(bytes))
}
fn media_part(sheet: &str, id: &str) -> String {

    format!(workbook_format!(XL_MEDIA_BACKGROUND_ID_PNG),
        &digest(sheet.as_bytes())[..16],
        id = id
    )
}
fn relation_part(sheet: &str) -> Result<String> {
    let (dir, file) = sheet
        .rsplit_once('/')
        .ok_or_else(|| crate::coded("object-malformed", "Invalid background worksheet part"))?;
    Ok(format!("{dir}/_rels/{file}.rels"))
}
fn part(zip: &mut ZipArchive<crate::cancel::Reader<File>>, name: &str) -> Result<Vec<u8>> {
    part_bytes(zip, name, &serde_json::Map::new())
}
fn png(zip: &mut ZipArchive<crate::cancel::Reader<File>>, name: &str) -> Result<Vec<u8>> {
    let mut member = zip.by_name(name).map_err(err)?;
    if member.size() > PNG_LIMIT as u64 {
        return Err(crate::coded(
            "object-limit",
            "Background PNG exceeds its bound",
        ));
    }
    let mut result = Vec::new();
    member.read_to_end(&mut result).map_err(err)?;
    Ok(result)
}
fn relation_tree(
    zip: &mut ZipArchive<crate::cancel::Reader<File>>,
    part_name: &str,
) -> Result<Tree> {
    if zip.file_names().any(|name| name == part_name) {
        Tree::parse(&part(zip, part_name)?)
    } else {
        Tree::parse(
            format!(
                "<Relationships xmlns=\"{}\"/>",
                String::from_utf8_lossy(PACKAGE_REL_NS)
            )
            .as_bytes(),
        )
    }
}
fn media_incoming(zip: &mut ZipArchive<crate::cancel::Reader<File>>, media: &str) -> Result<usize> {
    let names = clear_hyperlinks::package_names(zip)?;
    let mut count = 0;
    for name in names.iter().filter(|name| name.ends_with(".rels")) {
        let tree = Tree::parse(&part(zip, name)?)?;
        let edges = theme_authoring::relationships(&tree)?;
        let owner = if name == "_rels/.rels" {
            String::new()
        } else {
            let (dir, file) = name.rsplit_once("/_rels/").ok_or_else(|| {
                crate::coded("object-malformed", "Invalid background relationship part")
            })?;
            format!(
                "{dir}/{}",
                file.strip_suffix(".rels").ok_or_else(|| crate::coded(
                    "object-malformed",
                    "Invalid relationship suffix"
                ))?
            )
        };
        for index in edges {
            let node = &tree.nodes[index];
            if node
                .attrs
                .get("TargetMode")
                .is_some_and(|mode| mode == "External")
            {
                continue;
            }
            if theme_authoring::target(&owner, &node.attrs["Target"])? == media {
                count += 1;
            }
        }
    }
    Ok(count)
}
struct Existing {
    revision: String,
    sample: String,
    media: String,
    edge: usize,
    id: String,
}
fn existing(
    root: &Path,
    zip: &mut ZipArchive<crate::cancel::Reader<File>>,
    sheet: &str,
    tree: &Tree,
    rels: &Tree,
) -> Result<Option<Existing>> {
    let Some(picture) = tree.direct(0, "picture")? else {
        return Ok(None);
    };
    tree.clean(picture, &["r:id"])?;
    let id = tree.nodes[picture].attrs.get("r:id").ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing background relationship id",
        )
    })?;
    let edges = theme_authoring::relationships(rels)?;
    let edge = *edges
        .iter()
        .find(|index| rels.nodes[**index].attrs.get("Id") == Some(id))
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing background image edge"))?;
    let node = &rels.nodes[edge];
    if node.attrs["Type"] != IMAGE_EDGE
        || node
            .attrs
            .get("TargetMode")
            .is_some_and(|mode| mode != "Internal")
    {
        return Err(crate::coded(
            "object-protected",
            "External/foreign worksheet backgrounds stay protected",
        ));
    }
    let media = theme_authoring::target(sheet, &node.attrs["Target"])?;
    let id_sample = ["graph-paper", "soft-dots"]
        .iter()
        .find(|sample| {
            media_part(sheet, sample) == media
        })
        .ok_or_else(|| {
            crate::coded(
                "object-protected",
                "Imported worksheet backgrounds stay protected",
            )
        })?;
    let bytes = png(zip, &media)?;
    if bytes != sample(id_sample)? {
        return Err(crate::coded(
            "object-protected",
            "Modified background assets stay protected",
        ));
    }
    if media_incoming(zip, &media)? != 1 {
        return Err(crate::coded(
            "object-protected",
            "Shared background assets stay protected",
        ));
    }
    let candidates = BTreeSet::from([id.clone()]);
    let (_, _, references) =
        clear_hyperlinks::worksheet_tag_hash(zip, sheet, Some(&candidates), b"picture")?;
    if !references.is_empty() {
        return Err(crate::coded(
            "object-refused",
            "Background relationship has another worksheet consumer",
        ));
    }
    let revision = digest(format!("{sheet}\0{id}\0{media}\0{}", digest(&bytes)).as_bytes());
    let _ = root;
    Ok(Some(Existing {
        revision,
        sample: (*id_sample).into(),
        media,
        edge,
        id: id.clone(),
    }))
}
fn definition(root: &Path, sheet: &str) -> Result<Option<Value>> {
    let mut metadata = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let tree = Tree::parse(&part(&mut metadata, sheet)?)?;
    tree.root(NS, "worksheet")?;
    if tree.direct(0, "picture")?.is_none() {
        return Ok(None);
    }
    let rel_part = relation_part(sheet)?;
    let rels = relation_tree(&mut metadata, &rel_part)?;
    let mut source = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "source.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let value = existing(root, &mut source, sheet, &tree, &rels)?
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing background"))?;
    Ok(Some(
        json!({"schema":workbook_format!(SCHEMA_SHEET_BACKGROUND_V1),"sheet":sheet,"revision":value.revision,"sample":value.sample,"canEdit":true,"protectedReason":null}),
    ))
}
/// Import projects only the independently verified sample identity, never PNG
/// bytes or renderer-provided paths. Foreign sources stay present and protected.
pub(super) fn inspect(root: &Path, sheet: &str) -> Value {
    match definition(root, sheet) {
        Ok(Some(value)) => value,
        Ok(None) => Value::Null,
        Err(reason) => {
            json!({"schema":workbook_format!(SCHEMA_SHEET_BACKGROUND_V1),"sheet":sheet,"revision":null,"sample":null,"canEdit":false,"protectedReason":reason})
        }
    }
}
pub(super) struct Prepared {
    pub plan: Value,
    pub command: Value,
    pub additions: Vec<(String, Vec<u8>)>,
    pub changed: bool,
    pub prior: Option<String>,
    pub sample: Option<String>,
    pub sheet: String,
}
impl Prepared {
    pub fn result(&self) -> Value {
        json!({"schema":workbook_format!(SCHEMA_BACKGROUND_WRITER_V1),"file":"package.xlsx","contentChanged":false,"changed":self.changed,"priorBackgroundRevision":self.prior,"sample":self.sample,"sheet":self.sheet})
    }
}
fn prepare_from(root: &Path, source: &str, command: &Value) -> Result<Prepared> {
    let remove = command["type"] == "removeSheetBackground";
    theme_authoring::fields(
        command,
        &["type", "sheet", "expectedBackgroundRevision"],
        if remove { &[] } else { &["sample"] },
    )?;
    if !remove && (command["type"] != "setSheetBackground" || command.get("sample").is_none()) {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid background command",
        ));
    }
    let selected_sample = if remove {
        None
    } else {
        let id = field(command, "sample")?;
        sample(id)?;
        Some(id.to_owned())
    };
    let expected = &command["expectedBackgroundRevision"];
    if !expected.is_null()
        && !expected.as_str().is_some_and(|s| {
            s.len() == 64
                && s.bytes()
                    .all(|b| b.is_ascii_digit() || matches!(b, b'a'..=b'f'))
        })
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid captured background revision",
        ));
    }
    let inventory = crate::used_range_transfer::worksheet_inventory(&local(root, source)?)?;
    let sheet = inventory
        .iter()
        .find(|sheet| sheet.name == field(command, "sheet").unwrap_or(""))
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing background worksheet"))?
        .path
        .clone();
    let mut metadata = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let tree = Tree::parse(&part(&mut metadata, &sheet)?)?;
    tree.root(NS, "worksheet")?;
    if tree
        .nodes
        .iter()
        .any(|node| node.ns == NS && node.name == "sheetProtection")
    {
        return Err(crate::coded(
            "object-protected",
            "Protected worksheet background stays read-only",
        ));
    }
    let rel_part = relation_part(&sheet)?;
    let rels = relation_tree(&mut metadata, &rel_part)?;
    let edges = theme_authoring::relationships(&rels)?;
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, source)?).map_err(err)?,
    ))
    .map_err(err)?;
    let names = clear_hyperlinks::package_names(&zip)?;
    let workbook = Tree::parse(&part(&mut zip, "xl/workbook.xml")?)?;
    workbook.root(NS, "workbook")?;
    if workbook
        .nodes
        .iter()
        .any(|node| node.ns == NS && node.name == "workbookProtection")
        || names.iter().any(|name| {
            name.starts_with("_xmlsignatures/")
                || name.to_ascii_lowercase().ends_with("vbaproject.bin")
        })
    {
        return Err(crate::coded(
            "object-protected",
            "Protected/macro/signed background source stays read-only",
        ));
    }
    let current = existing(root, &mut zip, &sheet, &tree, &rels)?;
    let prior = current.as_ref().map(|value| value.revision.clone());
    if expected != &json!(prior) {
        return Err(crate::coded(
            "object-stale",
            "Background changed since the captured revision",
        ));
    }
    let same = remove && current.is_none()
        || current
            .as_ref()
            .is_some_and(|current| Some(&current.sample) == selected_sample.as_ref());
    if same {
        return Ok(Prepared {
            plan: json!({"kind":command["type"],"copies":{},"removed":[],"replacements":{},"fragments":{}}),
            command: command.clone(),
            additions: vec![],
            changed: false,
            prior,
            sample: selected_sample,
            sheet,
        });
    }
    let mut replacements = serde_json::Map::new();
    let mut removed = Vec::new();
    let mut additions = Vec::new();
    let mut revised_rels = if let Some(current) = &current {
        String::from_utf8(rels.mask(&[current.edge])?).map_err(err)?
    } else {
        String::from_utf8(rels.bytes.clone()).map_err(err)?
    };
    let mut types = Tree::parse(&part(&mut zip, "[Content_Types].xml")?)?;
    types.root(TYPE_NS, "Types")?;
    if let Some(current) = &current {
        removed.push(current.media.clone());
        let matches = types.nodes[0]
            .children
            .iter()
            .filter(|index| {
                types.nodes[**index].attrs.get("PartName") == Some(&format!("/{}", current.media))
            })
            .copied()
            .collect::<Vec<_>>();
        if matches.len() > 1 {
            return Err(crate::coded(
                "object-ambiguous",
                "Ambiguous background content type",
            ));
        }
        if let Some(index) = matches.first() {
            types.clean(*index, &["PartName", "ContentType"])?;
            if types.nodes[*index].ns != TYPE_NS
                || types.nodes[*index].name != "Override"
                || types.nodes[*index]
                    .attrs
                    .get("ContentType")
                    .map(String::as_str)
                    != Some("image/png")
            {
                return Err(crate::coded(
                    "object-unsupported",
                    "Foreign background content type",
                ));
            }
            types = Tree::parse(&types.mask(&[*index])?)?;
        }
    }
    let picture = if let Some(id) = &selected_sample {
        let media = media_part(&sheet, id);
        if names.contains(&media) {
            return Err(crate::coded(
                "object-conflict",
                "Background media part name is occupied",
            ));
        }
        let mut suffix = 1;
        let relation_id = if let Some(current) = &current {
            current.id.clone()
        } else {
            while edges.iter().any(|index| {
                rels.nodes[*index].attrs["Id"]
                    == format!(workbook_format!(BACKGROUND_SUFFIX), suffix = suffix)
            }) {
                suffix += 1;
            }
            format!(workbook_format!(BACKGROUND_SUFFIX), suffix = suffix)
        };
        revised_rels = theme_authoring::append(
            &Tree::parse(revised_rels.as_bytes())?,
            &format!(
                "<Relationship xmlns=\"{}\" Id=\"{relation_id}\" Type=\"{IMAGE_EDGE}\" Target=\"/{media}\"/>",
                String::from_utf8_lossy(PACKAGE_REL_NS)
            ),
        )?;
        if types.nodes[0]
            .children
            .iter()
            .any(|index| types.nodes[*index].attrs.get("PartName") == Some(&format!("/{media}")))
        {
            return Err(crate::coded(
                "object-conflict",
                "Background content type is occupied",
            ));
        }
        let revised = theme_authoring::append(
            &types,
            &format!(
                "<Override xmlns=\"{}\" PartName=\"/{media}\" ContentType=\"image/png\"/>",
                String::from_utf8_lossy(TYPE_NS)
            ),
        )?;
        types = Tree::parse(revised.as_bytes())?;
        additions.push((media, sample(id)?.to_vec()));
        format!(
            "<picture xmlns=\"{}\" xmlns:r=\"{}\" r:id=\"{relation_id}\"/>",
            String::from_utf8_lossy(NS),
            String::from_utf8_lossy(REL_NS)
        )
    } else {
        String::new()
    };
    replacements.insert(rel_part, json!(revised_rels));
    replacements.insert(
        "[Content_Types].xml".into(),
        json!(String::from_utf8(types.bytes).map_err(err)?),
    );
    let plan = json!({"kind":command["type"],"copies":{},"removed":removed,"replacements":replacements,"fragments":{sheet.clone():{"picture":picture}}});
    Ok(Prepared {
        plan,
        command: command.clone(),
        additions,
        changed: true,
        prior,
        sample: selected_sample,
        sheet,
    })
}
pub(super) fn prepare(root: &Path, input: &Value) -> Result<Prepared> {
    theme_authoring::fields(
        input,
        &[
            "action",
            "operation",
            "root",
            "version",
            "source",
            "rewriteInputId",
            "command",
        ],
        &["diagnostics"],
    )?;
    if input["action"] != "workbook"
        || input["operation"] != "package-edit"
        || input["source"] != "export.xlsx"
        || input.get("diagnostics").is_some_and(|v| !v.is_boolean())
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid native background request",
        ));
    }
    prepare_from(root, "export.xlsx", &input["command"])
}
pub(super) fn verify_output(
    root: &Path,
    source: &str,
    output: &str,
    prepared: &Prepared,
) -> Result<()> {
    let mut before = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, source)?).map_err(err)?,
    ))
    .map_err(err)?;
    let mut after = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, output)?).map_err(err)?,
    ))
    .map_err(err)?;
    let mut names = clear_hyperlinks::package_names(&before)?;
    let replacements = prepared.plan["replacements"]
        .as_object()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing background replacements"))?;
    for removed in prepared.plan["removed"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing background removals"))?
    {
        names.remove(
            removed.as_str().ok_or_else(|| {
                crate::coded("internal-invalid-field", "Invalid background removal")
            })?,
        );
    }
    names.extend(replacements.keys().cloned());
    names.extend(prepared.additions.iter().map(|(name, _)| name.clone()));
    if names != clear_hyperlinks::package_names(&after)? {
        return Err(crate::coded(
            "object-stale",
            "Background writer changed unrelated package inventory",
        ));
    }
    for name in names {
        if let Some(fragment) = prepared.plan["fragments"].get(&name) {
            let (a, _, _) =
                clear_hyperlinks::worksheet_tag_hash(&mut before, &name, None, b"picture")?;
            let (b, picture, _) =
                clear_hyperlinks::worksheet_tag_hash(&mut after, &name, None, b"picture")?;
            if a != b || picture != field(fragment, "picture")?.as_bytes() {
                return Err(crate::coded(
                    "object-stale",
                    "Background writer changed nonbackground worksheet content",
                ));
            }
        } else if let Some(xml) = replacements.get(&name) {
            if clear_hyperlinks::member_hash(&mut after, &name)?
                != Sha256::digest(
                    xml.as_str()
                        .ok_or_else(|| crate::coded("object-malformed", "Invalid background XML"))?
                        .as_bytes(),
                )
                .to_vec()
            {
                return Err(crate::coded(
                    "internal-mismatch",
                    "Background output differs from native relationships/types",
                ));
            }
        } else if let Some((_, bytes)) = prepared.additions.iter().find(|(part, _)| part == &name) {
            if png(&mut after, &name)? != *bytes {
                return Err(crate::coded(
                    "internal-mismatch",
                    "Background output differs from native PNG sample",
                ));
            }
        } else if clear_hyperlinks::member_hash(&mut before, &name)?
            != clear_hyperlinks::member_hash(&mut after, &name)?
        {
            return Err(crate::coded(
                "object-stale",
                "Background writer changed unrelated package content",
            ));
        }
    }
    Ok(())
}
pub(super) fn verify_receipt_package(root: &Path, command: &Value, output: &str) -> Result<()> {
    let prepared = prepare_from(root, "export.xlsx", command)?;
    verify_output(root, "export.xlsx", output, &prepared)
}
#[cfg(test)]
mod tests;
