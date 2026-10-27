//! Closed Page Layout theme samples. Native-owned package and style changes;
//! caller XML, arbitrary part paths and imported/shared theme graphs are refused.
use super::{clear_hyperlinks, print_settings::Tree, *};
pub(super) const THEME_PART: &str = workbook_format!(XL_THEME_PAGE_THEME_XML);
const THEME_TYPE: &str = "application/vnd.openxmlformats-officedocument.theme+xml";
pub(super) const THEME_EDGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
pub(super) fn sample(id: &str) -> Result<&'static str> {
    match id {
        "neutral" => Ok(include_str!("../../assets/page-themes/neutral.xml")),
        "forest" => Ok(include_str!("../../assets/page-themes/forest.xml")),
        "plum" => Ok(include_str!("../../assets/page-themes/plum.xml")),
        _ => Err(crate::coded(
            "internal-unknown-reference",
            "Unknown page theme sample",
        )),
    }
}
pub(super) fn target(owner: &str, value: &str) -> Result<String> {
    if value.contains(['\\', ':', '?', '#', '%']) {
        return Err(crate::coded(
            "style-unsupported",
            "Unsupported theme relationship target",
        ));
    }
    let mut path = if value.starts_with('/') {
        Vec::new()
    } else {
        owner
            .rsplit_once('/')
            .map(|(dir, _)| dir.split('/').map(str::to_owned).collect())
            .unwrap_or_default()
    };
    for segment in value.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                path.pop().ok_or_else(|| {
                    crate::coded("style-unsupported", "Theme relationship escapes package")
                })?;
            }
            _ => path.push(segment.into()),
        }
    }
    let result = path.join("/");
    valid_part(&result)?;
    Ok(result)
}
pub(super) fn fields(value: &Value, required: &[&str], optional: &[&str]) -> Result<()> {
    let object = value
        .as_object()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid page theme record"))?;
    if required.iter().any(|key| !object.contains_key(*key))
        || object
            .keys()
            .any(|key| !required.contains(&key.as_str()) && !optional.contains(&key.as_str()))
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid page theme fields",
        ));
    }
    Ok(())
}
fn digest(bytes: &[u8]) -> String {
    format!("{:x}", Sha256::digest(bytes))
}
fn part(zip: &mut ZipArchive<crate::cancel::Reader<File>>, name: &str) -> Result<Vec<u8>> {
    part_bytes(zip, name, &serde_json::Map::new())
}
pub(super) fn append(tree: &Tree, xml: &str) -> Result<String> {
    let root = &tree.nodes[0];
    let raw = std::str::from_utf8(&tree.bytes).map_err(err)?;
    if root.children.is_empty() && tree.raw(0).ends_with(b"/>") {
        let start = root
            .end
            .checked_sub(2)
            .ok_or_else(|| crate::coded("style-malformed", "Invalid empty XML root"))?;
        return Ok(format!(
            "{}>{}</{}>{}",
            &raw[..start],
            xml,
            root.qname,
            &raw[root.end..]
        ));
    }
    let end = raw[..root.end]
        .rfind("</")
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing XML root end"))?;
    Ok(format!("{}{}{}", &raw[..end], xml, &raw[end..]))
}
pub(super) fn attributes(tree: &Tree, index: usize, allowed: &[&str]) -> Result<()> {
    if tree.nodes[index]
        .attrs
        .keys()
        .any(|key| key != "xmlns" && !key.starts_with("xmlns:") && !allowed.contains(&key.as_str()))
    {
        return Err(crate::coded(
            "style-protected",
            "Extended page theme metadata is protected",
        ));
    }
    Ok(())
}
pub(super) fn relationships(tree: &Tree) -> Result<Vec<usize>> {
    tree.root(PACKAGE_REL_NS, "Relationships")?;
    attributes(tree, 0, &[])?;
    let mut ids = BTreeSet::new();
    for index in &tree.nodes[0].children {
        let node = &tree.nodes[*index];
        if node.ns != PACKAGE_REL_NS || node.name != "Relationship" {
            return Err(crate::coded(
                "style-unsupported",
                "Foreign page theme relationship",
            ));
        }
        tree.clean(*index, &["Id", "Type", "Target", "TargetMode"])?;
        if ["Id", "Type", "Target"]
            .iter()
            .any(|key| node.attrs.get(*key).is_none_or(|value| value.is_empty()))
            || !ids.insert(node.attrs["Id"].clone())
        {
            return Err(crate::coded(
                "style-ambiguous",
                "Ambiguous page theme relationship",
            ));
        }
    }
    Ok(tree.nodes[0].children.clone())
}
/// Add the theme role to the Normal font; leave explicit cloned fonts, number
/// formats, XFs, fills and the original font's literal RGB meaning unchanged.
fn themed_styles(bytes: &[u8]) -> Result<String> {
    let tree = Tree::parse(bytes)?;
    tree.root(NS, "styleSheet")?;
    let bases = tree
        .direct(0, "cellStyleXfs")?
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing Normal style bases"))?;
    let mut normal = 0usize;
    if let Some(styles) = tree.direct(0, "cellStyles")? {
        let matches = tree.nodes[styles]
            .children
            .iter()
            .filter(|i| {
                let n = &tree.nodes[**i];
                n.attrs.get("builtinId").is_some_and(|v| v == "0")
                    || n.attrs.get("name").is_some_and(|v| v == "Normal")
            })
            .copied()
            .collect::<Vec<_>>();
        if matches.len() > 1 {
            return Err(crate::coded(
                "style-ambiguous",
                "Ambiguous Normal theme font",
            ));
        }
        if let Some(index) = matches.first() {
            normal = tree.nodes[*index]
                .attrs
                .get("xfId")
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing Normal style id"))?
                .parse::<usize>()
                .map_err(err)?;
        }
    }
    let base = *tree.nodes[bases]
        .children
        .get(normal)
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing Normal font base"))?;
    if tree.nodes[base].ns != NS || tree.nodes[base].name != "xf" {
        return Err(crate::coded(
            "style-unsupported",
            "Foreign Normal font base",
        ));
    }
    let font_id = tree.nodes[base]
        .attrs
        .get("fontId")
        .map(String::as_str)
        .unwrap_or("0")
        .parse::<usize>()
        .map_err(err)?;
    let fonts = tree
        .direct(0, "fonts")?
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing page theme fonts"))?;
    let font = *tree.nodes[fonts]
        .children
        .get(font_id)
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing Normal font"))?;
    if tree.nodes[font].ns != NS || tree.nodes[font].name != "font" {
        return Err(crate::coded(
            "style-unsupported",
            "Foreign Normal theme font",
        ));
    }
    attributes(&tree, font, &[])?;
    for child in &tree.nodes[font].children {
        let node = &tree.nodes[*child];
        if node.ns != NS {
            return Err(crate::coded(
                "style-unsupported",
                "Foreign Normal font child",
            ));
        }
        match node.name.as_str() {
            "name" | "sz" | "b" | "i" | "u" | "family" | "charset" | "scheme" => {
                tree.clean(*child, &["val"])?
            }
            "color" => tree.clean(*child, &["rgb", "auto", "indexed", "theme", "tint"])?,
            _ => {
                return Err(crate::coded(
                    "style-unsupported",
                    "Rich Normal font requires separate theme qualification",
                ));
            }
        }
    }
    if let Some(size) = tree.direct(font, "sz")? {
        let value = tree.nodes[size]
            .attrs
            .get("val")
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing Normal font size"))?
            .parse::<f64>()
            .map_err(err)?;
        if !value.is_finite() || value <= 0.0 || value > 409.5 {
            return Err(crate::coded(
                "style-unsupported",
                "Unresolved Normal font size",
            ));
        }
    }
    if let Some(name) = tree.direct(font, "name")? {
        let value = tree.nodes[name]
            .attrs
            .get("val")
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing Normal font name"))?;
        if value.trim().is_empty() || value.encode_utf16().count() > 128 {
            return Err(crate::coded(
                "style-unsupported",
                "Unresolved Normal font name",
            ));
        }
    }
    let scheme = tree.direct(font, "scheme")?;
    let color = tree.direct(font, "color")?;
    let mut replacement = std::str::from_utf8(tree.raw(font)).map_err(err)?.to_owned();
    if let Some(index) = scheme {
        if !["major", "minor", "none"].contains(
            &tree.nodes[index]
                .attrs
                .get("val")
                .map(String::as_str)
                .unwrap_or(""),
        ) {
            return Err(crate::coded(
                "internal-unknown-reference",
                "Unknown Normal font scheme",
            ));
        }
        let a = tree.nodes[index].start - tree.nodes[font].start;
        let b = tree.nodes[index].end - tree.nodes[font].start;
        replacement.replace_range(
            a..b,
            &format!(
                "<scheme xmlns=\"{}\" val=\"minor\"/>",
                String::from_utf8_lossy(NS)
            ),
        );
    }
    let mut additions = String::new();
    if scheme.is_none() {
        additions += &format!(
            "<scheme xmlns=\"{}\" val=\"minor\"/>",
            String::from_utf8_lossy(NS)
        );
    }
    if color.is_none() {
        additions += &format!(
            "<color xmlns=\"{}\" theme=\"1\"/>",
            String::from_utf8_lossy(NS)
        );
    }
    if !additions.is_empty() {
        if replacement.ends_with("/>") {
            replacement.truncate(replacement.len() - 2);
            replacement += &format!(">{}</{}>", additions, tree.nodes[font].qname);
        } else {
            let end = replacement
                .rfind("</")
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing font close"))?;
            replacement.insert_str(end, &additions);
        }
    }
    // Keep repeated selection of the same role history-pure. A canonical minor
    // scheme does not need to be rewritten solely to alter prefix spelling.
    if scheme.is_some_and(|index| {
        tree.nodes[index]
            .attrs
            .get("val")
            .is_some_and(|v| v == "minor")
    }) && additions.is_empty()
    {
        return String::from_utf8(bytes.to_vec()).map_err(err);
    }
    let mut output = tree.bytes.clone();
    output.splice(
        tree.nodes[font].start..tree.nodes[font].end,
        replacement.bytes(),
    );
    String::from_utf8(output).map_err(err)
}
pub(super) struct Prepared {
    pub plan: Value,
    pub command: Value,
    pub changed: bool,
    pub prior: Option<String>,
    pub revision: String,
}
impl Prepared {
    pub fn result(&self) -> Value {
        json!({"schema":workbook_format!(SCHEMA_PAGE_THEME_WRITER_V1),"file":"package.xlsx","contentChanged":false,"changed":self.changed,"priorThemeRevision":self.prior,"themeRevision":self.revision,"sample":self.command["sample"]})
    }
}
fn prepare_from(root: &Path, source: &str, command: &Value) -> Result<Prepared> {
    fields(command, &["type", "sample", "expectedThemeRevision"], &[])?;
    if command["type"] != "applyPageTheme" {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid page theme command",
        ));
    }
    let xml = sample(field(command, "sample")?)?;
    let expected = &command["expectedThemeRevision"];
    if !expected.is_null()
        && !expected.as_str().is_some_and(|s| {
            s.len() == 64
                && s.bytes()
                    .all(|b| b.is_ascii_digit() || matches!(b, b'a'..=b'f'))
        })
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid expected page theme revision",
        ));
    }
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, source)?).map_err(err)?,
    ))
    .map_err(err)?;
    let names = clear_hyperlinks::package_names(&zip)?;
    let rels = Tree::parse(&part(&mut zip, "xl/_rels/workbook.xml.rels")?)?;
    let edges = relationships(&rels)?;
    let style_edges = edges
        .iter()
        .filter(|index| {
            rels.nodes[**index].attrs["Type"]
                == format!("{}/styles", String::from_utf8_lossy(REL_NS))
        })
        .copied()
        .collect::<Vec<_>>();
    if style_edges.len() != 1 {
        return Err(crate::coded(
            "style-ambiguous",
            "Ambiguous page theme style ownership",
        ));
    }
    let style_edge = &rels.nodes[style_edges[0]];
    if style_edge
        .attrs
        .get("TargetMode")
        .is_some_and(|mode| mode != "Internal")
        || target("xl/workbook.xml", &style_edge.attrs["Target"])? != "xl/styles.xml"
    {
        return Err(crate::coded(
            "style-unsupported",
            "Foreign page theme style ownership",
        ));
    }
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
            "style-protected",
            "Protected/macro/signed page theme source stays read-only",
        ));
    }
    let worksheet_parts = crate::used_range_transfer::worksheet_inventory(&local(root, source)?)?;
    let mut metadata = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    for sheet in worksheet_parts {
        let tree = Tree::parse(&part(&mut metadata, &sheet.path)?)?;
        tree.root(NS, "worksheet")?;
        if tree
            .nodes
            .iter()
            .any(|node| node.ns == NS && node.name == "sheetProtection")
        {
            return Err(crate::coded(
                "style-protected",
                "Protected worksheet page theme source stays read-only",
            ));
        }
    }
    let theme_edges = edges
        .iter()
        .filter(|index| rels.nodes[**index].attrs["Type"] == THEME_EDGE)
        .copied()
        .collect::<Vec<_>>();
    if theme_edges.len() > 1 {
        return Err(crate::coded(
            "style-ambiguous",
            "Ambiguous workbook theme graph",
        ));
    }
    let prior_part = theme_edges
        .first()
        .map(|index| target("xl/workbook.xml", &rels.nodes[*index].attrs["Target"]))
        .transpose()?;
    let theme_part = prior_part.as_deref().unwrap_or(THEME_PART);

    let prior = if let Some(index) = theme_edges.first() {
        let node = &rels.nodes[*index];
        if node
            .attrs
            .get("TargetMode")
            .is_some_and(|value| value != "Internal")
            || !format_matches!(
                target("xl/workbook.xml", &node.attrs["Target"])?,
                XL_THEME_PAGE_THEME_XML
            )
        {
            return Err(crate::coded(
                "style-protected",
                "Imported/external workbook themes stay protected",
            ));
        }
        let bytes = part(&mut zip, theme_part)?;
        if !["neutral", "forest", "plum"].iter().any(|id| {
            sample(id).is_ok_and(|xml| crate::workbook_format::page_theme_matches(&bytes, xml))
        }) {
            return Err(crate::coded(
                "style-protected",
                "Modified/foreign owned-theme XML stays protected",
            ));
        }
        Some(digest(&bytes))
    } else {
        None
    };
    if expected != &json!(prior) {
        return Err(crate::coded(
            "style-stale",
            "Page theme changed since the captured revision",
        ));
    }
    if prior.is_none()
        && (names.contains(theme_part)
            || names.contains(workbook_format!(XL_THEME_PAGE_THEME_XML)))
    {
        return Err(crate::coded(
            "style-conflict",
            "Page theme part name is occupied",
        ));
    }
    // A workbook theme may not secretly be shared by an unrelated package owner.
    for name in names.iter().filter(|name| name.ends_with(".rels")) {
        let tree = Tree::parse(&part(&mut zip, name)?)?;
        let edges = relationships(&tree)?;
        let owner = if name == "_rels/.rels" {
            String::new()
        } else {
            let (directory, file) = name
                .rsplit_once("/_rels/")
                .ok_or_else(|| crate::coded("style-malformed", "Invalid relationship part"))?;
            format!(
                "{}/{}",
                directory,
                file.strip_suffix(".rels").ok_or_else(|| crate::coded(
                    "style-malformed",
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
            if target(&owner, &node.attrs["Target"])? == theme_part
                && (name != "xl/_rels/workbook.xml.rels" || node.attrs["Type"] != THEME_EDGE)
            {
                return Err(crate::coded(
                    "style-protected",
                    "Shared owned-theme graph stays protected",
                ));
            }
        }
    }
    let mut replacements = serde_json::Map::new();
    replacements.insert(theme_part.into(), json!(xml));
    if prior.is_none() {
        let mut n = 1;
        while edges.iter().any(|index| {
            rels.nodes[*index].attrs["Id"] == format!(workbook_format!(PAGETHEME_N), n = n)
        }) {
            n += 1;
        }
        replacements.insert(
            "xl/_rels/workbook.xml.rels".into(),
            json!(append(
                &rels,
                &format!(
                    workbook_format!(XML_BECA37B7DA),
                    String::from_utf8_lossy(PACKAGE_REL_NS),
                    n = n,
                    THEME_EDGE = THEME_EDGE
                )
            )?),
        );
    }
    let types = Tree::parse(&part(&mut zip, "[Content_Types].xml")?)?;
    types.root(TYPE_NS, "Types")?;
    let matches = types.nodes[0]
        .children
        .iter()
        .filter(|index| {
            types.nodes[**index]
                .attrs
                .get("PartName")
                .is_some_and(|name| name == &format!("/{theme_part}"))
        })
        .copied()
        .collect::<Vec<_>>();
    if matches.len() > 1 {
        return Err(crate::coded(
            "style-ambiguous",
            "Ambiguous theme content type",
        ));
    }
    if let Some(index) = matches.first() {
        let node = &types.nodes[*index];
        types.clean(*index, &["PartName", "ContentType"])?;
        if node.ns != TYPE_NS
            || node.name != "Override"
            || node.attrs.get("ContentType").map(String::as_str) != Some(THEME_TYPE)
        {
            return Err(crate::coded(
                "style-conflict",
                "Occupied theme content type",
            ));
        }
    } else {
        replacements.insert("[Content_Types].xml".into(),json!(append(&types,&format!("<Override xmlns=\"{}\" PartName=\"/{theme_part}\" ContentType=\"{THEME_TYPE}\"/>",String::from_utf8_lossy(TYPE_NS)))?));
    }
    let styles = part(&mut zip, "xl/styles.xml")?;
    let revised = themed_styles(&styles)?;
    if revised.as_bytes() != styles {
        replacements.insert("xl/styles.xml".into(), json!(revised));
    }
    let mut changed = false;
    for (name, value) in &replacements {
        if !names.contains(name)
            || part(&mut zip, name)?
                != value
                    .as_str()
                    .ok_or_else(|| crate::coded("style-malformed", "Invalid native theme XML"))?
                    .as_bytes()
        {
            changed = true;
        }
    }
    let plan = json!({"kind":"applyPageTheme","sheets":[],"copies":{},"removed":[],"replacements":replacements,"fragments":{}});
    Ok(Prepared {
        plan,
        command: command.clone(),
        changed,
        prior,
        revision: digest(xml.as_bytes()),
    })
}
pub(super) fn prepare(root: &Path, input: &Value) -> Result<Prepared> {
    fields(
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
            "Invalid native page theme request",
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
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing theme replacements"))?;
    names.extend(replacements.keys().cloned());
    if names != clear_hyperlinks::package_names(&after)? {
        return Err(crate::coded(
            "style-stale",
            "Theme writer changed unrelated package inventory",
        ));
    }
    for name in names {
        if let Some(xml) = replacements.get(&name) {
            if clear_hyperlinks::member_hash(&mut after, &name)?
                != Sha256::digest(
                    xml.as_str()
                        .ok_or_else(|| {
                            crate::coded("internal-invalid-field", "Invalid theme replacement")
                        })?
                        .as_bytes(),
                )
                .to_vec()
            {
                return Err(crate::coded(
                    "internal-mismatch",
                    "Theme output differs from native plan",
                ));
            }
        } else if clear_hyperlinks::member_hash(&mut before, &name)?
            != clear_hyperlinks::member_hash(&mut after, &name)?
        {
            return Err(crate::coded(
                "style-stale",
                "Theme writer changed unrelated package content",
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
pub(super) mod tests;
