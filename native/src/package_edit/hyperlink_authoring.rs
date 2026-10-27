//! Single-cell HTTP(S)/worksheet-link authoring. The native Clear engine admits
//! the complete source graph; this writer never edits cells, styles or formulas.
use super::{clear_hyperlinks, print_settings::Tree, *};
const KIND: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\r', "&#13;")
        .replace('\n', "&#10;")
        .replace('\t', "&#9;")
}
fn text(value: &str, max: usize) -> Result<()> {
    if value.encode_utf16().count() > max
        || value.chars().any(|c| {
            c < ' ' && !['\n', '\r', '\t'].contains(&c) || ['\u{fffe}', '\u{ffff}'].contains(&c)
        })
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid hyperlink text",
        ));
    }
    Ok(())
}
fn cell(value: &str) -> Result<()> {
    let p = crate::xml::position(value)?;
    if crate::xml::address(p.0, p.1) != value {
        return Err(crate::coded(
            "hyperlink-bad-input",
            "Hyperlink authoring requires a canonical single cell",
        ));
    }
    Ok(())
}
fn fields(value: &Value, required: &[&str], optional: &[&str]) -> Result<()> {
    let m = value.as_object().ok_or_else(|| {
        crate::coded(
            "internal-invalid-field",
            "Invalid hyperlink authoring fields",
        )
    })?;
    if required.iter().any(|k| !m.contains_key(*k))
        || m.keys()
            .any(|k| !required.contains(&k.as_str()) && !optional.contains(&k.as_str()))
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid hyperlink authoring fields",
        ));
    }
    Ok(())
}
fn url(value: &str) -> Result<()> {
    text(value, 2048)?;
    let tail = value
        .strip_prefix("https://")
        .or_else(|| value.strip_prefix("http://"))
        .ok_or_else(|| {
            crate::coded(
                "hyperlink-bad-input",
                "Only ordinary HTTP(S) hyperlinks are supported",
            )
        })?;
    if value
        .chars()
        .any(|c| c.is_whitespace() || c.is_control() || c == '\\')
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid web hyperlink",
        ));
    }
    let authority = tail.split(['/', '?', '#']).next().unwrap_or("");
    if authority.is_empty() || authority.contains(['@', '%']) {
        return Err(crate::coded(
            "internal-missing-field",
            "Web hyperlink authority is missing or uses credentials/escapes",
        ));
    }
    let port = if authority.starts_with('[') {
        let end = authority.find(']').ok_or_else(|| {
            crate::coded("internal-invalid-field", "Invalid web hyperlink IPv6 host")
        })?;
        authority[1..end]
            .parse::<std::net::Ipv6Addr>()
            .map_err(err)?;
        let rest = &authority[end + 1..];
        if rest.is_empty() {
            None
        } else {
            Some(rest.strip_prefix(':').ok_or_else(|| {
                crate::coded("internal-invalid-field", "Invalid web hyperlink port")
            })?)
        }
    } else {
        let mut pieces = authority.split(':');
        let host = pieces.next().unwrap_or("");
        if host.is_empty()
            || host.split('.').any(|s| {
                s.is_empty()
                    || s.starts_with('-')
                    || s.ends_with('-')
                    || !s.bytes().all(|b| b.is_ascii_alphanumeric() || b == b'-')
            })
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid web hyperlink host",
            ));
        }
        let port = pieces.next();
        if pieces.next().is_some() {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid web hyperlink authority",
            ));
        }
        port
    };
    if let Some(port) = port {
        if port.is_empty()
            || !port.bytes().all(|b| b.is_ascii_digit())
            || port.parse::<u16>().map_err(err)? == 0
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid web hyperlink port",
            ));
        }
    }
    let bytes = value.as_bytes();
    for at in 0..bytes.len() {
        if bytes[at] == b'%'
            && (!bytes.get(at + 1).is_some_and(u8::is_ascii_hexdigit)
                || !bytes.get(at + 2).is_some_and(u8::is_ascii_hexdigit))
        {
            return Err(crate::coded(
                "hyperlink-malformed",
                "Invalid web hyperlink percent escape",
            ));
        }
    }
    Ok(())
}
fn prefix(tree: &Tree) -> String {
    tree.nodes[0]
        .qname
        .rsplit_once(':')
        .map(|(p, _)| format!("{p}:"))
        .unwrap_or_default()
}
fn append(tree: &Tree, body: &str) -> Result<String> {
    let source = std::str::from_utf8(&tree.bytes).map_err(err)?;
    let node = &tree.nodes[0];
    let raw = std::str::from_utf8(tree.raw(0)).map_err(err)?;
    let result = if let Some(head) = raw.strip_suffix("/>") {
        format!(
            "{}{}>{}</{}>{}",
            &source[..node.start],
            head,
            body,
            node.qname,
            &source[node.end..]
        )
    } else {
        let close = raw.rfind("</").ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing hyperlink XML closing tag",
            )
        })?;
        format!(
            "{}{}{}{}",
            &source[..node.start + close],
            body,
            &source[node.start + close..node.end],
            &source[node.end..]
        )
    };
    if result.len() > XML_LIMIT {
        return Err(crate::coded(
            "hyperlink-limit",
            "Hyperlink authoring metadata exceeds its bound",
        ));
    }
    Ok(result)
}
fn relation_part(sheet: &str) -> Result<String> {
    let (dir, file) = sheet
        .rsplit_once('/')
        .ok_or_else(|| crate::coded("hyperlink-malformed", "Invalid hyperlink worksheet part"))?;
    Ok(format!("{dir}/_rels/{file}.rels"))
}
fn part(zip: &mut ZipArchive<crate::cancel::Reader<File>>, name: &str) -> Result<Vec<u8>> {
    part_bytes(zip, name, &serde_json::Map::new())
}
fn fingerprint(tree: &Tree, index: usize, rels: Option<&Tree>) -> Result<String> {
    let mut digest = Sha256::new();
    digest.update(tree.raw(index));
    digest.update([0]);
    if let Some(id) = tree.nodes[index].attrs.get("r:id") {
        let rels = rels.ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing hyperlink relationships")
        })?;
        let matches = rels.nodes[0]
            .children
            .iter()
            .filter(|i| rels.nodes[**i].attrs.get("Id") == Some(id))
            .collect::<Vec<_>>();
        if matches.len() != 1 {
            return Err(crate::coded(
                "hyperlink-ambiguous",
                "Ambiguous hyperlink relationship revision",
            ));
        }
        digest.update(rels.raw(*matches[0]));
    }
    Ok(format!("{:x}", digest.finalize()))
}
pub(super) struct Prepared {
    pub plan: Value,
    pub command: Value,
    pub revision: Option<String>,
    pub changed: bool,
    pub remaining: u64,
}
impl Prepared {
    pub fn result(&self) -> Value {
        json!({"schema":workbook_format!(SCHEMA_HYPERLINK_WRITER_V1),"file":"package.xlsx","contentChanged":false,"changed":self.changed,"priorHyperlinkRevision":self.revision,"remainingHyperlinks":self.remaining})
    }
}
fn prepare_from(root: &Path, source: &str, command: &Value) -> Result<Prepared> {
    let remove = command["type"] == "removeHyperlink";
    fields(
        command,
        &["type", "sheet", "cell", "expectedHyperlinkRevision"],
        if remove { &[] } else { &["link", "tooltip"] },
    )?;
    if !remove && command["type"] != "setHyperlink" || !remove && command.get("link").is_none() {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid hyperlink authoring command",
        ));
    }
    let reference = field(command, "cell")?;
    cell(reference)?;
    let sheet = field(command, "sheet")?;
    if !command["expectedHyperlinkRevision"].is_null()
        && !command["expectedHyperlinkRevision"]
            .as_str()
            .is_some_and(|s| {
                s.len() == 64
                    && s.bytes()
                        .all(|b| b.is_ascii_digit() || matches!(b, b'a'..=b'f'))
            })
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid expected hyperlink revision",
        ));
    }
    if let Some(tooltip) = command.get("tooltip") {
        text(
            tooltip.as_str().ok_or_else(|| {
                crate::coded("internal-invalid-field", "Invalid hyperlink tooltip")
            })?,
            256,
        )?;
    }
    let inventory = crate::used_range_transfer::worksheet_inventory(&local(root, source)?)?;
    let selected = inventory.iter().find(|s| s.name == sheet).ok_or_else(|| {
        crate::coded(
            "internal-unknown-reference",
            "Unknown hyperlink authoring worksheet",
        )
    })?;
    let sheet_part = &selected.path;
    let rel_part = relation_part(sheet_part)?;
    let mut metadata = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let tree = Tree::parse(&part(&mut metadata, sheet_part)?)?;
    tree.root(NS, "worksheet")?;
    let has_rels = metadata.file_names().any(|n| n == rel_part);
    let rels = if has_rels {
        Some(Tree::parse(&part(&mut metadata, &rel_part)?)?)
    } else {
        None
    };
    let collection = tree.direct(0, "hyperlinks")?;
    let mut hit = None;
    if let Some(index) = collection {
        for child in &tree.nodes[index].children {
            let n = &tree.nodes[*child];
            let r = field(&json!(n.attrs), "ref")?.to_string();
            let (a, b) = crate::bulk::bounds(&r)?;
            let p = crate::xml::position(reference)?;
            if p.0 >= a.0 && p.0 <= b.0 && p.1 >= a.1 && p.1 <= b.1 {
                if a != b || hit.is_some() {
                    return Err(crate::coded(
                        "hyperlink-unsupported",
                        "Range/shared hyperlink authoring is unsupported; preserve its original graph",
                    ));
                }
                hit = Some(*child);
            }
        }
    }
    // Complete graph validation and independently derived single-record removal.
    let baseline = clear_hyperlinks::prepare_from(
        root,
        source,
        &json!({"type":"clearHyperlinks","sheet":sheet,"ranges":[reference]}),
    )?;
    let revision = hit
        .map(|i| fingerprint(&tree, i, rels.as_ref()))
        .transpose()?;
    if command["expectedHyperlinkRevision"] != json!(revision) {
        return Err(crate::coded(
            "hyperlink-stale",
            "The hyperlink changed; inspect it again",
        ));
    }
    if let Some(index) = hit {
        let n = &tree.nodes[index];
        if n.attrs.contains_key("r:id") && n.attrs.contains_key("location") {
            return Err(crate::coded(
                "hyperlink-unsupported",
                "Combined external/location hyperlink authoring is unsupported",
            ));
        }
        if let Some(location) = n.attrs.get("location") {
            destination(sheet, location, &inventory)?;
        }
        if let Some(id) = n.attrs.get("r:id") {
            let rels = rels.as_ref().ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing hyperlink relationship")
            })?;
            let edge = rels.nodes[0]
                .children
                .iter()
                .find(|i| rels.nodes[**i].attrs.get("Id") == Some(id))
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing hyperlink edge"))?;
            let edge = &rels.nodes[*edge];
            if edge.attrs.get("TargetMode").map(String::as_str) != Some("External") {
                return Err(crate::coded(
                    "hyperlink-unsupported",
                    "Package-local hyperlink authoring is unsupported",
                ));
            }
            url(edge.attrs.get("Target").ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing web hyperlink target")
            })?)?;
        }
    }
    let mut plan = baseline.plan;
    plan["kind"] = json!(if remove {
        "removeHyperlink"
    } else {
        "setHyperlink"
    });
    if remove {
        return Ok(Prepared {
            plan,
            command: command.clone(),
            revision,
            changed: hit.is_some(),
            remaining: baseline.counts.remaining,
        });
    }
    if baseline.counts.remaining >= 10_000 {
        return Err(crate::coded(
            "hyperlink-limit",
            "Hyperlink collection exceeds its authoring bound",
        ));
    }
    let link = &command["link"];
    let mut attrs = format!(" ref=\"{}\"", escape(reference));
    match field(link, "kind")? {
        "url" => {
            fields(link, &["kind", "target"], &[])?;
            let target = field(link, "target")?;
            url(target)?;
            let types = Tree::parse(&part(&mut metadata, "[Content_Types].xml")?)?;
            types.root(TYPE_NS, "Types")?;
            let defaults = types.nodes[0]
                .children
                .iter()
                .filter(|i| {
                    types.nodes[**i].name == "Default"
                        && types.nodes[**i].attrs.get("Extension").map(String::as_str)
                            == Some("rels")
                })
                .collect::<Vec<_>>();
            if defaults.len() != 1
                || types.nodes[*defaults[0]].ns != TYPE_NS
                || types.nodes[*defaults[0]]
                    .attrs
                    .get("ContentType")
                    .map(String::as_str)
                    != Some("application/vnd.openxmlformats-package.relationships+xml")
            {
                return Err(crate::coded(
                    "hyperlink-refused",
                    "Hyperlink relationships lack the admitted content type",
                ));
            }
            let raw = plan["replacements"]
                .get(&rel_part)
                .and_then(Value::as_str)
                .map(str::to_owned)
                .or_else(|| {
                    rels.as_ref()
                        .map(|t| String::from_utf8_lossy(&t.bytes).into_owned())
                })
                .unwrap_or_else(|| {
                    format!(
                        "<Relationships xmlns=\"{}\"/>",
                        String::from_utf8_lossy(PACKAGE_REL_NS)
                    )
                });
            let remaining = Tree::parse(raw.as_bytes())?;
            remaining.root(PACKAGE_REL_NS, "Relationships")?;
            let mut index = 1usize;
            let id = loop {
                let id = format!(workbook_format!(HYPERLINK_INDEX), index = index);
                if !tree.nodes.iter().any(|n| {
                    n.attrs
                        .values()
                        .any(|v| v.split_whitespace().any(|t| t == id))
                        || n.text.split_whitespace().any(|t| t == id)
                }) && !rels
                    .as_ref()
                    .is_some_and(|t| t.nodes.iter().any(|n| n.attrs.get("Id") == Some(&id)))
                {
                    break id;
                }
                index += 1;
                if index > 50_000 {
                    return Err(crate::coded(
                        "hyperlink-bad-input",
                        "No bounded unique hyperlink relationship ID",
                    ));
                }
            };
            let edge = format!(
                "<{}Relationship Id=\"{}\" Type=\"{}\" Target=\"{}\" TargetMode=\"External\"/>",
                prefix(&remaining),
                id,
                KIND,
                escape(target)
            );
            plan["replacements"][&rel_part] = json!(append(&remaining, &edge)?);
            attrs.push_str(&format!(
                " xmlns:r=\"{}\" r:id=\"{}\"",
                String::from_utf8_lossy(REL_NS),
                id
            ));
        }
        "internal" => {
            fields(link, &["kind", "sheet", "cell"], &[])?;
            let target_sheet = field(link, "sheet")?;
            let target_cell = field(link, "cell")?;
            cell(target_cell)?;
            if !inventory.iter().any(|s| s.name == target_sheet) {
                return Err(crate::coded(
                    "internal-missing-field",
                    "Internal hyperlink target worksheet is missing",
                ));
            }
            attrs.push_str(&format!(
                " location=\"{}\"",
                escape(&format!(
                    "'{}'!{}",
                    target_sheet.replace('\'', "''"),
                    target_cell
                ))
            ));
        }
        _ => {
            return Err(crate::coded(
                "hyperlink-unsupported",
                "Unsupported hyperlink destination",
            ));
        }
    }
    if let Some(t) = command.get("tooltip") {
        attrs.push_str(&format!(
            " tooltip=\"{}\"",
            escape(t.as_str().ok_or_else(|| crate::coded(
                "internal-invalid-field",
                "Invalid hyperlink tooltip"
            ))?)
        ));
    }
    let new_link = format!(
        "<hyperlink xmlns=\"{}\"{attrs}/>",
        String::from_utf8_lossy(NS)
    );
    let raw = plan["fragments"]
        .get(sheet_part)
        .and_then(|f| f["hyperlinks"].as_str())
        .map(str::to_owned)
        .unwrap_or_else(|| {
            collection
                .map(|i| String::from_utf8_lossy(tree.raw(i)).into_owned())
                .unwrap_or_default()
        });
    let fresh = raw.is_empty();
    let raw = if fresh {
        format!("<hyperlinks xmlns=\"{}\"/>", String::from_utf8_lossy(NS))
    } else {
        raw
    };
    // Existing fragments inherit the worksheet's prefix/default namespace. Bind
    // that root explicitly before parsing it as a standalone bounded fragment.
    let qualified = if let Some(index) = collection.filter(|_| !fresh) {
        let node = &tree.nodes[index];
        let inherited = tree.nodes[0]
            .attrs
            .iter()
            .filter(|(k, _)| {
                (k.as_str() == "xmlns" || k.starts_with("xmlns:")) && !node.attrs.contains_key(*k)
            })
            .map(|(k, v)| format!(" {k}=\"{}\"", escape(v)))
            .collect::<String>();
        raw.replacen(
            &format!("<{}", node.qname),
            &format!("<{}{inherited}", node.qname),
            1,
        )
    } else {
        raw
    };
    let links = Tree::parse(qualified.as_bytes())?;
    links.root(NS, "hyperlinks")?;
    plan["fragments"][sheet_part] = json!({"hyperlinks":append(&links,&new_link)?});
    Ok(Prepared {
        plan,
        command: command.clone(),
        revision,
        changed: true,
        remaining: baseline.counts.remaining + 1,
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
            "Invalid native hyperlink authoring request",
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
    let names = clear_hyperlinks::package_names(&before)?;
    let mut expected = names.clone();
    let replacements = prepared.plan["replacements"]
        .as_object()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing hyperlink replacements"))?;
    expected.extend(replacements.keys().cloned());
    if expected != clear_hyperlinks::package_names(&after)? {
        return Err(crate::coded(
            "hyperlink-stale",
            "Hyperlink authoring changed unrelated package inventory",
        ));
    }
    let fragments = prepared.plan["fragments"]
        .as_object()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing hyperlink fragments"))?;
    for name in expected {
        crate::cancel::check()?;
        if let Some(f) = fragments.get(&name) {
            let (a, _, _) = clear_hyperlinks::worksheet_hash(&mut before, &name, None)?;
            let (b, links, _) = clear_hyperlinks::worksheet_hash(&mut after, &name, None)?;
            if a != b || links != field(f, "hyperlinks")?.as_bytes() {
                return Err(crate::coded(
                    "hyperlink-stale",
                    "Hyperlink authoring changed nonlink worksheet content",
                ));
            }
        } else if let Some(xml) = replacements.get(&name) {
            if clear_hyperlinks::member_hash(&mut after, &name)?
                != Sha256::digest(
                    xml.as_str()
                        .ok_or_else(|| {
                            crate::coded("internal-invalid-field", "Invalid hyperlink replacement")
                        })?
                        .as_bytes(),
                )
                .to_vec()
            {
                return Err(crate::coded(
                    "internal-mismatch",
                    "Hyperlink relationship output differs from the native plan",
                ));
            }
        } else if clear_hyperlinks::member_hash(&mut before, &name)?
            != clear_hyperlinks::member_hash(&mut after, &name)?
        {
            return Err(crate::coded(
                "hyperlink-stale",
                "Hyperlink authoring changed unrelated package content",
            ));
        }
    }
    Ok(())
}
pub(super) fn verify_receipt_package(root: &Path, command: &Value, output: &str) -> Result<()> {
    let prepared = prepare_from(root, "export.xlsx", command)?;
    verify_output(root, "export.xlsx", output, &prepared)
}
fn destination(
    sheet: &str,
    location: &str,
    inventory: &[crate::used_range_transfer::SheetBinding],
) -> Result<Value> {
    let (name, reference) = if let Some((name, reference)) = location.rsplit_once('!') {
        let name = if name.starts_with('\'') && name.ends_with('\'') && name.len() >= 2 {
            name[1..name.len() - 1].replace("''", "'")
        } else {
            name.to_string()
        };
        (name, reference.to_string())
    } else {
        (sheet.to_string(), location.to_string())
    };
    let p = clear_hyperlinks::point(&reference)?;
    let reference = crate::xml::address(p.0, p.1);
    if !inventory.iter().any(|s| s.name == name) {
        return Err(crate::coded(
            "hyperlink-bad-input",
            "Named/external/unknown-sheet hyperlink destinations are not editable",
        ));
    }
    Ok(json!({"kind":"internal","sheet":name,"cell":reference}))
}
pub(super) fn read(root: &Path, input: &Value) -> Result<Value> {
    fields(
        input,
        &[
            "action",
            "operation",
            "root",
            "version",
            "sourceRevision",
            "sheet",
            "cell",
        ],
        &["diagnostics"],
    )?;
    if input["action"] != "workbook"
        || input["operation"] != "hyperlink"
        || input.get("diagnostics").is_some_and(|v| !v.is_boolean())
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid native hyperlink inspection operation",
        ));
    }
    let version = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid hyperlink version"))?;
    let db = crate::database(root)?;
    if !db
        .query_row(
            "SELECT EXISTS(SELECT 1 FROM versions WHERE id=?1)",
            [version],
            |r| r.get::<_, bool>(0),
        )
        .map_err(err)?
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown hyperlink inspection version",
        ));
    }
    let index: Value =
        serde_json::from_slice(&fs::read(local(root, "index.json")?).map_err(err)?).map_err(err)?;
    if index["revision"] != input["sourceRevision"]
        || crate::hash(&local(root, "source.xlsx")?)? != field(input, "sourceRevision")?
    {
        return Err(crate::coded(
            "hyperlink-stale",
            "Hyperlink inspection source revision changed",
        ));
    }
    let sheet = field(input, "sheet")?;
    let reference = field(input, "cell")?;
    cell(reference)?;
    let inventory = crate::used_range_transfer::worksheet_inventory(&local(root, "source.xlsx")?)?;
    let selected = inventory.iter().find(|s| s.name == sheet).ok_or_else(|| {
        crate::coded(
            "internal-unknown-reference",
            "Unknown hyperlink inspection worksheet",
        )
    })?;
    // This native graph admission is read-only: it validates all relationships
    // and opaque references but neither publishes a plan nor changes the owner.
    clear_hyperlinks::prepare_from(
        root,
        "source.xlsx",
        &json!({"type":"clearHyperlinks","sheet":sheet,"ranges":[reference]}),
    )?;
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let tree = Tree::parse(&part(&mut zip, &selected.path)?)?;
    let rel_part = relation_part(&selected.path)?;
    let rels = if zip.file_names().any(|n| n == rel_part) {
        Some(Tree::parse(&part(&mut zip, &rel_part)?)?)
    } else {
        None
    };
    let mut reply = json!({"schema":workbook_format!(SCHEMA_HYPERLINK_DEFINITION_V1),"sheet":sheet,"cell":reference,"hyperlinkRevision":null,"link":null,"tooltip":"","canEdit":true,"protectedReason":null});
    if let Some(index) = tree.direct(0, "hyperlinks")? {
        for child in &tree.nodes[index].children {
            let n = &tree.nodes[*child];
            let (a, b) = crate::bulk::bounds(n.attrs.get("ref").ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing hyperlink reference")
            })?)?;
            let p = crate::xml::position(reference)?;
            if !(p.0 >= a.0 && p.0 <= b.0 && p.1 >= a.1 && p.1 <= b.1) {
                continue;
            }
            reply["hyperlinkRevision"] = json!(fingerprint(&tree, *child, rels.as_ref())?);
            let parsed: Result<Value> = (|| {
                if a != b {
                    return Err(crate::coded(
                        "hyperlink-unsupported",
                        "Range/shared hyperlink authoring is unsupported; preserve its original graph",
                    ));
                }
                if n.attrs.contains_key("r:id") && n.attrs.contains_key("location") {
                    return Err(crate::coded(
                        "hyperlink-unsupported",
                        "Combined external/location hyperlink authoring is unsupported",
                    ));
                }
                let tooltip = n.attrs.get("tooltip").map(String::as_str).unwrap_or("");
                text(tooltip, 256)?;
                let link = if let Some(id) = n.attrs.get("r:id") {
                    let rels = rels.as_ref().ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing hyperlink relationships")
                    })?;
                    let edge = rels.nodes[0]
                        .children
                        .iter()
                        .find(|i| rels.nodes[**i].attrs.get("Id") == Some(id))
                        .ok_or_else(|| {
                            crate::coded("internal-missing-field", "Missing hyperlink edge")
                        })?;
                    let edge = &rels.nodes[*edge];
                    if edge.attrs.get("TargetMode").map(String::as_str) != Some("External") {
                        return Err(crate::coded(
                            "hyperlink-unsupported",
                            "Package-local hyperlink authoring is unsupported",
                        ));
                    }
                    let target = edge.attrs.get("Target").ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing web hyperlink target")
                    })?;
                    url(target)?;
                    json!({"kind":"url","target":target})
                } else {
                    destination(
                        sheet,
                        n.attrs.get("location").ok_or_else(|| {
                            crate::coded(
                                "internal-missing-field",
                                "Missing internal hyperlink destination",
                            )
                        })?,
                        &inventory,
                    )?
                };
                Ok(json!({"link":link,"tooltip":tooltip}))
            })();
            match parsed {
                Ok(value) => {
                    reply["link"] = value["link"].clone();
                    reply["tooltip"] = value["tooltip"].clone();
                }
                Err(reason) => {
                    reply["canEdit"] = json!(false);
                    reply["protectedReason"] = json!(reason);
                }
            }
            break;
        }
    }
    Ok(reply)
}

#[cfg(test)]
mod tests;
