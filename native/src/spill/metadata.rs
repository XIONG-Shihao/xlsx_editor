//! Independent admission of closed spill metadata deltas.
//! No XML part or master index is accepted merely because its path is familiar.
use crate::{Result, err};
use serde_json::Value;
use std::collections::{BTreeMap, BTreeSet};
mod writer;
pub(crate) use writer::build;
const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const XDA: &str = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray";
const RICH: &str = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";
const RICH2: &str = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2";
const REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
pub(crate) const PARTS: [&str; 6] = [
    "xl/metadata.xml",
    "xl/richData/rdrichvalue.xml",
    "xl/richData/rdrichvaluestructure.xml",
    "xl/richData/rdRichValueTypes.xml",
    "xl/_rels/workbook.xml.rels",
    "[Content_Types].xml",
];
const GRAPH: [(&str, &str, &str, &str); 4] = [
    (
        PARTS[0],
        "metadata.xml",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
    ),
    (
        PARTS[1],
        "richData/rdrichvalue.xml",
        "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue",
        "application/vnd.ms-excel.rdrichvalue+xml",
    ),
    (
        PARTS[2],
        "richData/rdrichvaluestructure.xml",
        "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure",
        "application/vnd.ms-excel.rdrichvaluestructure+xml",
    ),
    (
        PARTS[3],
        "richData/rdRichValueTypes.xml",
        "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes",
        "application/vnd.ms-excel.rdrichvaluetypes+xml",
    ),
];
#[derive(Clone, Debug)]
struct Node {
    ns: String,
    name: String,
    attrs: BTreeMap<String, String>,
    scope: BTreeMap<String, String>,
    text: String,
    tail: String,
    children: Vec<usize>,
}
struct Tree {
    nodes: Vec<Node>,
}
impl Tree {
    fn parse(bytes: &[u8]) -> Result<Self> {
        use quick_xml::{NsReader, events::Event, name::ResolveResult};
        if bytes.len() > 8 * 1024 * 1024 {
            return Err(crate::coded(
                "spill-limit",
                "Spill metadata exceeds its XML bound",
            ));
        }
        let mut reader = NsReader::from_reader(bytes);
        let mut buffer = Vec::new();
        let mut nodes: Vec<Node> = Vec::new();
        let mut stack: Vec<usize> = Vec::new();
        let mut roots = 0;
        loop {
            crate::cancel::check()?;
            let (binding, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
            let empty = matches!(&event, Event::Empty(_));
            let namespace = match binding {
                ResolveResult::Bound(ns) => {
                    std::str::from_utf8(ns.as_ref()).map_err(err)?.to_owned()
                }
                ResolveResult::Unbound => String::new(),
                _ => {
                    return Err(crate::coded(
                        "spill-unsupported",
                        "Unbound spill metadata namespace prefix",
                    ));
                }
            };
            match event {
                Event::Start(element) | Event::Empty(element) => {
                    if stack.len() > 64 || nodes.len() > 200000 {
                        return Err(crate::coded(
                            "spill-limit",
                            "Spill metadata node/depth bound exceeded",
                        ));
                    }
                    if stack.is_empty() {
                        roots += 1;
                        if roots != 1 {
                            return Err(crate::coded(
                                "spill-malformed",
                                "Multiple spill metadata roots",
                            ));
                        }
                    }
                    let mut scope = stack
                        .last()
                        .map(|index| nodes[*index].scope.clone())
                        .unwrap_or_default();
                    let mut attributes = BTreeMap::new();
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(err)?;
                        let key = std::str::from_utf8(attribute.key.as_ref()).map_err(err)?;
                        let value = attribute.unescape_value().map_err(err)?.into_owned();
                        if key == "xmlns" {
                            scope.insert(String::new(), value);
                        } else if let Some(prefix) = key.strip_prefix("xmlns:") {
                            scope.insert(prefix.into(), value);
                        } else {
                            let (binding, local) = reader.resolve_attribute(attribute.key);
                            let key = match binding {
                                ResolveResult::Bound(ns) => format!(
                                    "{{{}}}{}",
                                    std::str::from_utf8(ns.as_ref()).map_err(err)?,
                                    std::str::from_utf8(local.as_ref()).map_err(err)?
                                ),
                                ResolveResult::Unbound => key.to_owned(),
                                _ => {
                                    return Err(crate::coded(
                                        "spill-unsupported",
                                        "Unbound spill metadata attribute prefix",
                                    ));
                                }
                            };
                            if attributes.insert(key, value).is_some() {
                                return Err(crate::coded(
                                    "spill-duplicate",
                                    "Duplicate spill metadata attribute",
                                ));
                            }
                        }
                    }
                    let index = nodes.len();
                    nodes.push(Node {
                        ns: namespace,
                        name: std::str::from_utf8(element.local_name().as_ref())
                            .map_err(err)?
                            .into(),
                        attrs: attributes,
                        scope,
                        text: String::new(),
                        tail: String::new(),
                        children: Vec::new(),
                    });
                    if let Some(parent) = stack.last() {
                        nodes[*parent].children.push(index);
                    }
                    if !empty {
                        stack.push(index);
                    }
                }
                Event::End(_) => {
                    stack.pop().ok_or_else(|| {
                        crate::coded("spill-malformed", "Unbalanced spill metadata")
                    })?;
                }
                Event::Text(text) => {
                    let text = crate::xml::xml_text(text.as_ref())?;
                    if let Some(parent) = stack.last() {
                        if let Some(child) = nodes[*parent].children.last().copied() {
                            nodes[child].tail.push_str(&text);
                        } else {
                            nodes[*parent].text.push_str(&text);
                        }
                    } else if !text.trim().is_empty() {
                        return Err(crate::coded(
                            "spill-malformed",
                            "Text outside spill metadata root",
                        ));
                    }
                }
                Event::CData(text) => {
                    let text = std::str::from_utf8(text.as_ref())
                        .map_err(err)?
                        .replace("\r\n", "\n")
                        .replace('\r', "\n");
                    let parent = *stack.last().ok_or_else(|| {
                        crate::coded("spill-unsupported", "CDATA outside spill metadata root")
                    })?;
                    if let Some(child) = nodes[parent].children.last().copied() {
                        nodes[child].tail.push_str(&text);
                    } else {
                        nodes[parent].text.push_str(&text);
                    }
                }
                Event::Comment(_) | Event::PI(_) => {
                    return Err(crate::coded(
                        "spill-unsupported",
                        "Spill metadata comments/processing instructions require preservation qualification",
                    ));
                }
                Event::DocType(_) => {
                    return Err(crate::coded(
                        "spill-malformed",
                        "Spill metadata XML entities are refused",
                    ));
                }
                Event::Eof => break,
                _ => (),
            }
            buffer.clear();
        }
        if roots != 1 || !stack.is_empty() {
            return Err(crate::coded(
                "spill-malformed",
                "Unbalanced spill metadata root",
            ));
        }
        Ok(Self { nodes })
    }
    fn node(&self, index: usize) -> Result<&Node> {
        self.nodes.get(index).ok_or_else(|| {
            crate::coded(
                "spill-unsupported",
                "Spill metadata index is outside its container",
            )
        })
    }
    fn root(&self, ns: &str, name: &str) -> Result<&Node> {
        let root = self.node(0)?;
        if root.ns != ns || root.name != name {
            return Err(crate::coded(
                "spill-unsupported",
                "Unqualified spill metadata root",
            ));
        }
        Ok(root)
    }
    fn one(&self, parent: &Node, ns: &str, name: &str, required: bool) -> Result<Option<usize>> {
        let matches = parent
            .children
            .iter()
            .copied()
            .filter(|index| {
                let n = &self.nodes[*index];
                n.ns == ns && n.name == name
            })
            .collect::<Vec<_>>();
        if matches.len() > 1 || required && matches.is_empty() {
            return Err(crate::coded(
                "internal-missing-field",
                "Missing/ambiguous spill metadata container",
            ));
        }
        Ok(matches.first().copied())
    }
    fn children<'a>(&'a self, parent: &Node, ns: &str, name: &str) -> Result<Vec<&'a Node>> {
        parent
            .children
            .iter()
            .map(|index| {
                let n = self.node(*index)?;
                if n.ns != ns || n.name != name {
                    return Err(crate::coded(
                        "spill-unsupported",
                        "Unqualified spill metadata child",
                    ));
                }
                Ok(n)
            })
            .collect()
    }
    fn counted<'a>(&'a self, parent: &Node, ns: &str, name: &str) -> Result<Vec<&'a Node>> {
        let nodes = self.children(parent, ns, name)?;
        if integer(attr(parent, "count")?, false)? != nodes.len() {
            return Err(crate::coded(
                "spill-bad-input",
                "Spill metadata count does not match children",
            ));
        }
        Ok(nodes)
    }
}
fn attr<'a>(node: &'a Node, key: &str) -> Result<&'a str> {
    node.attrs
        .get(key)
        .map(String::as_str)
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing spill metadata attribute"))
}
fn integer(value: &str, one: bool) -> Result<usize> {
    if value.is_empty()
        || value.len() > 1 && value.starts_with('0')
        || !value.bytes().all(|c| c.is_ascii_digit())
    {
        return Err(crate::coded(
            "spill-malformed",
            "Noncanonical spill metadata index",
        ));
    }
    let n = value.parse::<usize>().map_err(err)?;
    if one {
        n.checked_sub(1).ok_or_else(|| {
            crate::coded("spill-malformed", "Invalid one-based spill metadata index")
        })
    } else {
        Ok(n)
    }
}
fn json_index(value: &Value, one: bool) -> Result<usize> {
    let n = value
        .as_u64()
        .ok_or_else(|| crate::coded("spill-malformed", "Invalid mapped spill metadata index"))?;
    let n = usize::try_from(n).map_err(err)?;
    if one {
        n.checked_sub(1)
            .ok_or_else(|| crate::coded("spill-malformed", "Invalid one-based mapped spill index"))
    } else {
        Ok(n)
    }
}
fn whitespace(tree: &Tree, node: &Node) -> Result<()> {
    if !node.text.trim().is_empty()
        || node
            .children
            .iter()
            .any(|i| !tree.nodes[*i].tail.trim().is_empty())
    {
        return Err(crate::coded(
            "spill-malformed",
            "Unexpected spill metadata text/tail",
        ));
    }
    Ok(())
}
fn leaf(tree: &Tree, node: &Node) -> Result<()> {
    whitespace(tree, node)?;
    if !node.children.is_empty() {
        return Err(crate::coded(
            "spill-malformed",
            "Unexpected spill metadata leaf child",
        ));
    }
    Ok(())
}
fn attributes(node: &Node, expected: &[(&str, &str)]) -> Result<()> {
    let map = expected
        .iter()
        .map(|(key, value)| ((*key).to_string(), (*value).to_string()))
        .collect::<BTreeMap<_, _>>();
    if node.attrs != map {
        return Err(crate::coded(
            "spill-unsupported",
            "Unqualified spill metadata attributes",
        ));
    }
    Ok(())
}
fn identical(old_tree: &Tree, old: &Node, new_tree: &Tree, new: &Node, append: bool) -> Result<()> {
    if old.ns != new.ns
        || old.name != new.name
        || old.text != new.text
        || old.tail != new.tail
        || old
            .scope
            .iter()
            .any(|(key, value)| new.scope.get(key) != Some(value))
    {
        return Err(crate::coded(
            "spill-stale",
            "Opaque spill metadata text/namespaces changed",
        ));
    }
    let filtered = |node: &Node| {
        node.attrs
            .iter()
            .filter(|(key, _)| !append || key.as_str() != "count")
            .map(|(key, value)| (key.clone(), value.clone()))
            .collect::<BTreeMap<_, _>>()
    };
    if filtered(old) != filtered(new)
        || new.children.len() < old.children.len()
        || !append && new.children.len() != old.children.len()
    {
        return Err(crate::coded(
            "spill-stale",
            "Opaque spill metadata attributes/children changed",
        ));
    }
    for (a, b) in old.children.iter().zip(&new.children) {
        identical(
            old_tree,
            old_tree.node(*a)?,
            new_tree,
            new_tree.node(*b)?,
            false,
        )?;
    }
    Ok(())
}
fn extension<'a>(
    tree: &'a Tree,
    block: &Node,
    uri: &str,
    ns: &str,
    name: &str,
) -> Result<&'a Node> {
    attributes(block, &[])?;
    whitespace(tree, block)?;
    let list = tree.children(block, NS, "extLst")?;
    if list.len() != 1 {
        return Err(crate::coded(
            "spill-unsupported",
            "Unqualified spill extension ancestry",
        ));
    }
    attributes(list[0], &[])?;
    whitespace(tree, list[0])?;
    let extensions = tree.children(list[0], NS, "ext")?;
    if extensions.len() != 1 {
        return Err(crate::coded(
            "spill-unsupported",
            "Unqualified spill extension ancestry",
        ));
    }
    attributes(extensions[0], &[("uri", uri)])?;
    whitespace(tree, extensions[0])?;
    let leaves = tree.children(extensions[0], ns, name)?;
    if leaves.len() != 1 {
        return Err(crate::coded(
            "spill-unsupported",
            "Unqualified spill extension leaf",
        ));
    }
    leaf(tree, leaves[0])?;
    Ok(leaves[0])
}
fn record<'a>(tree: &'a Tree, block: &Node) -> Result<&'a Node> {
    attributes(block, &[])?;
    whitespace(tree, block)?;
    let records = tree.children(block, NS, "rc")?;
    if records.len() != 1
        || records[0].attrs.len() != 2
        || !records[0].attrs.contains_key("t")
        || !records[0].attrs.contains_key("v")
    {
        return Err(crate::coded(
            "spill-unsupported",
            "Unqualified spill metadata record",
        ));
    }
    leaf(tree, records[0])?;
    Ok(records[0])
}
fn old_container<'a>(
    tree: Option<&'a Tree>,
    name: &str,
    future: Option<&str>,
) -> Result<Option<&'a Node>> {
    let Some(tree) = tree else {
        return Ok(None);
    };
    let matches = tree.nodes[0]
        .children
        .iter()
        .map(|index| &tree.nodes[*index])
        .filter(|node| {
            node.ns == NS
                && node.name == name
                && future
                    .is_none_or(|name| node.attrs.get("name").map(String::as_str) == Some(name))
        })
        .collect::<Vec<_>>();
    if matches.len() > 1 {
        return Err(crate::coded(
            "spill-ambiguous",
            "Ambiguous original spill metadata container",
        ));
    }
    Ok(matches.first().copied())
}
fn no_unused(old: usize, new: usize, required: &BTreeSet<usize>) -> Result<()> {
    if (old..new).any(|index| !required.contains(&index)) {
        return Err(crate::coded(
            "spill-unsupported",
            "Unused appended spill metadata record",
        ));
    }
    Ok(())
}
fn rich_types(tree: &Tree) -> Result<()> {
    let root = tree.root(RICH2, "rvTypesInfo")?;
    attributes(root, &[(&format!("{{{MC}}}Ignorable"), "x")])?;
    if root.scope.get("x").map(String::as_str) != Some(NS)
        || root.scope.get("mc").map(String::as_str) != Some(MC)
    {
        return Err(crate::coded(
            "spill-unsupported",
            "Unqualified rich type namespace bindings",
        ));
    }
    whitespace(tree, root)?;
    let global = tree.children(root, RICH2, "global")?;
    if global.len() != 1 {
        return Err(crate::coded(
            "spill-unsupported",
            "Unqualified rich type global",
        ));
    }
    attributes(global[0], &[])?;
    whitespace(tree, global[0])?;
    let flags = tree.children(global[0], RICH2, "keyFlags")?;
    if flags.len() != 1 {
        return Err(crate::coded(
            "spill-unsupported",
            "Unqualified rich type flags",
        ));
    }
    attributes(flags[0], &[])?;
    whitespace(tree, flags[0])?;
    let names = [
        "_Self",
        "_DisplayString",
        "_Flags",
        "_Format",
        "_SubLabel",
        "_Attribution",
        "_Icon",
        "_Display",
        "_CanonicalPropertyNames",
        "_ClassificationId",
    ];
    let keys = tree.children(flags[0], RICH2, "key")?;
    if keys.len() != names.len() {
        return Err(crate::coded(
            "spill-unsupported",
            "Unqualified rich type keys",
        ));
    }
    for (key, name) in keys.iter().zip(names) {
        attributes(key, &[("name", name)])?;
        whitespace(tree, key)?;
        let flags = tree.children(key, RICH2, "flag")?;
        let expected = if name == "_Self" {
            vec!["ExcludeFromFile", "ExcludeFromCalcComparison"]
        } else {
            vec!["ExcludeFromCalcComparison"]
        };
        if flags.len() != expected.len() {
            return Err(crate::coded(
                "spill-unsupported",
                "Unqualified rich key flags",
            ));
        }
        for (flag, name) in flags.iter().zip(expected) {
            attributes(flag, &[("name", name), ("value", "1")])?;
            leaf(tree, flag)?;
        }
    }
    Ok(())
}
fn canonical_rect(value: &str) -> Result<super::Geometry> {
    let shape = super::Geometry::from_reference(value)?;
    if shape.reference() != value || shape.count() > u64::from(super::MAX_OUTPUTS) {
        return Err(crate::coded(
            "internal-invalid-field",
            "Noncanonical mapped spill geometry",
        ));
    }
    Ok(shape)
}

pub(crate) fn source(root: &std::path::Path) -> Result<BTreeMap<String, Vec<u8>>> {
    use std::io::Read;
    let mut zip =
        zip::ZipArchive::new(std::fs::File::open(root.join("metadata.xlsx")).map_err(err)?)
            .map_err(err)?;
    let mut parts = BTreeMap::new();
    for part in PARTS {
        let mut file = match zip.by_name(part) {
            Ok(file) => file,
            Err(zip::result::ZipError::FileNotFound) => continue,
            Err(error) => return Err(err(error)),
        };
        if file.size() > 8 * 1024 * 1024 {
            return Err(crate::coded(
                "spill-limit",
                "Spill metadata source part exceeds bound",
            ));
        }
        let mut bytes = Vec::new();
        file.read_to_end(&mut bytes).map_err(err)?;
        parts.insert(part.into(), bytes);
    }
    Ok(parts)
}
pub(crate) fn check(root: &std::path::Path, input: &Value) -> Result<Value> {
    let fields = input["plan"]["replacements"].as_object().ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing spill metadata replacements",
        )
    })?;
    let replacements = fields
        .iter()
        .map(|(part, xml)| {
            Ok((
                part.clone(),
                xml.as_str()
                    .ok_or_else(|| crate::coded("spill-malformed", "Invalid spill metadata XML"))?
                    .to_owned(),
            ))
        })
        .collect::<Result<BTreeMap<_, _>>>()?;
    validate(&source(root)?, &replacements, &input["plan"]["masters"])?;
    Ok(serde_json::json!({"valid":true}))
}
/// Admit the calculation owner's exact calcChain graph removal before the
/// independent spill append proof. Every other relationship/type stays equal.
pub(crate) fn calculation_base(
    before: &BTreeMap<String, Vec<u8>>,
    replacements: &Value,
    removed: &Value,
) -> Result<BTreeMap<String, Vec<u8>>> {
    let mut base = before.clone();
    let remove_chain = removed
        .as_array()
        .is_some_and(|names| names.iter().any(|name| name == "xl/calcChain.xml"));
    for part in [PARTS[4], PARTS[5]] {
        let Some(xml) = replacements.get(part).and_then(Value::as_str) else {
            continue;
        };
        let old = Tree::parse(before.get(part).ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing source calculation graph")
        })?)?;
        let new = Tree::parse(xml.as_bytes())?;
        let a = &old.nodes[0];
        let b = &new.nodes[0];
        if a.ns != b.ns
            || a.name != b.name
            || a.attrs != b.attrs
            || a.text != b.text
            || a.tail != b.tail
            || a.scope
                .iter()
                .any(|(key, value)| b.scope.get(key) != Some(value))
        {
            return Err(crate::coded(
                "spill-stale",
                "Calculation graph root changed outside calcChain normalization",
            ));
        }
        let is_chain = |node: &Node| {
            if part == PARTS[4] {
                node.ns == REL
                    && node.name == "Relationship"
                    && node.attrs.get("Type").map(String::as_str)
                        == Some(
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",
                        )
                    && node.attrs.get("Target").is_some_and(|target| {
                        ["calcChain.xml", "/xl/calcChain.xml"].contains(&target.as_str())
                    })
                    && node
                        .attrs
                        .get("TargetMode")
                        .is_none_or(|mode| mode == "Internal")
            } else {
                node.ns == CT
                    && node.name == "Override"
                    && node.attrs.get("PartName").map(String::as_str) == Some("/xl/calcChain.xml")
                    && node.attrs.get("ContentType").map(String::as_str)
                        == Some(
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
                        )
            }
        };
        let originals = a
            .children
            .iter()
            .map(|index| &old.nodes[*index])
            .filter(|node| !remove_chain || !is_chain(node))
            .collect::<Vec<_>>();
        if originals.len() != b.children.len()
            || a.children
                .iter()
                .filter(|index| is_chain(&old.nodes[**index]))
                .count()
                > 1
        {
            return Err(crate::coded(
                "spill-bad-input",
                "Calculation graph delta is not the exact unique calcChain removal",
            ));
        }
        for (node, index) in originals.iter().zip(&b.children) {
            identical(&old, node, &new, &new.nodes[*index], false)?;
        }
        base.insert(part.into(), xml.as_bytes().to_vec());
    }
    Ok(base)
}

pub(crate) fn validate(
    before: &BTreeMap<String, Vec<u8>>,
    replacements: &BTreeMap<String, String>,
    masters: &Value,
) -> Result<()> {
    let request = masters
        .as_object()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing native spill master map"))?;
    let mut after = before.clone();
    for (part, xml) in replacements {
        if !PARTS.contains(&part.as_str()) {
            return Err(crate::coded(
                "spill-unsupported",
                "Unqualified spill metadata replacement part",
            ));
        }
        after.insert(part.clone(), xml.as_bytes().to_vec());
    }
    if request.is_empty() {
        return if replacements.is_empty() {
            Ok(())
        } else {
            Err(crate::coded(
                "spill-refused",
                "Unrequested spill metadata delta without active owners",
            ))
        };
    }
    let mut old_trees = BTreeMap::new();
    let mut trees = BTreeMap::new();
    for (part, bytes) in before {
        if !PARTS.contains(&part.as_str()) {
            return Err(crate::coded(
                "spill-malformed",
                "Unexpected spill metadata source part",
            ));
        }
        old_trees.insert(part.clone(), Tree::parse(bytes)?);
    }
    for (part, bytes) in &after {
        trees.insert(part.clone(), Tree::parse(bytes)?);
    }
    let metadata = trees
        .get(PARTS[0])
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing spill metadata part"))?;
    let root = metadata.root(NS, "metadata")?;
    let old_metadata = old_trees.get(PARTS[0]);
    for (part, old) in &old_trees {
        let new = trees.get(part).ok_or_else(|| {
            crate::coded("spill-refused", "Spill metadata part removal is refused")
        })?;
        if part == PARTS[3] {
            if before[part] != after[part] {
                return Err(crate::coded(
                    "spill-unsupported",
                    "Existing rich type part must remain byte-identical",
                ));
            }
            continue;
        }
        if part == PARTS[0] {
            let old_root = old.root(NS, "metadata")?;
            if old_root.attrs != root.attrs
                || old_root.text != root.text
                || old_root.tail != root.tail
                || old_root
                    .scope
                    .iter()
                    .any(|(key, value)| root.scope.get(key) != Some(value))
            {
                return Err(crate::coded(
                    "spill-stale",
                    "Original metadata root changed",
                ));
            }
            let mut last = None;
            for index in &old_root.children {
                let child = old.node(*index)?;
                let matches = root
                    .children
                    .iter()
                    .enumerate()
                    .filter(|(_, index)| {
                        let node = &metadata.nodes[**index];
                        node.ns == child.ns
                            && node.name == child.name
                            && node.attrs.get("name") == child.attrs.get("name")
                    })
                    .collect::<Vec<_>>();
                if matches.len() != 1 || last.is_some_and(|position| position >= matches[0].0) {
                    return Err(crate::coded(
                        "spill-stale",
                        "Original metadata container order changed",
                    ));
                }
                last = Some(matches[0].0);
                let append = child.ns == NS
                    && (["metadataTypes", "cellMetadata", "valueMetadata"]
                        .contains(&child.name.as_str())
                        || child.name == "futureMetadata"
                            && child.attrs.get("name").is_some_and(|name| {
                                ["XLDAPR", "XLRICHVALUE"].contains(&name.as_str())
                            }));
                identical(old, child, metadata, metadata.node(*matches[0].1)?, append)?;
            }
        } else {
            identical(old, &old.nodes[0], new, &new.nodes[0], true)?;
        }
    }
    if root.children.iter().any(|index| {
        let node = &metadata.nodes[*index];
        node.ns != NS
            || ![
                "metadataTypes",
                "futureMetadata",
                "cellMetadata",
                "valueMetadata",
                "extLst",
            ]
            .contains(&node.name.as_str())
    }) {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown spill metadata root container",
        ));
    }
    let type_node = metadata.node(metadata.one(root, NS, "metadataTypes", true)?.unwrap())?;
    let types = metadata.counted(type_node, NS, "metadataType")?;
    let mut names = BTreeSet::new();
    for kind in &types {
        if attr(kind, "name")?.is_empty() || !names.insert(attr(kind, "name")?.to_owned()) {
            return Err(crate::coded(
                "spill-ambiguous",
                "Ambiguous spill metadata type names",
            ));
        }
    }
    let mut future = BTreeMap::new();
    for index in &root.children {
        let node = metadata.node(*index)?;
        if node.name == "futureMetadata" {
            let name = attr(node, "name")?;
            if name.is_empty()
                || future
                    .insert(name.to_owned(), metadata.counted(node, NS, "bk")?)
                    .is_some()
            {
                return Err(crate::coded(
                    "spill-ambiguous",
                    "Ambiguous spill future metadata",
                ));
            }
        }
    }
    let cells = metadata.counted(
        metadata.node(metadata.one(root, NS, "cellMetadata", true)?.unwrap())?,
        NS,
        "bk",
    )?;
    let values = metadata
        .one(root, NS, "valueMetadata", false)?
        .map(|index| metadata.counted(&metadata.nodes[index], NS, "bk"))
        .transpose()?
        .unwrap_or_default();
    for block in cells.iter().chain(&values) {
        for rc in metadata.children(block, NS, "rc")? {
            let kind = types
                .get(integer(attr(rc, "t")?, true)?)
                .ok_or_else(|| crate::coded("spill-malformed", "Invalid metadata type index"))?;
            if integer(attr(rc, "v")?, false)?
                >= future
                    .get(attr(kind, "name")?)
                    .ok_or_else(|| {
                        crate::coded(
                            "internal-missing-field",
                            "Missing metadata future container",
                        )
                    })?
                    .len()
            {
                return Err(crate::coded(
                    "spill-malformed",
                    "Invalid metadata future index",
                ));
            }
        }
    }
    let relations = trees.get(PARTS[4]).ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing workbook metadata relationships",
        )
    })?;
    let rels = relations.children(relations.root(REL, "Relationships")?, REL, "Relationship")?;
    let mut ids = BTreeSet::new();
    for rel in &rels {
        if attr(rel, "Id")?.is_empty() || !ids.insert(attr(rel, "Id")?.to_owned()) {
            return Err(crate::coded(
                "spill-duplicate",
                "Duplicate/empty metadata relationship identity",
            ));
        }
    }
    let content = trees
        .get(PARTS[5])
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing metadata content types"))?;
    let content_root = content.root(CT, "Types")?;
    let mut override_names = BTreeSet::new();
    for index in &content_root.children {
        let node = content.node(*index)?;
        if node.ns != CT || !["Default", "Override"].contains(&node.name.as_str()) {
            return Err(crate::coded(
                "spill-unsupported",
                "Unqualified metadata content type child",
            ));
        }
        if node.name == "Override" && !override_names.insert(attr(node, "PartName")?.to_owned()) {
            return Err(crate::coded(
                "spill-duplicate",
                "Duplicate metadata content type part",
            ));
        }
    }
    let mut required_parts = BTreeSet::new();
    if !request.is_empty() {
        required_parts.insert(PARTS[0]);
    }
    if request.values().any(|master| master.get("vm").is_some()) {
        required_parts.extend(PARTS[1..4].iter().copied());
    }
    let added = required_parts
        .iter()
        .filter(|part| !before.contains_key(**part))
        .copied()
        .collect::<BTreeSet<_>>();
    if after
        .keys()
        .filter(|part| !before.contains_key(*part))
        .map(String::as_str)
        .collect::<BTreeSet<_>>()
        != added
    {
        return Err(crate::coded(
            "spill-unsupported",
            "Unused/unrequested new spill metadata part",
        ));
    }
    let old_rel_count = old_trees
        .get(PARTS[4])
        .ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing original workbook relationships",
            )
        })?
        .nodes[0]
        .children
        .len();
    let old_type_count = old_trees
        .get(PARTS[5])
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing original content types"))?
        .nodes[0]
        .children
        .len();
    if rels.len() != old_rel_count + added.len()
        || content_root.children.len() != old_type_count + added.len()
    {
        return Err(crate::coded(
            "spill-malformed",
            "Unexpected spill package graph delta",
        ));
    }
    for rel in &rels[old_rel_count..] {
        if !GRAPH.iter().any(|(part, target, typ, _)| {
            added.contains(part)
                && rel.attrs
                    == [
                        ("Id".into(), attr(rel, "Id").unwrap_or("").into()),
                        ("Type".into(), (*typ).into()),
                        ("Target".into(), (*target).into()),
                    ]
                    .into_iter()
                    .collect()
        }) {
            return Err(crate::coded(
                "spill-unsupported",
                "Unqualified new spill relationship",
            ));
        }
    }
    for index in &content_root.children[old_type_count..] {
        let node = content.node(*index)?;
        if !GRAPH.iter().any(|(part, _, _, typ)| {
            added.contains(part)
                && node.name == "Override"
                && node.attrs
                    == [
                        ("PartName".into(), format!("/{part}")),
                        ("ContentType".into(), (*typ).into()),
                    ]
                    .into_iter()
                    .collect()
        }) {
            return Err(crate::coded(
                "spill-unsupported",
                "Unqualified new spill content type",
            ));
        }
    }
    for (part, target, typ, content_type) in GRAPH {
        if after.contains_key(part) {
            if rels
                .iter()
                .filter(|r| {
                    r.attrs.get("Type").map(String::as_str) == Some(typ)
                        && r.attrs.get("Target").map(String::as_str) == Some(target)
                        && r.attrs
                            .get("TargetMode")
                            .is_none_or(|mode| mode == "Internal")
                })
                .count()
                != 1
                || content_root
                    .children
                    .iter()
                    .map(|index| &content.nodes[*index])
                    .filter(|n| {
                        n.name == "Override"
                            && n.attrs.get("PartName") == Some(&format!("/{part}"))
                            && n.attrs.get("ContentType").map(String::as_str) == Some(content_type)
                    })
                    .count()
                    != 1
            {
                return Err(crate::coded(
                    "spill-unsupported",
                    "Spill part lacks a unique qualified package graph",
                ));
            }
        }
    }
    let rich_values = trees
        .get(PARTS[1])
        .map(|tree| tree.counted(tree.root(RICH, "rvData")?, RICH, "rv"))
        .transpose()?
        .unwrap_or_default();
    let rich_structures = trees
        .get(PARTS[2])
        .map(|tree| tree.counted(tree.root(RICH, "rvStructures")?, RICH, "s"))
        .transpose()?
        .unwrap_or_default();
    if !rich_values.is_empty() {
        let value_tree = &trees[PARTS[1]];
        let structure_tree = &trees[PARTS[2]];
        for value in &rich_values {
            let structure = rich_structures
                .get(integer(attr(value, "s")?, false)?)
                .ok_or_else(|| crate::coded("spill-malformed", "Invalid rich structure index"))?;
            if value_tree.children(value, RICH, "v")?.len()
                != structure_tree.children(structure, RICH, "k")?.len()
            {
                return Err(crate::coded(
                    "internal-mismatch",
                    "Rich value field/structure count mismatch",
                ));
            }
        }
        let rich_types_tree = trees.get(PARTS[3]).ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing rich value type part")
        })?;
        if !before.contains_key(PARTS[3]) {
            rich_types(rich_types_tree)?;
        } else {
            rich_types_tree.root(RICH2, "rvTypesInfo")?;
        }
    }
    let mut needed: BTreeMap<&str, BTreeSet<usize>> = [
        "types",
        "cells",
        "values",
        "XLDAPR",
        "XLRICHVALUE",
        "richValues",
        "richStructures",
    ]
    .into_iter()
    .map(|name| (name, BTreeSet::new()))
    .collect();
    for (key, master) in request {
        let fields = master
            .as_object()
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid mapped spill master"))?;
        if fields
            .keys()
            .any(|key| !["cm", "vm", "range", "intendedRange", "aca", "ca"].contains(&key.as_str()))
        {
            return Err(crate::coded(
                "spill-malformed",
                "Unexpected mapped spill master field",
            ));
        }
        let (sheet, address) = key.split_once('\0').ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing mapped master sheet/address",
            )
        })?;
        if !sheet.starts_with("xl/worksheets/")
            || !sheet.ends_with(".xml")
            || sheet["xl/worksheets/".len()..].contains('/')
            || address.contains('\0')
        {
            return Err(crate::coded(
                "spill-malformed",
                "Invalid mapped spill worksheet part",
            ));
        }
        let origin = canonical_rect(address)?;
        if origin.count() != 1 {
            return Err(crate::coded(
                "spill-bad-input",
                "Mapped spill master is not a point",
            ));
        }
        let range = canonical_rect(master["range"].as_str().ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing mapped spill range")
        })?)?;
        if range.first != origin.first {
            return Err(crate::coded("spill-stale", "Mapped spill origin changed"));
        }
        let blocked = master.get("vm").is_some();
        if blocked {
            if range != origin || master["aca"] != "1" || master["ca"] != "1" {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid mapped blocked spill flags/range",
                ));
            }
        } else if ["intendedRange", "aca", "ca"]
            .iter()
            .any(|key| fields.contains_key(*key))
        {
            return Err(crate::coded(
                "spill-refused",
                "Calculated spill has blocked metadata",
            ));
        }
        let cell_index = json_index(&master["cm"], true)?;
        let rc = record(
            metadata,
            cells.get(cell_index).ok_or_else(|| {
                crate::coded(
                    "spill-unsupported",
                    "Mapped cm index is outside cellMetadata",
                )
            })?,
        )?;
        let type_index = integer(attr(rc, "t")?, true)?;
        let kind = types.get(type_index).ok_or_else(|| {
            crate::coded("internal-invalid-field", "Invalid mapped metadata type")
        })?;
        if attr(kind, "name")? != "XLDAPR" || attr(kind, "cellMeta")? != "1" {
            return Err(crate::coded(
                "spill-bad-input",
                "Mapped cm is not dynamic cell metadata",
            ));
        }
        let future_index = integer(attr(rc, "v")?, false)?;
        let flag = extension(
            metadata,
            future
                .get("XLDAPR")
                .and_then(|items| items.get(future_index))
                .ok_or_else(|| {
                    crate::coded(
                        "internal-missing-field",
                        "Mapped XLDAPR future index is missing",
                    )
                })?,
            "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}",
            XDA,
            "dynamicArrayProperties",
        )?;
        if flag
            .attrs
            .keys()
            .any(|key| !["fDynamic", "fCollapsed"].contains(&key.as_str()))
            || attr(flag, "fDynamic")? != "1"
            || flag.attrs.get("fCollapsed").is_some_and(|flag| flag != "0")
        {
            return Err(crate::coded(
                "spill-unsupported",
                "Mapped dynamic-array flags are unqualified",
            ));
        }
        needed.get_mut("types").unwrap().insert(type_index);
        needed.get_mut("cells").unwrap().insert(cell_index);
        needed.get_mut("XLDAPR").unwrap().insert(future_index);
        if blocked {
            let intended = canonical_rect(master["intendedRange"].as_str().ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing intended spill geometry")
            })?)?;
            if intended.first != origin.first {
                return Err(crate::coded("spill-stale", "Intended spill origin changed"));
            }
            let value_index = json_index(&master["vm"], true)?;
            let rc = record(
                metadata,
                values.get(value_index).ok_or_else(|| {
                    crate::coded(
                        "spill-unsupported",
                        "Mapped vm index is outside valueMetadata",
                    )
                })?,
            )?;
            let type_index = integer(attr(rc, "t")?, true)?;
            let kind = types.get(type_index).ok_or_else(|| {
                crate::coded(
                    "internal-invalid-field",
                    "Invalid mapped value metadata type",
                )
            })?;
            if attr(kind, "name")? != "XLRICHVALUE"
                || kind.attrs.get("cellMeta").is_some_and(|flag| flag != "0")
            {
                return Err(crate::coded(
                    "spill-unsupported",
                    "Mapped vm is not qualified value metadata",
                ));
            }
            let future_index = integer(attr(rc, "v")?, false)?;
            let value = extension(
                metadata,
                future
                    .get("XLRICHVALUE")
                    .and_then(|items| items.get(future_index))
                    .ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing rich future record")
                    })?,
                "{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}",
                RICH,
                "rvb",
            )?;
            if value.attrs.len() != 1 {
                return Err(crate::coded(
                    "spill-unsupported",
                    "Extended rich future record",
                ));
            }
            let rich_index = integer(attr(value, "i")?, false)?;
            let value = rich_values.get(rich_index).ok_or_else(|| {
                crate::coded(
                    "internal-missing-field",
                    "Mapped rich value index is missing",
                )
            })?;
            if value.attrs.len() != 1 {
                return Err(crate::coded(
                    "spill-unsupported",
                    "Extended rich value attributes",
                ));
            }
            let structure_index = integer(attr(value, "s")?, false)?;
            let structure = rich_structures.get(structure_index).ok_or_else(|| {
                crate::coded(
                    "internal-missing-field",
                    "Mapped rich structure index is missing",
                )
            })?;
            attributes(structure, &[("t", "_error")])?;
            let structure_tree = &trees[PARTS[2]];
            let value_tree = &trees[PARTS[1]];
            whitespace(structure_tree, structure)?;
            let keys = structure_tree.children(structure, RICH, "k")?;
            let values = value_tree.children(value, RICH, "v")?;
            let expected = ["colOffset", "errorType", "rwOffset", "subType"];
            if keys.len() != 4 || values.len() != 4 || !value.text.trim().is_empty() {
                return Err(crate::coded(
                    "spill-unsupported",
                    "Unqualified rich error structure",
                ));
            }
            let mut numbers = Vec::new();
            for ((key, value), name) in keys.iter().zip(&values).zip(expected) {
                attributes(key, &[("n", name), ("t", "i")])?;
                leaf(structure_tree, key)?;
                attributes(value, &[])?;
                if !value.children.is_empty() || !value.tail.trim().is_empty() {
                    return Err(crate::coded(
                        "spill-unsupported",
                        "Extended rich integer value",
                    ));
                }
                numbers.push(integer(&value.text, false)?);
            }
            if numbers
                != vec![
                    (intended.last.1 - intended.first.1) as usize,
                    8,
                    (intended.last.0 - intended.first.0) as usize,
                    1,
                ]
            {
                return Err(crate::coded(
                    "internal-mismatch",
                    "Rich SPILL error/offsets differ from intended geometry",
                ));
            }
            for (name, index) in [
                ("types", type_index),
                ("values", value_index),
                ("XLRICHVALUE", future_index),
                ("richValues", rich_index),
                ("richStructures", structure_index),
            ] {
                needed.get_mut(name).unwrap().insert(index);
            }
        }
    }
    for (name, items, required) in [
        ("metadataTypes", &types, "types"),
        ("cellMetadata", &cells, "cells"),
        ("valueMetadata", &values, "values"),
    ] {
        let old = old_container(old_metadata, name, None)?.map_or(0, |node| node.children.len());
        no_unused(old, items.len(), &needed[required])?;
    }
    let old_types =
        old_container(old_metadata, "metadataTypes", None)?.map_or(0, |node| node.children.len());
    for kind in &types[old_types..] {
        let name = attr(kind, "name")?;
        if !["XLDAPR", "XLRICHVALUE"].contains(&name) {
            return Err(crate::coded(
                "spill-unsupported",
                "Unqualified appended metadata type",
            ));
        }
        let mut expected = BTreeMap::from([
            ("name".into(), name.into()),
            ("minSupportedVersion".into(), "120000".into()),
        ]);
        for flag in [
            "copy",
            "pasteAll",
            "pasteValues",
            "merge",
            "splitFirst",
            "rowColShift",
            "clearFormats",
            "clearComments",
            "assign",
            "coerce",
        ] {
            expected.insert(flag.into(), "1".into());
        }
        if name == "XLDAPR" {
            expected.insert("cellMeta".into(), "1".into());
        }
        if kind.attrs != expected {
            return Err(crate::coded(
                "spill-unsupported",
                "Unqualified appended metadata type attributes",
            ));
        }
        leaf(metadata, kind)?;
    }
    for (name, blocks) in &future {
        let old = old_container(old_metadata, "futureMetadata", Some(name))?;
        no_unused(
            old.map_or(0, |node| node.children.len()),
            blocks.len(),
            needed.get(name.as_str()).unwrap_or(&BTreeSet::new()),
        )?;
        if old.is_none()
            && (!["XLDAPR", "XLRICHVALUE"].contains(&name.as_str()) || blocks.is_empty())
        {
            return Err(crate::coded(
                "spill-unsupported",
                "Unused added future metadata container",
            ));
        }
    }
    let old_keys = old_metadata
        .map(|tree| {
            tree.nodes[0]
                .children
                .iter()
                .map(|index| {
                    let node = &tree.nodes[*index];
                    (
                        node.ns.clone(),
                        node.name.clone(),
                        node.attrs.get("name").cloned(),
                    )
                })
                .collect::<BTreeSet<_>>()
        })
        .unwrap_or_default();
    for index in &root.children {
        let node = metadata.node(*index)?;
        if old_keys.contains(&(
            node.ns.clone(),
            node.name.clone(),
            node.attrs.get("name").cloned(),
        )) {
            continue;
        }
        if node.name == "extLst" || node.children.is_empty() {
            return Err(crate::coded(
                "spill-unsupported",
                "Unused added metadata root container",
            ));
        }
        let mut expected = BTreeMap::from([("count".into(), node.children.len().to_string())]);
        if node.name == "futureMetadata" {
            expected.insert("name".into(), attr(node, "name")?.into());
        }
        if node.attrs != expected {
            return Err(crate::coded(
                "spill-unsupported",
                "Unqualified added metadata container attributes",
            ));
        }
        whitespace(metadata, node)?;
    }
    if old_metadata.is_none() {
        attributes(root, &[])?;
        whitespace(metadata, root)?;
    }
    for (part, items, required) in [
        (PARTS[1], &rich_values, "richValues"),
        (PARTS[2], &rich_structures, "richStructures"),
    ] {
        let old = old_trees
            .get(part)
            .map_or(0, |tree| tree.nodes[0].children.len());
        no_unused(old, items.len(), &needed[required])?;
        if !before.contains_key(part) && after.contains_key(part) {
            let tree = &trees[part];
            if items.is_empty() {
                return Err(crate::coded(
                    "spill-unsupported",
                    "Unused new rich metadata part",
                ));
            }
            attributes(&tree.nodes[0], &[("count", &items.len().to_string())])?;
            whitespace(tree, &tree.nodes[0])?;
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;
    use std::io::Read;
    fn retained() -> BTreeMap<String, Vec<u8>> {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../fixtures/excel-dynamic-reference.xlsx");
        let mut zip = zip::ZipArchive::new(std::fs::File::open(path).unwrap()).unwrap();
        let mut parts = BTreeMap::new();
        for part in PARTS {
            let mut bytes = Vec::new();
            zip.by_name(part).unwrap().read_to_end(&mut bytes).unwrap();
            parts.insert(part.into(), bytes);
        }
        parts
    }
    fn mappings() -> Value {
        json!({"xl/worksheets/sheet1.xml\u{0}A1":{"cm":1,"range":"A1:B3"},"xl/worksheets/sheet1.xml\u{0}A41":{"cm":1,"vm":2,"range":"A41","intendedRange":"A41:B42","aca":"1","ca":"1"}})
    }
    #[test]
    fn actual_retained_metadata_has_exact_dynamic_and_rich_spill_mapping() {
        validate(&retained(), &BTreeMap::new(), &mappings()).unwrap();
    }
    #[test]
    fn namespace_uri_indices_flags_offsets_and_unused_delta_are_not_qualification() {
        let before = retained();
        let source = std::str::from_utf8(&before[PARTS[0]]).unwrap();
        for (from, to) in [
            ("{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}", "{wrong-uri}"),
            ("{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}", "{wrong-uri}"),
            ("fDynamic=\"1\"", "fDynamic=\"0\""),
            (
                "</metadata>",
                "<futureMetadata name=\"OTHER\" count=\"0\"/></metadata>",
            ),
        ] {
            assert!(source.contains(from));
            assert!(
                validate(
                    &before,
                    &BTreeMap::from([(PARTS[0].into(), source.replace(from, to))]),
                    &mappings()
                )
                .is_err()
            );
        }
        for (field, value) in [
            ("cm", json!(0)),
            ("vm", json!(0)),
            ("intendedRange", json!("A41:C42")),
            ("intendedRange", json!("B41:C42")),
            ("aca", json!("0")),
        ] {
            let mut map = mappings();
            map["xl/worksheets/sheet1.xml\u{0}A41"][field] = value;
            assert!(validate(&before, &BTreeMap::new(), &map).is_err());
        }
    }
    #[test]
    fn opaque_mixed_text_and_namespace_bindings_are_preserved() {
        let mut before = retained();
        let text=std::str::from_utf8(&before[PARTS[0]]).unwrap().replace("</metadata>","<extLst xmlns:q=\"urn:original\"><q:opaque> leading <q:inner/> trailing </q:opaque></extLst></metadata>");
        before.insert(PARTS[0].into(), text.as_bytes().to_vec());
        validate(&before, &BTreeMap::new(), &mappings()).unwrap();
        for changed in [
            text.replace(" leading ", "leading"),
            text.replace(" trailing ", "trailing"),
            text.replace("urn:original", "urn:changed"),
        ] {
            assert!(
                validate(
                    &before,
                    &BTreeMap::from([(PARTS[0].into(), changed)]),
                    &mappings()
                )
                .is_err()
            );
        }
    }
}
