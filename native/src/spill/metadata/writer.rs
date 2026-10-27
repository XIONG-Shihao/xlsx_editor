//! Append records without reserializing the source XML. Admission remains in
//! the parent module; this producer cannot qualify spill ownership.
use super::*;
use serde_json::json;
use std::ops::Range;

const DYNAMIC_URI: &str = "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}";
const RICH_URI: &str = "{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}";
const ERROR_FIELDS: [&str; 4] = ["colOffset", "errorType", "rwOffset", "subType"];

struct Span {
    start: usize,
    end: usize,
    closing: Option<usize>,
    qname: String,
    count: Option<Range<usize>>,
}

struct Document {
    xml: String,
    tree: Tree,
    spans: Vec<Span>,
    appended: BTreeMap<usize, (bool, Vec<String>)>,
    inserted: Vec<(usize, String)>,
}

impl Document {
    fn parse(bytes: &[u8], ns: &str, name: &str) -> Result<Self> {
        let tree = Tree::parse(bytes)?;
        tree.root(ns, name)?;
        let xml = std::str::from_utf8(bytes).map_err(err)?.to_owned();
        // The semantic parser has already rejected entities, comments, unknown
        // prefixes and malformed nesting. This pass records only byte offsets.
        use quick_xml::{Reader, events::Event};
        let mut reader = Reader::from_reader(bytes);
        let mut spans: Vec<Span> = Vec::new();
        let mut stack = Vec::new();
        loop {
            crate::cancel::check()?;
            let start = reader.buffer_position() as usize;
            let event = reader.read_event().map_err(err)?;
            let end = reader.buffer_position() as usize;
            match event {
                Event::Start(ref element) | Event::Empty(ref element) => {
                    let mut count = None;
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(err)?;
                        if attribute.key.as_ref() == b"count" {
                            let value = attribute.value.as_ref();
                            let at = (value.as_ptr() as usize)
                                .checked_sub(bytes.as_ptr() as usize)
                                .ok_or_else(|| {
                                    crate::coded(
                                        "spill-malformed",
                                        "Invalid metadata attribute span",
                                    )
                                })?;
                            let finish = at.checked_add(value.len()).ok_or_else(|| {
                                crate::coded("spill-limit", "Metadata span overflow")
                            })?;
                            if at < start || finish > end || bytes.get(at..finish) != Some(value) {
                                return Err(crate::coded(
                                    "spill-malformed",
                                    "Invalid metadata attribute span",
                                ));
                            }
                            count = Some(at..finish);
                        }
                    }
                    let index = spans.len();
                    spans.push(Span {
                        start,
                        end,
                        closing: None,
                        qname: std::str::from_utf8(element.name().as_ref())
                            .map_err(err)?
                            .into(),
                        count,
                    });
                    if matches!(event, Event::Start(_)) {
                        stack.push(index);
                    }
                }
                Event::End(_) => {
                    let index = stack.pop().ok_or_else(|| {
                        crate::coded("spill-malformed", "Unbalanced metadata spans")
                    })?;
                    spans[index].closing = Some(start);
                    spans[index].end = end;
                }
                Event::Eof => break,
                _ => {}
            }
        }
        if spans.len() != tree.nodes.len() || !stack.is_empty() {
            return Err(crate::coded(
                "internal-mismatch",
                "Metadata semantic/span indexes differ",
            ));
        }
        Ok(Self {
            xml,
            tree,
            spans,
            appended: BTreeMap::new(),
            inserted: Vec::new(),
        })
    }

    fn load(
        before: &BTreeMap<String, Vec<u8>>,
        part: &str,
        ns: &str,
        name: &str,
        counted: bool,
    ) -> Result<Self> {
        let fresh = format!(
            "<{name} xmlns=\"{ns}\"{}/>",
            if counted { " count=\"0\"" } else { "" }
        );
        Self::parse(
            before
                .get(part)
                .map(Vec::as_slice)
                .unwrap_or(fresh.as_bytes()),
            ns,
            name,
        )
    }

    fn append(&mut self, index: usize, item: String, counted: bool) -> Result<usize> {
        let count = self.tree.node(index)?.children.len();
        let (previous_counted, items) = self.appended.entry(index).or_insert((counted, Vec::new()));
        if *previous_counted != counted {
            return Err(crate::coded(
                "spill-unsupported",
                "Inconsistent metadata container append",
            ));
        }
        let result = count + items.len();
        items.push(item);
        Ok(result)
    }

    fn finish(self) -> Result<String> {
        let mut edits: Vec<(Range<usize>, String)> = self
            .inserted
            .into_iter()
            .map(|(at, xml)| (at..at, xml))
            .collect();
        for (index, (counted, items)) in self.appended {
            let span = &self.spans[index];
            if counted {
                let count = span.count.clone().ok_or_else(|| {
                    crate::coded("spill-bad-input", "Counted metadata container has no count")
                })?;
                edits.push((
                    count,
                    (self.tree.nodes[index].children.len() + items.len()).to_string(),
                ));
            }
            let added = items.concat();
            if let Some(at) = span.closing {
                edits.push((at..at, added));
            } else {
                if self
                    .xml
                    .as_bytes()
                    .get(span.end.saturating_sub(2)..span.end)
                    != Some(b"/>")
                {
                    return Err(crate::coded(
                        "spill-malformed",
                        "Invalid empty metadata element span",
                    ));
                }
                edits.push((span.end - 2..span.end, format!(">{added}</{}>", span.qname)));
            }
        }
        edits.sort_by_key(|(span, _)| span.start);
        let mut result = String::new();
        let mut previous = 0;
        for (span, replacement) in edits {
            if span.start < previous || span.end > self.xml.len() {
                return Err(crate::coded(
                    "spill-conflict",
                    "Overlapping metadata append edits",
                ));
            }
            result.push_str(&self.xml[previous..span.start]);
            result.push_str(&replacement);
            previous = span.end;
        }
        result.push_str(&self.xml[previous..]);
        Ok(result)
    }
}

#[derive(Clone, PartialEq, Eq, PartialOrd, Ord)]
enum ContainerKey {
    Types,
    Future(String),
    Cells,
    Values,
}
impl ContainerKey {
    fn name(&self) -> &str {
        match self {
            Self::Types => "metadataTypes",
            Self::Future(_) => "futureMetadata",
            Self::Cells => "cellMetadata",
            Self::Values => "valueMetadata",
        }
    }
    fn rank(&self) -> usize {
        match self {
            Self::Types => 0,
            Self::Future(_) => 1,
            Self::Cells => 2,
            Self::Values => 3,
        }
    }
    fn child(&self) -> &str {
        if matches!(self, Self::Types) {
            "metadataType"
        } else {
            "bk"
        }
    }
}
struct Container {
    original: Option<usize>,
    children: Vec<usize>,
    added: Vec<String>,
}
struct Metadata {
    document: Document,
    containers: BTreeMap<ContainerKey, Container>,
    records: BTreeMap<(ContainerKey, usize, usize), usize>,
}
impl Metadata {
    fn new(before: &BTreeMap<String, Vec<u8>>) -> Result<Self> {
        Ok(Self {
            document: Document::load(before, PARTS[0], NS, "metadata", false)?,
            containers: BTreeMap::new(),
            records: BTreeMap::new(),
        })
    }
    fn ensure(&mut self, key: &ContainerKey) -> Result<()> {
        if self.containers.contains_key(key) {
            return Ok(());
        }
        let tree = &self.document.tree;
        let matches = tree.nodes[0]
            .children
            .iter()
            .copied()
            .filter(|index| {
                let node = &tree.nodes[*index];
                node.ns == NS
                    && node.name == key.name()
                    && match key {
                        ContainerKey::Future(name) => node.attrs.get("name") == Some(name),
                        _ => true,
                    }
            })
            .collect::<Vec<_>>();
        if matches.len() > 1 {
            return Err(crate::coded(
                "spill-ambiguous",
                "Ambiguous metadata append container",
            ));
        }
        let original = matches.first().copied();
        let children = if let Some(index) = original {
            tree.counted(&tree.nodes[index], NS, key.child())?;
            tree.nodes[index].children.clone()
        } else {
            Vec::new()
        };
        self.containers.insert(
            key.clone(),
            Container {
                original,
                children,
                added: Vec::new(),
            },
        );
        Ok(())
    }
    fn children(&mut self, key: &ContainerKey) -> Result<Vec<usize>> {
        self.ensure(key)?;
        Ok(self.containers[key].children.clone())
    }
    fn push(&mut self, key: &ContainerKey, item: String) -> Result<usize> {
        self.ensure(key)?;
        let container = self.containers.get_mut(key).ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing metadata append container",
            )
        })?;
        let index = container.children.len() + container.added.len();
        container.added.push(item);
        Ok(index)
    }
    fn kind(&mut self, name: &str, cell: bool) -> Result<usize> {
        let key = ContainerKey::Types;
        let children = self.children(&key)?;
        let matches = children
            .iter()
            .enumerate()
            .filter(|(_, index)| {
                self.document.tree.nodes[**index]
                    .attrs
                    .get("name")
                    .map(String::as_str)
                    == Some(name)
            })
            .collect::<Vec<_>>();
        if matches.len() > 1 {
            return Err(crate::coded(
                "spill-ambiguous",
                "Ambiguous spill metadata type",
            ));
        }
        if let Some((position, index)) = matches.first() {
            let kind = &self.document.tree.nodes[**index];
            if cell && kind.attrs.get("cellMeta").map(String::as_str) != Some("1")
                || !cell && kind.attrs.get("cellMeta").is_some_and(|flag| flag != "0")
            {
                return Err(crate::coded(
                    "spill-unsupported",
                    "Existing spill type has incompatible metadata role",
                ));
            }
            return Ok(*position + 1);
        }
        let flags = [
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
        ]
        .into_iter()
        .map(|flag| format!(" {flag}=\"1\""))
        .collect::<String>();
        Ok(self.push(&key, format!("<metadataType xmlns=\"{NS}\" name=\"{name}\" minSupportedVersion=\"120000\"{flags}{}/>", if cell { " cellMeta=\"1\"" } else { "" }))? + 1)
    }
    fn record(&mut self, key: &ContainerKey, kind: usize, future: usize) -> Result<usize> {
        let cache_key = (key.clone(), kind, future);
        if let Some(index) = self.records.get(&cache_key) {
            return Ok(*index);
        }
        let mut found = None;
        for (position, index) in self.children(key)?.into_iter().enumerate() {
            let tree = &self.document.tree;
            if let Ok(rc) = super::record(tree, &tree.nodes[index]) {
                if integer(attr(rc, "t")?, false)? == kind
                    && integer(attr(rc, "v")?, false)? == future
                {
                    found = Some(position + 1);
                    break;
                }
            }
        }
        let index = match found {
            Some(index) => index,
            None => {
                self.push(
                    key,
                    format!("<bk xmlns=\"{NS}\"><rc t=\"{kind}\" v=\"{future}\"/></bk>"),
                )? + 1
            }
        };
        self.records.insert(cache_key, index);
        Ok(index)
    }
    fn dynamic_cell(&mut self) -> Result<usize> {
        let kind = self.kind("XLDAPR", true)?;
        let key = ContainerKey::Future("XLDAPR".into());
        let mut found = None;
        for (position, index) in self.children(&key)?.into_iter().enumerate() {
            let tree = &self.document.tree;
            if let Ok(flag) = extension(
                tree,
                &tree.nodes[index],
                DYNAMIC_URI,
                XDA,
                "dynamicArrayProperties",
            ) {
                if flag
                    .attrs
                    .keys()
                    .all(|key| ["fDynamic", "fCollapsed"].contains(&key.as_str()))
                    && flag.attrs.get("fDynamic").map(String::as_str) == Some("1")
                    && flag.attrs.get("fCollapsed").is_none_or(|flag| flag == "0")
                {
                    found = Some(position);
                    break;
                }
            }
        }
        let future = match found {
            Some(index) => index,
            None => self.push(
                &key,
                extension_xml(
                    DYNAMIC_URI,
                    &format!(
                        "<dynamicArrayProperties xmlns=\"{XDA}\" fDynamic=\"1\" fCollapsed=\"0\"/>"
                    ),
                ),
            )?,
        };
        self.record(&ContainerKey::Cells, kind, future)
    }
    fn finish(mut self) -> Result<String> {
        for (key, container) in self.containers {
            if container.added.is_empty() {
                continue;
            }
            if let Some(index) = container.original {
                for item in container.added {
                    self.document.append(index, item, true)?;
                }
            } else {
                let name = key.name();
                let future_name = match &key {
                    ContainerKey::Future(name) => format!(" name=\"{name}\""),
                    _ => String::new(),
                };
                let xml = format!(
                    "<{name} xmlns=\"{NS}\"{future_name} count=\"{}\">{}</{name}>",
                    container.added.len(),
                    container.added.concat()
                );
                let next = self.document.tree.nodes[0]
                    .children
                    .iter()
                    .copied()
                    .find(|index| {
                        let node = &self.document.tree.nodes[*index];
                        node.ns == NS
                            && match node.name.as_str() {
                                "metadataTypes" => 0,
                                "futureMetadata" => 1,
                                "cellMetadata" => 2,
                                "valueMetadata" => 3,
                                "extLst" => 4,
                                _ => 0,
                            } > key.rank()
                    });
                if let Some(index) = next {
                    self.document
                        .inserted
                        .push((self.document.spans[index].start, xml));
                } else {
                    self.document.append(0, xml, false)?;
                }
            }
        }
        self.document.finish()
    }
}

fn extension_xml(uri: &str, leaf: &str) -> String {
    format!("<bk xmlns=\"{NS}\"><extLst><ext uri=\"{uri}\">{leaf}</ext></extLst></bk>")
}

struct Rich {
    values: Document,
    structures: Document,
    structure: usize,
    kind: usize,
    mapped: BTreeMap<(u32, u32), usize>,
}
impl Rich {
    fn new(before: &BTreeMap<String, Vec<u8>>, metadata: &mut Metadata) -> Result<Self> {
        let values = Document::load(before, PARTS[1], RICH, "rvData", true)?;
        values.tree.counted(&values.tree.nodes[0], RICH, "rv")?;
        let mut structures = Document::load(before, PARTS[2], RICH, "rvStructures", true)?;
        let list = structures
            .tree
            .counted(&structures.tree.nodes[0], RICH, "s")?;
        let structure = list.iter().position(|item| {
            attributes(item, &[("t", "_error")]).is_ok()
                && whitespace(&structures.tree, item).is_ok()
                && structures.tree.children(item, RICH, "k").is_ok_and(|keys| {
                    keys.len() == ERROR_FIELDS.len()
                        && keys.iter().zip(ERROR_FIELDS).all(|(key, name)| {
                            attributes(key, &[("n", name), ("t", "i")]).is_ok()
                                && leaf(&structures.tree, key).is_ok()
                        })
                })
        });
        let structure = match structure {
            Some(index) => index,
            None => {
                let keys = ERROR_FIELDS
                    .into_iter()
                    .map(|name| format!("<k n=\"{name}\" t=\"i\"/>"))
                    .collect::<String>();
                structures.append(
                    0,
                    format!("<s xmlns=\"{RICH}\" t=\"_error\">{keys}</s>"),
                    true,
                )?
            }
        };
        let kind = metadata.kind("XLRICHVALUE", false)?;
        metadata.ensure(&ContainerKey::Future("XLRICHVALUE".into()))?;
        metadata.ensure(&ContainerKey::Values)?;
        Ok(Self {
            values,
            structures,
            structure,
            kind,
            mapped: BTreeMap::new(),
        })
    }
    fn spill_error(&mut self, metadata: &mut Metadata, rows: u32, columns: u32) -> Result<usize> {
        if let Some(index) = self.mapped.get(&(rows, columns)) {
            return Ok(*index);
        }
        let expected = [columns - 1, 8, rows - 1, 1].map(|n| n.to_string());
        let tree = &self.values.tree;
        let found = tree
            .counted(&tree.nodes[0], RICH, "rv")?
            .iter()
            .position(|item| {
                attributes(item, &[("s", &self.structure.to_string())]).is_ok()
                    && item.text.trim().is_empty()
                    && tree.children(item, RICH, "v").is_ok_and(|values| {
                        values.len() == expected.len()
                            && values.iter().zip(&expected).all(|(value, expected)| {
                                value.attrs.is_empty()
                                    && value.children.is_empty()
                                    && value.tail.trim().is_empty()
                                    && &value.text == expected
                            })
                    })
            });
        let rich_value = match found {
            Some(index) => index,
            None => {
                let values = expected
                    .iter()
                    .map(|value| format!("<v>{value}</v>"))
                    .collect::<String>();
                self.values.append(
                    0,
                    format!(
                        "<rv xmlns=\"{RICH}\" s=\"{}\">{values}</rv>",
                        self.structure
                    ),
                    true,
                )?
            }
        };
        let key = ContainerKey::Future("XLRICHVALUE".into());
        let mut found = None;
        for (position, index) in metadata.children(&key)?.into_iter().enumerate() {
            let tree = &metadata.document.tree;
            if let Ok(value) = extension(tree, &tree.nodes[index], RICH_URI, RICH, "rvb") {
                if attributes(value, &[("i", &rich_value.to_string())]).is_ok() {
                    found = Some(position);
                    break;
                }
            }
        }
        let future = match found {
            Some(index) => index,
            None => metadata.push(
                &key,
                extension_xml(
                    RICH_URI,
                    &format!("<rvb xmlns=\"{RICH}\" i=\"{rich_value}\"/>"),
                ),
            )?,
        };
        let vm = metadata.record(&ContainerKey::Values, self.kind, future)?;
        self.mapped.insert((rows, columns), vm);
        Ok(vm)
    }
}

fn rich_types_xml() -> String {
    let keys = [
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
    ]
    .into_iter()
    .map(|name| {
        format!(
            "<key name=\"{name}\">{}<flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>",
            if name == "_Self" {
                "<flag name=\"ExcludeFromFile\" value=\"1\"/>"
            } else {
                ""
            }
        )
    })
    .collect::<String>();
    format!(
        "<rvTypesInfo xmlns=\"{RICH2}\" xmlns:mc=\"{MC}\" xmlns:x=\"{NS}\" mc:Ignorable=\"x\"><global><keyFlags>{keys}</keyFlags></global></rvTypesInfo>"
    )
}

fn package_graph(
    before: &BTreeMap<String, Vec<u8>>,
    required: &[&str],
) -> Result<(String, String)> {
    let mut rels = Document::parse(
        before.get(PARTS[4]).ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing original workbook relationships",
            )
        })?,
        REL,
        "Relationships",
    )?;
    let mut types = Document::parse(
        before.get(PARTS[5]).ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing original content types")
        })?,
        CT,
        "Types",
    )?;
    let relationships = rels
        .tree
        .children(&rels.tree.nodes[0], REL, "Relationship")?;
    let mut ids = BTreeSet::new();
    for node in &relationships {
        if attr(node, "Id")?.is_empty() || !ids.insert(attr(node, "Id")?.to_owned()) {
            return Err(crate::coded(
                "spill-ambiguous",
                "Ambiguous package relationship identity",
            ));
        }
    }
    let mut overrides = BTreeMap::new();
    for index in &types.tree.nodes[0].children {
        let node = &types.tree.nodes[*index];
        if node.ns != CT || !["Override", "Default"].contains(&node.name.as_str()) {
            return Err(crate::coded(
                "spill-unsupported",
                "Unqualified content type child",
            ));
        }
        if node.name == "Override"
            && overrides
                .insert(
                    attr(node, "PartName")?.to_owned(),
                    attr(node, "ContentType")?.to_owned(),
                )
                .is_some()
        {
            return Err(crate::coded(
                "spill-ambiguous",
                "Ambiguous package content type identity",
            ));
        }
    }
    let mut additions = Vec::new();
    for (part, target, kind, content_type) in GRAPH {
        if !required.contains(&part) {
            continue;
        }
        let matches = relationships
            .iter()
            .filter(|node| node.attrs.get("Type").map(String::as_str) == Some(kind))
            .collect::<Vec<_>>();
        let exists = before.contains_key(part);
        if exists {
            if matches.len() != 1
                || matches[0].attrs.get("Target").map(String::as_str) != Some(target)
                || matches[0]
                    .attrs
                    .get("TargetMode")
                    .is_some_and(|mode| mode != "Internal")
                || overrides.get(&format!("/{part}")).map(String::as_str) != Some(content_type)
            {
                return Err(crate::coded(
                    "spill-unsupported",
                    "Existing spill part lacks a qualified package graph",
                ));
            }
        } else {
            if !matches.is_empty() || overrides.contains_key(&format!("/{part}")) {
                return Err(crate::coded(
                    "spill-refused",
                    "New spill part has a dangling original package mapping",
                ));
            }
            let mut number = 1;
            while ids.contains(&format!(workbook_format!(SPILL_NUMBER), number = number)) {
                number += 1;
            }
            let id = format!(workbook_format!(SPILL_NUMBER), number = number);
            ids.insert(id.clone());
            additions.push((format!("<Relationship xmlns=\"{REL}\" Id=\"{id}\" Type=\"{kind}\" Target=\"{target}\"/>"), format!("<Override xmlns=\"{CT}\" PartName=\"/{part}\" ContentType=\"{content_type}\"/>")));
        }
    }
    for (rel, content_type) in additions {
        rels.append(0, rel, false)?;
        types.append(0, content_type, false)?;
    }
    Ok((rels.finish()?, types.finish()?))
}

/// Groups must be derived from the stored, selected native version. The final
/// validator independently proves every index, record and package graph delta.
pub(crate) fn build(before: &BTreeMap<String, Vec<u8>>, groups: &[Value]) -> Result<Value> {
    if groups.is_empty() {
        return Ok(json!({"replacements":{},"masters":{}}));
    }
    let mut metadata = Metadata::new(before)?;
    let cm = metadata.dynamic_cell()?;
    let mut rich = None;
    let mut masters = serde_json::Map::new();
    let mut rectangles: Vec<(String, super::super::Geometry)> = Vec::new();
    for group in groups {
        crate::cancel::check()?;
        if group["schema"] != super::super::SCHEMA {
            return Err(crate::coded(
                "spill-unsupported",
                "Unsupported native spill export schema",
            ));
        }
        let sheet = crate::field(group, "sheet")?;
        let master = crate::field(group, "master")?;
        let origin = canonical_rect(master)?;
        let range = canonical_rect(crate::field(group, "range")?)?;
        if origin.count() != 1 || range.first != origin.first {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid native spill export origin",
            ));
        }
        if rectangles
            .iter()
            .any(|(prior_sheet, prior)| prior_sheet == sheet && prior.intersects(&range))
        {
            return Err(crate::coded(
                "spill-conflict",
                "Overlapping native spill export groups",
            ));
        }
        rectangles.push((sheet.into(), range.clone()));
        let key = format!("{sheet}\0{master}");
        if masters.contains_key(&key) {
            return Err(crate::coded(
                "spill-duplicate",
                "Duplicate native spill export master",
            ));
        }
        let mapped = match group["status"].as_str() {
            Some("calculated") => json!({"cm":cm,"range":range.reference()}),
            Some("error") => {
                if group["error"] != "#SPILL!" || range != origin {
                    return Err(crate::coded(
                        "spill-unsupported",
                        "Unqualified native spill error",
                    ));
                }
                let intended = canonical_rect(crate::field(group, "intendedRange")?)?;
                if intended.first != origin.first {
                    return Err(crate::coded(
                        "internal-invalid-field",
                        "Invalid native blocked spill origin",
                    ));
                }
                if rich.is_none() {
                    rich = Some(Rich::new(before, &mut metadata)?);
                }
                let vm = rich
                    .as_mut()
                    .ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing rich spill producer")
                    })?
                    .spill_error(
                        &mut metadata,
                        intended.last.0 - intended.first.0 + 1,
                        intended.last.1 - intended.first.1 + 1,
                    )?;
                json!({"cm":cm,"vm":vm,"range":master,"intendedRange":intended.reference(),"aca":"1","ca":"1"})
            }
            _ => {
                return Err(crate::coded(
                    "spill-unsupported",
                    "Unsupported native spill export status",
                ));
            }
        };
        masters.insert(key, mapped);
    }
    let mut candidates = BTreeMap::from([(PARTS[0].to_string(), metadata.finish()?)]);
    let mut required = vec![PARTS[0]];
    if let Some(rich) = rich {
        candidates.insert(PARTS[1].into(), rich.values.finish()?);
        candidates.insert(PARTS[2].into(), rich.structures.finish()?);
        if !before.contains_key(PARTS[3]) {
            candidates.insert(PARTS[3].into(), rich_types_xml());
        }
        required.extend_from_slice(&PARTS[1..4]);
    }
    let (rels, types) = package_graph(before, &required)?;
    candidates.insert(PARTS[4].into(), rels);
    candidates.insert(PARTS[5].into(), types);
    let replacements = candidates
        .into_iter()
        .filter(|(part, xml)| {
            before
                .get(part)
                .is_none_or(|original| original.as_slice() != xml.as_bytes())
        })
        .collect::<BTreeMap<_, _>>();
    let masters = Value::Object(masters);
    validate(before, &replacements, &masters)?;
    Ok(json!({"replacements":replacements,"masters":masters}))
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Read;

    fn fixture(name: &str) -> BTreeMap<String, Vec<u8>> {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../fixtures")
            .join(name);
        let mut zip = zip::ZipArchive::new(std::fs::File::open(path).unwrap()).unwrap();
        let mut parts = BTreeMap::new();
        for part in PARTS {
            if let Ok(mut file) = zip.by_name(part) {
                let mut bytes = Vec::new();
                file.read_to_end(&mut bytes).unwrap();
                parts.insert(part.into(), bytes);
            }
        }
        parts
    }
    fn pristine() -> BTreeMap<String, Vec<u8>> {
        BTreeMap::from([
            (PARTS[4].into(), format!(workbook_format!(XML_FFB1491B7F), REL = REL).into_bytes()),
            (PARTS[5].into(), format!("<Types xmlns=\"{CT}\"><Default Extension=\"xml\" ContentType=\"application/xml\"/></Types>").into_bytes()),
        ])
    }
    fn calculated(master: &str, range: &str) -> Value {
        json!({"schema":1,"sheet":"xl/worksheets/sheet1.xml","master":master,"range":range,"status":"calculated"})
    }
    fn blocked(master: &str, intended: &str) -> Value {
        json!({"schema":1,"sheet":"xl/worksheets/sheet1.xml","master":master,"range":master,"status":"error","error":"#SPILL!","intendedRange":intended})
    }
    fn replacements(plan: &Value) -> BTreeMap<String, String> {
        plan["replacements"]
            .as_object()
            .unwrap()
            .iter()
            .map(|(part, xml)| (part.clone(), xml.as_str().unwrap().into()))
            .collect()
    }
    fn after(before: &BTreeMap<String, Vec<u8>>, plan: &Value) -> BTreeMap<String, Vec<u8>> {
        let mut result = before.clone();
        result.extend(
            replacements(plan)
                .into_iter()
                .map(|(part, xml)| (part, xml.into_bytes())),
        );
        result
    }

    #[test]
    fn retained_actual_records_are_reused_without_reserializing_any_part() {
        let before = fixture("excel-dynamic-reference.xlsx");
        let plan = build(
            &before,
            &[calculated("A1", "A1:B3"), blocked("A41", "A41:B42")],
        )
        .unwrap();
        assert!(replacements(&plan).is_empty());
        assert_eq!(
            plan["masters"]["xl/worksheets/sheet1.xml\0A1"],
            json!({"cm":1,"range":"A1:B3"})
        );
        assert_eq!(
            plan["masters"]["xl/worksheets/sheet1.xml\0A41"],
            json!({"cm":1,"vm":2,"range":"A41","intendedRange":"A41:B42","aca":"1","ca":"1"})
        );
    }

    #[test]
    fn pristine_graph_gets_only_requested_dynamic_and_rich_records() {
        let before = pristine();
        let plan = build(
            &before,
            &[
                calculated("A1", "A1:B3"),
                blocked("A41", "A41:C43"),
                blocked("A51", "A51:C53"),
            ],
        )
        .unwrap();
        assert_eq!(replacements(&plan).len(), 6);
        assert_eq!(plan["masters"]["xl/worksheets/sheet1.xml\0A41"]["vm"], 1);
        assert_eq!(plan["masters"]["xl/worksheets/sheet1.xml\0A51"]["vm"], 1);
        let produced = after(&before, &plan);
        let values = Tree::parse(&produced[PARTS[1]]).unwrap();
        let records = values.counted(&values.nodes[0], RICH, "rv").unwrap();
        assert_eq!(
            records.len(),
            1,
            "Equal blocked geometries share the rich error record"
        );
        assert_eq!(
            values
                .children(records[0], RICH, "v")
                .unwrap()
                .iter()
                .map(|node| node.text.as_str())
                .collect::<Vec<_>>(),
            ["2", "8", "2", "1"]
        );
        let relationships = std::str::from_utf8(&produced[PARTS[4]]).unwrap();
        assert!(relationships.contains(workbook_format!(XML_9A9AA58ED2)));
        assert!(relationships.contains(workbook_format!(XML_906A8FF106)));
        let second = build(
            &produced,
            &[calculated("A1", "A1:B3"), blocked("A41", "A41:C43")],
        )
        .unwrap();
        assert!(replacements(&second).is_empty());
        let calculated = build(&before, &[calculated("A1", "A1:B3")]).unwrap();
        assert_eq!(
            replacements(&calculated)
                .keys()
                .map(String::as_str)
                .collect::<BTreeSet<_>>(),
            BTreeSet::from([PARTS[0], PARTS[4], PARTS[5]])
        );
    }

    #[test]
    fn source_fixture_appends_rich_errors_and_keeps_existing_dynamic_indices() {
        let before = fixture("excel-dynamic-source.xlsx");
        let plan = build(
            &before,
            &[blocked("A41", "A41:B42"), blocked("A50", "A50:C53")],
        )
        .unwrap();
        let produced = after(&before, &plan);
        assert_eq!(plan["masters"]["xl/worksheets/sheet1.xml\0A41"]["cm"], 1);
        assert_eq!(plan["masters"]["xl/worksheets/sheet1.xml\0A41"]["vm"], 1);
        assert_eq!(plan["masters"]["xl/worksheets/sheet1.xml\0A50"]["vm"], 2);
        validate(&before, &replacements(&plan), &plan["masters"]).unwrap();
        let old = Tree::parse(&before[PARTS[0]]).unwrap();
        let new = Tree::parse(&produced[PARTS[0]]).unwrap();
        for name in ["cellMetadata", "futureMetadata"] {
            let index = old.one(&old.nodes[0], NS, name, true).unwrap().unwrap();
            let original = &old.nodes[index];
            let retained = new.nodes[0]
                .children
                .iter()
                .map(|index| &new.nodes[*index])
                .find(|node| {
                    node.name == name && node.attrs.get("name") == original.attrs.get("name")
                })
                .unwrap();
            identical(&old, original, &new, retained, false).unwrap();
        }
    }

    #[test]
    fn appending_preserves_opaque_mixed_text_and_prefix_bindings_byte_for_byte() {
        let mut before = fixture("excel-dynamic-reference.xlsx");
        let opaque = "<extLst xmlns:q=\"urn:original\"><q:opaque q:attribute=\" spaced \" > leading <![CDATA[literal < text]]><q:inner/> trailing </q:opaque></extLst>";
        let source = std::str::from_utf8(&before[PARTS[0]])
            .unwrap()
            .replace("</metadata>", &format!("{opaque}</metadata>"));
        before.insert(PARTS[0].into(), source.as_bytes().to_vec());
        let original_types = before[PARTS[3]].clone();
        let plan = build(&before, &[blocked("A41", "A41:C43")]).unwrap();
        let produced = after(&before, &plan);
        assert!(
            std::str::from_utf8(&produced[PARTS[0]])
                .unwrap()
                .contains(opaque)
        );
        assert_eq!(produced[PARTS[3]], original_types);
        validate(&before, &replacements(&plan), &plan["masters"]).unwrap();
    }

    #[test]
    fn qualified_prefixed_roots_and_self_closing_containers_accept_local_namespaces() {
        let mut before = fixture("excel-dynamic-source.xlsx");
        before.insert(PARTS[0].into(), format!("<m:metadata xmlns:m=\"{NS}\" xmlns=\"urn:opaque\">\n <m:metadataTypes count='0'/>\n <m:cellMetadata count='0'/>\n <m:extLst><opaque> original </opaque></m:extLst>\n</m:metadata>").into_bytes());
        let plan = build(&before, &[blocked("A41", "A41:B42")]).unwrap();
        let xml = plan["replacements"][PARTS[0]].as_str().unwrap();
        assert!(xml.contains("<m:metadataTypes count='2'>"));
        assert!(xml.contains("<opaque> original </opaque>"));
        assert!(xml.find("name=\"XLDAPR\" count=").unwrap() < xml.find("<m:cellMetadata").unwrap());
        validate(&before, &replacements(&plan), &plan["masters"]).unwrap();
    }

    #[test]
    fn invalid_mapping_or_package_graph_is_refused() {
        let before = pristine();
        for groups in [
            vec![calculated("A1", "B1:C3")],
            vec![calculated("A1", "A1:B3"), calculated("B2", "B2:C4")],
            vec![blocked("A41", "B41:C42")],
            vec![blocked("A41", "A41:A1041")],
            vec![calculated("A1", "A1:B3"), blocked("A1", "A1:B2")],
        ] {
            assert!(build(&before, &groups).is_err());
        }
        let mut wrong_part = calculated("A1", "A1:B3");
        wrong_part["sheet"] = json!("xl/worksheets/../sheet1.xml");
        assert!(build(&before, &[wrong_part]).is_err());
        let mut wrong_graph = fixture("excel-dynamic-reference.xlsx");
        wrong_graph.insert(
            PARTS[4].into(),
            format!("<Relationships xmlns=\"{REL}\"/>").into_bytes(),
        );
        assert!(build(&wrong_graph, &[calculated("A1", "A1:B3")]).is_err());
        assert_eq!(
            build(&BTreeMap::new(), &[]).unwrap(),
            json!({"replacements":{},"masters":{}})
        );
    }

    #[test]
    fn calculation_chain_removal_is_proved_before_spill_graph_append() {
        let mut before = pristine();
        let relationship = "<Relationship Id=\"chain\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain\" Target=\"calcChain.xml\"/>";
        let content_type = "<Override PartName=\"/xl/calcChain.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml\"/>";
        for (part, closing, fragment) in [
            (PARTS[4], "</Relationships>", relationship),
            (PARTS[5], "</Types>", content_type),
        ] {
            let xml = std::str::from_utf8(&before[part])
                .unwrap()
                .replace(closing, &format!("{fragment}{closing}"));
            before.insert(part.into(), xml.into_bytes());
        }
        let calc = json!({PARTS[4]:std::str::from_utf8(&before[PARTS[4]]).unwrap().replace(relationship,""),PARTS[5]:std::str::from_utf8(&before[PARTS[5]]).unwrap().replace(content_type,"")});
        let base = calculation_base(&before, &calc, &json!(["xl/calcChain.xml"])).unwrap();
        let plan = build(&base, &[blocked("A41", "A41:B42")]).unwrap();
        assert!(
            !plan["replacements"][PARTS[4]]
                .as_str()
                .unwrap()
                .contains("calcChain")
        );
        assert!(
            !plan["replacements"][PARTS[5]]
                .as_str()
                .unwrap()
                .contains("calcChain")
        );
        validate(&base, &replacements(&plan), &plan["masters"]).unwrap();
        assert!(validate(&before, &replacements(&plan), &plan["masters"]).is_err());
    }

    #[test]
    fn export_producer_uses_stored_native_owners_and_rejects_stale_version() {
        let directory = tempfile::tempdir().unwrap();
        let root = directory.path();
        let source = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../fixtures/excel-dynamic-reference.xlsx");
        std::fs::copy(source, root.join("source.xlsx")).unwrap();
        let source_hash = crate::hash(&root.join("source.xlsx")).unwrap();
        crate::import::index(root).unwrap();
        let view = json!({"sheets":[{"name":"Cases","path":"xl/worksheets/sheet1.xml","merges":[],"cells":{}}],"tables":[],"namedRanges":[]});
        let staged = crate::spill::stage_qualified(
            root,
            &json!({"version":0,"view":view,"revision":"native-writer-test"}),
        )
        .unwrap();
        let version = staged["version"].as_u64().unwrap();
        crate::export_state::prepare(root, &json!({"version":version,"view":view})).unwrap();
        assert!(
            crate::package_edit::export_metadata::prepare(root, &json!({"version":0}))
                .unwrap_err()
                .contains("Stale")
        );
        crate::package_edit::export_metadata::prepare(
            root,
            &json!({"version":version,"groups":[{"master":"forged"}],"view":{"sheets":[]}}),
        )
        .unwrap();
        let metadata: Value =
            serde_json::from_slice(&std::fs::read(root.join("export-metadata.json")).unwrap())
                .unwrap();
        assert_eq!(
            metadata["spillMetadata"]["masters"]
                .as_object()
                .unwrap()
                .len(),
            6
        );
        assert_eq!(
            metadata["spillMetadata"]["masters"]["xl/worksheets/sheet1.xml\0A41"]["vm"],
            2
        );
        assert_eq!(crate::hash(&root.join("source.xlsx")).unwrap(), source_hash);
        let index: Value =
            serde_json::from_slice(&std::fs::read(root.join("index.json")).unwrap()).unwrap();
        crate::export::write(
            root,
            &json!({"version":version,"revision":index["revision"],"calculationMetadata":true}),
        )
        .unwrap();
        let mut exported =
            zip::ZipArchive::new(std::fs::File::open(root.join("export.xlsx")).unwrap()).unwrap();
        let mut sheet = String::new();
        exported
            .by_name("xl/worksheets/sheet1.xml")
            .unwrap()
            .read_to_string(&mut sheet)
            .unwrap();
        assert!(sheet.contains("cm=\"1\""));
        assert!(sheet.contains("vm=\"2\""));
        assert_eq!(crate::hash(&root.join("source.xlsx")).unwrap(), source_hash);
    }
}
