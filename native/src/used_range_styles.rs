//! A bounded source style inventory. Style IDs alone do not prove that a blank
//! cell has meaningful formatting; only a validated XF equivalent to XF0 does.
use crate::{Result, err};
use quick_xml::{NsReader, events::Event, name::ResolveResult};
use std::{
    cell::RefCell,
    collections::{BTreeMap, BTreeSet},
    fs::File,
    io::Read,
    path::Path,
};

const MAX_BYTES: u64 = 16 * 1024 * 1024;
const MAX_NODES: usize = 262_144;
const MAX_DEPTH: usize = 32;
const MAX_ATTRIBUTES: usize = 32;

#[derive(Debug)]
struct Node {
    name: String,
    main: bool,
    attrs: BTreeMap<String, String>,
    children: Vec<usize>,
    non_whitespace: bool,
}

/// `count` is the actual cellXfs inventory, never the untrusted count attribute.
/// An absent styles part permits only the implicit style 0. A present but empty
/// cellXfs table does not manufacture a default entry.
#[derive(Debug)]
pub(super) struct Inventory {
    pub count: u64,
    implicit: bool,
    nodes: Vec<Node>,
    cell_xfs: Vec<usize>,
    style_xfs: Vec<usize>,
    fonts: Vec<usize>,
    fills: Vec<usize>,
    borders: Vec<usize>,
    num_formats: BTreeMap<u64, usize>,
    present: BTreeSet<String>,
    // Only IDs inside the bounded actual XF inventory can enter this cache.
    // Failures are lazy too: an unsupported unused XF does not poison XF0.
    default_cache: RefCell<BTreeMap<u64, Result<bool>>>,
}

impl Inventory {
    pub(super) fn load(root: &Path) -> Result<Self> {
        crate::cancel::check()?;
        let mut zip =
            zip::ZipArchive::new(File::open(crate::local(root, "metadata.xlsx")?).map_err(err)?)
                .map_err(err)?;
        let matches = zip
            .file_names()
            .filter(|name| *name == "xl/styles.xml")
            .count();
        if matches == 0 {
            return Ok(Self::implicit());
        }
        if matches != 1 {
            return Err(crate::coded(
                "style-ambiguous",
                "Ambiguous source styles part",
            ));
        }
        let member = zip.by_name("xl/styles.xml").map_err(err)?;
        if member.size() > MAX_BYTES {
            return Err(crate::coded(
                "style-limit",
                "UsedRange source style metadata exceeds its bound",
            ));
        }
        let mut bytes = Vec::new();
        crate::cancel::Reader(member)
            .take(MAX_BYTES + 1)
            .read_to_end(&mut bytes)
            .map_err(err)?;
        if bytes.len() as u64 > MAX_BYTES {
            return Err(crate::coded(
                "style-limit",
                "UsedRange source style metadata exceeds its bound",
            ));
        }
        Self::parse(&bytes)
    }

    fn implicit() -> Self {
        Self {
            count: 1,
            implicit: true,
            nodes: Vec::new(),
            cell_xfs: Vec::new(),
            style_xfs: Vec::new(),
            fonts: Vec::new(),
            fills: Vec::new(),
            borders: Vec::new(),
            num_formats: BTreeMap::new(),
            present: BTreeSet::new(),
            default_cache: RefCell::new(BTreeMap::new()),
        }
    }

    pub(super) fn from_xml(xml: &[u8]) -> Result<Self> {
        Self::parse(xml)
    }

    fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() as u64 > MAX_BYTES {
            return Err(crate::coded(
                "style-limit",
                "UsedRange source style metadata exceeds its bound",
            ));
        }
        let mut reader = NsReader::from_reader(xml);
        let mut buffer = Vec::new();
        let mut nodes: Vec<Node> = Vec::new();
        let mut stack: Vec<usize> = Vec::new();
        let mut root = None;
        loop {
            crate::cancel::check()?;
            let (namespace, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
            if matches!(namespace, ResolveResult::Unknown(_)) {
                return Err(crate::coded(
                    "style-unsupported",
                    "Unbound source style element namespace",
                ));
            }
            let main =
                namespace == ResolveResult::Bound(quick_xml::name::Namespace(crate::xml::NS));
            match event {
                Event::Start(ref element) | Event::Empty(ref element) => {
                    if nodes.len() >= MAX_NODES || stack.len() >= MAX_DEPTH {
                        return Err(crate::coded(
                            "style-limit",
                            "UsedRange source style graph exceeds its bound",
                        ));
                    }
                    let name = std::str::from_utf8(element.local_name().as_ref())
                        .map_err(err)?
                        .to_owned();
                    let mut attrs = BTreeMap::new();
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(err)?;
                        let key = std::str::from_utf8(attribute.key.as_ref()).map_err(err)?;
                        if key == "xmlns" || key.starts_with("xmlns:") {
                            continue;
                        }
                        if attrs.len() >= MAX_ATTRIBUTES {
                            return Err(crate::coded(
                                "style-limit",
                                "UsedRange source style attributes exceed their bound",
                            ));
                        }
                        let (binding, local) = reader.resolve_attribute(attribute.key);
                        let key = match binding {
                            ResolveResult::Unbound => key.to_owned(),
                            ResolveResult::Bound(ns) => format!(
                                "{{{}}}{}",
                                std::str::from_utf8(ns.as_ref()).map_err(err)?,
                                std::str::from_utf8(local.as_ref()).map_err(err)?
                            ),
                            _ => {
                                return Err(crate::coded(
                                    "style-unsupported",
                                    "Unbound source style attribute namespace",
                                ));
                            }
                        };
                        if attrs
                            .insert(key, attribute.unescape_value().map_err(err)?.into_owned())
                            .is_some()
                        {
                            return Err(crate::coded(
                                "style-duplicate",
                                "Duplicate source style attribute",
                            ));
                        }
                    }
                    let index = nodes.len();
                    if let Some(parent) = stack.last() {
                        nodes[*parent].children.push(index);
                    } else if root.replace(index).is_some() {
                        return Err(crate::coded(
                            "style-malformed",
                            "Multiple source style roots",
                        ));
                    }
                    nodes.push(Node {
                        name,
                        main,
                        attrs,
                        children: Vec::new(),
                        non_whitespace: false,
                    });
                    if matches!(event, Event::Start(_)) {
                        stack.push(index);
                    }
                }
                Event::End(_) => {
                    stack.pop().ok_or_else(|| {
                        crate::coded("style-malformed", "Unbalanced source style XML")
                    })?;
                }
                Event::Text(text) => {
                    let text = crate::xml::xml_text(text.as_ref())?;
                    if !xml_whitespace(&text) {
                        if let Some(parent) = stack.last() {
                            nodes[*parent].non_whitespace = true;
                        } else {
                            return Err(crate::coded(
                                "style-malformed",
                                "Text outside source style root",
                            ));
                        }
                    }
                }
                Event::CData(text) => {
                    let text = std::str::from_utf8(text.as_ref()).map_err(err)?;
                    if !xml_whitespace(text) {
                        if let Some(parent) = stack.last() {
                            nodes[*parent].non_whitespace = true;
                        } else {
                            return Err(crate::coded(
                                "style-unsupported",
                                "CDATA outside source style root",
                            ));
                        }
                    }
                }
                Event::DocType(_) => {
                    return Err(crate::coded(
                        "style-malformed",
                        "Source style XML entities are unsupported",
                    ));
                }
                Event::Eof => break,
                _ => {}
            }
            buffer.clear();
        }
        if !stack.is_empty() {
            return Err(crate::coded("style-malformed", "Unclosed source style XML"));
        }
        let root =
            root.ok_or_else(|| crate::coded("style-bad-input", "Empty source styles part"))?;
        let root_node = &nodes[root];
        if !root_node.main || root_node.name != "styleSheet" || root_node.non_whitespace {
            return Err(crate::coded(
                "style-unsupported",
                "Unqualified source style root",
            ));
        }
        let mut inventory = Self {
            count: 0,
            implicit: false,
            nodes,
            cell_xfs: Vec::new(),
            style_xfs: Vec::new(),
            fonts: Vec::new(),
            fills: Vec::new(),
            borders: Vec::new(),
            num_formats: BTreeMap::new(),
            present: BTreeSet::new(),
            default_cache: RefCell::new(BTreeMap::new()),
        };
        let mut seen = BTreeSet::new();
        for index in inventory.nodes[root].children.clone() {
            crate::cancel::check()?;
            let section = &inventory.nodes[index];
            if !section.main {
                return Err(crate::coded(
                    "internal-unknown-reference",
                    "Unknown source style inventory namespace",
                ));
            }
            if !seen.insert(section.name.clone()) {
                return Err(crate::coded(
                    "style-duplicate",
                    "Duplicate source style inventory",
                ));
            }
            let record = match section.name.as_str() {
                "cellXfs" | "cellStyleXfs" => "xf",
                "fonts" => "font",
                "fills" => "fill",
                "borders" => "border",
                "numFmts" => "numFmt",
                // These inventories do not supply a referenced cell XF graph.
                // In particular global slicer/timeline extLst is preserved by
                // the package writer, rather than mistaken for an XF extension.
                "cellStyles" | "dxfs" | "tableStyles" | "colors" | "extLst" => continue,
                _ => {
                    return Err(crate::coded(
                        "internal-unknown-reference",
                        "Unknown source style inventory",
                    ));
                }
            };
            if section.non_whitespace {
                return Err(crate::coded(
                    "style-malformed",
                    "Text in source style inventory",
                ));
            }
            if let Some(count) = section.attrs.get("count") {
                unsigned(count)?;
            }
            // Section-level extension attributes (such as x14ac:knownFonts)
            // are not properties of an individual referenced font or XF.
            let entries = section.children.clone();
            for child in &entries {
                let node = &inventory.nodes[*child];
                if !node.main || node.name != record {
                    return Err(crate::coded(
                        "internal-unknown-reference",
                        "Unknown source style inventory member",
                    ));
                }
            }
            match section.name.as_str() {
                "cellXfs" => inventory.cell_xfs = entries,
                "cellStyleXfs" => inventory.style_xfs = entries,
                "fonts" => inventory.fonts = entries,
                "fills" => inventory.fills = entries,
                "borders" => inventory.borders = entries,
                "numFmts" => {
                    for entry in entries {
                        let node = &inventory.nodes[entry];
                        let id = unsigned(node.attrs.get("numFmtId").ok_or_else(|| {
                            crate::coded("style-bad-input", "Source number format has no ID")
                        })?)?;
                        if inventory.num_formats.insert(id, entry).is_some() {
                            return Err(crate::coded(
                                "style-duplicate",
                                "Duplicate source number format ID",
                            ));
                        }
                    }
                }
                _ => unreachable!(),
            }
        }
        inventory.present = seen;
        inventory.count = inventory.cell_xfs.len() as u64;
        Ok(inventory)
    }

    /// Both the requested XF and XF0 must have a known, valid referenced graph.
    /// Equivalent syntax is accepted; an unsupported graph returns an error,
    /// rather than granting permission to drop formatting or guessing its value.
    pub(super) fn is_default(&self, id: u64) -> Result<bool> {
        crate::cancel::check()?;
        if id >= self.count {
            return Err(crate::coded(
                "style-unsupported",
                "UsedRange style index is outside its source style table",
            ));
        }
        if self.implicit {
            return Ok(id == 0);
        }
        if let Some(value) = self.default_cache.borrow().get(&id) {
            return value.clone();
        }
        let outcome: Result<bool> = (|| {
            let requested = self.xf(self.cell_xfs[id as usize], true)?;
            let default = self.xf(self.cell_xfs[0], true)?;
            if requested == default {
                return Ok(true);
            }
            if requested.get("numFmtId") != default.get("numFmtId") {
                let requested_format: serde_json::Value =
                    serde_json::from_str(&requested["numFmtId"]).map_err(err)?;
                let default_format: serde_json::Value =
                    serde_json::from_str(&default["numFmtId"]).map_err(err)?;
                if requested_format.get("custom").is_some()
                    != default_format.get("custom").is_some()
                {
                    // Only these exact Home-authored formats have a closed
                    // nondefault policy against builtin General. This does
                    // not qualify arbitrary code equivalence or bypass the
                    // application/suppression checks below.
                    let general = |format: &serde_json::Value| {
                        format.get("builtin").and_then(serde_json::Value::as_u64) == Some(0)
                    };
                    let home_numeric = |format: &serde_json::Value| {
                        matches!(
                            format.get("custom").and_then(serde_json::Value::as_str),
                            Some("0.0" | "\"$\"#,##0.00;(\"$\"#,##0.00)")
                        )
                    };
                    if !(general(&requested_format) && home_numeric(&default_format)
                        || general(&default_format) && home_numeric(&requested_format))
                    {
                        return Err(crate::coded(
                            "style-unsupported",
                            "Unqualified builtin/custom source number-format equivalence",
                        ));
                    }
                }
            }
            let policies = [
                ("applyNumberFormat", "numFmtId"),
                ("applyFont", "fontId"),
                ("applyFill", "fillId"),
                ("applyBorder", "borderId"),
                ("applyAlignment", "child:alignment"),
                ("applyProtection", "child:protection"),
            ];
            let mut requested_graph = requested.clone();
            let mut default_graph = default.clone();
            for (flag, _) in policies {
                requested_graph.remove(flag);
                default_graph.remove(flag);
            }
            if requested_graph == default_graph {
                return Err(crate::coded(
                    "style-unsupported",
                    "Unqualified source XF application policy",
                ));
            }
            for (flag, component) in policies {
                // An explicitly suppressed component can inherit from the
                // base XF. A differing direct graph or parent cannot prove
                // nondefault formatting without a qualified apply policy.
                let suppressed = requested.get(flag).is_some_and(|v| v == "false")
                    || default.get(flag).is_some_and(|v| v == "false");
                if suppressed
                    && (requested.get(component) != default.get(component)
                        || requested.get("xfId") != default.get("xfId"))
                {
                    return Err(crate::coded(
                        "style-unsupported",
                        "Unqualified source XF application policy",
                    ));
                }
            }
            Ok(false)
        })();
        // A cancellation belongs to this request, not to the immutable graph.
        if !crate::cancel::interrupted() {
            self.default_cache.borrow_mut().insert(id, outcome.clone());
        }
        outcome
    }

    fn node(&self, index: usize, name: &str) -> Result<&Node> {
        let node = &self.nodes[index];
        if !node.main || node.name != name || node.non_whitespace {
            return Err(crate::coded(
                "internal-unknown-reference",
                "Unknown referenced source style node",
            ));
        }
        Ok(node)
    }

    fn xf(&self, index: usize, cell: bool) -> Result<BTreeMap<String, String>> {
        crate::cancel::check()?;
        let node = self.node(index, "xf")?;
        let mut signature = BTreeMap::new();
        let mut inherited_children = BTreeSet::new();
        let mut inherited_num_format = None;
        for (key, value) in &node.attrs {
            match key.as_str() {
                "fontId" | "fillId" | "borderId" | "numFmtId" | "xfId" => {
                    let id = unsigned(value)?;
                    let resolved = match key.as_str() {
                        "xfId" if cell => {
                            let parent =
                                self.xf(self.resource(&self.style_xfs, id, "style XF")?, false)?;
                            inherited_num_format = parent.get("numFmtId").cloned();
                            for child in ["alignment", "protection"] {
                                if parent.contains_key(&format!("child:{child}")) {
                                    inherited_children.insert(child);
                                }
                            }
                            serde_json::to_string(&parent).map_err(err)?
                        }
                        "xfId" => {
                            return Err(crate::coded(
                                "style-unsupported",
                                "Referenced base style XF has an unsupported parent",
                            ));
                        }
                        _ => self.component(key, id, true)?,
                    };
                    signature.insert(key.clone(), resolved);
                }
                "applyNumberFormat" | "applyFont" | "applyFill" | "applyBorder"
                | "applyAlignment" | "applyProtection" => {
                    signature.insert(key.clone(), boolean(value)?.to_string());
                }
                "pivotButton" | "quotePrefix" => {
                    // CT_Xf explicitly defaults these attributes to false.
                    if boolean(value)? {
                        signature.insert(key.clone(), "true".into());
                    }
                }
                _ => {
                    return Err(crate::coded(
                        "internal-unknown-reference",
                        "Unknown referenced source XF attribute",
                    ));
                }
            }
        }
        // Missing component attributes use an implicit default. An explicit
        // reference was validated above, even when its numeric ID is zero.
        for key in ["fontId", "fillId", "borderId", "numFmtId"] {
            if !node.attrs.contains_key(key) {
                // When a real zero component exists it is still part of the
                // implicit graph, so an unknown property cannot be hidden by
                // simply omitting the reference attribute. Minimal source
                // fixtures with no component inventory use an implicit default.
                signature.insert(key.into(), self.component(key, 0, false)?);
            }
        }
        let mut seen = BTreeSet::new();
        for child in &node.children {
            let child = &self.nodes[*child];
            if !child.main || !seen.insert(child.name.as_str()) {
                return Err(crate::coded(
                    "internal-unknown-reference",
                    "Unknown or duplicate referenced XF child",
                ));
            }
            let canonical = match child.name.as_str() {
                "alignment" => self.alignment(child)?,
                "protection" => self.protection(child)?,
                _ => {
                    return Err(crate::coded(
                        "internal-unknown-reference",
                        "Unknown referenced source XF child",
                    ));
                }
            };
            if canonical.is_empty() {
                if inherited_children.contains(child.name.as_str()) {
                    return Err(crate::coded(
                        "style-unsupported",
                        "Unqualified source XF default-child inheritance",
                    ));
                }
                // An explicitly applied, validated schema-default alignment
                // adds no formatting when it cannot override a nondefault
                // parent. Bare apply flags and suppression remain guarded.
                if child.name == "alignment"
                    && signature
                        .get("applyAlignment")
                        .is_some_and(|value| value == "true")
                {
                    signature.remove("applyAlignment");
                }
                continue;
            }
            // JSON is an unambiguous internal comparison encoding, not caller
            // supplied evidence. Child order and empty-element syntax vanish.
            signature.insert(
                format!("child:{}", child.name),
                serde_json::to_string(&canonical).map_err(err)?,
            );
        }
        // Positive application of an explicit, validated builtin General
        // adds nothing only when it cannot override a nondefault parent.
        // A custom definition at ID 0, bare flag or suppression stays guarded.
        if node.attrs.contains_key("numFmtId")
            && signature
                .get("numFmtId")
                .is_some_and(|value| value == r#"{"builtin":0}"#)
            && signature
                .get("applyNumberFormat")
                .is_some_and(|value| value == "true")
        {
            if inherited_num_format
                .as_deref()
                .is_some_and(|value| value != r#"{"builtin":0}"#)
            {
                return Err(crate::coded(
                    "style-unsupported",
                    "Unqualified source XF General number-format inheritance",
                ));
            }
            signature.remove("applyNumberFormat");
        }
        Ok(signature)
    }

    fn component(&self, key: &str, id: u64, explicit: bool) -> Result<String> {
        crate::cancel::check()?;
        if key == "numFmtId" {
            self.num_format(id)?;
            return if let Some(index) = self.num_formats.get(&id) {
                // Exact custom code aliases are supported. Equivalence between
                // builtin and custom codes requires a separate format policy.
                serde_json::to_string(
                    &serde_json::json!({"custom": self.nodes[*index].attrs["formatCode"]}),
                )
                .map_err(err)
            } else {
                serde_json::to_string(&serde_json::json!({"builtin": id})).map_err(err)
            };
        }
        let (section, kind, entries) = match key {
            "fontId" => ("fonts", "font", &self.fonts),
            "fillId" => ("fills", "fill", &self.fills),
            "borderId" => ("borders", "border", &self.borders),
            _ => {
                return Err(crate::coded(
                    "internal-unknown-reference",
                    "Unknown source style component reference",
                ));
            }
        };
        if !explicit && !self.present.contains(section) {
            return Ok(format!("implicit:{kind}"));
        }
        let index = self.resource(entries, id, kind)?;
        match kind {
            "font" => self.font(index)?,
            "fill" => self.fill(index)?,
            "border" => self.border(index)?,
            _ => unreachable!(),
        }
        self.canonical_resource(index)
    }

    /// Only already-validated resource graphs enter this comparison. This
    /// removes syntax and inventory aliases without guessing at theme fonts,
    /// visually similar colors, or unrelated number-format codes.
    fn canonical_resource(&self, index: usize) -> Result<String> {
        crate::cancel::check()?;
        let node = &self.nodes[index];
        let mut attrs = BTreeMap::new();
        for (key, value) in &node.attrs {
            let canonical = match key.as_str() {
                "auto" | "diagonalUp" | "diagonalDown" | "outline" => boolean(value)?.to_string(),
                "indexed" | "theme" => unsigned(value)?.to_string(),
                "tint" | "degree" | "left" | "right" | "top" | "bottom" | "position" => {
                    canonical_number(value)?
                }
                "rgb" => value.to_ascii_uppercase(),
                "style" | "patternType" | "type" => value.trim_matches(xml_space).to_owned(),
                "val" => match node.name.as_str() {
                    "b" | "i" | "strike" | "outline" | "shadow" | "condense" | "extend" => {
                        boolean(value)?.to_string()
                    }
                    "sz" => canonical_number(value)?,
                    "family" | "charset" => unsigned(value)?.to_string(),
                    "u" | "scheme" | "vertAlign" => value.trim_matches(xml_space).to_owned(),
                    _ => value.clone(),
                },
                _ => value.clone(),
            };
            attrs.insert(key.clone(), canonical);
        }
        if [
            "b", "i", "strike", "outline", "shadow", "condense", "extend",
        ]
        .contains(&node.name.as_str())
        {
            attrs.entry("val".into()).or_insert_with(|| "true".into());
        } else if node.name == "u" {
            attrs.entry("val".into()).or_insert_with(|| "single".into());
        }
        let mut children = Vec::new();
        for index in &node.children {
            let child = &self.nodes[*index];
            if node.name == "font" {
                // An absent boolean font property is false. Explicit false is
                // the same graph, rather than meaningful nondefault formatting.
                if [
                    "b", "i", "strike", "outline", "shadow", "condense", "extend",
                ]
                .contains(&child.name.as_str())
                    && child
                        .attrs
                        .get("val")
                        .is_some_and(|v| boolean(v) == Ok(false))
                {
                    continue;
                }
                if child.name == "u"
                    && child
                        .attrs
                        .get("val")
                        .is_some_and(|v| v.trim_matches(xml_space) == "none")
                {
                    continue;
                }
            }
            children.push(self.canonical_resource(*index)?);
        }
        if ["font", "border", "patternFill"].contains(&node.name.as_str()) {
            children.sort();
        }
        // Gradient stop order remains exact; equal positions can be significant.
        serde_json::to_string(
            &serde_json::json!({"node": node.name, "attrs": attrs, "children": children}),
        )
        .map_err(err)
    }

    fn resource(&self, entries: &[usize], id: u64, kind: &str) -> Result<usize> {
        entries
            .get(usize::try_from(id).map_err(err)?)
            .copied()
            .ok_or_else(|| {
                crate::coded(
                    "style-unsupported",
                    format!("Referenced source {kind} index is outside its inventory"),
                )
            })
    }

    fn num_format(&self, id: u64) -> Result<()> {
        if let Some(index) = self.num_formats.get(&id) {
            let node = self.node(*index, "numFmt")?;
            attrs_only(node, &["numFmtId", "formatCode"])?;
            leaf(node)?;
            unsigned(node.attrs.get("numFmtId").ok_or_else(|| {
                crate::coded("style-bad-input", "Source number format has no ID")
            })?)?;
            if node
                .attrs
                .get("formatCode")
                .is_none_or(|value| value.is_empty())
            {
                return Err(crate::coded(
                    "style-bad-input",
                    "Referenced source number format has no code",
                ));
            }
            Ok(())
        } else if id < 164 {
            Ok(())
        } else {
            Err(crate::coded(
                "style-unsupported",
                "Referenced custom number format is outside its inventory",
            ))
        }
    }

    fn alignment(&self, node: &Node) -> Result<BTreeMap<String, String>> {
        leaf(node)?;
        let mut canonical = BTreeMap::new();
        for (key, value) in &node.attrs {
            let value = match key.as_str() {
                "horizontal" => enumeration(
                    value,
                    &[
                        "general",
                        "left",
                        "center",
                        "right",
                        "fill",
                        "justify",
                        "centerContinuous",
                        "distributed",
                    ],
                )?
                .to_owned(),
                "vertical" => enumeration(
                    value,
                    &["top", "center", "bottom", "justify", "distributed"],
                )?
                .to_owned(),
                "textRotation" => {
                    let n = unsigned(value)?;
                    if n > 180 && n != 255 {
                        return Err(crate::coded(
                            "internal-invalid-field",
                            "Invalid source text rotation",
                        ));
                    }
                    n.to_string()
                }
                "indent" => bounded_unsigned(value, 250)?.to_string(),
                "relativeIndent" => {
                    let n = value.trim_matches(xml_space).parse::<i32>().map_err(err)?;
                    if !(-15..=15).contains(&n) {
                        return Err(crate::coded(
                            "internal-invalid-field",
                            "Invalid source relative indent",
                        ));
                    }
                    n.to_string()
                }
                "readingOrder" => bounded_unsigned(value, 2)?.to_string(),
                "wrapText" | "shrinkToFit" | "justifyLastLine" => boolean(value)?.to_string(),
                _ => {
                    return Err(crate::coded(
                        "internal-unknown-reference",
                        "Unknown referenced source alignment attribute",
                    ));
                }
            };
            let default = match key.as_str() {
                "horizontal" => "general",
                "vertical" => "bottom",
                "wrapText" | "shrinkToFit" | "justifyLastLine" => "false",
                _ => "0",
            };
            if value != default {
                canonical.insert(key.clone(), value);
            }
        }
        Ok(canonical)
    }

    fn protection(&self, node: &Node) -> Result<BTreeMap<String, String>> {
        leaf(node)?;
        attrs_only(node, &["locked", "hidden"])?;
        let locked = node
            .attrs
            .get("locked")
            .map(|value| boolean(value))
            .transpose()?
            .unwrap_or(true);
        let hidden = node
            .attrs
            .get("hidden")
            .map(|value| boolean(value))
            .transpose()?
            .unwrap_or(false);
        let mut canonical = BTreeMap::new();
        if !locked {
            canonical.insert("locked".into(), "false".into());
        }
        if hidden {
            canonical.insert("hidden".into(), "true".into());
        }
        Ok(canonical)
    }

    fn font(&self, index: usize) -> Result<()> {
        crate::cancel::check()?;
        let node = self.node(index, "font")?;
        attrs_only(node, &[])?;
        let mut seen = BTreeSet::new();
        for index in &node.children {
            let child = &self.nodes[*index];
            if !child.main || !seen.insert(child.name.as_str()) {
                return Err(crate::coded(
                    "internal-unknown-reference",
                    "Unknown or duplicate referenced font property",
                ));
            }
            leaf(child)?;
            if child.name == "color" {
                color(child)?;
                continue;
            }
            attrs_only(child, &["val"])?;
            let value = child.attrs.get("val");
            match child.name.as_str() {
                "b" | "i" | "strike" | "outline" | "shadow" | "condense" | "extend" => {
                    if let Some(value) = value {
                        boolean(value)?;
                    }
                }
                "u" => {
                    if let Some(value) = value {
                        enumeration(
                            value,
                            &[
                                "none",
                                "single",
                                "double",
                                "singleAccounting",
                                "doubleAccounting",
                            ],
                        )?;
                    }
                }
                "name" => {
                    if value.is_none_or(|value| value.is_empty()) {
                        return Err(crate::coded(
                            "style-bad-input",
                            "Referenced source font has no name",
                        ));
                    }
                }
                "sz" => {
                    let size = number(value.ok_or_else(|| {
                        crate::coded("internal-missing-field", "Source font size is missing")
                    })?)?;
                    if !(0.0..=409.0).contains(&size) || size == 0.0 {
                        return Err(crate::coded(
                            "internal-invalid-field",
                            "Invalid source font size",
                        ));
                    }
                }
                "charset" => {
                    bounded_unsigned(
                        value.ok_or_else(|| {
                            crate::coded("internal-missing-field", "Source font charset is missing")
                        })?,
                        255,
                    )?;
                }
                "family" => {
                    bounded_unsigned(
                        value.ok_or_else(|| {
                            crate::coded("internal-missing-field", "Source font family is missing")
                        })?,
                        5,
                    )?;
                }
                "scheme" => {
                    enumeration(
                        value.ok_or_else(|| {
                            crate::coded("internal-missing-field", "Source font scheme is missing")
                        })?,
                        &["none", "major", "minor"],
                    )?;
                }
                "vertAlign" => {
                    enumeration(
                        value.ok_or_else(|| {
                            crate::coded(
                                "internal-missing-field",
                                "Source font alignment is missing",
                            )
                        })?,
                        &["baseline", "superscript", "subscript"],
                    )?;
                }
                _ => {
                    return Err(crate::coded(
                        "internal-unknown-reference",
                        "Unknown referenced source font property",
                    ));
                }
            }
        }
        Ok(())
    }

    fn fill(&self, index: usize) -> Result<()> {
        crate::cancel::check()?;
        let node = self.node(index, "fill")?;
        attrs_only(node, &[])?;
        if node.children.len() > 1 {
            return Err(crate::coded(
                "style-ambiguous",
                "Ambiguous referenced source fill",
            ));
        }
        for child in &node.children {
            let child = &self.nodes[*child];
            if !child.main || child.non_whitespace {
                return Err(crate::coded(
                    "internal-unknown-reference",
                    "Unknown referenced source fill",
                ));
            }
            match child.name.as_str() {
                "patternFill" => {
                    attrs_only(child, &["patternType"])?;
                    if let Some(value) = child.attrs.get("patternType") {
                        enumeration(
                            value,
                            &[
                                "none",
                                "solid",
                                "mediumGray",
                                "darkGray",
                                "lightGray",
                                "darkHorizontal",
                                "darkVertical",
                                "darkDown",
                                "darkUp",
                                "darkGrid",
                                "darkTrellis",
                                "lightHorizontal",
                                "lightVertical",
                                "lightDown",
                                "lightUp",
                                "lightGrid",
                                "lightTrellis",
                                "gray125",
                                "gray0625",
                            ],
                        )?;
                    }
                    let mut seen = BTreeSet::new();
                    for index in &child.children {
                        let color_node = &self.nodes[*index];
                        if !color_node.main
                            || !["fgColor", "bgColor"].contains(&color_node.name.as_str())
                            || !seen.insert(color_node.name.as_str())
                        {
                            return Err(crate::coded(
                                "internal-unknown-reference",
                                "Unknown referenced pattern fill color",
                            ));
                        }
                        color(color_node)?;
                    }
                }
                "gradientFill" => {
                    attrs_only(child, &["type", "degree", "left", "right", "top", "bottom"])?;
                    for (key, value) in &child.attrs {
                        match key.as_str() {
                            "type" => {
                                enumeration(value, &["linear", "path"])?;
                            }
                            "degree" => {
                                let value = number(value)?;
                                if !(0.0..=360.0).contains(&value) {
                                    return Err(crate::coded(
                                        "internal-invalid-field",
                                        "Invalid source gradient degree",
                                    ));
                                }
                            }
                            _ => {
                                unit_number(value)?;
                            }
                        }
                    }
                    for index in &child.children {
                        let stop = self.node(*index, "stop")?;
                        attrs_only(stop, &["position"])?;
                        unit_number(stop.attrs.get("position").ok_or_else(|| {
                            crate::coded("style-bad-input", "Source gradient stop has no position")
                        })?)?;
                        if stop.children.len() != 1 {
                            return Err(crate::coded(
                                "style-bad-input",
                                "Source gradient stop has no single color",
                            ));
                        }
                        color(self.node(stop.children[0], "color")?)?;
                    }
                }
                _ => {
                    return Err(crate::coded(
                        "internal-unknown-reference",
                        "Unknown referenced source fill child",
                    ));
                }
            }
        }
        Ok(())
    }

    fn border(&self, index: usize) -> Result<()> {
        crate::cancel::check()?;
        let node = self.node(index, "border")?;
        attrs_only(node, &["diagonalUp", "diagonalDown", "outline"])?;
        for value in node.attrs.values() {
            boolean(value)?;
        }
        let mut seen = BTreeSet::new();
        for index in &node.children {
            let side = &self.nodes[*index];
            if !side.main
                || side.non_whitespace
                || ![
                    "left",
                    "right",
                    "top",
                    "bottom",
                    "diagonal",
                    "vertical",
                    "horizontal",
                    "start",
                    "end",
                ]
                .contains(&side.name.as_str())
                || !seen.insert(side.name.as_str())
            {
                return Err(crate::coded(
                    "internal-unknown-reference",
                    "Unknown referenced source border side",
                ));
            }
            attrs_only(side, &["style"])?;
            if let Some(value) = side.attrs.get("style") {
                enumeration(
                    value,
                    &[
                        "none",
                        "thin",
                        "medium",
                        "dashed",
                        "dotted",
                        "thick",
                        "double",
                        "hair",
                        "mediumDashed",
                        "dashDot",
                        "mediumDashDot",
                        "dashDotDot",
                        "mediumDashDotDot",
                        "slantDashDot",
                    ],
                )?;
            }
            if side.children.len() > 1 {
                return Err(crate::coded(
                    "style-ambiguous",
                    "Ambiguous referenced source border color",
                ));
            }
            for index in &side.children {
                color(self.node(*index, "color")?)?;
            }
        }
        Ok(())
    }
}

fn xml_space(character: char) -> bool {
    matches!(character, ' ' | '\t' | '\r' | '\n')
}
fn xml_whitespace(value: &str) -> bool {
    value.chars().all(xml_space)
}
fn unsigned(value: &str) -> Result<u64> {
    let value = value.trim_matches(xml_space);
    let digits = value.strip_prefix('+').unwrap_or(value);
    if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid source style unsigned integer",
        ));
    }
    let parsed = digits.parse::<u64>().map_err(err)?;
    if parsed > u64::from(u32::MAX) {
        return Err(crate::coded(
            "style-limit",
            "Source style unsigned integer exceeds its schema range",
        ));
    }
    Ok(parsed)
}
fn bounded_unsigned(value: &str, max: u64) -> Result<u64> {
    let value = unsigned(value)?;
    if value > max {
        return Err(crate::coded(
            "style-limit",
            "Source style integer exceeds its allowed range",
        ));
    }
    Ok(value)
}
fn boolean(value: &str) -> Result<bool> {
    match value.trim_matches(xml_space) {
        "0" | "false" => Ok(false),
        "1" | "true" => Ok(true),
        _ => Err(crate::coded(
            "internal-invalid-field",
            "Invalid source style boolean",
        )),
    }
}
fn number(value: &str) -> Result<f64> {
    let value = value.trim_matches(xml_space).parse::<f64>().map_err(err)?;
    if !value.is_finite() {
        return Err(crate::coded(
            "style-unsupported",
            "Nonfinite source style number",
        ));
    }
    Ok(value)
}
fn canonical_number(value: &str) -> Result<String> {
    let value = number(value)?;
    Ok(if value == 0.0 {
        "0".into()
    } else {
        value.to_string()
    })
}
fn unit_number(value: &str) -> Result<f64> {
    let value = number(value)?;
    if !(0.0..=1.0).contains(&value) {
        return Err(crate::coded(
            "style-unsupported",
            "Source style fraction is outside its allowed range",
        ));
    }
    Ok(value)
}
fn enumeration<'a>(value: &'a str, allowed: &[&str]) -> Result<&'a str> {
    let value = value.trim_matches(xml_space);
    if !allowed.contains(&value) {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown source style enumeration",
        ));
    }
    Ok(value)
}
fn attrs_only(node: &Node, allowed: &[&str]) -> Result<()> {
    if node
        .attrs
        .keys()
        .any(|key| !allowed.contains(&key.as_str()))
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown referenced source style attribute",
        ));
    }
    Ok(())
}
fn leaf(node: &Node) -> Result<()> {
    if node.non_whitespace || !node.children.is_empty() {
        return Err(crate::coded(
            "style-malformed",
            "Unexpected referenced source style content",
        ));
    }
    Ok(())
}
fn color(node: &Node) -> Result<()> {
    leaf(node)?;
    attrs_only(node, &["auto", "indexed", "rgb", "theme", "tint"])?;
    if ["auto", "indexed", "rgb", "theme"]
        .iter()
        .filter(|key| node.attrs.contains_key(**key))
        .count()
        > 1
    {
        return Err(crate::coded(
            "style-ambiguous",
            "Ambiguous referenced source style color",
        ));
    }
    for (key, value) in &node.attrs {
        match key.as_str() {
            "auto" => {
                boolean(value)?;
            }
            "indexed" => {
                bounded_unsigned(value, 65)?;
            }
            "theme" => {
                bounded_unsigned(value, 11)?;
            }
            "rgb" => {
                if value.len() != 8 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
                    return Err(crate::coded(
                        "internal-invalid-field",
                        "Invalid source style RGB color",
                    ));
                }
            }
            "tint" => {
                let tint = number(value)?;
                if !(-1.0..=1.0).contains(&tint) {
                    return Err(crate::coded(
                        "internal-invalid-field",
                        "Invalid source style color tint",
                    ));
                }
            }
            _ => unreachable!(),
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;

    fn styles(xfs: &str) -> Inventory {
        Inventory::from_xml(format!(r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><name val="Calibri"/><sz val="11"/></font><font><name val="Calibri"/><sz val="11"/><b/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="999">{xfs}</cellXfs><extLst><ext uri="global"><outside xmlns="urn:unrelated" arbitrary="kept"/></ext></extLst></styleSheet>"#).as_bytes()).unwrap()
    }

    #[test]
    fn equivalent_xf_zero_uses_real_inventory_and_canonical_syntax() {
        let inventory = styles(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf xfId="000" borderId="+0" fillId="0" fontId="00" numFmtId="0" quotePrefix="false" pivotButton="0"> 
        </xf><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>"#,
        );
        assert_eq!(inventory.count, 3);
        assert!(inventory.is_default(0).unwrap());
        assert!(inventory.is_default(1).unwrap());
        assert!(!inventory.is_default(2).unwrap());
        assert!(inventory.is_default(3).is_err());
        assert!(inventory.is_default(u64::MAX).is_err());
    }

    #[test]
    fn implicit_missing_styles_are_distinct_from_present_empty_xfs() {
        assert!(Inventory::implicit().is_default(0).unwrap());
        assert!(Inventory::implicit().is_default(1).is_err());
        for xml in [
            r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="1"/></styleSheet>"#,
            r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#,
        ] {
            let inventory = Inventory::from_xml(xml.as_bytes()).unwrap();
            assert_eq!(inventory.count, 0);
            assert!(inventory.is_default(0).is_err());
        }
        let inventory = Inventory::from_xml(br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs><xf/><xf numFmtId="0"/></cellXfs></styleSheet>"#).unwrap();
        assert_eq!(inventory.count, 2);
        assert!(inventory.is_default(1).unwrap());
    }

    #[test]
    fn load_reads_the_actual_bounded_metadata_part_without_changing_it() {
        let root = tempfile::tempdir().unwrap();
        let path = root.path().join("metadata.xlsx");
        for text in [
            None,
            Some(
                r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="50"><xf/><xf/></cellXfs></styleSheet>"#,
            ),
        ] {
            let mut zip = zip::ZipWriter::new(File::create(&path).unwrap());
            if let Some(text) = text {
                zip.start_file("xl/styles.xml", zip::write::SimpleFileOptions::default())
                    .unwrap();
                zip.write_all(text.as_bytes()).unwrap();
            }
            zip.finish().unwrap();
            let before = std::fs::read(&path).unwrap();
            let inventory = Inventory::load(root.path()).unwrap();
            assert_eq!(inventory.count, if text.is_some() { 2 } else { 1 });
            assert!(inventory.is_default(0).unwrap());
            assert_eq!(std::fs::read(&path).unwrap(), before);
        }
    }

    #[test]
    fn invalid_resource_references_and_parent_graph_never_grant_equivalence() {
        for attributes in [
            "fontId=\"2\"",
            "fillId=\"1\"",
            "borderId=\"1\"",
            "xfId=\"1\"",
            "numFmtId=\"164\"",
            "fontId=\"-1\"",
            "numFmtId=\"4294967296\"",
            "applyFont=\"yes\"",
        ] {
            let inventory = styles(&format!("<xf/><xf {attributes}/>"));
            assert!(inventory.is_default(1).is_err(), "{attributes}");
        }
        let inventory = Inventory::from_xml(br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs><xf/><xf fontId="0"/></cellXfs></styleSheet>"#).unwrap();
        assert!(inventory.is_default(0).unwrap());
        assert!(inventory.is_default(1).is_err());
        let inventory = Inventory::from_xml(br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellStyleXfs><xf xfId="0"/></cellStyleXfs><cellXfs><xf xfId="0"/></cellXfs></styleSheet>"#).unwrap();
        assert!(inventory.is_default(0).is_err());
    }

    #[test]
    fn unknown_referenced_xf_and_resource_properties_are_refused_locally() {
        for xf in [
            r#"<xf invented="0"/>"#,
            r#"<xf xmlns:x="urn:unknown" x:fontId="0"/>"#,
            r#"<xf xmlns="urn:unknown"/>"#,
            r#"<xf><extLst/></xf>"#,
            r#"<xf><alignment mystery="0"/></xf>"#,
            r#"<xf><alignment/><alignment/></xf>"#,
            r#"<xf><protection locked="maybe"/></xf>"#,
        ] {
            // A foreign XF is not an actual inventory member and is rejected
            // during parsing. Unknown properties on a real XF are checked only
            // when that XF is referenced, preserving unrelated styles.
            let xml = format!(
                r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs><xf/>{xf}</cellXfs></styleSheet>"#
            );
            match Inventory::from_xml(xml.as_bytes()) {
                Err(_) => assert!(xf.contains("xmlns=\"urn:unknown\"")),
                Ok(inventory) => {
                    assert!(inventory.is_default(0).unwrap());
                    assert!(inventory.is_default(1).is_err(), "{xf}");
                }
            }
        }
        let inventory = Inventory::from_xml(br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts><font/><font><b val="invalid"/></font><font><extLst/></font></fonts><cellXfs><xf fontId="0"/><xf fontId="1"/><xf fontId="2"/></cellXfs></styleSheet>"#).unwrap();
        assert!(inventory.is_default(0).unwrap());
        assert!(inventory.is_default(1).is_err());
        assert!(inventory.is_default(2).is_err());
        let inventory = Inventory::from_xml(br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts><font><extLst/></font></fonts><cellXfs><xf/></cellXfs></styleSheet>"#).unwrap();
        assert!(
            inventory.is_default(0).is_err(),
            "Omitting fontId cannot hide an unknown real default font"
        );
        let inventory = Inventory::from_xml(br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts/><cellXfs><xf/></cellXfs></styleSheet>"#).unwrap();
        assert!(
            inventory.is_default(0).is_err(),
            "A present empty font inventory has no implicit entry"
        );
    }

    #[test]
    fn child_order_empty_tags_and_boolean_lexemes_are_canonical() {
        let inventory = styles(
            r#"<xf><alignment wrapText="1"/><protection/></xf><xf><protection hidden="0" locked="1"></protection><alignment wrapText="true"> </alignment></xf><xf><alignment wrapText="0"/><protection/></xf>"#,
        );
        assert!(inventory.is_default(1).unwrap());
        assert!(!inventory.is_default(2).unwrap());
        let inventory = styles(r#"<xf/><xf applyFont="0"/>"#);
        // Optional apply flags are not defaulted without evidence that absence
        // and explicit false have the same inheritance behavior.
        assert!(inventory.is_default(1).is_err());
    }

    #[test]
    fn exact_resource_and_base_xf_aliases_are_default_without_matching_ids() {
        let xml = br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts><numFmt numFmtId="164" formatCode="0.000"/><numFmt numFmtId="165" formatCode="0.000"/></numFmts><fonts><font><name val="Calibri"/><sz val="11"/></font><font><sz val="11.0"></sz><name val="Calibri"/></font><font><sz val="11"/><name val="Calibri"/><b/></font></fonts><fills><fill><patternFill patternType="solid"><fgColor rgb="FF00ABCD"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><bgColor indexed="+064"/><fgColor rgb="ff00abcd"/></patternFill></fill></fills><borders><border outline="true"><left style="thin"/><top/></border><border outline="1"><top></top><left style="thin"/></border></borders><cellStyleXfs><xf numFmtId="164" fontId="0" fillId="0" borderId="0"/><xf borderId="1" fillId="1" fontId="1" numFmtId="165"/></cellStyleXfs><cellXfs><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="165" fontId="1" fillId="1" borderId="1" xfId="1"/><xf numFmtId="165" fontId="2" fillId="1" borderId="1" xfId="1" applyFont="1"/></cellXfs></styleSheet>"#;
        let inventory = Inventory::from_xml(xml).unwrap();
        assert!(inventory.is_default(1).unwrap());
        assert!(!inventory.is_default(2).unwrap());
        let inventory = Inventory::from_xml(br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts><font><b/><sz val="11"/></font><font><sz val="11.0"/><b val="true"/></font></fonts><cellXfs><xf fontId="0"/><xf fontId="1"/></cellXfs></styleSheet>"#).unwrap();
        assert!(inventory.is_default(1).unwrap());
        let inventory = Inventory::from_xml(br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts><font><sz val="11"/></font><font><sz val="11.0"/><b val="false"/><i val="0"/><u val="none"/></font></fonts><cellXfs><xf fontId="0"/><xf fontId="1"/></cellXfs></styleSheet>"#).unwrap();
        assert!(inventory.is_default(1).unwrap());
        let inventory = styles(
            r#"<xf/><xf><alignment/></xf><xf><alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false" shrinkToFit="0" justifyLastLine="false" indent="0" relativeIndent="0" readingOrder="0"/></xf><xf><protection locked="true" hidden="0"/></xf>"#,
        );
        for id in 1..4 {
            assert!(
                inventory.is_default(id).unwrap(),
                "Exact default XF child at {id}"
            );
        }
    }

    #[test]
    fn unobserved_application_policy_is_incomplete_even_for_known_graphs() {
        for flag in [
            "applyNumberFormat",
            "applyFont",
            "applyFill",
            "applyBorder",
            "applyAlignment",
            "applyProtection",
        ] {
            for value in ["0", "1"] {
                let inventory = styles(&format!("<xf/><xf {flag}=\"{value}\"/>"));
                assert!(
                    inventory.is_default(1).is_err(),
                    "{flag}={value} differs only by apply state"
                );
            }
        }
        let inventory = styles(
            r#"<xf fontId="0"/><xf fontId="1" applyFont="0"/><xf fontId="1" applyFont="1"/>"#,
        );
        assert!(
            inventory.is_default(1).is_err(),
            "Suppressed changed font has unqualified inheritance"
        );
        assert!(
            !inventory.is_default(2).unwrap(),
            "Observed explicitly applied bold font remains supported"
        );
        let inventory = styles(r#"<xf fontId="0" applyFont="0"/><xf fontId="1" applyFont="1"/>"#);
        assert!(
            inventory.is_default(1).is_err(),
            "Suppressed baseline is guarded symmetrically"
        );
        let inventory = Inventory::from_xml(br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellStyleXfs><xf><alignment wrapText="1"/><protection hidden="1"/></xf></cellStyleXfs><cellXfs><xf xfId="0"/><xf xfId="0"><alignment/></xf><xf xfId="0"><protection locked="1" hidden="0"/></xf></cellXfs></styleSheet>"#).unwrap();
        assert!(inventory.is_default(0).unwrap());
        assert!(
            inventory.is_default(1).is_err(),
            "Empty direct alignment can override a nondefault parent"
        );
        assert!(
            inventory.is_default(2).is_err(),
            "Default protection can override a nondefault parent"
        );
    }

    #[test]
    fn explicitly_applied_default_alignment_is_qualified_without_widening_apply_policy() {
        for flag in ["1", "true"] {
            for alignment in [
                "vertical=\"bottom\"",
                "horizontal=\"general\" wrapText=\"0\"",
                "horizontal=\"general\" vertical=\"bottom\" textRotation=\"0\" wrapText=\"false\" shrinkToFit=\"0\" justifyLastLine=\"false\" indent=\"0\" relativeIndent=\"0\" readingOrder=\"0\"",
            ] {
                let inventory = styles(&format!(
                    "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"{flag}\"><alignment {alignment}/></xf>"
                ));
                assert!(inventory.is_default(1).unwrap(), "{flag}: {alignment}");
            }
        }
        for xf in [
            r#"<xf applyAlignment="1"/>"#,
            r#"<xf applyAlignment="0"><alignment vertical="bottom"/></xf>"#,
            r#"<xf applyAlignment="false"><alignment vertical="bottom"/></xf>"#,
            r#"<xf applyAlignment="1"><alignment vertical="bottom" unknown="0"/></xf>"#,
            r#"<xf applyAlignment="1"><alignment vertical="bottom"/><alignment/></xf>"#,
        ] {
            assert!(styles(&format!("<xf/>{xf}")).is_default(1).is_err(), "{xf}");
        }
        let inherited = Inventory::from_xml(br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellStyleXfs><xf><alignment vertical="top"/></xf></cellStyleXfs><cellXfs><xf xfId="0"/><xf xfId="0" applyAlignment="1"><alignment vertical="bottom"/></xf></cellXfs></styleSheet>"#).unwrap();
        assert!(inherited.is_default(0).unwrap());
        assert!(
            inherited.is_default(1).is_err(),
            "Explicit default cannot bypass parent override guard"
        );
    }

    #[test]
    fn explicitly_applied_general_requires_a_validated_default_number_format_parent() {
        for flag in ["1", "true"] {
            for (base, requested) in [
                (
                    r#"<xf numFmtId="0" xfId="0"/>"#.to_owned(),
                    format!(r#"<xf numFmtId="0" xfId="0" applyNumberFormat="{flag}"/>"#),
                ),
                (
                    r#"<xf numFmtId="0"/>"#.to_owned(),
                    format!(r#"<xf numFmtId="+000" applyNumberFormat="{flag}"/>"#),
                ),
            ] {
                for xfs in [format!("{base}{requested}"), format!("{requested}{base}")] {
                    assert!(styles(&xfs).is_default(1).unwrap(), "{xfs}");
                }
            }
        }
        let aliases = Inventory::from_xml(br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellStyleXfs><xf numFmtId="0"/><xf numFmtId="000" applyNumberFormat="true"/></cellStyleXfs><cellXfs><xf numFmtId="0" xfId="0"/><xf numFmtId="0" xfId="1" applyNumberFormat="1"/></cellXfs></styleSheet>"#).unwrap();
        assert!(
            aliases.is_default(1).unwrap(),
            "Equivalent validated parent aliases stay equivalent"
        );
        for xf in [
            r#"<xf applyNumberFormat="1"/>"#,
            r#"<xf numFmtId="0" applyNumberFormat="0"/>"#,
            r#"<xf numFmtId="0" applyNumberFormat="false"/>"#,
            r#"<xf numFmtId="0" applyNumberFormat="1" unknown="0"/>"#,
            r#"<xf numFmtId="0" applyNumberFormat="1"><unknown/></xf>"#,
        ] {
            assert!(styles(&format!("<xf/>{xf}")).is_default(1).is_err(), "{xf}");
        }
        let custom_zero = Inventory::from_xml(br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts><numFmt numFmtId="0" formatCode="0.0"/></numFmts><cellXfs><xf numFmtId="0"/><xf numFmtId="0" applyNumberFormat="1"/></cellXfs></styleSheet>"#).unwrap();
        assert!(
            custom_zero.is_default(1).is_err(),
            "A custom code at ID zero is not builtin General"
        );
        for parent in [
            r#"<xf numFmtId="2"/>"#,
            r#"<xf numFmtId="164"/>"#,
            r#"<xf numFmtId="0" unknown="0"/>"#,
        ] {
            for xfs in [
                r#"<xf numFmtId="0" xfId="1"/><xf numFmtId="0" xfId="1" applyNumberFormat="1"/>"#,
                r#"<xf numFmtId="0" xfId="0"/><xf numFmtId="0" xfId="1" applyNumberFormat="1"/>"#,
                r#"<xf numFmtId="0" xfId="1" applyNumberFormat="1"/><xf numFmtId="0" xfId="0"/>"#,
            ] {
                let xml = format!(
                    r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts><numFmt numFmtId="164" formatCode="0.0"/></numFmts><cellStyleXfs><xf numFmtId="0"/>{parent}</cellStyleXfs><cellXfs>{xfs}</cellXfs></styleSheet>"#
                );
                assert!(
                    Inventory::from_xml(xml.as_bytes())
                        .unwrap()
                        .is_default(1)
                        .is_err(),
                    "{parent}: {xfs}"
                );
            }
        }
        let unknown_resource = Inventory::from_xml(br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts><font unknown="0"/></fonts><cellXfs><xf numFmtId="0"/><xf numFmtId="0" applyNumberFormat="1"/></cellXfs></styleSheet>"#).unwrap();
        assert!(
            unknown_resource.is_default(1).is_err(),
            "General application cannot bypass referenced font validation"
        );
    }

    #[test]
    fn closed_home_numeric_codes_are_nondefault_to_general_in_both_directions() {
        for code in ["0.0", "\"$\"#,##0.00;(\"$\"#,##0.00)"] {
            for flag in ["1", "true"] {
                let custom = format!(r#"<xf numFmtId="164" applyNumberFormat="{flag}"/>"#);
                let general = format!(r#"<xf numFmtId="0" applyNumberFormat="{flag}"/>"#);
                for xfs in [format!("{general}{custom}"), format!("{custom}{general}")] {
                    let xml = format!(
                        r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts><numFmt numFmtId="164" formatCode='{code}'/></numFmts><cellXfs>{xfs}</cellXfs></styleSheet>"#
                    );
                    assert!(
                        !Inventory::from_xml(xml.as_bytes())
                            .unwrap()
                            .is_default(1)
                            .unwrap(),
                        "{code}: {xfs}"
                    );
                }
            }
            for flag in ["0", "false"] {
                for (custom, general) in [
                    (
                        format!(r#"<xf numFmtId="164" applyNumberFormat="{flag}"/>"#),
                        r#"<xf numFmtId="0" applyNumberFormat="1"/>"#.to_owned(),
                    ),
                    (
                        r#"<xf numFmtId="164" applyNumberFormat="1"/>"#.to_owned(),
                        format!(r#"<xf numFmtId="0" applyNumberFormat="{flag}"/>"#),
                    ),
                ] {
                    for xfs in [format!("{general}{custom}"), format!("{custom}{general}")] {
                        let xml = format!(
                            r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts><numFmt numFmtId="164" formatCode='{code}'/></numFmts><cellXfs>{xfs}</cellXfs></styleSheet>"#
                        );
                        assert!(
                            Inventory::from_xml(xml.as_bytes())
                                .unwrap()
                                .is_default(1)
                                .is_err(),
                            "Suppressed {code}: {xfs}"
                        );
                    }
                }
            }
            for (definition, xf) in [
                (
                    format!(r#"<numFmt numFmtId="164" formatCode='{code}' unknown="0"/>"#),
                    r#"<xf numFmtId="164" applyNumberFormat="1"/>"#,
                ),
                (
                    format!(r#"<numFmt numFmtId="164" formatCode='{code}'/>"#),
                    r#"<xf numFmtId="164" applyNumberFormat="1"><unknown/></xf>"#,
                ),
            ] {
                let xml = format!(
                    r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts>{definition}</numFmts><cellXfs><xf numFmtId="0"/>{xf}</cellXfs></styleSheet>"#
                );
                assert!(
                    Inventory::from_xml(xml.as_bytes())
                        .unwrap()
                        .is_default(1)
                        .is_err(),
                    "Known code still validates its graph: {xml}"
                );
            }
        }
    }

    #[test]
    fn closed_home_numeric_policy_keeps_other_codes_and_builtin_pairs_guarded() {
        for (code, builtin) in [
            ("General", 0),
            ("0.000", 0),
            ("0.00", 0),
            ("0.0 ", 0),
            ("$#,##0.00", 0),
            (" \"$\"#,##0.00;(\"$\"#,##0.00)", 0),
            ("\"$\"#,##0.00;(\"$\"#,##0.00) ", 0),
            ("0.0", 1),
            ("0.0", 2),
            ("0.0", 14),
            ("\"$\"#,##0.00;(\"$\"#,##0.00)", 44),
        ] {
            let custom = r#"<xf numFmtId="164" applyNumberFormat="1"/>"#;
            let builtin = format!(r#"<xf numFmtId="{builtin}"/>"#);
            for xfs in [format!("{builtin}{custom}"), format!("{custom}{builtin}")] {
                let xml = format!(
                    r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts><numFmt numFmtId="164" formatCode='{code}'/></numFmts><cellXfs>{xfs}</cellXfs></styleSheet>"#
                );
                assert!(
                    Inventory::from_xml(xml.as_bytes())
                        .unwrap()
                        .is_default(1)
                        .is_err(),
                    "{code}: {xfs}"
                );
            }
        }
    }

    #[test]
    fn custom_number_formats_and_known_resource_graphs_are_validated() {
        let xml = br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts><numFmt numFmtId="164" formatCode="0.000"/></numFmts><fonts><font><color theme="1" tint="-0.2"/><scheme val="minor"/></font></fonts><fills><fill><gradientFill degree="45"><stop position="0"><color rgb="FF000000"/></stop><stop position="1"><color rgb="FFFFFFFF"/></stop></gradientFill></fill></fills><borders><border diagonalUp="1"><diagonal style="thin"><color indexed="64"/></diagonal></border></borders><cellXfs><xf numFmtId="164" fontId="0" fillId="0" borderId="0"/><xf borderId="0" fillId="0" fontId="0" numFmtId="0164"/></cellXfs></styleSheet>"#;
        assert!(Inventory::from_xml(xml).unwrap().is_default(1).unwrap());
        for (before, after) in [
            ("theme=\"1\"", "theme=\"12\""),
            ("degree=\"45\"", "degree=\"NaN\""),
            ("position=\"1\"", "position=\"2\""),
            ("diagonalUp=\"1\"", "diagonalUp=\"2\""),
            ("formatCode=\"0.000\"", "formatCode=\"\""),
            ("style=\"thin\"", "style=\"unknown\""),
        ] {
            let changed = std::str::from_utf8(xml).unwrap().replace(before, after);
            assert!(
                Inventory::from_xml(changed.as_bytes())
                    .unwrap()
                    .is_default(0)
                    .is_err(),
                "{after}"
            );
        }
    }

    #[test]
    fn builtin_custom_format_crossing_is_guarded_without_a_format_engine() {
        let inventory = Inventory::from_xml(br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts><numFmt numFmtId="164" formatCode="General"/><numFmt numFmtId="165" formatCode="0.000"/></numFmts><cellXfs><xf numFmtId="0"/><xf numFmtId="164"/><xf numFmtId="165" applyNumberFormat="1"/><xf numFmtId="2" applyNumberFormat="1"/></cellXfs></styleSheet>"#).unwrap();
        assert!(
            inventory.is_default(1).is_err(),
            "Custom General cannot be guessed nondefault"
        );
        assert!(
            inventory.is_default(2).is_err(),
            "Crossing to another opaque custom code stays unqualified"
        );
        assert!(
            !inventory.is_default(3).unwrap(),
            "Known positive builtin formatting remains supported"
        );
        let inventory = Inventory::from_xml(br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts><numFmt numFmtId="164" formatCode="General"/><numFmt numFmtId="165" formatCode="General"/></numFmts><cellXfs><xf numFmtId="164"/><xf numFmtId="165"/><xf numFmtId="0"/></cellXfs></styleSheet>"#).unwrap();
        assert!(inventory.is_default(0).unwrap());
        assert!(
            inventory.is_default(1).unwrap(),
            "Exact custom-code aliases remain supported"
        );
        assert!(
            inventory.is_default(2).is_err(),
            "Crossing from custom to builtin is guarded symmetrically"
        );
    }

    #[test]
    fn malformed_and_oversized_style_documents_are_refused() {
        for xml in [
            "",
            "<!DOCTYPE styleSheet><styleSheet/>",
            "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><cellXfs><xf></cellXfs></styleSheet>",
            "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><cellXfs/><cellXfs/></styleSheet>",
        ] {
            assert!(Inventory::from_xml(xml.as_bytes()).is_err());
        }
        assert!(Inventory::from_xml(&vec![b' '; MAX_BYTES as usize + 1]).is_err());
    }
}
