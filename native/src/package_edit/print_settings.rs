//! Narrow print-definition admission. No cells, drawings or printer parts are
//! replaced here; their source package members remain in the streaming writer.
use super::*;
use crate::xml::address;
const TAGS: [&str; 8] = [
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "legacyDrawingHF",
    "customSheetViews",
];
#[derive(Default)]
pub(super) struct Node {
    pub(super) ns: Vec<u8>,
    pub(super) name: String,
    pub(super) qname: String,
    pub(super) attrs: BTreeMap<String, String>,
    pub(super) text: String,
    pub(super) children: Vec<usize>,
    pub(super) start: usize,
    pub(super) end: usize,
}
pub(super) struct Tree {
    pub(super) bytes: Vec<u8>,
    pub(super) nodes: Vec<Node>,
}
impl Tree {
    pub(super) fn parse(bytes: &[u8]) -> Result<Self> {
        if bytes.len() > XML_LIMIT {
            return Err(crate::coded("print-limit", "Print XML exceeds its bound"));
        }
        let source = std::str::from_utf8(bytes).map_err(err)?;
        if source.to_ascii_lowercase().contains("<!doctype")
            || source.to_ascii_lowercase().contains("<!entity")
        {
            return Err(crate::coded(
                "print-malformed",
                "Print XML entities are protected",
            ));
        }
        let mut reader = NsReader::from_reader(bytes);
        let mut buffer = Vec::new();
        let mut nodes: Vec<Node> = Vec::new();
        let mut stack: Vec<usize> = Vec::new();
        let mut roots = 0;
        loop {
            crate::cancel::check()?;
            let start = reader.buffer_position() as usize;
            let (binding, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
            let namespace = match binding {
                quick_xml::name::ResolveResult::Bound(ns) => ns.as_ref().to_vec(),
                _ => Vec::new(),
            };
            let end = reader.buffer_position() as usize;
            match event {
                Event::Start(ref element) | Event::Empty(ref element) => {
                    if nodes.len() >= 50000 || stack.len() > 64 {
                        return Err(crate::coded(
                            "print-limit",
                            "Print XML node/depth bound exceeded",
                        ));
                    }
                    if stack.is_empty() {
                        roots += 1;
                        if roots != 1 {
                            return Err(crate::coded(
                                "print-malformed",
                                "Multiple print XML roots",
                            ));
                        }
                    }
                    let ns = namespace;
                    let mut values = BTreeMap::new();
                    for attribute in element.attributes().with_checks(true) {
                        let attribute = attribute.map_err(err)?;
                        let (binding, name) = reader.resolve_attribute(attribute.key);
                        let key = match binding {
                            quick_xml::name::ResolveResult::Unknown(_)
                                if !attribute.key.as_ref().starts_with(b"xmlns:") =>
                            {
                                return Err(crate::coded(
                                    "print-protected",
                                    "Unbound XML attribute namespace is protected",
                                ));
                            }
                            quick_xml::name::ResolveResult::Unbound
                                if attribute.key.as_ref().contains(&b':')
                                    && !attribute.key.as_ref().starts_with(b"xmlns:") =>
                            {
                                return Err(crate::coded(
                                    "print-protected",
                                    "Unbound XML attribute namespace is protected",
                                ));
                            }
                            quick_xml::name::ResolveResult::Bound(ns) if ns.as_ref() == REL_NS => {
                                format!("r:{}", String::from_utf8_lossy(name.as_ref()))
                            }
                            quick_xml::name::ResolveResult::Bound(ns)
                                if ns.as_ref() != b"http://www.w3.org/2000/xmlns/" =>
                            {
                                format!(
                                    "{{{}}}{}",
                                    String::from_utf8_lossy(ns.as_ref()),
                                    String::from_utf8_lossy(name.as_ref())
                                )
                            }
                            _ => String::from_utf8_lossy(attribute.key.as_ref()).into_owned(),
                        };
                        if values
                            .insert(
                                key,
                                attribute
                                    .decode_and_unescape_value(reader.decoder())
                                    .map_err(err)?
                                    .into_owned(),
                            )
                            .is_some()
                        {
                            return Err(crate::coded(
                                "print-duplicate",
                                "Duplicate print XML attribute",
                            ));
                        }
                    }
                    let index = nodes.len();
                    nodes.push(Node {
                        ns,
                        name: String::from_utf8_lossy(element.local_name().as_ref()).into_owned(),
                        qname: String::from_utf8_lossy(element.name().as_ref()).into_owned(),
                        attrs: values,
                        start,
                        end,
                        ..Default::default()
                    });
                    if let Some(parent) = stack.last() {
                        nodes[*parent].children.push(index);
                    }
                    if matches!(event, Event::Start(_)) {
                        stack.push(index);
                    }
                }
                Event::End(_) => {
                    let index = stack
                        .pop()
                        .ok_or_else(|| crate::coded("print-malformed", "Unbalanced print XML"))?;
                    nodes[index].end = end;
                }
                Event::Text(text) => {
                    let value = text.unescape().map_err(err)?;
                    if let Some(index) = stack.last() {
                        nodes[*index].text.push_str(&value);
                    } else if !value.trim().is_empty() {
                        return Err(crate::coded("print-malformed", "Text outside print XML"));
                    }
                }
                Event::CData(text) => {
                    if let Some(index) = stack.last() {
                        nodes[*index]
                            .text
                            .push_str(std::str::from_utf8(&text).map_err(err)?);
                    } else {
                        return Err(crate::coded("print-unsupported", "CDATA outside print XML"));
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buffer.clear();
        }
        if roots != 1 || !stack.is_empty() {
            return Err(crate::coded("print-malformed", "Invalid print XML root"));
        }
        Ok(Self {
            bytes: bytes.to_vec(),
            nodes,
        })
    }
    pub(super) fn root(&self, ns: &[u8], name: &str) -> Result<()> {
        if self.nodes[0].ns != ns || self.nodes[0].name != name {
            return Err(crate::coded(
                "print-malformed",
                "Invalid print XML namespace/root",
            ));
        }
        Ok(())
    }
    pub(super) fn direct(&self, parent: usize, name: &str) -> Result<Option<usize>> {
        let found = self.nodes[parent]
            .children
            .iter()
            .copied()
            .filter(|index| self.nodes[*index].name == name)
            .collect::<Vec<_>>();
        if found.len() > 1 || found.iter().any(|index| self.nodes[*index].ns != NS) {
            return Err(crate::coded(
                "print-ambiguous",
                format!("Ambiguous/foreign print {name}"),
            ));
        }
        Ok(found.first().copied())
    }
    pub(super) fn raw(&self, index: usize) -> &[u8] {
        let node = &self.nodes[index];
        &self.bytes[node.start..node.end]
    }
    pub(super) fn clean(&self, index: usize, allowed: &[&str]) -> Result<()> {
        let node = &self.nodes[index];
        if !node.children.is_empty()
            || !node.text.trim().is_empty()
            // An attribute from another namespace (Excel's xr:uid revision id)
            // says nothing about the record itself.
            || node.attrs.keys().any(|key| {
                key != "xmlns"
                    && !key.starts_with("xmlns:")
                    && !key.starts_with('{')
                    && !allowed.contains(&key.as_str())
            })
        {
            return Err(crate::coded(
                "print-protected",
                format!("Extended print {} is protected", node.name),
            ));
        }
        Ok(())
    }
    pub(super) fn mask(&self, indexes: &[usize]) -> Result<Vec<u8>> {
        let mut ranges = indexes
            .iter()
            .map(|index| (self.nodes[*index].start, self.nodes[*index].end))
            .collect::<Vec<_>>();
        ranges.sort();
        let mut result = Vec::new();
        let mut previous = 0;
        for (start, end) in ranges {
            if start < previous {
                return Err(crate::coded(
                    "print-conflict",
                    "Overlapping print definitions",
                ));
            }
            result.extend_from_slice(&self.bytes[previous..start]);
            previous = end;
        }
        result.extend_from_slice(&self.bytes[previous..]);
        Ok(result)
    }
}
struct Definition {
    sheet_pr: Option<usize>,
    setup_pr: Option<usize>,
    tags: Vec<Option<usize>>,
    container: Option<usize>,
    names: Vec<usize>,
    revision: Option<String>,
}
fn definition(
    sheet: &Tree,
    workbook: &Tree,
    part: &str,
    name: &str,
    scope: u64,
) -> Result<Definition> {
    sheet.root(NS, "worksheet")?;
    workbook.root(NS, "workbook")?;
    let sheet_pr = sheet.direct(0, "sheetPr")?;
    let setup_pr = sheet_pr
        .map(|index| sheet.direct(index, "pageSetUpPr"))
        .transpose()?
        .flatten();
    let tags = TAGS
        .iter()
        .map(|tag| sheet.direct(0, tag))
        .collect::<Result<Vec<_>>>()?;
    let container = workbook.direct(0, "definedNames")?;
    let names = container
        .map(|index| {
            workbook.nodes[index]
                .children
                .iter()
                .copied()
                .filter(|index| {
                    let node = &workbook.nodes[*index];
                    node.ns == NS
                        && node.name == "definedName"
                        && node.attrs.get("name").is_some_and(|name| {
                            ["_xlnm.print_area", "_xlnm.print_titles"]
                                .contains(&name.to_ascii_lowercase().as_str())
                        })
                        && node
                            .attrs
                            .get("localSheetId")
                            .is_none_or(|id| id == &scope.to_string())
                })
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    let mut raw = vec![
        /* Saved V1 CAS identity: hash the original domain separator. */
        workbook_format!(PRINT_DEFINITION_V1).as_bytes().to_vec(),
        part.as_bytes().to_vec(),
        name.as_bytes().to_vec(),
        scope.to_string().into_bytes(),
    ];
    let mut selected = 0;
    for index in setup_pr.into_iter().chain(tags.iter().flatten().copied()) {
        raw.push(sheet.raw(index).to_vec());
        selected += 1;
    }
    for index in &names {
        raw.push(workbook.raw(*index).to_vec());
        selected += 1;
    }
    let revision = if selected == 0 {
        None
    } else {
        let mut digest = Sha256::new();
        for (index, record) in raw.iter().enumerate() {
            if index > 0 {
                digest.update([0]);
            }
            digest.update(record);
        }
        Some(format!("{:x}", digest.finalize()))
    };
    Ok(Definition {
        sheet_pr,
        setup_pr,
        tags,
        container,
        names,
        revision,
    })
}
fn allowed_fields(value: &Value, fields: &[&str]) -> Result<()> {
    let object = value
        .as_object()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid print definition object"))?;
    if object.len() != fields.len() || object.keys().any(|key| !fields.contains(&key.as_str())) {
        return Err(crate::coded(
            "print-unsupported",
            "Unqualified print definition fields",
        ));
    }
    Ok(())
}
fn boolean(value: Option<&String>, default: bool) -> Result<bool> {
    match value.map(String::as_str) {
        None | Some("") => Ok(default),
        Some("1" | "true") => Ok(true),
        Some("0" | "false") => Ok(false),
        _ => Err(crate::coded(
            "internal-invalid-field",
            "Invalid print boolean",
        )),
    }
}
fn number(node: &Node, key: &str, default: f64, min: f64, max: f64) -> Result<f64> {
    let value = node
        .attrs
        .get(key)
        .map(|value| value.parse::<f64>().map_err(err))
        .transpose()?
        .unwrap_or(default);
    if !value.is_finite() || value < min || value > max {
        return Err(crate::coded(
            "internal-invalid-field",
            format!("Invalid print {key}"),
        ));
    }
    Ok(value)
}
fn numeric(value: &Value, min: f64, max: f64) -> Result<f64> {
    let value = value
        .as_f64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing print numeric setting"))?;
    if !value.is_finite() || value < min || value > max {
        return Err(crate::coded(
            "print-limit",
            "Print number exceeds qualified bounds",
        ));
    }
    Ok(value)
}
fn whole(value: &Value, min: u64, max: u64) -> Result<u64> {
    value
        .as_u64()
        .filter(|value| *value >= min && *value <= max)
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid integer print setting"))
}
fn references<'a>(text: &'a str, sheet: &str) -> Result<Vec<&'a str>> {
    let quoted = format!("'{}'!", sheet.replace('\'', "''"));
    let simple = format!("{sheet}!");
    let mut remaining = text;
    let mut result = Vec::new();
    loop {
        let reference = remaining
            .strip_prefix(&quoted)
            .or_else(|| remaining.strip_prefix(&simple))
            .ok_or_else(|| {
                crate::coded("internal-mismatch", "External/mismatched print reference")
            })?;
        let (value, next) = reference
            .split_once(',')
            .map(|(value, next)| (value, Some(next)))
            .unwrap_or((reference, None));
        result.push(value);
        if result.len() > 2 {
            return Err(crate::coded(
                "print-unsupported",
                "Disconnected print references",
            ));
        }
        if let Some(next) = next {
            remaining = next
        } else {
            break;
        }
    }
    Ok(result)
}
fn absolute_range(reference: &str) -> Result<((u32, u32), (u32, u32))> {
    let fields = reference.split(':').collect::<Vec<_>>();
    if fields.len() != 2 {
        return Err(crate::coded(
            "print-bad-input",
            "Print area needs one rectangle",
        ));
    }
    for field in &fields {
        let values = field
            .strip_prefix('$')
            .ok_or_else(|| crate::coded("print-bad-input", "Print area must be absolute"))?
            .split('$')
            .collect::<Vec<_>>();
        if values.len() != 2
            || values[0].is_empty()
            || !values[0].bytes().all(|ch| ch.is_ascii_uppercase())
            || values[1].is_empty()
            || !values[1].bytes().all(|ch| ch.is_ascii_digit())
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid absolute print area",
            ));
        }
    }
    let bounds = crate::bulk::bounds(&reference.replace('$', ""))?;
    if u64::from(bounds.1.0 - bounds.0.0 + 1) * u64::from(bounds.1.1 - bounds.0.1 + 1) > 10000 {
        return Err(crate::coded(
            "print-limit",
            "Print area exceeds 10,000 positions",
        ));
    }
    Ok(bounds)
}

fn plain_sections(encoded: &str) -> Result<Value> {
    if encoded.encode_utf16().count() > 255
        || (!encoded.is_empty()
            && !["&L", "&C", "&R"]
                .iter()
                .any(|prefix| encoded.starts_with(prefix)))
    {
        return Err(crate::coded(
            "print-unsupported",
            "Unqualified header/footer sections",
        ));
    }
    let mut sections = [String::new(), String::new(), String::new()];
    let mut side = 0usize;
    let mut chars = encoded.chars();
    while let Some(ch) = chars.next() {
        if ch == '&' {
            match chars.next() {
                Some('&') => sections[side].push('&'),
                Some('L') => side = 0,
                Some('C') => side = 1,
                Some('R') => side = 2,
                _ => {
                    return Err(crate::coded(
                        "print-protected",
                        "Formatted, dynamic or image headers/footers are protected",
                    ));
                }
            }
        } else {
            if ch <= '\u{1f}' || ['\u{7f}', '\u{fffe}', '\u{ffff}'].contains(&ch) {
                return Err(crate::coded(
                    "print-bad-input",
                    "Use single-line XML-safe header/footer text",
                ));
            }
            sections[side].push(ch);
        }
    }
    Ok(serde_json::json!({"left":sections[0],"center":sections[1],"right":sections[2]}))
}
fn validate_plain_sections(value: &Value) -> Result<()> {
    allowed_fields(value, &["left", "center", "right"])?;
    let mut encoded = String::new();
    for (side, code) in [("left", "L"), ("center", "C"), ("right", "R")] {
        let text = field(value, side)?;
        if text
            .chars()
            .any(|ch| ch <= '\u{1f}' || ['\u{7f}', '\u{fffe}', '\u{ffff}'].contains(&ch))
        {
            return Err(crate::coded(
                "print-bad-input",
                "Use single-line XML-safe header/footer text",
            ));
        }
        if !text.is_empty() {
            encoded.push('&');
            encoded.push_str(code);
            encoded.push_str(&text.replace('&', "&&"));
        }
    }
    if plain_sections(&encoded)? != *value {
        return Err(crate::coded(
            "internal-mismatch",
            "Header/footer text differs from its encoded sections",
        ));
    }
    Ok(())
}
fn validate_plain_header_footer(value: &Value) -> Result<()> {
    allowed_fields(value, &["header", "footer", "scaleWithDoc"])?;
    validate_plain_sections(&value["header"])?;
    validate_plain_sections(&value["footer"])?;
    if !value["scaleWithDoc"].is_boolean() {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid header/footer scaling",
        ));
    }
    Ok(())
}
fn plain_header_footer(tree: &Tree, index: usize) -> Result<Value> {
    let node = &tree.nodes[index];
    let allowed = [
        "differentOddEven",
        "differentFirst",
        "scaleWithDoc",
        "alignWithMargins",
    ];
    if !node.text.trim().is_empty()
        || node.attrs.keys().any(|key| {
            key != "xmlns" && !key.starts_with("xmlns:") && !allowed.contains(&key.as_str())
        })
        || boolean(node.attrs.get("differentOddEven"), false)?
        || boolean(node.attrs.get("differentFirst"), false)?
        || !boolean(node.attrs.get("alignWithMargins"), true)?
    {
        return Err(crate::coded(
            "print-protected",
            "Advanced headers/footers are protected",
        ));
    }
    let empty = serde_json::json!({"left":"","center":"","right":""});
    let mut value = serde_json::json!({"header":empty,"footer":empty,"scaleWithDoc":boolean(node.attrs.get("scaleWithDoc"),true)?});
    let mut seen = BTreeSet::new();
    for child in &node.children {
        let item = &tree.nodes[*child];
        if item.ns != NS
            || !["oddHeader", "oddFooter"].contains(&item.name.as_str())
            || !seen.insert(&item.name)
            || !item.children.is_empty()
            || item
                .attrs
                .keys()
                .any(|key| key != "xmlns" && !key.starts_with("xmlns:"))
        {
            return Err(crate::coded(
                "print-protected",
                "Advanced headers/footers are protected",
            ));
        }
        let side = if item.name == "oddHeader" {
            "header"
        } else {
            "footer"
        };
        value[side] = plain_sections(&item.text)?;
    }
    validate_plain_header_footer(&value)?;
    Ok(value)
}

fn protect_source(
    sheet: &Tree,
    workbook: &Tree,
    definition: &Definition,
    name: &str,
    scope: u64,
) -> Result<()> {
    if definition.tags.iter().skip(4).any(Option::is_some) {
        return Err(crate::coded(
            "print-protected",
            "Advanced imported print settings are protected",
        ));
    }
    if let Some(index) = definition.tags[3] {
        plain_header_footer(sheet, index)?;
    }
    if let Some(index) = definition.setup_pr {
        sheet.clean(index, &["fitToPage", "autoPageBreaks"])?;
        boolean(sheet.nodes[index].attrs.get("fitToPage"), false)?;
        if !boolean(sheet.nodes[index].attrs.get("autoPageBreaks"), true)? {
            return Err(crate::coded(
                "print-protected",
                "Manual page setup is protected",
            ));
        }
    }
    if let Some(index) = definition.tags[0] {
        sheet.clean(
            index,
            &[
                "gridLines",
                "gridLinesSet",
                "headings",
                "horizontalCentered",
                "verticalCentered",
            ],
        )?;
        for key in ["headings", "horizontalCentered", "verticalCentered"] {
            if boolean(sheet.nodes[index].attrs.get(key), false)? {
                return Err(crate::coded(
                    "print-unsupported",
                    "Print headings/centering require qualification",
                ));
            }
        }
        boolean(sheet.nodes[index].attrs.get("gridLines"), false)?;
        boolean(sheet.nodes[index].attrs.get("gridLinesSet"), false)?;
    }
    if let Some(index) = definition.tags[1] {
        sheet.clean(
            index,
            &["top", "right", "bottom", "left", "header", "footer"],
        )?;
        for key in ["top", "right", "bottom", "left", "header", "footer"] {
            number(&sheet.nodes[index], key, 0.5, 0.0, 2.0)?;
        }
    }
    if let Some(index) = definition.tags[2] {
        sheet.clean(
            index,
            &[
                "paperSize",
                "orientation",
                "scale",
                "fitToWidth",
                "fitToHeight",
            ],
        )?;
        let node = &sheet.nodes[index];
        if node
            .attrs
            .get("paperSize")
            .is_some_and(|value| !["1", "9"].contains(&value.as_str()))
            || node
                .attrs
                .get("orientation")
                .is_some_and(|value| !["portrait", "landscape"].contains(&value.as_str()))
        {
            return Err(crate::coded(
                "print-protected",
                "Imported print paper/orientation is protected",
            ));
        }
        number(node, "scale", 100.0, 10.0, 400.0)?;
        number(node, "fitToWidth", 1.0, 1.0, 10.0)?;
        number(node, "fitToHeight", 0.0, 0.0, 10.0)?;
    }
    let mut seen = BTreeSet::new();
    let mut area = None;
    let mut titles = None;
    for index in &definition.names {
        let node = &workbook.nodes[*index];
        if !node.children.is_empty()
            || node.attrs.keys().any(|key| {
                key != "name"
                    && key != "localSheetId"
                    && key != "xmlns"
                    && !key.starts_with("xmlns:")
            })
            || node.attrs.get("localSheetId") != Some(&scope.to_string())
        {
            return Err(crate::coded(
                "print-protected",
                "Global/extended print names are protected",
            ));
        }
        let key = node
            .attrs
            .get("name")
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing print name"))?
            .to_ascii_lowercase();
        if !seen.insert(key.clone()) {
            return Err(crate::coded(
                "print-duplicate",
                "Duplicate reserved print name",
            ));
        }
        if key.ends_with("print_area") {
            let refs = references(&node.text, name)?;
            if refs.len() != 1 {
                return Err(crate::coded("print-unsupported", "Disconnected print area"));
            }
            area = Some(absolute_range(refs[0])?);
        } else {
            titles = Some(references(&node.text, name)?);
        }
    }
    if let Some(titles) = titles {
        let area =
            area.ok_or_else(|| crate::coded("print-bad-input", "Print titles require an area"))?;
        let mut row = false;
        let mut col = false;
        for range in titles {
            let fields = range
                .split(':')
                .map(|field| {
                    field.strip_prefix('$').ok_or_else(|| {
                        crate::coded("print-bad-input", "Print titles must be absolute")
                    })
                })
                .collect::<Result<Vec<_>>>()?;
            if fields.len() != 2 {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid print title range",
                ));
            }
            if fields
                .iter()
                .all(|field| field.bytes().all(|ch| ch.is_ascii_digit()))
            {
                let first = fields[0].parse::<u32>().map_err(err)?;
                let last = fields[1].parse::<u32>().map_err(err)?;
                if row
                    || first != area.0.0 + 1
                    || last < first
                    || last >= area.1.0 + 1
                    || last - first > 99
                {
                    return Err(crate::coded(
                        "print-unsupported",
                        "Unsupported repeat row prefix",
                    ));
                }
                row = true;
            } else {
                let first = position(&format!("{}1", fields[0]))?.1;
                let last = position(&format!("{}1", fields[1]))?.1;
                if col || first != area.0.1 || last < first || last >= area.1.1 || last - first > 99
                {
                    return Err(crate::coded(
                        "print-unsupported",
                        "Unsupported repeat column prefix",
                    ));
                }
                col = true;
            }
        }
    }
    Ok(())
}
fn expanded_empty(bytes: &[u8], qname: &str) -> Vec<u8> {
    let text = String::from_utf8_lossy(bytes);
    if let Some(before) = text.strip_suffix("/>") {
        format!("{before}></{qname}>").into_bytes()
    } else {
        bytes.to_vec()
    }
}
fn container_mask(tree: &Tree, container: usize, selected: &[usize]) -> Result<Vec<u8>> {
    let node = &tree.nodes[container];
    let bytes = tree.raw(container);
    let mut spans = selected
        .iter()
        .map(|index| {
            (
                tree.nodes[*index].start - node.start,
                tree.nodes[*index].end - node.start,
            )
        })
        .collect::<Vec<_>>();
    spans.sort();
    let mut result = Vec::new();
    let mut prior = 0;
    for (start, end) in spans {
        result.extend_from_slice(&bytes[prior..start]);
        prior = end;
    }
    result.extend_from_slice(&bytes[prior..]);
    Ok(expanded_empty(&result, &node.qname))
}
pub(super) fn validate(root: &Path, plan: &Value, input: &Value) -> Result<()> {
    allowed_fields(
        plan,
        &[
            "kind",
            "sheetPart",
            "sheetName",
            "scope",
            "priorPrintRevision",
            "settings",
            "replacements",
            "removed",
            "copies",
            "fragments",
        ],
    )?;
    if plan["removed"]
        .as_array()
        .is_none_or(|value| !value.is_empty())
        || plan["copies"]
            .as_object()
            .is_none_or(|value| !value.is_empty())
        || plan["replacements"]
            .as_object()
            .is_none_or(|value| value.len() != 1 || !value.contains_key("xl/workbook.xml"))
    {
        return Err(crate::coded(
            "print-bad-input",
            "Print edits touch only worksheet settings and workbook print names",
        ));
    }
    let part = field(plan, "sheetPart")?;
    let name = field(plan, "sheetName")?;
    let scope = plan["scope"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid print scope"))?;
    let empty = serde_json::Map::new();
    let mut metadata = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let source_sheet = Tree::parse(&part_bytes(&mut metadata, part, &empty)?)?;
    let source_workbook = Tree::parse(&part_bytes(&mut metadata, "xl/workbook.xml", &empty)?)?;
    let previous = definition(&source_sheet, &source_workbook, part, name, scope)?;
    protect_source(&source_sheet, &source_workbook, &previous, name, scope)?;
    let sheets = source_workbook.direct(0, "sheets")?.ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing print worksheet ownership",
        )
    })?;
    let sheet_node = source_workbook.nodes[sheets]
        .children
        .iter()
        .filter(|index| {
            source_workbook.nodes[**index].ns == NS
                && source_workbook.nodes[**index].name == "sheet"
        })
        .nth(scope as usize)
        .ok_or_else(|| crate::coded("print-refused", "Ungrantable print scope"))?;
    let sheet_node = &source_workbook.nodes[*sheet_node];
    if sheet_node.attrs.get("name").map(String::as_str) != Some(name) {
        return Err(crate::coded("print-stale", "Print worksheet name changed"));
    }
    let id = sheet_node.attrs.get("r:id").ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing print worksheet relationship",
        )
    })?;
    let relations = Tree::parse(&part_bytes(
        &mut metadata,
        "xl/_rels/workbook.xml.rels",
        &empty,
    )?)?;
    relations.root(PACKAGE_REL_NS, "Relationships")?;
    let edges = relations.nodes[0]
        .children
        .iter()
        .map(|index| &relations.nodes[*index])
        .filter(|node| node.attrs.get("Id") == Some(id))
        .collect::<Vec<_>>();
    if edges.len() != 1
        || edges[0].ns != PACKAGE_REL_NS
        || edges[0].name != "Relationship"
        || edges[0].attrs.get("TargetMode").is_some()
        || edges[0].attrs.get("Type")
            != Some(&format!("{}/worksheet", String::from_utf8_lossy(REL_NS)))
        || resolve(
            "xl/workbook.xml",
            edges[0].attrs.get("Target").ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing worksheet target")
            })?,
        )? != part
    {
        return Err(crate::coded(
            "print-unsupported",
            "Print worksheet relationship is unqualified",
        ));
    }
    match (&previous.revision, &plan["priorPrintRevision"]) {
        (None, Value::Null) => {}
        (Some(expected), Value::String(actual)) if expected == actual => {}
        _ => {
            return Err(crate::coded(
                "print-stale",
                "Conflict: native print definition changed",
            ));
        }
    }
    let settings = &plan["settings"];
    let mut settings_fields = vec![
        "area",
        "paper",
        "orientation",
        "marginsPt",
        "scaling",
        "repeatRows",
        "repeatColumns",
        "gridlines",
    ];
    let header_footer = settings.get("headerFooter");
    if let Some(value) = header_footer {
        settings_fields.push("headerFooter");
        validate_plain_header_footer(value)?;
    }
    allowed_fields(settings, &settings_fields)?;
    let has_area = !settings["area"].is_null();
    if !has_area && (!settings["repeatRows"].is_null() || !settings["repeatColumns"].is_null()) {
        return Err(crate::coded(
            "print-bad-input",
            "Print repeats require an area",
        ));
    }
    let area = if has_area {
        crate::bulk::bounds(field(settings, "area")?)?
    } else {
        ((0, 0), (0, 0))
    };
    if u64::from(area.1.0 - area.0.0 + 1) * u64::from(area.1.1 - area.0.1 + 1) > 10000 {
        return Err(crate::coded(
            "print-limit",
            "Print settings exceed the 10k slice",
        ));
    }
    let paper = field(settings, "paper")?;
    let orientation = field(settings, "orientation")?;
    if !["A4", "Letter"].contains(&paper) || !["portrait", "landscape"].contains(&orientation) {
        return Err(crate::coded(
            "print-unsupported",
            "Unsupported print paper/orientation",
        ));
    }
    allowed_fields(
        &settings["marginsPt"],
        &["top", "right", "bottom", "left", "header", "footer"],
    )?;
    for key in ["top", "right", "bottom", "left", "header", "footer"] {
        numeric(&settings["marginsPt"][key], 0.0, 144.0)?;
    }
    let (size_w, size_h) = if paper == "A4" {
        (210.0 / 25.4 * 72.0, 297.0 / 25.4 * 72.0)
    } else {
        (612.0, 792.0)
    };
    let (w, h) = if orientation == "landscape" {
        (size_h, size_w)
    } else {
        (size_w, size_h)
    };
    if numeric(&settings["marginsPt"]["left"], 0.0, 144.0)?
        + numeric(&settings["marginsPt"]["right"], 0.0, 144.0)?
        >= w - 18.0
        || numeric(&settings["marginsPt"]["top"], 0.0, 144.0)?
            + numeric(&settings["marginsPt"]["bottom"], 0.0, 144.0)?
            >= h - 18.0
    {
        return Err(crate::coded(
            "print-bad-input",
            "Print margins leave no body",
        ));
    }
    let fit = field(&settings["scaling"], "mode")? == "fit";
    if fit {
        allowed_fields(&settings["scaling"], &["mode", "width", "height"])?;
        whole(&settings["scaling"]["width"], 1, 10)?;
        whole(&settings["scaling"]["height"], 0, 10)?;
    } else {
        if field(&settings["scaling"], "mode")? != "scale" {
            return Err(crate::coded(
                "internal-unknown-reference",
                "Unknown print scaling",
            ));
        }
        allowed_fields(&settings["scaling"], &["mode", "percent"])?;
        whole(&settings["scaling"]["percent"], 10, 400)?;
    }
    let grid = settings["gridlines"]
        .as_bool()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid print gridlines"))?;
    let fragments = plan["fragments"]
        .as_object()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing print fragments"))?;
    if fragments.len() != 1 || !fragments.contains_key(part) {
        return Err(crate::coded(
            "print-refused",
            "Print fragments have an ungranted worksheet",
        ));
    }
    let mut fragment_fields = vec!["sheetPr", "printOptions", "pageMargins", "pageSetup"];
    if header_footer.is_some() {
        fragment_fields.push("headerFooter");
    }
    allowed_fields(&fragments[part], &fragment_fields)?;
    let raw = std::str::from_utf8(&source_sheet.bytes).map_err(err)?;
    let root_node = &source_sheet.nodes[0];
    let closing = raw[..root_node.end]
        .rfind("</")
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing worksheet close"))?;
    let inserted = fragment_fields
        .iter()
        .map(|tag| field(&fragments[part], tag).map(str::to_string))
        .collect::<Result<Vec<_>>>()?;
    let mut patches = Vec::new();
    for (tag, text) in fragment_fields.iter().zip(inserted) {
        if text.len() > 65536 {
            return Err(crate::coded(
                "print-limit",
                "Print fragment exceeds its bound",
            ));
        }
        let previous = source_sheet.direct(0, tag)?;
        patches.push((
            previous
                .map(|index| source_sheet.nodes[index].start)
                .unwrap_or(closing),
            previous
                .map(|index| source_sheet.nodes[index].end)
                .unwrap_or(closing),
            text,
        ));
    }
    patches.sort_by(|a, b| b.0.cmp(&a.0));
    let mut candidate = raw.to_string();
    for (start, end, text) in patches {
        candidate.replace_range(start..end, &text);
    }
    let sheet = Tree::parse(candidate.as_bytes())?;
    let workbook = Tree::parse(field(&plan["replacements"], "xl/workbook.xml")?.as_bytes())?;
    let next = definition(&sheet, &workbook, part, name, scope)?;
    protect_source(&sheet, &workbook, &next, name, scope)?;
    match (header_footer, next.tags[3]) {
        (Some(expected), Some(index)) if plain_header_footer(&sheet, index)? == *expected => {}
        (None, None) => {}
        _ => {
            return Err(crate::coded(
                "internal-mismatch",
                "Proposed header/footer differs from settings",
            ));
        }
    }
    let setup_pr = next
        .setup_pr
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing proposed pageSetUpPr"))?;
    if boolean(sheet.nodes[setup_pr].attrs.get("fitToPage"), false)? != fit
        || !boolean(sheet.nodes[setup_pr].attrs.get("autoPageBreaks"), false)?
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Print fit policy differs from settings",
        ));
    }
    let options = next.tags[0]
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing print options"))?;
    if boolean(sheet.nodes[options].attrs.get("gridLines"), false)? != grid
        || !boolean(sheet.nodes[options].attrs.get("gridLinesSet"), false)?
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Print gridline policy differs",
        ));
    }
    let margins = next.tags[1]
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing print margins"))?;
    for key in ["top", "right", "bottom", "left", "header", "footer"] {
        if (number(&sheet.nodes[margins], key, -1.0, 0.0, 2.0)? * 72.0
            - numeric(&settings["marginsPt"][key], 0.0, 144.0)?)
        .abs()
            > 1e-9
        {
            return Err(crate::coded(
                "internal-mismatch",
                "Print margins differ from settings",
            ));
        }
    }
    let setup = next.tags[2]
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing print page setup"))?;
    if sheet.nodes[setup]
        .attrs
        .get("paperSize")
        .map(String::as_str)
        != Some(if paper == "A4" { "9" } else { "1" })
        || sheet.nodes[setup]
            .attrs
            .get("orientation")
            .map(String::as_str)
            != Some(orientation)
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Print paper/orientation differs",
        ));
    }
    for (key, value) in if fit {
        vec![
            ("fitToWidth", settings["scaling"]["width"].as_f64().unwrap()),
            (
                "fitToHeight",
                settings["scaling"]["height"].as_f64().unwrap(),
            ),
        ]
    } else {
        vec![("scale", settings["scaling"]["percent"].as_f64().unwrap())]
    } {
        if number(&sheet.nodes[setup], key, -1.0, 0.0, 400.0)? != value {
            return Err(crate::coded("internal-mismatch", "Print scaling differs"));
        }
    }
    let quote = format!("'{}'!", name.replace('\'', "''"));
    let area_formula = format!(
        "{quote}${}${}:${}${}",
        crate::xml::address(0, area.0.1).trim_end_matches('1'),
        area.0.0 + 1,
        crate::xml::address(0, area.1.1).trim_end_matches('1'),
        area.1.0 + 1
    );
    let mut wanted = BTreeMap::new();
    if has_area {
        wanted.insert("_xlnm.Print_Area".to_string(), area_formula);
    }
    let mut titles = Vec::new();
    if !settings["repeatRows"].is_null() {
        allowed_fields(&settings["repeatRows"], &["first", "last"])?;
        let first = whole(&settings["repeatRows"]["first"], 1, 1048576)?;
        let last = whole(&settings["repeatRows"]["last"], first, first + 99)?;
        if first != u64::from(area.0.0) + 1 || last >= u64::from(area.1.0) + 1 {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid repeat row prefix",
            ));
        }
        titles.push(format!("{quote}${first}:${last}"));
    }
    if !settings["repeatColumns"].is_null() {
        allowed_fields(&settings["repeatColumns"], &["first", "last"])?;
        let first = field(&settings["repeatColumns"], "first")?;
        let last = field(&settings["repeatColumns"], "last")?;
        let a = position(&format!("{first}1"))?.1;
        let b = position(&format!("{last}1"))?.1;
        if a != area.0.1 || b < a || b >= area.1.1 || b - a > 99 {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid repeat column prefix",
            ));
        }
        titles.push(format!("{quote}${first}:${last}"));
    }
    if !titles.is_empty() {
        wanted.insert("_xlnm.Print_Titles".into(), titles.join(","));
    }
    let actual = next
        .names
        .iter()
        .map(|index| {
            let node = &workbook.nodes[*index];
            Ok((
                node.attrs
                    .get("name")
                    .ok_or_else(|| crate::coded("internal-missing-field", "Missing print name"))?
                    .clone(),
                node.text.clone(),
            ))
        })
        .collect::<Result<BTreeMap<_, _>>>()?;
    if actual != wanted {
        return Err(crate::coded(
            "internal-mismatch",
            "Reserved print names differ from settings",
        ));
    }
    // Only pageSetUpPr inside sheetPr may change. Its unrelated attributes,
    // outline/tab-color descendants and lexical content must survive exactly.
    let new_pr = next
        .sheet_pr
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing proposed sheetPr"))?;
    let mut new_envelope = container_mask(&sheet, new_pr, &[setup_pr])?;
    if let Some(old_pr) = previous.sheet_pr {
        let old = container_mask(
            &source_sheet,
            old_pr,
            &previous.setup_pr.into_iter().collect::<Vec<_>>(),
        )?;
        if old != new_envelope {
            return Err(crate::coded("print-stale", "Unrelated sheetPr XML changed"));
        }
    } else {
        let node = &sheet.nodes[new_pr];
        if node
            .attrs
            .keys()
            .any(|key| key != "xmlns" && !key.starts_with("xmlns:"))
            || node.children.len() != 1
            || !node.text.trim().is_empty()
        {
            return Err(crate::coded(
                "print-unsupported",
                "New sheetPr contains unrelated metadata",
            ));
        }
        new_envelope.clear();
    }
    // Preserve the actual accepted export's unrelated workbook bytes, including
    // calculation-readiness flags. The prior guard above still binds the
    // immutable owner metadata, not an arbitrary caller-supplied definition.
    let mut exported = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, field(input, "source")?)?).map_err(err)?,
    ))
    .map_err(err)?;
    let exported_workbook = Tree::parse(&part_bytes(&mut exported, "xl/workbook.xml", &empty)?)?;
    let exported_previous = definition(&source_sheet, &exported_workbook, part, name, scope)?;
    protect_source(
        &source_sheet,
        &exported_workbook,
        &exported_previous,
        name,
        scope,
    )?;
    let name_definitions = |tree: &Tree, indexes: &[usize]| {
        indexes
            .iter()
            .map(|index| {
                let node = &tree.nodes[*index];
                (
                    node.attrs.get("name").cloned(),
                    node.attrs.get("localSheetId").cloned(),
                    node.text.clone(),
                )
            })
            .collect::<Vec<_>>()
    };
    if name_definitions(&source_workbook, &previous.names)
        != name_definitions(&exported_workbook, &exported_previous.names)
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Exported print definition differs from its accepted owner",
        ));
    }
    if let Some(old_container) = exported_previous.container {
        let new_container = next
            .container
            .ok_or_else(|| crate::coded("print-refused", "Removed defined-name container"))?;
        if container_mask(&exported_workbook, old_container, &exported_previous.names)?
            != container_mask(&workbook, new_container, &next.names)?
        {
            return Err(crate::coded(
                "print-stale",
                "Unrelated defined names changed",
            ));
        }
        if exported_workbook.mask(&[old_container])? != workbook.mask(&[new_container])? {
            return Err(crate::coded(
                "print-stale",
                "Unrelated workbook XML changed",
            ));
        }
    } else if let Some(container) = next.container {
        if workbook.nodes[container].children.len() != next.names.len()
            || !workbook.nodes[container].text.trim().is_empty()
            || workbook.nodes[container]
                .attrs
                .keys()
                .any(|key| key != "xmlns" && !key.starts_with("xmlns:"))
            || workbook.mask(&[container])? != exported_workbook.bytes
        {
            return Err(crate::coded(
                "print-stale",
                "New print names changed unrelated workbook metadata",
            ));
        }
    } else if !wanted.is_empty() || workbook.bytes != exported_workbook.bytes {
        return Err(crate::coded(
            "print-stale",
            "Print clearing changed unrelated workbook metadata",
        ));
    }
    Ok(())
}
fn visual_attrs(tree: &Tree, index: usize, fields: &[&str]) -> Result<()> {
    if tree.nodes[index]
        .attrs
        .keys()
        .any(|key| key != "xmlns" && !key.starts_with("xmlns:") && !fields.contains(&key.as_str()))
    {
        return Err(crate::coded(
            "print-unsupported",
            "Opaque print style attributes require qualification",
        ));
    }
    Ok(())
}
fn visual_color(
    tree: &Tree,
    index: usize,
    black_only: bool,
    theme: Option<&super::print_theme::PrintTheme>,
) -> Result<()> {
    if tree.nodes[index].attrs.contains_key("theme") {
        return super::print_theme::color(tree, index, theme, black_only);
    }
    tree.clean(index, &["rgb", "auto"])?;
    let node = &tree.nodes[index];
    if node.attrs.get("auto").is_some_and(|value| value != "1") {
        return Err(crate::coded(
            "print-unsupported",
            "Unqualified automatic print color",
        ));
    }
    if let Some(rgb) = node.attrs.get("rgb") {
        if ![6, 8].contains(&rgb.len())
            || !rgb.bytes().all(|ch| ch.is_ascii_hexdigit())
            || rgb.len() == 8 && !rgb.starts_with("FF")
            || black_only && !rgb.ends_with("000000")
        {
            return Err(crate::coded(
                "print-unsupported",
                "Opaque print color requires qualification",
            ));
        }
    }
    Ok(())
}
pub(super) fn qualified_styles(
    bytes: Option<Vec<u8>>,
    theme: Option<&super::print_theme::PrintTheme>,
) -> Result<Vec<bool>> {
    let Some(bytes) = bytes else {
        return Ok(vec![true]);
    };
    let tree = Tree::parse(&bytes)?;
    tree.root(NS, "styleSheet")?;
    let items = |name: &str| -> Result<Vec<usize>> {
        Ok(tree
            .direct(0, name)?
            .map(|index| tree.nodes[index].children.clone())
            .unwrap_or_default())
    };
    let fonts = items("fonts")?;
    let fills = items("fills")?;
    let borders = items("borders")?;
    let xfs = items("cellXfs")?;
    let formats = items("numFmts")?
        .into_iter()
        .map(|index| {
            let node = &tree.nodes[index];
            Ok((
                node.attrs
                    .get("numFmtId")
                    .ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing print number format id")
                    })?
                    .parse::<u64>()
                    .map_err(err)?,
                node.attrs
                    .get("formatCode")
                    .ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing print format code")
                    })?
                    .clone(),
            ))
        })
        .collect::<Result<BTreeMap<_, _>>>()?;
    let font_ok = fonts
        .iter()
        .map(|index| -> Result<()> {
            let node = &tree.nodes[*index];
            if node.ns != NS || node.name != "font" {
                return Err(crate::coded("print-unsupported", "Foreign print font"));
            }
            visual_attrs(&tree, *index, &[])?;
            for child in &node.children {
                let element = &tree.nodes[*child];
                if element.ns != NS {
                    return Err(crate::coded(
                        "print-unsupported",
                        "Opaque print font namespace",
                    ));
                }
                match element.name.as_str() {
                    "name" | "sz" | "b" | "i" | "u" => {
                        tree.clean(*child, &["val"])?;
                        if ["b", "i"].contains(&element.name.as_str())
                            && element
                                .attrs
                                .get("val")
                                .is_some_and(|value| !["1", "true"].contains(&value.as_str()))
                            || element.name == "u"
                                && element
                                    .attrs
                                    .get("val")
                                    .is_some_and(|value| value != "single")
                        {
                            return Err(crate::coded(
                                "print-unsupported",
                                "Unqualified print font decoration",
                            ));
                        }
                    }
                    "color" => visual_color(&tree, *child, false, theme)?,
                    "scheme" => super::print_theme::scheme(&tree, *child, theme)?,
                    "family" | "charset" => {
                        tree.clean(*child, &["val"])?;
                    }
                    _ => {
                        return Err(crate::coded(
                            "print-unsupported",
                            "Rich/theme print font requires qualification",
                        ));
                    }
                }
            }
            Ok(())
        })
        .map(|result| result.is_ok())
        .collect::<Vec<_>>();
    let fill_ok = fills
        .iter()
        .map(|index| -> Result<()> {
            let node = &tree.nodes[*index];
            if node.ns != NS || node.name != "fill" {
                return Err(crate::coded("print-unsupported", "Foreign print fill"));
            }
            visual_attrs(&tree, *index, &[])?;
            let pattern = tree.direct(*index, "patternFill")?.ok_or_else(|| {
                crate::coded(
                    "print-unsupported",
                    "Gradient print fill requires qualification",
                )
            })?;
            if node.children.len() != 1 {
                return Err(crate::coded("print-unsupported", "Opaque print fill"));
            }
            visual_attrs(&tree, pattern, &["patternType"])?;
            let pattern_type = tree.nodes[pattern]
                .attrs
                .get("patternType")
                .map(String::as_str)
                .unwrap_or("none");
            if pattern_type == "none" && tree.nodes[pattern].children.is_empty() {
                return Ok(());
            }
            if pattern_type != "solid" {
                return Err(crate::coded(
                    "print-unsupported",
                    "Pattern print fill requires qualification",
                ));
            }
            let fg = tree.direct(pattern, "fgColor")?.ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing print fill color")
            })?;
            visual_color(&tree, fg, false, theme)?;
            for child in &tree.nodes[pattern].children {
                if *child != fg {
                    if tree.nodes[*child].name != "bgColor" {
                        return Err(crate::coded("print-unsupported", "Opaque print fill child"));
                    }
                    tree.clean(*child, &["indexed"])?;
                    if tree.nodes[*child].attrs.get("indexed").map(String::as_str) != Some("64") {
                        return Err(crate::coded(
                            "print-unsupported",
                            "Opaque print fill background",
                        ));
                    }
                }
            }
            Ok(())
        })
        .map(|result| result.is_ok())
        .collect::<Vec<_>>();
    let border_ok = borders
        .iter()
        .map(|index| -> Result<()> {
            let node = &tree.nodes[*index];
            if node.ns != NS || node.name != "border" {
                return Err(crate::coded("print-unsupported", "Foreign print border"));
            }
            visual_attrs(&tree, *index, &[])?;
            let mut painted = 0;
            for child in &node.children {
                let side = &tree.nodes[*child];
                if side.ns != NS
                    || !["left", "right", "top", "bottom", "diagonal"].contains(&side.name.as_str())
                {
                    return Err(crate::coded("print-unsupported", "Opaque print border"));
                }
                visual_attrs(&tree, *child, &["style"])?;
                if let Some(style) = side.attrs.get("style") {
                    if style != "thin" || side.name == "diagonal" {
                        return Err(crate::coded(
                            "print-unsupported",
                            "Alternative print border style",
                        ));
                    }
                    painted += 1;
                }
                for color in &side.children {
                    if tree.nodes[*color].name != "color" {
                        return Err(crate::coded(
                            "print-unsupported",
                            "Opaque print border child",
                        ));
                    }
                    visual_color(&tree, *color, true, theme)?;
                }
            }
            if painted != 0 && painted != 4 {
                return Err(crate::coded(
                    "print-unsupported",
                    "Partial print borders require qualification",
                ));
            }
            Ok(())
        })
        .map(|result| result.is_ok())
        .collect::<Vec<_>>();
    xfs.iter()
        .map(|index| -> Result<bool> {
            let node = &tree.nodes[*index];
            if node.ns != NS || node.name != "xf" {
                return Ok(false);
            }
            if visual_attrs(
                &tree,
                *index,
                &[
                    "numFmtId",
                    "fontId",
                    "fillId",
                    "borderId",
                    "xfId",
                    "applyNumberFormat",
                    "applyFont",
                    "applyFill",
                    "applyBorder",
                    "applyAlignment",
                    "applyProtection",
                ],
            )
            .is_err()
            {
                return Ok(false);
            }
            let id = |key: &str| -> Result<usize> {
                Ok(node
                    .attrs
                    .get(key)
                    .map(|value| value.parse::<usize>().map_err(err))
                    .transpose()?
                    .unwrap_or(0))
            };
            if !font_ok.get(id("fontId")?).copied().unwrap_or(false)
                || !fill_ok.get(id("fillId")?).copied().unwrap_or(false)
                || !border_ok.get(id("borderId")?).copied().unwrap_or(false)
            {
                return Ok(false);
            }
            let format = id("numFmtId")? as u64;
            if ![0, 1, 2, 3, 4, 9, 10, 14, 44, 49].contains(&format)
                && !formats.get(&format).is_some_and(|code| {
                    ["0.0", "0.00", "yyyy-mm-dd", "$#,##0.00", "@"].contains(&code.as_str())
                })
            {
                return Ok(false);
            }
            for child in &node.children {
                let element = &tree.nodes[*child];
                if element.ns != NS {
                    return Ok(false);
                }
                if element.name == "alignment" {
                    if tree
                        .clean(
                            *child,
                            &[
                                "horizontal",
                                "vertical",
                                "wrapText",
                                "textRotation",
                                "indent",
                                "shrinkToFit",
                            ],
                        )
                        .is_err()
                        || element.attrs.get("horizontal").is_some_and(|value| {
                            !["general", "left", "right", "center"].contains(&value.as_str())
                        })
                        || element.attrs.get("vertical").is_some_and(|value| {
                            !["top", "bottom", "center"].contains(&value.as_str())
                        })
                        || element
                            .attrs
                            .get("textRotation")
                            .is_some_and(|value| value != "0")
                        || element
                            .attrs
                            .get("indent")
                            .is_some_and(|value| value != "0")
                        || boolean(element.attrs.get("shrinkToFit"), false)?
                    {
                        return Ok(false);
                    }
                    boolean(element.attrs.get("wrapText"), false)?;
                } else if element.name == "protection" {
                    if tree.clean(*child, &["locked", "hidden"]).is_err() {
                        return Ok(false);
                    }
                    boolean(element.attrs.get("locked"), true)?;
                    boolean(element.attrs.get("hidden"), false)?;
                } else {
                    return Ok(false);
                }
            }
            Ok(true)
        })
        .collect()
}
/// Source XML is not represented fully by range values. Prove plain cell
/// content before a readonly snapshot can consume that flattened projection.
pub(super) fn check(root: &Path, input: &Value) -> Result<Value> {
    use rusqlite::params;
    use std::io::{Seek, SeekFrom};
    let range = field(input, "range")?;
    let (first, last) = crate::bulk::bounds(range)?;
    let count = u64::from(last.0 - first.0 + 1) * u64::from(last.1 - first.1 + 1);
    if count > 1000 {
        return Err(crate::coded(
            "print-refused",
            "Print format checks use <=1,000 positions per page",
        ));
    }
    let version = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing print version"))?;
    let part = field(input, "sheet")?;
    let db = crate::database(root)?;
    let effective = crate::range::read_public_db(
        &db,
        &json!({"version":version,"sheet":part,"range":range,"offset":0,"limit":1001}),
    )?;
    if !effective["nextOffset"].is_null() {
        return Err(crate::coded(
            "print-limit",
            "Print source page exceeded its bound",
        ));
    }
    if effective["cells"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing print cells"))?
        .iter()
        .any(|cell| cell[1]["calculation"]["pending"] == true)
    {
        return Err(crate::coded(
            "print-bad-input",
            "PRINT_PENDING_CALCULATION: calculate this page's stored formula or derived results before printing",
        ));
    }
    let index: Value =
        serde_json::from_slice(&fs::read(local(root, "index.json")?).map_err(err)?).map_err(err)?;
    let sheet = index["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing print source index"))?
        .iter()
        .find(|sheet| sheet["path"] == part)
        .ok_or_else(|| crate::coded("print-bad-input", "Print worksheet is not indexed"))?;
    if sheet["extended"] == true {
        return Err(crate::coded(
            "print-unsupported",
            "Special/opaque source worksheet print requires qualification",
        ));
    }
    // Source counts alone miss authored masters and hidden off-page followers.
    // Admit only complete current numeric ownership after a whole-sheet scan.
    let view = if input["view"].is_object() {
        &input["view"]
    } else {
        &index
    };
    let qualified_spills = crate::spill::printing::qualify(
        root,
        &db,
        version,
        part,
        view,
        &crate::spill::Geometry { first, last },
    )?;
    let mut metadata = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let empty = serde_json::Map::new();
    let meta_sheet = Tree::parse(&part_bytes(&mut metadata, part, &empty)?)?;
    meta_sheet.root(NS, "worksheet")?;
    for node in &meta_sheet.nodes {
        if node.name == "extLst" {
            return Err(crate::coded(
                "print-unsupported",
                "Extended print formatting requires qualification",
            ));
        }
    }
    let style_bytes = if metadata.file_names().any(|part| part == "xl/styles.xml") {
        Some(part_bytes(&mut metadata, "xl/styles.xml", &empty)?)
    } else {
        None
    };
    let page_theme = super::print_theme::owned(&mut metadata)?;
    let rule_styles = style_bytes
        .as_ref()
        .map(|bytes| Tree::parse(bytes))
        .transpose()?;
    if let Some(styles) = &rule_styles {
        styles.root(NS, "styleSheet")?;
    }
    super::rules::qualify_print(&meta_sheet, rule_styles.as_ref())?;
    let normal_font = rule_styles
        .as_ref()
        .map(|tree| super::print_theme::normal_font(tree, page_theme.as_ref()))
        .transpose()?
        .flatten();
    let styles = qualified_styles(style_bytes, page_theme.as_ref())?;
    let hidden_rows = sheet["hiddenRows"]
        .as_array()
        .map(|rows| {
            rows.iter()
                .filter_map(Value::as_u64)
                .collect::<BTreeSet<_>>()
        })
        .unwrap_or_default();
    let hidden_columns = sheet["hiddenColumns"]
        .as_array()
        .cloned()
        .unwrap_or_default();
    let visible = |r: u32, c: u32| {
        !hidden_rows.contains(&u64::from(r))
            && !hidden_columns.iter().any(|range| {
                range[0]
                    .as_u64()
                    .is_some_and(|first| u64::from(c) + 1 >= first)
                    && range[1]
                        .as_u64()
                        .is_some_and(|last| u64::from(c) + 1 <= last)
            })
    };
    for record in effective["cells"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing print source page"))?
    {
        let at = position(record[0].as_str().ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing print cell address")
        })?)?;
        if !visible(at.0, at.1) {
            continue;
        }
        let cell = &record[1];
        let proof =
            qualified_spills.get(record[0].as_str().ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing print reference")
            })?);
        if cell
            .get("spill")
            .is_some_and(|spill| proof.is_none_or(|proof| proof.get("spill") != Some(spill)))
            || cell.get("formulaType").is_some_and(|kind| kind != "shared") && proof.is_none()
            || cell.get("scenarioResult").is_some()
        {
            return Err(crate::coded("print-unsupported", "PRINT_SPECIAL_FORMULA"));
        }
        let style = cell["styleId"].as_u64().unwrap_or(0) as usize;
        if !styles.get(style).copied().unwrap_or(false) {
            return Err(crate::coded(
                "print-unsupported",
                format!("PRINT_UNSUPPORTED_STYLE {}", record[0]),
            ));
        }
    }
    let mut source = File::open(local(root, field(sheet, "artifact")?)?).map_err(err)?;
    let mut source_cells=db.prepare_cached("SELECT r,c,start,end,kind,raw FROM cells WHERE sheet=?1 AND r BETWEEN ?2 AND ?3 AND c BETWEEN ?4 AND ?5 ORDER BY r,c").map_err(err)?;
    let mut source_rows = source_cells
        .query(params![part, first.0, last.0, first.1, last.1])
        .map_err(err)?;
    let mut shared = BTreeSet::new();
    let mut source_bytes = 0u64;
    while let Some(row) = source_rows.next().map_err(err)? {
        crate::cancel::check()?;
        let r: u32 = row.get(0).map_err(err)?;
        let c: u32 = row.get(1).map_err(err)?;
        if !visible(r, c) {
            continue;
        }
        let start: u64 = row.get(2).map_err(err)?;
        let end: u64 = row.get(3).map_err(err)?;
        if end < start || end - start > 65536 {
            return Err(crate::coded(
                "print-limit",
                "Print source cell exceeds its bound",
            ));
        }
        source_bytes = source_bytes
            .checked_add(end - start)
            .ok_or_else(|| crate::coded("print-limit", "Print source byte overflow"))?;
        if source_bytes > 8 * 1024 * 1024 {
            return Err(crate::coded(
                "print-limit",
                "Print source page exceeds the 8 MiB proof bound",
            ));
        }
        source.seek(SeekFrom::Start(start)).map_err(err)?;
        let mut bytes = vec![0; usize::try_from(end - start).map_err(err)?];
        source.read_exact(&mut bytes).map_err(err)?;
        let text = std::str::from_utf8(&bytes).map_err(err)?;
        let wrapped = format!(
            "<worksheet xmlns=\"{}\">{text}</worksheet>",
            String::from_utf8_lossy(NS)
        );
        let tree = Tree::parse(wrapped.as_bytes())?;
        tree.root(NS, "worksheet")?;
        let cell = tree
            .direct(0, "c")?
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing source print cell"))?;
        if tree.nodes[0].children.len() != 1 {
            return Err(crate::coded(
                "print-unsupported",
                "Opaque source print cell",
            ));
        }
        let proof = qualified_spills.get(&address(r, c));
        visual_attrs(
            &tree,
            cell,
            if proof.is_some_and(|proof| proof["kind"] == "master" || proof["kind"] == "consumer") {
                &["r", "s", "t", "cm"]
            } else {
                &["r", "s", "t"]
            },
        )?;
        let mut seen = BTreeSet::new();
        for child in &tree.nodes[cell].children {
            let node = &tree.nodes[*child];
            if node.ns != NS || !seen.insert(node.name.as_str()) {
                return Err(crate::coded(
                    "print-unsupported",
                    "Rich/opaque print cell requires qualification",
                ));
            }
            match node.name.as_str() {
                "v" => {
                    visual_attrs(&tree, *child, &[])?;
                    if !node.children.is_empty() {
                        return Err(crate::coded(
                            "print-unsupported",
                            "Opaque source print value",
                        ));
                    }
                }
                "f" => {
                    visual_attrs(&tree, *child, &["t", "si", "ref", "ca", "aca"])?;
                    if !node.children.is_empty()
                        || node.attrs.get("t").is_some_and(|kind| kind != "shared")
                            && proof.is_none()
                    {
                        return Err(crate::coded(
                            "print-unsupported",
                            "Special source print formula",
                        ));
                    }
                }
                "is" => {
                    visual_attrs(&tree, *child, &[])?;
                    if node.children.len() != 1 {
                        return Err(crate::coded(
                            "print-unsupported",
                            "Rich inline print text requires qualification",
                        ));
                    }
                    let text = node.children[0];
                    if tree.nodes[text].ns != NS
                        || tree.nodes[text].name != "t"
                        || !tree.nodes[text].children.is_empty()
                    {
                        return Err(crate::coded(
                            "print-unsupported",
                            "Rich inline print text requires qualification",
                        ));
                    }
                    visual_attrs(
                        &tree,
                        text,
                        &["{http://www.w3.org/XML/1998/namespace}space"],
                    )?;
                }
                _ => {
                    return Err(crate::coded(
                        "print-unsupported",
                        "Opaque source print content requires qualification",
                    ));
                }
            }
        }
        let kind: String = row.get(4).map_err(err)?;
        if kind == "s" {
            let raw: String = row.get(5).map_err(err)?;
            shared.insert(raw.parse::<usize>().map_err(err)?);
        }
    }
    if !shared.is_empty() {
        let mut zip = ZipArchive::new(crate::cancel::Reader(
            File::open(local(root, "source.xlsx")?).map_err(err)?,
        ))
        .map_err(err)?;
        let mut entry = zip.by_name("xl/sharedStrings.xml").map_err(err)?;
        if entry.size() > 8 * 1024 * 1024 {
            return Err(crate::coded(
                "print-limit",
                "Print shared-string proof exceeds its 8 MiB bound",
            ));
        }
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).map_err(err)?;
        let tree = Tree::parse(&bytes)?;
        tree.root(NS, "sst")?;
        for index in shared {
            let si = *tree.nodes[0].children.get(index).ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing print shared string")
            })?;
            if tree.nodes[si].ns != NS
                || tree.nodes[si].name != "si"
                || tree.nodes[si].children.len() != 1
            {
                return Err(crate::coded(
                    "print-unsupported",
                    "Rich shared print text requires qualification",
                ));
            }
            visual_attrs(&tree, si, &[])?;
            let text = tree.nodes[si].children[0];
            if tree.nodes[text].ns != NS
                || tree.nodes[text].name != "t"
                || !tree.nodes[text].children.is_empty()
            {
                return Err(crate::coded(
                    "print-unsupported",
                    "Rich shared print text requires qualification",
                ));
            }
            visual_attrs(
                &tree,
                text,
                &["{http://www.w3.org/XML/1998/namespace}space"],
            )?;
        }
    }
    let special_cells = qualified_spills
        .into_iter()
        .filter(|(reference, _)| {
            position(reference).is_ok_and(|at| {
                first.0 <= at.0 && at.0 <= last.0 && first.1 <= at.1 && at.1 <= last.1
            })
        })
        .collect::<BTreeMap<_, _>>();
    Ok(
        json!({"qualified":true,"version":version,"sheet":part,"range":range,"positions":count,"specialCells":special_cells,"normalFont":normal_font,"pageTheme":page_theme.map(|theme|json!({"sample":theme.sample,"revision":theme.revision}))}),
    )
}

#[cfg(test)]
mod plain_header_footer_tests {
    use super::*;
    #[test]
    fn literal_sections_decode_without_promoting_fields() {
        assert_eq!(
            plain_sections("&LR&&D&CSales&R😀").unwrap(),
            serde_json::json!({"left":"R&D","center":"Sales","right":"😀"})
        );
        assert_eq!(
            plain_sections("&CLiteral &&P").unwrap()["center"],
            "Literal &P"
        );
        assert!(plain_sections("&C&P").is_err());
        assert!(plain_sections("&C&G").is_err());
        assert!(plain_sections("&C&BStyled").is_err());
    }
    #[test]
    fn bounds_and_extended_settings_refuse() {
        let mut value = serde_json::json!({"header":{"left":"","center":"Title","right":""},"footer":{"left":"","center":"","right":""},"scaleWithDoc":true});
        validate_plain_header_footer(&value).unwrap();
        value["header"]["center"] = Value::String("&".repeat(127));
        assert!(validate_plain_header_footer(&value).is_err());
        value["header"]["center"] = Value::String("bad\nline".into());
        assert!(validate_plain_header_footer(&value).is_err());
        value["header"]["center"] = Value::String("Title".into());
        value["scaleWithDoc"] = Value::String("true".into());
        assert!(validate_plain_header_footer(&value).is_err());
    }
    #[test]
    fn namespaced_plain_header_accepts_while_foreign_first_and_formatting_refuse() {
        let xml = format!(
            "<headerFooter xmlns=\"{}\" scaleWithDoc=\"0\"><oddHeader>&amp;CTitle</oddHeader><oddFooter>&amp;RConfidential</oddFooter></headerFooter>",
            String::from_utf8_lossy(NS)
        );
        let tree = Tree::parse(xml.as_bytes()).unwrap();
        let value = plain_header_footer(&tree, 0).unwrap();
        assert_eq!(value["header"]["center"], "Title");
        assert_eq!(value["scaleWithDoc"], false);
        for candidate in [
            xml.replace("scaleWithDoc=\"0\"", "differentFirst=\"1\""),
            xml.replace("&amp;CTitle", "&amp;C&amp;BTitle"),
            xml.replace("<oddHeader>", "<oddHeader xmlns=\"urn:foreign\">"),
        ] {
            assert!(plain_header_footer(&Tree::parse(candidate.as_bytes()).unwrap(), 0).is_err());
        }
    }
}
