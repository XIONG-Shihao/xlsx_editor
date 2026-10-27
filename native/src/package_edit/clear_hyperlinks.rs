//! Native-owned worksheet hyperlink removal. Rectangle subtraction is declared
//! editor policy; partial-range behavior has not been observed in source Excel.
//! No cells, styles, formulas, targets or unrelated package members are edited.
use super::{print_settings::Tree, *};
use std::collections::{BTreeMap, BTreeSet};
const SELECTOR_LIMIT: usize = 10_000;
const LINK_LIMIT: usize = 10_000;
const FRAGMENT_LIMIT: usize = 40_000;
const WORK_LIMIT: usize = 2_000_000;
const COMMAND_LIMIT: usize = 2 * 1024 * 1024;
const HYPERLINK_KIND: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct Rect {
    r0: u32,
    c0: u32,
    r1: u32,
    c1: u32,
}
impl Rect {
    fn reference(self) -> String {
        let a = crate::xml::address(self.r0, self.c0);
        let b = crate::xml::address(self.r1, self.c1);
        if a == b { a } else { format!("{a}:{b}") }
    }
    fn overlaps(self, b: Self) -> bool {
        self.r0 <= b.r1 && b.r0 <= self.r1 && self.c0 <= b.c1 && b.c0 <= self.c1
    }
    fn subtract(self, b: Self) -> Vec<Self> {
        if !self.overlaps(b) {
            return vec![self];
        }
        let (r0, c0, r1, c1) = (
            self.r0.max(b.r0),
            self.c0.max(b.c0),
            self.r1.min(b.r1),
            self.c1.min(b.c1),
        );
        let mut out = Vec::new();
        if self.r0 < r0 {
            out.push(Self { r1: r0 - 1, ..self });
        }
        if r1 < self.r1 {
            out.push(Self { r0: r1 + 1, ..self });
        }
        if self.c0 < c0 {
            out.push(Self {
                r0,
                r1,
                c1: c0 - 1,
                ..self
            });
        }
        if c1 < self.c1 {
            out.push(Self {
                r0,
                r1,
                c0: c1 + 1,
                ..self
            });
        }
        out
    }
}
pub(super) fn point(text: &str) -> Result<(u32, u32)> {
    let bytes = text.as_bytes();
    let mut at = usize::from(bytes.first() == Some(&b'$'));
    let first = at;
    while bytes.get(at).is_some_and(u8::is_ascii_alphabetic) {
        at += 1;
    }
    if at == first || at - first > 3 {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid hyperlink cell reference",
        ));
    }
    if bytes.get(at) == Some(&b'$') {
        at += 1;
    }
    if !bytes.get(at).is_some_and(|b| matches!(*b, b'1'..=b'9'))
        || !bytes[at..].iter().all(u8::is_ascii_digit)
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid hyperlink cell reference",
        ));
    }
    crate::xml::position(text)
}
fn rectangle(text: &str, canonical: bool) -> Result<Rect> {
    let mut parts = text.split(':');
    let a =
        point(parts.next().ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing hyperlink reference")
        })?)?;
    let b = parts.next().map(point).transpose()?.unwrap_or(a);
    if parts.next().is_some() || a.0 > b.0 || a.1 > b.1 {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid hyperlink rectangle",
        ));
    }
    let r = Rect {
        r0: a.0,
        c0: a.1,
        r1: b.0,
        c1: b.1,
    };
    if canonical && r.reference() != text {
        return Err(crate::coded(
            "hyperlink-bad-input",
            "Clear Hyperlinks requires canonical rectangles",
        ));
    }
    Ok(r)
}
/// Sweep rectangles, not cells. Active column intervals are disjoint, so one
/// predecessor query proves each insertion; full-sheet geometry stays compact.
fn disjoint(rectangles: &[Rect]) -> Result<()> {
    let mut events = Vec::with_capacity(rectangles.len() * 2);
    for r in rectangles {
        events.push((r.r0, true, r.c0, r.c1));
        events.push((r.r1 + 1, false, r.c0, r.c1));
    }
    events.sort();
    let mut active = BTreeMap::new();
    for (_, add, c0, c1) in events {
        crate::cancel::check()?;
        if add {
            if active
                .range(..=c1)
                .next_back()
                .is_some_and(|(_, end)| *end >= c0)
            {
                return Err(crate::coded(
                    "hyperlink-conflict",
                    "Overlapping hyperlink rectangles are ambiguous",
                ));
            }
            active.insert(c0, c1);
        } else if active.remove(&c0) != Some(c1) {
            return Err(crate::coded(
                "hyperlink-ambiguous",
                "Ambiguous hyperlink rectangle sweep",
            ));
        }
    }
    Ok(())
}
fn command(command: &Value) -> Result<(String, Vec<Rect>)> {
    if command.as_object().is_none_or(|m| {
        m.len() != 3
            || m.keys()
                .any(|k| !["type", "sheet", "ranges"].contains(&k.as_str()))
    }) || command["type"] != "clearHyperlinks"
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid Clear Hyperlinks command fields",
        ));
    }
    if serde_json::to_vec(command).map_err(err)?.len() > COMMAND_LIMIT {
        return Err(crate::coded(
            "hyperlink-limit",
            "Clear Hyperlinks command exceeds editor byte bound",
        ));
    }
    let sheet = field(command, "sheet")?;
    if sheet.is_empty() || sheet.encode_utf16().count() > 31 {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid Clear Hyperlinks worksheet",
        ));
    }
    let ranges = command["ranges"]
        .as_array()
        .filter(|a| !a.is_empty() && a.len() <= SELECTOR_LIMIT)
        .ok_or_else(|| {
            crate::coded(
                "hyperlink-bad-input",
                "Clear Hyperlinks requires bounded ranges",
            )
        })?;
    let ranges = ranges
        .iter()
        .map(|v| {
            rectangle(
                v.as_str().ok_or_else(|| {
                    crate::coded("internal-invalid-field", "Invalid hyperlink selector")
                })?,
                true,
            )
        })
        .collect::<Result<Vec<_>>>()?;
    disjoint(&ranges)?;
    Ok((sheet.into(), ranges))
}
fn bounded_append(out: &mut String, piece: &str) -> Result<()> {
    if piece.len() > XML_LIMIT.saturating_sub(out.len()) {
        return Err(crate::coded(
            "hyperlink-limit",
            "Clear Hyperlinks metadata output exceeds editor bound",
        ));
    }
    out.push_str(piece);
    Ok(())
}
fn replace_reference(raw: &str, reference: &str) -> Result<String> {
    let bytes = raw.as_bytes();
    let mut at = 1;
    while at < bytes.len() && !bytes[at].is_ascii_whitespace() && !matches!(bytes[at], b'>' | b'/')
    {
        at += 1;
    }
    loop {
        while bytes.get(at).is_some_and(u8::is_ascii_whitespace) {
            at += 1;
        }
        if bytes.get(at).is_none_or(|b| matches!(*b, b'>' | b'/')) {
            break;
        }
        let start = at;
        while bytes
            .get(at)
            .is_some_and(|b| !b.is_ascii_whitespace() && *b != b'=')
        {
            at += 1;
        }
        let key = &bytes[start..at];
        while bytes.get(at).is_some_and(u8::is_ascii_whitespace) {
            at += 1;
        }
        if bytes.get(at) != Some(&b'=') {
            return Err(crate::coded(
                "hyperlink-malformed",
                "Invalid hyperlink attribute envelope",
            ));
        }
        at += 1;
        while bytes.get(at).is_some_and(u8::is_ascii_whitespace) {
            at += 1;
        }
        let quote = *bytes
            .get(at)
            .filter(|b| matches!(**b, b'\'' | b'"'))
            .ok_or_else(|| {
                crate::coded("hyperlink-malformed", "Invalid hyperlink attribute quote")
            })?;
        at += 1;
        let start_value = at;
        while bytes.get(at).is_some_and(|b| *b != quote) {
            at += 1;
        }
        if at == bytes.len() {
            return Err(crate::coded(
                "hyperlink-malformed",
                "Unterminated hyperlink attribute",
            ));
        }
        if key == b"ref" {
            return Ok(format!(
                "{}{}{}",
                &raw[..start_value],
                reference,
                &raw[at..]
            ));
        }
        at += 1;
    }
    Err(crate::coded(
        "internal-missing-field",
        "Missing hyperlink reference envelope",
    ))
}
#[derive(Default, Clone, Debug, PartialEq, Eq)]
pub(super) struct Counts {
    pub removed: u64,
    pub trimmed: u64,
    pub relationships: u64,
    pub remaining: u64,
}
pub(super) struct Prepared {
    pub plan: Value,
    pub command: Value,
    pub counts: Counts,
}
impl Prepared {
    pub fn result(&self) -> Value {
        json!({"schema":workbook_format!(SCHEMA_CLEAR_HYPERLINKS_WRITER_V1),"file":"package.xlsx","contentChanged":false,"changed":self.counts.removed+self.counts.trimmed>0,"removedHyperlinks":self.counts.removed,"trimmedHyperlinks":self.counts.trimmed,"removedRelationships":self.counts.relationships,"remainingHyperlinks":self.counts.remaining})
    }
}
pub(super) fn package_names(
    zip: &ZipArchive<crate::cancel::Reader<File>>,
) -> Result<BTreeSet<String>> {
    let mut names = BTreeSet::new();
    for name in zip.file_names() {
        valid_part(name)?;
        if !names.insert(name.into()) {
            return Err(crate::coded(
                "hyperlink-duplicate",
                "Duplicate Clear Hyperlinks package member",
            ));
        }
    }
    Ok(names)
}
fn relationship_part(sheet: &str) -> Result<String> {
    let (dir, file) = sheet
        .rsplit_once('/')
        .ok_or_else(|| crate::coded("hyperlink-malformed", "Invalid hyperlink worksheet part"))?;
    Ok(format!("{dir}/_rels/{file}.rels"))
}
pub(super) fn prepare_from(root: &Path, source: &str, command_value: &Value) -> Result<Prepared> {
    let (sheet, ranges) = command(command_value)?;
    let inventory = crate::used_range_transfer::worksheet_inventory(&local(root, source)?)?;
    let selected = inventory.iter().find(|s| s.name == sheet).ok_or_else(|| {
        crate::coded(
            "internal-unknown-reference",
            "Unknown Clear Hyperlinks worksheet",
        )
    })?;
    let part = &selected.path;
    let mut metadata = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let names = package_names(&metadata)?;
    let tree = Tree::parse(&part_bytes(&mut metadata, part, &serde_json::Map::new())?)?;
    tree.root(NS, "worksheet")?;
    if tree.direct(0, "sheetProtection")?.is_some() {
        return Err(crate::coded(
            "hyperlink-protected",
            "Clear Hyperlinks worksheet is protected",
        ));
    }
    if tree
        .nodes
        .iter()
        .any(|n| ["hyperlinks", "hyperlink"].contains(&n.name.as_str()) && n.ns != NS)
    {
        return Err(crate::coded(
            "hyperlink-unsupported",
            "Extended hyperlink graph is unqualified",
        ));
    }
    let collection = tree.direct(0, "hyperlinks")?;
    let mut counts = Counts::default();
    let mut fragments = serde_json::Map::new();
    let mut replacements = serde_json::Map::new();
    let mut links = Vec::new();
    if let Some(index) = collection {
        let node = &tree.nodes[index];
        if !node.text.trim().is_empty()
            || node
                .attrs
                .keys()
                .any(|k| k != "xmlns" && !k.starts_with("xmlns:"))
        {
            return Err(crate::coded(
                "hyperlink-unsupported",
                "Opaque hyperlink collection",
            ));
        }
        if node.children.len() > LINK_LIMIT {
            return Err(crate::coded(
                "hyperlink-limit",
                "Hyperlink source record editor bound exceeded",
            ));
        }
        for child in &node.children {
            let n = &tree.nodes[*child];
            if n.ns != NS || n.name != "hyperlink" {
                return Err(crate::coded(
                    "hyperlink-unsupported",
                    "Foreign hyperlink record",
                ));
            }
            tree.clean(*child, &["ref", "r:id", "location", "display", "tooltip"])?;
            let rect = rectangle(
                n.attrs.get("ref").ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing hyperlink reference")
                })?,
                false,
            )?;
            if n.attrs.get("r:id").is_none_or(String::is_empty)
                && n.attrs.get("location").is_none_or(String::is_empty)
            {
                return Err(crate::coded(
                    "hyperlink-unsupported",
                    "Hyperlink has no qualified destination",
                ));
            }
            if n.attrs.get("r:id").is_some_and(String::is_empty)
                || n.attrs.get("location").is_some_and(String::is_empty)
            {
                return Err(crate::coded(
                    "hyperlink-bad-input",
                    "Empty hyperlink destination",
                ));
            }
            links.push((*child, rect));
        }
        disjoint(&links.iter().map(|(_, r)| *r).collect::<Vec<_>>())?;
    }
    let rel_part = relationship_part(part)?;
    let relationship_tree = if names.contains(&rel_part) {
        Some(Tree::parse(&part_bytes(
            &mut metadata,
            &rel_part,
            &serde_json::Map::new(),
        )?)?)
    } else {
        None
    };
    let mut edges = BTreeMap::new();
    if let Some(rels) = &relationship_tree {
        rels.root(PACKAGE_REL_NS, "Relationships")?;
        let root = &rels.nodes[0];
        if !root.text.trim().is_empty()
            || root
                .attrs
                .keys()
                .any(|k| k != "xmlns" && !k.starts_with("xmlns:"))
        {
            return Err(crate::coded(
                "hyperlink-unsupported",
                "Opaque hyperlink relationships",
            ));
        }
        for i in &root.children {
            let n = &rels.nodes[*i];
            if n.ns != PACKAGE_REL_NS || n.name != "Relationship" {
                return Err(crate::coded(
                    "hyperlink-unsupported",
                    "Foreign hyperlink relationship",
                ));
            }
            rels.clean(*i, &["Id", "Type", "Target", "TargetMode"])?;
            let id = n.attrs.get("Id").filter(|s| !s.is_empty()).ok_or_else(|| {
                crate::coded(
                    "internal-missing-field",
                    "Missing hyperlink relationship ID",
                )
            })?;
            let kind = n
                .attrs
                .get("Type")
                .filter(|s| !s.is_empty())
                .ok_or_else(|| {
                    crate::coded(
                        "internal-missing-field",
                        "Missing hyperlink relationship type",
                    )
                })?;
            let target = n
                .attrs
                .get("Target")
                .filter(|s| !s.is_empty())
                .ok_or_else(|| {
                    crate::coded(
                        "internal-missing-field",
                        "Missing hyperlink relationship target",
                    )
                })?;
            let scheme = kind.split_once(':').map(|(scheme, tail)| (scheme, tail));
            if scheme.is_none_or(|(scheme, tail)| {
                scheme.is_empty()
                    || tail.is_empty()
                    || !scheme.as_bytes()[0].is_ascii_alphabetic()
                    || !scheme
                        .bytes()
                        .all(|b| b.is_ascii_alphanumeric() || b"+.-".contains(&b))
            }) || kind.chars().any(char::is_whitespace)
            {
                return Err(crate::coded(
                    "hyperlink-malformed",
                    "Invalid hyperlink relationship type URI",
                ));
            }
            match n.attrs.get("TargetMode").map(String::as_str) {
                Some("External") => {}
                None | Some("Internal") => {
                    let target = resolve(part, target)?;
                    if !names.contains(&target) {
                        return Err(crate::coded(
                            "hyperlink-unsupported",
                            "Orphan hyperlink internal target",
                        ));
                    }
                }
                _ => {
                    return Err(crate::coded(
                        "internal-unknown-reference",
                        "Unknown hyperlink relationship mode",
                    ));
                }
            }
            if edges.insert(id.clone(), (*i, kind.clone())).is_some() {
                return Err(crate::coded(
                    "hyperlink-duplicate",
                    "Duplicate hyperlink relationship ID",
                ));
            }
        }
    }
    for (index, _) in &links {
        if let Some(id) = tree.nodes[*index].attrs.get("r:id") {
            if edges.get(id).is_none_or(|(_, kind)| kind != HYPERLINK_KIND) {
                return Err(crate::coded(
                    "hyperlink-refused",
                    "Hyperlink relationship missing or wrong type",
                ));
            }
        }
    }
    let mut work = 0usize;
    let mut replacement_bytes = 0usize;
    let mut edits = Vec::new();
    let mut whole = BTreeSet::new();
    let mut candidates = BTreeSet::new();
    let mut retained = BTreeSet::new();
    for (index, original) in &links {
        let mut residual = vec![*original];
        for target in &ranges {
            let mut next = Vec::new();
            for piece in residual {
                work += 1;
                if work > WORK_LIMIT {
                    return Err(crate::coded(
                        "hyperlink-limit",
                        "Clear Hyperlinks geometry work editor bound exceeded",
                    ));
                }
                next.extend(piece.subtract(*target));
                if next.len() > FRAGMENT_LIMIT {
                    return Err(crate::coded(
                        "hyperlink-limit",
                        "Clear Hyperlinks fragment editor bound exceeded",
                    ));
                }
            }
            residual = next;
            if residual.is_empty() {
                break;
            }
        }
        counts.remaining += residual.len() as u64;
        if counts.remaining > FRAGMENT_LIMIT as u64 {
            return Err(crate::coded(
                "hyperlink-limit",
                "Clear Hyperlinks fragment editor bound exceeded",
            ));
        }
        let node = &tree.nodes[*index];
        if residual == vec![*original] {
            if let Some(id) = node.attrs.get("r:id") {
                retained.insert(id.clone());
            }
            continue;
        }
        if let Some(id) = node.attrs.get("r:id") {
            candidates.insert(id.clone());
            if !residual.is_empty() {
                retained.insert(id.clone());
            }
        }
        let mut replacement = String::new();
        if residual.is_empty() {
            counts.removed += 1;
            whole.insert(*index);
        } else {
            counts.trimmed += 1;
            let raw = std::str::from_utf8(tree.raw(*index)).map_err(err)?;
            for r in residual {
                if raw.len() + 32 > XML_LIMIT.saturating_sub(replacement_bytes + replacement.len())
                {
                    return Err(crate::coded(
                        "hyperlink-limit",
                        "Clear Hyperlinks fragment byte bound exceeded",
                    ));
                }
                bounded_append(&mut replacement, &replace_reference(raw, &r.reference())?)?;
            }
        }
        replacement_bytes += replacement.len();
        edits.push((node.start, node.end, replacement));
    }
    // Preserve a candidate edge even for an opaque extension reference: only
    // demonstrably unused IDs may be deleted. Namespace aliases are resolved by Tree.
    for (index, node) in tree.nodes.iter().enumerate() {
        if !whole.contains(&index) {
            for value in node.attrs.values() {
                for token in value.split_whitespace() {
                    if candidates.contains(token) {
                        retained.insert(token.into());
                    }
                }
            }
            for token in node.text.split_whitespace() {
                if candidates.contains(token) {
                    retained.insert(token.into());
                }
            }
        }
    }
    if !edits.is_empty() {
        let index = collection
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing hyperlink container"))?;
        let node = &tree.nodes[index];
        let mut output = String::new();
        let mut previous = node.start;
        for (start, end, text) in &edits {
            bounded_append(
                &mut output,
                std::str::from_utf8(&tree.bytes[previous..*start]).map_err(err)?,
            )?;
            bounded_append(&mut output, text)?;
            previous = *end;
        }
        bounded_append(
            &mut output,
            std::str::from_utf8(&tree.bytes[previous..node.end]).map_err(err)?,
        )?;
        if counts.remaining == 0 {
            output.clear();
        }
        fragments.insert(part.clone(), json!({"hyperlinks":output}));
    }
    // Metadata strips sheetData. Stream the full owner too so an extended cell
    // reference cannot make a still-owned edge look exclusively unused. Only
    // candidate IDs are retained; no cell/attribute collection is projected.
    let mut source_zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, source)?).map_err(err)?,
    ))
    .map_err(err)?;
    package_names(&source_zip)?;
    let (_, source_links, source_references) =
        worksheet_hash(&mut source_zip, part, Some(&candidates))?;
    retained.extend(source_references);
    let expected = collection.map(|i| tree.raw(i)).unwrap_or_default();
    if source_links != expected {
        return Err(crate::coded(
            "internal-mismatch",
            "Hyperlink export source differs from native metadata",
        ));
    }
    let removable = candidates
        .difference(&retained)
        .map(|id| {
            edges
                .get(id)
                .map(|(index, _)| *index)
                .ok_or_else(|| crate::coded("hyperlink-refused", "Lost hyperlink relationship"))
        })
        .collect::<Result<Vec<_>>>()?;
    if !removable.is_empty() {
        let rels = relationship_tree.as_ref().ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing hyperlink relationship part",
            )
        })?;
        counts.relationships = removable.len() as u64;
        let masked = rels.mask(&removable)?;
        replacements.insert(rel_part, json!(String::from_utf8(masked).map_err(err)?));
    }
    Ok(Prepared {
        plan: json!({"kind":"clearHyperlinks","replacements":replacements,"removed":[],"copies":{},"fragments":fragments}),
        command: command_value.clone(),
        counts,
    })
}
pub(super) fn prepare(root: &Path, input: &Value) -> Result<Prepared> {
    let fields = input.as_object().ok_or_else(|| {
        crate::coded(
            "internal-invalid-field",
            "Invalid Clear Hyperlinks native request",
        )
    })?;
    let required = [
        "action",
        "operation",
        "root",
        "version",
        "source",
        "rewriteInputId",
        "command",
    ];
    if fields
        .keys()
        .filter(|k| k.as_str() != "diagnostics")
        .map(String::as_str)
        .collect::<BTreeSet<_>>()
        != required.into_iter().collect()
        || fields.get("diagnostics").is_some_and(|v| !v.is_boolean())
        || input["action"] != "workbook"
        || input["operation"] != "package-edit"
        || input["source"] != "export.xlsx"
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid Clear Hyperlinks native request fields",
        ));
    }
    prepare_from(root, "export.xlsx", &input["command"])
}
struct DigestWriter(Sha256);
impl Write for DigestWriter {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.0.update(bytes);
        Ok(bytes.len())
    }
    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}
pub(super) fn worksheet_hash(
    zip: &mut ZipArchive<crate::cancel::Reader<File>>,
    part: &str,
    candidates: Option<&BTreeSet<String>>,
) -> Result<(Vec<u8>, Vec<u8>, BTreeSet<String>)> {
    worksheet_tag_hash(zip, part, candidates, b"hyperlinks")
}
pub(super) fn worksheet_tag_hash(
    zip: &mut ZipArchive<crate::cancel::Reader<File>>,
    part: &str,
    candidates: Option<&BTreeSet<String>>,
    tag: &[u8],
) -> Result<(Vec<u8>, Vec<u8>, BTreeSet<String>)> {
    let member = zip.by_name(part).map_err(err)?;
    let mut reader = NsReader::from_reader(BufReader::new(member));
    let mut envelope = Writer::new(DigestWriter(Sha256::new()));
    let mut links = Writer::new(Vec::new());
    let mut depth = 0usize;
    let mut capture = None;
    let mut seen = false;
    let mut root = false;
    let mut buffer = Vec::new();
    let mut references = BTreeSet::new();
    loop {
        crate::cancel::check()?;
        let (ns, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        let main = ns == quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(NS));
        match &event {
            Event::DocType(_) => {
                return Err(crate::coded(
                    "hyperlink-malformed",
                    "Hyperlink worksheet entities are unqualified",
                ));
            }
            Event::Start(e) | Event::Empty(e) => {
                let empty = matches!(event, Event::Empty(_));
                if depth == 0 {
                    if root || !main || e.local_name().as_ref() != b"worksheet" {
                        return Err(crate::coded(
                            "hyperlink-malformed",
                            "Invalid hyperlink worksheet root",
                        ));
                    }
                    root = true;
                }
                if depth == 1 && main && e.local_name().as_ref() == tag {
                    if seen {
                        return Err(crate::coded(
                            "hyperlink-duplicate",
                            "Duplicate hyperlink source collection",
                        ));
                    }
                    seen = true;
                    capture = Some(depth);
                }
                if capture.is_none()
                    && let Some(candidates) = candidates
                {
                    for attr in e.attributes().with_checks(true) {
                        let attr = attr.map_err(err)?;
                        let value = attr
                            .decode_and_unescape_value(reader.decoder())
                            .map_err(err)?;
                        for token in value.split_whitespace() {
                            if candidates.contains(token) {
                                references.insert(token.into());
                            }
                        }
                    }
                }
                if !empty {
                    depth += 1;
                }
            }
            _ => {}
        }
        if matches!(event, Event::Eof) {
            break;
        }
        if capture.is_some() {
            // Event borrows the read buffer. Its byte slice plus exact XML
            // delimiters bounds the append without copying or reborrowing buffer.
            let overhead = match &event {
                Event::Start(_) => 2,
                Event::End(_) | Event::Empty(_) => 3,
                Event::Comment(_) => 7,
                Event::CData(_) => 12,
                Event::Decl(_) | Event::PI(_) => 4,
                Event::DocType(_) => 11,
                Event::Text(_) | Event::Eof => 0,
            };
            let event_bytes = event
                .len()
                .checked_add(overhead)
                .ok_or_else(|| crate::coded("hyperlink-limit", "Hyperlink event size overflow"))?;
            if event_bytes > XML_LIMIT.saturating_sub(links.get_ref().len()) {
                return Err(crate::coded(
                    "hyperlink-limit",
                    "Hyperlink source collection byte bound exceeded",
                ));
            }
            links.write_event(event.borrow()).map_err(err)?;
        } else {
            envelope.write_event(event.borrow()).map_err(err)?;
        }
        if matches!(event, Event::End(_)) {
            depth = depth.checked_sub(1).ok_or_else(|| {
                crate::coded(
                    "internal-invalid-field",
                    "Invalid hyperlink worksheet depth",
                )
            })?;
            if capture == Some(depth) {
                capture = None;
            }
        } else if matches!(event, Event::Empty(_)) && capture == Some(depth) {
            capture = None;
        }
        buffer.clear();
    }
    if !root || depth != 0 || capture.is_some() {
        return Err(crate::coded(
            "hyperlink-unsupported",
            "Incomplete hyperlink worksheet source",
        ));
    }
    Ok((
        envelope.into_inner().0.finalize().to_vec(),
        links.into_inner(),
        references,
    ))
}
pub(super) fn member_hash(
    zip: &mut ZipArchive<crate::cancel::Reader<File>>,
    part: &str,
) -> Result<Vec<u8>> {
    let mut member = zip.by_name(part).map_err(err)?;
    let mut digest = Sha256::new();
    let mut buffer = [0u8; 32768];
    loop {
        crate::cancel::check()?;
        let count = member.read(&mut buffer).map_err(err)?;
        if count == 0 {
            break;
        }
        digest.update(&buffer[..count]);
    }
    Ok(digest.finalize().to_vec())
}
/// Independent package proof: every cell/row/style/rule/opaque member is equal;
/// only the derived link collection and its exclusively unused edges may differ.
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
    let names = package_names(&before)?;
    if names != package_names(&after)? {
        return Err(crate::coded(
            "hyperlink-stale",
            "Clear Hyperlinks changed package inventory",
        ));
    }
    let sheets = prepared.plan["fragments"].as_object().ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing native hyperlink fragments",
        )
    })?;
    let replacements = prepared.plan["replacements"].as_object().ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing native hyperlink relationships",
        )
    })?;
    for part in names {
        crate::cancel::check()?;
        if let Some(fragment) = sheets.get(&part) {
            let (old_envelope, _, _) = worksheet_hash(&mut before, &part, None)?;
            let (new_envelope, new_links, _) = worksheet_hash(&mut after, &part, None)?;
            if old_envelope != new_envelope
                || new_links != field(fragment, "hyperlinks")?.as_bytes()
            {
                return Err(crate::coded(
                    "hyperlink-stale",
                    "Clear Hyperlinks changed nonlink worksheet fields or link policy",
                ));
            }
        } else if let Some(expected) = replacements.get(&part) {
            if member_hash(&mut after, &part)?
                != Sha256::digest(
                    expected
                        .as_str()
                        .ok_or_else(|| {
                            crate::coded(
                                "hyperlink-malformed",
                                "Invalid native hyperlink relationship replacement",
                            )
                        })?
                        .as_bytes(),
                )
                .to_vec()
            {
                return Err(crate::coded(
                    "internal-mismatch",
                    "Clear Hyperlinks relationship output differs from native plan",
                ));
            }
        } else if member_hash(&mut before, &part)? != member_hash(&mut after, &part)? {
            return Err(crate::coded(
                "hyperlink-stale",
                "Clear Hyperlinks changed unrelated package member",
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
pub(crate) mod tests;
