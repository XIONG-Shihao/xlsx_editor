//! Selected-source merge/table barriers for fill commands only.
use super::{Bounds, bounds};
use crate::xml::NS;
use crate::{Result, err, field, local};
use quick_xml::{NsReader, events::Event, name::ResolveResult};
use serde_json::Value;
use std::{
    collections::{BTreeMap, BTreeSet},
    fs::File,
    io::Read,
    path::Path,
};
use zip::ZipArchive;

const REL: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL: &[u8] = b"http://schemas.openxmlformats.org/package/2006/relationships";
const TABLE_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
const READ_LIMIT: usize = 16 * 1024 * 1024;
#[derive(Clone)]
pub(crate) struct Node {
    pub(crate) namespace: Vec<u8>,
    pub(crate) name: String,
    pub(crate) attrs: BTreeMap<String, String>,
}

pub(crate) fn walk(
    bytes: &[u8],
    namespace: &[u8],
    root: &str,
    mut visit: impl FnMut(&[Node], &Node) -> Result<()>,
) -> Result<()> {
    let mut reader = NsReader::from_reader(bytes);
    let mut buffer = Vec::new();
    let mut stack: Vec<Node> = Vec::new();
    let mut roots = 0;
    let mut nodes = 0;
    loop {
        crate::cancel::check()?;
        let (binding, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        let empty = matches!(event, Event::Empty(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                nodes += 1;
                if nodes > 50000 || stack.len() >= 64 {
                    return Err(crate::coded(
                        "fill-limit",
                        "Fill metadata exceeds its node/depth bound",
                    ));
                }
                let ns = match binding {
                    ResolveResult::Bound(ns) => ns.as_ref().to_vec(),
                    ResolveResult::Unbound => Vec::new(),
                    _ => {
                        return Err(crate::coded(
                            "fill-unsupported",
                            "Unbound fill metadata namespace",
                        ));
                    }
                };
                let mut attributes = BTreeMap::new();
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(err)?;
                    let raw = std::str::from_utf8(attribute.key.as_ref()).map_err(err)?;
                    let (binding, name) = reader.resolve_attribute(attribute.key);
                    let key = match binding {
                        ResolveResult::Bound(ns) if ns.as_ref() == REL => {
                            format!("r:{}", String::from_utf8_lossy(name.as_ref()))
                        }
                        ResolveResult::Bound(ns)
                            if ns.as_ref() == b"http://www.w3.org/XML/1998/namespace" =>
                        {
                            format!("xml:{}", String::from_utf8_lossy(name.as_ref()))
                        }
                        ResolveResult::Bound(ns)
                            if ns.as_ref() != b"http://www.w3.org/2000/xmlns/" =>
                        {
                            format!(
                                "{{{}}}{}",
                                String::from_utf8_lossy(ns.as_ref()),
                                String::from_utf8_lossy(name.as_ref())
                            )
                        }
                        ResolveResult::Unknown(_) if !raw.starts_with("xmlns:") => {
                            return Err(crate::coded(
                                "fill-unsupported",
                                "Unbound fill metadata attribute",
                            ));
                        }
                        ResolveResult::Unbound
                            if raw.contains(':') && !raw.starts_with("xmlns:") =>
                        {
                            return Err(crate::coded(
                                "fill-unsupported",
                                "Unbound fill metadata attribute",
                            ));
                        }
                        _ => raw.to_owned(),
                    };
                    if attributes
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
                            "fill-duplicate",
                            "Duplicate fill metadata attribute",
                        ));
                    }
                }
                let node = Node {
                    namespace: ns,
                    name: std::str::from_utf8(element.local_name().as_ref())
                        .map_err(err)?
                        .into(),
                    attrs: attributes,
                };
                if stack.is_empty() {
                    roots += 1;
                    if roots != 1 || node.namespace != namespace || node.name != root {
                        return Err(crate::coded(
                            "fill-unsupported",
                            "Fill metadata root is unqualified",
                        ));
                    }
                }
                if stack.last().is_some_and(|node| {
                    ["mergeCell", "tablePart", "Relationship"].contains(&node.name.as_str())
                }) {
                    return Err(crate::coded(
                        "fill-protected",
                        "Nested fill geometry metadata is protected",
                    ));
                }
                visit(&stack, &node)?;
                if !empty {
                    stack.push(node);
                }
            }
            Event::End(_) => {
                stack
                    .pop()
                    .ok_or_else(|| crate::coded("fill-malformed", "Unbalanced fill metadata"))?;
            }
            Event::Text(text) => {
                if (stack.is_empty()
                    || stack.last().is_some_and(|node| {
                        [
                            "mergeCells",
                            "mergeCell",
                            "tableParts",
                            "tablePart",
                            "Relationships",
                            "Relationship",
                        ]
                        .contains(&node.name.as_str())
                    }))
                    && !text.unescape().map_err(err)?.trim().is_empty()
                {
                    return Err(crate::coded(
                        "fill-protected",
                        "Opaque fill geometry text is protected",
                    ));
                }
            }
            Event::CData(_) => {
                if stack.is_empty()
                    || stack.last().is_some_and(|node| {
                        [
                            "mergeCells",
                            "mergeCell",
                            "tableParts",
                            "tablePart",
                            "Relationships",
                            "Relationship",
                        ]
                        .contains(&node.name.as_str())
                    })
                {
                    return Err(crate::coded(
                        "fill-protected",
                        "Opaque fill geometry text is protected",
                    ));
                }
            }
            Event::DocType(_) => {
                return Err(crate::coded(
                    "fill-malformed",
                    "Fill metadata XML entities are protected",
                ));
            }
            Event::Eof => break,
            _ => (),
        }
        buffer.clear();
    }
    if roots != 1 || !stack.is_empty() {
        return Err(crate::coded("fill-unsupported", "Incomplete fill metadata"));
    }
    Ok(())
}
fn attributes(node: &Node, allowed: &[&str]) -> Result<()> {
    if node
        .attrs
        .keys()
        .any(|key| key != "xmlns" && !key.starts_with("xmlns:") && !allowed.contains(&key.as_str()))
    {
        return Err(crate::coded(
            "fill-protected",
            "Extended fill geometry attributes are protected",
        ));
    }
    Ok(())
}
fn count(node: &Node) -> Result<Option<usize>> {
    node.attrs
        .get("count")
        .map(|value| {
            if value.is_empty() || !value.bytes().all(|b| b.is_ascii_digit()) {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid fill geometry count",
                ));
            }
            value
                .parse::<u32>()
                .map(|value| value as usize)
                .map_err(err)
        })
        .transpose()
}
pub(crate) fn part(
    zip: &mut ZipArchive<crate::cancel::Reader<File>>,
    name: &str,
    read: &mut usize,
) -> Result<Vec<u8>> {
    let mut member = zip.by_name(name).map_err(err)?;
    if member.size() > READ_LIMIT.saturating_sub(*read) as u64 {
        return Err(crate::coded(
            "fill-limit",
            "Fill metadata exceeds its read bound",
        ));
    }
    let mut bytes = Vec::new();
    member.read_to_end(&mut bytes).map_err(err)?;
    *read = read
        .checked_add(bytes.len())
        .ok_or_else(|| crate::coded("fill-limit", "Fill metadata size overflow"))?;
    if *read > READ_LIMIT {
        return Err(crate::coded(
            "fill-limit",
            "Fill metadata exceeds its read bound",
        ));
    }
    Ok(bytes)
}
fn target(owner: &str, target: &str) -> Result<String> {
    if target.is_empty()
        || target.contains(['\\', ':', '?', '#', '%'])
        || target.bytes().any(|b| b < 32)
    {
        return Err(crate::coded(
            "fill-unsupported",
            "Table target is outside the qualified package profile",
        ));
    }
    let mut parts = if target.starts_with('/') {
        Vec::new()
    } else {
        owner
            .rsplit_once('/')
            .ok_or_else(|| crate::coded("fill-malformed", "Invalid worksheet part"))?
            .0
            .split('/')
            .collect::<Vec<_>>()
    };
    for item in target.trim_start_matches('/').split('/') {
        match item {
            ".." => {
                parts.pop().ok_or_else(|| {
                    crate::coded("fill-unsupported", "Table target escapes package")
                })?;
            }
            "" | "." => return Err(crate::coded("fill-malformed", "Noncanonical table target")),
            _ => parts.push(item),
        }
    }
    Ok(parts.join("/"))
}
#[derive(Clone)]
struct Barrier {
    kind: &'static str,
    area: Bounds,
    /// Only writing here is a problem; a fill may read from it.
    writes_only: bool,
}
fn inventory(
    zip: &mut ZipArchive<crate::cancel::Reader<File>>,
    members: &BTreeSet<String>,
    sheet: &str,
    read: &mut usize,
) -> Result<(Vec<Barrier>, BTreeSet<String>)> {
    let worksheet = part(zip, sheet, read)?;
    let mut barriers = Vec::new();
    let mut table_ids = BTreeSet::new();
    let mut merge_count = None;
    let mut table_count = None;
    let mut merges = 0;
    let mut merge_containers = 0;
    let mut table_containers = 0;
    walk(&worksheet, NS, "worksheet", |stack, node| {
        if !["mergeCells", "mergeCell", "tableParts", "tablePart"].contains(&node.name.as_str()) {
            return Ok(());
        }
        if node.namespace != NS {
            return Err(crate::coded(
                "fill-protected",
                "Foreign fill geometry namespace is protected",
            ));
        }
        match node.name.as_str() {
            "mergeCells" | "tableParts" => {
                if stack.len() != 1 {
                    return Err(crate::coded(
                        "fill-protected",
                        "Nested fill geometry container is protected",
                    ));
                }
                attributes(node, &["count"])?;
                if node.name == "mergeCells" {
                    merge_containers += 1;
                    merge_count = count(node)?;
                } else {
                    table_containers += 1;
                    table_count = count(node)?;
                }
                if merge_containers > 1 || table_containers > 1 {
                    return Err(crate::coded(
                        "fill-duplicate",
                        "Duplicate fill geometry container",
                    ));
                }
            }
            "mergeCell" => {
                if stack.len() != 2 || stack[1].name != "mergeCells" || stack[1].namespace != NS {
                    return Err(crate::coded(
                        "fill-protected",
                        "Orphan merge geometry is protected",
                    ));
                }
                attributes(node, &["ref"])?;
                barriers.push(Barrier {
                    kind: "merged cells",
                    area: bounds(node.attrs.get("ref").ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing merge range")
                    })?)?,
                    writes_only: false,
                });
                merges += 1;
            }
            "tablePart" => {
                if stack.len() != 2 || stack[1].name != "tableParts" || stack[1].namespace != NS {
                    return Err(crate::coded(
                        "fill-protected",
                        "Orphan table geometry is protected",
                    ));
                }
                attributes(node, &["r:id"])?;
                let id = node
                    .attrs
                    .get("r:id")
                    .filter(|id| !id.is_empty())
                    .ok_or_else(|| {
                        crate::coded(
                            "internal-missing-field",
                            "Missing table relationship identity",
                        )
                    })?;
                if !table_ids.insert(id.clone()) {
                    return Err(crate::coded(
                        "fill-duplicate",
                        "Duplicate table owner identity",
                    ));
                }
            }
            _ => (),
        }
        Ok(())
    })?;
    if merge_count.is_some_and(|count| count != merges)
        || table_count.is_some_and(|count| count != table_ids.len())
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Fill geometry count disagrees with source",
        ));
    }
    let (directory, file) = sheet
        .rsplit_once('/')
        .ok_or_else(|| crate::coded("fill-malformed", "Invalid worksheet part"))?;
    let rels = format!("{directory}/_rels/{file}.rels");
    if !members.contains(&rels) {
        if !table_ids.is_empty() {
            return Err(crate::coded(
                "internal-missing-field",
                "Missing table owner relationships",
            ));
        }
        return Ok((barriers, BTreeSet::new()));
    }
    let bytes = part(zip, &rels, read)?;
    let mut edges = BTreeMap::new();
    walk(&bytes, PACKAGE_REL, "Relationships", |stack, node| {
        if stack.is_empty() {
            attributes(node, &[])?;
            return Ok(());
        }
        if stack.len() != 1 || node.namespace != PACKAGE_REL || node.name != "Relationship" {
            return Err(crate::coded(
                "fill-unsupported",
                "Opaque worksheet relationship inventory",
            ));
        }
        attributes(node, &["Id", "Type", "Target", "TargetMode"])?;
        let id = node
            .attrs
            .get("Id")
            .filter(|id| !id.is_empty())
            .ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing relationship identity")
            })?;
        if !node.attrs.contains_key("Type")
            || !node.attrs.contains_key("Target")
            || edges.insert(id.clone(), node.attrs.clone()).is_some()
        {
            return Err(crate::coded(
                "fill-ambiguous",
                "Ambiguous worksheet relationship inventory",
            ));
        }
        Ok(())
    })?;
    for (id, edge) in &edges {
        if edge["Type"] == TABLE_REL && !table_ids.contains(id) {
            return Err(crate::coded(
                "fill-protected",
                "Orphan table relationship is protected",
            ));
        }
    }
    let mut table_parts = BTreeSet::new();
    for id in table_ids {
        let edge = edges.get(&id).ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing table owner relationship")
        })?;
        if edge["Type"] != TABLE_REL
            || edge
                .get("TargetMode")
                .is_some_and(|mode| mode != "Internal")
        {
            return Err(crate::coded(
                "fill-unsupported",
                "Table relationship is unqualified",
            ));
        }
        let name = target(sheet, &edge["Target"])?;
        if !members.contains(&name) || !table_parts.insert(name.clone()) {
            return Err(crate::coded(
                "internal-missing-field",
                "Missing or ambiguous table part",
            ));
        }
        let bytes = part(zip, &name, read)?;
        let mut geometry = None;
        let (mut header, mut totals) = (1u32, 0u32);
        walk(&bytes, NS, "table", |stack, node| {
            if stack.is_empty() {
                geometry = Some(bounds(node.attrs.get("ref").ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing table range")
                })?)?);
                let rows = |key: &str, default: u32| {
                    node.attrs
                        .get(key)
                        .and_then(|text| text.parse::<u32>().ok())
                        .unwrap_or(default)
                };
                header = rows("headerRowCount", 1);
                totals = rows("totalsRowCount", 0);
            }
            Ok(())
        })?;
        // Filling values into a table's data rows is ordinary work, and reading
        // from a table changes nothing. What must not be overwritten are the
        // header row, whose cells are the column names, and the totals row.
        let (first, last) = geometry
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing table geometry"))?;
        if header > 0 {
            barriers.push(Barrier {
                kind: "a table's header row",
                area: (first, (first.0 + header - 1, last.1)),
                writes_only: true,
            });
        }
        if totals > 0 && last.0 + 1 >= first.0 + totals {
            barriers.push(Barrier {
                kind: "a table's totals row",
                area: ((last.0 + 1 - totals, first.1), last),
                writes_only: true,
            });
        }
    }
    Ok((barriers, table_parts))
}
#[cfg(test)]
pub(super) fn guard(root: &Path, index: &Value, commands: &[Value]) -> Result<()> {
    guard_inner(root, index, commands, None)
}
pub(super) fn guard_version(
    root: &Path,
    index: &Value,
    commands: &[Value],
    db: &rusqlite::Connection,
    version: u64,
) -> Result<()> {
    guard_inner(root, index, commands, Some((db, version)))
}
fn guard_inner(
    root: &Path,
    index: &Value,
    commands: &[Value],
    selected: Option<(&rusqlite::Connection, u64)>,
) -> Result<()> {
    let mut requests = Vec::new();
    for command in commands {
        if !command["type"].as_str().is_some_and(|kind| {
            [
                "fillSeries",
                "fillRange",
                "pasteFormulas",
                "pasteValues",
                "fillAcrossSheets",
            ]
            .contains(&kind)
        }) {
            continue;
        }
        let sheet = field(command, "sheet")?;
        requests.push((
            sheet.to_owned(),
            bounds(field(command, "range")?)?,
            "destination",
        ));
        if command["type"] != "fillSeries" && command["type"] != "pasteValues" {
            requests.push((
                command["sourceSheet"].as_str().unwrap_or(sheet).to_owned(),
                bounds(field(
                    command,
                    if command["type"] == "fillAcrossSheets" {
                        "range"
                    } else {
                        "source"
                    },
                )?)?,
                "source",
            ));
        }
    }
    if requests.is_empty() {
        return Ok(());
    }
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let members = zip.file_names().map(str::to_owned).collect::<BTreeSet<_>>();
    if members.len() != zip.len() {
        return Err(crate::coded(
            "fill-ambiguous",
            "Ambiguous fill package members",
        ));
    }
    let sheets = index["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing indexed worksheets"))?;
    let mut inventories = BTreeMap::new();
    let mut table_owners = BTreeMap::new();
    let mut read = 0;
    for (sheet, area, role) in requests {
        if !sheets.iter().any(|entry| entry["path"] == sheet) {
            return Err(crate::coded(
                "internal-unknown-reference",
                "Unknown fill worksheet",
            ));
        }
        if !inventories.contains_key(&sheet) {
            let (mut barriers, table_parts) = inventory(&mut zip, &members, &sheet, &mut read)?;
            if let Some((db, version)) = selected
                && let Some(merges) = crate::merges::selected_override(db, version, &sheet)?
            {
                barriers.retain(|barrier| barrier.kind != "merged cells");
                for range in merges {
                    barriers.push(Barrier {
                        kind: "merged cells",
                        area: bounds(&range)?,
                        writes_only: false,
                    });
                }
            }
            for part in table_parts {
                if table_owners
                    .insert(part, sheet.clone())
                    .is_some_and(|owner| owner != sheet)
                {
                    return Err(crate::coded(
                        "fill-unsupported",
                        "A table part has shared worksheet ownership",
                    ));
                }
            }
            inventories.insert(sheet.clone(), barriers);
        }
        for barrier in &inventories[&sheet] {
            if barrier.writes_only && role == "source" {
                continue;
            }
            let (a, b) = area;
            let (c, d) = barrier.area;
            if a.0 <= d.0 && b.0 >= c.0 && a.1 <= d.1 && b.1 >= c.1 {
                return Err(crate::coded(
                    "fill-conflict",
                    format!("Fill {role} intersects {}", barrier.kind),
                ));
            }
        }
    }
    Ok(())
}

/// Across is a bounded cell-copy policy. Worksheet-owned notes, links, rules and
/// objects are deliberately not interpreted as contents or formats. These are
/// conservative whole-selected-worksheet barriers, unlike merge/table geometry.
pub(super) fn across_metadata(root: &Path, sheets: &[&str]) -> Result<u64> {
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let members = zip.file_names().map(str::to_owned).collect::<BTreeSet<_>>();
    if members.len() != zip.len() {
        return Err(crate::coded(
            "fill-ambiguous",
            "Ambiguous Across Worksheets package members",
        ));
    }
    let mut read = 0;
    let mut states = BTreeMap::new();
    let bytes = part(&mut zip, "xl/workbook.xml", &mut read)?;
    walk(&bytes, NS, "workbook", |stack, node| {
        if node.name == "workbookProtection" {
            return Err(crate::coded(
                "fill-protected",
                "Across Worksheets protected workbook metadata is not qualified",
            ));
        }
        if node.name == "sheet" {
            if node.namespace != NS || stack.len() != 2 || stack[1].name != "sheets" {
                return Err(crate::coded(
                    "fill-unsupported",
                    "Across Worksheets worksheet inventory is unqualified",
                ));
            }
            let id = node
                .attrs
                .get("r:id")
                .filter(|id| !id.is_empty())
                .ok_or_else(|| {
                    crate::coded(
                        "internal-missing-field",
                        "Missing Across Worksheets worksheet identity",
                    )
                })?;
            let state = node
                .attrs
                .get("state")
                .map(String::as_str)
                .unwrap_or("visible");
            if !["visible", "hidden", "veryHidden"].contains(&state)
                || states.insert(id.clone(), state.to_owned()).is_some()
            {
                return Err(crate::coded(
                    "fill-refused",
                    "Across Worksheets worksheet state is ambiguous",
                ));
            }
        }
        Ok(())
    })?;
    let bytes = part(&mut zip, "xl/_rels/workbook.xml.rels", &mut read)?;
    let mut visible = BTreeSet::new();
    let mut identities = BTreeSet::new();
    walk(&bytes, PACKAGE_REL, "Relationships", |stack, node| {
        if stack.is_empty() {
            return Ok(());
        }
        if stack.len() != 1 || node.name != "Relationship" || node.namespace != PACKAGE_REL {
            return Err(crate::coded(
                "fill-unsupported",
                "Across Worksheets workbook relationship inventory is unqualified",
            ));
        }
        let id = node
            .attrs
            .get("Id")
            .filter(|id| !id.is_empty())
            .ok_or_else(|| {
                crate::coded(
                    "internal-missing-field",
                    "Missing Across Worksheets workbook relationship",
                )
            })?;
        if !identities.insert(id.clone()) {
            return Err(crate::coded(
                "fill-refused",
                "Across Worksheets workbook relationships are ambiguous",
            ));
        }
        if states.get(id).is_some_and(|state| state == "visible") {
            if node.attrs.get("Type").is_none_or(|kind| kind != "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet") || node.attrs.get("TargetMode").is_some_and(|mode| mode != "Internal") { return Err(crate::coded("fill-unsupported", "Across Worksheets worksheet relationship is unqualified")); }
            let path = target(
                "xl/workbook.xml",
                node.attrs.get("Target").ok_or_else(|| {
                    crate::coded(
                        "internal-missing-field",
                        "Missing Across Worksheets worksheet target",
                    )
                })?,
            )?;
            if !members.contains(&path) || !visible.insert(path) {
                return Err(crate::coded(
                    "fill-refused",
                    "Across Worksheets worksheet target is ambiguous",
                ));
            }
        }
        Ok(())
    })?;
    if sheets.iter().any(|sheet| !visible.contains(*sheet)) {
        return Err(crate::coded(
            "fill-bad-input",
            "Across Worksheets source and destinations must be visible worksheets",
        ));
    }
    let mut styles = 0_u64;
    let bytes = part(&mut zip, "xl/styles.xml", &mut read)?;
    let mut containers = 0;
    let mut declared_styles = None;
    walk(&bytes, NS, "styleSheet", |stack, node| {
        if node.name == "cellXfs" {
            if node.namespace != NS || stack.len() != 1 {
                return Err(crate::coded(
                    "fill-unsupported",
                    "Across Worksheets style inventory is unqualified",
                ));
            }
            containers += 1;
            declared_styles = count(node)?;
        }
        if stack.len() == 2 && stack[1].name == "cellXfs" && node.name == "xf" {
            if node.namespace != NS {
                return Err(crate::coded(
                    "fill-unsupported",
                    "Across Worksheets style inventory is unqualified",
                ));
            }
            styles += 1;
        }
        Ok(())
    })?;
    if containers != 1 || styles == 0 || declared_styles.is_some_and(|count| count as u64 != styles)
    {
        return Err(crate::coded(
            "fill-bad-input",
            "Across Worksheets needs one complete cell style inventory",
        ));
    }
    for sheet in sheets {
        let bytes = part(&mut zip, sheet, &mut read)?;
        walk(&bytes, NS, "worksheet", |stack, node| {
            if node.namespace != NS {
                return Err(crate::coded(
                    "fill-unsupported",
                    "Across Worksheets extended worksheet metadata is not qualified",
                ));
            }
            if stack.len() == 1
                && ![
                    "sheetPr",
                    "dimension",
                    "sheetViews",
                    "sheetFormatPr",
                    "cols",
                    "sheetData",
                    "mergeCells",
                    "tableParts",
                    "pageMargins",
                    "pageSetup",
                    "printOptions",
                    "headerFooter",
                    "sheetCalcPr",
                ]
                .contains(&node.name.as_str())
            {
                return Err(crate::coded(
                    "fill-unsupported",
                    format!("Across Worksheets {} metadata is not qualified", node.name),
                ));
            }
            if [
                "sheetProtection",
                "conditionalFormatting",
                "dataValidations",
                "hyperlinks",
                "legacyDrawing",
                "legacyDrawingHF",
                "drawing",
                "extLst",
                "oleObjects",
                "controls",
                "pivotTableParts",
            ]
            .contains(&node.name.as_str())
            {
                return Err(crate::coded(
                    "fill-unsupported",
                    format!("Across Worksheets {} metadata is not qualified", node.name),
                ));
            }
            Ok(())
        })?;
        let (directory, file) = sheet
            .rsplit_once('/')
            .ok_or_else(|| crate::coded("fill-malformed", "Invalid Across Worksheets part"))?;
        let rels = format!("{directory}/_rels/{file}.rels");
        if members.contains(&rels) {
            let bytes = part(&mut zip, &rels, &mut read)?;
            walk(&bytes, PACKAGE_REL, "Relationships", |stack, node| {
                if stack.is_empty() {
                    attributes(node, &[])?;
                    return Ok(());
                }
                if stack.len() != 1 || node.namespace != PACKAGE_REL || node.name != "Relationship"
                {
                    return Err(crate::coded(
                        "fill-unsupported",
                        "Across Worksheets relationship inventory is unqualified",
                    ));
                }
                attributes(node, &["Id", "Type", "Target", "TargetMode"])?;
                if node.attrs.get("Type").is_none_or(|kind| kind != TABLE_REL) {
                    return Err(crate::coded(
                        "fill-unsupported",
                        "Across Worksheets note, link or object relationships are not qualified",
                    ));
                }
                Ok(())
            })?;
        }
    }
    Ok(styles)
}

/// Shared-string reads use the immutable source, since metadata.xlsx intentionally
/// omits string bodies. Only selected plain strings can enter this copy policy.
pub(super) fn across_shared_strings(
    root: &Path,
    db: &rusqlite::Connection,
    version: u64,
    sheet: &str,
    area: Bounds,
) -> Result<()> {
    let mut selected = BTreeSet::new();
    let mut query = db.prepare_cached("SELECT r,c,raw FROM cells WHERE sheet=?1 AND kind='s' AND r BETWEEN ?2 AND ?3 AND c BETWEEN ?4 AND ?5").map_err(err)?;
    let mut rows = query
        .query(rusqlite::params![
            sheet, area.0.0, area.1.0, area.0.1, area.1.1
        ])
        .map_err(err)?;
    while let Some(row) = rows.next().map_err(err)? {
        crate::cancel::check()?;
        let edit = super::inherited_edit(
            db,
            version,
            sheet,
            row.get(0).map_err(err)?,
            row.get(1).map_err(err)?,
        )?;
        if edit.get("value").is_some() || edit.get("formula").is_some() {
            continue;
        }
        let raw: String = row.get(2).map_err(err)?;
        selected.insert(raw.parse::<u64>().map_err(err)?);
    }
    if selected.is_empty() {
        return Ok(());
    }
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "source.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let mut read = 0;
    let bytes = part(&mut zip, "xl/sharedStrings.xml", &mut read)?;
    let mut count = 0_u64;
    let mut plain_texts = 0;
    walk(&bytes, NS, "sst", |stack, node| {
        if stack.len() == 1 && node.name != "si" {
            return Err(crate::coded(
                "fill-unsupported",
                "Across Worksheets shared string inventory is unqualified",
            ));
        }
        if stack.len() == 1 && node.name == "si" {
            if node.namespace != NS {
                return Err(crate::coded(
                    "fill-unsupported",
                    "Across Worksheets shared string identity is unqualified",
                ));
            }
            if selected.contains(&count) {
                attributes(node, &[])?;
            }
            count += 1;
            plain_texts = 0;
        } else if stack.len() >= 2 && stack[1].name == "si" && selected.contains(&(count - 1)) {
            if stack.len() != 2 || node.namespace != NS || node.name != "t" {
                return Err(crate::coded(
                    "fill-unsupported",
                    "Across Worksheets rich or extended shared strings are not qualified",
                ));
            }
            attributes(node, &["xml:space"])?;
            if node
                .attrs
                .get("xml:space")
                .is_some_and(|value| !["preserve", "default"].contains(&value.as_str()))
            {
                return Err(crate::coded(
                    "fill-unsupported",
                    "Across Worksheets shared string whitespace is unqualified",
                ));
            }
            plain_texts += 1;
            if plain_texts > 1 {
                return Err(crate::coded(
                    "fill-unsupported",
                    "Across Worksheets duplicate shared string text is unqualified",
                ));
            }
        }
        Ok(())
    })?;
    if selected.last().is_some_and(|id| *id >= count) {
        return Err(crate::coded(
            "internal-missing-field",
            "Across Worksheets shared string identity is missing",
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;
    use std::io::Write;
    use zip::{ZipWriter, write::SimpleFileOptions};
    const SHEET: &str = "xl/worksheets/sheet1.xml";
    fn package(sheet: &str, extra: &[(&str, String)]) -> tempfile::TempDir {
        let directory = tempfile::tempdir().unwrap();
        let mut zip = ZipWriter::new(File::create(directory.path().join("metadata.xlsx")).unwrap());
        for (name, bytes) in std::iter::once((SHEET, sheet.to_owned())).chain(extra.iter().cloned())
        {
            zip.start_file(name, SimpleFileOptions::default()).unwrap();
            zip.write_all(bytes.as_bytes()).unwrap();
        }
        zip.finish().unwrap();
        directory
    }
    fn index() -> Value {
        json!({"sheets":[{"path":SHEET}]})
    }
    fn command(kind: &str, range: &str, source: Option<&str>) -> Value {
        let mut value = json!({"type":kind,"sheet":SHEET,"range":range});
        if let Some(source) = source {
            value["source"] = json!(source);
        }
        value
    }
    fn tables(worksheet: &str, relation: &str, table: &str) -> tempfile::TempDir {
        package(
            worksheet,
            &[
                ("xl/worksheets/_rels/sheet1.xml.rels", relation.into()),
                ("xl/tables/table1.xml", table.into()),
            ],
        )
    }
    fn worksheet() -> String {
        format!(
            r#"<worksheet xmlns="{}" xmlns:edge="{}"><sheetData/><tableParts count="1"><tablePart edge:id="t1"/></tableParts></worksheet>"#,
            String::from_utf8_lossy(NS),
            String::from_utf8_lossy(REL)
        )
    }
    fn relation() -> String {
        format!(
            r#"<Relationships xmlns="{}"><Relationship Id="t1" Type="{}" Target="../tables/table1.xml"/></Relationships>"#,
            String::from_utf8_lossy(PACKAGE_REL),
            TABLE_REL
        )
    }
    fn table() -> String {
        format!(
            r#"<table xmlns="{}" ref="D4:E8"><tableColumns count="2"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"/></tableColumns></table>"#,
            String::from_utf8_lossy(NS)
        )
    }
    #[test]
    fn actual_prefixed_merges_guard_source_and_destination_only() {
        let xml = format!(
            r#"<s:worksheet xmlns:s="{}"><s:sheetData/><s:mergeCells count="1"><s:mergeCell ref="B3:C4"/></s:mergeCells></s:worksheet>"#,
            String::from_utf8_lossy(NS)
        );
        let p = package(&xml, &[]);
        let before = std::fs::read(p.path().join("metadata.xlsx")).unwrap();
        assert!(guard(p.path(), &index(), &[command("fillSeries", "A1:A5", None)]).is_ok());
        assert!(
            guard(p.path(), &index(), &[command("fillSeries", "C1:C5", None)])
                .unwrap_err()
                .contains("merged")
        );
        assert!(
            guard(
                p.path(),
                &index(),
                &[command("fillRange", "Z1", Some("B3"))]
            )
            .unwrap_err()
            .contains("source")
        );
        assert!(
            guard(
                p.path(),
                &index(),
                &[command("pasteFormulas", "B3", Some("Z1"))]
            )
            .unwrap_err()
            .contains("destination")
        );
        assert_eq!(
            std::fs::read(p.path().join("metadata.xlsx")).unwrap(),
            before
        );
    }
    #[test]
    fn actual_table_relationships_and_geometry_cannot_be_omitted_by_caller() {
        let p = tables(&worksheet(), &relation(), &table());
        assert!(guard(p.path(), &index(), &[command("fillSeries", "A1:A8", None)]).is_ok());
        for kind in ["fillSeries", "fillRange", "pasteFormulas"] {
            let source = (kind != "fillSeries").then_some("Z1:Z5");
            assert!(
                guard(p.path(), &index(), &[command(kind, "D4:D8", source)])
                    .unwrap_err()
                    .contains("table")
            );
        }
        for kind in ["fillRange", "pasteFormulas"] {
            // Reading from a table changes nothing, and its data rows may be filled.
            assert!(guard(p.path(), &index(), &[command(kind, "Z1:Z5", Some("D4:D8"))]).is_ok());
            assert!(guard(p.path(), &index(), &[command(kind, "D5:D8", Some("Z1:Z4"))]).is_ok());
            // The header row holds the column names and is not overwritten.
            assert!(
                guard(p.path(), &index(), &[command(kind, "D4:D5", Some("Z1:Z2"))])
                    .unwrap_err()
                    .contains("header row")
            );
        }
    }
    #[test]
    fn malformed_or_ambiguous_selected_geometry_refuses() {
        for (sheet, rels, part) in [
            (
                worksheet().replace("count=\"1\"", "count=\"2\""),
                relation(),
                table(),
            ),
            (
                worksheet().replace("</tableParts>", "<tablePart edge:id=\"t1\"/></tableParts>"),
                relation(),
                table(),
            ),
            (
                worksheet(),
                relation().replace("Target=", "TargetMode=\"External\" Target="),
                table(),
            ),
            (
                worksheet(),
                relation().replace("../tables/table1.xml", "../../../escape.xml"),
                table(),
            ),
            (
                worksheet(),
                relation().replace("Id=\"t1\"", "Id=\"missing\""),
                table(),
            ),
            (worksheet(), relation(), table().replace("D4:E8", "D0:E8")),
            (
                worksheet(),
                relation(),
                table().replace(String::from_utf8_lossy(NS).as_ref(), "urn:foreign"),
            ),
        ] {
            let p = tables(&sheet, &rels, &part);
            assert!(guard(p.path(), &index(), &[command("fillSeries", "A1:A8", None)]).is_err());
        }
        for xml in [
            format!(
                r#"<worksheet xmlns="{}"><mergeCells count="2"><mergeCell ref="D3:E4"/></mergeCells></worksheet>"#,
                String::from_utf8_lossy(NS)
            ),
            format!(
                r#"<worksheet xmlns="{}" xmlns:x="urn:foreign"><x:mergeCells><x:mergeCell ref="D3:E4"/></x:mergeCells></worksheet>"#,
                String::from_utf8_lossy(NS)
            ),
            format!(
                r#"<worksheet xmlns="{}"><mergeCells/><mergeCells/></worksheet>"#,
                String::from_utf8_lossy(NS)
            ),
        ] {
            let p = package(&xml, &[]);
            assert!(guard(p.path(), &index(), &[command("fillSeries", "A1:A8", None)]).is_err());
        }
    }
    #[test]
    fn other_bulk_policies_do_not_acquire_new_geometry_restrictions() {
        let directory = tempfile::tempdir().unwrap();
        assert!(
            guard(
                directory.path(),
                &index(),
                &[command("clearRange", "A1", None)]
            )
            .is_ok()
        );
    }
    fn indexed_source(table_source: bool) -> tempfile::TempDir {
        let directory = tempfile::tempdir().unwrap();
        let mut zip = ZipWriter::new(File::create(directory.path().join("source.xlsx")).unwrap());
        let body = r#"<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="Z1"><v>9</v></c></row><row r="2"><c r="A2"><v>3</v></c></row></sheetData>"#;
        let sheet = if table_source {
            worksheet().replace("<sheetData/>", body)
        } else {
            format!(
                r#"<worksheet xmlns="{}">{}<mergeCells count="1"><mergeCell ref="D4:E8"/></mergeCells></worksheet>"#,
                String::from_utf8_lossy(NS),
                body
            )
        };
        let workbook = format!(
            r#"<workbook xmlns="{}" xmlns:r="{}"><sheets><sheet name="Data" sheetId="1" r:id="s1"/></sheets></workbook>"#,
            String::from_utf8_lossy(NS),
            String::from_utf8_lossy(REL)
        );
        let links = format!(
            r#"<Relationships xmlns="{}"><Relationship Id="s1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            String::from_utf8_lossy(PACKAGE_REL)
        );
        let types = format!(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/{SHEET}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#
        );
        let mut parts = vec![
            (SHEET, sheet),
            ("xl/workbook.xml", workbook),
            ("xl/_rels/workbook.xml.rels", links),
            ("[Content_Types].xml", types),
        ];
        if table_source {
            parts.push(("xl/worksheets/_rels/sheet1.xml.rels", relation()));
            parts.push(("xl/tables/table1.xml", table()));
        }
        for (name, text) in parts {
            zip.start_file(name, SimpleFileOptions::default()).unwrap();
            zip.write_all(text.as_bytes()).unwrap();
        }
        zip.finish().unwrap();
        crate::import::index(directory.path()).unwrap();
        directory
    }
    #[test]
    fn values_checks_only_actual_destination_geometry_and_preserves_readonly_source_access() {
        for table_source in [false, true] {
            let directory = indexed_source(table_source);
            let root = directory.path();
            let before = std::fs::read(root.join("cells.sqlite")).unwrap();
            let error=super::super::apply(root,&json!({"version":0,"commands":[command("pasteValues","D4",Some("Z1"))],"view":{"tables":[],"sheets":[{"path":SHEET,"merges":[]}]}})).unwrap_err();
            assert!(
                error.contains(if table_source {
                    "table"
                } else {
                    "merged cells"
                }),
                "{error}"
            );
            assert_eq!(std::fs::read(root.join("cells.sqlite")).unwrap(), before);
            assert!(
                super::super::apply(
                    root,
                    &json!({"version":0,"commands":[command("pasteValues","J10",Some("D4"))]})
                )
                .is_ok()
            );
        }
    }
    #[test]
    fn real_native_bulk_refuses_forged_view_and_late_bad_command_before_any_persistent_write() {
        for table_source in [false, true] {
            let directory = indexed_source(table_source);
            let root = directory.path();
            let source = std::fs::read(root.join("source.xlsx")).unwrap();
            let metadata = std::fs::read(root.join("metadata.xlsx")).unwrap();
            let sql = std::fs::read(root.join("cells.sqlite")).unwrap();
            let input = json!({"version":0,"commands":[command("fillSeries","A1:A5",None),command("fillRange","D4",Some("Z1"))],"view":{"sheets":[{"path":SHEET,"merges":[],"tableRanges":[],"cells":{}}],"tables":[]}});
            let error = super::super::apply(root, &input).unwrap_err();
            assert!(
                error.contains(if table_source { "table" } else { "merged" }),
                "{error}"
            );
            assert_eq!(std::fs::read(root.join("source.xlsx")).unwrap(), source);
            assert_eq!(std::fs::read(root.join("metadata.xlsx")).unwrap(), metadata);
            assert_eq!(std::fs::read(root.join("cells.sqlite")).unwrap(), sql);
            let db = crate::database(root).unwrap();
            assert_eq!(
                db.prepare_cached("SELECT count(*) FROM versions")
                    .unwrap()
                    .query_row([], |r| r.get::<_, i64>(0))
                    .unwrap(),
                1
            );
            assert_eq!(
                db.prepare_cached("SELECT count(*) FROM changes")
                    .unwrap()
                    .query_row([], |r| r.get::<_, i64>(0))
                    .unwrap(),
                0
            );
            let accepted = super::super::apply(
                root,
                &json!({"version":0,"commands":[command("fillSeries","A1:A5",None)]}),
            )
            .unwrap();
            let values = crate::range::read_db(
                &db,
                &json!({"version":accepted["version"],"sheet":SHEET,"range":"A1:A5"}),
            )
            .unwrap();
            assert_eq!(values["cells"][4][1]["value"].as_f64(), Some(9.0));
            assert_eq!(std::fs::read(root.join("source.xlsx")).unwrap(), source);
        }
    }
    #[test]
    fn selected_source_and_destination_cannot_share_one_table_part() {
        let source = "xl/worksheets/sheet2.xml";
        let p = package(
            &worksheet(),
            &[
                (source, worksheet()),
                ("xl/worksheets/_rels/sheet1.xml.rels", relation()),
                ("xl/worksheets/_rels/sheet2.xml.rels", relation()),
                ("xl/tables/table1.xml", table()),
            ],
        );
        let mut command = command("fillRange", "A3", Some("Z1"));
        command["sourceSheet"] = json!(source);
        let index = json!({"sheets":[{"path":SHEET},{"path":source}]});
        assert!(
            guard(p.path(), &index, &[command])
                .unwrap_err()
                .contains("shared worksheet")
        );
    }
    #[test]
    fn unrelated_external_hyperlinks_are_not_geometry_barriers() {
        let relation=relation().replace("</Relationships>",r#"<Relationship Id="link" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/" TargetMode="External"/></Relationships>"#);
        let p = tables(&worksheet(), &relation, &table());
        assert!(guard(p.path(), &index(), &[command("fillSeries", "A1:A8", None)]).is_ok());
        assert!(
            guard(
                p.path(),
                &index(),
                &[command("fillRange", "A3", Some("Z1"))]
            )
            .is_ok()
        );
    }
    #[test]
    fn array_source_follower_cache_is_not_a_literal_or_silently_skipped_formula() {
        let directory = indexed_source(false);
        let root = directory.path();
        std::fs::remove_file(root.join("cells.sqlite")).unwrap();
        let mut source = ZipArchive::new(File::open(root.join("source.xlsx")).unwrap()).unwrap();
        let mut output = ZipWriter::new(File::create(root.join("replacement.xlsx")).unwrap());
        for index in 0..source.len() {
            let member = source.by_index(index).unwrap();
            if member.name() == SHEET {
                output
                    .start_file(SHEET, SimpleFileOptions::default())
                    .unwrap();
                let xml = format!(
                    r#"<worksheet xmlns="{}"><sheetData><row r="1"><c r="C1"><f t="array" ref="C1:C2">UNKNOWN()</f><v>913</v></c></row><row r="2"><c r="B2"><f>A3+1</f><v>5</v></c><c r="C2"><v>914</v></c></row><row r="3"><c r="A3"><v>4</v></c></row></sheetData></worksheet>"#,
                    String::from_utf8_lossy(NS)
                );
                output.write_all(xml.as_bytes()).unwrap();
            } else {
                output.raw_copy_file(member).unwrap();
            }
        }
        output.finish().unwrap();
        drop(source);
        std::fs::rename(root.join("replacement.xlsx"), root.join("source.xlsx")).unwrap();
        crate::import::index(root).unwrap();
        let mut view = crate::package_edit::auditing_metadata(root).unwrap();
        view["readOnly"] = json!(false);
        view["sheets"][0]["storedFormulaCount"] = json!(2);
        crate::calculation_policy::command(root,&json!({"version":0,"method":"initialize","view":view,"initialPolicy":{"mode":"auto","recalculateBeforeSave":false}})).unwrap();
        let db = crate::database(root).unwrap();
        let raw =
            crate::range::read_db(&db, &json!({"version":0,"sheet":SHEET,"range":"C2"})).unwrap();
        let mut follower = raw["cells"][0][1].clone();
        crate::calculation_policy::mask(&db, 0, SHEET, "C2", &mut follower).unwrap();
        assert_eq!(follower["value"].as_f64(), Some(914.0));
        assert_eq!(
            follower["formulaOutput"],
            json!({"schema":1,"master":"C1","range":"C1:C2"})
        );
        assert!(follower.get("formulaType").is_none());
        assert!(follower["calculation"]["pending"] != true);
        let before = std::fs::read(root.join("cells.sqlite")).unwrap();
        let source = std::fs::read(root.join("source.xlsx")).unwrap();
        let metadata = std::fs::read(root.join("metadata.xlsx")).unwrap();
        for (kind, range, source_range) in [
            ("fillRange", "G2", "C2"),
            ("pasteFormulas", "G2", "C2"),
            ("fillRange", "G2:H2", "B2:C2"),
            ("pasteFormulas", "G2:H2", "B2:C2"),
        ] {
            let mut forged = view.clone();
            forged["sheets"][0]["cells"] = json!({"C2":{"value":914,"formulaOutput":null}});
            let error=super::super::apply(root,&json!({"version":0,"commands":[command(kind,range,Some(source_range))],"view":forged})).unwrap_err();
            assert!(error.contains("Source formula output"), "{error}");
            assert_eq!(std::fs::read(root.join("cells.sqlite")).unwrap(), before);
            assert_eq!(std::fs::read(root.join("source.xlsx")).unwrap(), source);
            assert_eq!(std::fs::read(root.join("metadata.xlsx")).unwrap(), metadata);
        }
        assert_eq!(
            db.prepare_cached("SELECT count(*) FROM versions")
                .unwrap()
                .query_row([], |r| r.get::<_, i64>(0))
                .unwrap(),
            1
        );
        assert_eq!(
            db.prepare_cached("SELECT count(*) FROM changes")
                .unwrap()
                .query_row([], |r| r.get::<_, i64>(0))
                .unwrap(),
            0
        );
        let accepted = super::super::apply(
            root,
            &json!({"version":0,"commands":[command("fillRange","G2",Some("B2"))]}),
        )
        .unwrap();
        let read = crate::range::read_db(
            &db,
            &json!({"version":accepted["version"],"sheet":SHEET,"range":"G2"}),
        )
        .unwrap();
        assert_eq!(read["cells"][0][1]["formula"], "F3+1");
    }
}
