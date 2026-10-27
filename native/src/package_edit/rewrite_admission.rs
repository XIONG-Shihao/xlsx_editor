//! Independent admission of the geometry described by compact JavaScript plans.
//! Cell movement is performed by the native writers; these checks keep metadata
//! fragments from smuggling a second worksheet or coordinate transformation.
use super::print_settings::Tree;
use super::*;
use crate::used_range_transfer::{NativeMapping, worksheet_inventory};

fn clean(node: &print_settings::Node) -> BTreeMap<String, String> {
    node.attrs
        .iter()
        .filter(|(key, _)| *key != "xmlns" && !key.starts_with("xmlns:"))
        .map(|(key, value)| (key.clone(), value.clone()))
        .collect()
}
fn model(tree: &Tree, index: usize) -> Value {
    let node = &tree.nodes[index];
    json!({"ns":String::from_utf8_lossy(&node.ns),"tag":node.name,"attrs":clean(node),"text":if node.text.trim().is_empty(){""}else{&node.text},"children":node.children.iter().map(|index|model(tree,*index)).collect::<Vec<_>>()})
}
/// The records directly under a worksheet's root that carry one of the wanted
/// names, read without normalising anything: element names keep their
/// prefixes and attributes keep their spelling, so a record can be changed and
/// written back in place and still mean the same thing.
fn records(bytes: &[u8], wanted: &[&str]) -> Result<BTreeMap<String, Vec<Value>>> {
    let mut reader = NsReader::from_reader(bytes);
    let mut buffer = Vec::new();
    let mut found: BTreeMap<String, Vec<Value>> = BTreeMap::new();
    // The elements being read, outermost first. Empty while outside a wanted record.
    let mut open: Vec<Value> = Vec::new();
    let (mut depth, mut nodes) = (0usize, 0usize);
    loop {
        crate::cancel::check()?;
        let (binding, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let empty = matches!(event, Event::Empty(_));
                let namespace = match binding {
                    quick_xml::name::ResolveResult::Bound(ns) => ns.as_ref().to_vec(),
                    _ => Vec::new(),
                };
                let local = String::from_utf8_lossy(element.local_name().as_ref()).into_owned();
                let wanted_here = depth == 1 && namespace == NS && wanted.contains(&local.as_str());
                if !open.is_empty() || wanted_here {
                    nodes += 1;
                    if nodes > 200000 || open.len() > 64 {
                        return Err(crate::coded(
                            "package-limit",
                            "Worksheet records exceed their bound",
                        ));
                    }
                    let mut attributes = serde_json::Map::new();
                    for attribute in element.attributes().with_checks(true) {
                        let attribute = attribute.map_err(err)?;
                        attributes.insert(
                            String::from_utf8_lossy(attribute.key.as_ref()).into_owned(),
                            json!(attribute.unescape_value().map_err(err)?),
                        );
                    }
                    let node = json!({
                        "q": String::from_utf8_lossy(element.name().as_ref()),
                        "ns": String::from_utf8_lossy(&namespace),
                        "tag": local,
                        "attrs": attributes,
                        "text": "",
                        "children": []
                    });
                    if empty {
                        match open.last_mut() {
                            Some(parent) => parent["children"].as_array_mut().unwrap().push(node),
                            None => found.entry(local).or_default().push(node),
                        }
                    } else {
                        open.push(node);
                    }
                }
                if !empty {
                    depth += 1;
                }
            }
            Event::Text(text) => {
                if let Some(node) = open.last_mut() {
                    let text = text.unescape().map_err(err)?;
                    node["text"] = json!(format!("{}{}", node["text"].as_str().unwrap(), text));
                }
            }
            Event::CData(text) => {
                if let Some(node) = open.last_mut() {
                    let text = String::from_utf8_lossy(&text).into_owned();
                    node["text"] = json!(format!("{}{}", node["text"].as_str().unwrap(), text));
                }
            }
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| crate::coded("package-malformed", "Unbalanced worksheet"))?;
                if let Some(mut node) = open.pop() {
                    if node["text"].as_str().unwrap().trim().is_empty() {
                        node["text"] = json!("");
                    }
                    match open.last_mut() {
                        Some(parent) => parent["children"].as_array_mut().unwrap().push(node),
                        None => found
                            .entry(node["tag"].as_str().unwrap().to_owned())
                            .or_default()
                            .push(node),
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(found)
}
fn escape(text: &str, attribute: bool, out: &mut String) {
    for character in text.chars() {
        match character {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' if attribute => out.push_str("&quot;"),
            '\n' if attribute => out.push_str("&#10;"),
            '\r' => out.push_str("&#13;"),
            '\t' if attribute => out.push_str("&#9;"),
            other => out.push(other),
        }
    }
}
fn to_xml(node: &Value, out: &mut String) -> Result<()> {
    let name = node["q"]
        .as_str()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid record name"))?;
    out.push('<');
    out.push_str(name);
    for (key, value) in node["attrs"]
        .as_object()
        .ok_or_else(|| crate::coded("package-malformed", "Invalid record attributes"))?
    {
        out.push(' ');
        out.push_str(key);
        out.push_str("=\"");
        escape(value.as_str().unwrap_or(""), true, out);
        out.push('"');
    }
    let text = node["text"].as_str().unwrap_or("");
    let children = node["children"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid record children"))?;
    if text.is_empty() && children.is_empty() {
        out.push_str("/>");
        return Ok(());
    }
    out.push('>');
    escape(text, false, out);
    for child in children {
        to_xml(child, out)?;
    }
    out.push_str("</");
    out.push_str(name);
    out.push('>');
    Ok(())
}
fn children(tree: &Tree, tag: &str) -> Result<Vec<Value>> {
    Ok(tree.nodes[0]
        .children
        .iter()
        .filter(|index| tree.nodes[**index].ns == NS && tree.nodes[**index].name == tag)
        .map(|index| model(tree, *index))
        .collect())
}
fn fragment(bytes: &[u8], tag: &str) -> Result<Vec<Value>> {
    let mut wrapped = format!(
        "<worksheet xmlns=\"{}\" xmlns:r=\"{}\">",
        String::from_utf8_lossy(NS),
        String::from_utf8_lossy(REL_NS)
    )
    .into_bytes();
    wrapped.extend_from_slice(bytes);
    wrapped.extend_from_slice(b"</worksheet>");
    let tree = Tree::parse(&wrapped)?;
    tree.root(NS, "worksheet")?;
    if !tree.nodes[0].text.trim().is_empty()
        || tree.nodes[0]
            .children
            .iter()
            .any(|index| tree.nodes[*index].ns != NS || tree.nodes[*index].name != tag)
    {
        return Err(crate::coded(
            "package-unsupported",
            "Rewrite fragment contains unrelated worksheet geometry",
        ));
    }
    children(&tree, tag)
}
fn metadata_sheet(root: &Path, sheet: &str) -> Result<Tree> {
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let tree = Tree::parse(&part_bytes(&mut zip, sheet, &serde_json::Map::new())?)?;
    tree.root(NS, "worksheet")?;
    Ok(tree)
}
fn integer(value: Option<&Value>, limit: u32) -> Result<u32> {
    value
        .and_then(Value::as_u64)
        .filter(|value| *value < u64::from(limit))
        .map(|value| value as u32)
        .ok_or_else(|| {
            crate::coded(
                "internal-invalid-field",
                "Invalid native rewrite coordinate",
            )
        })
}
fn attr_number(node: &Value, key: &str) -> Result<u32> {
    node["attrs"][key]
        .as_str()
        .ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing rewrite geometry attribute",
            )
        })?
        .parse()
        .map_err(err)
}
fn empty_node(tag: &str, attrs: Value) -> Value {
    json!({"ns":String::from_utf8_lossy(NS),"tag":tag,"attrs":attrs,"text":"","children":[]})
}
fn range(reference: &str) -> Result<((u32, u32), (u32, u32))> {
    let bounds = crate::bulk::bounds(reference)?;
    let canonical = if bounds.0 == bounds.1 {
        crate::xml::address(bounds.0.0, bounds.0.1)
    } else {
        format!(
            "{}:{}",
            crate::xml::address(bounds.0.0, bounds.0.1),
            crate::xml::address(bounds.1.0, bounds.1.1)
        )
    };
    if canonical != reference {
        return Err(crate::coded(
            "internal-invalid-field",
            "Noncanonical rewrite range",
        ));
    }
    Ok(bounds)
}

pub(super) fn layout(root: &Path, input: &Value, bytes: Option<&[u8]>) -> Result<()> {
    let command = &input["command"];
    let operation = field(command, "operation")?;
    let inventory = worksheet_inventory(&local(root, field(input, "source")?)?)?;
    let target = inventory
        .iter()
        .find(|sheet| sheet.path == input["sheet"])
        .ok_or_else(|| {
            crate::coded("internal-unknown-reference", "Unknown rewrite layout sheet")
        })?;
    if command["sheet"] != target.name {
        return Err(crate::coded(
            "internal-mismatch",
            "Layout command differs from native worksheet identity",
        ));
    }
    let allowed = match operation {
        "unhideAll" => vec!["operation", "sheet"],
        "resizeRow" => vec!["operation", "sheet", "index", "height"],
        "resizeColumn" => vec!["operation", "sheet", "index", "width"],
        "freeze" => vec!["operation", "sheet", "rows", "columns"],
        "mergeCells" | "unmergeCells" => vec!["operation", "sheet", "range"],
        _ => {
            return Err(crate::coded(
                "internal-unknown-reference",
                "Unknown rewrite layout command",
            ));
        }
    };
    if command
        .as_object()
        .is_none_or(|object| object.keys().any(|key| !allowed.contains(&key.as_str())))
    {
        return Err(crate::coded(
            "package-unsupported",
            "Unqualified rewrite layout command fields",
        ));
    }
    if operation == "unhideAll" {
        if input.get("tag").is_some()
            || input.get("row").is_some()
            || input.get("height").is_some()
            || bytes.is_some()
        {
            return Err(crate::coded(
                "package-bad-input",
                "Unhide admits no caller layout fragment",
            ));
        }
        let source = metadata_sheet(root, &target.path)?;
        if source.nodes[0].children.iter().any(|i| {
            source.nodes[*i].ns == NS
                && ["sheetProtection", "autoFilter", "tableParts"]
                    .contains(&source.nodes[*i].name.as_str())
        }) {
            return Err(crate::coded(
                "package-protected",
                "Unhide requires an unprotected worksheet without linked filters/tables",
            ));
        }
        return Ok(());
    }
    if operation == "resizeRow" {
        if input.get("tag").is_some()
            || input["row"] != command["index"]
            || input["height"] != command["height"]
        {
            return Err(crate::coded(
                "internal-mismatch",
                "Row layout differs from admitted command",
            ));
        }
        integer(command.get("index"), 1048576)?;
        if command["height"]
            .as_f64()
            .is_none_or(|height| !height.is_finite() || !(8.0..=200.0).contains(&height))
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid rewrite row height",
            ));
        }
        return Ok(());
    }
    let tag = field(input, "tag")?;
    if tag
        != match operation {
            "resizeColumn" => "cols",
            "freeze" => "sheetViews",
            _ => "mergeCells",
        }
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Layout tag differs from admitted command",
        ));
    }
    let source = metadata_sheet(root, &target.path)?;
    let mut expected = children(&source, tag)?;
    let actual = fragment(
        bytes.ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing admitted layout fragment")
        })?,
        tag,
    )?;
    if operation == "resizeColumn" {
        let coordinate = integer(command.get("index"), 16384)? + 1;
        let width = command["width"]
            .as_f64()
            .filter(|width| width.is_finite() && (0.0..=255.0).contains(width))
            .ok_or_else(|| {
                crate::coded("internal-invalid-field", "Invalid rewrite column width")
            })?;
        if expected.len() > 1 {
            return Err(crate::coded(
                "package-ambiguous",
                "Ambiguous source column geometry",
            ));
        }
        let mut columns = expected
            .first()
            .map(|node| {
                node["children"]
                    .as_array()
                    .cloned()
                    .ok_or_else(|| crate::coded("internal-missing-field", "Missing columns"))
            })
            .transpose()?
            .unwrap_or_default();
        let mut replacement = Vec::new();
        let mut modified = false;
        let mut previous = 0;
        for mut column in columns.drain(..) {
            if column["ns"] != String::from_utf8_lossy(NS).as_ref()
                || column["tag"] != "col"
                || column["children"]
                    .as_array()
                    .is_none_or(|items| !items.is_empty())
            {
                return Err(crate::coded(
                    "package-unsupported",
                    "Opaque source column geometry",
                ));
            }
            let first = attr_number(&column, "min")?;
            let last = attr_number(&column, "max")?;
            if first == 0 || first <= previous || last < first || last > 16384 {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid source column intervals",
                ));
            }
            previous = last;
            if coordinate < first || coordinate > last {
                replacement.push(column);
                continue;
            }
            modified = true;
            if first < coordinate {
                let mut left = column.clone();
                left["attrs"]["max"] = json!((coordinate - 1).to_string());
                replacement.push(left);
            }
            let right = if last > coordinate {
                let mut right = column.clone();
                right["attrs"]["min"] = json!((coordinate + 1).to_string());
                Some(right)
            } else {
                None
            };
            column["attrs"]["min"] = json!(coordinate.to_string());
            column["attrs"]["max"] = json!(coordinate.to_string());
            column["attrs"]["width"] = json!(width.to_string());
            column["attrs"]["customWidth"] = json!("1");
            replacement.push(column);
            if let Some(right) = right {
                replacement.push(right);
            }
        }
        if !modified {
            let column = empty_node(
                "col",
                json!({"min":coordinate.to_string(),"max":coordinate.to_string(),"width":width.to_string(),"customWidth":"1"}),
            );
            let at = replacement
                .iter()
                .position(|node| attr_number(node, "min").is_ok_and(|first| first > coordinate))
                .unwrap_or(replacement.len());
            replacement.insert(at, column);
        }
        if expected.is_empty() {
            expected.push(empty_node("cols", json!({})));
        }
        expected[0]["children"] = json!(replacement);
    } else if operation == "freeze" {
        let rows = integer(command.get("rows"), 1048576)?;
        let columns = integer(command.get("columns"), 16384)?;
        if expected.len() > 1 {
            return Err(crate::coded(
                "package-ambiguous",
                "Ambiguous source sheetViews",
            ));
        }
        if expected.is_empty() {
            expected.push(empty_node("sheetViews", json!({})));
        }
        let views = expected[0]["children"]
            .as_array_mut()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheet views"))?;
        if views.is_empty() {
            views.push(empty_node("sheetView", json!({"workbookViewId":"0"})));
        }
        let view = views
            .first_mut()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing primary sheet view"))?;
        if view["tag"] != "sheetView" || view["ns"] != String::from_utf8_lossy(NS).as_ref() {
            return Err(crate::coded(
                "package-unsupported",
                "Opaque primary sheet view",
            ));
        }
        let nodes = view["children"]
            .as_array_mut()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheet view children"))?;
        nodes.retain(|node| node["tag"] != "pane");
        for selection in nodes.iter_mut().filter(|node| node["tag"] == "selection") {
            selection["attrs"]
                .as_object_mut()
                .ok_or_else(|| crate::coded("package-malformed", "Invalid selection attributes"))?
                .remove("pane");
        }
        if rows > 0 || columns > 0 {
            nodes.insert(0,empty_node("pane",json!({"xSplit":columns.to_string(),"ySplit":rows.to_string(),"topLeftCell":crate::xml::address(rows,columns),"state":"frozen","activePane":if rows>0&&columns>0{"bottomRight"}else if rows>0{"bottomLeft"}else{"topRight"}})));
        }
    } else {
        if input["kind"] != operation || input["range"] != command["range"] {
            return Err(crate::coded(
                "internal-mismatch",
                "Merge command differs from native geometry",
            ));
        }
        let reference = field(command, "range")?;
        let selected = range(reference)?;
        if selected.0 == selected.1 {
            return Err(crate::coded(
                "package-bad-input",
                "Merge requires multiple positions",
            ));
        }
        if expected.len() > 1 {
            return Err(crate::coded(
                "package-ambiguous",
                "Ambiguous source merge geometry",
            ));
        }
        let mut entries = expected
            .first()
            .map(|node| {
                node["children"]
                    .as_array()
                    .cloned()
                    .ok_or_else(|| crate::coded("internal-missing-field", "Missing merge entries"))
            })
            .transpose()?
            .unwrap_or_default();
        for entry in &entries {
            if entry["tag"] != "mergeCell"
                || entry["ns"] != String::from_utf8_lossy(NS).as_ref()
                || entry["attrs"]
                    .as_object()
                    .is_none_or(|attrs| attrs.len() != 1)
                || entry["children"]
                    .as_array()
                    .is_none_or(|nodes| !nodes.is_empty())
            {
                return Err(crate::coded(
                    "package-unsupported",
                    "Opaque source merge entry",
                ));
            }
            range(field(&entry["attrs"], "ref")?)?;
        }
        if operation == "mergeCells" {
            for entry in &entries {
                let other = range(field(&entry["attrs"], "ref")?)?;
                if selected.0.0 <= other.1.0
                    && selected.1.0 >= other.0.0
                    && selected.0.1 <= other.1.1
                    && selected.1.1 >= other.0.1
                {
                    return Err(crate::coded(
                        "package-conflict",
                        "Rewrite merge overlaps existing geometry",
                    ));
                }
            }
            entries.push(empty_node("mergeCell", json!({"ref":reference})));
        } else {
            let at = entries
                .iter()
                .position(|entry| entry["attrs"]["ref"] == reference)
                .ok_or_else(|| {
                    crate::coded("internal-missing-field", "Exact source merge is missing")
                })?;
            entries.remove(at);
        }
        if entries.is_empty() {
            expected.clear();
        } else {
            if expected.is_empty() {
                expected.push(empty_node("mergeCells", json!({})));
            }
            expected[0]["attrs"]["count"] = json!(entries.len().to_string());
            expected[0]["children"] = json!(entries);
        }
    }
    if expected != actual {
        return Err(crate::coded(
            "internal-mismatch",
            "Layout fragment differs from independently derived command geometry",
        ));
    }
    Ok(())
}

pub(super) fn package_mapping(
    root: &Path,
    input: &Value,
    plan: &Value,
    output: &str,
) -> Result<NativeMapping> {
    let before = worksheet_inventory(&local(root, field(input, "source")?)?)?;
    let after = worksheet_inventory(&local(root, output)?)?;
    let kind = field(plan, "kind")?;
    let add = ["addSheet", "duplicateSheet", "createPivot"].contains(&kind);
    if kind == "createTable" {
        table_creation(root, input, plan)?;
    }
    if kind == "addConditionalRule" {
        conditional_addition(root, plan)?;
    }
    if after.len() != before.len() + usize::from(add) {
        return Err(crate::coded(
            "package-stale",
            "Package rewrite changed undeclared worksheets",
        ));
    }
    let rename = kind == "renameSheet";
    let rename_old = if rename || kind == "duplicateSheet" {
        Some(field(&plan["rename"], "old")?)
    } else {
        None
    };
    let rename_new = if rename || kind == "duplicateSheet" {
        Some(field(&plan["rename"], "new")?)
    } else {
        None
    };
    let mut renamed = 0;
    let mut predecessors = BTreeMap::new();
    for (old, new) in before.iter().zip(&after) {
        if old.path != new.path {
            return Err(crate::coded(
                "internal-mismatch",
                "Package rewrite moved or reordered an existing worksheet part",
            ));
        }
        let expected = if rename && Some(old.name.as_str()) == rename_old {
            renamed += 1;
            rename_new.unwrap()
        } else {
            &old.name
        };
        if new.name != expected {
            return Err(crate::coded(
                "package-stale",
                "Package rewrite changed an undeclared worksheet name",
            ));
        }
        predecessors.insert(new.path.clone(), Some(old.path.clone()));
    }
    if rename && renamed != 1 {
        return Err(crate::coded(
            "package-bad-input",
            "Rename requires one exact native predecessor",
        ));
    }
    let originals = before
        .iter()
        .map(|sheet| sheet.path.as_str())
        .collect::<BTreeSet<_>>();
    let replacements = plan["replacements"].as_object().ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing package rewrite replacements",
        )
    })?;
    let copies = plan["copies"]
        .as_object()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing package rewrite copies"))?;
    let removed = plan["removed"].as_array().ok_or_else(|| {
        crate::coded("internal-missing-field", "Missing package rewrite removals")
    })?;
    if removed
        .iter()
        .any(|part| part.as_str().is_none_or(|part| originals.contains(part)))
        || copies.keys().any(|part| originals.contains(part.as_str()))
    {
        return Err(crate::coded(
            "package-unsupported",
            "Package rewrite cannot replace an existing worksheet identity",
        ));
    }
    // Only the separately qualified pivot writer replaces existing cell XML.
    if replacements.keys().any(|part| {
        originals.contains(part.as_str())
            && !(["refreshPivot", "updatePivot", "setPivotLayout"].contains(&kind)
                && plan["targetPath"] == *part)
    }) {
        return Err(crate::coded(
            "package-unsupported",
            "Package rewrite contains an unqualified worksheet replacement",
        ));
    }
    if add {
        let new = after
            .last()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing new worksheet"))?;
        if kind == "duplicateSheet" {
            let source = before
                .iter()
                .find(|sheet| Some(sheet.name.as_str()) == rename_old)
                .ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing copied native worksheet")
                })?;
            if Some(new.name.as_str()) != rename_new
                || plan["sheet"] != source.path
                || plan["copyOnly"] != true
                || copies.get(&new.path).and_then(Value::as_str) != Some(source.path.as_str())
                || plan["fragments"]
                    .as_object()
                    .is_none_or(|fragments| !fragments.is_empty())
            {
                return Err(crate::coded(
                    "internal-mismatch",
                    "Copied worksheet differs from its admitted native copy command",
                ));
            }
            predecessors.insert(new.path.clone(), Some(source.path.clone()));
        } else if kind == "createPivot" {
            if plan["targetPath"] != new.path || plan["targetSheet"] != new.name {
                return Err(crate::coded(
                    "package-unsupported",
                    "Pivot creation added a different worksheet",
                ));
            }
            predecessors.insert(new.path.clone(), None);
        } else {
            if !copies.is_empty()
                || !removed.is_empty()
                || plan["fragments"]
                    .as_object()
                    .is_none_or(|fragments| !fragments.is_empty())
            {
                return Err(crate::coded(
                    "package-unsupported",
                    "New worksheet plan contains unrelated movement",
                ));
            }
            let mut zip = ZipArchive::new(crate::cancel::Reader(
                File::open(local(root, output)?).map_err(err)?,
            ))
            .map_err(err)?;
            let tree = Tree::parse(&part_bytes(&mut zip, &new.path, &serde_json::Map::new())?)?;
            tree.root(NS, "worksheet")?;
            if tree.nodes[0].children.len() != 1
                || tree.nodes[0].children.iter().any(|index| {
                    tree.nodes[*index].ns != NS
                        || tree.nodes[*index].name != "sheetData"
                        || !tree.nodes[*index].children.is_empty()
                        || !tree.nodes[*index].text.trim().is_empty()
                })
            {
                return Err(crate::coded(
                    "package-bad-input",
                    "New worksheet must have empty native geometry",
                ));
            }
            predecessors.insert(new.path.clone(), None);
        }
    } else if !copies.is_empty() {
        return Err(crate::coded(
            "package-refused",
            "Identity package rewrite contains undeclared copies",
        ));
    }
    Ok(NativeMapping::Sheets { predecessors })
}

fn one_fragment<'a>(plan: &'a Value, sheet: &str, tag: &str) -> Result<&'a str> {
    let sheets = plan["fragments"].as_object().ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing metadata rewrite fragments",
        )
    })?;
    if sheets.len() != 1 {
        return Err(crate::coded(
            "package-stale",
            "Metadata rewrite changed an unrelated worksheet",
        ));
    }
    let fragments = sheets
        .get(sheet)
        .and_then(Value::as_object)
        .ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing owned worksheet fragment")
        })?;
    if fragments.len() != 1 {
        return Err(crate::coded(
            "package-stale",
            "Metadata rewrite changed an unrelated worksheet fragment",
        ));
    }
    fragments.get(tag).and_then(Value::as_str).ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing qualified metadata fragment",
        )
    })
}
fn unchanged_except_added(before: &Tree, after: &Tree, added: &Value) -> Result<()> {
    let old = model(before, 0);
    let mut new = model(after, 0);
    let children = new["children"]
        .as_array_mut()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing metadata graph children"))?;
    let indexes = children
        .iter()
        .enumerate()
        .filter(|(_, node)| *node == added)
        .map(|(index, _)| index)
        .collect::<Vec<_>>();
    if indexes.len() != 1 {
        return Err(crate::coded(
            "internal-missing-field",
            "Metadata addition is missing or ambiguous",
        ));
    }
    children.remove(indexes[0]);
    if old != new {
        return Err(crate::coded(
            "package-stale",
            "Metadata rewrite changed an unrelated package graph",
        ));
    }
    Ok(())
}
fn table_creation(root: &Path, input: &Value, plan: &Value) -> Result<()> {
    let sheet = field(plan, "sheet")?;
    let selected = range(field(plan, "range")?)?;
    let positions =
        u64::from(selected.1.0 - selected.0.0 + 1) * u64::from(selected.1.1 - selected.0.1 + 1);
    if selected.0.0 == selected.1.0 || positions > 10000 {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid native table creation rectangle",
        ));
    }
    let replacements = plan["replacements"].as_object().ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing table creation replacements",
        )
    })?;
    if !plan["copies"]
        .as_object()
        .is_some_and(|copies| copies.is_empty())
        || !plan["removed"]
            .as_array()
            .is_some_and(|removed| removed.is_empty())
    {
        return Err(crate::coded(
            "package-unsupported",
            "Table creation contains unrelated package movement",
        ));
    }
    let xml = one_fragment(plan, sheet, "tableParts")?;
    let source = metadata_sheet(root, sheet)?;
    let prior = children(&source, "tableParts")?;
    let actual = fragment(xml.as_bytes(), "tableParts")?;
    if prior.len() > 1 || actual.len() != 1 {
        return Err(crate::coded(
            "package-ambiguous",
            "Ambiguous table creation owner fragment",
        ));
    }
    let old = prior
        .first()
        .and_then(|node| node["children"].as_array())
        .cloned()
        .unwrap_or_default();
    let new = actual[0]["children"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing new table edges"))?;
    if new.len() != old.len() + 1
        || new[..old.len()] != old
        || actual[0]["attrs"] != json!({"count":new.len().to_string()})
    {
        return Err(crate::coded(
            "package-stale",
            "Table creation changed unrelated owner edges",
        ));
    }
    let edge = new
        .last()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing new table edge"))?;
    if edge["tag"] != "tablePart"
        || edge["ns"] != String::from_utf8_lossy(NS).as_ref()
        || edge["attrs"]
            .as_object()
            .is_none_or(|attrs| attrs.len() != 1)
        || edge["children"] != json!([])
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid new table owner edge",
        ));
    }
    let id = field(&edge["attrs"], "r:id")?;
    let (directory, file) = sheet
        .rsplit_once('/')
        .ok_or_else(|| crate::coded("package-malformed", "Invalid table worksheet path"))?;
    let rel_part = format!("{directory}/_rels/{file}.rels");
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let relation = Tree::parse(
        replacements
            .get(&rel_part)
            .and_then(Value::as_str)
            .ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing new table relationship")
            })?
            .as_bytes(),
    )?;
    relation.root(PACKAGE_REL_NS, "Relationships")?;
    let links = relation.nodes[0]
        .children
        .iter()
        .filter(|index| relation.nodes[**index].attrs.get("Id").map(String::as_str) == Some(id))
        .copied()
        .collect::<Vec<_>>();
    if links.len() != 1 {
        return Err(crate::coded(
            "package-refused",
            "New table edge is ambiguous",
        ));
    }
    let link = &relation.nodes[links[0]];
    if link.name != "Relationship"
        || link.ns != PACKAGE_REL_NS
        || clean(link).len() != 3
        || !link.children.is_empty()
        || link.attrs.get("Type").map(String::as_str)
            != Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships/table")
    {
        return Err(crate::coded(
            "package-malformed",
            "Invalid new table relationship kind",
        ));
    }
    let table_part = resolve(
        sheet,
        link.attrs.get("Target").ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing new table relationship target",
            )
        })?,
    )?;
    if zip.file_names().any(|part| part == table_part)
        || replacements.len() != 3
        || !replacements.contains_key("[Content_Types].xml")
    {
        return Err(crate::coded(
            "package-bad-input",
            "Table creation destination exists or replacement set is not closed",
        ));
    }
    let table = Tree::parse(
        replacements
            .get(&table_part)
            .and_then(Value::as_str)
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing new table definition"))?
            .as_bytes(),
    )?;
    table.root(NS, "table")?;
    let name = table.nodes[0]
        .attrs
        .get("name")
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing new table name"))?;
    if !name_addition::legal_name(name)
        || table.nodes[0].attrs.get("displayName") != Some(name)
        || table.nodes[0].attrs.get("ref").map(String::as_str) != Some(field(plan, "range")?)
        || table.nodes[0]
            .attrs
            .get("headerRowCount")
            .map(String::as_str)
            != Some("1")
        || table.nodes[0]
            .attrs
            .get("totalsRowCount")
            .map(String::as_str)
            != Some("0")
    {
        return Err(crate::coded(
            "internal-mismatch",
            "New table definition differs from native creation geometry",
        ));
    }
    let inventory = super::auditing_metadata(root)?;
    let target = inventory["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing native table sheets"))?
        .iter()
        .find(|node| node["path"] == sheet)
        .ok_or_else(|| crate::coded("internal-unknown-reference", "Unknown native table sheet"))?;
    for old in inventory["tables"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing native table inventory"))?
    {
        if field(old, "name")?.eq_ignore_ascii_case(name) {
            return Err(crate::coded(
                "package-conflict",
                "Native table name collision",
            ));
        }
        if old["sheet"] == target["name"] {
            let bounds = range(field(old, "ref")?)?;
            if selected.0.0 <= bounds.1.0
                && selected.1.0 >= bounds.0.0
                && selected.0.1 <= bounds.1.1
                && selected.1.1 >= bounds.0.1
            {
                return Err(crate::coded(
                    "package-conflict",
                    "Native table creation overlaps an existing table",
                ));
            }
        }
    }
    for node in source
        .nodes
        .iter()
        .filter(|node| node.ns == NS && node.name == "mergeCell")
    {
        let bounds = range(node.attrs.get("ref").ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing source merge range")
        })?)?;
        if selected.0.0 <= bounds.1.0
            && selected.1.0 >= bounds.0.0
            && selected.0.1 <= bounds.1.1
            && selected.1.1 >= bounds.0.1
        {
            return Err(crate::coded(
                "package-conflict",
                "Native table creation overlaps a source merge",
            ));
        }
    }
    let columns = table
        .direct(0, "tableColumns")?
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing new table columns"))?;
    let count = (selected.1.1 - selected.0.1 + 1) as usize;
    if table.nodes[columns].children.len() != count
        || table.nodes[columns].attrs.get("count").map(String::as_str)
            != Some(count.to_string().as_str())
    {
        return Err(crate::coded(
            "internal-mismatch",
            "New table column inventory differs from native width",
        ));
    }
    let db = crate::database(root)?;
    let headers = crate::range::read_public_db(
        &db,
        &json!({"version":input["version"],"sheet":sheet,"range":format!("{}:{}",crate::xml::address(selected.0.0,selected.0.1),crate::xml::address(selected.0.0,selected.1.1)),"limit":count+1}),
    )?;
    let values = headers["cells"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing native table headers"))?;
    if values.len() != count || !headers["nextCursor"].is_null() {
        return Err(crate::coded(
            "package-unsupported",
            "Incomplete native table header inventory",
        ));
    }
    let mut labels = BTreeSet::new();
    for (index, (entry, value)) in table.nodes[columns].children.iter().zip(values).enumerate() {
        let node = &table.nodes[*entry];
        let header = value[1]["value"].as_str().ok_or_else(|| {
            crate::coded(
                "package-bad-input",
                "Table header is not accepted literal text",
            )
        })?;
        if value[1].get("formula").is_some()
            || value[1].get("formulaType").is_some()
            || header.trim().is_empty()
            || !labels.insert(header.to_lowercase())
            || node.name != "tableColumn"
            || node.ns != NS
            || !node.children.is_empty()
            || node.attrs.get("id").map(String::as_str) != Some((index + 1).to_string().as_str())
            || node.attrs.get("name").map(String::as_str) != Some(header)
        {
            return Err(crate::coded(
                "internal-mismatch",
                "New table headers differ from accepted native source cells",
            ));
        }
    }
    if zip.file_names().any(|part| part == rel_part) {
        let old = Tree::parse(&part_bytes(&mut zip, &rel_part, &serde_json::Map::new())?)?;
        unchanged_except_added(&old, &relation, &model(&relation, links[0]))?;
    } else if relation.nodes[0].children.len() != 1 {
        return Err(crate::coded(
            "package-unsupported",
            "Table creation added unrelated relationships",
        ));
    }
    let old_types = Tree::parse(&part_bytes(
        &mut zip,
        "[Content_Types].xml",
        &serde_json::Map::new(),
    )?)?;
    let new_types = Tree::parse(
        replacements["[Content_Types].xml"]
            .as_str()
            .ok_or_else(|| crate::coded("package-malformed", "Invalid table content types"))?
            .as_bytes(),
    )?;
    let added = json!({"ns":String::from_utf8_lossy(TYPE_NS),"tag":"Override","attrs":{"PartName":format!("/{table_part}"),"ContentType":"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"},"text":"","children":[]});
    unchanged_except_added(&old_types, &new_types, &added)?;
    Ok(())
}
fn conditional_addition(root: &Path, plan: &Value) -> Result<()> {
    let sheet = field(plan, "sheet")?;
    let selected = range(field(plan, "range")?)?;
    if u64::from(selected.1.0 - selected.0.0 + 1) * u64::from(selected.1.1 - selected.0.1 + 1)
        > 10000
    {
        return Err(crate::coded(
            "package-limit",
            "Oversized native conditional range",
        ));
    }
    let replacements = plan["replacements"].as_object().ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing conditional style replacement",
        )
    })?;
    if replacements.len() > 1
        || replacements.keys().any(|part| part != "xl/styles.xml")
        || !plan["copies"]
            .as_object()
            .is_some_and(|copies| copies.is_empty())
        || !plan["removed"]
            .as_array()
            .is_some_and(|removed| removed.is_empty())
    {
        return Err(crate::coded(
            "package-unsupported",
            "Conditional addition contains unrelated package movement",
        ));
    }
    let source = metadata_sheet(root, sheet)?;
    let prior = children(&source, "conditionalFormatting")?;
    let actual = fragment(
        one_fragment(plan, sheet, "conditionalFormatting")?.as_bytes(),
        "conditionalFormatting",
    )?;
    if actual.len() != prior.len() + 1 || actual[..prior.len()] != prior {
        return Err(crate::coded(
            "package-stale",
            "Conditional addition changed existing rules",
        ));
    }
    let added = actual.last().ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing new conditional collection",
        )
    })?;
    if added["attrs"] != json!({"sqref":field(plan,"range")?})
        || added["children"]
            .as_array()
            .is_none_or(|nodes| nodes.len() != 1)
    {
        return Err(crate::coded(
            "package-stale",
            "New conditional collection changed declared geometry",
        ));
    }
    let rule = &added["children"][0];
    let kind = field(&rule["attrs"], "type")?;
    let priority = source
        .nodes
        .iter()
        .filter(|node| node.ns == NS && node.name == "cfRule")
        .map(|node| {
            node.attrs
                .get("priority")
                .ok_or_else(|| {
                    crate::coded(
                        "internal-missing-field",
                        "Missing source conditional priority",
                    )
                })
                .and_then(|value| {
                    value.parse::<u32>().map_err(|_| {
                        crate::coded(
                            "internal-invalid-field",
                            "Invalid source conditional priority",
                        )
                    })
                })
        })
        .collect::<std::result::Result<Vec<_>, _>>()
        .map_err(err)?
        .into_iter()
        .max()
        .unwrap_or(0)
        .checked_add(1)
        .ok_or_else(|| crate::coded("package-limit", "Conditional priority overflow"))?;
    if rule["ns"] != String::from_utf8_lossy(NS).as_ref()
        || rule["tag"] != "cfRule"
        || rule["attrs"]["priority"] != priority.to_string()
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "New conditional rule priority or namespace is invalid",
        ));
    }
    let attrs = rule["attrs"]
        .as_object()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing conditional rule fields"))?;
    if kind == "duplicateValues" {
        if attrs.len() != 3 || rule["children"] != json!([]) {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid native duplicate-values rule",
            ));
        }
    } else if kind == "cellIs" {
        let nodes = rule["children"].as_array().ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing native conditional formula",
            )
        })?;
        if attrs.len() != 4
            || !["greaterThan", "lessThan"].contains(&field(&rule["attrs"], "operator")?)
            || nodes.len() != 1
            || nodes[0]["tag"] != "formula"
            || nodes[0]["children"] != json!([])
            || field(&nodes[0], "text")?
                .parse::<f64>()
                .map_err(err)?
                .is_finite()
                == false
        {
            return Err(crate::coded(
                "package-unsupported",
                "Unqualified native scalar conditional rule",
            ));
        }
    } else {
        return Err(crate::coded(
            "package-unsupported",
            "Unsupported native conditional addition",
        ));
    }
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let before = Tree::parse(&part_bytes(
        &mut zip,
        "xl/styles.xml",
        &serde_json::Map::new(),
    )?)?;
    let after = Tree::parse(
        replacements
            .get("xl/styles.xml")
            .map(|value| {
                value.as_str().map(str::as_bytes).ok_or_else(|| {
                    crate::coded("package-malformed", "Invalid differential style XML")
                })
            })
            .transpose()?
            .unwrap_or(&before.bytes),
    )?;
    before.root(NS, "styleSheet")?;
    after.root(NS, "styleSheet")?;
    let old = before.direct(0, "dxfs")?;
    let new = after
        .direct(0, "dxfs")?
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing new differential styles"))?;
    let count = old
        .map(|index| before.nodes[index].children.len())
        .unwrap_or(0);
    let dxf = field(&rule["attrs"], "dxfId")?
        .parse::<usize>()
        .map_err(err)?;
    let added = after.nodes[new].children.len() == count + 1;
    if !((added && dxf == count)
        || (!added && after.nodes[new].children.len() == count && dxf < count))
    {
        return Err(crate::coded(
            "package-unsupported",
            "Conditional rule has no exact qualified differential style",
        ));
    }
    let mut expected = model(&before, 0);
    let mut actual_styles = model(&after, 0);
    let actual_nodes = actual_styles["children"]
        .as_array_mut()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing style collections"))?;
    let at = actual_nodes
        .iter()
        .position(|node| node["tag"] == "dxfs")
        .ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing differential style model")
        })?;
    let entry = actual_nodes[at]["children"][dxf].clone();
    if added {
        actual_nodes[at]["children"]
            .as_array_mut()
            .ok_or_else(|| {
                crate::coded(
                    "internal-missing-field",
                    "Missing differential style entries",
                )
            })?
            .pop();
    }
    actual_nodes[at]["attrs"]["count"] = json!(count.to_string());
    if old.is_none() {
        actual_nodes.remove(at);
    } else {
        let expected_nodes = expected["children"]
            .as_array_mut()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing source style nodes"))?;
        let at = expected_nodes
            .iter()
            .position(|node| node["tag"] == "dxfs")
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing source dxfs"))?;
        expected_nodes[at]["attrs"]["count"] = json!(count.to_string());
    }
    if expected != actual_styles {
        return Err(crate::coded(
            "package-stale",
            "Conditional addition changed unrelated cell style geometry",
        ));
    }
    let fill = &entry["children"][0];
    let pattern = &fill["children"][0];
    let color = field(&pattern["children"][0]["attrs"], "rgb")?;
    if color.len() != 8
        || !color.starts_with("FF")
        || !color.bytes().all(|byte| byte.is_ascii_hexdigit())
        || entry != empty_differential_fill(color)
    {
        return Err(crate::coded(
            "package-unsupported",
            "Unqualified new conditional fill",
        ));
    }
    Ok(())
}
fn empty_differential_fill(color: &str) -> Value {
    let mut dxf = empty_node("dxf", json!({}));
    let mut fill = empty_node("fill", json!({}));
    let mut pattern = empty_node("patternFill", json!({"patternType":"solid"}));
    pattern["children"] = json!([
        empty_node("fgColor", json!({"rgb":color})),
        empty_node("bgColor", json!({"rgb":color}))
    ]);
    fill["children"] = json!([pattern]);
    dxf["children"] = json!([fill]);
    dxf
}

const DELETED_GEOMETRY: &str = "Deleting complete metadata geometry requires qualification";
thread_local! {
    /// Table columns added or removed by the change being written, so a formula
    /// that names a table column is adjusted the same way everywhere.
    static TABLE_CHANGES: std::cell::RefCell<Value> = const { std::cell::RefCell::new(Value::Null) };
}
fn rewrite_reference(
    reference: &str,
    owner: &str,
    target: &str,
    axis: &str,
    index: u32,
    count: u32,
    remove: bool,
) -> Result<String> {
    let tables = TABLE_CHANGES.with(|changes| changes.borrow().clone());
    let result = crate::formula::rewrite(
        &json!({"target":target,"axis":axis,"index":index,"count":count,"remove":remove,"tableChanges":tables,"items":[{"formula":reference,"owner":owner}]}),
    )?;
    let value = result["formulas"][0].as_str().ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing native metadata reference rewrite",
        )
    })?;
    if value.contains("#REF!") {
        return Err(crate::coded("package-refused", DELETED_GEOMETRY));
    }
    Ok(value.to_owned())
}
/// A range that ends just before an insertion grows to take in the new rows
/// or columns; any other range is returned as it is.
fn extended(reference: &str, columns: bool, index: u32, count: u32) -> String {
    let Ok((first, mut last)) = range(reference) else {
        return reference.to_owned();
    };
    let (end, limit) = if columns {
        (&mut last.1, 16384)
    } else {
        (&mut last.0, 1048576)
    };
    if index == 0 || *end != index - 1 || *end + count >= limit {
        return reference.to_owned();
    }
    *end += count;
    format!(
        "{}:{}",
        crate::xml::address(first.0, first.1),
        crate::xml::address(last.0, last.1)
    )
}
fn rewrite_nodes(
    node: &mut Value,
    moving: bool,
    owner: &str,
    target: &str,
    axis: &str,
    index: u32,
    count: u32,
    remove: bool,
) -> Result<()> {
    let tag = field(node, "tag")?.to_owned();
    if moving && tag == "autoFilter" {
        // A filter's criteria name their column by its place in the filtered
        // range: find that column on the sheet, move it, then renumber it
        // against the moved range.
        let first_column = |range: &str| -> Result<u32> {
            Ok(crate::xml::position(
                range
                    .split(':')
                    .next()
                    .unwrap_or(range)
                    .replace('$', "")
                    .as_str(),
            )?
            .1)
        };
        let before = first_column(
            node["attrs"]["ref"]
                .as_str()
                .ok_or_else(|| crate::coded("package-bad-input", "The filter has no range"))?,
        )?;
        let after = first_column(&rewrite_reference(
            node["attrs"]["ref"]
                .as_str()
                .ok_or_else(|| crate::coded("package-bad-input", "The filter has no range"))?,
            owner,
            target,
            axis,
            index,
            count,
            remove,
        )?)?;
        for child in node["children"]
            .as_array_mut()
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid metadata children"))?
        {
            if child["tag"] != "filterColumn" {
                continue;
            }
            let mut at = before
                + child["attrs"]["colId"]
                    .as_str()
                    .and_then(|text| text.parse::<u32>().ok())
                    .ok_or_else(|| {
                        crate::coded("package-refused", "The filter has an invalid column")
                    })?;
            if axis == "column" {
                if remove && at >= index && at < index + count {
                    return Err(crate::coded(
                        "filtered-column",
                        "A filtered column cannot be deleted; clear its filter first",
                    ));
                }
                if at >= index {
                    at = if remove { at - count } else { at + count };
                }
            }
            child["attrs"]["colId"] = json!(
                at.checked_sub(after)
                    .ok_or_else(|| crate::coded(
                        "package-refused",
                        "The filter has an invalid column"
                    ))?
                    .to_string()
            );
        }
    }
    let mut dropped = false;
    let attributes = node["attrs"]
        .as_object_mut()
        .ok_or_else(|| crate::coded("package-malformed", "Invalid metadata attributes"))?;
    for key in ["ref", "sqref"] {
        if moving
            && [
                "mergeCell",
                "dataValidation",
                "conditionalFormatting",
                "hyperlink",
                "autoFilter",
                "sortState",
                "sortCondition",
            ]
            .contains(&tag.as_str())
        {
            if let Some(value) = attributes.get_mut(key) {
                // Deleting every row or column an object covers deletes the
                // object, as in Excel: a merged range, a link, or that part of
                // a rule's or dropdown's range. A filter is not dropped this way.
                let droppable = [
                    "mergeCell",
                    "dataValidation",
                    "conditionalFormatting",
                    "hyperlink",
                ]
                .contains(&tag.as_str());
                let mut kept = Vec::new();
                for reference in value
                    .as_str()
                    .ok_or_else(|| {
                        crate::coded("internal-invalid-field", "Invalid metadata range")
                    })?
                    .split_whitespace()
                {
                    match rewrite_reference(reference, owner, target, axis, index, count, remove) {
                        // A merged range left with one cell is no longer a merge.
                        Ok(moved)
                            if tag == "mergeCell"
                                && moved.split_once(':').is_none_or(|(a, b)| a == b) => {}
                        // A rule or dropdown that ends on the row above (or the
                        // column to the left of) the insertion covers the new
                        // cells too: they take everything from that neighbour.
                        Ok(moved)
                            if !remove
                                && moved == reference
                                && ["dataValidation", "conditionalFormatting"]
                                    .contains(&tag.as_str()) =>
                        {
                            kept.push(extended(&moved, axis == "column", index, count))
                        }
                        Ok(moved) => kept.push(moved),
                        Err(error) if droppable && crate::sentence(&error) == DELETED_GEOMETRY => {}
                        Err(error) => return Err(error),
                    }
                }
                if kept.is_empty() {
                    dropped = true;
                }
                *value = json!(kept.join(" "));
            }
        }
    }
    // A link's "location" is a cell in this workbook only when the link has no
    // web address. With one, it is a place inside that web page and is left
    // alone, as is anything that does not read as a reference (a name, say).
    if tag == "hyperlink" && !attributes.keys().any(|key| key.ends_with(":id")) {
        if let Some(value) = attributes.get_mut("location") {
            let location = value.as_str().ok_or_else(|| {
                crate::coded("internal-invalid-field", "Invalid hyperlink location")
            })?;
            match rewrite_reference(location, owner, target, axis, index, count, remove) {
                Ok(moved) => *value = json!(moved),
                Err(error) if crate::sentence(&error) == DELETED_GEOMETRY => return Err(error),
                Err(_) => {}
            }
        }
    }
    if ["formula", "formula1", "formula2"].contains(&tag.as_str()) {
        let text = field(node, "text")?;
        if !(text.trim().starts_with('"') && text.trim().ends_with('"')) {
            node["text"] = json!(rewrite_reference(
                text, owner, target, axis, index, count, remove
            )?);
        }
    }
    if dropped {
        node["dropped"] = json!(true);
        return Ok(());
    }
    let children = node["children"]
        .as_array_mut()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid metadata children"))?;
    for child in children.iter_mut() {
        rewrite_nodes(child, moving, owner, target, axis, index, count, remove)?;
    }
    let before = children.len();
    children.retain(|child| child["dropped"] != true);
    let left = children.len();
    if left != before && ["mergeCells", "hyperlinks", "dataValidations"].contains(&tag.as_str()) {
        if left == 0 {
            node["dropped"] = json!(true);
        } else if node["attrs"].get("count").is_some() {
            node["attrs"]["count"] = json!(left.to_string());
        }
    }
    Ok(())
}
fn deleted_table_edges(
    root: &Path,
    sheet: &str,
    plan: &Value,
    columns: bool,
    index: u32,
    count: u32,
) -> Result<BTreeSet<String>> {
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let (directory, file) = sheet
        .rsplit_once('/')
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid table owner worksheet"))?;
    let relationship_part = format!("{directory}/_rels/{file}.rels");
    if !zip.file_names().any(|part| part == relationship_part) {
        return Ok(BTreeSet::new());
    }
    let relationships = Tree::parse(&part_bytes(
        &mut zip,
        &relationship_part,
        &serde_json::Map::new(),
    )?)?;
    relationships.root(PACKAGE_REL_NS, "Relationships")?;
    let removed = plan["removed"].as_array().ok_or_else(|| {
        crate::coded("internal-missing-field", "Missing insertion table removals")
    })?;
    let mut edges = BTreeSet::new();
    for edge in relationships.nodes[0]
        .children
        .iter()
        .map(|index| &relationships.nodes[*index])
    {
        if edge.ns != PACKAGE_REL_NS || edge.name != "Relationship" || !edge.children.is_empty() {
            return Err(crate::coded(
                "package-unsupported",
                "Opaque insertion table edge",
            ));
        }
        if edge.attrs.get("Type").map(String::as_str)
            != Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships/table")
        {
            continue;
        }
        if edge.attrs.contains_key("TargetMode") {
            return Err(crate::coded(
                "package-refused",
                "External insertion table edge",
            ));
        }
        let target = resolve(
            sheet,
            edge.attrs.get("Target").ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing insertion table target")
            })?,
        )?;
        if !removed.iter().any(|part| part == &target) {
            continue;
        }
        let table = Tree::parse(&part_bytes(&mut zip, &target, &serde_json::Map::new())?)?;
        table.root(NS, "table")?;
        let (first, last) = range(table.nodes[0].attrs.get("ref").ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing removed table range")
        })?)?;
        let (first, last) = if columns {
            (first.1, last.1)
        } else {
            (first.0, last.0)
        };
        if index > first || index + count <= last {
            return Err(crate::coded(
                "package-unsupported",
                "Insertion removed a table outside its complete deleted axis band",
            ));
        }
        if !edges.insert(
            edge.attrs
                .get("Id")
                .ok_or_else(|| {
                    crate::coded(
                        "internal-missing-field",
                        "Missing removed table relationship",
                    )
                })?
                .clone(),
        ) {
            return Err(crate::coded(
                "package-duplicate",
                "Duplicate removed table relationship",
            ));
        }
    }
    Ok(edges)
}
/// How many rows or columns stay frozen once rows or columns are inserted or
/// deleted, as Excel does it (watched in Excel for Mac, October 2026): a row
/// inserted above the last frozen row joins the frozen rows; one inserted
/// just below them does not; deleting frozen rows leaves fewer frozen.
fn frozen_after(frozen: u32, index: u32, count: u32, remove: bool) -> u32 {
    if remove {
        frozen - frozen.min(index + count).saturating_sub(index)
    } else if index < frozen {
        frozen + count
    } else {
        frozen
    }
}
fn panes(views: &mut [Value], columns: bool, index: u32, count: u32, remove: bool) -> Result<()> {
    let (split, limit) = if columns {
        ("xSplit", 16384)
    } else {
        ("ySplit", 1048576)
    };
    for view in views
        .iter_mut()
        .flat_map(|views| views["children"].as_array_mut().into_iter().flatten())
    {
        let Some(children) = view["children"].as_array_mut() else {
            continue;
        };
        let Some(at) = children.iter().position(|child| child["tag"] == "pane") else {
            continue;
        };
        let pane = &mut children[at];
        // A window split by dragging is placed in points, not rows: it stays
        // where it is on the screen.
        if !["frozen", "frozenSplit"].contains(&pane["attrs"]["state"].as_str().unwrap_or("")) {
            continue;
        }
        let number = |pane: &Value, key: &str| -> Result<u32> {
            match pane["attrs"][key].as_str() {
                None => Ok(0),
                Some(text) => text
                    .parse::<f64>()
                    .ok()
                    .filter(|value| *value >= 0.0 && value.fract() == 0.0 && *value < 1048576.0)
                    .map(|value| value as u32)
                    .ok_or_else(|| {
                        crate::coded(
                            "package-bad-input",
                            "The frozen panes are not a whole number of rows or columns",
                        )
                    }),
            }
        };
        let before = number(pane, split)?;
        let after = frozen_after(before, index, count, remove);
        if after == before {
            continue;
        }
        if after >= limit {
            return Err(crate::coded(
                "package-refused",
                "The frozen panes would run off the sheet",
            ));
        }
        let other = number(pane, if columns { "ySplit" } else { "xSplit" })?;
        let attributes = pane["attrs"].as_object_mut().unwrap();
        if after == 0 {
            attributes.remove(split);
        } else {
            attributes.insert(split.into(), json!(after.to_string()));
        }
        if after == 0 && other == 0 {
            // Nothing is frozen any more: the window is one pane again.
            children.remove(at);
            children.retain(|child| child["tag"] != "selection");
            continue;
        }
        let (rows, cols) = if columns {
            (other, after)
        } else {
            (after, other)
        };
        // The cell shown first below and right of the frozen area. It records
        // how far the person had scrolled; that is kept, moved by as much as
        // the frozen area grew or shrank on this axis.
        let shown = attributes
            .get("topLeftCell")
            .and_then(Value::as_str)
            .and_then(|cell| crate::xml::position(cell).ok())
            .unwrap_or((rows, cols));
        let moved = |at: u32| (at + after).saturating_sub(before).max(after);
        let shown = if columns {
            (shown.0.max(rows), moved(shown.1))
        } else {
            (moved(shown.0), shown.1.max(cols))
        };
        attributes.insert(
            "topLeftCell".into(),
            json!(crate::xml::address(shown.0, shown.1)),
        );
        let active = match (rows > 0, cols > 0) {
            (true, true) => "bottomRight",
            (true, false) => "bottomLeft",
            _ => "topRight",
        };
        attributes.insert("activePane".into(), json!(active));
        // A selection kept for a pane that no longer exists is not valid.
        let kept = |name: &str| match name {
            "bottomRight" => rows > 0 && cols > 0,
            "bottomLeft" => rows > 0,
            "topRight" => cols > 0,
            _ => true,
        };
        children.retain(|child| {
            child["tag"] != "selection"
                || kept(child["attrs"]["pane"].as_str().unwrap_or("topLeft"))
        });
    }
    Ok(())
}
/// Manual page breaks follow their rows or columns. A break is stored as the
/// number of the last row above it, which is the 0-based place of the row that
/// starts the new page; that row carries the break with it, and deleting that
/// row deletes the break.
fn breaks(lists: &mut Vec<Value>, index: u32, count: u32, remove: bool, limit: u32) -> Result<()> {
    for list in lists.iter_mut() {
        let entries = match list["children"].take() {
            Value::Array(entries) => entries,
            _ => {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid page breaks",
                ));
            }
        };
        let mut kept = Vec::new();
        for mut entry in entries {
            if entry["tag"] != "brk" {
                return Err(crate::coded(
                    "package-refused",
                    "Unrecognised page break record",
                ));
            }
            let mut at = attr_number(&entry, "id")?;
            if at >= index {
                if remove && at < index + count {
                    continue;
                }
                at = if remove { at - count } else { at + count };
            }
            if at >= limit {
                continue;
            }
            entry["attrs"]["id"] = json!(at.to_string());
            kept.push(entry);
        }
        let manual = kept
            .iter()
            .filter(|entry| ["1", "true"].contains(&entry["attrs"]["man"].as_str().unwrap_or("")))
            .count();
        if list["attrs"].get("count").is_some() {
            list["attrs"]["count"] = json!(kept.len().to_string());
        }
        if list["attrs"].get("manualBreakCount").is_some() {
            list["attrs"]["manualBreakCount"] = json!(manual.to_string());
        }
        if kept.is_empty() {
            list["dropped"] = json!(true);
        }
        list["children"] = Value::Array(kept);
    }
    lists.retain(|list| list["dropped"] != true);
    Ok(())
}
/// A pivot table remembers which cells it summarises. When those cells are on
/// the changed sheet the remembered range follows them, as in Excel. Returns
/// the rewritten part, or nothing when this part needs no change.
pub(super) fn pivot_source(
    root: &Path,
    input: &Value,
    plan: &Value,
    bytes: &[u8],
) -> Result<Option<Vec<u8>>> {
    let Some(mut start) = bytes
        .windows(15)
        .position(|window| window == b"worksheetSource")
    else {
        return Ok(None);
    };
    while start > 0 && bytes[start] != b'<' {
        start -= 1;
    }
    let end = start
        + bytes[start..]
            .iter()
            .position(|byte| *byte == b'>')
            .ok_or_else(|| crate::coded("package-malformed", "Unclosed pivot source"))?
        + 1;
    let mut reader = quick_xml::Reader::from_reader(&bytes[start..end]);
    let element = match reader.read_event().map_err(err)? {
        Event::Start(element) | Event::Empty(element) => element.into_owned(),
        _ => return Err(crate::coded("package-refused", "Unreadable pivot source")),
    };
    let attributes = attrs(&element)?;
    let (Some(reference), Some(sheet)) = (attributes.get("ref"), attributes.get("sheet")) else {
        return Ok(None);
    };
    // A source in another workbook has a relationship id; it does not move.
    if attributes.keys().any(|key| key.ends_with(":id")) {
        return Ok(None);
    }
    let inventory = worksheet_inventory(&local(root, field(input, "source")?)?)?;
    let target = inventory
        .iter()
        .find(|item| plan["target"] == item.path)
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing native insertion target"))?;
    if *sheet != target.name {
        return Ok(None);
    }
    let moved = rewrite_reference(
        reference,
        &target.name,
        &target.name,
        field(plan, "axis")?,
        plan["index"].as_u64().ok_or_else(|| crate::coded("internal-missing-field", "Missing insertion index"))? as u32,
        plan["count"].as_u64().ok_or_else(|| crate::coded("internal-missing-field", "Missing insertion count"))? as u32,
        plan["remove"] == true,
    )
    .map_err(|error| {
        if crate::sentence(&error) == DELETED_GEOMETRY {
            crate::coded(
                "pivot-source-deleted",
                "These rows or columns are all of a pivot table's data; delete the pivot table first",
            )
        } else {
            error
        }
    })?;
    if moved == *reference {
        return Ok(None);
    }
    let text = std::str::from_utf8(&bytes[start..end]).map_err(err)?;
    let mut escaped = String::new();
    escape(&moved, true, &mut escaped);
    let at = text
        .find(" ref=")
        .ok_or_else(|| crate::coded("package-refused", "Unreadable pivot source range"))?
        + 5;
    let quote = text[at..]
        .chars()
        .next()
        .ok_or_else(|| crate::coded("package-refused", "Unreadable pivot source range"))?;
    let close = at
        + 1
        + text[at + 1..]
            .find(quote)
            .ok_or_else(|| crate::coded("package-refused", "Unreadable pivot source range"))?;
    let mut out = bytes[..start + at + 1].to_vec();
    out.extend_from_slice(escaped.as_bytes());
    out.extend_from_slice(&bytes[start + close..]);
    Ok(Some(out))
}
/// The sheet records that move when rows or columns are inserted or deleted.
const MOVED_RECORDS: [&str; 10] = [
    "sheetViews",
    "cols",
    "autoFilter",
    "mergeCells",
    "conditionalFormatting",
    "dataValidations",
    "hyperlinks",
    "rowBreaks",
    "colBreaks",
    "tableParts",
];
/// Work out, for every worksheet, the records that change when rows or columns
/// are inserted or deleted, as XML ready to be written in place of the old
/// ones. This is the only place those records are rewritten.
pub(super) fn insertion(root: &Path, input: &Value, plan: &Value) -> Result<Value> {
    let inventory = worksheet_inventory(&local(root, field(input, "source")?)?)?;
    let target = inventory
        .iter()
        .find(|sheet| plan["target"] == sheet.path)
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing native insertion target"))?;
    let axis = field(plan, "axis")?;
    let columns = axis == "column";
    if !["row", "column"].contains(&axis) || !plan["remove"].is_boolean() {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid insertion geometry kind",
        ));
    }
    let index = integer(plan.get("index"), if columns { 16384 } else { 1048576 })?;
    let count = plan["count"]
        .as_u64()
        .filter(|count| {
            *count > 0
                && *count
                    <= u64::from(if columns { 16384u32 } else { 1048576u32 }) - u64::from(index)
        })
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid insertion geometry count"))?
        as u32;
    let remove = plan["remove"] == true;
    let paths = inventory
        .iter()
        .map(|sheet| sheet.path.as_str())
        .collect::<BTreeSet<_>>();
    if plan["replacements"]
        .as_object()
        .is_none_or(|parts| parts.keys().any(|part| paths.contains(part.as_str())))
        || plan["removed"].as_array().is_none_or(|parts| {
            parts
                .iter()
                .any(|part| part.as_str().is_none_or(|part| paths.contains(part)))
        })
    {
        return Err(crate::coded(
            "package-unsupported",
            "Insertion metadata cannot replace worksheet cells or identities",
        ));
    }
    TABLE_CHANGES.with(|changes| *changes.borrow_mut() = plan["tableChanges"].clone());
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let mut fragments = serde_json::Map::new();
    for owner in &inventory {
        let path = &owner.path;
        let moving = owner.path == target.path;
        let mut found = records(
            &part_bytes(&mut zip, path, &serde_json::Map::new())?,
            &MOVED_RECORDS,
        )?;
        let mut parts = serde_json::Map::new();
        for tag in MOVED_RECORDS {
            let Some(source) = found.remove(tag) else {
                continue;
            };
            let mut expected = source.clone();
            if tag == "sheetViews" {
                if moving {
                    panes(&mut expected, columns, index, count, remove)?;
                }
            } else if tag == "rowBreaks" || tag == "colBreaks" {
                if moving && columns == (tag == "colBreaks") {
                    let limit = if columns { 16384 } else { 1048576 };
                    breaks(&mut expected, index, count, remove, limit)?;
                }
            } else if tag == "cols" {
                if moving && columns {
                    for container in &mut expected {
                        let entries = container["children"].as_array_mut().ok_or_else(|| {
                            crate::coded("internal-missing-field", "Missing source columns")
                        })?;
                        let mut retained = Vec::new();
                        for mut entry in entries.drain(..) {
                            if entry["tag"] != "col"
                                || entry["ns"] != String::from_utf8_lossy(NS).as_ref()
                            {
                                return Err(crate::coded(
                                    "package-unsupported",
                                    "Opaque insertion columns",
                                ));
                            }
                            let mut low =
                                attr_number(&entry, "min")?.checked_sub(1).ok_or_else(|| {
                                    crate::coded(
                                        "internal-invalid-field",
                                        "Invalid insertion column",
                                    )
                                })?;
                            let mut high =
                                attr_number(&entry, "max")?.checked_sub(1).ok_or_else(|| {
                                    crate::coded(
                                        "internal-invalid-field",
                                        "Invalid insertion column",
                                    )
                                })?;
                            if high < low || high >= 16384 {
                                return Err(crate::coded(
                                    "package-malformed",
                                    "Invalid source column span",
                                ));
                            }
                            if remove && low >= index && high < index + count {
                                continue;
                            }
                            if low >= index {
                                low = if remove {
                                    index.max(low.saturating_sub(count))
                                } else {
                                    low.checked_add(count).ok_or_else(|| {
                                        crate::coded("package-limit", "Column overflow")
                                    })?
                                };
                            }
                            if high >= index {
                                high = if remove {
                                    index.saturating_sub(1).max(high.saturating_sub(count))
                                } else {
                                    high.checked_add(count).ok_or_else(|| {
                                        crate::coded("package-limit", "Column overflow")
                                    })?
                                };
                            }
                            // A new column takes the width and format of the
                            // column on its left, as in Excel (but is not hidden).
                            if !remove
                                && index > 0
                                && high == index - 1
                                && !["1", "true"]
                                    .contains(&entry["attrs"]["hidden"].as_str().unwrap_or(""))
                            {
                                high += count;
                            }
                            // Formatting that runs to the last column (very common)
                            // stops at the edge when a column is inserted, as in
                            // Excel; a record pushed wholly off the sheet is dropped.
                            if low >= 16384 {
                                continue;
                            }
                            high = high.min(16383);
                            if high >= 16384 || high < low {
                                return Err(crate::coded(
                                    "package-limit",
                                    "Moved column metadata exceeds worksheet",
                                ));
                            }
                            entry["attrs"]["min"] = json!((low + 1).to_string());
                            entry["attrs"]["max"] = json!((high + 1).to_string());
                            retained.push(entry);
                        }
                        *entries = retained;
                    }
                }
            } else if tag == "tableParts" {
                if moving && remove && expected.len() == 1 {
                    let deleted = deleted_table_edges(root, path, plan, columns, index, count)?;
                    let entries = expected[0]["children"].as_array_mut().ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing source table owner edges")
                    })?;
                    entries.retain(|entry| {
                        !entry["attrs"].as_object().is_some_and(|attributes| {
                            attributes.iter().any(|(key, id)| {
                                key.ends_with(":id")
                                    && id.as_str().is_some_and(|id| deleted.contains(id))
                            })
                        })
                    });
                    let count = entries.len();
                    if count == 0 {
                        expected.clear();
                    } else if expected[0]["attrs"].get("count").is_some() {
                        expected[0]["attrs"]["count"] = json!(count.to_string());
                    }
                }
            } else {
                for node in &mut expected {
                    rewrite_nodes(
                        node,
                        moving,
                        &owner.name,
                        &target.name,
                        axis,
                        index,
                        count,
                        remove,
                    )?;
                }
                expected.retain(|node| node["dropped"] != true);
            }
            if expected != source {
                let mut xml = String::new();
                for node in &expected {
                    to_xml(node, &mut xml)?;
                }
                parts.insert(tag.into(), json!(xml));
            }
        }
        if !parts.is_empty() {
            fragments.insert(path.clone(), Value::Object(parts));
        }
    }
    Ok(Value::Object(fragments))
}

#[cfg(test)]
mod tests {
    use super::*;
    fn workbook(file: &Path, sheets: &[(&str, &str, &str)]) {
        let mut zip = ZipWriter::new(File::create(file).unwrap());
        let mut workbook = format!(
            "<workbook xmlns=\"{}\" xmlns:r=\"{}\"><sheets>",
            String::from_utf8_lossy(NS),
            String::from_utf8_lossy(REL_NS)
        );
        let mut relationships = format!(
            "<Relationships xmlns=\"{}\">",
            String::from_utf8_lossy(PACKAGE_REL_NS)
        );
        let mut types = format!(
            "<Types xmlns=\"{}\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>",
            String::from_utf8_lossy(TYPE_NS)
        );
        for (index, (name, path, xml)) in sheets.iter().enumerate() {
            workbook.push_str(&format!(
                "<sheet name=\"{name}\" sheetId=\"{}\" r:id=\"rId{}\"/>",
                index + 1,
                index + 1
            ));
            relationships.push_str(&format!(
                "<Relationship Id=\"rId{}\" Type=\"{}/worksheet\" Target=\"{}\"/>",
                index + 1,
                String::from_utf8_lossy(REL_NS),
                path.strip_prefix("xl/").unwrap()
            ));
            types.push_str(&format!("<Override PartName=\"/{path}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"));
            zip.start_file(*path, SimpleFileOptions::default()).unwrap();
            zip.write_all(xml.as_bytes()).unwrap();
        }
        workbook.push_str("</sheets></workbook>");
        relationships.push_str("</Relationships>");
        types.push_str("</Types>");
        let root_relationship = format!(
            "<Relationships xmlns=\"{}\"><Relationship Id=\"rId1\" Type=\"{}/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>",
            String::from_utf8_lossy(PACKAGE_REL_NS),
            String::from_utf8_lossy(REL_NS)
        );
        for (path, xml) in [
            ("xl/workbook.xml", workbook),
            ("xl/_rels/workbook.xml.rels", relationships),
            ("[Content_Types].xml", types),
            ("_rels/.rels", root_relationship),
        ] {
            zip.start_file(path, SimpleFileOptions::default()).unwrap();
            zip.write_all(xml.as_bytes()).unwrap();
        }
        zip.finish().unwrap();
    }
    fn fixture(metadata: &str) -> tempfile::TempDir {
        let root = tempfile::tempdir().unwrap();
        let xml = format!(
            "<worksheet xmlns=\"{}\">{metadata}<sheetData/></worksheet>",
            String::from_utf8_lossy(NS)
        );
        for file in ["source.xlsx", "metadata.xlsx"] {
            workbook(
                &root.path().join(file),
                &[("Sheet1", "xl/worksheets/sheet1.xml", &xml)],
            );
        }
        root
    }
    fn layout_request(command: Value, tag: &str) -> Value {
        let mut input = json!({"source":"source.xlsx","sheet":"xl/worksheets/sheet1.xml","tag":tag,"command":command});
        if tag == "mergeCells" {
            input["kind"] = input["command"]["operation"].clone();
            input["range"] = input["command"]["range"].clone();
        }
        input
    }
    #[test]
    fn merge_admission_rejects_an_extra_unrequested_range() {
        let root = fixture("");
        let input = layout_request(
            json!({"operation":"mergeCells","sheet":"Sheet1","range":"D4:G6"}),
            "mergeCells",
        );
        assert!(
            layout(
                root.path(),
                &input,
                Some(b"<mergeCells count=\"1\"><mergeCell ref=\"D4:G6\"/></mergeCells>")
            )
            .is_ok()
        );
        assert!(layout(root.path(),&input,Some(b"<mergeCells count=\"2\"><mergeCell ref=\"D4:G6\"/><mergeCell ref=\"Z80:AA80\"/></mergeCells>")).is_err());
    }
    #[test]
    fn unmerge_admission_preserves_other_existing_ranges() {
        let root = fixture(
            "<mergeCells count=\"2\"><mergeCell ref=\"D4:G6\"/><mergeCell ref=\"Z80:AA80\"/></mergeCells>",
        );
        let input = layout_request(
            json!({"operation":"unmergeCells","sheet":"Sheet1","range":"D4:G6"}),
            "mergeCells",
        );
        assert!(
            layout(
                root.path(),
                &input,
                Some(b"<mergeCells count=\"1\"><mergeCell ref=\"Z80:AA80\"/></mergeCells>")
            )
            .is_ok()
        );
        assert!(layout(root.path(), &input, Some(b"")).is_err());
    }
    #[test]
    fn resize_column_admission_preserves_styles_on_split_intervals() {
        let root = fixture("<cols><col min=\"2\" max=\"4\" width=\"8\" style=\"1\"/></cols>");
        let input = layout_request(
            json!({"operation":"resizeColumn","sheet":"Sheet1","index":2,"width":12}),
            "cols",
        );
        let good=b"<cols><col min=\"2\" max=\"2\" width=\"8\" style=\"1\"/><col min=\"3\" max=\"3\" width=\"12\" style=\"1\" customWidth=\"1\"/><col min=\"4\" max=\"4\" width=\"8\" style=\"1\"/></cols>";
        assert!(layout(root.path(), &input, Some(good)).is_ok());
        let forged = String::from_utf8(good.to_vec())
            .unwrap()
            .replace("style=\"1\" customWidth", "style=\"2\" customWidth");
        assert!(layout(root.path(), &input, Some(forged.as_bytes())).is_err());
    }
    #[test]
    fn unhide_admission_has_no_caller_fragment_or_foreign_target() {
        let root = fixture("");
        let input = json!({"source":"source.xlsx","sheet":"xl/worksheets/sheet1.xml","command":{"operation":"unhideAll","sheet":"Sheet1"}});
        assert!(layout(root.path(), &input, None).is_ok());
        assert!(layout(root.path(), &input, Some(b"<cols/>")).is_err());
        let mut foreign = input.clone();
        foreign["command"]["sheet"] = json!("Other");
        assert!(layout(root.path(), &foreign, None).is_err());
        let mut extra = input.clone();
        extra["command"]["range"] = json!("A1:B2");
        assert!(layout(root.path(), &extra, None).is_err());
        for body in [
            "<sheetProtection/>",
            "<autoFilter ref=\"A1:B2\"/>",
            "<tableParts/>",
        ] {
            let protected = fixture(body);
            assert!(layout(protected.path(), &input, None).is_err());
        }
    }
    #[test]
    fn freeze_admission_keeps_existing_selection_and_other_views() {
        let root = fixture(
            "<sheetViews><sheetView workbookViewId=\"0\"><selection activeCell=\"B2\" sqref=\"B2\"/></sheetView><sheetView workbookViewId=\"1\"/></sheetViews>",
        );
        let input = layout_request(
            json!({"operation":"freeze","sheet":"Sheet1","rows":1,"columns":0}),
            "sheetViews",
        );
        let good=b"<sheetViews><sheetView workbookViewId=\"0\"><pane xSplit=\"0\" ySplit=\"1\" topLeftCell=\"A2\" state=\"frozen\" activePane=\"bottomLeft\"/><selection activeCell=\"B2\" sqref=\"B2\"/></sheetView><sheetView workbookViewId=\"1\"/></sheetViews>";
        assert!(layout(root.path(), &input, Some(good)).is_ok());
        let forged = String::from_utf8(good.to_vec())
            .unwrap()
            .replace("workbookViewId=\"1\"", "workbookViewId=\"2\"");
        assert!(layout(root.path(), &input, Some(forged.as_bytes())).is_err());
    }
    #[test]
    fn duplicate_mapping_is_derived_from_the_native_sheet_graph() {
        let root = fixture("");
        let xml = format!(
            "<worksheet xmlns=\"{}\"><sheetData/></worksheet>",
            String::from_utf8_lossy(NS)
        );
        workbook(
            &root.path().join("package.xlsx"),
            &[
                ("Sheet1", "xl/worksheets/sheet1.xml", &xml),
                ("Copy", "xl/worksheets/sheet2.xml", &xml),
            ],
        );
        let mut plan = json!({"kind":"duplicateSheet","sheet":"xl/worksheets/sheet1.xml","rename":{"old":"Sheet1","new":"Copy"},"copyOnly":true,"copies":{"xl/worksheets/sheet2.xml":"xl/worksheets/sheet1.xml"},"replacements":{},"removed":[],"fragments":{}});
        let input = json!({"source":"source.xlsx"});
        let mapping = package_mapping(root.path(), &input, &plan, "package.xlsx").unwrap();
        match mapping {
            NativeMapping::Sheets { predecessors } => assert_eq!(
                predecessors["xl/worksheets/sheet2.xml"],
                Some("xl/worksheets/sheet1.xml".into())
            ),
            _ => panic!("Unexpected mapping"),
        }
        plan["copies"]["xl/worksheets/sheet2.xml"] = json!("xl/worksheets/foreign.xml");
        assert!(package_mapping(root.path(), &input, &plan, "package.xlsx").is_err());
    }
    #[test]
    fn insertion_ignores_records_offered_from_outside_and_never_replaces_a_sheet() {
        let root = fixture("");
        let input = json!({"source":"source.xlsx"});
        let mut plan = json!({"axis":"row","remove":false,"target":"xl/worksheets/sheet1.xml","index":2,"count":1,"replacements":{},"removed":[],"fragments":{"xl/worksheets/sheet1.xml":{"sheetData":"<sheetData/>"}}});
        // Nothing on the sheet moves, so nothing is written, whatever was offered.
        assert_eq!(insertion(root.path(), &input, &plan).unwrap(), json!({}));
        plan["replacements"] = json!({"xl/worksheets/sheet1.xml":"<worksheet/>"});
        assert!(insertion(root.path(), &input, &plan).is_err());
    }
    fn moved(metadata: &str, axis: &str, index: u32, remove: bool) -> Value {
        let root = fixture(metadata);
        let plan = json!({"axis":axis,"remove":remove,"target":"xl/worksheets/sheet1.xml","index":index,"count":1,"replacements":{},"removed":[]});
        insertion(root.path(), &json!({"source":"source.xlsx"}), &plan).unwrap()
            ["xl/worksheets/sheet1.xml"]
            .clone()
    }
    #[test]
    fn insertion_moves_frozen_panes_the_way_excel_was_seen_to() {
        let frozen = concat!(
            "<sheetViews><sheetView workbookViewId=\"0\">",
            "<pane ySplit=\"2\" topLeftCell=\"A3\" activePane=\"bottomLeft\" state=\"frozen\"/>",
            "<selection pane=\"bottomLeft\" activeCell=\"A3\" sqref=\"A3\"/></sheetView></sheetViews>"
        );
        // Inside the frozen rows: one more row is frozen.
        let inside = moved(frozen, "row", 1, false);
        assert!(
            inside["sheetViews"]
                .as_str()
                .unwrap()
                .contains("ySplit=\"3\"")
        );
        assert!(
            inside["sheetViews"]
                .as_str()
                .unwrap()
                .contains("topLeftCell=\"A4\"")
        );
        // Just below them, or a column anywhere: nothing changes.
        assert!(moved(frozen, "row", 2, false).get("sheetViews").is_none());
        assert!(
            moved(frozen, "column", 0, false)
                .get("sheetViews")
                .is_none()
        );
        // Deleting a frozen row leaves one frozen.
        let fewer = moved(frozen, "row", 0, true);
        assert!(
            fewer["sheetViews"]
                .as_str()
                .unwrap()
                .contains("ySplit=\"1\"")
        );
        assert!(
            fewer["sheetViews"]
                .as_str()
                .unwrap()
                .contains("topLeftCell=\"A2\"")
        );
        // A person who had scrolled down keeps their place below the frozen rows.
        let scrolled = moved(
            &frozen.replace("topLeftCell=\"A3\"", "topLeftCell=\"A30\""),
            "row",
            1,
            false,
        );
        assert!(
            scrolled["sheetViews"]
                .as_str()
                .unwrap()
                .contains("topLeftCell=\"A31\"")
        );
        // A window split by dragging is measured in points and is left alone.
        let split = "<sheetViews><sheetView workbookViewId=\"0\"><pane ySplit=\"2400\" topLeftCell=\"A9\"/></sheetView></sheetViews>";
        assert!(moved(split, "row", 1, false).get("sheetViews").is_none());
    }
    #[test]
    fn insertion_moves_page_breaks_and_keeps_prefixed_records_intact() {
        let breaks = "<rowBreaks count=\"2\" manualBreakCount=\"2\"><brk id=\"4\" max=\"16383\" man=\"1\"/><brk id=\"7\" max=\"16383\" man=\"1\"/></rowBreaks>";
        assert_eq!(
            moved(breaks, "row", 4, true)["rowBreaks"],
            "<rowBreaks count=\"1\" manualBreakCount=\"1\"><brk id=\"6\" max=\"16383\" man=\"1\"/></rowBreaks>"
        );
        assert!(moved(breaks, "column", 1, false).get("rowBreaks").is_none());
        // A rule with an extension child written with a prefix comes back with it,
        // attributes in their original order.
        let tail = concat!(
            "<extLst><ext xmlns:x14=\"urn:x14\" uri=\"{1}\"><x14:id>{A-1}</x14:id></ext></extLst>",
            "</cfRule></conditionalFormatting>"
        );
        let rule = format!(
            "<conditionalFormatting sqref=\"C3:C8\"><cfRule type=\"expression\" priority=\"1\">{}{tail}",
            "<formula>$C3&gt;\"a&amp;b\"</formula>"
        );
        assert_eq!(
            moved(&rule, "row", 1, false)["conditionalFormatting"],
            format!(
                "<conditionalFormatting sqref=\"C4:C9\"><cfRule type=\"expression\" priority=\"1\">{}{tail}",
                "<formula>$C4&gt;\"a&amp;b\"</formula>"
            )
        );
    }
    #[test]
    fn a_pivot_table_source_range_follows_its_rows() {
        let root = fixture("");
        let input = json!({"source":"source.xlsx"});
        let plan = |index: u32, count: u32, remove: bool| json!({"axis":"row","remove":remove,"target":"xl/worksheets/sheet1.xml","index":index,"count":count});
        let cache = |source: &str| {
            format!(
                "<pivotCacheDefinition><cacheSource type=\"worksheet\">{source}</cacheSource><cacheFields/></pivotCacheDefinition>"
            )
        };
        let on_sheet = cache("<worksheetSource ref=\"A1:D20\" sheet=\"Sheet1\"/>");
        let moved = pivot_source(root.path(), &input, &plan(5, 2, false), on_sheet.as_bytes())
            .unwrap()
            .unwrap();
        assert_eq!(
            String::from_utf8(moved).unwrap(),
            cache("<worksheetSource ref=\"A1:D22\" sheet=\"Sheet1\"/>")
        );
        // Rows below the data, another sheet, or a source given by name: unchanged.
        assert!(
            pivot_source(
                root.path(),
                &input,
                &plan(30, 1, false),
                on_sheet.as_bytes()
            )
            .unwrap()
            .is_none()
        );
        let elsewhere = cache("<worksheetSource ref=\"A1:D20\" sheet=\"Other\"/>");
        assert!(
            pivot_source(
                root.path(),
                &input,
                &plan(5, 2, false),
                elsewhere.as_bytes()
            )
            .unwrap()
            .is_none()
        );
        let named = cache("<worksheetSource name=\"Sales\"/>");
        assert!(
            pivot_source(root.path(), &input, &plan(5, 2, false), named.as_bytes())
                .unwrap()
                .is_none()
        );
        // Deleting every row of the data is refused with a reason.
        let error =
            pivot_source(root.path(), &input, &plan(0, 20, true), on_sheet.as_bytes()).unwrap_err();
        assert!(error.contains("pivot table"), "{error}");
    }
}
