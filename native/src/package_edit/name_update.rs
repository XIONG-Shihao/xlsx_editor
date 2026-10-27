//! Retarget static names and remove unused names with scoped formula admission.
//! Dynamic/object consumers stay protected; print-name neighbors are preserved.
use super::print_settings::Tree;
use super::*;

fn print_builtin(name: &str) -> bool {
    ["_xlnm.print_area", "_xlnm.print_titles"].contains(&name.to_ascii_lowercase().as_str())
}
fn print_name(node: &super::print_settings::Node, tree: &Tree, scope: u64) -> Result<()> {
    let sheet = sheets(tree)?
        .get(scope as usize)
        .cloned()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing print-name sheet"))?;
    let prefix = format!("'{}'!", sheet.replace('\'', "''"));
    let mut text = node.text.as_str();
    let mut count = 0;
    loop {
        let remaining = text.strip_prefix(&prefix).ok_or_else(|| {
            crate::coded(
                "name-bad-input",
                "Print name needs its actual scoped worksheet",
            )
        })?;
        let (part, next) = remaining
            .split_once(',')
            .map(|(a, b)| (a, Some(b)))
            .unwrap_or((remaining, None));
        let fields = part.split(':').collect::<Vec<_>>();
        if fields.len() != 2 {
            return Err(crate::coded(
                "name-unsupported",
                "Unqualified print-name reference",
            ));
        }
        if node.attrs["name"].eq_ignore_ascii_case("_xlnm.Print_Area") {
            if count > 0 {
                return Err(crate::coded(
                    "name-unsupported",
                    "Disconnected print area requires qualification",
                ));
            }
            static_range(&format!("{prefix}{part}"), tree)?;
        } else {
            let values = fields
                .iter()
                .map(|v| {
                    v.strip_prefix('$').ok_or_else(|| {
                        crate::coded("name-refused", "Absolute print titles required")
                    })
                })
                .collect::<Result<Vec<_>>>()?;
            let row = values
                .iter()
                .all(|v| !v.is_empty() && v.bytes().all(|c| c.is_ascii_digit()));
            let col = values
                .iter()
                .all(|v| !v.is_empty() && v.bytes().all(|c| c.is_ascii_uppercase()));
            if row {
                crate::bulk::bounds(&format!("A{}:A{}", values[0], values[1]))?;
            } else if col {
                crate::bulk::bounds(&format!("{}1:{}1", values[0], values[1]))?;
            } else {
                return Err(crate::coded(
                    "name-unsupported",
                    "Unqualified print title literal",
                ));
            }
        }
        count += 1;
        if count > 2 {
            return Err(crate::coded(
                "name-unsupported",
                "Excessive print-name areas",
            ));
        }
        match next {
            Some(next) => text = next,
            None => break,
        }
    }
    Ok(())
}

fn names(tree: &Tree) -> Result<Vec<usize>> {
    tree.root(NS, "workbook")?;
    let containers = tree.nodes[0]
        .children
        .iter()
        .filter(|i| tree.nodes[**i].name == "definedNames")
        .copied()
        .collect::<Vec<_>>();
    if containers.len() != 1 || tree.nodes[containers[0]].ns != NS {
        return Err(crate::coded(
            "name-ambiguous",
            "Ambiguous static name container",
        ));
    }
    let parent = &tree.nodes[containers[0]];
    if parent
        .attrs
        .keys()
        .any(|key| key != "xmlns" && !key.starts_with("xmlns:"))
        || !parent.text.trim().is_empty()
    {
        return Err(crate::coded(
            "name-unsupported",
            "Extended static name container",
        ));
    }
    let result = parent.children.clone();
    if result.len() > 1000
        || result
            .iter()
            .any(|i| tree.nodes[*i].ns != NS || tree.nodes[*i].name != "definedName")
    {
        return Err(crate::coded(
            "name-unsupported",
            "Extended or excessive static names require qualification",
        ));
    }
    let mut identities = BTreeSet::new();
    for index in &result {
        let node = &tree.nodes[*index];
        let name = node
            .attrs
            .get("name")
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing static name"))?;
        if !super::name_addition::legal_name(name) && !print_builtin(name)
            || !node.children.is_empty()
            || node
                .attrs
                .keys()
                .any(|key| !["name", "localSheetId"].contains(&key.as_str()))
        {
            return Err(crate::coded(
                "name-protected",
                "Extended or dependent static name is protected",
            ));
        }
        let scope = node
            .attrs
            .get("localSheetId")
            .map(|v| v.parse::<u64>().map_err(err))
            .transpose()?;
        if scope.is_some_and(|v| v >= sheets(tree).map(|s| s.len() as u64).unwrap_or(0))
            || !identities.insert((name.to_ascii_lowercase(), scope))
        {
            return Err(crate::coded(
                "name-ambiguous",
                "Ambiguous or invalid static name scope",
            ));
        }
        if print_builtin(name) {
            print_name(
                node,
                tree,
                scope.ok_or_else(|| {
                    crate::coded("name-bad-input", "Print names need worksheet scope")
                })?,
            )?;
        } else {
            static_range(&node.text, tree)?;
        }
    }
    Ok(result)
}
fn sheets(tree: &Tree) -> Result<Vec<String>> {
    let containers = tree.nodes[0]
        .children
        .iter()
        .copied()
        .filter(|index| tree.nodes[*index].name == "sheets")
        .collect::<Vec<_>>();
    if containers.len() != 1 || tree.nodes[containers[0]].ns != NS {
        return Err(crate::coded(
            "name-ambiguous",
            "Ambiguous workbook sheet identity",
        ));
    }
    let mut result = Vec::new();
    for index in &tree.nodes[containers[0]].children {
        let node = &tree.nodes[*index];
        if node.ns != NS || node.name != "sheet" || !node.children.is_empty() {
            return Err(crate::coded(
                "name-unsupported",
                "Extended workbook sheet identity",
            ));
        }
        let name = node
            .attrs
            .get("name")
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing workbook sheet name"))?;
        if result.contains(name) {
            return Err(crate::coded(
                "name-duplicate",
                "Duplicate workbook sheet identity",
            ));
        }
        result.push(name.clone());
    }
    if result.is_empty() {
        return Err(crate::coded(
            "internal-missing-field",
            "Missing workbook sheets",
        ));
    }
    Ok(result)
}
fn static_range(formula: &str, tree: &Tree) -> Result<()> {
    let (quoted, range) = formula.rsplit_once('!').ok_or_else(|| {
        crate::coded(
            "name-unsupported",
            "Only a static name rectangle is qualified",
        )
    })?;
    let sheet = quoted
        .strip_prefix('\'')
        .and_then(|value| value.strip_suffix('\''))
        .ok_or_else(|| {
            crate::coded(
                "name-unsupported",
                "Only a quoted static name sheet is qualified",
            )
        })?
        .replace("''", "'");
    if quoted != format!("'{}'", sheet.replace('\'', "''")) {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid static name sheet quoting",
        ));
    }
    if !sheets(tree)?.contains(&sheet) {
        return Err(crate::coded(
            "internal-missing-field",
            "Static name sheet is missing",
        ));
    }
    let pair = range.split(':').collect::<Vec<_>>();
    if pair.len() != 2 {
        return Err(crate::coded(
            "name-unsupported",
            "Only a static name rectangle is qualified",
        ));
    }
    for part in &pair {
        let bytes = part.as_bytes();
        if bytes.first() != Some(&b'$') {
            return Err(crate::coded(
                "name-bad-input",
                "Static name must be absolute",
            ));
        }
        let col = part[1..]
            .find('$')
            .ok_or_else(|| crate::coded("name-bad-input", "Static name must be absolute"))?
            + 1;
        if col == 1
            || !part[1..col].bytes().all(|ch| ch.is_ascii_uppercase())
            || !part[col + 1..].bytes().all(|ch| ch.is_ascii_digit())
            || part[col + 1..].is_empty()
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid absolute name rectangle",
            ));
        }
    }
    crate::bulk::bounds(&range.replace('$', ""))?;
    Ok(())
}
fn envelope(tree: &Tree, index: usize) -> Result<Vec<u8>> {
    let node = &tree.nodes[index];
    if !node.children.is_empty() {
        return Err(crate::coded(
            "name-protected",
            "Extended static name text is protected",
        ));
    }
    let bytes = tree.raw(index);
    let opening = bytes
        .iter()
        .position(|ch| *ch == b'>')
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing name opening"))?
        + 1;
    let close = bytes
        .windows(2)
        .rposition(|pair| pair == b"</")
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing name body"))?;
    let mut out = tree.bytes.clone();
    out.drain(node.start + opening..node.start + close);
    Ok(out)
}
fn legal_name(name: &str) -> bool {
    let chars = name.as_bytes();
    if chars.is_empty()
        || chars.len() > 255
        || !(chars[0].is_ascii_alphabetic() || chars[0] == b'_')
        || chars
            .iter()
            .any(|ch| !(ch.is_ascii_alphanumeric() || [b'_', b'.'].contains(ch)))
    {
        return false;
    }
    let upper = name.to_ascii_uppercase();
    if upper.starts_with("_XLNM.") || ["R", "C"].contains(&upper.as_str()) {
        return false;
    }
    let letters = chars
        .iter()
        .take_while(|ch| ch.is_ascii_alphabetic())
        .count();
    if (1..=3).contains(&letters)
        && letters < chars.len()
        && chars[letters..].iter().all(|ch| ch.is_ascii_digit())
    {
        return false;
    }
    if let Some(tail) = upper.strip_prefix('R') {
        if let Some((row, col)) = tail.split_once('C') {
            if !row.is_empty()
                && !col.is_empty()
                && row.bytes().all(|ch| ch.is_ascii_digit())
                && col.bytes().all(|ch| ch.is_ascii_digit())
            {
                return false;
            }
        }
    }
    true
}
fn references(
    zip: &mut ZipArchive<crate::cancel::Reader<File>>,
    workbook: &Tree,
    plan: &Value,
) -> Result<()> {
    if workbook
        .nodes
        .iter()
        .any(|node| ["extLst", "externalReferences", "pivotCaches"].contains(&node.name.as_str()))
    {
        return Err(crate::coded(
            "name-unsupported",
            "Extended workbook references require qualification",
        ));
    }
    let empty = serde_json::Map::new();
    let parts = zip.file_names().map(str::to_owned).collect::<Vec<_>>();
    let types = Tree::parse(&part_bytes(zip, "[Content_Types].xml", &empty)?)?;
    types.root(TYPE_NS, "Types")?;
    let mut overrides = BTreeMap::new();
    let mut defaults = BTreeMap::new();
    let mut worksheets = BTreeSet::new();
    let mut worksheet_edges = BTreeMap::new();
    for index in &types.nodes[0].children {
        let node = &types.nodes[*index];
        if node.ns != TYPE_NS || !node.children.is_empty() {
            return Err(crate::coded(
                "name-malformed",
                "Invalid name source content-type graph",
            ));
        }
        let mime = node.attrs.get("ContentType").ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing name source content type")
        })?;
        let lower = mime.to_ascii_lowercase();
        if [
            "table",
            "drawing",
            "chart",
            "pivot",
            "externallink",
            "querytable",
            "connections",
            "activex",
            "vbaproject",
            "signature",
        ]
        .iter()
        .any(|word| lower.contains(word))
        {
            return Err(crate::coded(
                "name-unsupported",
                "Referenced names require qualified object/reference rewriting",
            ));
        }
        if node.name == "Override" {
            let part = node
                .attrs
                .get("PartName")
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing type override"))?
                .trim_start_matches('/')
                .to_string();
            if overrides.insert(part, mime.clone()).is_some() {
                return Err(crate::coded(
                    "name-duplicate",
                    "Duplicate name source content type",
                ));
            }
        } else if node.name == "Default" {
            let extension = node
                .attrs
                .get("Extension")
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing type extension"))?
                .to_ascii_lowercase();
            if defaults.insert(extension, mime.clone()).is_some() {
                return Err(crate::coded(
                    "name-duplicate",
                    "Duplicate name source content type",
                ));
            }
        } else {
            return Err(crate::coded(
                "internal-unknown-reference",
                "Unknown name source content-type entry",
            ));
        }
    }
    let benign = |kind: &str| {
        let suffix = kind
            .strip_prefix("http://schemas.openxmlformats.org/officeDocument/2006/relationships/");
        suffix.is_some_and(|value| {
            [
                "officeDocument",
                "worksheet",
                "styles",
                "sharedStrings",
                "theme",
                "metadata",
                "printerSettings",
                "extended-properties",
            ]
            .contains(&value)
        }) || [
            "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
            "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail",
        ]
        .contains(&kind)
    };
    for name in &parts {
        crate::cancel::check()?;
        if [
            "xl/charts/",
            "xl/tables/",
            "xl/pivot",
            "xl/drawings/",
            "customXml/",
            "xl/externalLinks/",
        ]
        .iter()
        .any(|prefix| name.starts_with(prefix))
        {
            return Err(crate::coded(
                "name-unsupported",
                "Referenced names require qualified object/reference rewriting",
            ));
        }
        let mime = overrides
            .get(name)
            .or_else(|| defaults.get(&name.rsplit('.').next().unwrap_or("").to_ascii_lowercase()));
        if mime.is_some_and(|value| {
            value.eq_ignore_ascii_case(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            )
        }) || name.starts_with("xl/worksheets/")
            && name.ends_with(".xml")
            && !name.contains("/_rels/")
        {
            worksheets.insert(name.clone());
        }
        if name.ends_with(".rels") {
            let tree = Tree::parse(&part_bytes(zip, name, &empty)?)?;
            tree.root(PACKAGE_REL_NS, "Relationships")?;
            let owner = rel_owner(name)?;
            let mut ids = BTreeSet::new();
            for index in &tree.nodes[0].children {
                let node = &tree.nodes[*index];
                let kind = node.attrs.get("Type").ok_or_else(|| {
                    crate::coded(
                        "internal-missing-field",
                        "Missing name source relationship type",
                    )
                })?;
                if node.ns != PACKAGE_REL_NS
                    || node.name != "Relationship"
                    || !node.children.is_empty()
                    || !benign(kind)
                    || node
                        .attrs
                        .get("TargetMode")
                        .is_some_and(|mode| mode == "External")
                {
                    return Err(crate::coded(
                        "name-unsupported",
                        "Referenced names require qualified relationship rewriting",
                    ));
                }
                if !ids.insert(node.attrs.get("Id").ok_or_else(|| {
                    crate::coded(
                        "internal-missing-field",
                        "Missing name source relationship identity",
                    )
                })?) {
                    return Err(crate::coded(
                        "name-duplicate",
                        "Duplicate name source relationship",
                    ));
                }
                if kind.ends_with("/worksheet") {
                    let target = resolve(
                        &owner,
                        node.attrs.get("Target").ok_or_else(|| {
                            crate::coded(
                                "internal-missing-field",
                                "Missing worksheet relationship target",
                            )
                        })?,
                    )?;
                    if !parts.contains(&target) {
                        return Err(crate::coded(
                            "internal-missing-field",
                            "Missing actual worksheet part",
                        ));
                    }
                    if owner == "xl/workbook.xml" {
                        worksheet_edges.insert(node.attrs["Id"].clone(), target.clone());
                    }
                    worksheets.insert(target);
                }
            }
        }
    }

    let declared = workbook.direct(0, "sheets")?.ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing name-consumer worksheet bindings",
        )
    })?;
    let mut owners = BTreeMap::new();
    for index in &workbook.nodes[declared].children {
        let node = &workbook.nodes[*index];
        let id = node.attrs.get("r:id").ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing worksheet relationship identity",
            )
        })?;
        let part = worksheet_edges.get(id).ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing actual name-consumer worksheet",
            )
        })?;
        if owners
            .insert(part.clone(), node.attrs["name"].clone())
            .is_some()
        {
            return Err(crate::coded(
                "name-ambiguous",
                "Ambiguous name-consumer worksheet binding",
            ));
        }
    }
    let scope =
        if plan["scope"].is_null() {
            None
        } else {
            Some(
                sheets(workbook)?
                    .get(plan["scope"].as_u64().ok_or_else(|| {
                        crate::coded("internal-invalid-field", "Invalid name scope")
                    })? as usize)
                    .cloned()
                    .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid name scope"))?,
            )
        };
    let target = field(plan, "name")?;
    let name_records = names(workbook)?;
    let mut formula_count = 0usize;
    for part in worksheets {
        let owner = owners.get(&part).ok_or_else(|| {
            crate::coded(
                "name-unsupported",
                "Unbound worksheet part requires name-consumer qualification",
            )
        })?;
        let owner_scope = sheets(workbook)?
            .iter()
            .position(|n| n == owner)
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing actual sheet scope"))?
            as u64;
        let shadow = name_records.iter().any(|i| {
            workbook.nodes[*i]
                .attrs
                .get("name")
                .is_some_and(|n| n.eq_ignore_ascii_case(target))
                && workbook.nodes[*i]
                    .attrs
                    .get("localSheetId")
                    .and_then(|s| s.parse::<u64>().ok())
                    == Some(owner_scope)
        });
        let bytes = part_bytes(zip, &part, &empty)?;
        xml_root(&bytes, NS, b"worksheet")?;
        xml_events(&bytes, |_, local, _| {
            if [
                b"formula".as_slice(),
                b"formula1",
                b"formula2",
                b"dataValidation",
                b"dataValidations",
                b"conditionalFormatting",
                b"tableParts",
                b"drawing",
                b"legacyDrawing",
                b"legacyDrawingHF",
                b"oleObjects",
                b"controls",
                b"extLst",
                b"hyperlinks",
            ]
            .contains(&local)
            {
                return Err(crate::coded(
                    "name-unsupported",
                    "Referenced names require qualified formula/reference rewriting",
                ));
            }
            Ok(())
        })?;
        let mut reader = NsReader::from_reader(bytes.as_slice());
        let mut buffer = Vec::new();
        let mut formula: Option<String> = None;
        loop {
            crate::cancel::check()?;
            let (binding, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
            let ns = match binding {
                quick_xml::name::ResolveResult::Bound(ns) => ns.as_ref().to_vec(),
                _ => Vec::new(),
            };
            match event {
                Event::Start(ref e) | Event::Empty(ref e) if e.local_name().as_ref() == b"f" => {
                    if ns != NS || formula.is_some() || matches!(event, Event::Empty(_)) {
                        return Err(crate::coded(
                            "name-unsupported",
                            "Shared/array/extended name consumers require qualification",
                        ));
                    }
                    // Our accepted exporter materializes namespace declarations
                    // on f. They do not change ordinary formula semantics.
                    for attribute in e.attributes() {
                        let attribute = attribute.map_err(err)?;
                        let key = attribute.key.as_ref();
                        if key == b"xmlns" || key.starts_with(b"xmlns:") {
                            continue;
                        }
                        if key == b"t"
                            && attribute.unescape_value().map_err(err)?.as_ref() == "normal"
                        {
                            continue;
                        }
                        return Err(crate::coded(
                            "name-unsupported",
                            "Shared/array/extended name consumers require qualification",
                        ));
                    }
                    formula = Some(String::new());
                }
                Event::Start(_) | Event::Empty(_) if formula.is_some() => {
                    return Err(crate::coded(
                        "name-unsupported",
                        "Nested name-consumer formula",
                    ));
                }
                Event::Text(ref text) if formula.is_some() => {
                    formula
                        .as_mut()
                        .unwrap()
                        .push_str(&text.unescape().map_err(err)?);
                }
                Event::CData(ref text) if formula.is_some() => {
                    formula
                        .as_mut()
                        .unwrap()
                        .push_str(std::str::from_utf8(text).map_err(err)?);
                }
                Event::End(ref e) if e.local_name().as_ref() == b"f" => {
                    let text = formula.take().ok_or_else(|| {
                        crate::coded("name-malformed", "Unbalanced name-consumer formula")
                    })?;
                    formula_count += 1;
                    if formula_count > 100000 || text.len() > 32768 {
                        return Err(crate::coded(
                            "name-limit",
                            "Name-consumer scan exceeds its bound",
                        ));
                    }
                    let used = crate::formula::static_name_use(
                        &text,
                        target,
                        owner,
                        scope.as_deref(),
                        shadow,
                    )?;
                    if plan["kind"] == "removeName" && used {
                        return Err(crate::coded(
                            "name-unsupported",
                            "Cannot remove a referenced name; change its formula consumers first",
                        ));
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            if formula.as_ref().is_some_and(|s| s.len() > 32768) {
                return Err(crate::coded(
                    "name-limit",
                    "Name-consumer formula exceeds its bound",
                ));
            }
            buffer.clear();
        }
        if formula.is_some() {
            return Err(crate::coded(
                "name-unsupported",
                "Incomplete name-consumer formula",
            ));
        }
    }
    Ok(())
}
pub(super) fn validate(root: &Path, plan: &Value, input: &Value) -> Result<()> {
    if plan.as_object().is_none_or(|fields| {
        fields.len() != if plan["kind"] == "removeName" { 8 } else { 9 }
            || fields.keys().any(|key| {
                ![
                    "kind",
                    "name",
                    "scope",
                    "priorFormula",
                    "formula",
                    "replacements",
                    "removed",
                    "copies",
                    "fragments",
                ]
                .contains(&key.as_str())
            })
    }) || plan["replacements"]
        .as_object()
        .is_none_or(|parts| parts.len() != 1 || !parts.contains_key("xl/workbook.xml"))
        || !plan["removed"]
            .as_array()
            .is_some_and(|value| value.is_empty())
        || !plan["copies"]
            .as_object()
            .is_some_and(|value| value.is_empty())
        || !plan["fragments"]
            .as_object()
            .is_some_and(|value| value.is_empty())
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid static name update plan",
        ));
    }
    if !legal_name(field(plan, "name")?) {
        return Err(crate::coded(
            "name-unsupported",
            "Reserved or invalid static range name",
        ));
    }
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, field(input, "source")?)?).map_err(err)?,
    ))
    .map_err(err)?;
    let empty = serde_json::Map::new();
    let before = Tree::parse(&part_bytes(&mut zip, "xl/workbook.xml", &empty)?)?;
    let after = Tree::parse(field(&plan["replacements"], "xl/workbook.xml")?.as_bytes())?;
    let old_names = names(&before)?;
    let new_names = names(&after)?;
    let matches = |tree: &Tree, indexes: &Vec<usize>| -> Result<Vec<usize>> {
        let mut found = Vec::new();
        for index in indexes {
            let node = &tree.nodes[*index];
            let scope = node
                .attrs
                .get("localSheetId")
                .map(|v| v.parse::<u64>().map_err(err))
                .transpose()?;
            if node.attrs.get("name").map(String::as_str) == Some(field(plan, "name")?)
                && (plan["scope"].is_null() && scope.is_none()
                    || plan["scope"].as_u64().is_some() && plan["scope"].as_u64() == scope)
            {
                found.push(*index)
            }
        }
        Ok(found)
    };
    let old_matches = matches(&before, &old_names)?;
    if old_matches.len() != 1 {
        return Err(crate::coded(
            "name-refused",
            "Static name prior identity conflict",
        ));
    }
    let old = old_matches[0];
    if plan["kind"] == "removeName" {
        if plan.get("formula").is_some()
            || before.nodes[old].text != field(plan, "priorFormula")?
            || old_names.len() != new_names.len() + 1
            || before.mask(&[old])? != after.bytes
        {
            return Err(crate::coded(
                "name-stale",
                "Static name removal changed prior identity or unrelated bytes",
            ));
        }
        references(&mut zip, &before, plan)?;
        return Ok(());
    }
    if old_names.len() != new_names.len() {
        return Err(crate::coded(
            "name-stale",
            "Static name update changed the name inventory",
        ));
    }
    let new_matches = matches(&after, &new_names)?;
    if new_matches.len() != 1 {
        return Err(crate::coded(
            "name-refused",
            "Static name next identity conflict",
        ));
    }
    let new = new_matches[0];
    for (tree, index, formula) in [(&before, old, "priorFormula"), (&after, new, "formula")] {
        let node = &tree.nodes[index];
        if node
            .attrs
            .keys()
            .any(|key| !["name", "localSheetId"].contains(&key.as_str()))
            || node.attrs.get("name").map(String::as_str) != Some(field(plan, "name")?)
            || node.text != field(plan, formula)?
        {
            return Err(crate::coded(
                "name-refused",
                "Static name prior identity/formula conflict",
            ));
        }
        let scope = node
            .attrs
            .get("localSheetId")
            .map(|value| value.parse::<u64>().map_err(err))
            .transpose()?;
        let actual_sheets = sheets(tree)?;
        if scope.is_some_and(|value| value >= actual_sheets.len() as u64) {
            return Err(crate::coded(
                "name-unsupported",
                "Static name scope is outside actual workbook sheets",
            ));
        }
        if plan["scope"].is_null() && scope.is_some()
            || !plan["scope"].is_null() && plan["scope"].as_u64() != scope
        {
            return Err(crate::coded("name-stale", "Static name scope changed"));
        }
        static_range(&node.text, tree)?;
    }
    if before.nodes[old].attrs != after.nodes[new].attrs
        || envelope(&before, old)? != envelope(&after, new)?
    {
        return Err(crate::coded(
            "name-stale",
            "Static name update changed unrelated workbook bytes",
        ));
    }
    references(&mut zip, &before, plan)?;
    Ok(())
}
