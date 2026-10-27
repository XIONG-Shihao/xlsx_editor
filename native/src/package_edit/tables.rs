//! Ordinary table symbol/row-boundary edits, independently admitted in native.
use super::print_settings::Tree;
use super::*;
use rusqlite::{Connection, params};
pub(super) mod columns;
fn attributes(tree: &Tree, index: usize, allowed: &[&str]) -> Result<()> {
    let n = &tree.nodes[index];
    if n.attrs
        .keys()
        .any(|key| key != "xmlns" && !key.starts_with("xmlns:") && !allowed.contains(&key.as_str()))
        || !n.text.trim().is_empty()
    {
        return Err(crate::coded(
            "table-protected",
            "Extended table attributes/text are protected",
        ));
    }
    Ok(())
}
fn leaf(tree: &Tree, index: usize, allowed: &[&str]) -> Result<()> {
    attributes(tree, index, allowed)?;
    if !tree.nodes[index].children.is_empty() {
        return Err(crate::coded(
            "table-unsupported",
            "Nested table metadata requires qualification",
        ));
    }
    Ok(())
}
fn table_name(value: &str) -> bool {
    let upper = value.to_ascii_uppercase();
    !value.is_empty()
        && value.len() <= 255
        && value.bytes().enumerate().all(|(i, c)| {
            if i == 0 {
                c.is_ascii_alphabetic() || c == b'_'
            } else {
                c.is_ascii_alphanumeric() || c == b'_' || c == b'.'
            }
        })
        && !upper.starts_with("_XLNM.")
        && !matches!(upper.as_str(), "R" | "C")
        && !position(value).is_ok()
        && !(upper.starts_with('R')
            && upper[1..].split_once('C').is_some_and(|(a, b)| {
                !a.is_empty()
                    && !b.is_empty()
                    && a.bytes().chain(b.bytes()).all(|c| c.is_ascii_digit())
            }))
}
fn integer(value: &str) -> Result<u64> {
    if value.is_empty()
        || !value.bytes().all(|c| c.is_ascii_digit())
        || value.len() > 1 && value.starts_with('0')
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Noncanonical table integer",
        ));
    }
    value
        .parse::<u64>()
        .ok()
        .filter(|v| *v <= 9007199254740991)
        .ok_or_else(|| crate::coded("table-limit", "Table integer exceeds its bound"))
}
fn attr<'a>(tree: &'a Tree, index: usize, key: &str) -> Result<&'a str> {
    tree.nodes[index]
        .attrs
        .get(key)
        .map(String::as_str)
        .ok_or_else(|| crate::coded("internal-missing-field", format!("Missing table {key}")))
}
fn direct(tree: &Tree, index: usize, name: &str, namespace: &[u8]) -> Result<Vec<usize>> {
    let nodes = tree.nodes[index]
        .children
        .iter()
        .copied()
        .filter(|i| tree.nodes[*i].name == name)
        .collect::<Vec<_>>();
    if nodes.iter().any(|i| tree.nodes[*i].ns != namespace) {
        return Err(crate::coded("table-unsupported", "Foreign table namespace"));
    }
    Ok(nodes)
}
fn target(owner: &str, value: &str) -> Result<String> {
    let mut path = if value.starts_with('/') {
        Vec::new()
    } else {
        owner
            .rsplit_once('/')
            .map(|(dir, _)| dir.split('/').collect::<Vec<_>>())
            .unwrap_or_default()
    };
    for item in value.trim_start_matches('/').split('/') {
        match item {
            ".." => {
                path.pop().ok_or_else(|| {
                    crate::coded("table-unsupported", "Table relationship escapes package")
                })?;
            }
            "." | "" => {
                return Err(crate::coded(
                    "table-malformed",
                    "Noncanonical table relationship target",
                ));
            }
            _ => path.push(item),
        }
    }
    let result = path.join("/");
    valid_part(&result)?;
    Ok(result)
}
fn relationships(bytes: &[u8]) -> Result<BTreeMap<String, (String, String)>> {
    let tree = Tree::parse(bytes)?;
    tree.root(PACKAGE_REL_NS, "Relationships")?;
    attributes(&tree, 0, &[])?;
    let mut result = BTreeMap::new();
    for i in &tree.nodes[0].children {
        let n = &tree.nodes[*i];
        if n.ns != PACKAGE_REL_NS || n.name != "Relationship" {
            return Err(crate::coded(
                "table-malformed",
                "Invalid table relationship graph",
            ));
        }
        leaf(&tree, *i, &["Id", "Type", "Target", "TargetMode"])?;
        let id = attr(&tree, *i, "Id")?;
        let kind = attr(&tree, *i, "Type")?;
        let dest = attr(&tree, *i, "Target")?;
        if id.is_empty()
            || dest.is_empty()
            || n.attrs.get("TargetMode").is_some_and(|v| v != "Internal")
            || result
                .insert(id.into(), (kind.into(), dest.into()))
                .is_some()
        {
            return Err(crate::coded(
                "table-ambiguous",
                "Ambiguous/external table relationship",
            ));
        }
    }
    Ok(result)
}
fn ordinary_cell_attributes(
    root: &Path,
    db: &Connection,
    sheet: &str,
    at: (u32, u32),
) -> Result<()> {
    if crate::spill::source_attributes(root, db, sheet, at)?.is_some_and(|attrs| {
        attrs
            .keys()
            .any(|key| !["r", "s", "t", "xmlns"].contains(&key.as_str()))
            || attrs.get("xmlns").is_some_and(|ns| ns.as_bytes() != NS)
    }) {
        return Err(crate::coded(
            "table-unsupported",
            "Opaque table header/growth cell attributes require qualification",
        ));
    }
    crate::spill::source_content_proof(root, db, sheet, at, true)?;
    Ok(())
}
#[derive(Clone)]
struct Model {
    name: String,
    reference: String,
    columns: Vec<String>,
}
fn model(tree: &Tree) -> Result<Model> {
    tree.root(NS, "table")?;
    attributes(
        tree,
        0,
        &[
            "id",
            "name",
            "displayName",
            "ref",
            "headerRowCount",
            "totalsRowCount",
            "totalsRowShown",
        ],
    )?;
    if integer(attr(tree, 0, "id")?)? == 0 {
        return Err(crate::coded("table-bad-input", "Table id must be positive"));
    }
    let name = attr(tree, 0, "name")?;
    if !table_name(name) || attr(tree, 0, "displayName")? != name {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid/ambiguous table name",
        ));
    }
    let reference = attr(tree, 0, "ref")?;
    let (first, last) = crate::bulk::bounds(reference)?;
    let count = u64::from(last.0 - first.0 + 1) * u64::from(last.1 - first.1 + 1);
    if count > 10000
        || last.1 - first.1 + 1 > 1000
        || first.0 == last.0
        || tree.nodes[0]
            .attrs
            .get("headerRowCount")
            .map(|v| integer(v))
            .transpose()?
            .unwrap_or(1)
            != 1
        || tree.nodes[0]
            .attrs
            .get("totalsRowCount")
            .map(|v| integer(v))
            .transpose()?
            .unwrap_or(0)
            != 0
        || tree.nodes[0]
            .attrs
            .get("totalsRowShown")
            .is_some_and(|v| !["0", "1", "false", "true"].contains(&v.as_str()))
    {
        return Err(crate::coded(
            "table-unsupported",
            "Only bounded ordinary one-header/no-total tables are qualified",
        ));
    }
    let mut seen = BTreeSet::new();
    let mut columns = Vec::new();
    let mut ids = BTreeSet::new();
    let mut labels = BTreeSet::new();
    for index in &tree.nodes[0].children {
        let n = &tree.nodes[*index];
        if n.ns != NS || !seen.insert(n.name.clone()) {
            return Err(crate::coded(
                "table-duplicate",
                "Duplicate/foreign table children",
            ));
        }
        match n.name.as_str() {
            "autoFilter" => {
                leaf(tree, *index, &["ref"])?;
                if attr(tree, *index, "ref")? != reference {
                    return Err(crate::coded(
                        "table-unsupported",
                        "Active or unrelated table filter requires qualification",
                    ));
                }
            }
            "tableColumns" => {
                attributes(tree, *index, &["count"])?;
                if integer(attr(tree, *index, "count")?)? != u64::from(last.1 - first.1 + 1)
                    || n.children.len() != usize::try_from(last.1 - first.1 + 1).map_err(err)?
                {
                    return Err(crate::coded(
                        "internal-mismatch",
                        "Table column count mismatch",
                    ));
                }
                for column in &n.children {
                    let c = &tree.nodes[*column];
                    if c.ns != NS || c.name != "tableColumn" {
                        return Err(crate::coded("table-unsupported", "Foreign table column"));
                    }
                    leaf(tree, *column, &["id", "name"])?;
                    let id = integer(attr(tree, *column, "id")?)?;
                    let label = attr(tree, *column, "name")?;
                    if id == 0
                        || !ids.insert(id)
                        || label.is_empty()
                        || label.len() > 1024
                        || !labels.insert(label.to_ascii_lowercase())
                    {
                        return Err(crate::coded(
                            "table-duplicate",
                            "Duplicate/invalid table columns",
                        ));
                    }
                    columns.push(label.into());
                }
            }
            "tableStyleInfo" => {
                leaf(
                    tree,
                    *index,
                    &[
                        "name",
                        "showFirstColumn",
                        "showLastColumn",
                        "showRowStripes",
                        "showColumnStripes",
                    ],
                )?;
                for key in [
                    "showFirstColumn",
                    "showLastColumn",
                    "showRowStripes",
                    "showColumnStripes",
                ] {
                    if n.attrs
                        .get(key)
                        .is_some_and(|v| !["0", "1", "false", "true"].contains(&v.as_str()))
                    {
                        return Err(crate::coded(
                            "internal-invalid-field",
                            "Invalid table style boolean",
                        ));
                    }
                }
            }
            _ => {
                return Err(crate::coded(
                    "table-unsupported",
                    "Extended table metadata requires qualification",
                ));
            }
        }
    }
    if columns.is_empty() {
        return Err(crate::coded(
            "internal-missing-field",
            "Missing table columns",
        ));
    }
    Ok(Model {
        name: name.into(),
        reference: reference.into(),
        columns,
    })
}
/// Mask only the authorized root attributes and its matching filter ref. Every
/// remaining raw byte (including opaque comments/whitespace) must be identical.
fn masked(tree: &Tree, rename: bool, column: Option<usize>) -> Result<Vec<u8>> {
    let mut ranges = Vec::new();
    let indexes = if let Some(column) = column {
        let container = direct(tree, 0, "tableColumns", NS)?;
        vec![
            *tree.nodes[*container
                .first()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing columns"))?]
            .children
            .get(column)
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing column index"))?,
        ]
    } else if rename {
        vec![0]
    } else {
        std::iter::once(0)
            .chain(direct(tree, 0, "autoFilter", NS)?)
            .collect()
    };
    for index in indexes {
        let node = &tree.nodes[index];
        let bytes = &tree.bytes[node.start..node.end];
        let end = bytes
            .iter()
            .position(|c| *c == b'>')
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid table opening"))?;
        let head = &bytes[..end];
        let mut at = 0;
        while at < head.len() {
            while at < head.len() && head[at] != b'=' {
                at += 1;
            }
            if at == head.len() {
                break;
            }
            let equals = at;
            let mut end = equals;
            while end > 0 && head[end - 1].is_ascii_whitespace() {
                end -= 1;
            }
            let mut start = end;
            while start > 0 && !head[start - 1].is_ascii_whitespace() {
                start -= 1;
            }
            let key = std::str::from_utf8(&head[start..end]).map_err(err)?;
            at = equals + 1;
            while at < head.len() && head[at].is_ascii_whitespace() {
                at += 1;
            }
            let quote = *head
                .get(at)
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing table quote"))?;
            if ![b'\'', b'"'].contains(&quote) {
                return Err(crate::coded("table-malformed", "Unquoted table attribute"));
            }
            let a = at + 1;
            at = a;
            while at < head.len() && head[at] != quote {
                at += 1;
            }
            if at == head.len() {
                return Err(crate::coded(
                    "table-malformed",
                    "Unterminated table attribute",
                ));
            }
            if if column.is_some() {
                key == "name"
            } else if rename {
                ["name", "displayName"].contains(&key)
            } else {
                key == "ref"
            } {
                ranges.push((node.start + a, node.start + at));
            }
            at += 1;
        }
    }
    ranges.sort();
    let mut result = Vec::new();
    let mut at = 0;
    for (start, end) in ranges {
        result.extend_from_slice(&tree.bytes[at..start]);
        result.extend_from_slice(b"@PLATFORM_TABLE_CHANGE@");
        at = end;
    }
    result.extend_from_slice(&tree.bytes[at..]);
    Ok(result)
}
pub(super) fn validate(root: &Path, db: &Connection, plan: &Value, input: &Value) -> Result<()> {
    let kind = field(plan, "kind")?;
    let rename = kind == "renameTable";
    let column_rename = kind == "renameTableColumn";
    let structure = ["insertTableColumn", "deleteTableColumn"].contains(&kind);
    let fields = plan
        .as_object()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid table plan"))?;
    if fields.keys().any(|key| {
        ![
            "kind",
            "sheet",
            "tablePart",
            "tableName",
            "priorRef",
            "priorTableSha256",
            "renameTable",
            "renameTableColumn",
            "columnChange",
            "sheets",
            "range",
            "replacements",
            "removed",
            "copies",
            "fragments",
        ]
        .contains(&key.as_str())
    }) || plan["removed"].as_array().is_none_or(|v| !v.is_empty())
        || plan["copies"].as_object().is_none_or(|v| !v.is_empty())
        || plan["fragments"].as_object().is_none_or(|v| !v.is_empty())
    {
        return Err(crate::coded(
            "table-unsupported",
            "Table plan contains an unqualified delta",
        ));
    }
    if !structure && plan.get("columnChange").is_some() {
        return Err(crate::coded(
            "table-malformed",
            "Unexpected column structure change",
        ));
    }
    let table_part = field(plan, "tablePart")?;
    let sheet = field(plan, "sheet")?;
    let replacements = plan["replacements"]
        .as_object()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing table replacement"))?;
    if replacements.len() != 1 || !replacements.contains_key(table_part) {
        return Err(crate::coded(
            "table-unsupported",
            "Table edits may replace only their existing owned table part",
        ));
    }
    let version = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing table edit version"))?;
    if crate::spill::any_active(db, version)? {
        return Err(crate::coded(
            "table-unsupported",
            "Table edits with active spill ownership require qualification",
        ));
    }
    let index: Value =
        serde_json::from_slice(&fs::read(local(root, "index.json")?).map_err(err)?).map_err(err)?;
    if index["sheets"].as_array().is_none_or(|sheets| {
        sheets.iter().any(|sheet| {
            sheet["extended"] == true || sheet["specialFormulaCount"].as_u64().unwrap_or(0) > 0
        })
    }) {
        return Err(crate::coded(
            "table-unsupported",
            "Special/opaque source table dependencies require qualification",
        ));
    }
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let empty = serde_json::Map::new();
    let workbook = Tree::parse(&part_bytes(&mut zip, "xl/workbook.xml", &empty)?)?;
    workbook.root(NS, "workbook")?;
    if workbook.nodes.iter().any(|n| {
        [
            "definedNames",
            "extLst",
            "externalReferences",
            "pivotCaches",
        ]
        .contains(&n.name.as_str())
    }) {
        return Err(crate::coded(
            "table-unsupported",
            "Name/pivot/extended table dependencies require qualification",
        ));
    }
    let wb_rel = relationships(&part_bytes(&mut zip, "xl/_rels/workbook.xml.rels", &empty)?)?;
    let mut sheets = Vec::new();
    let mut formula_sheets = BTreeMap::new();
    let mut sheet_names = BTreeSet::new();
    for parent in direct(&workbook, 0, "sheets", NS)? {
        for index in &workbook.nodes[parent].children {
            let n = &workbook.nodes[*index];
            if n.ns != NS || n.name != "sheet" || !n.children.is_empty() {
                return Err(crate::coded(
                    "table-unsupported",
                    "Opaque workbook worksheet identity",
                ));
            }
            let name = attr(&workbook, *index, "name")?;
            if !sheet_names.insert(name) {
                return Err(crate::coded(
                    "table-duplicate",
                    "Duplicate table worksheet name",
                ));
            }
            let link = wb_rel
                .get(attr(&workbook, *index, "r:id")?)
                .ok_or_else(|| {
                    crate::coded(
                        "internal-missing-field",
                        "Missing table worksheet relationship",
                    )
                })?;
            if link.0 != format!("{}/worksheet", String::from_utf8_lossy(REL_NS)) {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Table worksheet relationship type is invalid",
                ));
            }
            let path = target("xl/workbook.xml", &link.1)?;
            formula_sheets.insert(path.clone(), name.to_owned());
            sheets.push(path);
        }
    }
    if sheets.is_empty()
        || !sheets.iter().any(|p| p == sheet)
        || sheets.iter().collect::<BTreeSet<_>>().len() != sheets.len()
    {
        return Err(crate::coded(
            "table-refused",
            "Table workbook worksheet graph is ambiguous",
        ));
    }
    let types = Tree::parse(&part_bytes(&mut zip, "[Content_Types].xml", &empty)?)?;
    types.root(TYPE_NS, "Types")?;
    attributes(&types, 0, &[])?;
    let mut overrides = BTreeMap::new();
    let mut defaults = BTreeSet::new();
    for i in &types.nodes[0].children {
        let n = &types.nodes[*i];
        if n.ns != TYPE_NS || !n.children.is_empty() {
            return Err(crate::coded(
                "table-malformed",
                "Malformed table content type graph",
            ));
        }
        let mime = attr(&types, *i, "ContentType")?;
        if [
            "macroenabled",
            "vbaproject",
            "activex",
            "signature",
            "drawing",
            "chart",
            "pivot",
            "externallink",
            "querytable",
            "connections",
        ]
        .iter()
        .any(|word| mime.to_ascii_lowercase().contains(word))
        {
            return Err(crate::coded(
                "table-protected",
                "Object/external/protected table dependency requires qualification",
            ));
        }
        if n.name == "Override" {
            leaf(&types, *i, &["PartName", "ContentType"])?;
            let part = attr(&types, *i, "PartName")?;
            if overrides.insert(part.to_owned(), mime.to_owned()).is_some() {
                return Err(crate::coded(
                    "table-duplicate",
                    "Duplicate table content type",
                ));
            }
        } else if n.name == "Default" {
            leaf(&types, *i, &["Extension", "ContentType"])?;
            let extension = attr(&types, *i, "Extension")?;
            if extension.is_empty() || !defaults.insert(extension.to_ascii_lowercase()) {
                return Err(crate::coded(
                    "table-duplicate",
                    "Duplicate/empty default table content type",
                ));
            }
        } else {
            return Err(crate::coded(
                "internal-unknown-reference",
                "Unknown table content type entry",
            ));
        }
    }
    if overrides.get(&format!("/{table_part}")).map(String::as_str)
        != Some("application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")
    {
        return Err(crate::coded(
            "table-bad-input",
            "Table part has no exact content type",
        ));
    }
    if zip.len() > 10000 {
        return Err(crate::coded(
            "table-limit",
            "Table metadata package part count exceeds its bound",
        ));
    }
    let mut expanded = 0u64;
    for index in 0..zip.len() {
        expanded = expanded
            .checked_add(zip.by_index(index).map_err(err)?.size())
            .ok_or_else(|| crate::coded("table-limit", "Table metadata size overflow"))?;
        if expanded > 16 * 1024 * 1024 {
            return Err(crate::coded(
                "table-limit",
                "Table metadata expansion exceeds16 MiB",
            ));
        }
    }
    let parts = zip.file_names().map(str::to_owned).collect::<Vec<_>>();
    if parts.iter().collect::<BTreeSet<_>>().len() != parts.len() {
        return Err(crate::coded(
            "table-duplicate",
            "Duplicate table package parts",
        ));
    }
    let mut incoming = 0;
    let mut table_names = BTreeSet::new();
    let mut owned = None;
    let mut all_models = Vec::new();
    for part in &parts {
        if part.ends_with(".rels") {
            for (_, (kind, dest)) in relationships(&part_bytes(&mut zip, part, &empty)?)? {
                let suffix = kind.strip_prefix(&format!("{}/", String::from_utf8_lossy(REL_NS)));
                if !suffix.is_some_and(|s| {
                    [
                        "officeDocument",
                        "worksheet",
                        "table",
                        "styles",
                        "sharedStrings",
                        "theme",
                        "sheetMetadata",
                        "extended-properties",
                    ]
                    .contains(&s)
                }) && ![
                    format!(
                        "{}/metadata/core-properties",
                        String::from_utf8_lossy(PACKAGE_REL_NS)
                    ),
                    format!(
                        "{}/metadata/thumbnail",
                        String::from_utf8_lossy(PACKAGE_REL_NS)
                    ),
                ]
                .contains(&kind)
                {
                    return Err(crate::coded(
                        "table-unsupported",
                        "Opaque table relationship type",
                    ));
                }
                let directory = part
                    .rsplit_once("/_rels/")
                    .map(|(dir, file)| {
                        format!("{dir}/{}", file.strip_suffix(".rels").unwrap_or(file))
                    })
                    .unwrap_or_default();
                if kind.ends_with("/table") && target(&directory, &dest)? == table_part {
                    incoming += 1;
                }
            }
        }
    }
    for path in &sheets {
        let tree = Tree::parse(&part_bytes(&mut zip, path, &empty)?)?;
        tree.root(NS, "worksheet")?;
        if tree.nodes.iter().any(|n| {
            [
                "extLst",
                "conditionalFormatting",
                "dataValidations",
                "drawing",
                "legacyDrawing",
                "oleObjects",
                "controls",
            ]
            .contains(&n.name.as_str())
        }) {
            return Err(crate::coded(
                "table-unsupported",
                "Rule/object table dependencies require qualification",
            ));
        }
        let rels = path
            .rsplit_once('/')
            .map(|(dir, file)| format!("{dir}/_rels/{file}.rels"))
            .ok_or_else(|| crate::coded("table-malformed", "Invalid table worksheet part"))?;
        let links = if parts.contains(&rels) {
            relationships(&part_bytes(&mut zip, &rels, &empty)?)?
        } else {
            BTreeMap::new()
        };
        let containers = direct(&tree, 0, "tableParts", NS)?;
        if containers.len() > 1 {
            return Err(crate::coded(
                "table-duplicate",
                "Duplicate tableParts ownership container",
            ));
        }
        for container in containers {
            attributes(&tree, container, &["count"])?;
            if integer(attr(&tree, container, "count")?)?
                != tree.nodes[container].children.len() as u64
            {
                return Err(crate::coded(
                    "internal-mismatch",
                    "Table owner count mismatch",
                ));
            }
            for index in &tree.nodes[container].children {
                let n = &tree.nodes[*index];
                if n.ns != NS || n.name != "tablePart" {
                    return Err(crate::coded("table-unsupported", "Foreign table owner"));
                }
                leaf(&tree, *index, &["r:id"])?;
                let rel = links.get(attr(&tree, *index, "r:id")?).ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing table owner edge")
                })?;
                if rel.0 != format!("{}/table", String::from_utf8_lossy(REL_NS)) {
                    return Err(crate::coded(
                        "table-malformed",
                        "Invalid table owner relationship type",
                    ));
                }
                let part = target(path, &rel.1)?;
                let bytes = part_bytes(&mut zip, &part, &empty)?;
                let tree = Tree::parse(&bytes)?;
                let m = model(&tree)?;
                if !table_names.insert(m.name.to_ascii_lowercase()) {
                    return Err(crate::coded(
                        "table-duplicate",
                        "Duplicate workbook table name",
                    ));
                }
                all_models.push((path.clone(), m.reference.clone()));
                if part == table_part {
                    if owned.is_some() || path != sheet {
                        return Err(crate::coded(
                            "table-refused",
                            "Table part has ambiguous worksheet ownership",
                        ));
                    }
                    owned = Some((bytes, tree, m));
                }
            }
        }
    }
    if incoming != 1 {
        return Err(crate::coded(
            "table-unsupported",
            "Table part is shared or unreferenced",
        ));
    }
    let (bytes, before, old) = owned.ok_or_else(|| {
        crate::coded(
            "table-bad-input",
            "Table part is not owned by its requested worksheet",
        )
    })?;
    if format!("{:x}", Sha256::digest(&bytes)) != field(plan, "priorTableSha256")?
        || old.name != field(plan, "tableName")?
        || old.reference != field(plan, "priorRef")?
    {
        return Err(crate::coded("table-stale", "Stale table definition guard"));
    }
    let after = Tree::parse(
        replacements[table_part]
            .as_str()
            .ok_or_else(|| crate::coded("table-malformed", "Invalid table replacement XML"))?
            .as_bytes(),
    )?;
    let new = model(&after)?;
    let column =
        if column_rename {
            Some(
                usize::try_from(plan["renameTableColumn"]["index"].as_u64().ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing column index")
                })?)
                .map_err(err)?,
            )
        } else {
            None
        };
    if !structure
        && ((!column_rename && old.columns != new.columns)
            || masked(&before, rename, column)? != masked(&after, rename, column)?)
    {
        return Err(crate::coded(
            "table-stale",
            "Table edit changed unrelated raw metadata",
        ));
    }
    let source_bytes = {
        let mut source = ZipArchive::new(crate::cancel::Reader(
            File::open(local(root, field(input, "source")?)?).map_err(err)?,
        ))
        .map_err(err)?;
        part_bytes(&mut source, table_part, &empty)?
    };
    if source_bytes != bytes {
        return Err(crate::coded(
            "internal-mismatch",
            "Prepared export table differs from accepted metadata",
        ));
    }
    crate::calculation::snapshot(db, version)?;
    db.execute_batch("CREATE TEMP TABLE renamed_formulas(sheet TEXT,ref TEXT,formula TEXT,PRIMARY KEY(sheet,ref)) WITHOUT ROWID;").map_err(err)?;
    let movement = if structure {
        let c = columns::definition(&before, &after, &old, &new, plan)?;
        if plan["sheets"]
            .as_array()
            .is_none_or(|v| v.len() != sheets.len() || v.iter().zip(&sheets).any(|(a, b)| a != b))
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid table column sheet scope",
            ));
        }
        let tree = Tree::parse(&part_bytes(&mut zip, sheet, &empty)?)?;
        columns::cells(root, db, plan, input, &c, &all_models, &tree)?;
        Some(c)
    } else {
        None
    };
    if structure {
    } else if column_rename {
        let change = &plan["renameTableColumn"];
        let i = column.ok_or_else(|| crate::coded("internal-missing-field", "Missing column"))?;
        let prior = field(change, "old")?;
        let next = field(change, "new")?;
        let label = |v: &str| {
            !v.trim().is_empty()
                && v.trim() == v
                && v.len() <= 255
                && !v
                    .chars()
                    .any(|c| c.is_control() || ['[', ']', ':', '\'', ',', '#', '@'].contains(&c))
        };
        let mut expected = old.columns.clone();
        if i >= expected.len() {
            return Err(crate::coded(
                "table-unsupported",
                "Column index outside table",
            ));
        }
        expected[i] = next.into();
        if change.as_object().is_none_or(|v| v.len() != 3)
            || !label(prior)
            || !label(next)
            || prior == next
            || old.columns[i] != prior
            || new.columns != expected
            || new.name != old.name
            || new.reference != old.reference
            || plan.get("range").is_some()
            || plan.get("renameTable").is_some()
            || plan["sheets"].as_array().is_none_or(|v| {
                v.len() != sheets.len() || v.iter().zip(&sheets).any(|(a, b)| a != b)
            })
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid table column rename scope or label",
            ));
        }
    } else if rename {
        if plan.get("renameTableColumn").is_some() {
            return Err(crate::coded(
                "table-unsupported",
                "Table rename cannot carry a column change",
            ));
        }
        if plan.get("range").is_some()
            || plan["renameTable"].as_object().is_none_or(|v| v.len() != 2)
            || plan["renameTable"]["old"] != old.name
            || plan["renameTable"]["new"] != new.name
            || new.reference != old.reference
            || new.name == old.name
            || table_names.contains(&new.name.to_ascii_lowercase())
                && !new.name.eq_ignore_ascii_case(&old.name)
            || plan["sheets"].as_array().is_none_or(|v| {
                v.len() != sheets.len() || v.iter().zip(&sheets).any(|(a, b)| a != b)
            })
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid table rename scope or name collision",
            ));
        }
    } else {
        if plan.get("renameTableColumn").is_some()
            || plan.get("renameTable").is_some()
            || plan.get("sheets").is_some()
            || new.name != old.name
            || plan["range"] != new.reference
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid table resize plan",
            ));
        }
        let (a, b) = crate::bulk::bounds(&old.reference)?;
        let (c, d) = crate::bulk::bounds(&new.reference)?;
        if a != c || b.1 != d.1 || b.0 == d.0 {
            return Err(crate::coded(
                "table-bad-input",
                "Table resize must retain header origin and columns",
            ));
        }
        for (path, reference) in &all_models {
            if path == sheet && reference != &old.reference {
                let (e, f) = crate::bulk::bounds(reference)?;
                if c.0 <= f.0 && d.0 >= e.0 && c.1 <= f.1 && d.1 >= e.1 {
                    return Err(crate::coded(
                        "table-conflict",
                        "Table resize intersects another table",
                    ));
                }
            }
        }
        let sheet_tree = Tree::parse(&part_bytes(&mut zip, sheet, &empty)?)?;
        for n in &sheet_tree.nodes {
            if n.ns == NS && n.name == "mergeCell" {
                let (e, f) =
                    crate::bulk::bounds(n.attrs.get("ref").ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing merge ref")
                    })?)?;
                if c.0 <= f.0 && d.0 >= e.0 && c.1 <= f.1 && d.1 >= e.1 {
                    return Err(crate::coded(
                        "table-conflict",
                        "Table resize intersects a merge",
                    ));
                }
            }
        }
        if d.0 > b.0 {
            let rows = (1000 / (d.1 - c.1 + 1)).max(1);
            let mut row = b.0 + 1;
            while row <= d.0 {
                let last = (row + rows - 1).min(d.0);
                let page = crate::range::read_db(
                    db,
                    &json!({"version":version,"sheet":sheet,"range":format!("{}:{}",crate::xml::address(row,c.1),crate::xml::address(last,d.1)),"limit":1001}),
                )?;
                if !page["nextOffset"].is_null()
                    || page["cells"]
                        .as_array()
                        .ok_or_else(|| {
                            crate::coded("internal-missing-field", "Missing growth cells")
                        })?
                        .iter()
                        .any(|record| {
                            let cell = &record[1];
                            cell.get("formula").is_some()
                                || cell.get("formulaType").is_some()
                                || cell.get("spill").is_some()
                                || !cell["value"].is_null()
                        })
                {
                    return Err(crate::coded(
                        "table-unsupported",
                        "Table growth cannot absorb existing content or spill ownership",
                    ));
                }
                for record in page["cells"]
                    .as_array()
                    .ok_or_else(|| crate::coded("internal-missing-field", "Missing growth cells"))?
                {
                    ordinary_cell_attributes(
                        root,
                        db,
                        sheet,
                        position(record[0].as_str().ok_or_else(|| {
                            crate::coded("internal-missing-field", "Missing growth position")
                        })?)?,
                    )?;
                }
                row = last + 1;
            }
        }
    }
    let (a, b) = crate::bulk::bounds(&old.reference)?;
    let headers = crate::range::read_db(
        db,
        &json!({"version":version,"sheet":sheet,"range":format!("{}:{}",crate::xml::address(a.0,a.1),crate::xml::address(a.0,b.1)),"limit":1001}),
    )?;
    let values = headers["cells"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing table header source"))?;
    for record in values {
        ordinary_cell_attributes(
            root,
            db,
            sheet,
            position(record[0].as_str().ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing table header position")
            })?)?,
        )?;
    }
    if !headers["nextOffset"].is_null()
        || values.len() != old.columns.len()
        || values.iter().enumerate().any(|(i, record)| {
            record[1].get("formula").is_some()
                || record[1]["value"].as_str() != Some(old.columns[i].as_str())
        })
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Current literal headers differ from their table definition",
        ));
    }
    if let Some(c) = &movement {
        columns::header_style(db, &headers, c)?;
    }
    let mut count = 0;
    for path in &sheets {
        let mut cursor = None;
        loop {
            let page = crate::calculation::page(db, path, (0, 0), (1048575, 16383), cursor, true)?;
            for (reference, cell) in &page {
                crate::cancel::check()?;
                count += 1;
                if count > 100000
                    || cell.get("formulaType").is_some()
                    || cell.get("scenarioResult").is_some()
                {
                    return Err(crate::coded(
                        "table-unsupported",
                        "Special or excessive formula table dependencies require qualification",
                    ));
                }
                let formula = field(cell, "formula")?;
                if rename || column_rename || structure {
                    let p = position(reference)?;
                    let inside =
                        path == sheet && p.0 > a.0 && p.0 <= b.0 && p.1 >= a.1 && p.1 <= b.1;
                    let updated = if let Some(c) = &movement {
                        columns::formula(formula, &old, plan, c, path, reference, &formula_sheets)?
                    } else if column_rename {
                        crate::formula::rename_table_column(
                            formula,
                            &old.name,
                            field(&plan["renameTableColumn"], "old")?,
                            field(&plan["renameTableColumn"], "new")?,
                            inside,
                        )?
                    } else {
                        crate::formula::rename_table(formula, &old.name, &new.name)?
                    };
                    if updated != formula {
                        db.execute(
                            "INSERT INTO renamed_formulas VALUES(?1,?2,?3)",
                            params![path, reference, updated],
                        )
                        .map_err(err)?;
                    }
                } else {
                    let (a, b) = crate::bulk::bounds(&old.reference)?;
                    let (_, d) = crate::bulk::bounds(&new.reference)?;
                    let p = position(reference)?;
                    if p.0 > d.0 && p.0 <= b.0 && p.1 >= a.1 && p.1 <= b.1 {
                        return Err(crate::coded(
                            "table-unsupported",
                            "Table shrink cannot detach formula-bearing body cells",
                        ));
                    }
                }
            }
            if page.len() < 512 {
                break;
            }
            cursor = Some(position(&page.last().unwrap().0)?);
        }
    }
    Ok(())
}
