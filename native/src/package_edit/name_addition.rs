//! One static name addition, checked against source OPC bindings and raw XML.
use super::print_settings::Tree;
use super::*;

pub(super) fn legal_name(name: &str) -> bool {
    let bytes = name.as_bytes();
    if bytes.is_empty()
        || bytes.len() > 255
        || !(bytes[0].is_ascii_alphabetic() || bytes[0] == b'_')
        || bytes
            .iter()
            .any(|b| !b.is_ascii_alphanumeric() && ![b'_', b'.'].contains(b))
    {
        return false;
    }
    let upper = name.to_ascii_uppercase();
    if upper.starts_with("_XLNM.") || ["R", "C", "TRUE", "FALSE"].contains(&upper.as_str()) {
        return false;
    }
    let letters = bytes.iter().take_while(|b| b.is_ascii_alphabetic()).count();
    if (1..=3).contains(&letters)
        && letters < bytes.len()
        && bytes[letters..].iter().all(u8::is_ascii_digit)
    {
        return false;
    }
    if let Some(rest) = upper.strip_prefix('R') {
        if let Some((row, col)) = rest.split_once('C') {
            if !row.is_empty()
                && !col.is_empty()
                && row.bytes().all(|b| b.is_ascii_digit())
                && col.bytes().all(|b| b.is_ascii_digit())
            {
                return false;
            }
        }
    }
    true
}
fn escape(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}
fn container(tree: &Tree) -> Result<Option<usize>> {
    tree.root(NS, "workbook")?;
    let found = tree.nodes[0]
        .children
        .iter()
        .copied()
        .filter(|i| tree.nodes[*i].name == "definedNames")
        .collect::<Vec<_>>();
    if found.len() > 1 || found.first().is_some_and(|i| tree.nodes[*i].ns != NS) {
        return Err(crate::coded("name-ambiguous", "Ambiguous name container"));
    }
    if let Some(index) = found.first() {
        let node = &tree.nodes[*index];
        if !node.text.trim().is_empty()
            || node
                .attrs
                .keys()
                .any(|k| k != "xmlns" && !k.starts_with("xmlns:"))
            || node.children.iter().any(|i| {
                tree.nodes[*i].ns != NS
                    || tree.nodes[*i].name != "definedName"
                    || !tree.nodes[*i].children.is_empty()
            })
        {
            return Err(crate::coded(
                "name-protected",
                "Extended name container is protected",
            ));
        }
    }
    Ok(found.first().copied())
}
fn scope(tree: &Tree, index: usize, sheets: usize) -> Result<Option<u32>> {
    let Some(text) = tree.nodes[index].attrs.get("localSheetId") else {
        return Ok(None);
    };
    if text.is_empty() || !text.bytes().all(|b| b.is_ascii_digit()) {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid existing name scope",
        ));
    }
    let value = text.parse::<u32>().map_err(err)?;
    if value as usize >= sheets {
        return Err(crate::coded(
            "name-unsupported",
            "Existing name scope outside actual workbook",
        ));
    }
    Ok(Some(value))
}
fn expected(tree: &Tree, parent: Option<usize>, addition: &str) -> Result<Vec<u8>> {
    let mut bytes = tree.bytes.clone();
    if let Some(index) = parent {
        let node = &tree.nodes[index];
        let raw = tree.raw(index);
        let mut at = raw
            .len()
            .checked_sub(2)
            .ok_or_else(|| crate::coded("name-malformed", "Invalid name container boundary"))?;
        while raw[at].is_ascii_whitespace() {
            at = at
                .checked_sub(1)
                .ok_or_else(|| crate::coded("name-malformed", "Invalid name boundary"))?;
        }
        if raw[at] == b'/' {
            bytes.remove(node.start + at);
            let insert = node.end - 1;
            bytes.splice(
                insert..insert,
                format!("{addition}</{}>", node.qname).bytes(),
            );
        } else {
            let close = format!("</{}>", node.qname);
            if !raw.ends_with(close.as_bytes()) {
                return Err(crate::coded(
                    "name-malformed",
                    "Invalid name container closing boundary",
                ));
            }
            let insert = node.end - close.len();
            bytes.splice(insert..insert, addition.bytes());
        }
    } else {
        let prefix = tree.nodes[0]
            .qname
            .rsplit_once(':')
            .map(|(prefix, _)| format!("{prefix}:"))
            .unwrap_or_default();
        let rank = [
            "calcPr",
            "oleSize",
            "customWorkbookViews",
            "pivotCaches",
            "smartTagPr",
            "smartTagTypes",
            "webPublishing",
            "fileRecoveryPr",
            "webPublishObjects",
            "extLst",
        ];
        let following = tree.nodes[0]
            .children
            .iter()
            .find(|i| tree.nodes[**i].ns == NS && rank.contains(&tree.nodes[**i].name.as_str()));
        let close = format!("</{}>", tree.nodes[0].qname);
        let insert = if let Some(index) = following {
            tree.nodes[*index].start
        } else {
            let at = tree.nodes[0].end.checked_sub(close.len()).ok_or_else(|| {
                crate::coded(
                    "internal-missing-field",
                    "Missing workbook closing boundary",
                )
            })?;
            if bytes.get(at..tree.nodes[0].end) != Some(close.as_bytes()) {
                return Err(crate::coded(
                    "internal-missing-field",
                    "Missing workbook closing boundary",
                ));
            }
            at
        };
        bytes.splice(
            insert..insert,
            format!("<{prefix}definedNames>{addition}</{prefix}definedNames>").bytes(),
        );
    }
    Ok(bytes)
}
pub(super) fn validate(root: &Path, plan: &Value, input: &Value) -> Result<()> {
    let fields = plan
        .as_object()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid name addition plan"))?;
    if fields.len() != 8
        || fields.keys().any(|key| {
            ![
                "kind",
                "name",
                "scope",
                "formula",
                "replacements",
                "removed",
                "copies",
                "fragments",
            ]
            .contains(&key.as_str())
        })
        || plan["kind"] != "defineName"
        || plan["replacements"]
            .as_object()
            .is_none_or(|r| r.len() != 1 || !r.contains_key("xl/workbook.xml"))
        || plan["removed"].as_array().is_none_or(|v| !v.is_empty())
        || plan["copies"].as_object().is_none_or(|v| !v.is_empty())
        || plan["fragments"].as_object().is_none_or(|v| !v.is_empty())
    {
        return Err(crate::coded(
            "name-refused",
            "Closed name addition delta required",
        ));
    }
    let name = field(plan, "name")?;
    if !legal_name(name) {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid or reserved name addition",
        ));
    }
    let version = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing name addition version"))?;
    let db = crate::database(root)?;
    crate::calculation::snapshot(&db, version)?;
    let index: Value =
        serde_json::from_slice(&fs::read(local(root, "index.json")?).map_err(err)?).map_err(err)?;
    if crate::hash(&local(root, "source.xlsx")?)? != field(&index, "revision")? {
        return Err(crate::coded(
            "name-stale",
            "Name addition source changed since indexing",
        ));
    }
    let source = field(input, "source")?;
    if !["source.xlsx", "export.xlsx"].contains(&source) {
        return Err(crate::coded(
            "name-bad-input",
            "Name addition requires an accepted source export",
        ));
    }
    if source == "export.xlsx" {
        let exported: u64 = db
            .prepare_cached("SELECT version FROM export_state")
            .map_err(crate::err)?
            .query_row([], |row| row.get(0))
            .map_err(err)?;
        if exported != version {
            return Err(crate::coded(
                "name-unsupported",
                "Name addition export belongs to a different version",
            ));
        }
    }
    let bindings = auditing_metadata(root)?;
    let sheets = bindings["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing actual name worksheets"))?;
    let new_scope = if plan["scope"].is_null() {
        None
    } else {
        let scope = plan["scope"]
            .as_u64()
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid name scope"))?;
        if scope >= sheets.len() as u64 {
            return Err(crate::coded(
                "name-unsupported",
                "Name scope outside actual workbook",
            ));
        }
        Some(scope as u32)
    };
    let formula = field(plan, "formula")?;
    let (quoted, range) = formula.rsplit_once('!').ok_or_else(|| {
        crate::coded(
            "name-bad-input",
            "Static name needs an actual sheet rectangle",
        )
    })?;
    let sheet = quoted
        .strip_prefix('\'')
        .and_then(|text| text.strip_suffix('\''))
        .ok_or_else(|| crate::coded("name-bad-input", "Static name needs quoted sheet identity"))?
        .replace("''", "'");
    if quoted != format!("'{}'", sheet.replace('\'', "''"))
        || !sheets.iter().any(|item| item["name"] == sheet)
    {
        return Err(crate::coded(
            "name-bad-input",
            "Static name target is not an actual worksheet",
        ));
    }
    let pieces = range.split(':').collect::<Vec<_>>();
    if pieces.len() != 2 {
        return Err(crate::coded(
            "name-bad-input",
            "Static name requires a canonical absolute rectangle",
        ));
    }
    let (first, last) = crate::bulk::bounds(&range.replace('$', ""))?;
    let absolute = |point: (u32, u32)| {
        let reference = crate::xml::address(point.0, point.1);
        let at = reference.find(|ch: char| ch.is_ascii_digit()).unwrap();
        format!("${}${}", &reference[..at], &reference[at..])
    };
    if range != format!("{}:{}", absolute(first), absolute(last)) {
        return Err(crate::coded(
            "name-bad-input",
            "Static name requires canonical absolute endpoints",
        ));
    }
    if bindings["tables"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing actual table names"))?
        .iter()
        .any(|table| {
            table["name"]
                .as_str()
                .is_some_and(|table| table.eq_ignore_ascii_case(name))
        })
    {
        return Err(crate::coded(
            "name-conflict",
            "Name addition collides with a source table",
        ));
    }
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, source)?).map_err(err)?,
    ))
    .map_err(err)?;
    let before = Tree::parse(&part_bytes(
        &mut zip,
        "xl/workbook.xml",
        &serde_json::Map::new(),
    )?)?;
    let parent = container(&before)?;
    if source == "export.xlsx" {
        validate_calculation_metadata(
            root,
            &db,
            version,
            &json!({"replacements":{"xl/workbook.xml":std::str::from_utf8(&before.bytes).map_err(err)?}}),
        )?;
    } else {
        let mut metadata = ZipArchive::new(crate::cancel::Reader(
            File::open(local(root, "metadata.xlsx")?).map_err(err)?,
        ))
        .map_err(err)?;
        if part_bytes(&mut metadata, "xl/workbook.xml", &serde_json::Map::new())? != before.bytes {
            return Err(crate::coded(
                "name-bad-input",
                "Name addition metadata does not describe indexed workbook source",
            ));
        }
    }
    let existing = parent
        .map(|index| before.nodes[index].children.as_slice())
        .unwrap_or_default();
    if existing.len() >= 1000 {
        return Err(crate::coded("name-limit", "Name limit reached"));
    }
    for index in existing {
        if before.nodes[*index]
            .attrs
            .get("name")
            .is_some_and(|prior| prior.eq_ignore_ascii_case(name))
        {
            if scope(&before, *index, sheets.len())? == new_scope {
                return Err(crate::coded(
                    "name-conflict",
                    "Name addition collides in its source scope",
                ));
            }
        }
    }
    let prefix = parent
        .map(|index| &before.nodes[index].qname)
        .unwrap_or(&before.nodes[0].qname)
        .rsplit_once(':')
        .map(|(prefix, _)| format!("{prefix}:"))
        .unwrap_or_default();
    let local = new_scope
        .map(|scope| format!(" localSheetId=\"{scope}\""))
        .unwrap_or_default();
    let addition = format!(
        "<{prefix}definedName name=\"{name}\"{local}>{}</{prefix}definedName>",
        escape(formula)
    );
    let after = Tree::parse(field(&plan["replacements"], "xl/workbook.xml")?.as_bytes())?;
    let after_parent = container(&after)?
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing new name container"))?;
    if after.nodes[after_parent].children.len() != existing.len() + 1
        || after.bytes != expected(&before, parent, &addition)?
    {
        return Err(crate::coded(
            "name-stale",
            "Name addition changed unrelated workbook bytes or existing names",
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn ascii_name_profile_is_closed() {
        for name in ["Revenue", "_Input", "Cost.total", "SS"] {
            assert!(legal_name(name));
        }
        for name in [
            "",
            "TRUE",
            "false",
            "R",
            "c",
            "A1",
            "R12C7",
            "_xlnm.Print_Area",
            "has space",
            "Ω",
            "\\Name",
        ] {
            assert!(!legal_name(name), "{name}");
        }
        assert!(legal_name(&"a".repeat(255)));
        assert!(!legal_name(&"a".repeat(256)));
    }
    #[test]
    fn raw_addition_preserves_prefixed_empty_container_bytes() {
        let bytes = format!(
            "<s:workbook xmlns:s=\"{}\"><s:sheets/><s:definedNames /><!-- tail --><s:calcPr calcMode=\"manual\"/></s:workbook>",
            String::from_utf8_lossy(NS)
        );
        let tree = Tree::parse(bytes.as_bytes()).unwrap();
        let parent = container(&tree).unwrap();
        let addition = "<s:definedName name=\"Revenue\">'Sheet1'!$B$2:$B$4</s:definedName>";
        let expected_text = bytes.replace(
            "<s:definedNames />",
            &format!("<s:definedNames >{addition}</s:definedNames>"),
        );
        assert_eq!(
            expected(&tree, parent, addition).unwrap(),
            expected_text.as_bytes()
        );
    }
    #[test]
    fn raw_addition_preserves_existing_comments_and_name_attributes() {
        let bytes = format!(
            "<workbook xmlns=\"{}\"><sheets/><definedNames><!-- keep --><definedName name='Prior' hidden='1'>Sheet1!A1</definedName>  </definedNames><calcPr/></workbook>",
            String::from_utf8_lossy(NS)
        );
        let tree = Tree::parse(bytes.as_bytes()).unwrap();
        let addition = "<definedName name=\"Next\">'Sheet1'!$B$2:$B$4</definedName>";
        assert_eq!(
            expected(&tree, container(&tree).unwrap(), addition).unwrap(),
            bytes
                .replace("</definedNames>", &format!("{addition}</definedNames>"))
                .as_bytes()
        );
    }
}
