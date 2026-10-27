//! Independently qualify a closed real worksheet PivotTable cache and its display.
use super::print_settings::Tree;
use super::*;
use crate::xml::address;
const MAX_SAFE: i64 = 9_007_199_254_740_991;
fn safe_integer(text: &str) -> Result<i64> {
    text.parse::<i64>()
        .map_err(err)?
        .checked_add(0)
        .filter(|number| number.unsigned_abs() <= MAX_SAFE as u64)
        .ok_or_else(|| {
            crate::coded(
                "pivot-limit",
                "Pivot cached numeric value exceeds exact safe integers",
            )
        })
}

#[derive(Clone, Copy)]
struct Parts {
    table: &'static str,
    definition: &'static str,
    records: &'static str,
}
impl Parts {
 fn current()->Self{Self{table:workbook_format!(XL_PIVOTTABLES_PIVOT1_XML),definition:workbook_format!(XL_PIVOTCACHE_CACHEDEFINITION1_XML),records:workbook_format!(XL_PIVOTCACHE_CACHERECORDS1_XML)}}
 fn from_zip(_zip:&ZipArchive<crate::cancel::Reader<File>>)->Result<Self>{Ok(Self::current())}
 fn from_plan(plan:&Value)->Result<Self>{let parts=Self::current();if plan["parts"]!=json!({"table":parts.table,"definition":parts.definition,"records":parts.records}){return Err(crate::coded("pivot-bad-input","Unknown pivot part identities"));}Ok(parts)}
}

#[derive(Clone, Copy)]
struct Presentation {
    sample: &'static str,
    compact: &'static str,
    outline: &'static str,
    style: &'static str,
    stripes: &'static str,
    grand: bool,
}
fn presentation(sample: &str, grand: bool) -> Result<Presentation> {
    let (id, compact, outline, style, stripes) = match sample {
        "tabular-light" => ("tabular-light", "0", "0", "PivotStyleLight16", "0"),
        "compact-blue" => ("compact-blue", "1", "1", "PivotStyleMedium9", "1"),
        "outline-dark" => ("outline-dark", "0", "1", "PivotStyleDark2", "1"),
        _ => {
            return Err(crate::coded(
                "pivot-unsupported",
                "Unsupported pivot layout sample",
            ));
        }
    };
    Ok(Presentation {
        sample: id,
        compact,
        outline,
        style,
        stripes,
        grand,
    })
}
fn plan_presentation(plan: &Value) -> Result<Presentation> {
    match (plan.get("sample"), plan.get("grandTotals")) {
        (None, None) => presentation("tabular-light", true),
        (Some(sample), Some(grand)) => presentation(
            sample
                .as_str()
                .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid pivot sample"))?,
            grand.as_bool().ok_or_else(|| {
                crate::coded("internal-invalid-field", "Invalid grand-total Boolean")
            })?,
        ),
        _ => Err(crate::coded(
            "pivot-unsupported",
            "Incomplete pivot presentation plan",
        )),
    }
}
fn table_presentation(table: &Tree) -> Result<Presentation> {
    let style = child(table, 0, "pivotTableStyleInfo")?;
    let sample = match table.nodes[style].attrs.get("name").map(String::as_str) {
        Some("PivotStyleLight16") => "tabular-light",
        Some("PivotStyleMedium9") => "compact-blue",
        Some("PivotStyleDark2") => "outline-dark",
        _ => return Err(crate::coded("pivot-unsupported", "Unsupported pivot style")),
    };
    let total = table.nodes[0]
        .attrs
        .get("colGrandTotals")
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing grand totals"))?;
    if !["0", "1"].contains(&total.as_str())
        || table.nodes[0].attrs.get("rowGrandTotals") != Some(total)
    {
        return Err(crate::coded(
            "pivot-unsupported",
            "Unqualified pivot grand totals",
        ));
    }
    presentation(sample, total == "1")
}
fn child(tree: &Tree, parent: usize, tag: &str) -> Result<usize> {
    tree.direct(parent, tag)?
        .ok_or_else(|| crate::coded("internal-missing-field", format!("Missing pivot {tag}")))
}
fn scalar(cell: &Value) -> Result<Value> {
    if !cell["formula"].is_null()
        || !cell["formulaType"].is_null()
        || cell.get("formulaOutput").is_some()
        || cell.get("spill").is_some()
        || cell["calculation"]["pending"] == true
        || cell["dataType"] == "error"
    {
        return Err(crate::coded(
            "pivot-protected",
            "Pivot source formulas/errors are protected",
        ));
    }
    let value = &cell["value"];
    if value.is_null() || value.as_str() == Some("") {
        return Ok(Value::Null);
    }
    if let Some(n) = value.as_f64() {
        if cell["rawValue"]
            .as_str()
            .is_some_and(|raw| crate::bulk::meaning(raw) != crate::bulk::meaning(&n.to_string()))
        {
            return Err(crate::coded(
                "pivot-bad-input",
                "Pivot literal precision is not exact",
            ));
        }
        if !n.is_finite() || n.fract() != 0.0 || n.abs() > MAX_SAFE as f64 {
            return Err(crate::coded(
                "pivot-bad-input",
                "Pivot source requires exact safe integers",
            ));
        }
        return Ok(json!(n as i64));
    }
    if let Some(text) = value.as_str() {
        if text.encode_utf16().count() > 128
            || text.to_ascii_lowercase().contains("_x")
                && text.as_bytes().windows(7).any(|w| {
                    w[0] == b'_'
                        && [b'x', b'X'].contains(&w[1])
                        && w[2..6].iter().all(u8::is_ascii_hexdigit)
                        && w[6] == b'_'
                })
        {
            return Err(crate::coded(
                "pivot-unsupported",
                "Pivot text needs separate qualification",
            ));
        }
        return Ok(value.clone());
    }
    if value.is_boolean() {
        return Ok(value.clone());
    }
    Err(crate::coded(
        "pivot-unsupported",
        "Unsupported pivot scalar",
    ))
}
#[cfg(test)]
mod formula_output_scalar_tests {
    use super::*;
    #[test]
    fn cached_source_output_cannot_become_a_pivot_literal() {
        let marker = json!({"schema":1,"master":"A1","range":"A1:A3"});
        for cell in [
            json!({"value":"Category","dataType":"text","rawValue":"Category","formulaOutput":marker}),
            json!({"value":7,"dataType":"number","rawValue":"7","formulaOutput":marker,"calculation":{"status":"calculated","value":7,"revision":"current"}}),
        ] {
            assert_eq!(
                crate::split_code(&scalar(&cell).unwrap_err()).0,
                Some("pivot-protected")
            );
        }
        assert_eq!(
            scalar(&json!({"value":"Category","dataType":"text","rawValue":"Category"})).unwrap(),
            json!("Category")
        );
        assert_eq!(
            scalar(&json!({"value":7,"dataType":"number","rawValue":"7"})).unwrap(),
            json!(7)
        );
    }
}
fn attrs(tree: &Tree, index: usize, allowed: &[&str]) -> Result<()> {
    let n = &tree.nodes[index];
    if n.ns != NS
        || n.attrs
            .keys()
            .any(|k| !k.starts_with("xmlns") && !allowed.contains(&k.as_str()))
    {
        return Err(crate::coded(
            "pivot-protected",
            "Extended pivot XML is protected",
        ));
    }
    Ok(())
}
fn exact(tree: &Tree, index: usize, tag: &str, pairs: &[(&str, String)]) -> Result<()> {
    attrs(
        tree,
        index,
        &pairs.iter().map(|(k, _)| *k).collect::<Vec<_>>(),
    )?;
    let n = &tree.nodes[index];
    if n.name != tag
        || !n.text.trim().is_empty()
        || pairs.iter().any(|(k, v)| n.attrs.get(*k) != Some(v))
    {
        return Err(crate::coded(
            "internal-invalid-field",
            format!("Invalid closed pivot {tag} profile"),
        ));
    }
    Ok(())
}
fn leaf(tree: &Tree, index: usize) -> Result<()> {
    if !tree.nodes[index].children.is_empty() {
        return Err(crate::coded("pivot-unsupported", "Extended pivot leaf"));
    }
    Ok(())
}
fn sheet_name(name: &str) -> bool {
    !name.is_empty()
        && name.encode_utf16().count() <= 31
        && !name
            .chars()
            .any(|c| c.is_control() || "[]:*?/\\".contains(c))
        && !name.starts_with('\'')
        && !name.ends_with('\'')
}
fn worksheet_identity(
    zip: &mut ZipArchive<crate::cancel::Reader<File>>,
    sheet: &str,
    file: &str,
) -> Result<()> {
    let wb = Tree::parse(&part(zip, "xl/workbook.xml")?)?;
    wb.root(NS, "workbook")?;
    let sheets = child(&wb, 0, "sheets")?;
    let matching = wb.nodes[sheets]
        .children
        .iter()
        .copied()
        .filter(|i| {
            wb.nodes[*i].ns == NS
                && wb.nodes[*i].attrs.get("name").map(String::as_str) == Some(sheet)
        })
        .collect::<Vec<_>>();
    if matching.len() != 1 {
        return Err(crate::coded(
            "pivot-refused",
            "Pivot worksheet identity is ambiguous",
        ));
    }
    let id = wb.nodes[matching[0]].attrs.get("r:id").ok_or_else(|| {
        crate::coded("internal-missing-field", "Missing pivot worksheet relation")
    })?;
    let rel = Tree::parse(&part(zip, "xl/_rels/workbook.xml.rels")?)?;
    rel.root(PACKAGE_REL_NS, "Relationships")?;
    let links = rel.nodes[0]
        .children
        .iter()
        .filter(|i| rel.nodes[**i].attrs.get("Id") == Some(id))
        .collect::<Vec<_>>();
    if links.len() != 1 {
        return Err(crate::coded(
            "pivot-ambiguous",
            "Ambiguous pivot worksheet relationship",
        ));
    }
    let n = &rel.nodes[*links[0]];
    let target = n
        .attrs
        .get("Target")
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing pivot worksheet target"))?;
    let resolved = if let Some(t) = target.strip_prefix('/') {
        t.to_owned()
    } else {
        format!("xl/{target}")
    };
    if n.ns != PACKAGE_REL_NS
        || n.attrs.contains_key("TargetMode")
        || n.attrs.get("Type").map(String::as_str)
            != Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
        || resolved != file
    {
        return Err(crate::coded("pivot-stale", "Pivot worksheet path changed"));
    }
    Ok(())
}
fn tree(plan: &Value, file: &str, root: &str) -> Result<Tree> {
    let t = Tree::parse(field(&plan["replacements"], file)?.as_bytes())?;
    t.root(NS, root)?;
    Ok(t)
}
fn relation_file(owner: &str) -> String {
    let at = owner.rfind('/').unwrap_or(0);
    format!("{}/_rels/{}.rels", &owner[..at], &owner[at + 1..])
}
fn relation(plan: &Value, file: &str, id: &str, kind: &str, target: &str) -> Result<()> {
    let t = Tree::parse(field(&plan["replacements"], file)?.as_bytes())?;
    t.root(PACKAGE_REL_NS, "Relationships")?;
    if t.nodes.len() != 2 || t.nodes[0].children.len() != 1 {
        return Err(crate::coded(
            "pivot-bad-input",
            "Pivot relationship is not isolated",
        ));
    }
    let n = &t.nodes[1];
    if n.ns != PACKAGE_REL_NS
        || n.name != "Relationship"
        || !n.children.is_empty()
        || n.attrs.get("Id").map(String::as_str) != Some(id)
        || n.attrs.get("Type").map(String::as_str)
            != Some(
                format!(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}"
                )
                .as_str(),
            )
        || n.attrs.get("Target").map(String::as_str) != Some(target)
        || n.attrs
            .keys()
            .any(|k| !k.starts_with("xmlns") && !["Id", "Type", "Target"].contains(&k.as_str()))
    {
        return Err(crate::coded(
            "pivot-malformed",
            "Invalid pivot relationship",
        ));
    }
    Ok(())
}
fn part(zip: &mut ZipArchive<crate::cancel::Reader<File>>, file: &str) -> Result<Vec<u8>> {
    let mut member = zip.by_name(file).map_err(err)?;
    if member.size() > XML_LIMIT as u64 {
        return Err(crate::coded(
            "pivot-limit",
            "Pivot source metadata exceeds its bound",
        ));
    }
    let mut bytes = Vec::new();
    member.read_to_end(&mut bytes).map_err(err)?;
    Ok(bytes)
}
fn envelope(tree: &Tree, remove: &[usize]) -> Vec<u8> {
    let mut spans = remove
        .iter()
        .map(|i| (tree.nodes[*i].start, tree.nodes[*i].end))
        .collect::<Vec<_>>();
    spans.sort();
    let mut out = Vec::new();
    let mut cursor = 0;
    for (a, b) in spans {
        out.extend_from_slice(&tree.bytes[cursor..a]);
        cursor = b;
    }
    out.extend_from_slice(&tree.bytes[cursor..]);
    out
}
fn source_identity(zip: &mut ZipArchive<crate::cancel::Reader<File>>, plan: &Value) -> Result<()> {
    let wb = Tree::parse(&part(zip, "xl/workbook.xml")?)?;
    wb.root(NS, "workbook")?;
    let sheets = child(&wb, 0, "sheets")?;
    let matching = wb.nodes[sheets]
        .children
        .iter()
        .copied()
        .filter(|i| {
            wb.nodes[*i].ns == NS
                && wb.nodes[*i].attrs.get("name").map(String::as_str)
                    == Some(field(plan, "sourceSheet").unwrap_or(""))
        })
        .collect::<Vec<_>>();
    if matching.len() != 1 {
        return Err(crate::coded(
            "pivot-refused",
            "Pivot source sheet is ambiguous",
        ));
    }
    let id = wb.nodes[matching[0]]
        .attrs
        .get("r:id")
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing source sheet relation"))?;
    let rel = Tree::parse(&part(zip, "xl/_rels/workbook.xml.rels")?)?;
    rel.root(PACKAGE_REL_NS, "Relationships")?;
    let links = rel.nodes[0]
        .children
        .iter()
        .filter(|i| rel.nodes[**i].attrs.get("Id") == Some(id))
        .collect::<Vec<_>>();
    if links.len() != 1 {
        return Err(crate::coded(
            "pivot-ambiguous",
            "Ambiguous pivot source relationship",
        ));
    }
    let n = &rel.nodes[*links[0]];
    let target = n
        .attrs
        .get("Target")
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing worksheet target"))?;
    let resolved = if let Some(t) = target.strip_prefix('/') {
        t.to_owned()
    } else {
        format!("xl/{target}")
    };
    if n.ns != PACKAGE_REL_NS
        || n.attrs.contains_key("TargetMode")
        || n.attrs.get("Type").map(String::as_str)
            != Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
        || resolved != field(plan, "sourcePath")?
    {
        return Err(crate::coded("pivot-stale", "Pivot source path changed"));
    }
    Ok(())
}
fn output_cells(t: &Tree, output: ((u32, u32), (u32, u32))) -> Result<BTreeMap<String, Value>> {
    let sd = child(t, 0, "sheetData")?;
    exact(t, sd, "sheetData", &[])?;
    let mut actual = BTreeMap::new();
    let mut seen = BTreeSet::new();
    let mut prior_row = 0u32;
    for row in &t.nodes[sd].children {
        let n = &t.nodes[*row];
        attrs(t, *row, &["r"])?;
        let r = n
            .attrs
            .get("r")
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing pivot row"))?
            .parse::<u32>()
            .map_err(err)?;
        if n.ns != NS
            || n.name != "row"
            || r <= prior_row
            || r > 1048576
            || !n.text.trim().is_empty()
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid pivot row ordering/profile",
            ));
        }
        prior_row = r;
        let mut prior_col = None;
        for cell in &n.children {
            let n = &t.nodes[*cell];
            attrs(t, *cell, &["r", "s", "t"])?;
            let reference = n
                .attrs
                .get("r")
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing pivot cell"))?;
            let p = position(reference)?;
            if n.ns != NS
                || n.name != "c"
                || p.0 + 1 != r
                || address(p.0, p.1) != *reference
                || prior_col.is_some_and(|c| p.1 <= c)
                || !seen.insert(reference.clone())
                || seen.len() > 10000
                || !n.text.trim().is_empty()
            {
                return Err(crate::coded(
                    "pivot-malformed",
                    "Invalid pivot target cell profile",
                ));
            }
            prior_col = Some(p.1);
            let inside =
                p.0 >= output.0.0 && p.0 <= output.1.0 && p.1 >= output.0.1 && p.1 <= output.1.1;
            if !inside {
                if n.attrs.get("s").is_some_and(|style| style != "0") || n.children.len() > 2 {
                    return Err(crate::coded(
                        "pivot-unsupported",
                        "Neighboring pivot style/content is unqualified",
                    ));
                }
                let mut kinds = BTreeSet::new();
                for node in &n.children {
                    let item = &t.nodes[*node];
                    if item.ns != NS || !kinds.insert(item.name.clone()) {
                        return Err(crate::coded(
                            "pivot-unsupported",
                            "Opaque neighboring pivot content",
                        ));
                    }
                    match item.name.as_str() {
                        "f" | "v" => {
                            attrs(t, *node, &[])?;
                            leaf(t, *node)?;
                        }
                        "is" => {
                            attrs(t, *node, &[])?;
                            let text = child(t, *node, "t")?;
                            if item.children.len() != 1 {
                                return Err(crate::coded(
                                    "pivot-unsupported",
                                    "Rich neighboring pivot text is unqualified",
                                ));
                            }
                            attrs(t, text, &["{http://www.w3.org/XML/1998/namespace}space"])?;
                            leaf(t, text)?;
                        }
                        _ => {
                            return Err(crate::coded(
                                "pivot-unsupported",
                                "Opaque neighboring pivot child",
                            ));
                        }
                    }
                }
                continue;
            }
            if n.attrs.contains_key("s") || n.children.len() != 1 {
                return Err(crate::coded(
                    "pivot-stale",
                    "Owned pivot output cell changed",
                ));
            }
            let value = if n.attrs.get("t").map(String::as_str) == Some("inlineStr") {
                let is = child(t, *cell, "is")?;
                exact(t, is, "is", &[])?;
                let text = child(t, is, "t")?;
                attrs(t, text, &["{http://www.w3.org/XML/1998/namespace}space"])?;
                leaf(t, text)?;
                if t.nodes[is].children.len() != 1
                    || t.nodes[text]
                        .attrs
                        .get("{http://www.w3.org/XML/1998/namespace}space")
                        .map(String::as_str)
                        != Some("preserve")
                {
                    return Err(crate::coded(
                        "pivot-stale",
                        "Pivot text whitespace/profile changed",
                    ));
                }
                json!(t.nodes[text].text)
            } else {
                if n.attrs.contains_key("t") {
                    return Err(crate::coded("pivot-stale", "Pivot numeric type changed"));
                }
                let v = child(t, *cell, "v")?;
                attrs(t, v, &[])?;
                leaf(t, v)?;
                json!(safe_integer(&t.nodes[v].text)?)
            };
            actual.insert(reference.clone(), value);
        }
    }
    Ok(actual)
}
fn target_mask(
    t: &Tree,
    a: ((u32, u32), (u32, u32)),
    b: ((u32, u32), (u32, u32)),
) -> Result<Vec<u8>> {
    let inside = |p: (u32, u32)| {
        [a, b].iter().any(|range| {
            p.0 >= range.0.0 && p.0 <= range.1.0 && p.1 >= range.0.1 && p.1 <= range.1.1
        })
    };
    let sd = child(t, 0, "sheetData")?;
    let mut spans = vec![(
        t.nodes[child(t, 0, "dimension")?].start,
        t.nodes[child(t, 0, "dimension")?].end,
    )];
    for row in &t.nodes[sd].children {
        let cells = &t.nodes[*row].children;
        let owned = cells
            .iter()
            .map(|cell| {
                position(t.nodes[*cell].attrs.get("r").ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing target coordinate")
                })?)
                .map(inside)
            })
            .collect::<Result<Vec<_>>>()?;
        if owned.iter().all(|value| *value) {
            spans.push((t.nodes[*row].start, t.nodes[*row].end));
        } else {
            for (cell, owned) in cells.iter().zip(owned) {
                if owned {
                    spans.push((t.nodes[*cell].start, t.nodes[*cell].end));
                }
            }
        }
    }
    spans.sort();
    let mut out = Vec::new();
    let mut cursor = 0;
    for (start, end) in spans {
        out.extend_from_slice(&t.bytes[cursor..start]);
        cursor = end;
    }
    out.extend_from_slice(&t.bytes[cursor..]);
    Ok(out)
}
fn prior_display(zip: &mut ZipArchive<crate::cancel::Reader<File>>, target: &str) -> Result<()> {
    let parts = Parts::from_zip(zip)?;
    let t = Tree::parse(&part(zip, target)?)?;
    t.root(NS, "worksheet")?;
    attrs(&t, 0, &[])?;
    if t.nodes[0].children.len() != 5 {
        return Err(crate::coded(
            "pivot-unsupported",
            "Extended prior pivot display",
        ));
    }
    let dim = child(&t, 0, "dimension")?;
    attrs(&t, dim, &["ref"])?;
    leaf(&t, dim)?;
    let (a, b) =
        crate::bulk::bounds(t.nodes[dim].attrs.get("ref").ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing prior pivot dimension")
        })?)?;
    if a != (0, 0) {
        return Err(crate::coded(
            "pivot-stale",
            "Prior pivot display geometry changed",
        ));
    }
    let views = child(&t, 0, "sheetViews")?;
    exact(&t, views, "sheetViews", &[])?;
    if t.nodes[views].children.len() != 1 {
        return Err(crate::coded(
            "pivot-stale",
            "Prior pivot sheet views changed",
        ));
    }
    let view = t.nodes[views].children[0];
    exact(&t, view, "sheetView", &[("workbookViewId", "0".into())])?;
    leaf(&t, view)?;
    let format = child(&t, 0, "sheetFormatPr")?;
    exact(
        &t,
        format,
        "sheetFormatPr",
        &[("defaultRowHeight", "15".into())],
    )?;
    leaf(&t, format)?;
    let cols = child(&t, 0, "cols")?;
    exact(&t, cols, "cols", &[])?;
    if t.nodes[cols].children.len() != 2 {
        return Err(crate::coded("pivot-stale", "Prior pivot columns changed"));
    }
    for (i, col) in t.nodes[cols].children.iter().enumerate() {
        exact(
            &t,
            *col,
            "col",
            &[
                ("min", (i + 1).to_string()),
                ("max", (i + 1).to_string()),
                ("width", if i == 0 { "24" } else { "20" }.into()),
                ("customWidth", "1".into()),
            ],
        )?;
        leaf(&t, *col)?;
    }
    let sd = child(&t, 0, "sheetData")?;
    exact(&t, sd, "sheetData", &[])?;
    let prior_table = Tree::parse(&part(zip, parts.table)?)?;
    prior_table.root(NS, "pivotTableDefinition")?;
    let location = child(&prior_table, 0, "location")?;
    let output = crate::bulk::bounds(
        prior_table.nodes[location]
            .attrs
            .get("ref")
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing prior pivot output"))?,
    )?;
    if output.0 != (0, 0)
        || output.1.1 != 1
        || output.1.0 < 1
        || output.1.0 > 513
        || b.0 < output.1.0
        || b.1 < output.1.1
    {
        return Err(crate::coded(
            "pivot-unsupported",
            "Prior pivot output geometry is unqualified",
        ));
    }
    let actual = output_cells(&t, output)?;
    let prior_presentation = table_presentation(&prior_table)?;
    // Reconstruct the previously displayed aggregate from its stored cache, not
    // the newly edited source. This guards refresh against erased user changes.
    let def = Tree::parse(&part(zip, parts.definition)?)?;
    def.root(NS, "pivotCacheDefinition")?;
    let fields = child(&def, 0, "cacheFields")?;
    let table = Tree::parse(&part(zip, parts.table)?)?;
    table.root(NS, "pivotTableDefinition")?;
    let rf = child(&table, 0, "rowFields")?;
    let df = child(&table, 0, "dataFields")?;
    if table.nodes[rf].children.len() != 1 || table.nodes[df].children.len() != 1 {
        return Err(crate::coded("pivot-stale", "Prior pivot axes changed"));
    }
    let row = table.nodes[table.nodes[rf].children[0]]
        .attrs
        .get("x")
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing prior row field"))?
        .parse::<usize>()
        .map_err(err)?;
    let data = &table.nodes[table.nodes[df].children[0]];
    let val = data
        .attrs
        .get("fld")
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing prior value field"))?
        .parse::<usize>()
        .map_err(err)?;
    let aggregate = data
        .attrs
        .get("subtotal")
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing prior aggregate"))?;
    if !["sum", "count"].contains(&aggregate.as_str()) || row == val {
        return Err(crate::coded(
            "pivot-unsupported",
            "Unsupported prior pivot axes",
        ));
    }
    let field_values = def.nodes[fields]
        .children
        .iter()
        .map(|f| {
            let si = child(&def, *f, "sharedItems")?;
            def.nodes[si]
                .children
                .iter()
                .map(|i| {
                    let n = &def.nodes[*i];
                    match n.name.as_str() {
                        "s" => Ok(json!(n.attrs.get("v").ok_or_else(|| crate::coded(
                            "internal-missing-field",
                            "Missing prior string"
                        ))?)),
                        "n" => Ok(json!(safe_integer(n.attrs.get("v").ok_or_else(|| {
                            crate::coded("internal-missing-field", "Missing prior number")
                        })?)?)),
                        "m" => Ok(Value::Null),
                        "b" => Ok(json!(n.attrs.get("v").map(String::as_str) == Some("1"))),
                        _ => Err(crate::coded(
                            "pivot-unsupported",
                            "Unsupported prior shared item",
                        )),
                    }
                })
                .collect::<Result<Vec<Value>>>()
        })
        .collect::<Result<Vec<_>>>()?;
    let records = Tree::parse(&part(zip, parts.records)?)?;
    records.root(NS, "pivotCacheRecords")?;
    let mut groups: Vec<(String, i64)> = Vec::new();
    let mut total = 0i64;
    for record in &records.nodes[0].children {
        let r = &records.nodes[*record];
        let item = |field: usize| -> Result<&Value> {
            let cell = *r.children.get(field).ok_or_else(|| {
                crate::coded("internal-missing-field", "Prior record field missing")
            })?;
            let at = records.nodes[cell]
                .attrs
                .get("v")
                .ok_or_else(|| {
                    crate::coded("internal-missing-field", "Prior record index missing")
                })?
                .parse::<usize>()
                .map_err(err)?;
            field_values
                .get(field)
                .and_then(|v| v.get(at))
                .ok_or_else(|| crate::coded("pivot-limit", "Prior cache index out of bounds"))
        };
        let label = item(row)?
            .as_str()
            .ok_or_else(|| crate::coded("pivot-bad-input", "Prior pivot row is not text"))?
            .to_owned();
        let v = item(val)?;
        let amount = if v.is_null() {
            0
        } else if aggregate == "count" {
            v.as_i64().ok_or_else(|| {
                crate::coded("pivot-bad-input", "Prior count value is not numeric")
            })?;
            1
        } else {
            v.as_i64()
                .ok_or_else(|| crate::coded("pivot-bad-input", "Prior sum value is not numeric"))?
        };
        let at = groups.iter().position(|g| g.0 == label).unwrap_or_else(|| {
            groups.push((label, 0));
            groups.len() - 1
        });
        groups[at].1 = groups[at]
            .1
            .checked_add(amount)
            .filter(|n| n.unsigned_abs() <= MAX_SAFE as u64)
            .ok_or_else(|| crate::coded("pivot-limit", "Prior aggregate overflow"))?;
        total = total
            .checked_add(amount)
            .filter(|n| n.unsigned_abs() <= MAX_SAFE as u64)
            .ok_or_else(|| crate::coded("pivot-limit", "Prior total overflow"))?;
    }
    let header = def.nodes[*def.nodes[fields]
        .children
        .get(row)
        .ok_or_else(|| crate::coded("internal-missing-field", "Prior row header missing"))?]
    .attrs
    .get("name")
    .ok_or_else(|| crate::coded("internal-missing-field", "Prior row header missing"))?;
    let mut expected = BTreeMap::from([
        (
            "A1".to_owned(),
            json!(if prior_presentation.compact == "1" {
                "Row Labels"
            } else {
                header.as_str()
            }),
        ),
        (
            "B1".to_owned(),
            json!(data.attrs.get("name").ok_or_else(|| crate::coded(
                "internal-missing-field",
                "Prior aggregate caption missing"
            ))?),
        ),
    ]);
    for (i, g) in groups.iter().enumerate() {
        expected.insert(address(i as u32 + 1, 0), json!(g.0));
        expected.insert(address(i as u32 + 1, 1), json!(g.1));
    }
    if prior_presentation.grand {
        expected.insert(address(groups.len() as u32 + 1, 0), json!("Grand Total"));
        expected.insert(address(groups.len() as u32 + 1, 1), json!(total));
    }
    if actual != expected {
        return Err(crate::coded(
            "pivot-stale",
            "Prior pivot output changed outside its cache; preserved and protected",
        ));
    }
    Ok(())
}
fn creation_graph(zip: &mut ZipArchive<crate::cancel::Reader<File>>, plan: &Value) -> Result<()> {
    let parts = Parts::from_plan(plan)?;
    let original = Tree::parse(&part(zip, "xl/workbook.xml")?)?;
    if original.nodes[0].children.iter().any(|i| {
        ["pivotCaches", "extLst", "externalReferences"].contains(&original.nodes[*i].name.as_str())
    }) || zip
        .file_names()
        .any(|n| n.starts_with("xl/pivotTables/") || n.starts_with("xl/pivotCache/"))
    {
        return Err(crate::coded(
            "pivot-protected",
            "Existing or extended pivot graph is protected",
        ));
    }
    let wb = Tree::parse(field(&plan["replacements"], "xl/workbook.xml")?.as_bytes())?;
    let sheets = child(&wb, 0, "sheets")?;
    let mut sheet_ids = BTreeSet::new();
    let mut names = BTreeSet::new();
    let mut sheet_links = BTreeSet::new();
    let mut added = None;
    for i in &wb.nodes[sheets].children {
        let n = &wb.nodes[*i];
        if n.ns != NS
            || n.name != "sheet"
            || !sheet_ids.insert(
                n.attrs
                    .get("sheetId")
                    .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheet ID"))?
                    .parse::<u32>()
                    .map_err(err)?,
            )
            || !names.insert(
                n.attrs
                    .get("name")
                    .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheet name"))?
                    .to_lowercase(),
            )
            || !sheet_links.insert(
                n.attrs
                    .get("r:id")
                    .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheet link"))?
                    .clone(),
            )
        {
            return Err(crate::coded(
                "pivot-ambiguous",
                "Ambiguous pivot worksheet graph",
            ));
        }
        if n.attrs.get("name").map(String::as_str) == Some(field(plan, "targetSheet")?) {
            attrs(&wb, *i, &["name", "sheetId", "r:id"])?;
            leaf(&wb, *i)?;
            added = Some(*i);
        }
    }
    let added = added
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing added pivot worksheet"))?;
    let caches = child(&wb, 0, "pivotCaches")?;
    exact(&wb, caches, "pivotCaches", &[])?;
    if wb.nodes[caches].children.len() != 1 {
        return Err(crate::coded("pivot-stale", "Pivot cache ownership changed"));
    }
    let cache = wb.nodes[caches].children[0];
    attrs(&wb, cache, &["cacheId", "r:id"])?;
    leaf(&wb, cache)?;
    if wb.nodes[cache].name != "pivotCache"
        || wb.nodes[cache].attrs.get("cacheId").map(String::as_str) != Some("1")
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid pivot cache owner",
        ));
    }
    let rel = Tree::parse(field(&plan["replacements"], "xl/_rels/workbook.xml.rels")?.as_bytes())?;
    let mut ids = BTreeSet::new();
    for i in &rel.nodes[0].children {
        let n = &rel.nodes[*i];
        if n.ns != PACKAGE_REL_NS
            || n.name != "Relationship"
            || !ids.insert(
                n.attrs
                    .get("Id")
                    .ok_or_else(|| crate::coded("internal-missing-field", "Missing relation ID"))?
                    .clone(),
            )
        {
            return Err(crate::coded(
                "pivot-duplicate",
                "Duplicate pivot relationship",
            ));
        }
    }
    for (id, kind, target) in [
        (
            wb.nodes[added]
                .attrs
                .get("r:id")
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheet link"))?,
            "worksheet",
            field(plan, "targetPath")?,
        ),
        (
            wb.nodes[cache]
                .attrs
                .get("r:id")
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing cache link"))?,
            "pivotCacheDefinition",
            parts.definition,
        ),
    ] {
        let link = rel.nodes[0]
            .children
            .iter()
            .find(|i| rel.nodes[**i].attrs.get("Id") == Some(id))
            .ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing owned pivot relationship")
            })?;
        let n = &rel.nodes[*link];
        leaf(&rel, *link)?;
        if n.attrs
            .keys()
            .any(|k| !k.starts_with("xmlns") && !["Id", "Type", "Target"].contains(&k.as_str()))
            || n.attrs.get("Type")
                != Some(&format!(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}"
                ))
            || n.attrs.get("Target").map(String::as_str) != target.strip_prefix("xl/")
        {
            return Err(crate::coded(
                "pivot-stale",
                "Pivot relationship MIME/path changed",
            ));
        }
    }
    let types = Tree::parse(field(&plan["replacements"], "[Content_Types].xml")?.as_bytes())?;
    let mut overrides = BTreeMap::new();
    let mut defaults = BTreeSet::new();
    for i in &types.nodes[0].children {
        let n = &types.nodes[*i];
        if n.ns != TYPE_NS {
            return Err(crate::coded("pivot-unsupported", "Extended content type"));
        }
        if n.name == "Override" {
            let name = n.attrs.get("PartName").ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing content type part")
            })?;
            if overrides.insert(name.clone(), *i).is_some() {
                return Err(crate::coded(
                    "pivot-duplicate",
                    "Duplicate content type override",
                ));
            }
        } else if n.name == "Default" {
            if !defaults.insert(
                n.attrs
                    .get("Extension")
                    .ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing content type extension")
                    })?
                    .to_ascii_lowercase(),
            ) {
                return Err(crate::coded(
                    "pivot-duplicate",
                    "Duplicate content type default",
                ));
            }
        } else {
            return Err(crate::coded("pivot-unsupported", "Extended content types"));
        }
    }
    for (file, mime) in [
        (
            field(plan, "targetPath")?,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        ),
        (
            parts.table,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
        ),
        (
            parts.definition,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
        ),
        (
            parts.records,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml",
        ),
    ] {
        let i = *overrides
            .get(&format!("/{file}"))
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing pivot content type"))?;
        let n = &types.nodes[i];
        leaf(&types, i)?;
        if n.attrs
            .keys()
            .any(|k| !k.starts_with("xmlns") && !["PartName", "ContentType"].contains(&k.as_str()))
            || n.attrs.get("ContentType").map(String::as_str) != Some(mime)
        {
            return Err(crate::coded("pivot-stale", "Pivot content type changed"));
        }
    }
    Ok(())
}
pub(super) fn validate(
    root: &Path,
    db: &rusqlite::Connection,
    plan: &Value,
    input: &Value,
) -> Result<()> {
    let parts = Parts::from_plan(plan)?;
    let existing = if ["refreshPivot", "updatePivot", "setPivotLayout"]
        .iter()
        .any(|kind| plan["kind"] == *kind)
    {
        cached_profile(root)?
    } else {
        None
    };
    if ["refreshPivot", "updatePivot", "setPivotLayout"]
        .iter()
        .any(|kind| plan["kind"] == *kind)
    {
        let (original, _) = existing.as_ref().ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing qualified owned pivot source profile",
            )
        })?;
        let binding = json!({"name":original["name"],"targetSheet":original["targetSheet"],"targetPath":original["targetPath"],"revision":original["priorRevision"],"sourceSheet":original["sourceSheet"],"sourcePath":original["sourcePath"],"range":original["sourceRange"]});
        for (key, bound) in [
            ("name", "name"),
            ("targetSheet", "targetSheet"),
            ("targetPath", "targetPath"),
            ("priorRevision", "revision"),
        ] {
            if plan[key] != binding[bound] {
                return Err(crate::coded(
                    "pivot-unsupported",
                    "Pivot editing cannot retarget the original identity or cache",
                ));
            }
        }
        if plan["kind"] == "refreshPivot" {
            let mut zip = ZipArchive::new(crate::cancel::Reader(
                File::open(local(root, "source.xlsx")?).map_err(err)?,
            ))
            .map_err(err)?;
            let table = Tree::parse(&part(&mut zip, parts.table)?)?;
            let rows = child(&table, 0, "rowFields")?;
            let values = child(&table, 0, "dataFields")?;
            let row = table.nodes[*table.nodes[rows].children.first().ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing original row field")
            })?]
            .attrs
            .get("x")
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing row index"))?
            .parse::<usize>()
            .map_err(err)?;
            let value = &table.nodes[*table.nodes[values].children.first().ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing original value field")
            })?];
            if plan["sourceSheet"] != binding["sourceSheet"]
                || plan["sourcePath"] != binding["sourcePath"]
                || plan["sourceRange"] != binding["range"]
                || plan["rowField"] != row
                || plan["valueField"]
                    != value
                        .attrs
                        .get("fld")
                        .ok_or_else(|| {
                            crate::coded("internal-missing-field", "Missing value index")
                        })?
                        .parse::<usize>()
                        .map_err(err)?
                || plan["aggregate"]
                    != value
                        .attrs
                        .get("subtotal")
                        .ok_or_else(|| crate::coded("internal-missing-field", "Missing aggregate"))?
                        .as_str()
            {
                return Err(crate::coded(
                    "pivot-unsupported",
                    "Refresh preserves pivot fields; use field editing to change them",
                ));
            }
        }
    }
    if ["refreshPivot", "updatePivot", "setPivotLayout"]
        .iter()
        .any(|kind| plan["kind"] == *kind)
    {
        let (original, dataset) = existing.as_ref().ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing qualified pivot presentation",
            )
        })?;
        let before = plan_presentation(original)?;
        let after = plan_presentation(plan)?;
        if plan["kind"] == "setPivotLayout" {
            for key in [
                "sourceSheet",
                "sourcePath",
                "sourceRange",
                "rowField",
                "valueField",
                "aggregate",
            ] {
                if plan[key] != original[key] {
                    return Err(crate::coded(
                        "pivot-unsupported",
                        "Pivot layout cannot change source or fields",
                    ));
                }
            }
            for file in [parts.definition, parts.records] {
                if plan["replacements"][file] != original["replacements"][file] {
                    return Err(crate::coded(
                        "pivot-unsupported",
                        "Pivot layout cannot refresh or replace its cache",
                    ));
                }
            }
            return validate_inner(root, db, plan, input, Some(dataset));
        }
        if before.sample != after.sample || before.grand != after.grand {
            return Err(crate::coded(
                "pivot-unsupported",
                "Refresh and field editing preserve pivot layout/totals",
            ));
        }
    }
    validate_inner(root, db, plan, input, None)
}
fn validate_inner(
    root: &Path,
    db: &rusqlite::Connection,
    plan: &Value,
    input: &Value,
    cached: Option<&BTreeMap<(u32, u32), Value>>,
) -> Result<()> {
    let parts = Parts::from_plan(plan)?;
    let keys = [
        "kind",
        "name",
        "sourceSheet",
        "sourcePath",
        "sourceRange",
        "rowField",
        "valueField",
        "aggregate",
        "targetSheet",
        "targetPath",
        "cacheId",
        "parts",
        "priorRevision",
        "replacements",
        "removed",
        "copies",
        "fragments",
        "sample",
        "grandTotals",
    ];
    if plan.as_object().is_none_or(|o| {
        (o.len() != keys.len() && o.len() != keys.len() - 2)
            || o.keys().any(|k| !keys.contains(&k.as_str()))
    }) || plan["removed"].as_array().is_none_or(|v| !v.is_empty())
        || plan["copies"].as_object().is_none_or(|v| !v.is_empty())
        || plan["fragments"].as_object().is_none_or(|v| !v.is_empty())
        || plan["cacheId"] != 1
        || plan["parts"]
            != json!({"table":parts.table,"definition":parts.definition,"records":parts.records})
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid closed pivot package plan",
        ));
    }
    let presentation = plan_presentation(plan)?;
    let refresh = plan["kind"] == "refreshPivot"
        || plan["kind"] == "updatePivot"
        || plan["kind"] == "setPivotLayout";
    let target = field(plan, "targetPath")?;
    let name = field(plan, "name")?;
    if name.is_empty()
        || name.len() > 128
        || !name
            .bytes()
            .enumerate()
            .all(|(i, c)| c == b'_' || c.is_ascii_alphabetic() || i > 0 && c.is_ascii_digit())
        || !sheet_name(field(plan, "targetSheet")?)
        || field(plan, "targetSheet")?.eq_ignore_ascii_case(field(plan, "sourceSheet")?)
        || target
            .strip_prefix("xl/worksheets/sheet")
            .and_then(|s| s.strip_suffix(".xml"))
            .is_none_or(|s| {
                s.is_empty() || !s.bytes().all(|c| c.is_ascii_digit()) || s.starts_with('0')
            })
        || crate::xml::position(name).is_ok()
        || ["r", "c"].contains(&name.to_ascii_lowercase().as_str())
    {
        return Err(crate::coded("pivot-malformed", "Invalid pivot identity"));
    }
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, field(input, "source")?)?).map_err(err)?,
    ))
    .map_err(err)?;
    source_identity(&mut zip, plan)?;
    if refresh {
        worksheet_identity(&mut zip, field(plan, "targetSheet")?, target)?;
        prior_display(&mut zip, target)?;
    }
    let defs = tree(plan, parts.definition, "pivotCacheDefinition")?;
    let records = tree(plan, parts.records, "pivotCacheRecords")?;
    let table = tree(plan, parts.table, "pivotTableDefinition")?;
    let display = tree(plan, target, "worksheet")?;
    attrs(
        &defs,
        0,
        &[
            "r:id",
            "saveData",
            "refreshOnLoad",
            "enableRefresh",
            "createdVersion",
            "refreshedVersion",
            "minRefreshableVersion",
            "recordCount",
        ],
    )?;
    for (k, v) in [
        ("saveData", "1"),
        ("enableRefresh", "1"),
        ("createdVersion", "6"),
        ("refreshedVersion", "6"),
        ("minRefreshableVersion", "3"),
    ] {
        if defs.nodes[0].attrs.get(k).map(String::as_str) != Some(v) {
            return Err(crate::coded("pivot-stale", "Pivot cache options changed"));
        }
    }
    if defs.nodes[0].attrs.get("r:id").map(String::as_str) != Some("rIdRecords")
        || defs.nodes[0].attrs.get("refreshOnLoad").map(String::as_str) != Some("0")
        || defs.nodes[0].children.len() != 2
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid pivot cache profile",
        ));
    }
    let source = child(&defs, 0, "cacheSource")?;
    attrs(&defs, source, &["type"])?;
    let ws = child(&defs, source, "worksheetSource")?;
    attrs(&defs, ws, &["ref", "sheet"])?;
    if defs.nodes[source].children.len() != 1 {
        return Err(crate::coded("pivot-unsupported", "Extended pivot source"));
    }
    leaf(&defs, ws)?;
    if defs.nodes[source].attrs.get("type").map(String::as_str) != Some("worksheet")
        || defs.nodes[ws].attrs.get("ref").map(String::as_str) != Some(field(plan, "sourceRange")?)
        || defs.nodes[ws].attrs.get("sheet").map(String::as_str)
            != Some(field(plan, "sourceSheet")?)
    {
        return Err(crate::coded(
            "pivot-bad-input",
            "Pivot cache source does not match",
        ));
    }
    let (a, b) = crate::bulk::bounds(field(plan, "sourceRange")?)?;
    let rows = b.0 - a.0 + 1;
    let columns = b.1 - a.1 + 1;
    if rows < 2 || rows > 1001 || columns < 2 || columns > 8 || rows * columns > 8000 {
        return Err(crate::coded(
            "pivot-limit",
            "Pivot source geometry exceeds its bound",
        ));
    }
    let row = plan["rowField"]
        .as_u64()
        .filter(|v| *v < columns as u64)
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid pivot row field"))?
        as usize;
    let value = plan["valueField"]
        .as_u64()
        .filter(|v| *v < columns as u64 && *v != row as u64)
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid pivot value field"))?
        as usize;
    let aggregate = field(plan, "aggregate")?;
    if !["sum", "count"].contains(&aggregate) {
        return Err(crate::coded(
            "pivot-unsupported",
            "Unsupported pivot aggregate",
        ));
    }
    let version = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing pivot version"))?;
    let mut source_view = crate::package_edit::auditing_metadata(root)?;
    let wb = Tree::parse(&part(&mut zip, "xl/workbook.xml")?)?;
    let calendar = wb.direct(0, "workbookPr")?.is_some_and(|i| {
        wb.nodes[i]
            .attrs
            .get("date1904")
            .is_some_and(|v| ["1", "true"].contains(&v.as_str()))
    });
    for sheet in source_view["sheets"]
        .as_array_mut()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing pivot source sheets"))?
    {
        sheet["date1904"] = json!(calendar);
    }
    if cached.is_none() {
        let (plan_id, scoped) = crate::calculation_policy::graph(root, db, version, &source_view)?;
        if !scoped {
            return Err(crate::coded(
                "pivot-unsupported",
                "Pivot formula source graph is not qualified",
            ));
        }
        db.execute_batch("CREATE TEMP TABLE spill_protected(sheet TEXT,mr INTEGER,mc INTEGER,r0 INTEGER,c0 INTEGER,r1 INTEGER,c1 INTEGER,PRIMARY KEY(sheet,mr,mc)) WITHOUT ROWID;").map_err(err)?;
        db.execute("INSERT OR IGNORE INTO spill_protected SELECT sheet,mr,mc,r0,c0,r1,c1 FROM calculation_outputs WHERE plan=?1",[plan_id]).map_err(err)?;
        db.execute("DELETE FROM dependency_selected", [])
            .map_err(err)?;
        db.execute("INSERT INTO dependency_selected SELECT plan,sheet,ref FROM dependency_formulas WHERE plan=?1 AND sheet=?2 AND r BETWEEN ?3 AND ?4 AND c BETWEEN ?5 AND ?6",rusqlite::params![plan_id,field(plan,"sourceSheet")?,a.0,b.0,a.1,b.1]).map_err(err)?;
        crate::calculation_policy::scalar_prerequisites(db, plan_id, &source_view)?;
        let mut query=db.prepare_cached("SELECT s.sheet,f.r,f.c FROM dependency_selected s JOIN dependency_formulas f ON f.plan=s.plan AND f.sheet=s.sheet AND f.ref=s.ref WHERE s.plan=?1").map_err(err)?;
        let prerequisites = query
            .query_map([plan_id], |row| {
                Ok((
                    row.get::<_, String>(0)?,
                    row.get::<_, u32>(1)?,
                    row.get::<_, u32>(2)?,
                ))
            })
            .map_err(err)?;
        for prerequisite in prerequisites {
            let (name, r, c) = prerequisite.map_err(err)?;
            let path = source_view["sheets"]
                .as_array()
                .unwrap()
                .iter()
                .find(|sheet| sheet["name"] == name)
                .and_then(|sheet| sheet["path"].as_str())
                .ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing prerequisite worksheet")
                })?;
            if crate::calculation_policy::pending(db, version, path, (r, c))? {
                return Err(crate::coded(
                    "pivot-refused",
                    "Pivot source prerequisite is pending calculation",
                ));
            }
        }
    } else {
        crate::calculation::snapshot(db, version)?;
    }
    db.execute_batch("CREATE TEMP TABLE IF NOT EXISTS calc_scenarios(sheet TEXT,r INTEGER,c INTEGER,model TEXT,input1 TEXT,value1 TEXT,input2 TEXT,value2 TEXT,PRIMARY KEY(sheet,r,c));").map_err(err)?;
    let mut evaluator = crate::formula::Eval::from_sqlite(&source_view, db, false);
    evaluator.fence_spill_inputs();
    let mut cells = BTreeMap::new();
    let mut after = None;
    loop {
        let page = crate::calculation::page(db, field(plan, "sourcePath")?, a, b, after, false)?;
        for (reference, cell) in &page {
            crate::spill::source_content_proof(
                root,
                db,
                field(plan, "sourcePath")?,
                position(reference)?,
                cell["formula"].is_string(),
            )?;
            if crate::spill::source_attributes(
                root,
                db,
                field(plan, "sourcePath")?,
                position(reference)?,
            )?
            .is_some_and(|attrs| {
                attrs
                    .keys()
                    .any(|k| !["r", "s", "t", "xmlns"].contains(&k.as_str()))
                    || attrs.get("xmlns").is_some_and(|ns| ns.as_bytes() != NS)
            }) {
                return Err(crate::coded(
                    "pivot-unsupported",
                    "Opaque pivot source cell attributes require qualification",
                ));
            }
            if let Some(cached) = cached {
                cells.insert(
                    position(reference)?,
                    cached
                        .get(&position(reference)?)
                        .cloned()
                        .unwrap_or(Value::Null),
                );
                continue;
            }
            let mut current = cell.clone();
            crate::calculation_policy::mask(
                db,
                version,
                field(plan, "sourcePath")?,
                reference,
                &mut current,
            )?;
            if current["formula"].is_string() {
                if !current["formulaType"].is_null()
                    || current.get("spill").is_some()
                    || current["calculation"]["pending"] == true
                    || current["calculation"]["status"] != "calculated"
                {
                    return Err(crate::coded(
                        "pivot-bad-input",
                        "Pivot formula source has no current ordinary calculation",
                    ));
                }
                let number = evaluator
                    .cell(field(plan, "sourceSheet")?, reference)?
                    .ok_or_else(|| crate::coded("pivot-refused", "Pivot formula input is blank"))?;
                let recorded = current["calculation"]["value"].as_f64().ok_or_else(|| {
                    crate::coded("pivot-bad-input", "Pivot formula input is not numeric")
                })?;
                if number.to_bits() != recorded.to_bits()
                    || !number.is_finite()
                    || number.fract() != 0.0
                    || number.abs() > MAX_SAFE as f64
                {
                    return Err(crate::coded(
                        "internal-mismatch",
                        "Pivot formula cache differs from independent current numeric proof",
                    ));
                }
                cells.insert(position(reference)?, json!(number as i64));
            } else {
                cells.insert(position(reference)?, scalar(&current)?);
            }
        }
        if page.len() < 512 {
            break;
        }
        after = Some(position(
            &page
                .last()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing pivot page"))?
                .0,
        )?);
    }
    let mut headers = Vec::new();
    let mut lower = BTreeSet::new();
    for c in a.1..=b.1 {
        let s = cells
            .get(&(a.0, c))
            .and_then(Value::as_str)
            .filter(|s| !s.trim().is_empty())
            .ok_or_else(|| crate::coded("pivot-bad-input", "Pivot headers require text"))?;
        if !lower.insert(s.to_lowercase()) {
            return Err(crate::coded("pivot-duplicate", "Duplicate pivot headers"));
        }
        headers.push(s.to_owned());
    }
    let fields = child(&defs, 0, "cacheFields")?;
    attrs(&defs, fields, &["count"])?;
    if defs.nodes[fields].children.len() != columns as usize
        || defs.nodes[fields].attrs.get("count").map(String::as_str)
            != Some(columns.to_string().as_str())
    {
        return Err(crate::coded(
            "pivot-stale",
            "Pivot cache field count changed",
        ));
    }
    let mut shared = Vec::new();
    for (i, index) in defs.nodes[fields].children.iter().enumerate() {
        attrs(&defs, *index, &["name", "numFmtId"])?;
        let n = &defs.nodes[*index];
        if n.name != "cacheField"
            || n.attrs.get("name") != Some(&headers[i])
            || n.attrs.get("numFmtId").map(String::as_str) != Some("0")
            || n.children.len() != 1
        {
            return Err(crate::coded(
                "pivot-stale",
                "Pivot cache field identity changed",
            ));
        }
        let si = child(&defs, *index, "sharedItems")?;
        attrs(
            &defs,
            si,
            &[
                "count",
                "containsBlank",
                "containsString",
                "containsNumber",
                "containsInteger",
                "containsDate",
                "containsNonDate",
                "containsSemiMixedTypes",
                "containsMixedTypes",
            ],
        )?;
        let mut values = Vec::new();
        for j in &defs.nodes[si].children {
            attrs(&defs, *j, &["v"])?;
            let n = &defs.nodes[*j];
            if !n.children.is_empty() || !n.text.is_empty() {
                return Err(crate::coded(
                    "pivot-unsupported",
                    "Extended pivot shared value",
                ));
            }
            let v = match n.name.as_str() {
                "m" => Value::Null,
                "s" => json!(n.attrs.get("v").ok_or_else(|| crate::coded(
                    "internal-missing-field",
                    "Missing pivot string"
                ))?),
                "b" => match n.attrs.get("v").map(String::as_str) {
                    Some("0") => json!(false),
                    Some("1") => json!(true),
                    _ => {
                        return Err(crate::coded(
                            "internal-invalid-field",
                            "Invalid pivot boolean",
                        ));
                    }
                },
                "n" => json!(safe_integer(n.attrs.get("v").ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing pivot number")
                })?)?),
                _ => {
                    return Err(crate::coded(
                        "pivot-unsupported",
                        "Unqualified pivot shared value",
                    ));
                }
            };
            if values.contains(&v) {
                return Err(crate::coded(
                    "pivot-duplicate",
                    "Duplicate pivot shared value",
                ));
            }
            values.push(v);
        }
        if values.len() > 1000
            || defs.nodes[si].attrs.get("count").map(String::as_str)
                != Some(values.len().to_string().as_str())
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid shared pivot count",
            ));
        }
        shared.push(values);
    }
    attrs(&records, 0, &["count"])?;
    if records.nodes[0].children.len() != rows as usize - 1
        || records.nodes[0].attrs.get("count").map(String::as_str)
            != Some((rows - 1).to_string().as_str())
        || defs.nodes[0].attrs.get("recordCount").map(String::as_str)
            != Some((rows - 1).to_string().as_str())
    {
        return Err(crate::coded("pivot-stale", "Pivot record count changed"));
    }
    let mut groups: Vec<(String, usize, i64)> = Vec::new();
    let mut labels = BTreeMap::new();
    let mut total = 0i64;
    let mut seen = vec![Vec::<Value>::new(); columns as usize];
    for (i, index) in records.nodes[0].children.iter().enumerate() {
        attrs(&records, *index, &[])?;
        let r = &records.nodes[*index];
        if r.name != "r" || r.children.len() != columns as usize {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid pivot record shape",
            ));
        }
        let mut indexes = Vec::new();
        let mut values = Vec::new();
        for (c, j) in r.children.iter().enumerate() {
            attrs(&records, *j, &["v"])?;
            let n = &records.nodes[*j];
            if n.name != "x" || !n.children.is_empty() {
                return Err(crate::coded(
                    "pivot-unsupported",
                    "Unqualified pivot record",
                ));
            }
            let x = n
                .attrs
                .get("v")
                .ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing pivot record index")
                })?
                .parse::<usize>()
                .map_err(err)?;
            let v = shared[c].get(x).ok_or_else(|| {
                crate::coded("pivot-limit", "Pivot shared index is out of bounds")
            })?;
            if v != cells
                .get(&(a.0 + i as u32 + 1, a.1 + c as u32))
                .unwrap_or(&Value::Null)
            {
                return Err(crate::coded(
                    "internal-mismatch",
                    "Pivot cached record differs from accepted literal source",
                ));
            }
            let expected = if let Some(at) = seen[c].iter().position(|s| s == v) {
                at
            } else {
                seen[c].push(v.clone());
                seen[c].len() - 1
            };
            if x != expected {
                return Err(crate::coded(
                    "pivot-unsupported",
                    "Pivot shared items are not canonical source order",
                ));
            }
            indexes.push(x);
            values.push(v);
        }
        let label = values[row]
            .as_str()
            .filter(|s| !s.is_empty())
            .ok_or_else(|| crate::coded("pivot-bad-input", "Pivot row label requires text"))?
            .to_owned();
        let fold = label.to_lowercase();
        if labels.get(&fold).is_some_and(|s| s != &label) {
            return Err(crate::coded(
                "pivot-conflict",
                "Case-colliding pivot labels are protected",
            ));
        }
        labels.insert(fold, label.clone());
        let n = if values[value].is_null() {
            0
        } else {
            values[value].as_i64().ok_or_else(|| {
                crate::coded(
                    "pivot-bad-input",
                    "Pivot aggregate requires integer or blank",
                )
            })?
        };
        let amount = if values[value].is_null() {
            0
        } else if aggregate == "count" {
            1
        } else {
            n
        };
        let at = if let Some(at) = groups.iter().position(|g| g.0 == label) {
            at
        } else {
            if groups.len() >= 512 {
                return Err(crate::coded(
                    "pivot-limit",
                    "Pivot groups exceed their bound",
                ));
            }
            groups.push((label, indexes[row], 0));
            groups.len() - 1
        };
        groups[at].2 = groups[at]
            .2
            .checked_add(amount)
            .filter(|n| n.unsigned_abs() <= MAX_SAFE as u64)
            .ok_or_else(|| crate::coded("pivot-limit", "Pivot group overflow"))?;
        total = total
            .checked_add(amount)
            .filter(|n| n.unsigned_abs() <= MAX_SAFE as u64)
            .ok_or_else(|| crate::coded("pivot-limit", "Pivot total overflow"))?;
    }
    if seen != shared {
        return Err(crate::coded(
            "pivot-unsupported",
            "Unused or unqualified pivot shared items",
        ));
    }
    for (i, field) in defs.nodes[fields].children.iter().enumerate() {
        let si = child(&defs, *field, "sharedItems")?;
        let types = shared[i]
            .iter()
            .filter(|v| !v.is_null())
            .map(|v| if v.is_number() { "number" } else { "string" })
            .collect::<BTreeSet<_>>();
        for (k, v) in [
            ("containsBlank", shared[i].iter().any(Value::is_null)),
            ("containsString", types.contains("string")),
            ("containsNumber", types.contains("number")),
            ("containsInteger", types.contains("number")),
            ("containsDate", false),
            ("containsNonDate", true),
            (
                "containsSemiMixedTypes",
                types.contains("string") || shared[i].iter().any(Value::is_null),
            ),
            ("containsMixedTypes", types.len() > 1),
        ] {
            if defs.nodes[si].attrs.get(k).map(String::as_str) != Some(if v { "1" } else { "0" }) {
                return Err(crate::coded(
                    "pivot-stale",
                    "Pivot shared type flags changed",
                ));
            }
        }
    }
    attrs(
        &table,
        0,
        &[
            "name",
            "cacheId",
            "dataCaption",
            "grandTotalCaption",
            "rowGrandTotals",
            "colGrandTotals",
            "compact",
            "compactData",
            "outline",
            "outlineData",
            "gridDropZones",
            "updatedVersion",
            "minRefreshableVersion",
            "createdVersion",
            "useAutoFormatting",
        ],
    )?;
    if table.nodes[0].attrs.get("name").map(String::as_str) != Some(name)
        || table.nodes[0].attrs.get("cacheId").map(String::as_str) != Some("1")
        || table.nodes[0].children.iter().any(|i| {
            ![
                "location",
                "pivotFields",
                "rowFields",
                "rowItems",
                "colItems",
                "dataFields",
                "pivotTableStyleInfo",
            ]
            .contains(&table.nodes[*i].name.as_str())
        })
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid pivot definition profile",
        ));
    }
    if table.nodes[0].children.len() != 7 {
        return Err(crate::coded(
            "pivot-unsupported",
            "Extended pivot definition",
        ));
    }
    for (k, v) in [
        ("dataCaption", "Values"),
        ("grandTotalCaption", "Grand Total"),
        ("rowGrandTotals", if presentation.grand { "1" } else { "0" }),
        ("colGrandTotals", if presentation.grand { "1" } else { "0" }),
        ("compact", presentation.compact),
        ("compactData", presentation.compact),
        ("outline", presentation.outline),
        ("outlineData", presentation.outline),
        ("gridDropZones", "0"),
        ("updatedVersion", "6"),
        ("minRefreshableVersion", "3"),
        ("createdVersion", "6"),
        ("useAutoFormatting", "0"),
    ] {
        if table.nodes[0].attrs.get(k).map(String::as_str) != Some(v) {
            return Err(crate::coded(
                "pivot-stale",
                "Pivot definition options changed",
            ));
        }
    }
    let location = child(&table, 0, "location")?;
    attrs(
        &table,
        location,
        &["ref", "firstHeaderRow", "firstDataRow", "firstDataCol"],
    )?;
    let item_count = groups.len() + usize::from(presentation.grand);
    let output = format!("A1:B{}", groups.len() + 1 + usize::from(presentation.grand));
    leaf(&table, location)?;
    if table.nodes[location].attrs.get("ref") != Some(&output) {
        return Err(crate::coded("pivot-stale", "Pivot output range changed"));
    }
    for (k, v) in [
        ("firstHeaderRow", "1"),
        ("firstDataRow", "1"),
        ("firstDataCol", "1"),
    ] {
        if table.nodes[location].attrs.get(k).map(String::as_str) != Some(v) {
            return Err(crate::coded("pivot-stale", "Pivot display offsets changed"));
        }
    }
    let pf = child(&table, 0, "pivotFields")?;
    exact(&table, pf, "pivotFields", &[("count", columns.to_string())])?;
    if table.nodes[pf].children.len() != columns as usize {
        return Err(crate::coded("pivot-stale", "Pivot field layout changed"));
    }
    for (i, f) in table.nodes[pf].children.iter().enumerate() {
        if i == row {
            exact(
                &table,
                *f,
                "pivotField",
                &[
                    ("axis", "axisRow".into()),
                    ("showAll", "0".into()),
                    ("defaultSubtotal", "0".into()),
                    ("compact", presentation.compact.into()),
                    ("outline", presentation.outline.into()),
                ],
            )?;
            let items = child(&table, *f, "items")?;
            exact(
                &table,
                items,
                "items",
                &[("count", groups.len().to_string())],
            )?;
            if table.nodes[*f].children.len() != 1
                || table.nodes[items].children.len() != groups.len()
            {
                return Err(crate::coded("pivot-stale", "Pivot row item count changed"));
            }
            for (j, item) in table.nodes[items].children.iter().enumerate() {
                exact(&table, *item, "item", &[("x", groups[j].1.to_string())])?;
                if !table.nodes[*item].children.is_empty() {
                    return Err(crate::coded("pivot-unsupported", "Extended pivot row item"));
                }
            }
        } else {
            let mut pairs = vec![("showAll", "0".into())];
            if i == value {
                pairs.push(("dataField", "1".into()));
            }
            exact(&table, *f, "pivotField", &pairs)?;
            if !table.nodes[*f].children.is_empty() {
                return Err(crate::coded(
                    "pivot-unsupported",
                    "Extended pivot unused field",
                ));
            }
        }
    }
    // A single value field has one implicit column item, with no row-label
    // cache index. Pivot linkage lives in OPC relationships, not worksheet XML.
    let ci = child(&table, 0, "colItems")?;
    exact(&table, ci, "colItems", &[("count", "1".into())])?;
    if table.nodes[ci].children.len() != 1 {
        return Err(crate::coded(
            "pivot-stale",
            "Pivot value column layout changed",
        ));
    }
    let column_item = table.nodes[ci].children[0];
    exact(&table, column_item, "i", &[])?;
    leaf(&table, column_item)?;
    let ri = child(&table, 0, "rowItems")?;
    exact(&table, ri, "rowItems", &[("count", item_count.to_string())])?;
    if table.nodes[ri].children.len() != item_count {
        return Err(crate::coded("pivot-stale", "Pivot row layout changed"));
    }
    for (i, item) in table.nodes[ri].children.iter().enumerate() {
        exact(
            &table,
            *item,
            "i",
            &if i == groups.len() {
                vec![("t", "grand".into())]
            } else {
                vec![]
            },
        )?;
        if table.nodes[*item].children.len() != 1 {
            return Err(crate::coded(
                "pivot-unsupported",
                "Extended pivot row layout",
            ));
        }
        let x = table.nodes[*item].children[0];
        exact(
            &table,
            x,
            "x",
            &if i == groups.len() {
                vec![]
            } else {
                vec![("v", i.to_string())]
            },
        )?;
        if !table.nodes[x].children.is_empty() {
            return Err(crate::coded(
                "pivot-unsupported",
                "Extended pivot row index",
            ));
        }
    }
    let rf = child(&table, 0, "rowFields")?;
    let rfield = table.nodes[rf]
        .children
        .first()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing row field"))?;
    if table.nodes[rf].children.len() != 1
        || table.nodes[*rfield].attrs.get("x").map(String::as_str) != Some(row.to_string().as_str())
    {
        return Err(crate::coded("pivot-stale", "Pivot row axis changed"));
    }
    let df = child(&table, 0, "dataFields")?;
    let data = table.nodes[df]
        .children
        .first()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing data field"))?;
    let caption = format!(
        "{} of {}",
        if aggregate == "sum" { "Sum" } else { "Count" },
        headers[value]
    );
    if table.nodes[df].children.len() != 1
        || table.nodes[*data].attrs.get("fld").map(String::as_str)
            != Some(value.to_string().as_str())
        || table.nodes[*data].attrs.get("subtotal").map(String::as_str) != Some(aggregate)
        || table.nodes[*data].attrs.get("name") != Some(&caption)
    {
        return Err(crate::coded("pivot-stale", "Pivot data axis changed"));
    }
    exact(&table, rf, "rowFields", &[("count", "1".into())])?;
    exact(&table, *rfield, "field", &[("x", row.to_string())])?;
    exact(&table, df, "dataFields", &[("count", "1".into())])?;
    exact(
        &table,
        *data,
        "dataField",
        &[
            ("name", caption.clone()),
            ("fld", value.to_string()),
            ("subtotal", aggregate.into()),
            ("baseField", "0".into()),
            ("baseItem", "0".into()),
            ("numFmtId", "0".into()),
        ],
    )?;
    let style = child(&table, 0, "pivotTableStyleInfo")?;
    exact(
        &table,
        style,
        "pivotTableStyleInfo",
        &[
            ("name", presentation.style.into()),
            ("showRowHeaders", "1".into()),
            ("showColHeaders", "1".into()),
            ("showRowStripes", presentation.stripes.into()),
            ("showColStripes", "0".into()),
            ("showLastColumn", "1".into()),
        ],
    )?;
    for i in [*rfield, *data, style] {
        leaf(&table, i)?;
    }
    attrs(&display, 0, &[])?;
    if display.nodes[0].children.len() != 5
        || display.nodes[0].children.iter().any(|i| {
            ![
                "dimension",
                "sheetViews",
                "sheetFormatPr",
                "cols",
                "sheetData",
            ]
            .contains(&display.nodes[*i].name.as_str())
        })
    {
        return Err(crate::coded(
            "pivot-unsupported",
            "Extended pivot worksheet metadata",
        ));
    }
    let dim = child(&display, 0, "dimension")?;
    attrs(&display, dim, &["ref"])?;
    let dimension = crate::bulk::bounds(
        display.nodes[dim]
            .attrs
            .get("ref")
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing display dimension"))?,
    )?;
    let new_output = crate::bulk::bounds(&output)?;
    if dimension.0 != (0, 0) || dimension.1.0 < new_output.1.0 || dimension.1.1 < new_output.1.1 {
        return Err(crate::coded(
            "pivot-refused",
            "Pivot output dimension omits owned cells",
        ));
    }
    leaf(&display, dim)?;
    let views = child(&display, 0, "sheetViews")?;
    exact(&display, views, "sheetViews", &[])?;
    if display.nodes[views].children.len() != 1 {
        return Err(crate::coded(
            "pivot-unsupported",
            "Extended pivot sheet view",
        ));
    }
    let view = display.nodes[views].children[0];
    exact(
        &display,
        view,
        "sheetView",
        &[("workbookViewId", "0".into())],
    )?;
    leaf(&display, view)?;
    let format = child(&display, 0, "sheetFormatPr")?;
    exact(
        &display,
        format,
        "sheetFormatPr",
        &[("defaultRowHeight", "15".into())],
    )?;
    leaf(&display, format)?;
    let cols = child(&display, 0, "cols")?;
    exact(&display, cols, "cols", &[])?;
    if display.nodes[cols].children.len() != 2 {
        return Err(crate::coded("pivot-stale", "Pivot columns changed"));
    }
    for (i, col) in display.nodes[cols].children.iter().enumerate() {
        exact(
            &display,
            *col,
            "col",
            &[
                ("min", (i + 1).to_string()),
                ("max", (i + 1).to_string()),
                ("width", if i == 0 { "24" } else { "20" }.into()),
                ("customWidth", "1".into()),
            ],
        )?;
        leaf(&display, *col)?;
    }
    let sd = child(&display, 0, "sheetData")?;
    exact(&display, sd, "sheetData", &[])?;
    let actual = output_cells(&display, new_output)?;
    if !refresh
        && (dimension != new_output
            || display.nodes[sd].children.len()
                != groups.len() + 1 + usize::from(presentation.grand)
            || display.nodes[sd]
                .children
                .iter()
                .any(|row| display.nodes[*row].children.len() != 2))
    {
        return Err(crate::coded(
            "pivot-refused",
            "Pivot creation has undeclared neighboring data",
        ));
    }
    let mut expected = BTreeMap::from([
        (
            "A1".into(),
            json!(if presentation.compact == "1" {
                "Row Labels"
            } else {
                headers[row].as_str()
            }),
        ),
        ("B1".into(), json!(caption)),
    ]);
    for (i, g) in groups.iter().enumerate() {
        expected.insert(address(i as u32 + 1, 0), json!(g.0));
        expected.insert(address(i as u32 + 1, 1), json!(g.2));
    }
    if presentation.grand {
        expected.insert(address(groups.len() as u32 + 1, 0), json!("Grand Total"));
        expected.insert(address(groups.len() as u32 + 1, 1), json!(total));
    }
    if actual != expected {
        return Err(crate::coded(
            "internal-mismatch",
            "Pivot displayed aggregate differs from accepted cache/source",
        ));
    }
    let replacements = plan["replacements"]
        .as_object()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing pivot replacements"))?;
    if refresh {
        if replacements.len() != 4 || plan["priorRevision"].as_str().is_none() {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid pivot refresh replacements",
            ));
        }
        let prior_target = Tree::parse(&part(&mut zip, target)?)?;
        let old_table = Tree::parse(&part(&mut zip, parts.table)?)?;
        let loc = child(&old_table, 0, "location")?;
        let old_output =
            crate::bulk::bounds(old_table.nodes[loc].attrs.get("ref").ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing prior location")
            })?)?;
        if target_mask(&prior_target, old_output, new_output)?
            != target_mask(&display, old_output, new_output)?
        {
            return Err(crate::coded(
                "pivot-stale",
                "Pivot refresh changed raw neighboring cells or worksheet metadata",
            ));
        }
        let sd = child(&prior_target, 0, "sheetData")?;
        for row in &prior_target.nodes[sd].children {
            for cell in &prior_target.nodes[*row].children {
                let p = position(prior_target.nodes[*cell].attrs.get("r").ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing prior coordinate")
                })?)?;
                let within = |range: ((u32, u32), (u32, u32))| {
                    p.0 >= range.0.0 && p.0 <= range.1.0 && p.1 >= range.0.1 && p.1 <= range.1.1
                };
                if within(new_output) && !within(old_output) {
                    return Err(crate::coded(
                        "pivot-conflict",
                        "Pivot growth intersects an occupied neighboring source cell",
                    ));
                }
            }
        }
        let before = [
            String::from_utf8(part(&mut zip, parts.table)?).map_err(err)?,
            String::from_utf8(part(&mut zip, parts.definition)?).map_err(err)?,
            String::from_utf8(part(&mut zip, parts.records)?).map_err(err)?,
        ];
        let digest = format!(
            "{:x}",
            Sha256::digest(serde_json::to_vec(&before).map_err(err)?)
        );
        if plan["priorRevision"] != digest {
            return Err(crate::coded("pivot-stale", "Conflict: pivot cache changed"));
        }
    } else {
        creation_graph(&mut zip, plan)?;
        if replacements.len() != 10
            || !plan["priorRevision"].is_null()
            || zip
                .file_names()
                .any(|n| [parts.table, parts.definition, parts.records, target].contains(&n))
        {
            return Err(crate::coded(
                "pivot-bad-input",
                "Pivot creation destination exists or plan is not isolated",
            ));
        }
        relation(
            plan,
            &relation_file(target),
            "rIdPivot",
            "pivotTable",
            workbook_format!(PIVOTTABLES_PIVOT1_XML),
        )?;
        relation(
            plan,
            &relation_file(parts.table),
            "rIdCache",
            "pivotCacheDefinition",
            workbook_format!(PIVOTCACHE_CACHEDEFINITION1_XML),
        )?;
        relation(
            plan,
            &relation_file(parts.definition),
            "rIdRecords",
            "pivotCacheRecords",
            workbook_format!(CACHERECORDS1_XML),
        )?;
        for (file, ns, root_name, tags, expected_count) in [
            (
                "xl/workbook.xml",
                NS,
                "workbook",
                vec!["sheet", "pivotCaches"],
                2usize,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                PACKAGE_REL_NS,
                "Relationships",
                vec!["Relationship"],
                2,
            ),
            ("[Content_Types].xml", TYPE_NS, "Types", vec!["Override"], 4),
        ] {
            let before = Tree::parse(&part(&mut zip, file)?)?;
            let after = Tree::parse(field(&plan["replacements"], file)?.as_bytes())?;
            before.root(ns, root_name)?;
            after.root(ns, root_name)?;
            let mut remove = Vec::new();
            for (i, n) in after.nodes.iter().enumerate().skip(1) {
                let added = if file == "xl/workbook.xml" {
                    n.name == "sheet"
                        && n.attrs.get("name").map(String::as_str)
                            == Some(field(plan, "targetSheet")?)
                        || n.name == "pivotCaches"
                } else if file.ends_with(".rels") {
                    n.name == "Relationship"
                        && n.attrs.get("Target").is_some_and(|v| {
                            v == target.strip_prefix("xl/").unwrap_or(target)
                                || v == parts
                                    .definition
                                    .strip_prefix("xl/")
                                    .unwrap_or(parts.definition)
                        })
                } else {
                    n.name == "Override"
                        && n.attrs.get("PartName").is_some_and(|v| {
                            [
                                format!("/{target}"),
                                format!("/{}", parts.table),
                                format!("/{}", parts.definition),
                                format!("/{}", parts.records),
                            ]
                            .contains(v)
                        })
                };
                if added && n.ns == ns && tags.contains(&n.name.as_str()) {
                    remove.push(i);
                }
            }
            if remove.len() != expected_count || envelope(&after, &remove) != before.bytes {
                return Err(crate::coded(
                    "pivot-stale",
                    "Pivot edit changed unrelated package metadata",
                ));
            }
        }
    }
    Ok(())
}

#[cfg(test)]
mod numeric_guard_tests {
    use super::*;
    #[test]
    fn cached_numeric_bounds_refuse_minimum_i64_without_abs_overflow() {
        for value in [i64::MIN, i64::MAX, MAX_SAFE + 1, -MAX_SAFE - 1] {
            assert!(safe_integer(&value.to_string()).is_err());
        }
        for value in [-MAX_SAFE, 0, MAX_SAFE] {
            assert_eq!(safe_integer(&value.to_string()).unwrap(), value);
        }
        assert!(safe_integer("1.5").is_err());
    }
}
/// Qualify only the stored package/model shape. Cached XML never grants current
/// source values: ordinary refresh separately evaluates the accepted SQL cells.
fn cached_profile(root: &Path) -> Result<Option<(Value, BTreeMap<(u32, u32), Value>)>> {
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "source.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let parts = Parts::from_zip(&zip)?;
    if !zip.file_names().any(|name| name == parts.table) {
        return Ok(None);
    }
    let table = Tree::parse(&part(&mut zip, parts.table)?)?;
    table.root(NS, "pivotTableDefinition")?;
    let definition = Tree::parse(&part(&mut zip, parts.definition)?)?;
    definition.root(NS, "pivotCacheDefinition")?;
    let cache = child(&definition, 0, "cacheSource")?;
    let ws = child(&definition, cache, "worksheetSource")?;
    let source_sheet = definition.nodes[ws]
        .attrs
        .get("sheet")
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing owned source sheet"))?
        .clone();
    let range = definition.nodes[ws]
        .attrs
        .get("ref")
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing owned source range"))?
        .clone();
    let (a, b) = crate::bulk::bounds(&range)?;
    let source_view = crate::package_edit::auditing_metadata(root)?;
    let source_path = source_view["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing source sheets"))?
        .iter()
        .find(|sheet| sheet["name"] == source_sheet)
        .and_then(|sheet| sheet["path"].as_str())
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing pivot source identity"))?
        .to_owned();
    let index: Value =
        serde_json::from_slice(&fs::read(local(root, "index.json")?).map_err(err)?).map_err(err)?;
    let mut target_path = None;
    for sheet in index["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing indexed sheets"))?
    {
        let path = field(sheet, "path")?;
        let rel = relation_file(path);
        let bytes = match zip.by_name(&rel) {
            Ok(mut entry) => {
                let mut bytes = Vec::new();
                entry.read_to_end(&mut bytes).map_err(err)?;
                bytes
            }
            Err(zip::result::ZipError::FileNotFound) => continue,
            Err(error) => return Err(err(error)),
        };
        let rels = Tree::parse(&bytes)?;
        for node in &rels.nodes[0].children {
            let node = &rels.nodes[*node];
            if node.attrs.get("Type").map(String::as_str)
                == Some(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable",
                )
                && node.attrs.get("Target").is_some_and(|target| {
                    super::resolve(path, target).is_ok_and(|target| target == parts.table)
                })
            {
                if target_path.replace(path.to_owned()).is_some() {
                    return Err(crate::coded(
                        "pivot-unsupported",
                        "Shared owned pivot target is unqualified",
                    ));
                }
            }
        }
    }
    let target_path = target_path
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing owned pivot target"))?;
    let target_sheet = source_view["sheets"]
        .as_array()
        .unwrap()
        .iter()
        .find(|sheet| sheet["path"] == target_path)
        .and_then(|sheet| sheet["name"].as_str())
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing pivot target name"))?
        .to_owned();
    let fields = child(&definition, 0, "cacheFields")?;
    let mut dataset = BTreeMap::new();
    let mut items = Vec::new();
    for (c, field_node) in definition.nodes[fields].children.iter().enumerate() {
        let field_node = &definition.nodes[*field_node];
        dataset.insert(
            (a.0, a.1 + c as u32),
            json!(
                field_node.attrs.get("name").ok_or_else(|| crate::coded(
                    "internal-missing-field",
                    "Missing cached header"
                ))?
            ),
        );
        let shared = child(
            &definition,
            *definition.nodes[fields].children.get(c).unwrap(),
            "sharedItems",
        )?;
        let values = definition.nodes[shared]
            .children
            .iter()
            .map(|item| {
                let n = &definition.nodes[*item];
                match n.name.as_str() {
                    "n" => Ok(json!(safe_integer(n.attrs.get("v").ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing cached integer")
                    })?)?)),
                    "s" => Ok(json!(n.attrs.get("v").ok_or_else(|| crate::coded(
                        "internal-missing-field",
                        "Missing cached text"
                    ))?)),
                    "b" => Ok(json!(n.attrs.get("v").map(String::as_str) == Some("1"))),
                    "m" => Ok(Value::Null),
                    _ => Err(crate::coded(
                        "pivot-unsupported",
                        "Unsupported cached source field",
                    )),
                }
            })
            .collect::<Result<Vec<_>>>()?;
        items.push(values);
    }
    let records = Tree::parse(&part(&mut zip, parts.records)?)?;
    records.root(NS, "pivotCacheRecords")?;
    for (r, record) in records.nodes[0].children.iter().enumerate() {
        for (c, cell) in records.nodes[*record].children.iter().enumerate() {
            let item = records.nodes[*cell]
                .attrs
                .get("v")
                .ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing cached source index")
                })?
                .parse::<usize>()
                .map_err(err)?;
            dataset.insert(
                (a.0 + r as u32 + 1, a.1 + c as u32),
                items
                    .get(c)
                    .and_then(|items| items.get(item))
                    .ok_or_else(|| {
                        crate::coded(
                            "pivot-unsupported",
                            "Cached source index is outside its field",
                        )
                    })?
                    .clone(),
            );
        }
    }
    if records.nodes[0].children.len() != (b.0 - a.0) as usize
        || items.len() != (b.1 - a.1 + 1) as usize
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Owned pivot cached source geometry differs",
        ));
    }
    let row_fields = child(&table, 0, "rowFields")?;
    let row = table.nodes[*table.nodes[row_fields]
        .children
        .first()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing cached row axis"))?]
    .attrs
    .get("x")
    .ok_or_else(|| crate::coded("internal-missing-field", "Missing row index"))?
    .parse::<usize>()
    .map_err(err)?;
    let data_fields = child(&table, 0, "dataFields")?;
    let data = &table.nodes[*table.nodes[data_fields]
        .children
        .first()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing cached aggregate"))?];
    let value = data
        .attrs
        .get("fld")
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing value index"))?
        .parse::<usize>()
        .map_err(err)?;
    let aggregate = data
        .attrs
        .get("subtotal")
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing aggregate"))?
        .clone();
    let name = table.nodes[0]
        .attrs
        .get("name")
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing pivot name"))?
        .clone();
    let originals = [
        String::from_utf8(part(&mut zip, parts.table)?).map_err(err)?,
        String::from_utf8(part(&mut zip, parts.definition)?).map_err(err)?,
        String::from_utf8(part(&mut zip, parts.records)?).map_err(err)?,
    ];
    let revision = format!(
        "{:x}",
        Sha256::digest(serde_json::to_vec(&originals).map_err(err)?)
    );
    let mut replacements = serde_json::Map::new();
    for part_name in [
        parts.table,
        parts.definition,
        parts.records,
        target_path.as_str(),
    ] {
        replacements.insert(
            part_name.into(),
            json!(String::from_utf8(part(&mut zip, part_name)?).map_err(err)?),
        );
    }
    let opts = table_presentation(&table)?;
    let plan = json!({"sample":opts.sample,"grandTotals":opts.grand,"kind":"refreshPivot","name":name,"sourceSheet":source_sheet,"sourcePath":source_path,"sourceRange":range,"rowField":row,"valueField":value,"aggregate":aggregate,"targetSheet":target_sheet,"targetPath":target_path,"cacheId":1,"parts":{"table":parts.table,"definition":parts.definition,"records":parts.records},"priorRevision":revision,"replacements":replacements,"removed":[],"copies":{},"fragments":{}});
    validate_inner(
        root,
        &crate::database(root)?,
        &plan,
        &json!({"version":0,"source":"source.xlsx"}),
        Some(&dataset),
    )?;
    Ok(Some((plan, dataset)))
}

pub(crate) fn cached_binding(root: &Path) -> Result<Option<Value>> {
    Ok(cached_profile(root)?.map(|(plan,_)|json!({"object":plan["parts"]["table"],"sourceSheet":plan["sourceSheet"],"sourcePath":plan["sourcePath"],"range":plan["sourceRange"],"targetSheet":plan["targetSheet"],"targetPath":plan["targetPath"],"name":plan["name"],"revision":plan["priorRevision"]})))
}
