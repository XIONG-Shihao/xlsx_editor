//! Table-local column structure: prove exact definition changes before streaming cells.
use super::*;
#[derive(Clone)]
pub(crate) struct Change {
    pub first: (u32, u32),
    pub last: (u32, u32),
    pub index: usize,
    pub remove: bool,
    pub label: String,
}
pub(crate) fn change(plan: &Value) -> Result<Option<Change>> {
    if !["insertTableColumn", "deleteTableColumn"].contains(&field(plan, "kind")?) {
        return Ok(None);
    }
    let c = &plan["columnChange"];
    if c.as_object().is_none_or(|v| v.len() != 3) {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid table column change fields",
        ));
    }
    let index = usize::try_from(
        c["index"]
            .as_u64()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing table column index"))?,
    )
    .map_err(err)?;
    let remove = c["remove"]
        .as_bool()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing table column direction"))?;
    if remove != (plan["kind"] == "deleteTableColumn") {
        return Err(crate::coded(
            "internal-mismatch",
            "Table column direction mismatch",
        ));
    }
    let label = field(c, "label")?;
    if label.trim().is_empty()
        || label.trim() != label
        || label.len() > 255
        || label
            .chars()
            .any(|c| c.is_control() || ['[', ']', ':', '\'', ',', '#', '@'].contains(&c))
    {
        return Err(crate::coded(
            "table-bad-input",
            "Use an ordinary table column label",
        ));
    }
    let (first, last) = crate::bulk::bounds(field(plan, "priorRef")?)?;
    Ok(Some(Change {
        first,
        last,
        index,
        remove,
        label: label.into(),
    }))
}
fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}
struct Patch {
    start: usize,
    end: usize,
    text: Vec<u8>,
}
fn attribute(tree: &Tree, index: usize, key: &str, value: &str) -> Result<Patch> {
    let n = &tree.nodes[index];
    let head = &tree.bytes[n.start..n.end];
    let close = head
        .iter()
        .position(|c| *c == b'>')
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing table opening"))?;
    let head = &head[..close];
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
        let name = std::str::from_utf8(&head[start..end]).map_err(err)?;
        at = equals + 1;
        while at < head.len() && head[at].is_ascii_whitespace() {
            at += 1;
        }
        let quote = *head
            .get(at)
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing attribute quote"))?;
        if ![b'\'', b'"'].contains(&quote) {
            return Err(crate::coded(
                "table-malformed",
                "Invalid table attribute quote",
            ));
        }
        let from = at + 1;
        at = from;
        while at < head.len() && head[at] != quote {
            at += 1;
        }
        if at == head.len() {
            return Err(crate::coded(
                "table-malformed",
                "Unterminated table attribute",
            ));
        }
        if name == key {
            return Ok(Patch {
                start: n.start + from,
                end: n.start + at,
                text: escape(value).into_bytes(),
            });
        }
        at += 1;
    }
    Err(crate::coded(
        "internal-missing-field",
        "Missing table column structure attribute",
    ))
}
pub(super) fn definition(
    before: &Tree,
    after: &Tree,
    old: &Model,
    new: &Model,
    plan: &Value,
) -> Result<Change> {
    let c = change(plan)?
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing table column structure"))?;
    let count = old.columns.len();
    if c.index > count || c.remove && (c.index >= count || count == 1) {
        return Err(crate::coded(
            "table-bad-input",
            "Keep at least one table column and select a valid index",
        ));
    }
    let mut expected = old.columns.clone();
    if c.remove {
        if expected[c.index] != c.label {
            return Err(crate::coded("table-stale", "Stale table column label"));
        }
        expected.remove(c.index);
    } else {
        if expected.iter().any(|v| v.eq_ignore_ascii_case(&c.label)) {
            return Err(crate::coded(
                "table-conflict",
                "Table column label collision",
            ));
        }
        expected.insert(c.index, c.label.clone());
    }
    let last = if c.remove {
        c.last.1.checked_sub(1)
    } else {
        c.last.1.checked_add(1)
    }
    .filter(|v| *v < 16384)
    .ok_or_else(|| crate::coded("table-unsupported", "Table column outside worksheet"))?;
    let range = format!(
        "{}:{}",
        crate::xml::address(c.first.0, c.first.1),
        crate::xml::address(c.last.0, last)
    );
    if new.name != old.name
        || new.reference != range
        || new.columns != expected
        || plan["range"] != range
        || plan.get("renameTable").is_some()
        || plan.get("renameTableColumn").is_some()
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid table column definition change",
        ));
    }
    let container = direct(before, 0, "tableColumns", NS)?[0];
    let entries = &before.nodes[container].children;
    let mut patches = vec![
        attribute(before, 0, "ref", &range)?,
        attribute(before, container, "count", &expected.len().to_string())?,
    ];
    for filter in direct(before, 0, "autoFilter", NS)? {
        patches.push(attribute(before, filter, "ref", &range)?);
    }
    if c.remove {
        let n = &before.nodes[entries[c.index]];
        patches.push(Patch {
            start: n.start,
            end: n.end,
            text: Vec::new(),
        });
    } else {
        let id = entries
            .iter()
            .map(|i| integer(attr(before, *i, "id")?))
            .collect::<Result<Vec<_>>>()?
            .into_iter()
            .max()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing table column ids"))?
            .checked_add(1)
            .filter(|id| *id <= 9007199254740991)
            .ok_or_else(|| crate::coded("table-limit", "Table column id exceeds its bound"))?;
        let at = if c.index < count {
            before.nodes[entries[c.index]].start
        } else {
            before.nodes[*entries.last().unwrap()].end
        };
        patches.push(Patch {
            start: at,
            end: at,
            text: format!(
                "<{} id=\"{}\" name=\"{}\"/>",
                before.nodes[entries[0]].qname,
                id,
                escape(&c.label)
            )
            .into_bytes(),
        });
    }
    let mut bytes = before.bytes.clone();
    for patch in {
        patches.sort_by_key(|p| std::cmp::Reverse(p.start));
        patches
    } {
        bytes.splice(patch.start..patch.end, patch.text);
    }
    if bytes != after.bytes {
        return Err(crate::coded(
            "table-stale",
            "Table column edit changed unrelated raw metadata",
        ));
    }
    Ok(c)
}
pub(super) fn cells(
    root: &Path,
    db: &Connection,
    plan: &Value,
    input: &Value,
    c: &Change,
    all_models: &[(String, String)],
    sheet_tree: &Tree,
) -> Result<()> {
    let sheet = field(plan, "sheet")?;
    let new_range = field(plan, "range")?;
    let (a, b) = crate::bulk::bounds(new_range)?;
    if all_models
        .iter()
        .filter(|(p, r)| p == sheet && r == field(plan, "priorRef").unwrap_or(""))
        .count()
        != 1
    {
        return Err(crate::coded(
            "table-conflict",
            "Overlapping table identity is protected",
        ));
    }
    for (path, reference) in all_models {
        if path == sheet && reference != field(plan, "priorRef")? {
            let (e, f) = crate::bulk::bounds(reference)?;
            if a.0 <= f.0 && b.0 >= e.0 && a.1 <= f.1 && b.1 >= e.1 {
                return Err(crate::coded(
                    "table-conflict",
                    "Table column change intersects another table",
                ));
            }
        }
    }
    for n in &sheet_tree.nodes {
        if n.ns == NS && n.name == "mergeCell" {
            let (e, f) = crate::bulk::bounds(
                n.attrs
                    .get("ref")
                    .ok_or_else(|| crate::coded("internal-missing-field", "Missing merge ref"))?,
            )?;
            if c.first.0 <= f.0 && c.last.0 >= e.0 && c.first.1 <= f.1 && c.last.1.max(b.1) >= e.1 {
                return Err(crate::coded(
                    "table-conflict",
                    "Table column change intersects a merge",
                ));
            }
        }
    }
    let version = input["version"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing table version"))?;
    let check = |first: (u32, u32), last: (u32, u32), growth: bool| -> Result<()> {
        let rows = (1000 / (last.1 - first.1 + 1)).max(1);
        let mut row = first.0;
        while row <= last.0 {
            let end = (row + rows - 1).min(last.0);
            let page = crate::range::read_db(
                db,
                &json!({"version":version,"sheet":sheet,"range":format!("{}:{}",crate::xml::address(row,first.1),crate::xml::address(end,last.1)),"limit":1001}),
            )?;
            if !page["nextOffset"].is_null() {
                return Err(crate::coded(
                    "table-unsupported",
                    "Incomplete table column cells",
                ));
            }
            for record in page["cells"].as_array().ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing table column cells")
            })? {
                let cell = &record[1];
                if growth
                    && (cell.get("formula").is_some()
                        || cell.get("formulaType").is_some()
                        || cell.get("spill").is_some()
                        || !cell["value"].is_null())
                {
                    return Err(crate::coded(
                        "table-unsupported",
                        "Table growth cannot absorb existing content or spill ownership",
                    ));
                }
                ordinary_cell_attributes(
                    root,
                    db,
                    sheet,
                    position(record[0].as_str().ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing table cell position")
                    })?)?,
                )?;
            }
            row = end + 1;
        }
        Ok(())
    };
    check(c.first, c.last, false)?;
    if !c.remove {
        check((c.first.0, b.1), (c.last.0, b.1), true)?;
    }
    Ok(())
}
pub(super) fn header_style(db: &Connection, headers: &Value, c: &Change) -> Result<()> {
    if !c.remove {
        let values = headers["cells"]
            .as_array()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing headers"))?;
        let source = &values[c.index.min(values.len() - 1)][1];
        let style = source["styleId"].as_i64();
        db.execute_batch("CREATE TEMP TABLE table_column_header_style(style INTEGER);")
            .map_err(err)?;
        db.execute(
            "INSERT INTO table_column_header_style VALUES(?1)",
            params![style],
        )
        .map_err(err)?;
    }
    Ok(())
}
pub(crate) fn header(db: &Connection, c: &Change) -> Result<Vec<u8>> {
    let style: Option<i64> = db
        .prepare_cached("SELECT style FROM table_column_header_style")
        .map_err(crate::err)?
        .query_row([], |row| row.get(0))
        .map_err(err)?;
    Ok(format!(
        "<c xmlns=\"{}\" r=\"{}\" t=\"inlineStr\"{}><is><t xml:space=\"preserve\">{}</t></is></c>",
        String::from_utf8_lossy(NS),
        crate::xml::address(c.first.0, c.first.1 + c.index as u32),
        style.map(|s| format!(" s=\"{s}\"")).unwrap_or_default(),
        escape(&c.label)
    )
    .into_bytes())
}
pub(super) fn formula(
    formula: &str,
    old: &Model,
    plan: &Value,
    c: &Change,
    path: &str,
    reference: &str,
    sheet_names: &BTreeMap<String, String>,
) -> Result<String> {
    let point = position(reference)?;
    let inside = path == field(plan, "sheet")?
        && point.0 > c.first.0
        && point.0 <= c.last.0
        && point.1 >= c.first.1
        && point.1 <= c.last.1;
    let value = if c.remove {
        crate::formula::rename_table_column(formula, &old.name, &c.label, "", inside)?
    } else {
        formula.into()
    };
    let owner = sheet_names
        .get(path)
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing formula owner name"))?;
    let target = sheet_names
        .get(field(plan, "sheet")?)
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing table sheet name"))?;
    let result = crate::formula::rewrite(
        &json!({"items":[{"formula":value,"owner":owner,"reference":reference}],"axis":"column","remove":c.remove,"target":target,"index":c.first.1+c.index as u32,"count":1,"tableRange":old.reference}),
    )?;
    Ok(result["formulas"][0]
        .as_str()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing table column formula"))?
        .into())
}

#[cfg(test)]
mod tests {
    use super::*;
    fn table(columns: &str, count: usize, range: &str) -> Tree {
        Tree::parse(format!("<table xmlns=\"{}\" id=\"1\" name=\"Prices\" displayName=\"Prices\" ref=\"{}\"><autoFilter ref=\"{}\"/><!--keep--><tableColumns count=\"{}\">{}</tableColumns><tableStyleInfo name=\"TableStyleMedium2\" showRowStripes=\"1\"/></table>",String::from_utf8_lossy(NS),range,range,count,columns).as_bytes()).unwrap()
    }
    const A: &str = "<tableColumn id=\"1\" name=\"Value\"/>";
    const B: &str = "<tableColumn id=\"5\" name=\"Cost\"/>";
    const C: &str = "<tableColumn id=\"9\" name=\"Total\"/>";
    #[test]
    fn definition_admits_each_insert_position_and_preserves_existing_ids() {
        let before = table(&format!("{A}{B}{C}"), 3, "A1:C4");
        let old = model(&before).unwrap();
        let parts = [A, B, C];
        for index in [0, 1, 3] {
            let mut next = parts.to_vec();
            next.insert(index, "<tableColumn id=\"10\" name=\"Tax &amp; Fee\"/>");
            let after = table(&next.join(""), 4, "A1:D4");
            let plan = json!({"kind":"insertTableColumn","priorRef":"A1:C4","range":"A1:D4","columnChange":{"index":index,"remove":false,"label":"Tax & Fee"}});
            definition(&before, &after, &old, &model(&after).unwrap(), &plan).unwrap();
        }
    }
    #[test]
    fn definition_admits_each_remove_position_but_refuses_unrelated_deltas() {
        let before = table(&format!("{A}{B}{C}"), 3, "A1:C4");
        let old = model(&before).unwrap();
        let parts = [A, B, C];
        for (index, label) in ["Value", "Cost", "Total"].iter().enumerate() {
            let mut next = parts.to_vec();
            next.remove(index);
            let after = table(&next.join(""), 2, "A1:B4");
            let plan = json!({"kind":"deleteTableColumn","priorRef":"A1:C4","range":"A1:B4","columnChange":{"index":index,"remove":true,"label":label}});
            definition(&before, &after, &old, &model(&after).unwrap(), &plan).unwrap();
            let bad = Tree::parse(&after.bytes.iter().copied().collect::<Vec<_>>()).unwrap();
            let replaced = Tree::parse(
                String::from_utf8(bad.bytes)
                    .unwrap()
                    .replace("TableStyleMedium2", "TableStyleMedium9")
                    .as_bytes(),
            )
            .unwrap();
            assert!(
                definition(&before, &replaced, &old, &model(&replaced).unwrap(), &plan).is_err()
            );
        }
    }
}
