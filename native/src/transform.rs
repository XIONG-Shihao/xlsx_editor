//! Native row ordering, visibility and unreferenced row/column transforms.
//! Cell XML streams through unchanged except its address. Sort spools rows to
//! SQLite, retaining one row at a time rather than decoding the cell graph.
use crate::bulk::{bounds, literal};
use crate::range::read_public_db as read_db;
use crate::xml::{NS, address, attrs, position};
use crate::{Result, database, err, field, local};
use quick_xml::{
    NsReader, Writer,
    events::{BytesEnd, BytesStart, BytesText, Event},
};
use rusqlite::{OptionalExtension, params};
use serde_json::{Value, json};
use std::{
    collections::BTreeMap,
    fs::{self, File},
    io::{BufReader, Read, Write},
    path::Path,
};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};
mod owned_filter;
fn element(e: &BytesStart<'_>, changes: &[(&str, Option<String>)]) -> Result<BytesStart<'static>> {
    let mut attributes = attrs(e)?;
    for (k, v) in changes {
        if let Some(v) = v {
            attributes.insert((*k).into(), v.clone());
        } else {
            attributes.remove(*k);
        }
    }
    let mut out = BytesStart::new(String::from_utf8_lossy(e.name().as_ref()).into_owned());
    for (k, v) in &attributes {
        out.push_attribute((k.as_str(), v.as_str()));
    }
    Ok(out)
}
fn text(value: &Value) -> String {
    match value {
        Value::Null => String::new(),
        Value::String(s) => s.clone(),
        _ => value.to_string(),
    }
}
fn sort_destination(db: &rusqlite::Connection, row: u32) -> Result<Option<u32>> {
    db.prepare_cached("SELECT destination FROM sort_order WHERE source=?1")
        .map_err(err)?
        .query_row([row], |r| r.get(0))
        .optional()
        .map_err(err)
}
fn same(a: &Value, b: &Value) -> bool {
    if let (Some(a), Some(b)) = (a.as_f64(), b.as_f64()) {
        a == b && (a != 0.0 || a.is_sign_negative() == b.is_sign_negative())
    } else {
        a == b
    }
}
fn filter_xml(input: &Value, owned: &owned_filter::Prepared) -> Result<Vec<u8>> {
    if owned.remove {
        return Ok(Vec::new());
    }
    let mut writer = Writer::new(Vec::new());
    let mut filter = BytesStart::new("autoFilter");
    filter.push_attribute(("xmlns", std::str::from_utf8(NS).unwrap()));
    filter.push_attribute(("ref", field(input, "range")?));
    writer.write_event(Event::Start(filter)).map_err(err)?;
    let values = input["values"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing filter values"))?;
    // Hidden or shown dropdown buttons and the "sorted by" record are not
    // criteria; they stay with the filter whatever is filtered.
    let button = |writer: &mut Writer<Vec<u8>>, column: u32| -> Result<()> {
        if let Some((_, kept)) = owned.buttons.iter().find(|(at, _)| *at == column) {
            let mut element = BytesStart::new("filterColumn");
            element.push_attribute(("colId", column.to_string().as_str()));
            for (key, value) in kept {
                element.push_attribute((key.as_str(), value.as_str()));
            }
            writer.write_event(Event::Empty(element)).map_err(err)?;
        }
        Ok(())
    };
    let mut passive = owned.buttons.iter().map(|(at, _)| *at).collect::<Vec<_>>();
    passive.sort_unstable();
    if values.is_empty() {
        for at in &passive {
            button(&mut writer, *at)?;
        }
        writer.get_mut().extend_from_slice(&owned.sorted);
        append_owned_filter(writer.get_mut(), owned)?;
        writer
            .write_event(Event::End(BytesEnd::new("autoFilter")))
            .map_err(err)?;
        return Ok(writer.into_inner());
    }
    let ((_, first), _) = bounds(field(input, "range")?)?;
    let col = position(&format!("{}1", field(input, "column")?))?.1;
    for at in passive.iter().filter(|at| **at < col - first) {
        button(&mut writer, *at)?;
    }
    let mut column = BytesStart::new("filterColumn");
    column.push_attribute(("colId", (col - first).to_string().as_str()));
    if let Some((_, kept)) = owned.buttons.iter().find(|(at, _)| *at == col - first) {
        for (key, value) in kept {
            column.push_attribute((key.as_str(), value.as_str()));
        }
    }
    writer.write_event(Event::Start(column)).map_err(err)?;
    let mut filters = BytesStart::new("filters");
    if values.iter().any(|v| v.is_null() || v == "") {
        filters.push_attribute(("blank", "1"));
    }
    writer.write_event(Event::Start(filters)).map_err(err)?;
    for value in input["values"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing filter values"))?
    {
        if value.is_null() || value == "" {
            continue;
        }
        let mut item = BytesStart::new("filter");
        let text = if value.is_boolean() {
            if value == true { "1" } else { "0" }.into()
        } else {
            text(value)
        };
        item.push_attribute(("val", text.as_str()));
        writer.write_event(Event::Empty(item)).map_err(err)?;
    }
    for name in ["filters", "filterColumn"] {
        writer
            .write_event(Event::End(BytesEnd::new(name)))
            .map_err(err)?;
    }
    for at in passive.iter().filter(|at| **at > col - first) {
        button(&mut writer, *at)?;
    }
    writer.get_mut().extend_from_slice(&owned.sorted);
    append_owned_filter(writer.get_mut(), owned)?;
    writer
        .write_event(Event::End(BytesEnd::new("autoFilter")))
        .map_err(err)?;
    Ok(writer.into_inner())
}
fn append_owned_filter(bytes: &mut Vec<u8>, owned: &owned_filter::Prepared) -> Result<()> {
    let mut w = Writer::new(Vec::new());
    w.write_event(Event::Start(BytesStart::new("extLst")))
        .map_err(err)?;
    let mut ext = BytesStart::new("ext");
    ext.push_attribute(("uri", owned_filter::URI));
    w.write_event(Event::Start(ext)).map_err(err)?;
    let mut state = BytesStart::new(workbook_format!(FILTERSTATE));
    state.push_attribute((workbook_format!(XMLNS), owned_filter::URI));
    state.push_attribute(("revision", owned.revision.as_str()));
    w.write_event(Event::Start(state)).map_err(err)?;
    w.write_event(Event::Text(BytesText::new(&owned.text)))
        .map_err(err)?;
    for name in [workbook_format!(FILTERSTATE), "ext", "extLst"] {
        w.write_event(Event::End(BytesEnd::new(name)))
            .map_err(err)?;
    }
    bytes.extend_from_slice(&w.into_inner());
    Ok(())
}
fn table_filter<R: std::io::BufRead>(
    source: R,
    input: &Value,
    owned: &owned_filter::Prepared,
) -> Result<Vec<u8>> {
    let mut reader = NsReader::from_reader(source);
    let mut writer = Writer::new(Vec::new());
    let mut buffer = Vec::new();
    let (mut depth, mut skip, mut inserted, mut filters, mut columns) = (0, 0, false, 0, 0);
    loop {
        crate::cancel::check()?;
        let (ns, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        let main = ns == quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(NS));
        match &event {
            Event::Start(e) | Event::Empty(e) => {
                let empty = matches!(event, Event::Empty(_));
                if skip > 0 {
                    if !empty {
                        skip += 1;
                    }
                    buffer.clear();
                    continue;
                }
                let name = e.local_name();
                if depth == 0
                    && (!main
                        || name.as_ref() != b"table"
                        || attrs(e)?.get("ref").map(String::as_str) != input["range"].as_str())
                {
                    return Err(crate::coded(
                        "filter-bad-input",
                        "Filter must match the complete table range",
                    ));
                }
                if main && depth == 1 && name.as_ref() == b"autoFilter" {
                    filters += 1;
                    if filters > 1
                        || attrs(e)?.get("ref").map(String::as_str) != input["range"].as_str()
                    {
                        return Err(crate::coded(
                            "filter-unsupported",
                            "Replacing unrelated filters is unsupported",
                        ));
                    }
                    if !empty {
                        skip = 1;
                    }
                    buffer.clear();
                    continue;
                }
                if main && depth == 1 && name.as_ref() == b"tableColumns" {
                    columns += 1;
                    if columns > 1 {
                        return Err(crate::coded(
                            "internal-invalid-field",
                            "Invalid table columns",
                        ));
                    }
                }
                if main
                    && depth == 1
                    && !inserted
                    && [b"sortState".as_slice(), b"tableColumns"].contains(&name.as_ref())
                {
                    {
                        writer
                            .get_mut()
                            .write_all(&filter_xml(input, owned)?)
                            .map_err(err)?;
                    }
                    inserted = true;
                }
                if !empty {
                    depth += 1;
                }
            }
            Event::End(_) => {
                if skip > 0 {
                    skip -= 1;
                    buffer.clear();
                    continue;
                }
                depth -= 1;
            }
            Event::Eof => break,
            _ if skip > 0 => {
                buffer.clear();
                continue;
            }
            _ => {}
        }
        writer.write_event(event).map_err(err)?;
        buffer.clear();
    }
    if columns != 1 {
        return Err(crate::coded(
            "internal-missing-field",
            "Missing table columns",
        ));
    }
    Ok(writer.into_inner())
}
fn sort_keys(input: &Value, first: u32, last: u32) -> Result<Vec<(u32, bool)>> {
    let object = input
        .as_object()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid sort command"))?;
    let legacy;
    let keys = if object.contains_key("keys") {
        if object.contains_key("column") || object.contains_key("direction") {
            return Err(crate::coded(
                "sort-unsupported",
                "Sort keys cannot be mixed with legacy column or direction",
            ));
        }
        input["keys"]
            .as_array()
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid sort keys"))?
    } else {
        legacy = vec![json!({"column":input["column"],"direction":input["direction"]})];
        &legacy
    };
    if keys.is_empty() || keys.len() > 16 {
        return Err(crate::coded(
            "sort-bad-input",
            "Sort requires 1–16 ordered keys",
        ));
    }
    let mut result = Vec::with_capacity(keys.len());
    for key in keys {
        let fields = key
            .as_object()
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid sort key fields"))?;
        if fields.len() != 2 || !fields.contains_key("column") || !fields.contains_key("direction")
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid sort key fields",
            ));
        }
        let column = field(key, "column")?;
        if column.is_empty() || column.len() > 3 || !column.bytes().all(|c| c.is_ascii_alphabetic())
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid sort column",
            ));
        }
        let col = position(&format!("{column}1"))?.1;
        if col < first || col > last {
            return Err(crate::coded(
                "sort-unsupported",
                "Sort key column is outside the requested range",
            ));
        }
        if result.iter().any(|(previous, _)| *previous == col) {
            return Err(crate::coded("sort-duplicate", "Duplicate sort key column"));
        }
        let descending = match field(key, "direction")? {
            "asc" => false,
            "desc" => true,
            _ => {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid sort direction",
                ));
            }
        };
        result.push((col, descending));
    }
    Ok(result)
}
pub fn apply(root: &Path, input: &Value) -> Result<Value> {
    let db = database(root)?;
    let rewrite = crate::used_range_transfer::admit_writer(
        root,
        &db,
        input,
        crate::used_range_transfer::WriterKind::Transform,
    )?;
    if rewrite.is_some() {
        let inventory = crate::used_range_transfer::worksheet_inventory(&local(
            root,
            field(input, "source")?,
        )?)?;
        let supplied = input["view"]["sheets"].as_array().ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing rewrite worksheet inventory",
            )
        })?;
        if supplied.len() != inventory.len()
            || supplied
                .iter()
                .zip(&inventory)
                .any(|(a, b)| a["name"] != b.name || a["path"] != b.path)
        {
            return Err(crate::coded(
                "internal-mismatch",
                "Transform rewrite inventory differs from its native input",
            ));
        }
    }
    // Owned filter admission has a closed public command envelope. The private
    // receipt was admitted above and is not an authored filter field.
    let mut authored = input.clone();
    authored
        .as_object_mut()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid transform request"))?
        .remove("rewriteInputId");
    let input = &authored;
    let kind = field(input, "kind")?;
    if ![
        "sortRows",
        "filterRows",
        "reapplyFilter",
        "clearFilter",
        "removeFilter",
        "insertRow",
        "deleteRow",
        "insertColumn",
        "deleteColumn",
    ]
    .contains(&kind)
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown native transform",
        ));
    }
    let owned = if ["filterRows", "reapplyFilter", "clearFilter", "removeFilter"].contains(&kind) {
        Some(owned_filter::prepare(root, input)?)
    } else {
        None
    };
    let base = owned.as_ref().map(|p| &p.input).unwrap_or(input);
    let normalized = if owned.is_some() {
        let mut normalized = base.clone();
        let values = normalized["values"]
            .as_array_mut()
            .filter(|v| v.len() <= 1000)
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid filter values"))?;
        for value in values {
            if value.as_object().is_some_and(|v| {
                v.len() == 1 && v.get("$workbookNumber").and_then(Value::as_str) == Some("-0")
            }) {
                *value = json!(-0.0);
            } else if value.is_object() || value.is_array() {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid filter values",
                ));
            }
        }
        Some(normalized)
    } else {
        None
    };
    let input = normalized.as_ref().unwrap_or(base);
    let sort = kind == "sortRows";
    let filter = owned.is_some();
    let keys = if sort {
        let (first, last) = bounds(field(input, "range")?)?;
        sort_keys(input, first.1, last.1)?
    } else {
        Vec::new()
    };
    let row_axis = kind.ends_with("Row");
    let remove = kind.starts_with("delete");
    let sheet = field(input, "sheet")?;
    if crate::spill::present(&db)?
        && crate::spill::any_active(
            &db,
            input["version"].as_u64().ok_or_else(|| {
                crate::coded(
                    "internal-missing-field",
                    "Missing spill-safe transform version",
                )
            })?,
        )?
    {
        return Err(crate::coded(
            "spill-unsupported",
            "Transforming accepted spill ownership requires qualified geometry/reference preservation",
        ));
    }
    let target_name = input["view"]["sheets"]
        .as_array()
        .and_then(|s| s.iter().find(|s| s["path"] == sheet))
        .and_then(|s| s["name"].as_str());
    let mut filter_table = None;
    if filter && let Some(tables) = input["view"]["tables"].as_array() {
        let tables = tables
            .iter()
            .filter(|t| t["sheet"].as_str() == target_name)
            .collect::<Vec<_>>();
        if !tables.is_empty() {
            let table = tables
                .into_iter()
                .find(|t| t["ref"] == input["range"])
                .ok_or_else(|| {
                    crate::coded(
                        "filter-bad-input",
                        "Filter must match the complete table range",
                    )
                })?;
            filter_table = Some(field(table, "path")?);
        }
    }
    // Named ranges describe areas, which a sort does not move; every filtered
    // sheet has one. Sorting the data rows of a table moves rows inside it and
    // leaves the table itself as it is, so that is allowed, and so is sorting
    // rows that no table touches. A sort that would carry a table's header or
    // totals row, or only part of its columns, is refused.
    if sort {
        let (first, last) = bounds(field(input, "range")?)?;
        for table in input["view"]["tables"]
            .as_array()
            .into_iter()
            .flatten()
            .filter(|t| t["sheet"].as_str() == target_name)
        {
            let (a, b) = bounds(field(table, "ref")?)?;
            if first.0 > b.0 || last.0 < a.0 {
                continue;
            }
            let header = table["headerRowCount"].as_u64().unwrap_or(1) as u32;
            let totals = table["totalsRowCount"].as_u64().unwrap_or(0) as u32;
            let inside_body = first.0 >= a.0 + header
                && last.0 + totals <= b.0
                && first.1 <= a.1
                && last.1 >= b.1;
            if !inside_body {
                return Err(crate::coded(
                    "table-or-protection",
                    "Sort the rows inside the table (all of its columns, without its header or totals row), or rows away from it",
                ));
            }
        }
    }
    if sort {
        let mut options = icu_collator::options::CollatorOptions::default();
        options.strength = Some(icu_collator::options::Strength::Secondary);
        let collator = icu_collator::Collator::try_new(Default::default(), options).map_err(err)?;
        db.create_collation("workbook_text", move |a, b| collator.compare(a, b))
            .map_err(err)?;
        let columns = (0..keys.len()).map(|i|format!("kind{i} INTEGER NOT NULL DEFAULT 4,number{i} REAL,text{i} TEXT COLLATE workbook_text")).collect::<Vec<_>>().join(",");
        db.execute_batch(&format!("CREATE TEMP TABLE sort_keys(row INTEGER PRIMARY KEY,{columns}); CREATE TEMP TABLE sort_order(source INTEGER PRIMARY KEY,destination INTEGER NOT NULL UNIQUE);")).map_err(err)?;
    }
    if filter {
        db.execute_batch("CREATE TEMP TABLE filter_keys(r INTEGER PRIMARY KEY,value TEXT NOT NULL); CREATE TEMP TABLE filter_hidden(r INTEGER PRIMARY KEY);").map_err(err)?;
    }
    let rectangle = if sort || filter {
        Some(bounds(field(input, "range")?)?)
    } else {
        None
    };
    if let Some((first, last)) = rectangle {
        let columns = if sort {
            keys.iter().map(|(column, _)| *column).collect::<Vec<_>>()
        } else {
            let col = position(&format!("{}1", field(input, "column")?))?.1;
            if col < first.1 || col > last.1 {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid sort/filter column",
                ));
            }
            vec![col]
        };

        if sort {
            db.execute("WITH RECURSIVE rows(n) AS (SELECT ?1 UNION ALL SELECT n+1 FROM rows WHERE n<?2) INSERT INTO sort_keys(row) SELECT n FROM rows", params![first.0,last.0]).map_err(err)?;
        }
        for (key, col) in columns.into_iter().enumerate() {
            if filter && owned.as_ref().unwrap().clear {
                break;
            }
            let mut after = None;
            loop {
                crate::cancel::check()?;
                let page = read_db(
                    &db,
                    &json!({"sheet":sheet,"version":input["version"],"range":format!("{}:{}",address(first.0,col),address(last.0,col)),"limit":2048,"after":after}),
                )?;
                for item in page["cells"].as_array().unwrap() {
                    let row = position(item[0].as_str().unwrap())?.0;
                    if filter && row == first.0 {
                        continue;
                    }
                    if filter
                        && row > first.0
                        && item[1].get("formula").is_none()
                        && item[1]["dataType"] == "error"
                    {
                        return Err(crate::coded(
                            "filter-unsupported",
                            "Filtering typed errors requires separate qualification",
                        ));
                    }
                    let value = literal(&item[1], true)?;
                    if sort {
                        // Excel type order, stable equal keys, and blanks last in
                        // both directions. Keep key values/permutation in SQLite.
                        let kind = if item[1]["dataType"] == "error" {
                            3
                        } else if value.is_null() || value == "" {
                            4
                        } else if value.is_number() {
                            0
                        } else if value.is_string() {
                            1
                        } else if value.is_boolean() {
                            2
                        } else {
                            return Err(crate::coded("sort-unsupported", "Unsupported sort key"));
                        };
                        let number = value
                            .as_f64()
                            .or_else(|| value.as_bool().map(|b| if b { 1.0 } else { 0.0 }));
                        let text = if kind == 1 { value.as_str() } else { None };
                        db.prepare_cached(&format!("UPDATE sort_keys SET kind{key}=?2,number{key}=?3,text{key}=?4 WHERE row=?1"))
                    .map_err(err)?
                    .execute(params![row, kind, number, text])
                    .map_err(err)?;
                    } else {
                        db.prepare_cached("INSERT INTO filter_keys VALUES(?1,?2)")
                            .map_err(err)?
                            .execute(params![row, value.to_string()])
                            .map_err(err)?;
                    }
                }
                after = page["nextCursor"].as_str().map(str::to_owned);
                if after.is_none() {
                    break;
                }
            }
        }
        if sort {
            // One lexicographic permutation, independent direction per key,
            // blanks last at each level and source-row ties always stable.
            let order=keys.iter().enumerate().map(|(i,(_,descending))|{let direction=if *descending{"DESC"}else{"ASC"};format!("kind{i}=4,kind{i} {direction},number{i} {direction},text{i} COLLATE workbook_text {direction}")}).collect::<Vec<_>>().join(",");
            let sql = format!(
                "INSERT INTO sort_order SELECT row,?1+row_number() OVER (ORDER BY {order},row)-1 FROM sort_keys"
            );
            db.execute(&sql, [first.0]).map_err(err)?;
            db.execute_batch("DROP TABLE sort_keys;").map_err(err)?;
        } else {
            let values = input["values"]
                .as_array()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing filter values"))?;
            if values.iter().any(|v| v.is_object() || v.is_array()) {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid filter values",
                ));
            }
            // Bare OOXML blank=1 represents both missing cells and empty text. A
            // primitive filter cannot preserve only one category on reopening.
            let selects_null = values.iter().any(Value::is_null);
            let selects_empty = values.iter().any(|v| v == "");
            if selects_null != selects_empty && owned.is_none() {
                let (stored, nulls, empty): (u64, u64, u64) = db
                    .query_row(
                        "SELECT count(*),coalesce(sum(value='null'),0),coalesce(sum(value='\"\"'),0) FROM filter_keys WHERE r>?1",
                        [first.0],
                        |r| Ok((r.get(0)?, r.get(1)?, r.get(2)?)),
                    )
                    .map_err(err)?;
                let has_null = nulls > 0 || stored < u64::from(last.0 - first.0);
                if has_null && empty > 0 {
                    return Err(crate::coded(
                        "filter-unsupported",
                        "Select both blanks and empty text; the saved filter cannot distinguish them",
                    ));
                }
            }
            for row in first.0 + 1..=last.0 {
                crate::cancel::check()?;
                let text: Option<String> = db
                    .prepare_cached("SELECT value FROM filter_keys WHERE r=?1")
                    .map_err(err)?
                    .query_row([row], |r| r.get(0))
                    .optional()
                    .map_err(err)?;
                let value: Value = text
                    .map(|t| serde_json::from_str(&t).map_err(err))
                    .transpose()?
                    .unwrap_or(Value::Null);
                if !values.is_empty() && !values.iter().any(|v| same(v, &value)) {
                    db.prepare_cached("INSERT INTO filter_hidden VALUES(?1)")
                        .map_err(err)?
                        .execute([row])
                        .map_err(err)?;
                }
            }
        }
    }
    // Validate every formula, including dependents on other sheets, before
    // publishing anything. Keep rewritten formulas in SQLite, not the JS heap.
    db.execute_batch(
        "CREATE TEMP TABLE sorted_formulas(ref TEXT PRIMARY KEY,formula TEXT NOT NULL);",
    )
    .map_err(err)?;
    if sort {
        let sheets = input["view"]["sheets"]
            .as_array()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing sort sheet metadata"))?;
        let target = sheets
            .iter()
            .find(|s| s["path"] == sheet)
            .ok_or_else(|| crate::coded("internal-unknown-reference", "Unknown sort worksheet"))?;
        let target_name = field(target, "name")?;
        let (first, last) = rectangle.unwrap();
        crate::calculation::snapshot(
            &db,
            input["version"]
                .as_u64()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing version"))?,
        )?;
        for owner in sheets {
            let mut after = None;
            loop {
                let page = crate::calculation::page(
                    &db,
                    field(owner, "path")?,
                    (0, 0),
                    (1048575, 16383),
                    after,
                    true,
                )?;
                for (reference, cell) in &page {
                    crate::cancel::check()?;
                    if !cell["formulaType"].is_null() {
                        return Err(crate::coded(
                            "sort-unsupported",
                            "Special formula sort is not qualified",
                        ));
                    }
                    let row = position(reference)?.0;
                    let destination = if owner["path"] == sheet {
                        sort_destination(&db, row)?
                    } else {
                        None
                    };
                    let formula = crate::formula::sorted_formula(
                        field(cell, "formula")?,
                        field(owner, "name")?,
                        target_name,
                        row,
                        (first.0, last.0),
                        destination,
                    )
                    .map_err(|reason| {
                        if reason.contains("Dynamic or unsupported function") {
                            format!("Dynamic formula sorting is not qualified: {reason}")
                        } else {
                            crate::coded(
                                "needs-structural-writer",
                                format!(
                                    "Formula sort requires the shared structural writer: {reason}"
                                ),
                            )
                        }
                    })?;
                    if destination.is_some() {
                        db.execute(
                            "INSERT INTO sorted_formulas VALUES(?1,?2)",
                            params![reference, formula],
                        )
                        .map_err(err)?;
                    }
                }
                if page.len() < 512 {
                    break;
                }
                after = Some(position(&page.last().unwrap().0)?);
            }
        }
    }
    let index = if rectangle.is_none() {
        let index = input["index"]
            .as_u64()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing transform index"))?;
        if index >= if row_axis { 1048576 } else { 16384 } {
            return Err(crate::coded(
                "workbook-malformed",
                "Invalid transform index",
            ));
        }
        index as u32
    } else {
        0
    };
    let count = if kind == "insertRow" {
        input["count"].as_u64().unwrap_or(1)
    } else {
        1
    };
    if count == 0 || count > 1048576 {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid insertion count",
        ));
    }
    let shift = |n: u32, limit: u32| -> Result<Option<u32>> {
        if remove && n == index {
            return Ok(None);
        }
        let result = if n < index {
            n as i64
        } else {
            n as i64 + if remove { -1 } else { count as i64 }
        };
        if result < 0 || result >= limit as i64 {
            return Err(crate::coded(
                "workbook-limit",
                "Structural transform exceeds worksheet bounds",
            ));
        }
        Ok(Some(result as u32))
    };
    let source = local(root, field(input, "source")?)?;
    let mut zip =
        ZipArchive::new(crate::cancel::Reader(File::open(&source).map_err(err)?)).map_err(err)?;
    let table_xml = if let Some(table) = filter_table {
        Some(table_filter(
            BufReader::new(zip.by_name(table).map_err(err)?),
            input,
            owned.as_ref().unwrap(),
        )?)
    } else {
        None
    };
    if let Some(owned) = owned.as_ref() {
        owned.metadata_budget(root, filter_xml(input, owned)?.len())?;
    }
    if sort {
        // A sort is judged by what it moves: see `sort_guard`.
        let (first, last) = rectangle.ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing admitted sort rectangle")
        })?;
        sort_guard(&mut zip, sheet, (first.0, last.0))?;
    } else if !filter
        && zip.file_names().any(|name| {
            [
                "comments",
                "threadedComments",
                "pivotTables",
                "externalLinks",
            ]
            .iter()
            .any(|part| name.contains(part))
        })
    {
        return Err(crate::coded(
            "needs-structural-writer",
            "Linked workbook transform requires the shared structural writer",
        ));
    }
    db.execute_batch("CREATE TEMP TABLE ordered_rows(r INTEGER PRIMARY KEY,xml BLOB NOT NULL);")
        .map_err(err)?;
    // First pass validates and spools transformed rows. No output is published yet.
    let member = zip.by_name(sheet).map_err(err)?;
    let mut reader = NsReader::from_reader(BufReader::new(member));
    let mut buffer = Vec::new();
    let mut depth = 0;
    let mut capture: Option<(u32, Writer<Vec<u8>>)> = None;
    // A sort may cover fewer columns than the sheet uses. Cells of a sorted row
    // that lie outside the sorted columns stay on their row, as in Excel:
    // `stay` collects them for the current row (left of the block, right of
    // it), `side` says where the current cell belongs (0 moves with the row),
    // and `sorted_row` holds the row being read and the length of its opening tag.
    db.execute_batch(
        "CREATE TEMP TABLE sort_rows(dest INTEGER PRIMARY KEY,xml BLOB NOT NULL,open INTEGER NOT NULL,close INTEGER NOT NULL);
         CREATE TEMP TABLE sort_stay(r INTEGER NOT NULL,side INTEGER NOT NULL,xml BLOB NOT NULL,PRIMARY KEY(r,side));",
    )
    .map_err(err)?;
    let mut stay = [Writer::new(Vec::new()), Writer::new(Vec::new())];
    let mut side = 0usize;
    let mut sorted_row: Option<(u32, usize)> = None;
    let mut row_name = String::from("row");
    let mut dropping = 0;
    let mut current_row = 0;
    let mut current_ref = String::new();
    let mut in_data = false;
    let mut inherited = BTreeMap::new();
    let mut skipping_filter = 0;
    loop {
        crate::cancel::check()?;
        let (ns, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        let main = ns == quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(NS));
        match &event {
            Event::Start(e) | Event::Empty(e) => {
                let empty = matches!(event, Event::Empty(_));
                let name = e.local_name();
                let name = name.as_ref();
                if skipping_filter > 0 {
                    if !empty {
                        skipping_filter += 1
                    }
                    buffer.clear();
                    continue;
                }
                if filter && main && depth == 1 && name == b"autoFilter" {
                    if !empty {
                        skipping_filter = 1
                    }
                    buffer.clear();
                    continue;
                }
                if dropping > 0 {
                    if !empty {
                        dropping += 1
                    }
                    buffer.clear();
                    continue;
                }

                if main && depth == 1 && name == b"sheetData" {
                    in_data = true;
                }
                if in_data && depth == 2 && (!main || name != b"row") {
                    return Err(crate::coded(
                        "needs-structural-writer",
                        "Extended sheetData requires the shared structural writer",
                    ));
                }

                // A sort moves cell content between rows. Records that describe a
                // fixed area of the sheet stay where they are, as in Excel: frozen
                // panes, highlight rules, validation, page breaks, drawings and a
                // filter that hides nothing. A table and an active filter still
                // block it, and so does a merged range or a link on a sorted row
                // (a link travels with its cell; that move is not made here yet).
                // A merged range or a link stops the sort only when it lies on
                // the sorted rows; one elsewhere on the sheet (a title, a link
                // beside the data) does not move.
                let on_sorted_rows = |reference: Option<&String>| -> bool {
                    let Some((first, last)) = rectangle else {
                        return true;
                    };
                    reference
                        .and_then(|value| bounds(&value.replace('$', "")).ok())
                        .is_none_or(|(a, b)| {
                            a.0 <= last.0 && b.0 >= first.0 && a.1 <= last.1 && b.1 >= first.1
                        })
                };
                // A link on one cell of the sorted rows travels with that cell
                // (the second pass moves it); a link over several cells does not.
                let travels = |reference: Option<&String>| -> bool {
                    let (Some((first, last)), Some(reference)) = (rectangle, reference) else {
                        return false;
                    };
                    !reference.contains(':')
                        && position(&reference.replace('$', "")).is_ok_and(|(row, col)| {
                            row >= first.0 && row <= last.0 && col >= first.1 && col <= last.1
                        })
                };
                // Tables were judged above, against the sorted rows. A filter
                // stops a sort only while it has criteria: rows are hidden then.
                // A filter that only remembers how the range was last sorted, or
                // which dropdown buttons to show, hides nothing.
                let blocks_sort = name == b"filterColumn" && !empty
                    || sort && name == b"mergeCell" && on_sorted_rows(attrs(e)?.get("ref"))
                    || sort
                        && name == b"hyperlink"
                        && on_sorted_rows(attrs(e)?.get("ref"))
                        && !travels(attrs(e)?.get("ref"));
                if main
                    && ((sort && blocks_sort)
                        || (!sort
                            && !filter
                            && [
                                b"mergeCells".as_slice(),
                                b"tableParts",
                                b"drawing",
                                b"conditionalFormatting",
                                b"dataValidations",
                                b"hyperlinks",
                                b"autoFilter",
                                b"rowBreaks",
                                b"colBreaks",
                                b"extLst",
                                b"pane",
                            ]
                            .contains(&name))
                        || !sort && !filter && name == b"f"
                        || (filter && filter_table.is_none() && name == b"tableParts"))
                {
                    return Err(crate::coded(
                        "needs-structural-writer",
                        "Referenced worksheet transform requires the shared structural writer",
                    ));
                }
                // Extension records describe fixed areas of the sheet (extended
                // rules, sparklines). A sort or a filter moves or hides rows and
                // leaves those areas where they are, so they are carried as is.
                if !main && !sort && !filter && matches!(&event, Event::Start(_) | Event::Empty(_))
                {
                    return Err(crate::coded(
                        "needs-structural-writer",
                        "Extended worksheet transform requires the shared structural writer",
                    ));
                }
                if filter
                    && filter_table.is_none()
                    && main
                    && name == b"autoFilter"
                    && attrs(e)?.get("ref").map(String::as_str) != input["range"].as_str()
                {
                    return Err(crate::coded(
                        "filter-unsupported",
                        "Replacing unrelated filters is unsupported",
                    ));
                }

                if main && depth == 2 && name == b"row" {
                    let row = attrs(e)?
                        .get("r")
                        .ok_or_else(|| crate::coded("internal-missing-field", "Missing row index"))?
                        .parse::<u32>()
                        .map_err(err)?
                        .checked_sub(1)
                        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid row"))?;
                    current_row = row;
                    let destination = if sort {
                        Some(sort_destination(&db, row)?.unwrap_or(row))
                    } else if !filter && row_axis {
                        shift(row, 1048576)?
                    } else {
                        Some(row)
                    };
                    if let Some(dest) = destination {
                        let mut changes = vec![("r", Some((dest + 1).to_string()))];
                        if filter && rectangle.is_some_and(|(a, b)| row > a.0 && row <= b.0) {
                            let failed = db
                                .prepare_cached(
                                    "SELECT EXISTS(SELECT 1 FROM filter_hidden WHERE r=?1)",
                                )
                                .map_err(err)?
                                .query_row([row], |r| r.get::<_, bool>(0))
                                .map_err(err)?;
                            changes.push(("hidden", owned.as_ref().unwrap().lexical(row, failed)));
                        }
                        if !row_axis && !sort && !filter {
                            changes.push(("spans", None));
                        }
                        let changed = element(e, &changes)?;
                        let mut writer = Writer::new(Vec::new());
                        if sort && sort_destination(&db, row)?.is_some() {
                            row_name = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                            writer.write_event(Event::Start(changed)).map_err(err)?;
                            let open = writer.get_ref().len();
                            if empty {
                                writer
                                    .write_event(Event::End(BytesEnd::new(row_name.as_str())))
                                    .map_err(err)?;
                                let bytes = writer.into_inner();
                                db.execute(
                                    "INSERT INTO sort_rows VALUES(?1,?2,?3,?4)",
                                    params![dest, bytes, open, bytes.len() - open],
                                )
                                .map_err(err)?;
                            } else {
                                sorted_row = Some((row, open));
                                capture = Some((dest, writer));
                                depth += 1;
                            }
                            buffer.clear();
                            continue;
                        }
                        writer
                            .write_event(if empty {
                                Event::Empty(changed)
                            } else {
                                Event::Start(changed)
                            })
                            .map_err(err)?;
                        if empty
                            && owned.as_ref().is_some_and(|p| {
                                p.clear
                                    && !p.baseline.contains_key(&row)
                                    && rectangle.is_some_and(|(a, b)| row > a.0 && row <= b.0)
                                    && attrs(e).is_ok_and(|a| {
                                        a.keys().all(|k| {
                                            k == "r" || k == "hidden" || k.starts_with("xmlns")
                                        })
                                    })
                            })
                        {
                            buffer.clear();
                            continue;
                        }
                        if empty {
                            db.execute(
                                "INSERT INTO ordered_rows VALUES(?1,?2)",
                                params![dest, writer.into_inner()],
                            )
                            .map_err(err)?;
                        } else {
                            capture = Some((dest, writer));
                            depth += 1;
                        }
                    } else if !empty {
                        dropping = 1;
                    }
                    buffer.clear();
                    continue;
                }
                let mut replacement = None;
                if main
                    && depth == 3
                    && name == b"c"
                    && let Some((dest, _)) = capture.as_ref()
                {
                    let (row, col) = position(attrs(e)?.get("r").ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing cell address")
                    })?)?;
                    current_ref = address(row, col);
                    if kind == "insertRow" && index > 0 && row == index - 1 {
                        inherited.insert(col, attrs(e)?.get("s").cloned());
                    }
                    if row != current_row {
                        return Err(crate::coded(
                            "workbook-bad-input",
                            "Cell row does not match row container",
                        ));
                    }
                    side = match rectangle {
                        Some((a, _)) if sorted_row.is_some() && col < a.1 => 1,
                        Some((_, b)) if sorted_row.is_some() && col > b.1 => 2,
                        _ => 0,
                    };
                    let c = if !row_axis && !sort && !filter {
                        shift(col, 16384)?
                    } else {
                        Some(col)
                    };
                    if let Some(c) = c {
                        // A cell that stays keeps its own row.
                        let to = if side == 0 { *dest } else { row };
                        replacement = Some(element(e, &[("r", Some(address(to, c)))])?);
                    } else {
                        if !empty {
                            dropping = 1;
                        }
                        buffer.clear();
                        continue;
                    }
                }
                if sort && side == 0 && main && depth == 4 && name == b"f" {
                    if attrs(e)?
                        .keys()
                        .any(|k| k != "xmlns" && !k.starts_with("xmlns:"))
                    {
                        return Err(crate::coded(
                            "sort-unsupported",
                            "Special formula sort is not qualified",
                        ));
                    }
                    let formula: Option<String> = db
                        .query_row(
                            "SELECT formula FROM sorted_formulas WHERE ref=?1",
                            [&current_ref],
                            |r| r.get(0),
                        )
                        .optional()
                        .map_err(err)?;
                    if let Some(formula) = formula {
                        let writer = &mut capture
                            .as_mut()
                            .ok_or_else(|| {
                                crate::coded("formula-unsupported", "Formula outside row")
                            })?
                            .1;
                        writer.write_event(Event::Start(e.clone())).map_err(err)?;
                        writer
                            .write_event(Event::Text(BytesText::new(&formula)))
                            .map_err(err)?;
                        writer.write_event(Event::End(e.to_end())).map_err(err)?;
                        if !empty {
                            dropping = 1;
                        }
                        buffer.clear();
                        continue;
                    }
                }
                let target = if side == 0 {
                    capture.as_mut().map(|(_, writer)| writer)
                } else {
                    capture.is_some().then_some(&mut stay[side - 1])
                };
                if let Some(writer) = target {
                    writer
                        .write_event(if let Some(e) = replacement {
                            if empty {
                                Event::Empty(e)
                            } else {
                                Event::Start(e)
                            }
                        } else {
                            event.clone()
                        })
                        .map_err(err)?;
                }
                if !empty {
                    depth += 1;
                }
            }
            Event::End(_) => {
                if skipping_filter > 0 {
                    skipping_filter -= 1;
                    buffer.clear();
                    continue;
                }
                if depth == 2 && dropping == 0 {
                    in_data = false;
                }

                if dropping > 0 {
                    dropping -= 1;
                    buffer.clear();
                    continue;
                }
                if depth == 3 && capture.is_some() && sorted_row.is_some() {
                    // The end of a sorted row: keep its moving part and what stays apart.
                    let (dest, mut writer) = capture.take().unwrap();
                    let (row, open) = sorted_row.take().unwrap();
                    let before = writer.get_ref().len();
                    writer.write_event(event.clone()).map_err(err)?;
                    let bytes = writer.into_inner();
                    db.execute(
                        "INSERT INTO sort_rows VALUES(?1,?2,?3,?4)",
                        params![dest, bytes, open, bytes.len() - before],
                    )
                    .map_err(err)?;
                    for (index, part) in stay.iter_mut().enumerate() {
                        let kept = std::mem::replace(part, Writer::new(Vec::new())).into_inner();
                        if !kept.is_empty() {
                            db.execute(
                                "INSERT INTO sort_stay VALUES(?1,?2,?3)",
                                params![row, index + 1, kept],
                            )
                            .map_err(err)?;
                        }
                    }
                    side = 0;
                    depth -= 1;
                    buffer.clear();
                    continue;
                }
                let target = if side == 0 {
                    capture.as_mut().map(|(_, writer)| writer)
                } else {
                    capture.is_some().then_some(&mut stay[side - 1])
                };
                if let Some(writer) = target {
                    writer.write_event(event.clone()).map_err(err)?;
                }
                if depth == 3 && capture.is_some() {
                    let (r, writer) = capture.take().unwrap();
                    db.execute(
                        "INSERT INTO ordered_rows VALUES(?1,?2)",
                        params![r, writer.into_inner()],
                    )
                    .map_err(err)?;
                }
                depth -= 1;
            }
            Event::Eof => break,
            _ => {
                let target = if side == 0 {
                    capture.as_mut().map(|(_, writer)| writer)
                } else {
                    capture.is_some().then_some(&mut stay[side - 1])
                };
                if dropping == 0
                    && let Some(writer) = target
                {
                    writer.write_event(event.clone()).map_err(err)?;
                }
            }
        }
        buffer.clear();
    }
    if sort {
        // Put each sorted row together: the cells that moved to this position,
        // with the cells that stayed here on either side of them.
        let positions = {
            let mut statement = db
                .prepare("SELECT dest FROM sort_rows UNION SELECT r FROM sort_stay ORDER BY 1")
                .map_err(err)?;
            statement
                .query_map([], |row| row.get::<_, u32>(0))
                .map_err(err)?
                .collect::<std::result::Result<Vec<_>, _>>()
                .map_err(err)?
        };
        let stayed = |position: u32, side: u32| -> Result<Vec<u8>> {
            Ok(db
                .query_row(
                    "SELECT xml FROM sort_stay WHERE r=?1 AND side=?2",
                    params![position, side],
                    |row| row.get::<_, Vec<u8>>(0),
                )
                .optional()
                .map_err(err)?
                .unwrap_or_default())
        };
        for position in positions {
            crate::cancel::check()?;
            let moved: Option<(Vec<u8>, usize, usize)> = db
                .query_row(
                    "SELECT xml,open,close FROM sort_rows WHERE dest=?1",
                    [position],
                    |row| Ok((row.get(0)?, row.get(1)?, row.get(2)?)),
                )
                .optional()
                .map_err(err)?;
            let (bytes, open, close) = moved.unwrap_or_else(|| {
                // No row moved here, but cells stayed: they need a row to sit in.
                let start = format!("<{row_name} r=\"{}\">", position + 1);
                let end = format!("</{row_name}>");
                let (open, close) = (start.len(), end.len());
                ((start + &end).into_bytes(), open, close)
            });
            let (left, right) = (stayed(position, 1)?, stayed(position, 2)?);
            let bytes = if left.is_empty() && right.is_empty() {
                bytes
            } else {
                let end = bytes.len() - close;
                // The row's "spans" hint described the row it came from.
                let head = String::from_utf8_lossy(&bytes[..open]).into_owned();
                let head = match head.find(" spans=\"") {
                    Some(at) => match head[at + 8..].find('"') {
                        Some(length) => format!("{}{}", &head[..at], &head[at + 9 + length..]),
                        None => head,
                    },
                    None => head,
                };
                let mut joined = head.into_bytes();
                joined.extend_from_slice(&left);
                joined.extend_from_slice(&bytes[open..end]);
                joined.extend_from_slice(&right);
                joined.extend_from_slice(&bytes[end..]);
                joined
            };
            db.execute(
                "INSERT INTO ordered_rows VALUES(?1,?2)",
                params![position, bytes],
            )
            .map_err(err)?;
        }
    }
    if kind == "insertRow" && !inherited.is_empty() {
        if u64::from(index) + count > 1048576 {
            return Err(crate::coded(
                "workbook-limit",
                "Insertion exceeds worksheet bounds",
            ));
        }
        for row in index..index + count as u32 {
            crate::cancel::check()?;
            let mut writer = Writer::new(Vec::new());
            let mut start = BytesStart::new("row");
            start.push_attribute(("xmlns", std::str::from_utf8(NS).unwrap()));
            start.push_attribute(("r", (row + 1).to_string().as_str()));
            writer.write_event(Event::Start(start)).map_err(err)?;
            for (col, style) in &inherited {
                let mut cell = BytesStart::new("c");
                cell.push_attribute(("r", address(row, *col).as_str()));
                if let Some(style) = style {
                    cell.push_attribute(("s", style.as_str()));
                }
                writer.write_event(Event::Empty(cell)).map_err(err)?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("row")))
                .map_err(err)?;
            db.execute(
                "INSERT INTO ordered_rows VALUES(?1,?2)",
                params![row, writer.into_inner()],
            )
            .map_err(err)?;
        }
    }
    if filter {
        let mut statement = db
            .prepare_cached("SELECT r FROM filter_hidden ORDER BY r")
            .map_err(err)?;
        let mut rows = statement.query([]).map_err(err)?;
        while let Some(record) = rows.next().map_err(err)? {
            crate::cancel::check()?;
            let row: u32 = record.get(0).map_err(err)?;
            if owned.as_ref().unwrap().inherited_hidden(row) {
                continue;
            }
            let present: Option<u32> = db
                .query_row("SELECT r FROM ordered_rows WHERE r=?1", [row], |r| r.get(0))
                .optional()
                .map_err(err)?;
            if present.is_none() {
                db.execute(
                    "INSERT INTO ordered_rows VALUES(?1,?2)",
                    params![
                        row,
                        format!(
                            "<row xmlns=\"{}\" r=\"{}\" hidden=\"1\"/>",
                            std::str::from_utf8(NS).unwrap(),
                            row + 1
                        )
                        .as_bytes()
                    ],
                )
                .map_err(err)?;
            }
        }
    }
    drop(reader);
    let partial = local(root, "transform.partial")?;
    let output = local(root, "transform.xlsx")?;
    let mut writer = ZipWriter::new(crate::quota::create(&partial).map_err(err)?);
    for i in 0..zip.len() {
        crate::cancel::check()?;
        let member = zip.by_index(i).map_err(err)?;
        if filter_table == Some(member.name()) {
            writer
                .start_file(
                    member.name(),
                    SimpleFileOptions::default()
                        .compression_method(zip::CompressionMethod::Deflated),
                )
                .map_err(err)?;
            writer.write_all(table_xml.as_ref().unwrap()).map_err(err)?;
            continue;
        }
        if member.name() != sheet {
            writer.raw_copy_file(member).map_err(err)?;
            continue;
        }
        writer
            .start_file(
                sheet,
                SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated),
            )
            .map_err(err)?;
        let mut xml = Writer::new(&mut writer);
        let mut reader = NsReader::from_reader(BufReader::new(member));
        let mut buffer = Vec::new();
        let (mut depth, mut skip) = (0, 0);
        loop {
            crate::cancel::check()?;
            let (ns, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
            let main = ns == quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(NS));
            match &event {
                Event::Start(e) | Event::Empty(e) => {
                    let empty = matches!(event, Event::Empty(_));
                    let name = e.local_name();
                    let name = name.as_ref();
                    if skip > 0 {
                        if !empty {
                            skip += 1
                        }
                        buffer.clear();
                        continue;
                    }
                    if main
                        && depth == 1
                        && (name == b"dimension"
                            || filter && filter_table.is_none() && name == b"autoFilter")
                    {
                        if !empty {
                            skip = 1
                        }
                        buffer.clear();
                        continue;
                    }
                    if main && depth == 1 && name == b"sheetData" {
                        xml.write_event(Event::Start(e.clone())).map_err(err)?;
                        let mut statement = db
                            .prepare_cached("SELECT xml FROM ordered_rows ORDER BY r")
                            .map_err(err)?;
                        let mut rows = statement.query([]).map_err(err)?;
                        while let Some(row) = rows.next().map_err(err)? {
                            crate::cancel::check()?;
                            let bytes: Vec<u8> = row.get(0).map_err(err)?;
                            xml.get_mut().write_all(&bytes).map_err(err)?;
                        }
                        xml.write_event(Event::End(BytesEnd::new(
                            String::from_utf8_lossy(e.name().as_ref()).into_owned(),
                        )))
                        .map_err(err)?;
                        if filter && filter_table.is_none() {
                            xml.get_mut()
                                .write_all(&filter_xml(input, owned.as_ref().unwrap())?)
                                .map_err(err)?;
                        }
                        if !empty {
                            skip = 1
                        }
                        buffer.clear();
                        continue;
                    }
                    // A link on a sorted cell goes where its cell went.
                    if sort && main && depth == 2 && name == b"hyperlink" {
                        let reference = attrs(e)?.get("ref").cloned().unwrap_or_default();
                        if !reference.contains(':')
                            && let Ok((row, col)) = position(&reference.replace('$', ""))
                            && let Some(destination) = sort_destination(&db, row)?
                            && rectangle.is_some_and(|(a, b)| col >= a.1 && col <= b.1)
                        {
                            let moved = element(e, &[("ref", Some(address(destination, col)))])?;
                            xml.write_event(if empty {
                                Event::Empty(moved)
                            } else {
                                Event::Start(moved)
                            })
                            .map_err(err)?;
                            if !empty {
                                depth += 1;
                            }
                            buffer.clear();
                            continue;
                        }
                    }
                    if main && depth == 2 && name == b"col" && !row_axis && !sort && !filter {
                        let a = attrs(e)?;
                        let low = a["min"].parse::<u32>().map_err(err)? - 1;
                        let high = a["max"].parse::<u32>().map_err(err)? - 1;
                        if !(remove && low == index && high == index) {
                            let low = if remove && low == index {
                                index
                            } else {
                                shift(low, 16384)?.unwrap()
                            };
                            let high = if remove && high == index {
                                index.saturating_sub(1)
                            } else {
                                shift(high, 16384)?.unwrap()
                            };
                            let changed = element(
                                e,
                                &[
                                    ("min", Some((low + 1).to_string())),
                                    ("max", Some((high + 1).to_string())),
                                ],
                            )?;
                            xml.write_event(if empty {
                                Event::Empty(changed)
                            } else {
                                Event::Start(changed)
                            })
                            .map_err(err)?;
                            if !empty {
                                depth += 1;
                            }
                        } else if !empty {
                            skip = 1;
                        }
                        buffer.clear();
                        continue;
                    }
                    if !empty {
                        depth += 1;
                    }
                }
                Event::End(_) => {
                    if skip > 0 {
                        skip -= 1;
                        buffer.clear();
                        continue;
                    }
                    depth -= 1;
                }
                Event::Eof => break,
                _ => {
                    if skip > 0 {
                        buffer.clear();
                        continue;
                    }
                }
            }
            xml.write_event(event).map_err(err)?;
            buffer.clear();
        }
    }
    writer.finish().map_err(err)?.sync_all().map_err(err)?;
    fs::rename(partial, output).map_err(err)?;
    let prepared = if let Some(rewrite) = rewrite {
        use crate::used_range_transfer::{Axis, NativeMapping};
        let mapping = if sort {
            let (first, last) = rectangle.ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing admitted sort rectangle")
            })?;
            let rows = db
                .prepare_cached("SELECT source,destination FROM sort_order ORDER BY source")
                .map_err(err)?
                .query_map([], |row| Ok((row.get::<_, u32>(0)?, row.get::<_, u32>(1)?)))
                .map_err(err)?
                .collect::<std::result::Result<Vec<_>, _>>()
                .map_err(err)?;
            NativeMapping::Rows {
                sheet: sheet.to_owned(),
                first: first.0,
                last: last.0,
                rows,
            }
        } else if filter {
            NativeMapping::Identity
        } else {
            NativeMapping::Axis {
                sheet: sheet.to_owned(),
                axis: if row_axis { Axis::Row } else { Axis::Column },
                index,
                count: count as u32,
                remove,
            }
        };
        Some(crate::used_range_transfer::prepare_output(
            root,
            &db,
            rewrite,
            mapping,
            "transform.xlsx",
        )?)
    } else {
        None
    };
    drop(db);
    let mut result = json!({"file":"transform.xlsx"});
    if let Some(prepared) = prepared {
        result["rewriteId"] = json!(crate::used_range_transfer::commit_output(root, prepared)?);
    }
    Ok(result)
}

/// What stops a sort outside the sheet's own records: a pivot table that sits on
/// the sorted sheet, or a note on a sorted row (a note travels with its cell in
/// Excel; that move is not made here yet). Notes and pivot tables elsewhere, and
/// links to other workbooks, are not affected by a sort.
fn sort_guard(
    zip: &mut ZipArchive<crate::cancel::Reader<File>>,
    sheet: &str,
    rows: (u32, u32),
) -> Result<()> {
    let refuse = || {
        crate::coded(
            "needs-structural-writer",
            "Linked workbook transform requires the shared structural writer",
        )
    };
    let (directory, file) = sheet
        .rsplit_once('/')
        .ok_or_else(|| crate::coded("workbook-malformed", "Invalid worksheet path"))?;
    let relationships = format!("{directory}/_rels/{file}.rels");
    if !zip.file_names().any(|name| name == relationships) {
        return Ok(());
    }
    let mut text = String::new();
    zip.by_name(&relationships)
        .map_err(err)?
        .read_to_string(&mut text)
        .map_err(err)?;
    let attribute = |element: &str, name: &str| -> Option<String> {
        let start = element.find(&format!("{name}=\""))? + name.len() + 2;
        Some(element[start..start + element[start..].find('"')?].to_owned())
    };
    for element in text.split("<Relationship ").skip(1) {
        let kind = attribute(element, "Type").unwrap_or_default();
        if kind.ends_with("/pivotTable") {
            return Err(crate::coded("workbook-refused", refuse()));
        }
        if !(kind.ends_with("/comments") || kind.ends_with("/threadedComment")) {
            continue;
        }
        let target = attribute(element, "Target").ok_or_else(refuse)?;
        let mut path: Vec<&str> = if target.starts_with('/') {
            Vec::new()
        } else {
            directory.split('/').collect()
        };
        for step in target.trim_start_matches('/').split('/') {
            match step {
                ".." => {
                    path.pop();
                }
                "." | "" => {}
                step => path.push(step),
            }
        }
        let mut notes = String::new();
        zip.by_name(&path.join("/"))
            .map_err(|_| refuse())?
            .read_to_string(&mut notes)
            .map_err(err)?;
        for note in notes.split('<').filter(|element| {
            element.starts_with("comment ") || element.starts_with("threadedComment ")
        }) {
            let cell = attribute(note, "ref").ok_or_else(refuse)?;
            let row = position(&cell)?.0;
            if row >= rows.0 && row <= rows.1 {
                return Err(crate::coded("workbook-refused", refuse()));
            }
        }
    }
    Ok(())
}
