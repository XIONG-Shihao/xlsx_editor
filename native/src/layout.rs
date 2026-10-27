//! Stream worksheet layout changes without decoding its cells. The shared
//! metadata writer supplies qualified cols/sheetViews fragments only.
use crate::xml::{NS, attrs, position};
use crate::{Result, err, field, local};
use quick_xml::{
    NsReader, Writer,
    events::{BytesEnd, BytesStart, Event},
};
use serde_json::{Value, json};
use std::{
    fs::{self, File},
    io::{BufReader, Write},
    path::Path,
};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

fn row_element(name: &str, index: u64, height: f64) -> BytesStart<'static> {
    let mut row = BytesStart::new(name.to_owned());
    row.push_attribute(("r", (index + 1).to_string().as_str()));
    row.push_attribute(("ht", height.to_string().as_str()));
    row.push_attribute(("customHeight", "1"));
    row
}
fn resized(element: &BytesStart<'_>, height: f64) -> Result<BytesStart<'static>> {
    let name = String::from_utf8_lossy(element.name().as_ref()).to_string();
    let mut row = BytesStart::new(name);
    let mut attributes = attrs(element)?;
    attributes.insert("ht".into(), height.to_string());
    attributes.insert("customHeight".into(), "1".into());
    for (k, v) in &attributes {
        row.push_attribute((k.as_str(), v.as_str()));
    }
    Ok(row)
}
// New batch inputs are private planner results, still checked before any ZIP
// writes. XML entities and document/cell payloads are not geometry fragments.
fn validate_fragment(tag: &str, xml: &str) -> Result<()> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut buffer = Vec::new();
    let (mut depth, mut roots, mut nodes) = (0usize, 0usize, 0usize);
    loop {
        let (namespace, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let main = namespace
                    == quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(NS));
                let name = element.local_name();
                if depth == 0 {
                    roots += 1;
                    if roots != 1 || !main || name.as_ref() != tag.as_bytes() {
                        return Err(crate::coded(
                            "layout-malformed",
                            "Invalid layout fragment root",
                        ));
                    }
                }
                if main
                    && [b"worksheet".as_slice(), b"sheetData", b"row", b"c"]
                        .contains(&name.as_ref())
                {
                    return Err(crate::coded(
                        "layout-bad-input",
                        "Cell payload is not a layout fragment",
                    ));
                }
                nodes += 1;
                if nodes > 500000 || depth >= 128 {
                    return Err(crate::coded(
                        "layout-limit",
                        "Layout fragment exceeds XML bounds",
                    ));
                }
                if matches!(event, Event::Start(_)) {
                    depth += 1;
                }
            }
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    crate::coded("internal-invalid-field", "Invalid layout fragment nesting")
                })?;
            }
            Event::DocType(_) => {
                return Err(crate::coded(
                    "layout-malformed",
                    "Layout entities are unsupported",
                ));
            }
            Event::Text(text) if depth == 0 && !text.iter().all(u8::is_ascii_whitespace) => {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid layout fragment text",
                ));
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if roots != 1 || depth != 0 {
        return Err(crate::coded(
            "layout-unsupported",
            "Incomplete layout fragment",
        ));
    }
    Ok(())
}
pub fn apply(root: &Path, input: &Value) -> Result<Value> {
    let apply_merges = input["applyMerges"] == true;
    if apply_merges && (input.get("tag").is_some() || input.get("rewriteInputId").is_some()) {
        return Err(crate::coded(
            "layout-unsupported",
            "Accepted merges cannot mix legacy layout modes",
        ));
    }
    let db = if input.get("rewriteInputId").is_some() || apply_merges {
        Some(crate::database(root)?)
    } else {
        None
    };
    let rewrite = db
        .as_ref()
        .filter(|_| input.get("rewriteInputId").is_some())
        .map(|db| {
            crate::used_range_transfer::admit_writer(
                root,
                db,
                input,
                crate::used_range_transfer::WriterKind::Layout,
            )
        })
        .transpose()?
        .flatten();
    let target = field(input, "sheet")?;
    let unhide = input["command"]["operation"] == "unhideAll";
    if unhide
        && ["tag", "row", "height", "kind", "range"]
            .iter()
            .any(|key| input.get(key).is_some())
    {
        return Err(crate::coded(
            "layout-unsupported",
            "Unhide cannot supply unrelated layout fields",
        ));
    }
    let tag = input["tag"].as_str();
    if tag.is_some_and(|t| !["cols", "sheetViews", "mergeCells"].contains(&t)) {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid layout fragment",
        ));
    }
    let mut fragments = std::collections::BTreeMap::new();
    if let Some(batch) = input.get("fragments") {
        if tag.is_some() || input.get("row").is_some() || input.get("rewriteInputId").is_some() {
            return Err(crate::coded(
                "layout-unsupported",
                "Batch fragments cannot mix legacy layout modes",
            ));
        }
        let batch = batch
            .as_array()
            .filter(|batch| !batch.is_empty() && batch.len() <= 2)
            .ok_or_else(|| {
                crate::coded("internal-invalid-field", "Invalid batch layout fragments")
            })?;
        for fragment in batch {
            let tag = field(fragment, "tag")?;
            let xml = field(fragment, "xml")?;
            if !["cols", "sheetViews"].contains(&tag)
                || xml.len() > 4 * 1024 * 1024
                || fragment.as_object().is_none_or(|fields| fields.len() != 2)
            {
                return Err(crate::coded(
                    "layout-unsupported",
                    "Unqualified batch layout fragment",
                ));
            }
            validate_fragment(tag, xml)?;
            if fragments
                .insert(tag.to_owned(), xml.as_bytes().to_vec())
                .is_some()
            {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid or duplicate layout fragment",
                ));
            }
        }
    }
    if apply_merges {
        let version = input["version"].as_u64().ok_or_else(|| {
            crate::coded("internal-missing-field", "Missing accepted merge version")
        })?;
        fragments.insert(
            "mergeCells".into(),
            crate::merges::fragment(
                db.as_ref().ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing merge database")
                })?,
                version,
                target,
            )?,
        );
    }
    let fragment = if tag.is_some() {
        Some(fs::read(local(root, "layout-fragment.xml")?).map_err(err)?)
    } else {
        None
    };
    if rewrite.is_some() || unhide {
        crate::package_edit::validate_rewrite_layout(root, input, fragment.as_deref())?;
    }
    // Merged ranges have a fixed place among the worksheet's records (after a
    // filter, before rules and links). The ordered fragments below put them there.
    let ordered_merge = tag == Some("mergeCells");
    if ordered_merge {
        fragments.insert("mergeCells".into(), fragment.clone().unwrap());
    }
    let merge = if tag == Some("mergeCells") {
        if !["mergeCells", "unmergeCells"].contains(&field(input, "kind")?) {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid merge operation",
            ));
        }
        let bounds = crate::bulk::bounds(field(input, "range")?)?;
        if bounds.0 == bounds.1 {
            return Err(crate::coded(
                "layout-bad-input",
                "Select at least two cells to merge or unmerge",
            ));
        }
        Some(bounds)
    } else {
        None
    };
    let row = input["row"].as_u64();
    let height = input["height"].as_f64();
    let mut row_changes = std::collections::BTreeMap::new();
    if let Some(rows) = input.get("rows") {
        if tag.is_some() || row.is_some() || input.get("rewriteInputId").is_some() {
            return Err(crate::coded(
                "layout-unsupported",
                "Batch row geometry cannot mix layout modes",
            ));
        }
        let rows = rows
            .as_array()
            .filter(|rows| !rows.is_empty() && rows.len() <= 10000)
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid batch row geometry"))?;
        for row in rows {
            let index = row["index"]
                .as_u64()
                .filter(|index| *index < 1048576)
                .ok_or_else(|| crate::coded("layout-malformed", "Invalid batch row index"))?;
            let height = row["height"]
                .as_f64()
                .filter(|height| (8.0..=200.0).contains(height))
                .ok_or_else(|| {
                    crate::coded("internal-invalid-field", "Invalid batch row height")
                })?;
            if row.as_object().is_none_or(|fields| fields.len() != 2)
                || row_changes.insert(index, height).is_some()
            {
                return Err(crate::coded(
                    "layout-duplicate",
                    "Duplicate or unqualified batch row geometry",
                ));
            }
        }
    } else if !unhide
        && tag.is_none()
        && fragments.is_empty()
        && (row.is_none_or(|n| n >= 1048576) || height.is_none_or(|h| !(8.0..=200.0).contains(&h)))
    {
        return Err(crate::coded("internal-invalid-field", "Invalid row height"));
    }
    if let (Some(row), Some(height)) = (row, height) {
        row_changes.insert(row, height);
    }
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, field(input, "source")?)?).map_err(err)?,
    ))
    .map_err(err)?;
    let partial = local(root, "layout.partial")?;
    let output = local(root, "layout.xlsx")?;
    let mut writer = ZipWriter::new(crate::quota::create(&partial).map_err(err)?);
    let mut found = false;
    for i in 0..zip.len() {
        let member = zip.by_index(i).map_err(err)?;
        if member.name() != target {
            writer.raw_copy_file(member).map_err(err)?;
            continue;
        }
        found = true;
        writer
            .start_file(
                target,
                SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated),
            )
            .map_err(err)?;
        let mut xml = Writer::new(&mut writer);
        let mut reader = NsReader::from_reader(BufReader::new(member));
        let mut buffer = Vec::new();
        let (mut depth, mut skip, mut inserted, mut resized_row) =
            (0usize, 0usize, ordered_merge, false);
        let mut in_data = false;
        let mut in_cols = false;
        let mut row_name = "row".to_owned();
        let mut current_cell = None;
        let (mut default_row, mut default_col) = (15.0, 8.43);
        loop {
            crate::cancel::check()?;
            let (namespace, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
            let main =
                namespace == quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(NS));
            let mut replacement = None;
            match &event {
                Event::Start(e) | Event::Empty(e) => {
                    let empty = matches!(event, Event::Empty(_));
                    let name = e.local_name();
                    let name = name.as_ref();
                    if skip > 0 {
                        if !empty {
                            skip += 1;
                        }
                        buffer.clear();
                        continue;
                    }
                    if unhide && main {
                        if depth == 1
                            && [b"sheetProtection".as_slice(), b"autoFilter", b"tableParts"]
                                .contains(&name)
                        {
                            return Err(crate::coded(
                                "layout-protected",
                                "Unhide requires an unprotected worksheet without linked filters/tables",
                            ));
                        }
                        if depth == 1 && name == b"cols" {
                            in_cols = !empty;
                        }
                        if depth == 1 && name == b"sheetFormatPr" {
                            (default_row, default_col) = crate::unhide::defaults(e)?;
                        }
                        if depth == 1 && name == b"sheetFormatPr"
                            || in_cols && depth == 2 && name == b"col"
                            || in_data && depth == 2 && name == b"row"
                        {
                            let changed =
                                crate::unhide::visible(e, name, default_row, default_col)?;
                            replacement = Some(if empty {
                                Event::Empty(changed)
                            } else {
                                Event::Start(changed)
                            });
                        }
                    }
                    if let Some((first, last)) = merge
                        && main
                    {
                        // Excel refuses a merge inside a table and allows one beside it.
                        // The caller names the sheet's table ranges; without them
                        // any table on the sheet stops the merge.
                        if name == b"tableParts" {
                            let tables = input["tableRanges"].as_array();
                            let inside = match tables {
                                None => true,
                                Some(ranges) => {
                                    let mut hit = false;
                                    for range in ranges {
                                        let (a, b) = crate::bulk::bounds(
                                            range.as_str().ok_or_else(|| {
                                                crate::coded(
                                                    "internal-invalid-field",
                                                    "Invalid table range",
                                                )
                                            })?,
                                        )?;
                                        hit |= a.0 <= last.0
                                            && b.0 >= first.0
                                            && a.1 <= last.1
                                            && b.1 >= first.1;
                                    }
                                    hit
                                }
                            };
                            if inside {
                                return Err(crate::coded(
                                    "table-or-protection",
                                    "Cells inside a table cannot be merged (Excel does not allow it either)",
                                ));
                            }
                        }
                        if input["kind"] == "mergeCells" {
                            if depth == 3 && name == b"c" {
                                current_cell =
                                    Some(position(attrs(e)?.get("r").ok_or_else(|| {
                                        crate::coded(
                                            "internal-missing-field",
                                            "Missing merge cell address",
                                        )
                                    })?)?);
                            }
                            if depth == 4
                                && [b"v".as_slice(), b"is", b"f"].contains(&name)
                                && let Some(cell) = current_cell
                                && cell != first
                                && cell.0 >= first.0
                                && cell.0 <= last.0
                                && cell.1 >= first.1
                                && cell.1 <= last.1
                            {
                                return Err(crate::coded(
                                    "layout-conflict",
                                    format!(
                                        "Merge would hide a nonempty cell: {}",
                                        crate::xml::address(cell.0, cell.1)
                                    ),
                                ));
                            }
                        }
                    }
                    if main && depth == 1 {
                        // Fixed planners produce metadata only. Apply all pending
                        // fragments and row geometry in one worksheet/ZIP pass.
                        for tag in ["sheetViews", "cols", "mergeCells"] {
                            let before = match tag {
                                "sheetViews" => {
                                    [b"sheetFormatPr".as_slice(), b"cols", b"sheetData"]
                                        .contains(&name)
                                }
                                "cols" => name == b"sheetData",
                                "mergeCells" => [
                                    b"phoneticPr".as_slice(),
                                    b"conditionalFormatting",
                                    b"dataValidations",
                                    b"hyperlinks",
                                    b"printOptions",
                                    b"pageMargins",
                                    b"pageSetup",
                                    b"headerFooter",
                                    b"rowBreaks",
                                    b"colBreaks",
                                    b"customProperties",
                                    b"cellWatches",
                                    b"ignoredErrors",
                                    b"smartTags",
                                    b"drawing",
                                    b"legacyDrawing",
                                    b"legacyDrawingHF",
                                    b"drawingHF",
                                    b"picture",
                                    b"oleObjects",
                                    b"controls",
                                    b"webPublishItems",
                                    b"tableParts",
                                    b"extLst",
                                ]
                                .contains(&name),
                                _ => false,
                            };
                            if before && let Some(fragment) = fragments.remove(tag) {
                                xml.get_mut().write_all(&fragment).map_err(err)?;
                            }
                        }
                        if let Ok(name) = std::str::from_utf8(name)
                            && let Some(fragment) = fragments.remove(name)
                        {
                            xml.get_mut().write_all(&fragment).map_err(err)?;
                            if !empty {
                                skip = 1;
                            }
                            buffer.clear();
                            continue;
                        }
                    }
                    if main && depth == 1 && Some(name) == tag.map(str::as_bytes) {
                        if !inserted {
                            xml.get_mut()
                                .write_all(fragment.as_ref().unwrap())
                                .map_err(err)?;
                            inserted = true;
                        }
                        if !empty {
                            skip = 1;
                        }
                        buffer.clear();
                        continue;
                    }
                    if main
                        && depth == 1
                        && !inserted
                        && tag.is_some()
                        && ((tag == Some("cols") && name == b"sheetData")
                            || (tag == Some("sheetViews")
                                && [b"sheetFormatPr".as_slice(), b"cols", b"sheetData"]
                                    .contains(&name)))
                    {
                        xml.get_mut()
                            .write_all(fragment.as_ref().unwrap())
                            .map_err(err)?;
                        inserted = true;
                    }
                    if main && depth == 1 && name == b"sheetData" {
                        in_data = !empty;
                        if empty && tag == Some("mergeCells") && !inserted {
                            xml.write_event(event.clone()).map_err(err)?;
                            xml.get_mut()
                                .write_all(fragment.as_ref().unwrap())
                                .map_err(err)?;
                            inserted = true;
                            in_data = false;
                            buffer.clear();
                            continue;
                        }
                        let full = String::from_utf8_lossy(e.name().as_ref()).to_string();
                        row_name = full.strip_suffix("sheetData").unwrap_or("").to_owned() + "row";
                        if empty && !row_changes.is_empty() {
                            xml.write_event(Event::Start(e.clone())).map_err(err)?;
                            for (index, height) in &row_changes {
                                xml.write_event(Event::Empty(row_element(
                                    &row_name, *index, *height,
                                )))
                                .map_err(err)?;
                            }
                            row_changes.clear();
                            xml.write_event(Event::End(BytesEnd::new(full)))
                                .map_err(err)?;
                            resized_row = true;
                            in_data = false;
                            buffer.clear();
                            continue;
                        }
                    }
                    if main && in_data && depth == 2 && name == b"row" && !row_changes.is_empty() {
                        let r = attrs(e)?
                            .get("r")
                            .ok_or_else(|| {
                                crate::coded("internal-missing-field", "Missing row index")
                            })?
                            .parse::<u64>()
                            .map_err(err)?;
                        while row_changes
                            .first_key_value()
                            .is_some_and(|(index, _)| index + 1 < r)
                        {
                            let (index, height) = row_changes.pop_first().unwrap();
                            xml.write_event(Event::Empty(row_element(&row_name, index, height)))
                                .map_err(err)?;
                            resized_row = true;
                        }
                        if let Some(height) = r
                            .checked_sub(1)
                            .and_then(|index| row_changes.remove(&index))
                        {
                            replacement = Some(if empty {
                                Event::Empty(resized(e, height)?)
                            } else {
                                Event::Start(resized(e, height)?)
                            });
                            resized_row = true;
                        }
                    }
                    if !empty {
                        depth += 1;
                    }
                }
                Event::End(e) => {
                    if main && depth == 2 && e.local_name().as_ref() == b"cols" {
                        in_cols = false;
                    }
                    if skip > 0 {
                        skip -= 1;
                        buffer.clear();
                        continue;
                    }
                    if main
                        && depth == 2
                        && e.local_name().as_ref() == b"sheetData"
                        && !row_changes.is_empty()
                    {
                        for (index, height) in &row_changes {
                            xml.write_event(Event::Empty(row_element(&row_name, *index, *height)))
                                .map_err(err)?;
                        }
                        row_changes.clear();
                        resized_row = true;
                    }
                    if main && depth == 2 && e.local_name().as_ref() == b"sheetData" {
                        in_data = false;
                        if tag == Some("mergeCells") && !inserted {
                            xml.write_event(event.clone()).map_err(err)?;
                            xml.get_mut()
                                .write_all(fragment.as_ref().unwrap())
                                .map_err(err)?;
                            inserted = true;
                            depth -= 1;
                            buffer.clear();
                            continue;
                        }
                    }
                    if main
                        && depth == 1
                        && e.local_name().as_ref() == b"worksheet"
                        && let Some(fragment) = fragments.remove("mergeCells")
                    {
                        xml.get_mut().write_all(&fragment).map_err(err)?;
                    }
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        crate::coded("internal-invalid-field", "Invalid worksheet nesting")
                    })?;
                }
                Event::Eof => break,
                _ => {
                    if skip > 0 {
                        buffer.clear();
                        continue;
                    }
                }
            }
            xml.write_event(replacement.unwrap_or(event)).map_err(err)?;
            buffer.clear();
        }
        if !fragments.is_empty()
            || tag.is_some() && !inserted
            || (row.is_some() || input.get("rows").is_some())
                && (!resized_row || !row_changes.is_empty())
        {
            return Err(crate::coded(
                "internal-missing-field",
                "Missing worksheet layout container",
            ));
        }
    }
    if !found {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown layout worksheet",
        ));
    }
    writer.finish().map_err(err)?.sync_all().map_err(err)?;
    fs::rename(partial, &output).map_err(err)?;
    let prepared = rewrite
        .map(|rewrite| {
            crate::used_range_transfer::prepare_output(
                root,
                db.as_ref().ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing admitted layout database")
                })?,
                rewrite,
                crate::used_range_transfer::NativeMapping::Identity,
                "layout.xlsx",
            )
        })
        .transpose()?;
    drop(db);
    let mut result = json!({"file":"layout.xlsx"});
    if let Some(prepared) = prepared {
        result["rewriteId"] = json!(crate::used_range_transfer::commit_output(root, prepared)?);
    }
    Ok(result)
}
