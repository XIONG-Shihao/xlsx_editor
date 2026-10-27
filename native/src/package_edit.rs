//! Small shared metadata plans; worksheet cells and copied sheets remain streamed.
use crate::xml::{NS, attrs, position};
use crate::{Result, err, field, local};
use quick_xml::{
    NsReader, Writer,
    events::{BytesText, Event},
};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use std::{
    collections::{BTreeMap, BTreeSet},
    fs::{self, File},
    io::{BufReader, Read, Write},
    path::Path,
};
const PNG_LIMIT: usize = 128 * 1024;
const XML_LIMIT: usize = 16 * 1024 * 1024;
const REL_NS: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const DRAW_NS: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const CHART_NS: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/chart";
const A_NS: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const PACKAGE_REL_NS: &[u8] = b"http://schemas.openxmlformats.org/package/2006/relationships";
const TYPE_NS: &[u8] = b"http://schemas.openxmlformats.org/package/2006/content-types";
mod background_authoring;
mod boolean_cell;
mod clear_hyperlinks;
pub(crate) mod export_metadata;
mod hyperlink_authoring;
mod managed_chart_styles;
mod managed_image;
mod managed_ink;
mod managed_model3d;
mod managed_shape;
mod managed_smartart;
mod note_authoring;
mod print_theme;
mod theme_authoring;
/// The XML of one of the page themes this editor offers (by sample id), if there is one.
pub(crate) fn page_theme_sample(id: &str) -> Option<&'static str> {
    theme_authoring::sample(id).ok()
}
pub(crate) fn chart_style_check(bytes: &[u8]) -> Result<Option<&'static str>> {
    managed_chart_styles::check(bytes)
}
mod name_addition;
mod name_update;
mod pivots;
mod print_settings;
mod rewrite_admission;
mod rules;
mod structural_parts;
mod tables;
mod xml_edit;
pub(crate) fn validate_rewrite_layout(
    root: &Path,
    input: &Value,
    fragment: Option<&[u8]>,
) -> Result<()> {
    rewrite_admission::layout(root, input, fragment)
}
/// A pivot table's remembered source range, moved with its rows and columns.
pub(crate) fn insertion_pivot_source(
    root: &Path,
    input: &Value,
    plan: &Value,
    bytes: &[u8],
) -> Result<Option<Vec<u8>>> {
    rewrite_admission::pivot_source(root, input, plan, bytes)
}
/// The worksheet records that move with inserted or deleted rows and columns.
pub(crate) fn insertion_records(root: &Path, input: &Value, plan: &Value) -> Result<Value> {
    rewrite_admission::insertion(root, input, plan)
}
/// Everything beside the worksheets that follows inserted or deleted rows and columns.
pub(crate) fn structural_plan(
    root: &Path,
    db: &rusqlite::Connection,
    version: u64,
    command: &Value,
) -> Result<structural_parts::Plan> {
    structural_parts::plan(root, db, version, command)
}
pub(crate) fn print_check(root: &Path, input: &Value) -> Result<Value> {
    print_settings::check(root, input)
}
pub(crate) fn source_calculation_policy(root: &Path) -> Result<Value> {
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let tree = print_settings::Tree::parse(&part_bytes(
        &mut zip,
        "xl/workbook.xml",
        &serde_json::Map::new(),
    )?)?;
    tree.root(NS, "workbook")?;
    let calc = tree.direct(0, "calcPr")?;
    let attrs = calc.map(|index| &tree.nodes[index].attrs);
    let mode = attrs
        .and_then(|a| a.get("calcMode"))
        .map(String::as_str)
        .unwrap_or("auto");
    let on_save = match attrs.and_then(|a| a.get("calcOnSave")).map(String::as_str) {
        None | Some("1" | "true") => true,
        Some("0" | "false") => false,
        _ => {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid source calculation before Save flag",
            ));
        }
    };
    let opaque = calc.is_some_and(|index| {
        !tree.nodes[index].children.is_empty()
            || !tree.nodes[index].text.trim().is_empty()
            || tree.nodes[index].attrs.keys().any(|key| {
                key != "xmlns"
                    && !key.starts_with("xmlns:")
                    && ![
                        "calcId",
                        "calcMode",
                        "fullCalcOnLoad",
                        "refMode",
                        "iterate",
                        "iterateCount",
                        "iterateDelta",
                        "fullPrecision",
                        "calcCompleted",
                        "calcOnSave",
                        "concurrentCalc",
                        "concurrentManualCount",
                        "forceFullCalc",
                    ]
                    .contains(&key.as_str())
            })
    }) || attrs
        .and_then(|a| a.get("iterate"))
        .is_some_and(|value| value == "1" || value == "true");
    Ok(
        json!({"mode":if mode=="manual"{"manual"}else{"auto"},"recalculateBeforeSave":on_save,"sourceMode":mode,"unsupportedMode":opaque||!["auto","manual"].contains(&mode)}),
    )
}

pub(crate) fn calculation_objects(root: &Path) -> Result<Value> {
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let names = zip.file_names().map(str::to_owned).collect::<BTreeSet<_>>();
    let mut objects = Vec::new();
    let mut outputs = Vec::new();
    let mut foreign = false;
    for part in &names {
        if part.starts_with("xl/charts/") && part.ends_with(".xml") {
            let own = format_prefix!(part, XL_CHARTS_CHART)
                .and_then(|s| s.strip_suffix(".xml"))
                .is_some_and(uuid);
            foreign |= !own;
            objects.push(json!({"key":part,"kind":"chart","owned":own}));
        } else if part.starts_with("xl/pivotTables/") && part.ends_with(".xml") {
            let own = format_matches!(part, XL_PIVOTTABLES_PIVOT1_XML);
            foreign |= !own;
            objects.push(json!({"key":part,"kind":"pivot","owned":own}));
        }
    }
    let owned_charts = objects
        .iter()
        .filter(|object| object["kind"] == "chart" && object["owned"] == true)
        .map(|object| field(object, "key").map(str::to_owned))
        .collect::<Result<BTreeSet<_>>>()?;
    let contexts = managed_chart_export_context(root, &owned_charts)?;
    for object in &mut objects {
        if let Some((sheet, name, _)) = object["key"].as_str().and_then(|key| contexts.get(key)) {
            object["sheet"] = json!(name);
            object["sheetPath"] = json!(sheet);
        }
    }
    let index: Value =
        serde_json::from_slice(&fs::read(local(root, "index.json")?).map_err(err)?).map_err(err)?;
    let empty = serde_json::Map::new();
    for sheet in index["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing indexed sheets"))?
    {
        let path = field(sheet, "path")?;
        let (dir, file) = path
            .rsplit_once('/')
            .ok_or_else(|| crate::coded("package-malformed", "Invalid worksheet path"))?;
        let rel = format!("{dir}/_rels/{file}.rels");
        if !names.contains(&rel) {
            continue;
        }
        let relationships = print_settings::Tree::parse(&part_bytes(&mut zip, &rel, &empty)?)?;
        relationships.root(PACKAGE_REL_NS, "Relationships")?;
        for node in relationships.nodes[0]
            .children
            .iter()
            .map(|i| &relationships.nodes[*i])
        {
            if node.ns != PACKAGE_REL_NS || node.name != "Relationship" {
                return Err(crate::coded(
                    "package-unsupported",
                    "Opaque worksheet relationship graph",
                ));
            }
            if node.attrs.get("Type").map(String::as_str)
                != Some(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable",
                )
            {
                continue;
            }
            if node.attrs.contains_key("TargetMode") {
                return Err(crate::coded(
                    "package-protected",
                    "External pivot graph is protected",
                ));
            }
            let target = resolve(
                path,
                node.attrs
                    .get("Target")
                    .ok_or_else(|| crate::coded("internal-missing-field", "Missing pivot edge"))?,
            )?;
            if !objects.iter().any(|object| object["key"] == target) {
                continue;
            }
            let tree = print_settings::Tree::parse(&part_bytes(&mut zip, &target, &empty)?)?;
            tree.root(NS, "pivotTableDefinition")?;
            let location = tree
                .direct(0, "location")?
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing pivot location"))?;
            let reference = tree.nodes[location]
                .attrs
                .get("ref")
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing pivot output"))?;
            crate::bulk::bounds(reference)?;
            outputs.push(json!({"object":target,"sheet":path,"range":reference}));
        }
    }
    let mut inputs = Vec::new();
    if objects
        .iter()
        .any(|object| format_matches!(object["key"], XL_PIVOTTABLES_PIVOT1_XML))
    {
        match pivots::cached_binding(root) {
            Ok(Some(binding)) => inputs.push(binding),
            Ok(None) => (),
            Err(reason) => {
                crate::cancel::check()?;
                for object in &mut objects {
                    if format_matches!(object["key"], XL_PIVOTTABLES_PIVOT1_XML) {
                        object["qualificationReason"] = json!(reason);
                    }
                }
            }
        }
    }
    Ok(json!({"objects":objects,"outputs":outputs,"inputs":inputs,"foreign":foreign}))
}
pub(crate) fn qualify_calculated_pivot(
    root: &Path,
    _db: &rusqlite::Connection,
    version: u64,
    plan: &Value,
) -> Result<()> {
    if plan["kind"] != "refreshPivot" {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid pivot calculation proof",
        ));
    }
    let mut source = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "source.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    for (part, xml) in plan["replacements"]
        .as_object()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing pivot proof parts"))?
    {
        if part_bytes(&mut source, part, &serde_json::Map::new())?
            != field(&json!({"xml":xml}), "xml")?.as_bytes()
        {
            return Err(crate::coded(
                "internal-mismatch",
                "Calculated pivot proof differs from accepted source cache",
            ));
        }
    }
    // The policy transaction already owns a paged dependency snapshot. A
    // separate readonly connection validates committed source cells without
    // rebuilding TEMP evaluator tables on that same connection.
    pivots::validate(
        root,
        &crate::database(root)?,
        plan,
        &json!({"version":version,"source":"source.xlsx"}),
    )
}
pub(crate) fn qualify_pivot_refresh(
    root: &Path,
    version: u64,
    source: &str,
    plan: &Value,
) -> Result<()> {
    pivots::validate(
        root,
        &crate::database(root)?,
        plan,
        &json!({"version":version,"source":source}),
    )
}

/// Only the scheduling flags may change during ordinary cell/cache export.
/// Every other workbook element/attribute and calcPr setting stays identical.
pub(crate) fn validate_calculation_metadata(
    root: &Path,
    db: &rusqlite::Connection,
    version: u64,
    metadata: &Value,
) -> Result<()> {
    use print_settings::Tree;
    let Some(replacement) = metadata["replacements"].get("xl/workbook.xml") else {
        return Err(crate::coded(
            "internal-missing-field",
            "Missing calculation policy metadata",
        ));
    };
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let before = Tree::parse(&part_bytes(
        &mut zip,
        "xl/workbook.xml",
        &serde_json::Map::new(),
    )?)?;
    let after = Tree::parse(
        replacement
            .as_str()
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid workbook replacement"))?
            .as_bytes(),
    )?;
    before.root(NS, "workbook")?;
    after.root(NS, "workbook")?;
    fn same(a: &Tree, ai: usize, b: &Tree, bi: usize) -> bool {
        let a_node = &a.nodes[ai];
        let b_node = &b.nodes[bi];
        let attrs = |node: &print_settings::Node| {
            node.attrs
                .iter()
                .filter(|(key, _)| *key != "xmlns" && !key.starts_with("xmlns:"))
                .map(|(k, v)| (k.clone(), v.clone()))
                .collect::<BTreeMap<_, _>>()
        };
        if a_node.ns != b_node.ns
            || a_node.name != b_node.name
            || attrs(a_node) != attrs(b_node)
            || a_node.text != b_node.text
        {
            return false;
        }
        let children = |tree: &Tree, index: usize| {
            tree.nodes[index]
                .children
                .iter()
                .copied()
                .filter(|child| {
                    !(index == 0
                        && tree.nodes[*child].ns == NS
                        && tree.nodes[*child].name == "calcPr")
                })
                .collect::<Vec<_>>()
        };
        let ac = children(a, ai);
        let bc = children(b, bi);
        ac.len() == bc.len() && ac.iter().zip(bc).all(|(ai, bi)| same(a, *ai, b, bi))
    }
    if !same(&before, 0, &after, 0) {
        return Err(crate::coded(
            "package-stale",
            "Calculation export changed unrelated workbook metadata",
        ));
    }
    let old = before.direct(0, "calcPr")?;
    let new = after.direct(0, "calcPr")?;
    let source = source_calculation_policy(root)?;
    let (mode, on_save) = crate::calculation_policy::options(db, version)?.unwrap_or((
        field(&source, "mode")?.into(),
        source["recalculateBeforeSave"]
            .as_bool()
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid source policy"))?,
    ));
    let pending = crate::calculation_policy::pending_count(db, version)?;
    let incomplete: bool = db
        .query_row(
            "SELECT incomplete FROM export_state WHERE version=?1",
            [version],
            |row| row.get(0),
        )
        .map_err(err)?;
    let refresh = if pending > 0
        || crate::calculation_policy::derived_count(db, version)? > 0
        || incomplete
    {
        "1"
    } else {
        "0"
    };
    let Some(new) = new else {
        // An unchanged workbook with no calcPr already expresses these
        // defaults. Accept absence only with complete selected cache authority.
        if old.is_none()
            && mode == "auto"
            && on_save
            && refresh == "0"
            && before.bytes == after.bytes
        {
            return Ok(());
        }
        return Err(crate::coded(
            "internal-missing-field",
            "Missing required calcPr export",
        ));
    };
    let flags = ["calcMode", "calcOnSave", "fullCalcOnLoad", "forceFullCalc"];
    let stable = |tree: &Tree, index: usize| {
        tree.nodes[index]
            .attrs
            .iter()
            .filter(|(key, _)| {
                *key != "xmlns" && !key.starts_with("xmlns:") && !flags.contains(&key.as_str())
            })
            .map(|(k, v)| (k.clone(), v.clone()))
            .collect::<BTreeMap<_, _>>()
    };
    let envelope = |tree: &Tree, index: Option<usize>| {
        let mut bytes = tree.bytes.clone();
        if let Some(index) = index {
            bytes.drain(tree.nodes[index].start..tree.nodes[index].end);
        }
        bytes
    };
    if envelope(&before, old) != envelope(&after, Some(new)) {
        return Err(crate::coded(
            "package-stale",
            "Calculation export changed opaque workbook bytes outside calcPr",
        ));
    }
    let preserving = crate::calculation_policy::preserving(db, version)?;
    let old_children = old
        .map(|i| before.nodes[i].children.as_slice())
        .unwrap_or(&[]);
    let child_same = preserving
        && old_children.len() == after.nodes[new].children.len()
        && old_children
            .iter()
            .zip(&after.nodes[new].children)
            .all(|(a, b)| same(&before, *a, &after, *b))
        && old.map(|i| before.nodes[i].text.as_str()).unwrap_or("") == after.nodes[new].text;
    if (!after.nodes[new].children.is_empty() || !after.nodes[new].text.trim().is_empty())
        && !child_same
        || old.map(|index| stable(&before, index)).unwrap_or_default() != stable(&after, new)
    {
        return Err(crate::coded(
            "package-stale",
            "Calculation export changed protected calcPr settings",
        ));
    }
    if crate::calculation_policy::preserving(db, version)? {
        crate::preservation::export_edits(root, db, version)?;
        if !metadata["replacements"].as_object().is_some_and(|items| {
            items.contains_key("xl/workbook.xml")
                && items
                    .keys()
                    .all(|key| ["xl/workbook.xml", "xl/styles.xml"].contains(&key.as_str()))
        }) {
            return Err(crate::coded(
                "package-unsupported",
                "Preservation export cannot replace unrelated package metadata",
            ));
        }
        if let Some(styles) = metadata["replacements"].get("xl/styles.xml") {
            validate_preserved_styles(
                &part_bytes(&mut zip, "xl/styles.xml", &serde_json::Map::new())?,
                styles
                    .as_str()
                    .ok_or_else(|| {
                        crate::coded("internal-invalid-field", "Invalid preservation styles")
                    })?
                    .as_bytes(),
            )?;
        }
        let new_node = &after.nodes[new];
        if new_node.attrs.get("fullCalcOnLoad").map(String::as_str) != Some("1")
            || new_node.attrs.get("forceFullCalc").map(String::as_str) != Some("1")
        {
            return Err(crate::coded(
                "package-unsupported",
                "Preserved results require Excel recalculation on open",
            ));
        }
        for key in ["calcMode", "calcOnSave"] {
            if old.and_then(|i| before.nodes[i].attrs.get(key)) != new_node.attrs.get(key) {
                return Err(crate::coded(
                    "package-stale",
                    "Preservation export changed source calculation options",
                ));
            }
        }
        if !metadata["removed"]
            .as_array()
            .is_some_and(|items| items.is_empty())
            || metadata.get("managedChartCaches").is_some()
            || metadata.get("spillMetadata").is_some()
        {
            return Err(crate::coded(
                "package-unsupported",
                "Preservation export cannot alter cache or package graphs",
            ));
        }
        return Ok(());
    }
    let attrs = &after.nodes[new].attrs;
    let actual_save = match attrs.get("calcOnSave").map(String::as_str) {
        None | Some("1" | "true") => true,
        Some("0" | "false") => false,
        _ => {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid exported Save calculation flag",
            ));
        }
    };
    if attrs.get("calcMode").map(String::as_str) != Some(mode.as_str())
        || actual_save != on_save
        || attrs.get("fullCalcOnLoad").map(String::as_str) != Some(refresh)
        || attrs.get("forceFullCalc").map(String::as_str) != Some(refresh)
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Calculation export flags differ from accepted policy/cache authority",
        ));
    }
    Ok(())
}

fn uuid(value: &str) -> bool {
    value.len() == 36
        && value.bytes().enumerate().all(|(i, ch)| {
            if [8, 13, 18, 23].contains(&i) {
                ch == b'-'
            } else {
                ch.is_ascii_digit() || (b'a'..=b'f').contains(&ch)
            }
        })
}
fn valid_hash(value: &str) -> bool {
    value.len() == 64
        && value
            .bytes()
            .all(|ch| ch.is_ascii_digit() || (b'a'..=b'f').contains(&ch))
}
fn base64(value: &str) -> Result<Vec<u8>> {
    if value.is_empty() || value.len() > PNG_LIMIT.div_ceil(3) * 4 || value.len() % 4 != 0 {
        return Err(crate::coded(
            "package-limit",
            "PNG base64 exceeds its bound",
        ));
    }
    let digit = |ch: u8| -> Result<u8> {
        Ok(match ch {
            b'A'..=b'Z' => ch - b'A',
            b'a'..=b'z' => ch - b'a' + 26,
            b'0'..=b'9' => ch - b'0' + 52,
            b'+' => 62,
            b'/' => 63,
            _ => {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Noncanonical PNG base64",
                ));
            }
        })
    };
    let mut out = Vec::with_capacity(value.len() / 4 * 3);
    for (index, part) in value.as_bytes().chunks_exact(4).enumerate() {
        crate::cancel::check()?;
        let a = digit(part[0])?;
        let b = digit(part[1])?;
        let last = index == value.len() / 4 - 1;
        out.push((a << 2) | (b >> 4));
        if part[2] == b'=' {
            if !last || part[3] != b'=' || b & 15 != 0 {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Noncanonical PNG padding",
                ));
            }
        } else {
            let c = digit(part[2])?;
            out.push((b << 4) | (c >> 2));
            if part[3] == b'=' {
                if !last || c & 3 != 0 {
                    return Err(crate::coded(
                        "internal-invalid-field",
                        "Noncanonical PNG padding",
                    ));
                }
            } else {
                out.push((c << 6) | digit(part[3])?);
            }
        }
    }
    if out.len() > PNG_LIMIT {
        return Err(crate::coded("package-limit", "PNG bytes exceed 128 KiB"));
    }
    Ok(out)
}
fn crc(bytes: &[u8]) -> u32 {
    let mut result = 0xffff_ffffu32;
    for byte in bytes {
        result ^= *byte as u32;
        for _ in 0..8 {
            result = if result & 1 != 0 {
                0xedb8_8320 ^ (result >> 1)
            } else {
                result >> 1
            };
        }
    }
    result ^ 0xffff_ffff
}
fn png(bytes: &[u8], descriptor: &Value) -> Result<()> {
    if bytes.len() > PNG_LIMIT || !bytes.starts_with(b"\x89PNG\r\n\x1a\n") {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid bounded PNG signature",
        ));
    }
    let mut at = 8;
    let mut size = None;
    let mut ended = false;
    let mut idat = false;
    while at < bytes.len() {
        crate::cancel::check()?;
        if at + 12 > bytes.len() {
            return Err(crate::coded("package-malformed", "Truncated PNG chunk"));
        }
        let length = u32::from_be_bytes(bytes[at..at + 4].try_into().map_err(err)?) as usize;
        let end = at
            .checked_add(length)
            .and_then(|n| n.checked_add(12))
            .ok_or_else(|| crate::coded("package-limit", "PNG bounds overflow"))?;
        if end > bytes.len()
            || crc(&bytes[at + 4..at + 8 + length])
                != u32::from_be_bytes(bytes[at + 8 + length..end].try_into().map_err(err)?)
        {
            return Err(crate::coded(
                "package-limit",
                "PNG chunk CRC/bounds mismatch",
            ));
        }
        let kind = &bytes[at + 4..at + 8];
        let data = &bytes[at + 8..at + 8 + length];
        if at == 8 && kind != b"IHDR" || [b"acTL".as_slice(), b"fcTL", b"fdAT"].contains(&kind) {
            return Err(crate::coded(
                "package-unsupported",
                "Only static PNG is qualified",
            ));
        }
        match kind {
            b"IHDR" => {
                if size.is_some() || length != 13 {
                    return Err(crate::coded("internal-invalid-field", "Invalid PNG IHDR"));
                }
                let width = u32::from_be_bytes(data[0..4].try_into().map_err(err)?);
                let height = u32::from_be_bytes(data[4..8].try_into().map_err(err)?);
                if width == 0
                    || height == 0
                    || width > 4096
                    || height > 4096
                    || width as u64 * height as u64 > 1024 * 1024
                    || data[8] != 8
                    || ![2, 6].contains(&data[9])
                    || data[10] != 0
                    || data[11] != 0
                    || data[12] != 0
                {
                    return Err(crate::coded(
                        "package-unsupported",
                        "PNG dimensions/pixel format are not qualified",
                    ));
                }
                if descriptor["widthPx"].as_u64() != Some(width as u64)
                    || descriptor["heightPx"].as_u64() != Some(height as u64)
                {
                    return Err(crate::coded(
                        "internal-mismatch",
                        "PNG dimensions disagree with metadata",
                    ));
                }
                size = Some((width, height));
            }
            b"IDAT" => {
                if length == 0 {
                    return Err(crate::coded("package-bad-input", "Empty PNG pixel chunk"));
                }
                idat = true;
            }
            b"IEND" => {
                if length != 0 || end != bytes.len() {
                    return Err(crate::coded("internal-invalid-field", "Invalid PNG end"));
                }
                ended = true;
            }
            _ => {
                if kind[0].is_ascii_uppercase() && kind != b"PLTE" {
                    return Err(crate::coded(
                        "internal-unknown-reference",
                        "Unknown critical PNG chunk",
                    ));
                }
            }
        }
        at = end;
    }
    if size.is_none() || !idat || !ended {
        return Err(crate::coded("package-unsupported", "Incomplete PNG"));
    }
    let expected = field(descriptor, "sha256")?;
    if !valid_hash(expected) || format!("{:x}", Sha256::digest(bytes)) != expected {
        return Err(crate::coded(
            "internal-mismatch",
            "PNG decoded SHA256 mismatch",
        ));
    }
    Ok(())
}
#[derive(Clone, PartialEq, Eq)]
struct Edge {
    kind: String,
    target: Option<String>,
}
fn rel_owner(part: &str) -> Result<String> {
    if part == "_rels/.rels" {
        return Ok(String::new());
    }
    if let Some((directory, file)) = part.rsplit_once("/_rels/") {
        return Ok(format!(
            "{directory}/{}",
            file.strip_suffix(".rels")
                .ok_or_else(|| crate::coded("package-malformed", "Invalid relationships part"))?
        ));
    }
    if let Some(file) = part.strip_prefix("_rels/") {
        return Ok(file
            .strip_suffix(".rels")
            .ok_or_else(|| crate::coded("package-malformed", "Invalid relationships part"))?
            .to_string());
    }
    Err(crate::coded(
        "package-malformed",
        "Invalid relationships owner",
    ))
}
fn resolve(owner: &str, target: &str) -> Result<String> {
    if target.is_empty() || target.contains(['\\', '?', '#', '%', '\0']) {
        return Err(crate::coded(
            "package-unsupported",
            "Unqualified relationship target",
        ));
    }
    let mut segments = if target.starts_with('/') {
        vec![]
    } else {
        owner
            .rsplit_once('/')
            .map(|(dir, _)| dir.split('/').map(str::to_string).collect())
            .unwrap_or_default()
    };
    for segment in target.trim_start_matches('/').split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                segments.pop().ok_or_else(|| {
                    crate::coded("package-unsupported", "Relationship escapes package")
                })?;
            }
            _ => segments.push(segment.to_string()),
        }
    }
    let result = segments.join("/");
    valid_part(&result)?;
    Ok(result)
}
fn xml_events(
    bytes: &[u8],
    mut visit: impl FnMut(&[u8], &[u8], &BTreeMap<String, String>) -> Result<()>,
) -> Result<()> {
    if bytes.len() > XML_LIMIT {
        return Err(crate::coded(
            "package-limit",
            "Object XML exceeds its bound",
        ));
    }
    let text = std::str::from_utf8(bytes).map_err(err)?;
    let lower = text.to_ascii_lowercase();
    if lower.contains("<!doctype") || lower.contains("<!entity") {
        return Err(crate::coded(
            "package-malformed",
            "Object XML entities are refused",
        ));
    }
    let mut reader = NsReader::from_reader(bytes);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut roots = 0usize;
    loop {
        crate::cancel::check()?;
        let (binding, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                if depth == 0 {
                    roots += 1;
                }
                if roots != 1 || depth >= 64 {
                    return Err(crate::coded(
                        "internal-invalid-field",
                        "Object XML root/depth is invalid",
                    ));
                }
                let namespace = match binding {
                    quick_xml::name::ResolveResult::Bound(ns) => ns.as_ref().to_vec(),
                    _ => Vec::new(),
                };
                let mut values = attrs(element)?;
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(err)?;
                    let (binding, name) = reader.resolve_attribute(attribute.key);
                    if binding
                        == quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(REL_NS))
                    {
                        values.insert(
                            format!("r:{}", String::from_utf8_lossy(name.as_ref())),
                            attribute
                                .decode_and_unescape_value(reader.decoder())
                                .map_err(err)?
                                .into_owned(),
                        );
                    }
                }
                visit(&namespace, element.local_name().as_ref(), &values)?;
                if matches!(event, Event::Start(_)) {
                    depth += 1;
                }
            }
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| crate::coded("package-refused", "Object XML is unbalanced"))?;
            }
            Event::Text(ref text) if depth == 0 => {
                if text.iter().any(|byte| !byte.is_ascii_whitespace()) {
                    return Err(crate::coded(
                        "package-malformed",
                        "Text outside object XML root",
                    ));
                }
            }
            Event::CData(_) if depth == 0 => {
                return Err(crate::coded(
                    "package-unsupported",
                    "CDATA outside object XML root",
                ));
            }
            Event::Eof => {
                if roots != 1 || depth != 0 {
                    return Err(crate::coded("package-unsupported", "Incomplete object XML"));
                }
                break;
            }
            _ => {}
        }
        buffer.clear();
    }
    Ok(())
}
fn xml_root(bytes: &[u8], namespace: &[u8], name: &[u8]) -> Result<()> {
    let mut first = true;
    xml_events(bytes, |ns, local, _| {
        if first && (ns != namespace || local != name) {
            return Err(crate::coded(
                "internal-mismatch",
                "Object XML root namespace/type mismatch",
            ));
        }
        first = false;
        Ok(())
    })
}
// Mask only the claimed top-level anchor. Root namespaces, comments, whitespace,
// and every opaque sibling must survive an object edit byte for byte.
fn drawing_envelope(bytes: &[u8], id: &str) -> Result<(Vec<u8>, usize, usize)> {
    xml_root(bytes, DRAW_NS, b"wsDr")?;
    let mut reader = NsReader::from_reader(bytes);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut begin = 0usize;
    let mut selected = false;
    let mut ranges = Vec::new();
    let mut children = 0usize;
    loop {
        crate::cancel::check()?;
        let at = reader.buffer_position() as usize;
        let (binding, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
        let drawing_namespace =
            binding == quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(DRAW_NS));
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref node) | Event::Empty(ref node) => {
                if depth == 1 {
                    begin = at;
                    selected = false;
                    children += 1;
                }
                if depth >= 1 && drawing_namespace && node.local_name().as_ref() == b"cNvPr" {
                    selected |= attrs(node)?
                        .get("name")
                        .is_some_and(|name| format_identity_matches!(name.as_str(), ID, id = id));
                }
                if matches!(event, Event::Start(_)) {
                    depth += 1;
                } else if depth == 1 && selected {
                    ranges.push((begin, end));
                }
            }
            Event::End(_) => {
                depth -= 1;
                if depth == 1 && selected {
                    ranges.push((begin, end));
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if children > 256 || ranges.len() > 1 {
        return Err(crate::coded(
            "package-unsupported",
            "Drawing anchor ownership is ambiguous/unbounded",
        ));
    }
    let mut envelope = bytes.to_vec();
    if let Some((begin, end)) = ranges.first() {
        envelope.drain(*begin..*end);
    }
    Ok((envelope, ranges.len(), children))
}
fn part_bytes(
    zip: &mut ZipArchive<crate::cancel::Reader<File>>,
    name: &str,
    replacements: &serde_json::Map<String, Value>,
) -> Result<Vec<u8>> {
    if let Some(data) = replacements.get(name) {
        let text = data
            .as_str()
            .ok_or_else(|| crate::coded("package-bad-input", "Object replacement must be XML"))?;
        if text.len() > XML_LIMIT {
            return Err(crate::coded(
                "package-limit",
                "Object XML exceeds its bound",
            ));
        }
        return Ok(text.as_bytes().to_vec());
    }
    let mut member = zip.by_name(name).map_err(err)?;
    if member.size() > XML_LIMIT as u64 {
        return Err(crate::coded(
            "package-limit",
            "Object XML exceeds its bound",
        ));
    }
    let mut data = Vec::new();
    member.read_to_end(&mut data).map_err(err)?;
    Ok(data)
}
#[derive(Default, PartialEq, Eq)]
struct OwnedObject {
    drawing: String,
    target: Option<String>,
    kind: String,
}
/// Whether two lists hold the same records, in any order. The order names and
/// tables are listed in means nothing, and two readers need not agree on it.
pub(crate) fn same_records(a: &Value, b: &Value) -> bool {
    let sorted = |list: &Value| {
        list.as_array().map(|items| {
            let mut texts = items.iter().map(Value::to_string).collect::<Vec<_>>();
            texts.sort_unstable();
            texts
        })
    };
    match (sorted(a), sorted(b)) {
        (Some(a), Some(b)) => a == b,
        _ => a == b,
    }
}
/// Auditing resolves references from immutable package metadata, never from a
/// caller-supplied worksheet/name/table projection. This reads no cell values.
pub(crate) fn auditing_metadata(root: &Path) -> Result<Value> {
    use print_settings::Tree;
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let names = zip.file_names().map(str::to_owned).collect::<BTreeSet<_>>();
    if names.len() != zip.len() {
        return Err(crate::coded(
            "package-ambiguous",
            "Ambiguous auditing package members",
        ));
    }
    let empty = serde_json::Map::new();
    let mut bytes_read = 0usize;
    let mut read = |part: &str| -> Result<Tree> {
        // Reserve the remaining aggregate budget before decompressing a part.
        let size = zip.by_name(part).map_err(err)?.size();
        if size > (64 * 1024 * 1024 - bytes_read) as u64 {
            return Err(crate::coded(
                "package-limit",
                "Auditing metadata exceeds its read bound",
            ));
        }
        let bytes = part_bytes(&mut zip, part, &empty)?;
        bytes_read = bytes_read
            .checked_add(bytes.len())
            .ok_or_else(|| crate::coded("package-limit", "Auditing metadata size overflow"))?;
        if bytes_read > 64 * 1024 * 1024 {
            return Err(crate::coded(
                "package-limit",
                "Auditing metadata exceeds its read bound",
            ));
        }
        Tree::parse(&bytes)
    };
    fn relations(
        tree: &Tree,
        owner: &str,
        names: &BTreeSet<String>,
    ) -> Result<BTreeMap<String, Edge>> {
        tree.root(PACKAGE_REL_NS, "Relationships")?;
        if !tree.nodes[0].text.trim().is_empty() {
            return Err(crate::coded(
                "package-unsupported",
                "Opaque auditing relationships",
            ));
        }
        let mut result = BTreeMap::new();
        for index in &tree.nodes[0].children {
            let node = &tree.nodes[*index];
            if node.ns != PACKAGE_REL_NS || node.name != "Relationship" {
                return Err(crate::coded(
                    "package-unsupported",
                    "Foreign auditing relationship",
                ));
            }
            tree.clean(*index, &["Id", "Type", "Target", "TargetMode"])?;
            let id = node
                .attrs
                .get("Id")
                .filter(|s| !s.is_empty())
                .ok_or_else(|| {
                    crate::coded(
                        "internal-missing-field",
                        "Missing auditing relationship identity",
                    )
                })?;
            let kind = node
                .attrs
                .get("Type")
                .filter(|s| !s.is_empty())
                .ok_or_else(|| {
                    crate::coded(
                        "internal-missing-field",
                        "Missing auditing relationship type",
                    )
                })?
                .clone();
            let target = match node.attrs.get("TargetMode").map(String::as_str) {
                Some("External") => None,
                None | Some("Internal") => {
                    let part = resolve(
                        owner,
                        node.attrs.get("Target").ok_or_else(|| {
                            crate::coded(
                                "internal-missing-field",
                                "Missing auditing relationship target",
                            )
                        })?,
                    )?;
                    if !names.contains(&part) {
                        return Err(crate::coded(
                            "package-unsupported",
                            "Orphan auditing relationship",
                        ));
                    }
                    Some(part)
                }
                _ => {
                    return Err(crate::coded(
                        "internal-unknown-reference",
                        "Unknown auditing relationship mode",
                    ));
                }
            };
            if result.insert(id.clone(), Edge { kind, target }).is_some() {
                return Err(crate::coded(
                    "package-duplicate",
                    "Duplicate auditing relationship identity",
                ));
            }
        }
        Ok(result)
    }
    let workbook = read("xl/workbook.xml")?;
    workbook.root(NS, "workbook")?;
    let workbook_edges = relations(
        &read("xl/_rels/workbook.xml.rels")?,
        "xl/workbook.xml",
        &names,
    )?;
    let container = workbook.direct(0, "sheets")?.ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing auditing worksheet inventory",
        )
    })?;
    let mut sheets = Vec::new();
    let mut parts = BTreeSet::new();
    let mut display_names = BTreeSet::new();
    for index in &workbook.nodes[container].children {
        let node = &workbook.nodes[*index];
        if node.ns != NS
            || node.name != "sheet"
            || !node.children.is_empty()
            || !node.text.trim().is_empty()
        {
            return Err(crate::coded(
                "package-unsupported",
                "Opaque auditing worksheet inventory",
            ));
        }
        let name = node
            .attrs
            .get("name")
            .filter(|s| !s.is_empty())
            .ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing auditing worksheet name")
            })?;
        let edge = workbook_edges
            .get(node.attrs.get("r:id").ok_or_else(|| {
                crate::coded(
                    "internal-missing-field",
                    "Missing auditing worksheet relationship",
                )
            })?)
            .ok_or_else(|| {
                crate::coded(
                    "internal-unknown-reference",
                    "Unknown auditing worksheet relationship",
                )
            })?;
        let part = edge
            .target
            .as_ref()
            .filter(|_| edge.kind == format!("{}/worksheet", String::from_utf8_lossy(REL_NS)))
            .ok_or_else(|| {
                crate::coded(
                    "package-unsupported",
                    "Auditing worksheet relationship is unqualified",
                )
            })?;
        if !parts.insert(part.clone()) || !display_names.insert(name.to_lowercase()) {
            return Err(crate::coded(
                "package-ambiguous",
                "Ambiguous auditing worksheet binding",
            ));
        }
        sheets.push(json!({"name":name,"path":part}));
    }
    if sheets.is_empty() || sheets.len() > 1024 {
        return Err(crate::coded(
            "package-limit",
            "Auditing worksheet inventory exceeds its bound",
        ));
    }
    let mut definitions = Vec::new();
    if let Some(container) = workbook.direct(0, "definedNames")? {
        for index in &workbook.nodes[container].children {
            let node = &workbook.nodes[*index];
            if node.ns != NS || node.name != "definedName" || !node.children.is_empty() {
                return Err(crate::coded(
                    "package-unsupported",
                    "Opaque auditing defined name",
                ));
            }
            let name = node
                .attrs
                .get("name")
                .filter(|s| !s.is_empty())
                .ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing auditing defined name")
                })?;
            let mut value = json!({"name":name,"formula":node.text});
            if let Some(local) = node.attrs.get("localSheetId") {
                let index = if !local.is_empty() && local.bytes().all(|byte| byte.is_ascii_digit())
                {
                    local
                        .parse::<u32>()
                        .ok()
                        .map(|index| index as usize)
                        .filter(|index| *index < sheets.len())
                } else {
                    None
                };
                if let Some(index) = index {
                    value["sheet"] = sheets[index]["name"].clone();
                } else {
                    value["scopeInvalid"] = json!(true);
                }
            }
            definitions.push(value);
        }
    }
    let mut tables = Vec::new();
    let mut table_parts = BTreeSet::new();
    let mut column_count = 0usize;
    for sheet in &sheets {
        crate::cancel::check()?;
        let part = field(sheet, "path")?;
        let worksheet = read(part)?;
        worksheet.root(NS, "worksheet")?;
        let Some(container) = worksheet.direct(0, "tableParts")? else {
            continue;
        };
        let (directory, file) = part
            .rsplit_once('/')
            .ok_or_else(|| crate::coded("package-malformed", "Invalid auditing worksheet part"))?;
        let edges = relations(
            &read(&format!("{directory}/_rels/{file}.rels"))?,
            part,
            &names,
        )?;
        for index in &worksheet.nodes[container].children {
            let node = &worksheet.nodes[*index];
            if node.ns != NS || node.name != "tablePart" {
                return Err(crate::coded(
                    "package-unsupported",
                    "Foreign auditing table binding",
                ));
            }
            worksheet.clean(*index, &["r:id"])?;
            let edge = edges
                .get(node.attrs.get("r:id").ok_or_else(|| {
                    crate::coded(
                        "internal-missing-field",
                        "Missing auditing table relationship",
                    )
                })?)
                .ok_or_else(|| {
                    crate::coded(
                        "internal-unknown-reference",
                        "Unknown auditing table relationship",
                    )
                })?;
            let table_part = edge
                .target
                .as_ref()
                .filter(|_| edge.kind == format!("{}/table", String::from_utf8_lossy(REL_NS)))
                .ok_or_else(|| {
                    crate::coded(
                        "package-unsupported",
                        "Auditing table relationship is unqualified",
                    )
                })?;
            if !table_parts.insert(table_part.clone()) {
                return Err(crate::coded(
                    "package-ambiguous",
                    "Ambiguous auditing table ownership",
                ));
            }
            let table = read(table_part)?;
            table.root(NS, "table")?;
            let attrs = &table.nodes[0].attrs;
            let name = attrs
                .get("displayName")
                .or_else(|| attrs.get("name"))
                .filter(|s| !s.is_empty())
                .ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing auditing table name")
                })?;
            let reference = attrs.get("ref").ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing auditing table range")
            })?;
            let geometry = crate::spill::Geometry::from_reference(reference)?;
            let integer = |key: &str, default: u64| -> Result<u64> {
                match attrs.get(key) {
                    None => Ok(default),
                    Some(value) => value.parse::<u64>().map_err(err),
                }
            };
            let columns = table.direct(0, "tableColumns")?.ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing auditing table columns")
            })?;
            let mut fields = Vec::new();
            let mut column_names = BTreeSet::new();
            let declared = table.nodes[columns]
                .attrs
                .get("count")
                .ok_or_else(|| {
                    crate::coded(
                        "internal-missing-field",
                        "Missing auditing table column count",
                    )
                })?
                .parse::<u64>()
                .map_err(err)?;
            let width = u64::from(geometry.last.1 - geometry.first.1 + 1);
            if declared != width || declared != table.nodes[columns].children.len() as u64 {
                return Err(crate::coded(
                    "internal-mismatch",
                    "Auditing table columns disagree with source geometry",
                ));
            }
            for index in &table.nodes[columns].children {
                let node = &table.nodes[*index];
                if node.ns != NS || node.name != "tableColumn" {
                    return Err(crate::coded(
                        "package-unsupported",
                        "Foreign auditing table column",
                    ));
                }
                let name = node
                    .attrs
                    .get("name")
                    .filter(|name| !name.is_empty())
                    .ok_or_else(|| {
                        crate::coded(
                            "internal-missing-field",
                            "Missing auditing table column name",
                        )
                    })?;
                if !column_names.insert(name.to_lowercase()) {
                    return Err(crate::coded(
                        "package-ambiguous",
                        "Ambiguous auditing table column names",
                    ));
                }
                fields.push(json!({"name":name}));
                column_count += 1;
            }
            tables.push(json!({"name":name,"sheet":sheet["name"],"ref":reference,"path":table_part,"headerRowCount":integer("headerRowCount",1)?,"totalsRowCount":integer("totalsRowCount",0)?,"columns":fields}));
        }
    }
    if definitions.len() > 4096 || tables.len() > 4096 || column_count > 32768 {
        return Err(crate::coded(
            "package-limit",
            "Auditing reference metadata exceeds its bound",
        ));
    }
    Ok(json!({"sheets":sheets,"namedRanges":definitions,"tables":tables}))
}
fn object_graph(
    zip: &mut ZipArchive<crate::cancel::Reader<File>>,
    names: &BTreeSet<String>,
    replacements: &serde_json::Map<String, Value>,
) -> Result<(
    BTreeMap<String, OwnedObject>,
    BTreeMap<String, BTreeMap<String, Edge>>,
)> {
    let mut relations = BTreeMap::new();
    let mut defaults = BTreeMap::new();
    let mut overrides = BTreeMap::new();
    xml_root(
        &part_bytes(zip, "[Content_Types].xml", replacements)?,
        TYPE_NS,
        b"Types",
    )?;
    xml_events(
        &part_bytes(zip, "[Content_Types].xml", replacements)?,
        |ns, name, values| {
            if ns == TYPE_NS && name == b"Default" {
                let extension = values
                    .get("Extension")
                    .ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing type extension")
                    })?
                    .to_ascii_lowercase();
                if defaults
                    .insert(
                        extension,
                        values
                            .get("ContentType")
                            .ok_or_else(|| {
                                crate::coded("internal-missing-field", "Missing content type")
                            })?
                            .clone(),
                    )
                    .is_some()
                {
                    return Err(crate::coded("package-duplicate", "Duplicate content type"));
                }
            }
            if ns == TYPE_NS && name == b"Override" {
                let part = values
                    .get("PartName")
                    .ok_or_else(|| crate::coded("internal-missing-field", "Missing type part"))?
                    .trim_start_matches('/')
                    .to_string();
                if !names.contains(&part)
                    || overrides
                        .insert(
                            part,
                            values
                                .get("ContentType")
                                .ok_or_else(|| {
                                    crate::coded("internal-missing-field", "Missing content type")
                                })?
                                .clone(),
                        )
                        .is_some()
                {
                    return Err(crate::coded(
                        "package-unsupported",
                        "Orphan/duplicate content type override",
                    ));
                }
            }
            Ok(())
        },
    )?;
    for name in names.iter().filter(|name| name.ends_with(".rels")) {
        let owner = rel_owner(name)?;
        if !owner.is_empty() && !names.contains(&owner) {
            return Err(crate::coded(
                "package-unsupported",
                "Orphan relationships owner",
            ));
        }
        let mut edges = BTreeMap::new();
        xml_root(
            &part_bytes(zip, name, replacements)?,
            PACKAGE_REL_NS,
            b"Relationships",
        )?;
        xml_events(
            &part_bytes(zip, name, replacements)?,
            |ns, local, values| {
                if ns != PACKAGE_REL_NS || local != b"Relationship" {
                    return Ok(());
                }
                let id = values.get("Id").ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing relationship ID")
                })?;
                let kind = values
                    .get("Type")
                    .ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing relationship type")
                    })?
                    .clone();
                let target = if values
                    .get("TargetMode")
                    .is_some_and(|mode| mode == "External")
                {
                    None
                } else {
                    let part = resolve(
                        &owner,
                        values.get("Target").ok_or_else(|| {
                            crate::coded("internal-missing-field", "Missing relationship target")
                        })?,
                    )?;
                    if !names.contains(&part) {
                        return Err(crate::coded(
                            "package-unsupported",
                            "Orphan internal relationship",
                        ));
                    }
                    Some(part)
                };
                if id.is_empty() || edges.insert(id.clone(), Edge { kind, target }).is_some() {
                    return Err(crate::coded(
                        "package-duplicate",
                        "Duplicate relationship ID",
                    ));
                }
                Ok(())
            },
        )?;
        relations.insert(owner, edges);
    }
    fn cycle(
        owner: &str,
        relations: &BTreeMap<String, BTreeMap<String, Edge>>,
        visiting: &mut BTreeSet<String>,
        done: &mut BTreeSet<String>,
        depth: usize,
    ) -> Result<()> {
        if done.contains(owner) {
            return Ok(());
        }
        if depth > 64 || !visiting.insert(owner.to_string()) {
            return Err(crate::coded(
                "package-malformed",
                "Cyclic/deep object package relationships are protected",
            ));
        }
        if let Some(edges) = relations.get(owner) {
            for edge in edges.values() {
                if let Some(target) = edge.target.as_deref() {
                    if relations.contains_key(target) {
                        cycle(target, relations, visiting, done, depth + 1)?;
                    }
                }
            }
        }
        visiting.remove(owner);
        done.insert(owner.to_string());
        Ok(())
    }
    let mut done = BTreeSet::new();
    for owner in relations.keys() {
        cycle(owner, &relations, &mut BTreeSet::new(), &mut done, 0)?;
    }
    let content_type = |name: &str| {
        overrides
            .get(name)
            .or_else(|| defaults.get(&name.rsplit('.').next().unwrap_or("").to_ascii_lowercase()))
            .map(String::as_str)
    };
    let mut objects = BTreeMap::new();
    for drawing in names.iter().filter(|name| {
        name.starts_with("xl/drawings/") && name.ends_with(".xml") && !name.contains("/_rels/")
    }) {
        if content_type(drawing)
            != Some("application/vnd.openxmlformats-officedocument.drawing+xml")
        {
            return Err(crate::coded(
                "internal-mismatch",
                "Drawing content type mismatch",
            ));
        }
        let mut ids = BTreeSet::new();
        let mut current = None;
        let mut targets = BTreeMap::new();
        let drawing_bytes = part_bytes(zip, drawing, replacements)?;
        xml_root(&drawing_bytes, DRAW_NS, b"wsDr")?;
        xml_events(&drawing_bytes, |ns, local, values| {
            if ns == DRAW_NS
                && [
                    b"oneCellAnchor".as_slice(),
                    b"twoCellAnchor",
                    b"absoluteAnchor",
                ]
                .contains(&local)
            {
                current = None;
            }
            if ns == DRAW_NS && local == b"cNvPr" {
                let id = values.get("id").ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing drawing nonvisual ID")
                })?;
                id.parse::<u32>().map_err(err)?;
                if !ids.insert(id.clone()) {
                    return Err(crate::coded(
                        "package-duplicate",
                        "Duplicate drawing nonvisual ID",
                    ));
                }
                current = None;
                if let Some(value) = values
                    .get("name")
                    .and_then(|name| format_prefix!(name, FORMAT_))
                    .filter(|id| uuid(id))
                {
                    if objects
                        .insert(
                            value.to_string(),
                            OwnedObject {
                                drawing: drawing.clone(),
                                ..Default::default()
                            },
                        )
                        .is_some()
                    {
                        return Err(crate::coded(
                            "package-duplicate",
                            "Duplicate managed object identity",
                        ));
                    }
                    current = Some(value.to_string());
                }
            }
            if ns == DRAW_NS && local == b"cNvSpPr" {
                if let Some(id) = current.as_ref() {
                    let object = objects.get_mut(id).ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing shape owner")
                    })?;
                    if !object.kind.is_empty() || object.target.is_some() {
                        return Err(crate::coded(
                            "package-ambiguous",
                            "Ambiguous shape ownership",
                        ));
                    }
                    object.kind = "shape".into();
                }
            }
            let is_chart = ns == CHART_NS && local == b"chart";
            let is_image = ns == A_NS && local == b"blip";
            if is_chart || is_image {
                let id = values
                    .get(if is_chart { "r:id" } else { "r:embed" })
                    .ok_or_else(|| {
                        crate::coded(
                            "internal-missing-field",
                            "Missing drawing object relationship",
                        )
                    })?;
                let edge = relations
                    .get(drawing)
                    .and_then(|edges| edges.get(id))
                    .ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing drawing object edge")
                    })?;
                if edge.target.is_none() {
                    if let Some(id) = current.as_ref() {
                        objects
                            .get_mut(id)
                            .ok_or_else(|| {
                                crate::coded("internal-missing-field", "Missing managed object")
                            })?
                            .kind = if is_chart { "chart" } else { "image" }.into();
                    }
                    return Ok(());
                }
                let part = edge.target.as_ref().ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing drawing target")
                })?;
                let suffix = if is_chart { "/chart" } else { "/image" };
                if edge.kind != format!("{}{suffix}", String::from_utf8_lossy(REL_NS)) {
                    return Err(crate::coded(
                        "internal-mismatch",
                        "Drawing relationship type mismatch",
                    ));
                }
                if is_chart
                    && content_type(part)
                        != Some("application/vnd.openxmlformats-officedocument.drawingml.chart+xml")
                    || is_image
                        && !content_type(part).is_some_and(|mime| mime.starts_with("image/"))
                {
                    return Err(crate::coded(
                        "internal-mismatch",
                        "Object content type mismatch",
                    ));
                }
                targets.insert(part.clone(), ());
                if let Some(id) = current.as_ref() {
                    let object = objects.get_mut(id).ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing managed object")
                    })?;
                    if object.target.replace(part.clone()).is_some() {
                        return Err(crate::coded(
                            "package-ambiguous",
                            "Ambiguous managed object graph",
                        ));
                    }
                    object.kind = if is_chart { "chart" } else { "image" }.into();
                }
            }
            Ok(())
        })?;
        for part in targets.keys().filter(|part| part.starts_with("xl/charts/")) {
            xml_root(
                &part_bytes(zip, part, replacements)?,
                CHART_NS,
                b"chartSpace",
            )?;
        }
    }
    if objects.len() > 32
        || objects
            .values()
            .filter(|object| object.kind == "chart")
            .count()
            > 16
    {
        return Err(crate::coded(
            "package-limit",
            "Managed object count exceeds its bound",
        ));
    }
    Ok((objects, relations))
}
/// Cache export may address only existing private managed charts. The tuple is
/// (owned worksheet part, actual workbook sheet name, immutable source XML).
/// Cache/model/value admission remains the export owner's separate responsibility.
pub(crate) fn managed_chart_export_context(
    root: &Path,
    requested_parts: &BTreeSet<String>,
) -> Result<BTreeMap<String, (String, String, Vec<u8>)>> {
    if requested_parts.len() > 16 {
        return Err(crate::coded(
            "package-limit",
            "Managed chart cache count exceeds its bound",
        ));
    }
    if requested_parts.is_empty() {
        return Ok(BTreeMap::new());
    }
    let mut metadata = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let names = metadata
        .file_names()
        .map(str::to_string)
        .collect::<BTreeSet<_>>();
    let empty = serde_json::Map::new();
    let (objects, relations) = object_graph(&mut metadata, &names, &empty)?;
    let index: Value =
        serde_json::from_slice(&fs::read(local(root, "index.json")?).map_err(err)?).map_err(err)?;
    let indexed = index["sheets"]
        .as_array()
        .ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing indexed worksheet ownership",
            )
        })?
        .iter()
        .map(|sheet| field(sheet, "path").map(str::to_string))
        .collect::<Result<BTreeSet<_>>>()?;
    fn relationship_id(
        reader: &NsReader<&[u8]>,
        element: &quick_xml::events::BytesStart<'_>,
    ) -> Result<String> {
        let mut id = None;
        for attribute in element.attributes().with_checks(true) {
            let attribute = attribute.map_err(err)?;
            let (binding, name) = reader.resolve_attribute(attribute.key);
            if binding == quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(REL_NS))
                && name.as_ref() == b"id"
            {
                if id
                    .replace(
                        attribute
                            .decode_and_unescape_value(reader.decoder())
                            .map_err(err)?
                            .into_owned(),
                    )
                    .is_some()
                {
                    return Err(crate::coded(
                        "package-ambiguous",
                        "Ambiguous relationship namespace identity",
                    ));
                }
            }
        }
        id.filter(|id| !id.is_empty()).ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing qualified relationship identity",
            )
        })
    }
    let workbook = part_bytes(&mut metadata, "xl/workbook.xml", &empty)?;
    xml_root(&workbook, NS, b"workbook")?;
    let mut sheet_names = BTreeMap::new();
    let mut display_names = BTreeSet::new();
    {
        let mut reader = NsReader::from_reader(workbook.as_slice());
        let mut buffer = Vec::new();
        let mut depth = 0usize;
        let mut in_sheets = false;
        loop {
            let (binding, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
            let main =
                binding == quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(NS));
            let is_empty = matches!(&event, Event::Empty(_));
            match event {
                Event::Start(element) | Event::Empty(element) => {
                    if depth == 1 && main && element.local_name().as_ref() == b"sheets" {
                        in_sheets = true;
                    }
                    if depth == 2 && in_sheets && main && element.local_name().as_ref() == b"sheet"
                    {
                        let id = relationship_id(&reader, &element)?;
                        let edge = relations
                            .get("xl/workbook.xml")
                            .and_then(|edges| edges.get(&id))
                            .ok_or_else(|| {
                                crate::coded(
                                    "internal-missing-field",
                                    "Workbook sheet relationship is missing",
                                )
                            })?;
                        if edge.kind != format!("{}/worksheet", String::from_utf8_lossy(REL_NS)) {
                            return Err(crate::coded(
                                "package-protected",
                                "Non-worksheet cache ownership is protected",
                            ));
                        }
                        let part = edge
                            .target
                            .as_ref()
                            .filter(|part| indexed.contains(*part))
                            .ok_or_else(|| {
                                crate::coded(
                                    "package-bad-input",
                                    "Sheet cache owner is not indexed",
                                )
                            })?;
                        let values = attrs(&element)?;
                        let name = values
                            .get("name")
                            .filter(|name| !name.is_empty())
                            .ok_or_else(|| {
                                crate::coded(
                                    "internal-missing-field",
                                    "Worksheet display name is missing",
                                )
                            })?;
                        if !display_names.insert(name.to_lowercase())
                            || sheet_names.insert(part.clone(), name.clone()).is_some()
                        {
                            return Err(crate::coded(
                                "package-ambiguous",
                                "Ambiguous workbook worksheet identity",
                            ));
                        }
                    }
                    if !is_empty {
                        depth += 1;
                    } else if depth == 1 && main && element.local_name().as_ref() == b"sheets" {
                        in_sheets = false;
                    }
                }
                Event::End(element) => {
                    if depth == 2 && main && element.local_name().as_ref() == b"sheets" {
                        in_sheets = false;
                    }
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        crate::coded("package-malformed", "Unbalanced workbook ownership")
                    })?;
                }
                Event::Eof => break,
                _ => (),
            }
            buffer.clear();
        }
    }
    let mut result = BTreeMap::new();
    for part in requested_parts {
        crate::cancel::check()?;
        let id = format_prefix!(part, XL_CHARTS_CHART)
            .and_then(|name| name.strip_suffix(".xml"))
            .filter(|id| uuid(id))
            .ok_or_else(|| {
                crate::coded(
                    "package-unsupported",
                    "Only existing private managed chart parts may refresh caches",
                )
            })?;
        let owned = objects
            .get(id)
            .filter(|owned| owned.kind == "chart" && owned.target.as_deref() == Some(part.as_str()))
            .ok_or_else(|| {
                crate::coded(
                    "package-bad-input",
                    "Managed chart identity does not own its part",
                )
            })?;
        if relations.contains_key(part) {
            return Err(crate::coded(
                "package-protected",
                "Dependent/external chart relationships are protected",
            ));
        }
        let incoming = |target: &str| {
            relations
                .iter()
                .flat_map(|(owner, edges)| edges.iter().map(move |(id, edge)| (owner, id, edge)))
                .filter(|(_, _, edge)| edge.target.as_deref() == Some(target))
                .collect::<Vec<_>>()
        };
        let drawing_edges = incoming(&owned.drawing);
        let chart_edges = incoming(part);
        if drawing_edges.len() != 1
            || chart_edges.len() != 1
            || chart_edges[0].0 != &owned.drawing
            || chart_edges[0].2.kind != format!("{}/chart", String::from_utf8_lossy(REL_NS))
            || drawing_edges[0].2.kind != format!("{}/drawing", String::from_utf8_lossy(REL_NS))
        {
            return Err(crate::coded(
                "package-protected",
                "Shared/ambiguous managed chart ownership is protected",
            ));
        }
        let sheet = drawing_edges[0].0;
        let display = sheet_names.get(sheet).ok_or_else(|| {
            crate::coded(
                "package-bad-input",
                "Drawing cache owner is not an actual workbook worksheet",
            )
        })?;
        let bytes = part_bytes(&mut metadata, sheet, &empty)?;
        xml_root(&bytes, NS, b"worksheet")?;
        let mut reader = NsReader::from_reader(bytes.as_slice());
        let mut buffer = Vec::new();
        let mut depth = 0usize;
        let mut links = Vec::new();
        loop {
            let (binding, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
            let main =
                binding == quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(NS));
            let is_empty = matches!(&event, Event::Empty(_));
            match event {
                Event::Start(element) | Event::Empty(element) => {
                    if depth == 1 && main && element.local_name().as_ref() == b"drawing" {
                        links.push(relationship_id(&reader, &element)?);
                    }
                    if !is_empty {
                        depth += 1;
                    }
                }
                Event::End(_) => {
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        crate::coded(
                            "package-malformed",
                            "Unbalanced worksheet drawing ownership",
                        )
                    })?
                }
                Event::Eof => break,
                _ => (),
            }
            buffer.clear();
        }
        if links.len() != 1 || links[0] != *drawing_edges[0].1 {
            return Err(crate::coded(
                "package-bad-input",
                "Actual worksheet drawing link does not own the managed chart",
            ));
        }
        let drawing = part_bytes(&mut metadata, &owned.drawing, &empty)?;
        let mut reader = NsReader::from_reader(drawing.as_slice());
        let mut buffer = Vec::new();
        let mut current = false;
        let mut chart_links = 0;
        loop {
            let (binding, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
            let namespace = match binding {
                quick_xml::name::ResolveResult::Bound(ns) => ns.as_ref().to_vec(),
                _ => Vec::new(),
            };
            match event {
                Event::Start(element) | Event::Empty(element) => {
                    if namespace == DRAW_NS
                        && [
                            b"oneCellAnchor".as_slice(),
                            b"twoCellAnchor",
                            b"absoluteAnchor",
                        ]
                        .contains(&element.local_name().as_ref())
                    {
                        current = false;
                    }
                    if namespace == DRAW_NS && element.local_name().as_ref() == b"cNvPr" {
                        current = attrs(&element)?.get("name")
                            == Some(&format!(workbook_format!(ID), id = id));
                    }
                    if current && namespace == CHART_NS && element.local_name().as_ref() == b"chart"
                    {
                        let edge_id = relationship_id(&reader, &element)?;
                        if edge_id != *chart_edges[0].1 {
                            return Err(crate::coded(
                                "package-refused",
                                "Managed drawing identity has another chart link",
                            ));
                        }
                        chart_links += 1;
                    }
                }
                Event::Eof => break,
                _ => (),
            }
            buffer.clear();
        }
        if chart_links != 1 {
            return Err(crate::coded(
                "package-unsupported",
                "Managed chart has no unique qualified drawing identity",
            ));
        }
        let source = part_bytes(&mut metadata, part, &empty)?;
        xml_root(&source, CHART_NS, b"chartSpace")?;
        xml_events(&source, |ns, name, _| {
            if ![CHART_NS, A_NS].contains(&ns)
                || [
                    b"extLst".as_slice(),
                    b"externalData",
                    b"pivotSource",
                    b"pivotFmts",
                    b"userShapes",
                ]
                .contains(&name)
            {
                return Err(crate::coded(
                    "package-protected",
                    "Extended/pivot/external managed chart cache is protected",
                ));
            }
            Ok(())
        })?;
        managed_chart_styles::check(&source)?;
        result.insert(part.clone(), (sheet.clone(), display.clone(), source));
    }
    Ok(result)
}
fn object_additions(
    root: &Path,
    plan: &Value,
    zip: &mut ZipArchive<crate::cancel::Reader<File>>,
) -> Result<BTreeMap<String, Vec<u8>>> {
    let is_object = [
        "createChart",
        "updateChart",
        "insertImage",
        "updateImage",
        "deleteObject",
        "createShape",
        "updateShape",
        "deleteShape",
        "createInkStroke",
        "updateInkStroke",
        "deleteInkStroke",
        "createSmartArt",
        "updateSmartArt",
        "deleteSmartArt",
        "createModel3D",
        "updateModel3D",
        "deleteModel3D",
    ]
    .contains(&field(plan, "kind")?);
    let empty = serde_json::Map::new();
    let binary = plan
        .get("binaryAdditions")
        .map(|value| {
            value
                .as_object()
                .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid binary additions"))
        })
        .transpose()?
        .unwrap_or(&empty);
    if !is_object {
        if !binary.is_empty() {
            return Err(crate::coded(
                "package-bad-input",
                "Binary additions require object authoring",
            ));
        }
        return Ok(BTreeMap::new());
    }
    let id = field(plan, "objectId")?;
    if !uuid(id) {
        return Err(crate::coded(
            "package-malformed",
            "Invalid managed object identity",
        ));
    }
    if plan.as_object().is_none_or(|fields| {
        fields.len() != 11
            || fields.keys().any(|key| {
                ![
                    "kind",
                    "sheet",
                    "objectId",
                    "priorObjectRevision",
                    "drawingPart",
                    "objectPart",
                    "replacements",
                    "removed",
                    "copies",
                    "fragments",
                    "binaryAdditions",
                ]
                .contains(&key.as_str())
            })
    }) {
        return Err(crate::coded(
            "package-unsupported",
            "Unqualified object plan fields",
        ));
    }
    let replacements = plan["replacements"]
        .as_object()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing replacement parts"))?;
    let copies = plan["copies"]
        .as_object()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing copies"))?;
    let removed = plan["removed"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing removed parts"))?;
    if !copies.is_empty() || binary.len() > 1 || plan.get("rename").is_some() {
        return Err(crate::coded(
            "package-unsupported",
            "Unqualified object package shape",
        ));
    }
    let mut original = BTreeSet::new();
    for name in zip.file_names() {
        valid_part(name)?;
        if !original.insert(name.to_string()) {
            return Err(crate::coded(
                "package-duplicate",
                "Duplicate source package part",
            ));
        }
    }
    let mut names = original.clone();
    let mut removal = BTreeSet::new();
    for part in removed {
        let part = part
            .as_str()
            .ok_or_else(|| crate::coded("package-malformed", "Invalid removed part"))?;
        valid_part(part)?;
        if !original.contains(part)
            || !removal.insert(part.to_string())
            || replacements.contains_key(part)
        {
            return Err(crate::coded(
                "package-duplicate",
                "Duplicate/conflicting object removal",
            ));
        }
        names.remove(part);
    }
    for name in replacements.keys() {
        valid_part(name)?;
        if binary.contains_key(name) {
            return Err(crate::coded(
                "package-duplicate",
                "Duplicate XML/binary part claim",
            ));
        }
        names.insert(name.clone());
    }
    let mut additions = BTreeMap::new();
    let model3d =
        ["createModel3D", "updateModel3D", "deleteModel3D"].contains(&field(plan, "kind")?);
    for (name, descriptor) in binary {
        if model3d {
            let sha = field(descriptor, "sha256")?;
            if !valid_hash(sha)
                || name
                    != &format!(
                        workbook_format!(XL_MEDIA_MODEL3D_ID_SHA_GLB),
                        id = id,
                        sha = sha
                    )
                || original.contains(name)
                || removal.contains(name)
                || descriptor.as_object().is_none_or(|f| {
                    f.len() != 2
                        || f.keys()
                            .any(|k| !["base64", "sha256"].contains(&k.as_str()))
                })
                || field(descriptor, "base64")?.len() > ((64 * 1024 + 2) / 3) * 4
            {
                return Err(crate::coded(
                    "package-unsupported",
                    "Unqualified 3D binary addition",
                ));
            }
            let bytes = base64(field(descriptor, "base64")?)?;
            managed_model3d::glb(&bytes)?;
            if format!("{:x}", Sha256::digest(&bytes)) != sha {
                return Err(crate::coded(
                    "internal-mismatch",
                    "3D addition hash mismatch",
                ));
            }
            names.insert(name.clone());
            additions.insert(name.clone(), bytes);
            continue;
        }
        if field(plan, "kind")? != "insertImage"
            || name
                != &format!(
                    workbook_format!(XL_MEDIA_IMAGE_ID_PNG),
                    field(descriptor, "sha256")?,
                    id = id
                )
            || original.contains(name)
            || removal.contains(name)
            || descriptor.as_object().is_none_or(|fields| {
                fields.len() != 4
                    || fields.keys().any(|key| {
                        !["base64", "sha256", "widthPx", "heightPx"].contains(&key.as_str())
                    })
            })
        {
            return Err(crate::coded(
                "package-conflict",
                "Colliding/unqualified PNG addition",
            ));
        }
        let bytes = base64(field(descriptor, "base64")?)?;
        png(&bytes, descriptor)?;
        names.insert(name.clone());
        additions.insert(name.clone(), bytes);
    }
    if field(plan, "kind")? == "insertImage" && additions.len() != 1 {
        return Err(crate::coded(
            "package-bad-input",
            "Image insertion requires one bounded PNG addition",
        ));
    }
    let sheet = field(plan, "sheet")?;
    let drawing = field(plan, "drawingPart")?;
    let object_part = field(plan, "objectPart")?;
    if !original.contains(sheet)
        || !sheet.starts_with("xl/worksheets/")
        || !drawing.starts_with("xl/drawings/")
    {
        return Err(crate::coded(
            "package-bad-input",
            "Object document scope is not granted",
        ));
    }
    let relation_file = |owner: &str| -> Result<String> {
        let (directory, file) = owner
            .rsplit_once('/')
            .ok_or_else(|| crate::coded("package-malformed", "Invalid object part"))?;
        Ok(format!("{directory}/_rels/{file}.rels"))
    };
    let drawing_rels = relation_file(drawing)?;
    let sheet_rels = relation_file(sheet)?;
    let smartart =
        ["createSmartArt", "updateSmartArt", "deleteSmartArt"].contains(&field(plan, "kind")?);
    let diagram_parts = if smartart {
        managed_smartart::planned_parts(id, object_part, field(plan, "kind")? == "createSmartArt")?
    } else {
        Vec::new()
    };
    if replacements.keys().any(|name| {
        !diagram_parts.contains(name)
            && ![
                "[Content_Types].xml",
                drawing,
                &drawing_rels,
                &sheet_rels,
                object_part,
            ]
            .contains(&name.as_str())
    }) || removal.iter().any(|name| {
        !diagram_parts.contains(name)
            && !(model3d
                && (name.starts_with(&format!(workbook_format!(XL_MEDIA_MODEL3D_ID), id = id))
                    || name.starts_with(&format!(workbook_format!(XL_MEDIA_MODEL3D_ID), id = id)))
                && name.ends_with(".glb"))
            && ![drawing, &drawing_rels, object_part].contains(&name.as_str())
    }) || plan["fragments"].as_object().is_none_or(|sheets| {
        sheets.iter().any(|(name, tags)| {
            name != sheet
                || tags
                    .as_object()
                    .is_none_or(|tags| tags.len() != 1 || !tags.contains_key("drawing"))
        })
    }) {
        return Err(crate::coded(
            "package-unsupported",
            "Object plan claims an unrelated package part",
        ));
    }
    let ink =
        ["createInkStroke", "updateInkStroke", "deleteInkStroke"].contains(&field(plan, "kind")?);
    let inline_shape =
        ink || ["createShape", "updateShape", "deleteShape"].contains(&field(plan, "kind")?);
    let create = [
        "createChart",
        "insertImage",
        "createShape",
        "createInkStroke",
        "createSmartArt",
        "createModel3D",
    ]
    .contains(&field(plan, "kind")?);
    let update_image = field(plan, "kind")? == "updateImage";
    let delete = [
        "deleteObject",
        "deleteShape",
        "deleteInkStroke",
        "deleteSmartArt",
        "deleteModel3D",
    ]
    .contains(&field(plan, "kind")?);
    if create && ((!inline_shape && original.contains(object_part)) || !removed.is_empty())
        || !delete && !model3d && !removed.is_empty()
        || delete
            && ((!inline_shape && replacements.contains_key(object_part)) || !binary.is_empty())
        || update_image
            && (replacements.contains_key(object_part)
                || !binary.is_empty()
                || !plan["fragments"]
                    .as_object()
                    .is_some_and(|value| value.is_empty()))
    {
        return Err(crate::coded(
            "package-conflict",
            "Colliding/unqualified object part mutation",
        ));
    }
    if !original.contains(drawing)
        && drawing != format!(workbook_format!(XL_DRAWINGS_DRAWING_ID_XML), id = id)
    {
        return Err(crate::coded(
            "package-bad-input",
            "New drawing identity is not scoped",
        ));
    }
    if create && plan.get("priorObjectRevision") != Some(&Value::Null)
        || !create
            && plan["priorObjectRevision"]
                .as_str()
                .is_none_or(|value| !valid_hash(value))
    {
        return Err(crate::coded(
            "internal-missing-field",
            "Object prior revision guard is missing",
        ));
    }
    let mut metadata = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, "metadata.xlsx")?).map_err(err)?,
    ))
    .map_err(err)?;
    let (before, before_relations) = object_graph(zip, &original, &empty)?;
    let (after, relations) = object_graph(zip, &names, replacements)?;
    if before
        .iter()
        .any(|(key, value)| key != id && after.get(key) != Some(value))
        || after
            .keys()
            .any(|key| key != id && !before.contains_key(key))
    {
        return Err(crate::coded(
            "package-stale",
            "Unrelated managed object ownership changed",
        ));
    }
    let before_drawing = if original.contains(drawing) {
        Some(drawing_envelope(&part_bytes(zip, drawing, &empty)?, id)?)
    } else {
        None
    };
    let after_drawing = if names.contains(drawing) {
        Some(drawing_envelope(
            &part_bytes(zip, drawing, replacements)?,
            id,
        )?)
    } else {
        None
    };
    match (&before_drawing, &after_drawing) {
        (Some((left, left_count, _)), Some((right, right_count, _)))
            if left == right
                && *left_count == usize::from(!create)
                && *right_count == usize::from(!delete) => {}
        (None, Some((_, 1, 1))) if create => {}
        (Some((_, 1, 1)), None) if delete => {}
        _ => {
            return Err(crate::coded(
                "package-stale",
                "Unscoped drawing bytes/anchor ownership changed",
            ));
        }
    }
    if create && before.contains_key(id)
        || !create && !before.contains_key(id)
        || delete && after.contains_key(id)
        || !delete && !after.contains_key(id)
    {
        return Err(crate::coded(
            "package-refused",
            "Managed object membership guard failed",
        ));
    }
    if model3d {
        managed_model3d::check(
            zip,
            &original,
            &names,
            replacements,
            &before_relations,
            &relations,
            &additions,
            plan,
        )?;
        return Ok(additions);
    }
    if smartart {
        managed_smartart::check(
            zip,
            &original,
            &names,
            replacements,
            &before_relations,
            &relations,
            plan,
        )?;
        return Ok(additions);
    }
    let owned = if delete {
        before.get(id)
    } else {
        after.get(id)
    }
    .ok_or_else(|| crate::coded("internal-missing-field", "Missing scoped object"))?;
    if owned.drawing != drawing
        || if inline_shape {
            owned.kind != "shape" || owned.target.is_some() || object_part != drawing
        } else {
            owned.target.as_deref() != Some(object_part)
        }
        || field(plan, "kind")? == "updateChart" && owned.kind != "chart"
        || update_image && (owned.kind != "image" || before.get(id) != after.get(id))
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Object ownership target mismatch",
        ));
    }
    if owned.kind == "chart"
        && object_part != format!(workbook_format!(XL_CHARTS_CHART_ID_XML), id = id)
        || owned.kind == "image"
            && (!object_part.starts_with(workbook_format!(XL_MEDIA_IMAGE))
                || !object_part.ends_with(".png"))
        || !["chart", "image", "shape"].contains(&owned.kind.as_str())
        || inline_shape != (owned.kind == "shape")
    {
        return Err(crate::coded(
            "package-protected",
            "Unmanaged object part is protected",
        ));
    }
    for owner in before_relations
        .keys()
        .chain(relations.keys())
        .collect::<BTreeSet<_>>()
    {
        let old = before_relations.get(owner);
        let new = relations.get(owner);
        if owner != sheet && owner != drawing {
            if old != new {
                return Err(crate::coded(
                    "package-stale",
                    "Unrelated relationship graph changed",
                ));
            }
            continue;
        }
        let allowed_target = if owner == sheet { drawing } else { object_part };
        let preserve = |edges: Option<&BTreeMap<String, Edge>>| {
            edges
                .into_iter()
                .flat_map(|edges| edges.iter())
                .filter(|(_, edge)| edge.target.as_deref() != Some(allowed_target))
                .map(|(key, edge)| (key.clone(), edge.clone()))
                .collect::<BTreeMap<_, _>>()
        };
        if preserve(old) != preserve(new) {
            return Err(crate::coded(
                "package-stale",
                "Unrelated scoped relationships changed",
            ));
        }
    }
    let types = |data: &[u8]| -> Result<BTreeMap<String, String>> {
        let mut result = BTreeMap::new();
        xml_events(data, |ns, local, values| {
            if ns == TYPE_NS && [b"Override".as_slice(), b"Default"].contains(&local) {
                let key = if local == b"Override" {
                    format!(
                        "O:{}",
                        values.get("PartName").ok_or_else(|| crate::coded(
                            "internal-missing-field",
                            "Missing type part"
                        ))?
                    )
                } else {
                    format!(
                        "D:{}",
                        values
                            .get("Extension")
                            .ok_or_else(|| crate::coded(
                                "internal-missing-field",
                                "Missing type extension"
                            ))?
                            .to_ascii_lowercase()
                    )
                };
                result.insert(
                    key,
                    values
                        .get("ContentType")
                        .ok_or_else(|| {
                            crate::coded("internal-missing-field", "Missing content type")
                        })?
                        .clone(),
                );
            }
            Ok(())
        })?;
        Ok(result)
    };
    let mut old_types = types(&part_bytes(zip, "[Content_Types].xml", &empty)?)?;
    let mut new_types = types(&part_bytes(zip, "[Content_Types].xml", replacements)?)?;
    for part in [drawing, object_part] {
        old_types.remove(&format!("O:/{part}"));
        new_types.remove(&format!("O:/{part}"));
    }
    if field(plan, "kind")? == "insertImage" && !old_types.contains_key("D:png") {
        if new_types.remove("D:png") != Some("image/png".into()) {
            return Err(crate::coded(
                "package-bad-input",
                "PNG content type does not match its bytes",
            ));
        }
    }
    if old_types != new_types {
        return Err(crate::coded(
            "package-stale",
            "Unrelated content types changed",
        ));
    }
    if inline_shape {
        if !binary.is_empty() {
            return Err(crate::coded(
                "package-unsupported",
                "Inline shapes cannot add binary media",
            ));
        }
        if original.contains(drawing)
            && names.contains(drawing)
            && before_relations.get(drawing) != relations.get(drawing)
        {
            return Err(crate::coded(
                "package-unsupported",
                "Inline shapes cannot alter drawing relationships",
            ));
        }
        if ["updateShape", "updateInkStroke"].contains(&field(plan, "kind")?)
            && (before_relations != relations
                || part_bytes(zip, "[Content_Types].xml", &empty)?
                    != part_bytes(zip, "[Content_Types].xml", replacements)?)
        {
            return Err(crate::coded(
                "package-unsupported",
                "Shape updates preserve their package graph",
            ));
        }
        let old = if !create {
            Some(part_bytes(zip, drawing, &empty)?)
        } else {
            None
        };
        let next = if !delete {
            Some(part_bytes(zip, drawing, replacements)?)
        } else {
            None
        };
        if ink {
            managed_ink::check(
                old.as_deref(),
                next.as_deref(),
                id,
                plan["priorObjectRevision"].as_str(),
            )?;
        } else {
            managed_shape::check(
                old.as_deref(),
                next.as_deref(),
                id,
                plan["priorObjectRevision"].as_str(),
            )?;
        }
    }
    if owned.kind == "image" {
        if update_image {
            if before_relations != relations
                || part_bytes(zip, "[Content_Types].xml", &empty)?
                    != part_bytes(zip, "[Content_Types].xml", replacements)?
            {
                return Err(crate::coded(
                    "package-unsupported",
                    "Image placement cannot change its package graph",
                ));
            }
            managed_image::check(
                &part_bytes(zip, drawing, &empty)?,
                &part_bytes(zip, drawing, replacements)?,
                &part_bytes(zip, object_part, &empty)?,
                id,
                object_part,
            )?;
        }
        let types = types(&part_bytes(
            zip,
            "[Content_Types].xml",
            if delete { &empty } else { replacements },
        )?)?;
        if types
            .get(&format!("O:/{object_part}"))
            .or_else(|| types.get("D:png"))
            .map(String::as_str)
            != Some("image/png")
        {
            return Err(crate::coded(
                "internal-mismatch",
                "Managed PNG content type mismatch",
            ));
        }
    }
    let incoming = |edges: &BTreeMap<String, BTreeMap<String, Edge>>, part: &str| {
        edges
            .values()
            .flat_map(|edges| edges.values())
            .filter(|edge| edge.target.as_deref() == Some(part))
            .count()
    };
    if !create
        && (incoming(&before_relations, drawing) != 1
            || owned.kind == "chart" && incoming(&before_relations, object_part) != 1)
        || !delete
            && (incoming(&relations, drawing) != 1
                || owned.kind == "chart" && incoming(&relations, object_part) != 1)
    {
        return Err(crate::coded(
            "package-protected",
            "Shared mutable object ownership is protected",
        ));
    }
    if owned.kind == "chart" {
        let data = part_bytes(zip, object_part, if delete { &empty } else { replacements })?;
        managed_chart_styles::check(&data)?;
        xml_events(&data, |_, local, _| {
            if [b"externalData".as_slice(), b"pivotSource", b"extLst"].contains(&local) {
                return Err(crate::coded(
                    "package-protected",
                    "Protected chart authoring",
                ));
            }
            Ok(())
        })?;
        if relations.contains_key(object_part) || before_relations.contains_key(object_part) {
            return Err(crate::coded(
                "package-protected",
                "Dependent chart graphs are protected",
            ));
        }
    }
    let fragment = plan["fragments"]
        .get(sheet)
        .and_then(|value| value.get("drawing"))
        .and_then(Value::as_str);
    let worksheet = if let Some(fragment) = fragment {
        fragment.as_bytes().to_vec()
    } else {
        part_bytes(&mut metadata, sheet, &empty)?
    };
    let mut links = Vec::new();
    if !worksheet.is_empty() {
        xml_events(&worksheet, |ns, local, values| {
            if ns == NS && local == b"drawing" {
                links.push(
                    values
                        .get("r:id")
                        .ok_or_else(|| {
                            crate::coded("internal-missing-field", "Missing sheet drawing edge")
                        })?
                        .clone(),
                );
            }
            Ok(())
        })?;
    }
    if links.len() > 1 {
        return Err(crate::coded(
            "package-ambiguous",
            "Ambiguous sheet drawing ownership",
        ));
    }
    if !delete || names.contains(drawing) {
        let edge = links
            .first()
            .and_then(|id| relations.get(sheet).and_then(|edges| edges.get(id)))
            .ok_or_else(|| {
                crate::coded(
                    "internal-missing-field",
                    "Missing owning worksheet drawing link",
                )
            })?;
        if edge.target.as_deref() != Some(drawing)
            || edge.kind != format!("{}/drawing", String::from_utf8_lossy(REL_NS))
        {
            return Err(crate::coded(
                "internal-mismatch",
                "Worksheet/drawing scope mismatch",
            ));
        }
    }
    for name in additions.keys() {
        if !relations.values().any(|edges| {
            edges.values().any(|edge| {
                edge.target.as_ref() == Some(name)
                    && edge.kind == format!("{}/image", String::from_utf8_lossy(REL_NS))
            })
        }) {
            return Err(crate::coded("package-unsupported", "Orphan PNG addition"));
        }
    }
    Ok(additions)
}
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};
fn valid_part(name: &str) -> Result<()> {
    // ZIP directory entries are legal package members and are copied unchanged.
    let name = name.strip_suffix('/').unwrap_or(name);
    if name.is_empty()
        || name.starts_with('/')
        || name.contains('\\')
        || name
            .split('/')
            .any(|s| s.is_empty() || s == "." || s == "..")
    {
        return Err(crate::coded("package-malformed", "Invalid package part"));
    }
    Ok(())
}
fn rank(name: &[u8]) -> usize {
    [
        b"sheetPr".as_slice(),
        b"dimension",
        b"sheetViews",
        b"sheetFormatPr",
        b"cols",
        b"sheetData",
        b"sheetCalcPr",
        b"sheetProtection",
        b"protectedRanges",
        b"scenarios",
        b"autoFilter",
        b"sortState",
        b"dataConsolidate",
        b"customSheetViews",
        b"mergeCells",
        b"phoneticPr",
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
        b"picture",
        b"oleObjects",
        b"controls",
        b"webPublishItems",
        b"tableParts",
        b"extLst",
    ]
    .iter()
    .position(|n| *n == name)
    .unwrap_or(usize::MAX)
}
fn worksheet<R: std::io::BufRead, W: Write>(
    source: R,
    out: &mut W,
    fragments: &Value,
    plan: &Value,
    db: Option<&rusqlite::Connection>,
    sheet: &str,
) -> Result<()> {
    let mut reader = NsReader::from_reader(source);
    let mut writer = Writer::new(out);
    let mut buffer = Vec::new();
    let (mut depth, mut skip) = (0usize, 0usize);
    let mut inserted = BTreeSet::new();
    let mut cell = None;
    let movement = if plan["sheet"] == sheet {
        tables::columns::change(plan)?
    } else {
        None
    };
    let mut header_inserted = false;
    let mut row_number = None;
    let mut fragments = fragments
        .as_object()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheet fragments"))?
        .iter()
        .collect::<Vec<_>>();
    for (tag, _) in &fragments {
        if ![
            "conditionalFormatting",
            "tableParts",
            "dataValidations",
            "hyperlinks",
            "legacyDrawing",
            "drawing",
            "picture",
            "sheetPr",
            "printOptions",
            "pageMargins",
            "pageSetup",
            "headerFooter",
        ]
        .contains(&tag.as_str())
        {
            return Err(crate::coded(
                "package-unsupported",
                "Unqualified package sheet fragment",
            ));
        }
        if tag.as_str() == "legacyDrawing"
            && !(["setNote", "removeNote"].contains(&field(plan, "kind")?))
        {
            return Err(crate::coded(
                "package-unsupported",
                "Legacy note fragments require native-owned note commands",
            ));
        }
        if tag.as_str() == "picture"
            && !(["setSheetBackground", "removeSheetBackground"].contains(&field(plan, "kind")?))
        {
            return Err(crate::coded(
                "package-unsupported",
                "Picture fragments require qualified background authoring",
            ));
        }
        if tag.as_str() == "drawing"
            && ![
                "createChart",
                "updateChart",
                "insertImage",
                "updateImage",
                "deleteObject",
                "createShape",
                "updateShape",
                "deleteShape",
                "createInkStroke",
                "updateInkStroke",
                "deleteInkStroke",
                "createSmartArt",
                "updateSmartArt",
                "deleteSmartArt",
                "createModel3D",
                "updateModel3D",
                "deleteModel3D",
            ]
            .contains(&field(plan, "kind")?)
        {
            return Err(crate::coded(
                "package-bad-input",
                "Drawing fragments require object authoring",
            ));
        }
        if [
            "sheetPr",
            "printOptions",
            "pageMargins",
            "pageSetup",
            "headerFooter",
        ]
        .contains(&tag.as_str())
            && plan["kind"] != "setPrintSettings"
        {
            return Err(crate::coded(
                "package-unsupported",
                "Print fragments require qualified print authoring",
            ));
        }
    }
    fragments.sort_by_key(|(tag, _)| rank(tag.as_bytes()));
    let bounds = if plan["kind"] == "createTable" {
        Some(crate::bulk::bounds(field(plan, "range")?)?)
    } else {
        None
    };
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
                let name = name.as_ref();
                if main && depth == 2 && name == b"row" {
                    header_inserted = false;
                    row_number = attrs(e)?
                        .get("r")
                        .and_then(|r| r.parse::<u32>().ok())
                        .and_then(|r| r.checked_sub(1));
                }
                if main && depth == 3 && name == b"c" {
                    cell = Some(position(attrs(e)?.get("r").ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing source cell address")
                    })?)?);
                }
                if main
                    && depth == 3
                    && name == b"c"
                    && let Some(c) = &movement
                {
                    let point = cell.ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing table cell")
                    })?;
                    if !c.remove
                        && point.0 == c.first.0
                        && !header_inserted
                        && point.1 >= c.first.1 + c.index as u32
                    {
                        writer
                            .get_mut()
                            .write_all(&tables::columns::header(
                                db.ok_or_else(|| {
                                    crate::coded(
                                        "package-bad-input",
                                        "Table column edit requires an indexed connection",
                                    )
                                })?,
                                c,
                            )?)
                            .map_err(err)?;
                        header_inserted = true;
                    }
                    let in_rows = point.0 >= c.first.0 && point.0 <= c.last.0;
                    let inside = in_rows && point.1 >= c.first.1 && point.1 <= c.last.1;
                    if inside && c.remove && point.1 == c.first.1 + c.index as u32
                        || in_rows && !c.remove && point.1 == c.last.1 + 1
                    {
                        if !empty {
                            skip = 1;
                        }
                        buffer.clear();
                        continue;
                    }
                    if inside && point.1 >= c.first.1 + c.index as u32 {
                        let next = if c.remove { point.1 - 1 } else { point.1 + 1 };
                        let mut values = attrs(e)?;
                        values.insert("r".into(), crate::xml::address(point.0, next));
                        let mut start = e.to_owned();
                        start.clear_attributes();
                        for (key, value) in &values {
                            start.push_attribute((key.as_str(), value.as_str()));
                        }
                        writer
                            .write_event(if empty {
                                Event::Empty(start)
                            } else {
                                Event::Start(start)
                            })
                            .map_err(err)?;
                        if !empty {
                            depth += 1;
                        }
                        buffer.clear();
                        continue;
                    }
                }
                // The independently admitted column change also updates its literal
                // header in this same package transaction, retaining cell attributes.
                if main
                    && depth == 3
                    && name == b"c"
                    && plan["kind"] == "renameTableColumn"
                    && plan["sheet"] == sheet
                {
                    let (first, _) = crate::bulk::bounds(field(plan, "priorRef")?)?;
                    let index =
                        u32::try_from(plan["renameTableColumn"]["index"].as_u64().ok_or_else(
                            || crate::coded("internal-missing-field", "Missing column index"),
                        )?)
                        .map_err(err)?;
                    if cell == Some((first.0, first.1 + index)) {
                        let mut attributes = attrs(e)?;
                        attributes.insert("t".into(), "inlineStr".into());
                        let mut start = e.to_owned();
                        start.clear_attributes();
                        for (key, value) in &attributes {
                            start.push_attribute((key.as_str(), value.as_str()));
                        }
                        writer
                            .write_event(Event::Start(start.clone()))
                            .map_err(err)?;
                        let label = field(&plan["renameTableColumn"], "new")?;
                        let text = format!(
                            "<is xmlns=\"{}\"><t xml:space=\"preserve\">{}</t></is>",
                            String::from_utf8_lossy(NS),
                            label
                                .replace('&', "&amp;")
                                .replace('<', "&lt;")
                                .replace('>', "&gt;")
                        );
                        writer.get_mut().write_all(text.as_bytes()).map_err(err)?;
                        writer
                            .write_event(Event::End(start.to_end()))
                            .map_err(err)?;
                        if !empty {
                            skip = 1;
                        }
                        buffer.clear();
                        continue;
                    }
                }
                if main && depth == 4 && name == b"f" && plan.get("rename").is_some() {
                    if attrs(e)?
                        .keys()
                        .any(|k| k != "xmlns" && !k.starts_with("xmlns:"))
                    {
                        return Err(crate::coded(
                            "package-unsupported",
                            "Special formulas require rename qualification",
                        ));
                    }
                    let c = cell.ok_or_else(|| {
                        crate::coded("package-unsupported", "Formula outside source cell")
                    })?;
                    let formula: String = db
                        .ok_or_else(|| {
                            crate::coded(
                                "package-bad-input",
                                "Rename requires an indexed connection",
                            )
                        })?
                        .query_row(
                            "SELECT formula FROM renamed_formulas WHERE sheet=?1 AND ref=?2",
                            rusqlite::params![sheet, crate::xml::address(c.0, c.1)],
                            |r| r.get(0),
                        )
                        .map_err(err)?;
                    writer.write_event(Event::Start(e.clone())).map_err(err)?;
                    writer
                        .write_event(Event::Text(BytesText::new(&formula)))
                        .map_err(err)?;
                    writer.write_event(Event::End(e.to_end())).map_err(err)?;
                    if !empty {
                        skip = 1;
                    }
                    buffer.clear();
                    continue;
                }
                if main
                    && depth == 4
                    && name == b"f"
                    && (plan.get("renameTable").is_some()
                        || (plan.get("renameTableColumn").is_some()
                            || plan.get("columnChange").is_some()))
                {
                    if attrs(e)?
                        .keys()
                        .any(|k| k != "xmlns" && !k.starts_with("xmlns:"))
                    {
                        return Err(crate::coded(
                            "package-unsupported",
                            "Special formulas require table rename qualification",
                        ));
                    }
                    let c = cell.ok_or_else(|| {
                        crate::coded("package-unsupported", "Table formula outside source cell")
                    })?;
                    use rusqlite::OptionalExtension;
                    let formula: Option<String> = db
                        .ok_or_else(|| {
                            crate::coded(
                                "package-bad-input",
                                "Rename requires an indexed connection",
                            )
                        })?
                        .query_row(
                            "SELECT formula FROM renamed_formulas WHERE sheet=?1 AND ref=?2",
                            rusqlite::params![sheet, crate::xml::address(c.0, c.1)],
                            |row| row.get(0),
                        )
                        .optional()
                        .map_err(err)?;
                    if let Some(formula) = formula {
                        writer.write_event(Event::Start(e.clone())).map_err(err)?;
                        writer
                            .write_event(Event::Text(BytesText::new(&formula)))
                            .map_err(err)?;
                        writer.write_event(Event::End(e.to_end())).map_err(err)?;
                        if !empty {
                            skip = 1;
                        }
                        buffer.clear();
                        continue;
                    }
                }
                if main
                    && name == b"f"
                    && let (Some((a, b)), Some(c)) = (bounds, cell)
                    && c.0 >= a.0
                    && c.0 <= b.0
                    && c.1 >= a.1
                    && c.1 <= b.1
                {
                    return Err(crate::coded(
                        "package-unsupported",
                        "Table creation over formulas is not qualified",
                    ));
                }
                if main
                    && depth == 0
                    && fragments
                        .iter()
                        .any(|(tag, _)| tag.as_str() == "tableParts")
                {
                    let a = attrs(e)?;
                    if a.get("xmlns:r").is_some_and(|v| {
                        v != "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    }) {
                        return Err(crate::coded(
                            "package-refused",
                            "Conflicting worksheet relationship namespace",
                        ));
                    }
                    let mut root = e.to_owned();
                    if !a.contains_key("xmlns:r") {
                        root.push_attribute((
                            "xmlns:r",
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                        ));
                    }
                    writer.write_event(Event::Start(root)).map_err(err)?;
                    depth += 1;
                    buffer.clear();
                    continue;
                }
                if main && depth == 1 {
                    for (tag, value) in &fragments {
                        if rank(tag.as_bytes()) <= rank(name) && inserted.insert((*tag).clone()) {
                            writer
                                .get_mut()
                                .write_all(
                                    value
                                        .as_str()
                                        .ok_or_else(|| {
                                            crate::coded(
                                                "internal-invalid-field",
                                                "Invalid fragment",
                                            )
                                        })?
                                        .as_bytes(),
                                )
                                .map_err(err)?;
                        }
                    }
                    if fragments.iter().any(|(tag, _)| tag.as_bytes() == name) {
                        if !empty {
                            skip = 1;
                        }
                        buffer.clear();
                        continue;
                    }
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
                if main
                    && depth == 3
                    && let Some(c) = &movement
                    && !c.remove
                    && !header_inserted
                    && row_number == Some(c.first.0)
                {
                    writer
                        .get_mut()
                        .write_all(&tables::columns::header(
                            db.ok_or_else(|| {
                                crate::coded(
                                    "package-bad-input",
                                    "Table column edit requires an indexed connection",
                                )
                            })?,
                            c,
                        )?)
                        .map_err(err)?;
                    header_inserted = true;
                }
                if main && depth == 1 {
                    for (tag, value) in &fragments {
                        if inserted.insert((*tag).clone()) {
                            writer
                                .get_mut()
                                .write_all(
                                    value
                                        .as_str()
                                        .ok_or_else(|| {
                                            crate::coded(
                                                "internal-invalid-field",
                                                "Invalid fragment",
                                            )
                                        })?
                                        .as_bytes(),
                                )
                                .map_err(err)?;
                        }
                    }
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    crate::coded("internal-invalid-field", "Invalid worksheet nesting")
                })?;
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
    if inserted.len() != fragments.len() {
        return Err(crate::coded(
            "internal-missing-field",
            "Missing worksheet metadata container",
        ));
    }
    Ok(())
}
pub(crate) fn verify_clear_hyperlinks_rewrite(
    root: &Path,
    command: &Value,
    output: &str,
) -> Result<()> {
    clear_hyperlinks::verify_receipt_package(root, command, output)
}
pub(crate) fn check_boolean_cell(root: &Path, input: &Value) -> Result<Value> {
    boolean_cell::check(root, input)
}
pub(crate) fn read_notes(root: &Path, input: &Value) -> Result<Value> {
    note_authoring::read(root, input)
}
pub(crate) fn verify_note_rewrite(root: &Path, command: &Value, output: &str) -> Result<()> {
    note_authoring::verify_receipt_package(root, command, output)
}
pub(crate) fn read_hyperlink(root: &Path, input: &Value) -> Result<Value> {
    hyperlink_authoring::read(root, input)
}
pub(crate) fn verify_hyperlink_rewrite(root: &Path, command: &Value, output: &str) -> Result<()> {
    hyperlink_authoring::verify_receipt_package(root, command, output)
}
pub(crate) fn verify_page_theme_rewrite(root: &Path, command: &Value, output: &str) -> Result<()> {
    theme_authoring::verify_receipt_package(root, command, output)
}
pub(crate) fn inspect_sheet_background(root: &Path, sheet: &str) -> Value {
    background_authoring::inspect(root, sheet)
}
pub(crate) fn verify_background_rewrite(root: &Path, command: &Value, output: &str) -> Result<()> {
    background_authoring::verify_receipt_package(root, command, output)
}
pub fn apply(root: &Path, input: &Value) -> Result<Value> {
    let note = if ["setNote", "removeNote"]
        .iter()
        .any(|kind| input["command"]["type"] == *kind)
    {
        Some(note_authoring::prepare(root, input)?)
    } else {
        None
    };
    let background = if ["setSheetBackground", "removeSheetBackground"]
        .iter()
        .any(|kind| input["command"]["type"] == *kind)
    {
        Some(background_authoring::prepare(root, input)?)
    } else {
        None
    };
    let theme = if input["command"]["type"] == "applyPageTheme" {
        Some(theme_authoring::prepare(root, input)?)
    } else {
        None
    };
    let hyperlink = if ["setHyperlink", "removeHyperlink"]
        .iter()
        .any(|kind| input["command"]["type"] == *kind)
    {
        Some(hyperlink_authoring::prepare(root, input)?)
    } else {
        None
    };
    let clear = if input["command"]["type"] == "clearHyperlinks" {
        Some(clear_hyperlinks::prepare(root, input)?)
    } else {
        if input.get("command").is_some()
            && hyperlink.is_none()
            && theme.is_none()
            && background.is_none()
            && note.is_none()
        {
            return Err(crate::coded(
                "internal-unknown-reference",
                "Unknown native package command",
            ));
        }
        None
    };
    let plan: Value = if let Some(prepared) = &note {
        prepared.plan.clone()
    } else if let Some(prepared) = &clear {
        prepared.plan.clone()
    } else if let Some(prepared) = &hyperlink {
        prepared.plan.clone()
    } else if let Some(prepared) = &theme {
        prepared.plan.clone()
    } else if let Some(prepared) = &background {
        prepared.plan.clone()
    } else {
        serde_json::from_slice(&fs::read(local(root, "package-metadata.json")?).map_err(err)?)
            .map_err(err)?
    };
    if ![
        "addSheet",
        "renameSheet",
        "duplicateSheet",
        "createTable",
        "renameTable",
        "resizeTable",
        "renameTableColumn",
        "insertTableColumn",
        "deleteTableColumn",
        "createPivot",
        "refreshPivot",
        "updatePivot",
        "setPivotLayout",
        "addConditionalRule",
        "defineName",
        "removeName",
        "updateName",
        "createChart",
        "updateChart",
        "insertImage",
        "updateImage",
        "deleteObject",
        "createShape",
        "updateShape",
        "deleteShape",
        "createInkStroke",
        "updateInkStroke",
        "deleteInkStroke",
        "createSmartArt",
        "updateSmartArt",
        "deleteSmartArt",
        "createModel3D",
        "updateModel3D",
        "deleteModel3D",
        "setPrintSettings",
        "setConditionalRules",
        "setValidationRules",
        "clearHyperlinks",
        "setHyperlink",
        "removeHyperlink",
        "setNote",
        "removeNote",
        "applyPageTheme",
        "setSheetBackground",
        "removeSheetBackground",
    ]
    .contains(&field(&plan, "kind")?)
    {
        return Err(crate::coded(
            "package-unsupported",
            "Unqualified package edit",
        ));
    }
    if ["setNote", "removeNote"].iter().any(|k| plan["kind"] == *k) && note.is_none() {
        return Err(crate::coded(
            "package-unsupported",
            "Note authoring requires a native-owned command",
        ));
    }
    if plan["kind"] == "clearHyperlinks" && clear.is_none() {
        return Err(crate::coded(
            "package-unsupported",
            "Clear Hyperlinks requires a native-owned command, never a metadata plan",
        ));
    }
    if ["setHyperlink", "removeHyperlink"]
        .iter()
        .any(|kind| plan["kind"] == *kind)
        && hyperlink.is_none()
    {
        return Err(crate::coded(
            "package-unsupported",
            "Hyperlink authoring requires a native-owned command",
        ));
    }
    if ["setSheetBackground", "removeSheetBackground"]
        .iter()
        .any(|kind| plan["kind"] == *kind)
        && background.is_none()
    {
        return Err(crate::coded(
            "package-unsupported",
            "Background writes require a native-owned command",
        ));
    }
    if plan["kind"] == "applyPageTheme" && theme.is_none() {
        return Err(crate::coded(
            "package-unsupported",
            "Page theme requires a native-owned command",
        ));
    }
    if plan["kind"] == "setPrintSettings" {
        print_settings::validate(root, &plan, input)?;
    }
    if ["updateName", "removeName"].contains(&field(&plan, "kind")?) {
        if !name_addition::legal_name(field(&plan, "name")?) {
            return Err(crate::coded(
                "package-unsupported",
                "Reserved or invalid static range name",
            ));
        }
        name_update::validate(root, &plan, input)?;
    }
    if plan["kind"] == "defineName" {
        name_addition::validate(root, &plan, input)?;
    }
    if ["setConditionalRules", "setValidationRules"].contains(&field(&plan, "kind")?) {
        rules::validate(root, &plan, input)?;
    }
    let db = crate::database(root)?;
    let rewrite = crate::used_range_transfer::admit_writer(
        root,
        &db,
        input,
        crate::used_range_transfer::WriterKind::Package,
    )?;
    if (clear.is_some()
        || hyperlink.is_some()
        || theme.is_some()
        || background.is_some()
        || note.is_some())
        && rewrite.is_none()
    {
        return Err(crate::coded(
            "package-bad-input",
            "Clear Hyperlinks requires a sealed native rewrite input",
        ));
    }
    if [
        "createPivot",
        "refreshPivot",
        "updatePivot",
        "setPivotLayout",
    ]
    .contains(&field(&plan, "kind")?)
    {
        pivots::validate(root, &db, &plan, input)?;
    }
    if ["insertTableColumn", "deleteTableColumn"].contains(&field(&plan, "kind")?)
        && rewrite.is_none()
    {
        return Err(crate::coded(
            "package-bad-input",
            "Table column structure requires a sealed native rewrite input",
        ));
    }
    if [
        "renameTable",
        "resizeTable",
        "renameTableColumn",
        "insertTableColumn",
        "deleteTableColumn",
    ]
    .contains(&field(&plan, "kind")?)
    {
        tables::validate(root, &db, &plan, input)?;
    }
    if let Some(rename) = plan.get("rename") {
        db.execute_batch("CREATE TEMP TABLE renamed_formulas(sheet TEXT,ref TEXT,formula TEXT,PRIMARY KEY(sheet,ref));").map_err(err)?;
        crate::calculation::snapshot(
            &db,
            input["version"]
                .as_u64()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing version"))?,
        )?;
        for sheet in plan["sheets"]
            .as_array()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing rename sheets"))?
        {
            let sheet = sheet
                .as_str()
                .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid rename sheet"))?;
            let mut after = None;
            loop {
                let page =
                    crate::calculation::page(&db, sheet, (0, 0), (1048575, 16383), after, true)?;
                for (reference, cell) in &page {
                    crate::cancel::check()?;
                    if !cell["formulaType"].is_null() {
                        return Err(crate::coded(
                            "package-unsupported",
                            "Special formula rename is not qualified",
                        ));
                    }
                    let mut formula = crate::formula::rename(
                        field(cell, "formula")?,
                        field(rename, "old")?,
                        field(rename, "new")?,
                    )?;
                    // A copied sheet's tables have new names; its formulas follow them.
                    for table in plan["tableRenames"].as_array().into_iter().flatten() {
                        formula = crate::formula::rename_table(
                            &formula,
                            field(table, "old")?,
                            field(table, "new")?,
                        )?;
                    }
                    db.execute(
                        "INSERT INTO renamed_formulas VALUES(?1,?2,?3)",
                        rusqlite::params![sheet, reference, formula],
                    )
                    .map_err(err)?;
                }
                if page.len() < 512 {
                    break;
                }
                after = Some(position(&page.last().unwrap().0)?);
            }
        }
    }
    let replacements = plan["replacements"]
        .as_object()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing replacement parts"))?;
    let copies = plan["copies"]
        .as_object()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing copies"))?;
    let removed = plan["removed"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing removed parts"))?;
    let mut zip = ZipArchive::new(crate::cancel::Reader(
        File::open(local(root, field(input, "source")?)?).map_err(err)?,
    ))
    .map_err(err)?;
    let additions = if let Some(background) = &background {
        background
            .additions
            .clone()
            .into_iter()
            .collect::<BTreeMap<_, _>>()
    } else {
        object_additions(root, &plan, &mut zip)?
    };
    let partial = local(root, "package.partial")?;
    let output = local(root, "package.xlsx")?;
    let outcome: Result<(Value, Option<crate::used_range_transfer::PreparedOutput>)> = (|| {
        let mut writer = ZipWriter::new(crate::quota::create(&partial).map_err(err)?);
        for i in 0..zip.len() {
            crate::cancel::check()?;
            let member = zip.by_index(i).map_err(err)?;
            let name = member.name().to_owned();
            valid_part(&name)?;
            if replacements.contains_key(&name) || removed.iter().any(|v| v == &name) {
                continue;
            }
            if plan["fragments"].get(&name).is_some()
                || (plan.get("rename").is_some()
                    || (plan.get("renameTable").is_some()
                        || (plan.get("renameTableColumn").is_some()
                            || plan.get("columnChange").is_some())))
                    && plan["copyOnly"] != true
                    && plan["sheets"]
                        .as_array()
                        .is_some_and(|v| v.iter().any(|s| s == &name))
            {
                let empty = json!({});
                let fragments = plan["fragments"].get(&name).unwrap_or(&empty);
                writer
                    .start_file(
                        &name,
                        SimpleFileOptions::default()
                            .compression_method(zip::CompressionMethod::Deflated),
                    )
                    .map_err(err)?;
                worksheet(
                    BufReader::new(member),
                    &mut writer,
                    fragments,
                    &plan,
                    Some(&db),
                    &name,
                )?;
            } else {
                writer.raw_copy_file(member).map_err(err)?;
            }
        }
        for (name, data) in replacements {
            crate::cancel::check()?;
            valid_part(name)?;
            writer
                .start_file(
                    name,
                    SimpleFileOptions::default()
                        .compression_method(zip::CompressionMethod::Deflated),
                )
                .map_err(err)?;
            writer
                .write_all(
                    data.as_str()
                        .ok_or_else(|| {
                            crate::coded("internal-invalid-field", "Invalid replacement")
                        })?
                        .as_bytes(),
                )
                .map_err(err)?;
        }
        for (name, source) in copies {
            crate::cancel::check()?;
            valid_part(name)?;
            if zip.file_names().any(|n| n == name) {
                return Err(crate::coded(
                    "package-duplicate",
                    "Copied worksheet already exists",
                ));
            }
            let member =
                zip.by_name(source.as_str().ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing copy source")
                })?)
                .map_err(err)?;
            if plan["copyOnly"] == true && plan["sheet"] == *source {
                writer
                    .start_file(
                        name,
                        SimpleFileOptions::default()
                            .compression_method(zip::CompressionMethod::Deflated),
                    )
                    .map_err(err)?;
                worksheet(
                    BufReader::new(member),
                    &mut writer,
                    &json!({}),
                    &plan,
                    Some(&db),
                    source.as_str().ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing copy source")
                    })?,
                )?;
            } else {
                writer.raw_copy_file_rename(member, name).map_err(err)?;
            }
        }
        for (name, bytes) in &additions {
            crate::cancel::check()?;
            writer
                .start_file(
                    name,
                    SimpleFileOptions::default()
                        .compression_method(zip::CompressionMethod::Deflated),
                )
                .map_err(err)?;
            for chunk in bytes.chunks(32768) {
                crate::cancel::check()?;
                writer.write_all(chunk).map_err(err)?;
            }
        }
        writer.finish().map_err(err)?.sync_all().map_err(err)?;
        crate::cancel::check()?;
        fs::rename(&partial, &output).map_err(err)?;
        if let Some(note) = &note {
            note_authoring::verify_output(root, field(input, "source")?, "package.xlsx", note)?;
        }
        if let Some(clear) = &clear {
            clear_hyperlinks::verify_output(root, field(input, "source")?, "package.xlsx", clear)?;
        }
        if let Some(hyperlink) = &hyperlink {
            hyperlink_authoring::verify_output(
                root,
                field(input, "source")?,
                "package.xlsx",
                hyperlink,
            )?;
        }
        if let Some(theme) = &theme {
            theme_authoring::verify_output(root, field(input, "source")?, "package.xlsx", theme)?;
        }
        if let Some(background) = &background {
            background_authoring::verify_output(
                root,
                field(input, "source")?,
                "package.xlsx",
                background,
            )?;
        }
        let prepared = if let Some(rewrite) = rewrite {
            let mapping = if clear.is_some()
                || hyperlink.is_some()
                || theme.is_some()
                || background.is_some()
                || note.is_some()
            {
                crate::used_range_transfer::NativeMapping::Identity
            } else {
                rewrite_admission::package_mapping(root, input, &plan, "package.xlsx")?
            };
            let mut prepared = crate::used_range_transfer::prepare_output(
                root,
                &db,
                rewrite,
                mapping,
                "package.xlsx",
            )?;
            if let Some(note) = &note {
                crate::used_range_transfer::mark_note_output(&mut prepared, &note.command)?;
            }
            if let Some(clear) = &clear {
                crate::used_range_transfer::mark_clear_hyperlinks_output(
                    &mut prepared,
                    &clear.command,
                )?;
            }
            if let Some(hyperlink) = &hyperlink {
                crate::used_range_transfer::mark_hyperlink_output(
                    &mut prepared,
                    &hyperlink.command,
                )?;
            }
            if let Some(theme) = &theme {
                crate::used_range_transfer::mark_page_theme_output(&mut prepared, &theme.command)?;
            }
            if let Some(background) = &background {
                crate::used_range_transfer::mark_background_output(
                    &mut prepared,
                    &background.command,
                )?;
            }
            Some(prepared)
        } else {
            None
        };
        Ok((
            if let Some(note) = &note {
                note.result()
            } else if let Some(clear) = &clear {
                clear.result()
            } else if let Some(hyperlink) = &hyperlink {
                hyperlink.result()
            } else if let Some(theme) = &theme {
                theme.result()
            } else if let Some(background) = &background {
                background.result()
            } else {
                json!({"file":"package.xlsx"})
            },
            prepared,
        ))
    })();
    if outcome.is_err() {
        let _ = fs::remove_file(&partial);
    }
    let (mut result, prepared) = outcome?;
    drop(db);
    if let Some(prepared) = prepared {
        result["rewriteId"] = json!(crate::used_range_transfer::commit_output(root, prepared)?);
    }
    Ok(result)
}

/// Owner-only rule publication. The supplied cell-free metadata represents the
/// selected rule history; native validation independently admits the closed
/// collection command. Cells and untouched members remain streamed/copied.
pub(crate) fn apply_rules(root: &Path, input: &Value) -> Result<Value> {
    let source_name = field(input, "source")?;
    let metadata_name = field(input, "metadataSource")?;
    let output_name = field(input, "output")?;
    let allowed = [
        "metadata.xlsx",
        "export.xlsx",
        "layout.xlsx",
        "rules-a.xlsx",
        "rules-b.xlsx",
        "rules-meta-a.xlsx",
        "rules-meta-b.xlsx",
    ];
    if !allowed.contains(&source_name)
        || !["metadata.xlsx", "rules-meta-a.xlsx", "rules-meta-b.xlsx"].contains(&metadata_name)
        || ![
            "rules-a.xlsx",
            "rules-b.xlsx",
            "rules-meta-a.xlsx",
            "rules-meta-b.xlsx",
        ]
        .contains(&output_name)
        || output_name == source_name
        || output_name == metadata_name
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid owned rule artifact",
        ));
    }
    let plan = &input["plan"];
    if serde_json::to_vec(plan).map_err(err)?.len() > 2 * 1024 * 1024 {
        return Err(crate::coded("package-limit", "Rule plan exceeds its bound"));
    }
    let source = local(root, source_name)?;
    rules::validate_files(&local(root, metadata_name)?, &source, plan)?;
    let mut zip =
        ZipArchive::new(crate::cancel::Reader(File::open(source).map_err(err)?)).map_err(err)?;
    let partial = local(root, "rules.partial")?;
    let output = local(root, output_name)?;
    let result = (|| -> Result<()> {
        let mut writer = ZipWriter::new(crate::quota::create(&partial).map_err(err)?);
        let mut seen = BTreeSet::new();
        for i in 0..zip.len() {
            crate::cancel::check()?;
            let member = zip.by_index(i).map_err(err)?;
            let name = member.name().to_owned();
            valid_part(&name)?;
            if !seen.insert(name.clone()) {
                return Err(crate::coded(
                    "package-duplicate",
                    "Duplicate rule package part",
                ));
            }
            if let Some(replacement) = plan["replacements"].get(&name) {
                writer
                    .start_file(
                        &name,
                        SimpleFileOptions::default()
                            .compression_method(zip::CompressionMethod::Deflated),
                    )
                    .map_err(err)?;
                writer
                    .write_all(
                        replacement
                            .as_str()
                            .ok_or_else(|| {
                                crate::coded("internal-invalid-field", "Invalid rule replacement")
                            })?
                            .as_bytes(),
                    )
                    .map_err(err)?;
            } else if let Some(fragments) = plan["fragments"].get(&name) {
                writer
                    .start_file(
                        &name,
                        SimpleFileOptions::default()
                            .compression_method(zip::CompressionMethod::Deflated),
                    )
                    .map_err(err)?;
                worksheet(
                    BufReader::new(member),
                    &mut writer,
                    fragments,
                    plan,
                    None,
                    &name,
                )?;
            } else {
                writer.raw_copy_file(member).map_err(err)?;
            }
        }
        if plan["fragments"]
            .as_object()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing rule fragments"))?
            .keys()
            .any(|name| !seen.contains(name))
        {
            return Err(crate::coded(
                "package-bad-input",
                "Rule worksheet was not present",
            ));
        }
        writer.finish().map_err(err)?.sync_all().map_err(err)?;
        crate::cancel::check()?;
        fs::rename(&partial, &output).map_err(err)?;
        Ok(())
    })();
    if result.is_err() {
        let _ = fs::remove_file(partial);
    }
    result?;
    Ok(json!({"file":output_name}))
}
/// Existing style records and unknown metadata are immutable; only known tables append.
fn validate_preserved_styles(before: &[u8], after: &[u8]) -> Result<()> {
    use print_settings::Tree;
    let a = Tree::parse(before)?;
    let b = Tree::parse(after)?;
    a.root(NS, "styleSheet")?;
    b.root(NS, "styleSheet")?;
    fn same(a: &Tree, ai: usize, b: &Tree, bi: usize) -> bool {
        let x = &a.nodes[ai];
        let y = &b.nodes[bi];
        x.ns == y.ns
            && x.name == y.name
            && x.attrs == y.attrs
            && x.text == y.text
            && x.children.len() == y.children.len()
            && x.children
                .iter()
                .zip(&y.children)
                .all(|(ai, bi)| same(a, *ai, b, *bi))
    }
    if a.nodes[0].attrs != b.nodes[0].attrs || a.nodes[0].text != b.nodes[0].text {
        return Err(crate::coded(
            "package-stale",
            "Preserved style root metadata changed",
        ));
    }
    let mut seen = std::collections::BTreeSet::new();
    for ai in &a.nodes[0].children {
        let old = &a.nodes[*ai];
        let key = (old.ns.clone(), old.name.clone());
        if !seen.insert(key.clone()) {
            return Err(crate::coded(
                "package-ambiguous",
                "Ambiguous source style containers are protected",
            ));
        }
        let matches = b.nodes[0]
            .children
            .iter()
            .filter(|i| b.nodes[**i].ns == old.ns && b.nodes[**i].name == old.name)
            .copied()
            .collect::<Vec<_>>();
        if matches.len() != 1 {
            return Err(crate::coded(
                "package-unsupported",
                "Preservation styles lost or duplicated source metadata",
            ));
        }
        let bi = matches[0];
        let new = &b.nodes[bi];
        if old.ns == NS
            && ["fonts", "fills", "borders", "cellXfs", "numFmts"].contains(&old.name.as_str())
        {
            let attrs = |node: &print_settings::Node| {
                node.attrs
                    .iter()
                    .filter(|(key, _)| key.as_str() != "count")
                    .map(|(key, value)| (key.clone(), value.clone()))
                    .collect::<std::collections::BTreeMap<_, _>>()
            };
            if attrs(old) != attrs(new)
                || old.text != new.text
                || new.children.len() < old.children.len()
                || !old
                    .children
                    .iter()
                    .zip(&new.children)
                    .all(|(ai, bi)| same(&a, *ai, &b, *bi))
            {
                return Err(crate::coded(
                    "package-stale",
                    "Preservation styles changed an existing record",
                ));
            }
        } else if !same(&a, *ai, &b, bi) {
            return Err(crate::coded(
                "package-stale",
                "Preservation styles changed opaque metadata",
            ));
        }
    }
    for bi in &b.nodes[0].children {
        let new = &b.nodes[*bi];
        if !seen.contains(&(new.ns.clone(), new.name.clone()))
            && !(new.ns == NS && new.name == "numFmts")
        {
            return Err(crate::coded(
                "package-unsupported",
                "Preservation styles added unrelated metadata",
            ));
        }
    }
    Ok(())
}
