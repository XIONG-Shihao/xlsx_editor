use crate::xml::{address, attrs, position};
use crate::{Result, database, err, field, hash, local};
use quick_xml::{Reader, events::Event};
use rusqlite::{Connection, OptionalExtension, params};
use serde_json::{Value, json};
use std::fs::{self, File};
use std::io::{self, Read, Seek, SeekFrom, Write};
use std::path::Path;
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

const CACHE_C: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/chart";
const CACHE_A: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
#[derive(Default)]
struct CacheNode {
    ns: Vec<u8>,
    name: String,
    attrs: std::collections::BTreeMap<String, String>,
    text: String,
    children: Vec<usize>,
    start: usize,
    end: usize,
}
struct CacheTree {
    bytes: Vec<u8>,
    nodes: Vec<CacheNode>,
}
impl CacheTree {
    fn parse(bytes: &[u8]) -> Result<Self> {
        if bytes.len() > 2 * 1024 * 1024 {
            return Err(crate::coded(
                "chart-limit",
                "Managed chart cache XML exceeds its bound",
            ));
        }
        let source = std::str::from_utf8(bytes).map_err(err)?;
        if source.to_ascii_lowercase().contains("<!doctype")
            || source.to_ascii_lowercase().contains("<!entity")
        {
            return Err(crate::coded(
                "chart-malformed",
                "Chart cache XML entities are refused",
            ));
        }
        let mut reader = quick_xml::NsReader::from_reader(bytes);
        let mut buffer = Vec::new();
        let mut nodes: Vec<CacheNode> = Vec::new();
        let mut stack: Vec<usize> = Vec::new();
        let mut roots = 0;
        loop {
            crate::cancel::check()?;
            let start = reader.buffer_position() as usize;
            let (binding, event) = reader.read_resolved_event_into(&mut buffer).map_err(err)?;
            let ns = match binding {
                quick_xml::name::ResolveResult::Bound(ns) => ns.as_ref().to_vec(),
                _ => Vec::new(),
            };
            let end = reader.buffer_position() as usize;
            match event {
                Event::Start(ref element) | Event::Empty(ref element) => {
                    if stack.len() > 64 || nodes.len() > 10000 {
                        return Err(crate::coded(
                            "chart-limit",
                            "Chart cache XML node/depth bound exceeded",
                        ));
                    }
                    if stack.is_empty() {
                        roots += 1;
                        if roots != 1 {
                            return Err(crate::coded(
                                "chart-malformed",
                                "Chart cache XML has multiple roots",
                            ));
                        }
                    }
                    let name = String::from_utf8_lossy(element.local_name().as_ref()).into_owned();
                    let values = attrs(element)?.into_iter().collect();
                    let index = nodes.len();
                    nodes.push(CacheNode {
                        ns,
                        name,
                        attrs: values,
                        start,
                        end,
                        ..Default::default()
                    });
                    if let Some(parent) = stack.last() {
                        nodes[*parent].children.push(index);
                    }
                    if matches!(event, Event::Start(_)) {
                        stack.push(index);
                    }
                }
                Event::End(_) => {
                    let index = stack.pop().ok_or_else(|| {
                        crate::coded("chart-malformed", "Unbalanced chart cache XML")
                    })?;
                    nodes[index].end = end;
                }
                Event::Text(ref value) => {
                    let text = value.unescape().map_err(err)?;
                    if let Some(index) = stack.last() {
                        nodes[*index].text.push_str(&text)
                    } else if !text.trim().is_empty() {
                        return Err(crate::coded(
                            "chart-malformed",
                            "Text outside chart cache root",
                        ));
                    }
                }
                Event::CData(ref value) => {
                    if let Some(index) = stack.last() {
                        nodes[*index]
                            .text
                            .push_str(std::str::from_utf8(value).map_err(err)?)
                    } else {
                        return Err(crate::coded(
                            "chart-unsupported",
                            "CDATA outside chart cache root",
                        ));
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buffer.clear();
        }
        if roots != 1
            || !stack.is_empty()
            || nodes[0].ns != CACHE_C
            || nodes[0].name != "chartSpace"
        {
            return Err(crate::coded(
                "chart-malformed",
                "Invalid managed chart cache root",
            ));
        }
        let c = [
            "chartSpace",
            "chart",
            "title",
            "tx",
            "rich",
            "plotArea",
            "layout",
            "barChart",
            "lineChart",
            "barDir",
            "grouping",
            "ser",
            "idx",
            "order",
            "strRef",
            "numRef",
            "f",
            "strCache",
            "numCache",
            "ptCount",
            "pt",
            "v",
            "cat",
            "val",
            "axId",
            "catAx",
            "valAx",
            "scaling",
            "orientation",
            "axPos",
            "crossAx",
            "legend",
            "legendPos",
            "plotVisOnly",
            "dispBlanksAs",
            "spPr",
        ];
        for node in &nodes {
            if !["v", "f", "t"].contains(&node.name.as_str()) && !node.text.trim().is_empty() {
                return Err(crate::coded(
                    "chart-malformed",
                    "Unexpected chart XML mixed text",
                ));
            }
            if !(node.ns == CACHE_C && c.contains(&node.name.as_str())
                || node.ns == CACHE_A
                    && [
                        "bodyPr",
                        "lstStyle",
                        "p",
                        "r",
                        "t",
                        "solidFill",
                        "srgbClr",
                        "noFill",
                        "ln",
                        "prstDash",
                    ]
                    .contains(&node.name.as_str()))
            {
                return Err(crate::coded(
                    "chart-unsupported",
                    "Unqualified managed chart XML model",
                ));
            }
            let allowed =
                if node.ns == CACHE_A && ["srgbClr", "prstDash"].contains(&node.name.as_str()) {
                    "val"
                } else if node.ns == CACHE_A && node.name == "ln" {
                    "w"
                } else if node.ns == CACHE_C && node.name == "pt" {
                    "idx"
                } else if node.ns == CACHE_C
                    && [
                        "barDir",
                        "grouping",
                        "idx",
                        "order",
                        "ptCount",
                        "axId",
                        "orientation",
                        "axPos",
                        "crossAx",
                        "legendPos",
                        "plotVisOnly",
                        "dispBlanksAs",
                    ]
                    .contains(&node.name.as_str())
                {
                    "val"
                } else {
                    ""
                };
            if node
                .attrs
                .keys()
                .any(|key| key != "xmlns" && !key.starts_with("xmlns:") && key != allowed)
            {
                return Err(crate::coded(
                    "chart-protected",
                    "Extended chart attributes are protected",
                ));
            }
        }
        crate::package_edit::chart_style_check(bytes)?;
        Ok(Self {
            bytes: bytes.to_vec(),
            nodes,
        })
    }
    fn children(&self, index: usize, name: &str) -> Vec<usize> {
        self.nodes[index]
            .children
            .iter()
            .copied()
            .filter(|i| self.nodes[*i].ns == CACHE_C && self.nodes[*i].name == name)
            .collect()
    }
    fn one(&self, index: usize, name: &str) -> Result<usize> {
        let found = self.children(index, name);
        if found.len() != 1 {
            return Err(crate::coded(
                "chart-ambiguous",
                format!("Ambiguous chart cache {name}"),
            ));
        }
        Ok(found[0])
    }
    fn val(&self, index: usize) -> Result<&str> {
        self.nodes[index]
            .attrs
            .get("val")
            .map(String::as_str)
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing chart model value"))
    }
    fn envelope(&self, caches: &[usize]) -> Result<Vec<u8>> {
        let mut spans = caches
            .iter()
            .map(|i| (self.nodes[*i].start, self.nodes[*i].end))
            .collect::<Vec<_>>();
        spans.sort();
        let mut bytes = Vec::new();
        let mut previous = 0;
        for (start, end) in spans {
            if start < previous || end < start || end > self.bytes.len() {
                return Err(crate::coded(
                    "chart-conflict",
                    "Overlapping chart cache boundaries",
                ));
            }
            bytes.extend_from_slice(&self.bytes[previous..start]);
            previous = end;
        }
        bytes.extend_from_slice(&self.bytes[previous..]);
        Ok(bytes)
    }
}
#[derive(Clone, PartialEq, Eq)]
struct ChartReference {
    sheet: String,
    first: (u32, u32),
    last: (u32, u32),
}
fn chart_reference(value: &str) -> Result<ChartReference> {
    let mut chars = value.chars().peekable();
    if chars.next() != Some('\'') {
        return Err(crate::coded(
            "chart-bad-input",
            "Chart references must quote their worksheet",
        ));
    }
    let mut sheet = String::new();
    loop {
        let ch = chars.next().ok_or_else(|| {
            crate::coded("chart-unsupported", "Incomplete chart worksheet reference")
        })?;
        if ch == '\'' {
            if chars.peek() == Some(&'\'') {
                chars.next();
                sheet.push('\'');
            } else {
                if chars.next() != Some('!') {
                    return Err(crate::coded(
                        "internal-invalid-field",
                        "Invalid chart worksheet reference",
                    ));
                }
                break;
            }
        } else {
            sheet.push(ch)
        }
    }
    let address = chars.collect::<String>();
    let mut fields = address.split(':');
    let parse = |field: &str| -> Result<(u32, u32)> {
        if !field.starts_with('$') {
            return Err(crate::coded(
                "chart-bad-input",
                "Chart source must use absolute references",
            ));
        }
        let mut parts = field[1..].split('$');
        let column = parts
            .next()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing chart column"))?;
        let row = parts
            .next()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing chart row"))?;
        if parts.next().is_some()
            || column.is_empty()
            || !column.bytes().all(|b| b.is_ascii_uppercase())
            || !row.bytes().all(|b| b.is_ascii_digit())
        {
            return Err(crate::coded(
                "chart-unsupported",
                "Unqualified chart source reference",
            ));
        }
        position(&format!("{column}{row}"))
    };
    let first = parse(
        fields
            .next()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing chart range"))?,
    )?;
    let last = fields.next().map(parse).transpose()?.unwrap_or(first);
    if fields.next().is_some() || first.0 > last.0 || first.1 > last.1 {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid chart source range",
        ));
    }
    Ok(ChartReference { sheet, first, last })
}
struct CacheSeries {
    header: ChartReference,
    category: ChartReference,
    values: ChartReference,
    caches: [usize; 3],
}
fn chart_cache_model(tree: &CacheTree, sheet: &str) -> Result<Vec<CacheSeries>> {
    let chart = tree.one(0, "chart")?;
    let plot = tree.one(chart, "plotArea")?;
    let bars = tree.children(plot, "barChart");
    let lines = tree.children(plot, "lineChart");
    if bars.len() + lines.len() != 1 {
        return Err(crate::coded(
            "chart-unsupported",
            "Only ordinary 2D charts are cache-qualified",
        ));
    }
    let chart_type = bars
        .first()
        .or_else(|| lines.first())
        .copied()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing chart plot"))?;
    let horizontal = !bars.is_empty() && tree.val(tree.one(chart_type, "barDir")?)? == "bar";
    if !bars.is_empty() && !["col", "bar"].contains(&tree.val(tree.one(chart_type, "barDir")?)?)
        || tree.val(tree.one(chart_type, "grouping")?)?
            != if bars.is_empty() {
                "standard"
            } else {
                "clustered"
            }
    {
        return Err(crate::coded(
            "chart-unsupported",
            "Unqualified chart grouping/direction",
        ));
    }
    if tree.val(tree.one(chart, "plotVisOnly")?)? != "1"
        || tree.val(tree.one(chart, "dispBlanksAs")?)? != "gap"
    {
        return Err(crate::coded(
            "chart-unsupported",
            "Unqualified chart visibility/blank policy",
        ));
    }
    let category_axis = tree.one(plot, "catAx")?;
    let value_axis = tree.one(plot, "valAx")?;
    let ids = [
        tree.val(tree.one(category_axis, "axId")?)?,
        tree.val(tree.one(value_axis, "axId")?)?,
    ];
    if ids[0] == ids[1]
        || ids
            .iter()
            .any(|id| id.parse::<u32>().is_err() || *id == "0")
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid chart primary axes",
        ));
    }
    let plot_ids = tree.children(chart_type, "axId");
    if plot_ids.len() != 2
        || tree.val(plot_ids[0])? == tree.val(plot_ids[1])?
        || !plot_ids
            .iter()
            .all(|id| ids.contains(&tree.val(*id).unwrap_or("")))
    {
        return Err(crate::coded("chart-unsupported", "Unbound chart plot axes"));
    }
    for (index, axis) in [category_axis, value_axis].iter().enumerate() {
        if tree.val(tree.one(*axis, "crossAx")?)? != ids[1 - index]
            || tree.val(tree.one(tree.one(*axis, "scaling")?, "orientation")?)? != "minMax"
            || tree.val(tree.one(*axis, "axPos")?)?
                != if horizontal {
                    if index == 0 { "l" } else { "b" }
                } else if index == 0 {
                    "b"
                } else {
                    "l"
                }
        {
            return Err(crate::coded(
                "chart-protected",
                "Alternative chart axes are protected",
            ));
        }
    }
    let series = tree.children(chart_type, "ser");
    if series.is_empty() || series.len() > 8 {
        return Err(crate::coded(
            "chart-limit",
            "Chart series count exceeds its bound",
        ));
    }
    let mut result = Vec::new();
    for (index, series) in series.iter().enumerate() {
        if tree.val(tree.one(*series, "idx")?)? != index.to_string()
            || tree.val(tree.one(*series, "order")?)? != index.to_string()
        {
            return Err(crate::coded(
                "chart-unsupported",
                "Chart series order is not qualified",
            ));
        }
        let refs = [
            tree.one(tree.one(*series, "tx")?, "strRef")?,
            tree.one(tree.one(*series, "cat")?, "strRef")?,
            tree.one(tree.one(*series, "val")?, "numRef")?,
        ];
        let header = chart_reference(&tree.nodes[tree.one(refs[0], "f")?].text)?;
        let category = chart_reference(&tree.nodes[tree.one(refs[1], "f")?].text)?;
        let values = chart_reference(&tree.nodes[tree.one(refs[2], "f")?].text)?;
        if header.sheet != sheet
            || category.sheet != sheet
            || values.sheet != sheet
            || header.first != header.last
            || header.first.0.checked_add(1) != Some(values.first.0)
            || header.first.1 != values.first.1
            || category.first.1 != category.last.1
            || values.first.1 != values.last.1
            || category.first.0 != values.first.0
            || category.last.0 != values.last.0
            || category.first.1.checked_add(index as u32 + 1) != Some(values.first.1)
        {
            return Err(crate::coded(
                "chart-bad-input",
                "Chart source model is not a same-sheet rectangle",
            ));
        }
        if let Some(previous) = result.first() {
            let previous: &CacheSeries = previous;
            if previous.category != category {
                return Err(crate::coded(
                    "internal-mismatch",
                    "Chart categories disagree across series",
                ));
            }
        }
        result.push(CacheSeries {
            header,
            category,
            values,
            caches: [
                tree.one(refs[0], "strCache")?,
                tree.one(refs[1], "strCache")?,
                tree.one(refs[2], "numCache")?,
            ],
        });
    }
    let first = &result[0];
    let rows = u64::from(first.values.last.0 - first.header.first.0 + 1);
    if rows * (result.len() as u64 + 1) > 1000 {
        return Err(crate::coded(
            "chart-limit",
            "Chart source exceeds 1,000 positions",
        ));
    }
    if tree
        .nodes
        .iter()
        .filter(|n| n.ns == CACHE_C && ["strCache", "numCache", "f"].contains(&n.name.as_str()))
        .count()
        != result.len() * 6
    {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown chart reference/cache graph is protected",
        ));
    }
    Ok(result)
}
fn cache_points(
    tree: &CacheTree,
    cache: usize,
    count: usize,
    numeric: bool,
) -> Result<std::collections::BTreeMap<usize, String>> {
    let count_node = tree.one(cache, "ptCount")?;
    if !tree.nodes[count_node].children.is_empty() {
        return Err(crate::coded(
            "chart-protected",
            "Nested chart cache count is protected",
        ));
    }
    let declared = tree.val(count_node)?.parse::<usize>().map_err(err)?;
    if declared != count || count > 1000 {
        return Err(crate::coded(
            "internal-mismatch",
            "Chart cache count disagrees with source",
        ));
    }
    let mut points = std::collections::BTreeMap::new();
    let mut previous = None;
    for index in tree.children(cache, "pt") {
        let offset = tree.nodes[index]
            .attrs
            .get("idx")
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing chart cache index"))?
            .parse::<usize>()
            .map_err(err)?;
        if offset >= count || previous.is_some_and(|old| offset <= old) {
            return Err(crate::coded(
                "chart-refused",
                "Chart cache indices are unordered/duplicated",
            ));
        }
        let value = tree.one(index, "v")?;
        if tree.nodes[index].children.len() != 1 {
            return Err(crate::coded(
                "internal-unknown-reference",
                "Unknown chart cache point fields",
            ));
        }
        if !tree.nodes[value].children.is_empty() {
            return Err(crate::coded(
                "session-protected",
                "Nested cache values are protected",
            ));
        }
        let text = tree.nodes[value].text.clone();
        if text.encode_utf16().count() > if numeric { 64 } else { 256 } {
            return Err(crate::coded(
                "chart-limit",
                "Chart cache value exceeds its bound",
            ));
        }
        if numeric && !text.parse::<f64>().map_err(err)?.is_finite() {
            return Err(crate::coded(
                "chart-unsupported",
                "Nonfinite numeric chart cache",
            ));
        }
        points.insert(offset, text);
        previous = Some(offset);
    }
    if tree.nodes[cache].children.len() != points.len() + 1 || !numeric && points.len() != count {
        return Err(crate::coded(
            "chart-unsupported",
            "Incomplete/unknown chart cache fields",
        ));
    }
    Ok(points)
}
fn chart_value(cell: &Value, formula: bool) -> Result<Value> {
    if cell["dataType"] == "error"
        || !cell["error"].is_null()
        || cell["scenarioResult"] == true
        || !cell["formulaType"].is_null()
    {
        return Err(crate::coded(
            "chart-protected",
            "Protected/error chart source cell",
        ));
    }
    if !formula && cell.get("formula").is_some() {
        return Err(crate::coded(
            "chart-bad-input",
            "Chart headers/categories must be literal",
        ));
    }
    let value = crate::bulk::literal(cell, true)?;
    if cell.get("formula").is_none()
        && cell["rawValue"].as_str().is_some_and(|raw| {
            raw.trim_start_matches(['+', '-'])
                .bytes()
                .all(|byte| byte.is_ascii_digit())
        })
        && value
            .as_f64()
            .is_some_and(|number| number.abs() > 9_007_199_254_740_991.0)
    {
        return Err(crate::coded(
            "chart-unsupported",
            "Unsafe raw integer chart source",
        ));
    }
    if value.as_object().is_some_and(|map| {
        map.len() == 1 && map.get("$workbookNumber").and_then(Value::as_str) == Some("-0")
    }) {
        return Ok(json!(-0.0));
    }
    if value.is_object() {
        return Err(crate::coded(
            "chart-unsupported",
            "Unsafe chart numeric source",
        ));
    }
    Ok(value)
}
pub(crate) fn checked_chart_caches(
    root: &Path,
    db: &Connection,
    version: u64,
    metadata: &Value,
) -> Result<std::collections::BTreeMap<String, String>> {
    use sha2::{Digest, Sha256};
    use std::collections::{BTreeMap, BTreeSet};
    let Some(caches) = metadata.get("managedChartCaches") else {
        return Ok(BTreeMap::new());
    };
    let caches = caches
        .as_object()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid managed chart cache map"))?;
    if caches.len() > 16 {
        return Err(crate::coded(
            "chart-limit",
            "Managed chart cache count exceeds its bound",
        ));
    }
    let mut cache_bytes = 0usize;
    for value in caches.values() {
        cache_bytes = cache_bytes
            .checked_add(field(value, "xml")?.len())
            .ok_or_else(|| crate::coded("chart-limit", "Chart cache byte accounting overflow"))?;
    }
    if cache_bytes > 4 * 1024 * 1024 {
        return Err(crate::coded(
            "chart-limit",
            "Managed chart caches exceed the 4 MiB export bound",
        ));
    }
    let requested = caches.keys().cloned().collect::<BTreeSet<_>>();
    let contexts = crate::package_edit::managed_chart_export_context(root, &requested)?;
    let mut replacements = BTreeMap::new();
    for (part, descriptor) in caches {
        crate::cancel::check()?;
        if descriptor.as_object().is_none_or(|fields| {
            fields.len() != 2 || !fields.contains_key("xml") || !fields.contains_key("sourceSha256")
        }) || metadata["replacements"].get(part).is_some()
            || metadata["removed"]
                .as_array()
                .is_some_and(|parts| parts.iter().any(|value| value.as_str() == Some(part)))
        {
            return Err(crate::coded(
                "chart-duplicate",
                "Duplicate/unqualified chart cache claim",
            ));
        }
        let context = contexts.get(part).ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Managed chart cache ownership is missing",
            )
        })?;
        let source = &context.2;
        let expected = field(descriptor, "sourceSha256")?;
        if expected.len() != 64 || format!("{:x}", Sha256::digest(source)) != expected {
            return Err(crate::coded(
                "chart-stale",
                "Managed chart cache source fingerprint changed",
            ));
        }
        let xml = field(descriptor, "xml")?;
        let old = CacheTree::parse(source)?;
        let new = CacheTree::parse(xml.as_bytes())?;
        let original = chart_cache_model(&old, &context.1)?;
        let proposed = chart_cache_model(&new, &context.1)?;
        let old_caches = original
            .iter()
            .flat_map(|series| series.caches)
            .collect::<Vec<_>>();
        let new_caches = proposed
            .iter()
            .flat_map(|series| series.caches)
            .collect::<Vec<_>>();
        if old.envelope(&old_caches)? != new.envelope(&new_caches)? {
            return Err(crate::coded(
                "chart-stale",
                "Chart cache export changed non-cache XML",
            ));
        }
        let first = &original[0];
        let last = &original[original.len() - 1];
        let range = format!(
            "{}:{}",
            address(first.header.first.0, first.category.first.1),
            address(last.values.last.0, last.values.last.1)
        );
        let read = crate::range::read_public_db(
            db,
            &json!({"version":version,"sheet":context.0,"range":range,"offset":0,"limit":1001}),
        )?;
        let mut cells = BTreeMap::new();
        for value in read["cells"]
            .as_array()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing chart source cells"))?
        {
            cells.insert(
                value[0]
                    .as_str()
                    .ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing chart cell address")
                    })?
                    .to_string(),
                value[1].clone(),
            );
        }
        if cells.len() > 1000 || !read["nextOffset"].is_null() {
            return Err(crate::coded(
                "chart-limit",
                "Chart source read exceeded its bound",
            ));
        }
        for (index, series) in proposed.iter().enumerate() {
            let rows = (series.values.last.0 - series.values.first.0 + 1) as usize;
            let headers = cache_points(&new, series.caches[0], 1, false)?;
            let categories = cache_points(&new, series.caches[1], rows, false)?;
            let values = cache_points(&new, series.caches[2], rows, true)?;
            let blank = json!({"value":null});
            let header = chart_value(
                cells
                    .get(&address(series.header.first.0, series.header.first.1))
                    .unwrap_or(&blank),
                false,
            )?;
            let header = header.as_str().ok_or_else(|| {
                crate::coded("chart-bad-input", "Chart series header needs literal text")
            })?;
            if header.trim().is_empty()
                || header.encode_utf16().count() > 128
                || headers.get(&0).map(String::as_str) != Some(header)
            {
                return Err(crate::coded(
                    "internal-mismatch",
                    "Chart title cache differs from accepted source",
                ));
            }
            for offset in 0..rows {
                let category = chart_value(
                    cells
                        .get(&address(
                            series.category.first.0 + offset as u32,
                            series.category.first.1,
                        ))
                        .unwrap_or(&blank),
                    false,
                )?;
                let label = categories.get(&offset).ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing chart category cache")
                })?;
                let matches = if category.is_null() {
                    label.is_empty()
                } else if let Some(text) = category.as_str() {
                    text == label
                } else if let Some(number) = category.as_f64() {
                    number.is_finite()
                        && label
                            .parse::<f64>()
                            .is_ok_and(|value| value.to_bits() == number.to_bits())
                } else {
                    false
                };
                if !matches {
                    return Err(crate::coded(
                        "internal-mismatch",
                        "Chart category cache differs from accepted source",
                    ));
                }
                let value = chart_value(
                    cells
                        .get(&address(
                            series.values.first.0 + offset as u32,
                            series.values.first.1,
                        ))
                        .unwrap_or(&blank),
                    true,
                )?;
                if value.is_null() {
                    if values.contains_key(&offset) {
                        return Err(crate::coded(
                            "chart-refused",
                            "Blank chart source has a numeric cache point",
                        ));
                    }
                } else {
                    let number = value.as_f64().ok_or_else(|| {
                        crate::coded(
                            "chart-bad-input",
                            "Chart values need accepted numeric results",
                        )
                    })?;
                    if !number.is_finite()
                        || values
                            .get(&offset)
                            .and_then(|text| text.parse::<f64>().ok())
                            .is_none_or(|value| value.to_bits() != number.to_bits())
                    {
                        return Err(crate::coded(
                            "internal-mismatch",
                            "Chart numeric cache differs from accepted source",
                        ));
                    }
                }
            }
            let _ = index;
        }
        replacements.insert(part.clone(), xml.to_string());
    }
    Ok(replacements)
}

struct Patch {
    start: u64,
    end: u64,
    text: Vec<u8>,
}
fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\r', "&#13;")
}
fn text(value: &str) -> Result<String> {
    if value
        .chars()
        .any(|c| c == '\u{fffe}' || c == '\u{ffff}' || c < ' ' && !['\t', '\n', '\r'].contains(&c))
    {
        return Err(crate::coded(
            "package-malformed",
            "Invalid XML text character",
        ));
    }
    let mut result = String::new();
    let bytes = value.as_bytes();
    let mut start = 0;
    for i in 0..bytes.len().saturating_sub(6) {
        if bytes[i] == b'_'
            && bytes[i + 1] == b'x'
            && bytes[i + 6] == b'_'
            && bytes[i + 2..i + 6].iter().all(u8::is_ascii_hexdigit)
        {
            result.push_str(&value[start..i]);
            result.push_str("_x005F_");
            start = i + 1;
        }
    }
    result.push_str(&value[start..]);
    Ok(escape(&result))
}
fn literal(value: &Value) -> Result<(&'static str, String)> {
    if value.get("$workbookNumber").and_then(Value::as_str) == Some("-0") {
        return Ok((
            "n",
            format!(
                "<v xmlns=\"{}\">-0</v>",
                String::from_utf8_lossy(crate::xml::NS)
            ),
        ));
    }
    if let Some(formula) = value.get("$formula").and_then(Value::as_str) {
        let ns = String::from_utf8_lossy(crate::xml::NS);
        let mut content = if value["$cache"] == true {
            String::new()
        } else {
            format!("<f xmlns=\"{ns}\">{}</f>", escape(formula))
        };
        let calculation = &value["calculation"];
        let (kind, cached) = match calculation["status"].as_str() {
            Some("error") => (
                "e",
                Some(escape(calculation["error"].as_str().unwrap_or("#VALUE!"))),
            ),
            Some("calculated") => match &calculation["value"] {
                Value::String(s) => ("str", Some(text(s)?)),
                Value::Bool(b) => ("b", Some(u8::from(*b).to_string())),
                Value::Number(n) => ("", Some(n.to_string())),
                Value::Object(map)
                    if map.get("$workbookNumber").and_then(Value::as_str) == Some("-0") =>
                {
                    ("", Some("-0".into()))
                }
                Value::Null => ("", Some(String::new())),
                _ => {
                    return Err(crate::coded(
                        "internal-invalid-field",
                        "Invalid calculated value",
                    ));
                }
            },
            _ => ("", None),
        };
        if let Some(cached) = cached {
            content.push_str(&format!("<v xmlns=\"{ns}\">{cached}</v>"));
        }
        return Ok((kind, content));
    }
    Ok(match value {
        Value::Null => ("", String::new()),
        Value::Bool(b) => (
            "b",
            format!(
                "<v xmlns=\"{}\">{}</v>",
                String::from_utf8_lossy(crate::xml::NS),
                u8::from(*b)
            ),
        ),
        Value::Number(n) => (
            "n",
            format!(
                "<v xmlns=\"{}\">{n}</v>",
                String::from_utf8_lossy(crate::xml::NS)
            ),
        ),
        Value::String(s) => (
            "inlineStr",
            format!(
                "<is xmlns=\"{}\"><t xml:space=\"preserve\">{}</t></is>",
                String::from_utf8_lossy(crate::xml::NS),
                text(s)?
            ),
        ),
        _ => {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid literal cell value",
            ));
        }
    })
}
fn edit_literal(change: &Value, value: &Value) -> Result<(&'static str, String)> {
    if let Some(kind) = change.get("valueType") {
        if kind != "error"
            || change.get("formula").is_some()
            || !value.as_str().is_some_and(crate::bulk::values::error_code)
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid declared error value",
            ));
        }
        return Ok((
            "e",
            format!(
                "<v xmlns=\"{}\">{}</v>",
                String::from_utf8_lossy(crate::xml::NS),
                escape(value.as_str().unwrap())
            ),
        ));
    }
    literal(value)
}
fn spill_literal(change: &Value) -> Result<(&'static str, String)> {
    let ns = String::from_utf8_lossy(crate::xml::NS);
    let metadata = &change["$spillMetadata"];
    let formula = change["formula"]
        .as_str()
        .ok_or_else(|| crate::coded("spill-bad-input", "Spill master has no formula"))?;
    let formula = spill_storage_formula(formula);
    let range = metadata["range"].as_str().ok_or_else(|| {
        crate::coded(
            "spill-unsupported",
            "Spill master has no qualified output rectangle",
        )
    })?;
    let mut attributes = format!(" t=\"array\" ref=\"{}\"", escape(range));
    for key in ["aca", "ca"] {
        if metadata[key] == "1" {
            attributes.push_str(&format!(" {key}=\"1\""));
        }
    }
    let content = format!("<f xmlns=\"{ns}\"{attributes}>{}</f>", escape(&formula));
    if metadata.get("vm").is_some() {
        if change["calculation"]["status"] != "error" || change["calculation"]["error"] != "#SPILL!"
        {
            return Err(crate::coded(
                "spill-unsupported",
                "Rich spill cache has no typed SPILL result",
            ));
        }
        Ok(("e", format!("{content}<v xmlns=\"{ns}\">#VALUE!</v>")))
    } else {
        let value = change["calculation"]["value"]
            .as_f64()
            .filter(|value| value.is_finite())
            .ok_or_else(|| {
                crate::coded(
                    "spill-bad-input",
                    "Spill master has no finite numeric result",
                )
            })?;
        Ok(("n", format!("{content}<v xmlns=\"{ns}\">{value}</v>")))
    }
}
pub(crate) fn spill_storage_formula(formula: &str) -> String {
    // Future-function OOXML spelling is storage syntax. Preserve an existing
    // prefix; insert it only at the proved root constructor token, retaining
    // authored argument, whitespace, and grouping bytes.
    let mut start = usize::from(formula.starts_with('='));
    while formula
        .as_bytes()
        .get(start)
        .is_some_and(|c| c.is_ascii_whitespace() || *c == b'(')
    {
        start += 1;
    }
    let end = start
        + formula[start..]
            .bytes()
            .take_while(|c| c.is_ascii_alphanumeric() || *c == b'_' || *c == b'.')
            .count();
    if formula[start..end].eq_ignore_ascii_case("SEQUENCE") {
        format!("{}_xlfn.{}", &formula[..start], &formula[start..])
    } else {
        formula.to_owned()
    }
}
fn fragment(file: &mut File, start: u64, end: u64) -> Result<Vec<u8>> {
    file.seek(SeekFrom::Start(start)).map_err(err)?;
    let mut bytes = vec![0; usize::try_from(end - start).map_err(err)?];
    file.read_exact(&mut bytes).map_err(err)?;
    Ok(bytes)
}
fn patch_cell(
    db: &Connection,
    file: &mut File,
    sheet: &str,
    row: u32,
    col: u32,
    change: &Value,
) -> Result<Option<Patch>> {
    let existing: Option<(u64, u64, u64, u64, String)> = db
        .query_row(
            "SELECT start,open,close,end,spans FROM cells WHERE sheet=?1 AND r=?2 AND c=?3",
            params![sheet, row, col],
            |r| Ok((r.get(0)?, r.get(1)?, r.get(2)?, r.get(3)?, r.get(4)?)),
        )
        .optional()
        .map_err(err)?;
    let Some((start, open, close, end, spans)) = existing else {
        return Ok(None);
    };
    let bytes = fragment(file, start, end)?;
    let mut reader = Reader::from_reader(bytes.as_slice());
    let element = match reader.read_event().map_err(err)? {
        Event::Start(e) | Event::Empty(e) => e,
        _ => return Err(crate::coded("workbook-malformed", "Invalid indexed cell")),
    };
    let name = std::str::from_utf8(element.name().as_ref())
        .map_err(err)?
        .to_owned();
    let mut attributes = attrs(&element)?;
    let replace = change.get("value").is_some() || change.get("formula").is_some();
    let cache = change.get("cache").is_some();
    let value = if cache {
        json!({"$formula":"","$cache":true,"calculation":change["cache"]})
    } else if change.get("formula").is_some() {
        json!({"$formula":change["formula"],"calculation":change["calculation"]})
    } else {
        change["value"].clone()
    };
    let (kind, content) = if change.get("$spillMetadata").is_some() {
        spill_literal(change)?
    } else {
        edit_literal(change, &value)?
    };
    if change["$spillRelease"] == true || change.get("$spillMetadata").is_some() {
        attributes.remove("cm");
        attributes.remove("vm");
    }
    if let Some(metadata) = change.get("$spillMetadata") {
        attributes.insert(
            "cm".into(),
            metadata["cm"]
                .as_u64()
                .ok_or_else(|| {
                    crate::coded("spill-malformed", "Invalid spill cell metadata index")
                })?
                .to_string(),
        );
        if let Some(vm) = metadata.get("vm") {
            attributes.insert(
                "vm".into(),
                vm.as_u64()
                    .ok_or_else(|| {
                        crate::coded("spill-malformed", "Invalid spill value metadata index")
                    })?
                    .to_string(),
            );
        }
    }
    if replace || cache {
        attributes.remove("t");
        if !kind.is_empty() {
            attributes.insert("t".into(), kind.into());
        }
    }
    if let Some(style) = change["styleId"].as_u64() {
        attributes.insert("s".into(), style.to_string());
    }
    let mut output = format!("<{name}");
    for (k, v) in attributes {
        output.push_str(&format!(" {k}=\"{}\"", escape(&v)));
    }
    output.push('>');
    if replace {
        output.push_str(&content);
    }
    let spans: Vec<(u64, u64)> = serde_json::from_str(&spans).map_err(err)?;
    let mut cursor = open;
    let mut cached = false;
    for (a, b) in spans.into_iter().filter(|_| replace || cache) {
        if a < cursor || b > close {
            return Err(crate::coded(
                "workbook-malformed",
                "Invalid indexed value span",
            ));
        }
        let mut child = Reader::from_reader(&bytes[(a - start) as usize..(b - start) as usize]);
        let is_formula = matches!(child.read_event().map_err(err)?,Event::Start(e)|Event::Empty(e) if e.local_name().as_ref()==b"f");
        output.push_str(
            std::str::from_utf8(&bytes[(cursor - start) as usize..(a - start) as usize])
                .map_err(err)?,
        );
        if cache && is_formula {
            output.push_str(
                std::str::from_utf8(&bytes[(a - start) as usize..(b - start) as usize])
                    .map_err(err)?,
            );
            output.push_str(&content);
            cached = true;
        }
        cursor = b;
    }
    if cache && !cached {
        output.push_str(&content);
    }
    output.push_str(
        std::str::from_utf8(&bytes[(cursor - start) as usize..(close - start) as usize])
            .map_err(err)?,
    );
    output.push_str(&format!("</{name}>"));
    if change["$removeBlank"] == true {
        let safe_attributes = attrs(&element)?.keys().all(|key| {
            ["r", "s", "t", "xmlns"].contains(&key.as_str()) || key.starts_with("xmlns:")
        });
        let open_end = output
            .find('>')
            .ok_or_else(|| crate::coded("package-malformed", "Invalid cleared cell XML"))?
            + 1;
        let close_start = output
            .rfind("</")
            .ok_or_else(|| crate::coded("package-malformed", "Invalid cleared cell XML"))?;
        if !safe_attributes || !output[open_end..close_start].trim().is_empty() {
            return Err(crate::coded(
                "workbook-unsupported",
                "Cleared blank has opaque source metadata; normalized extent export is unavailable",
            ));
        }
        output.clear();
    }
    Ok(Some(Patch {
        start,
        end,
        text: output.into_bytes(),
    }))
}
fn new_cell(row: u32, col: u32, change: &Value) -> Result<String> {
    let value = if change.get("formula").is_some() {
        json!({"$formula":change["formula"],"calculation":change["calculation"]})
    } else {
        change["value"].clone()
    };
    let (kind, content) = if change.get("$spillMetadata").is_some() {
        spill_literal(change)?
    } else {
        edit_literal(change, &value)?
    };
    let metadata = change
        .get("$spillMetadata")
        .map(|metadata| -> Result<String> {
            let mut attributes = format!(
                " cm=\"{}\"",
                metadata["cm"].as_u64().ok_or_else(|| crate::coded(
                    "spill-malformed",
                    "Invalid new spill cell metadata index"
                ))?
            );
            if let Some(vm) = metadata.get("vm") {
                attributes.push_str(&format!(
                    " vm=\"{}\"",
                    vm.as_u64().ok_or_else(|| crate::coded(
                        "spill-malformed",
                        "Invalid new spill value metadata index"
                    ))?
                ));
            }
            Ok(attributes)
        })
        .transpose()?
        .unwrap_or_default();
    Ok(format!(
        "<c xmlns=\"{}\" r=\"{}\"{}{}{metadata}>{content}</c>",
        String::from_utf8_lossy(crate::xml::NS),
        address(row, col),
        if kind.is_empty() {
            String::new()
        } else {
            format!(" t=\"{kind}\"")
        },
        change["styleId"]
            .as_u64()
            .map(|id| format!(" s=\"{id}\""))
            .unwrap_or_default()
    ))
}
fn checked_spill_metadata(
    root: &Path,
    db: &Connection,
    version: u64,
    metadata: &Value,
) -> Result<std::collections::BTreeMap<String, String>> {
    let Some(plan) = metadata.get("spillMetadata") else {
        if crate::spill::any_history(db, version)? {
            return Err(crate::coded(
                "internal-missing-field",
                "Missing accepted spill metadata plan",
            ));
        }
        return Ok(Default::default());
    };
    if !crate::spill::any_history(db, version)? {
        return Err(crate::coded(
            "spill-bad-input",
            "Spill metadata has no accepted ownership history",
        ));
    }
    let raw: String = db
        .query_row(
            "SELECT spill_view FROM export_state WHERE version=?1",
            [version],
            |row| row.get(0),
        )
        .map_err(err)?;
    let view: Value = serde_json::from_str(&raw).map_err(err)?;
    crate::spill::validate_current(root, db, &json!({"version":version,"view":view}))?;
    let groups = crate::spill::export_groups(db, version, &view)?;
    let masters = plan["masters"].as_object().ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing accepted spill master metadata",
        )
    })?;
    let mut expected = std::collections::BTreeSet::new();
    for group in &groups {
        let key = format!("{}\0{}", field(group, "sheet")?, field(group, "master")?);
        expected.insert(key.clone());
        let map = masters.get(&key).ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing accepted spill master metadata",
            )
        })?;
        if map["range"] != group["range"]
            || group["status"] == "error"
                && (map.get("vm").is_none() || map["intendedRange"] != group["intendedRange"])
            || group["status"] == "calculated" && map.get("vm").is_some()
        {
            return Err(crate::coded(
                "internal-mismatch",
                "Spill metadata differs from accepted geometry/status",
            ));
        }
    }
    if masters
        .keys()
        .cloned()
        .collect::<std::collections::BTreeSet<_>>()
        != expected
    {
        return Err(crate::coded(
            "spill-refused",
            "Spill metadata contains an unowned master",
        ));
    }
    let replacements = plan["replacements"]
        .as_object()
        .ok_or_else(|| {
            crate::coded(
                "internal-missing-field",
                "Missing spill metadata part replacements",
            )
        })?
        .iter()
        .map(|(part, xml)| {
            Ok((
                part.clone(),
                xml.as_str()
                    .ok_or_else(|| crate::coded("spill-malformed", "Invalid spill part XML"))?
                    .to_owned(),
            ))
        })
        .collect::<Result<std::collections::BTreeMap<_, _>>>()?;
    let before = crate::spill::metadata::calculation_base(
        &crate::spill::metadata::source(root)?,
        &metadata["replacements"],
        &metadata["removed"],
    )?;
    crate::spill::metadata::validate(&before, &replacements, &plan["masters"])?;
    for group in groups {
        let sheet = field(&group, "sheet")?;
        let reference = field(&group, "master")?;
        let at = position(reference)?;
        let key = format!("{sheet}\0{reference}");
        let raw: String = db
            .query_row(
                "SELECT data FROM export_changes WHERE sheet=?1 AND r=?2 AND c=?3",
                params![sheet, at.0, at.1],
                |row| row.get(0),
            )
            .map_err(err)?;
        let mut change: Value = serde_json::from_str(&raw).map_err(err)?;
        if change["spill"]["role"] != "master" || change["spill"]["range"] != group["range"] {
            return Err(crate::coded(
                "spill-bad-input",
                "Export spill master has no native prepared cell proof",
            ));
        }
        change["$spillMetadata"] = masters[&key].clone();
        db.execute(
            "UPDATE export_changes SET data=?1 WHERE sheet=?2 AND r=?3 AND c=?4",
            params![change.to_string(), sheet, at.0, at.1],
        )
        .map_err(err)?;
    }
    Ok(replacements)
}
fn add_patch(db: &Connection, patch: Patch) -> Result<()> {
    db.prepare_cached("INSERT INTO export_patches(start,end,text) VALUES(?1,?2,?3)")
        .map_err(err)?
        .execute(params![patch.start, patch.end, patch.text])
        .map_err(err)?;
    Ok(())
}
fn insertion(db: &Connection, at: u64, text: String) -> Result<()> {
    add_patch(
        db,
        Patch {
            start: at,
            end: at,
            text: text.into_bytes(),
        },
    )
}
fn stream_patches(db: &Connection, source: &mut File, destination: &mut impl Write) -> Result<()> {
    let mut statement = db
        .prepare_cached("SELECT start,end,text FROM export_patches ORDER BY start,end,id")
        .map_err(err)?;
    let mut rows = statement.query([]).map_err(err)?;
    let mut cursor = 0;
    while let Some(row) = rows.next().map_err(err)? {
        crate::cancel::check()?;
        let start: u64 = row.get(0).map_err(err)?;
        let end: u64 = row.get(1).map_err(err)?;
        let text: Vec<u8> = row.get(2).map_err(err)?;
        if start < cursor || end < start {
            return Err(crate::coded(
                "workbook-conflict",
                "Overlapping worksheet edits",
            ));
        }
        source.seek(SeekFrom::Start(cursor)).map_err(err)?;
        io::copy(&mut source.take(start - cursor), destination).map_err(err)?;
        destination.write_all(&text).map_err(err)?;
        cursor = end;
    }
    source.seek(SeekFrom::Start(cursor)).map_err(err)?;
    io::copy(source, destination).map_err(err)?;
    Ok(())
}
fn opening(source: &mut File, start: u64, end: u64) -> Result<(String, String)> {
    let bytes = fragment(source, start, end)?;
    let mut reader = Reader::from_reader(bytes.as_slice());
    let element = match reader.read_event().map_err(err)? {
        Event::Empty(e) | Event::Start(e) => e,
        _ => return Err(crate::coded("workbook-malformed", "Invalid empty element")),
    };
    let tag = String::from_utf8_lossy(element.name().as_ref()).to_string();
    let mut text = format!("<{tag}");
    for (k, v) in attrs(&element)? {
        text.push_str(&format!(" {k}=\"{}\"", escape(&v)));
    }
    text.push('>');
    Ok((text, format!("</{tag}>")))
}
pub fn write(root: &Path, input: &Value) -> Result<Value> {
    if input
        .get("publication")
        .is_some_and(|value| !value.is_boolean())
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid publication export flag",
        ));
    }
    if input["publication"] == true
        && (input["calculationMetadata"] != true || input["version"].as_u64().is_none())
    {
        return Err(crate::coded(
            "formula-bad-input",
            "Publication export requires selected-version native calculation metadata",
        ));
    }
    let index: Value =
        serde_json::from_slice(&fs::read(local(root, "index.json")?).map_err(err)?).map_err(err)?;
    if field(input, "revision")? != field(&index, "revision")? {
        return Err(crate::coded("workbook-stale", "Workbook revision is stale"));
    }
    let metadata = if input["calculationMetadata"].as_bool().unwrap_or(false) {
        serde_json::from_slice::<Value>(
            &fs::read(local(root, "export-metadata.json")?).map_err(err)?,
        )
        .map_err(err)?
    } else {
        // A literal-only export with no style edits needs no style resolver.
        // Explicitly empty mappings still refuse unresolved authored styles.
        json!({"styleIds":[]})
    };
    if index["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("workbook-malformed", "Invalid index"))?
        .iter()
        .any(|s| s["formulaCount"].as_u64().unwrap_or(0) > 0)
        && !input["calculationMetadata"].as_bool().unwrap_or(false)
    {
        return Err(crate::coded(
            "formula-bad-input",
            "Formula export requires the calculation owner",
        ));
    }
    let db = database(root)?;
    let preserving =
        crate::calculation_policy::preserving(&db, input["version"].as_u64().unwrap_or(0))?;
    if input
        .get("preservationExport")
        .is_some_and(|v| !v.is_boolean())
        || preserving != (input["preservationExport"] == true)
    {
        return Err(crate::coded(
            "workbook-unsupported",
            "Preservation export does not match native accepted policy",
        ));
    }
    let rewrite_candidate = crate::used_range_transfer::prepare_export(root, &db, input)?;
    if crate::spill::present(&db)?
        && crate::spill::any_history(
            &db,
            input["version"].as_u64().ok_or_else(|| {
                crate::coded(
                    "spill-bad-input",
                    "Spill export requires an accepted version",
                )
            })?,
        )?
        && metadata.get("spillMetadata").is_none()
    {
        return Err(crate::coded(
            "spill-unsupported",
            "Spill ownership export remains gated without complete native metadata admission",
        ));
    }
    if input["calculationMetadata"] == true {
        crate::package_edit::validate_calculation_metadata(
            root,
            &db,
            input["version"].as_u64().ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing policy export version")
            })?,
            &metadata,
        )?;
    }
    let prepared = input.get("version").is_some();
    if prepared {
        let version = input["version"]
            .as_u64()
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid export version"))?;
        if db
            .query_row("SELECT version FROM export_state", [], |r| {
                r.get::<_, u64>(0)
            })
            .map_err(err)?
            != version
        {
            return Err(crate::coded("workbook-stale", "Stale export preparation"));
        }
        db.execute_batch(
            "CREATE TEMP TABLE resolved_styles(id INTEGER PRIMARY KEY,style INTEGER NOT NULL)",
        )
        .map_err(err)?;
        for item in metadata["styleIds"]
            .as_array()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing export styles"))?
        {
            db.execute(
                "INSERT INTO resolved_styles VALUES(?1,?2)",
                params![
                    item["id"].as_u64().ok_or_else(|| crate::coded(
                        "internal-invalid-field",
                        "Invalid export style"
                    ))?,
                    item["styleId"].as_u64().ok_or_else(|| crate::coded(
                        "internal-invalid-field",
                        "Invalid resolved style"
                    ))?
                ],
            )
            .map_err(err)?;
        }
        let missing:u64=db.prepare_cached("SELECT count(*) FROM export_styles s LEFT JOIN resolved_styles r ON s.id=r.id WHERE r.id IS NULL").map_err(crate::err)?.query_row([],|r|r.get(0)).map_err(err)?;
        if missing != 0 {
            return Err(crate::coded(
                "style-unsupported",
                "Unresolved export styles",
            ));
        }
    } else {
        db.execute_batch("CREATE TEMP TABLE export_changes(sheet TEXT,r INTEGER,c INTEGER,data TEXT,style INTEGER,PRIMARY KEY(sheet,r,c)) WITHOUT ROWID").map_err(err)?;
        for change in input["changes"]
            .as_array()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing changes"))?
        {
            let sheet = field(change, "sheet")?;
            let (r, c) = position(field(change, "cell")?)?;
            let previous: Option<String> = db
                .query_row(
                    "SELECT data FROM export_changes WHERE sheet=?1 AND r=?2 AND c=?3",
                    params![sheet, r, c],
                    |r| r.get(0),
                )
                .optional()
                .map_err(err)?;
            let mut entry = previous
                .map(|s| serde_json::from_str::<Value>(&s).map_err(err))
                .transpose()?
                .unwrap_or(json!({}));
            if change.get("value").is_some() {
                entry.as_object_mut().unwrap().remove("formula");
                entry.as_object_mut().unwrap().remove("calculation");
                entry.as_object_mut().unwrap().remove("valueType");
            }
            if change.get("formula").is_some() {
                entry.as_object_mut().unwrap().remove("value");
                entry.as_object_mut().unwrap().remove("valueType");
            }
            for (key, value) in change
                .as_object()
                .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid change"))?
            {
                entry[key] = value.clone();
            }
            db.execute(
                "INSERT OR REPLACE INTO export_changes VALUES(?1,?2,?3,?4,NULL)",
                params![sheet, r, c, serde_json::to_string(&entry).map_err(err)?],
            )
            .map_err(err)?;
        }
    }
    let mut sheets = std::collections::BTreeSet::new();
    {
        let mut statement = db
            .prepare_cached("SELECT DISTINCT sheet FROM export_changes")
            .map_err(err)?;
        for sheet in statement
            .query_map([], |r| r.get::<_, String>(0))
            .map_err(err)?
        {
            sheets.insert(sheet.map_err(err)?);
        }
    }
    if prepared {
        for sheet in index["sheets"]
            .as_array()
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid source sheets"))?
        {
            sheets.insert(field(sheet, "path")?.to_owned());
        }
    }
    db.execute_batch("CREATE TEMP TABLE export_removed_cells(r INTEGER,c INTEGER,PRIMARY KEY(r,c)) WITHOUT ROWID; CREATE TEMP TABLE export_patches(id INTEGER PRIMARY KEY,start INTEGER,end INTEGER,text BLOB); CREATE INDEX patch_order ON export_patches(start,end,id); CREATE TEMP TABLE export_new_cells(r INTEGER,c INTEGER,text TEXT,PRIMARY KEY(r,c)) WITHOUT ROWID;").map_err(err)?;
    let mut zip =
        ZipArchive::new(File::open(local(root, "source.xlsx")?).map_err(err)?).map_err(err)?;
    let chart_caches =
        checked_chart_caches(root, &db, input["version"].as_u64().unwrap_or(0), &metadata)?;
    let spill_parts =
        checked_spill_metadata(root, &db, input["version"].as_u64().unwrap_or(0), &metadata)?;
    crate::spill::consumers::validate_export(root, &db, input["version"].as_u64().unwrap_or(0))?;
    crate::spill::scalars::validate_export(root, &db, input["version"].as_u64().unwrap_or(0))
        .map_err(|error| {
            crate::rephrase("spill-refused", error, |said| {
                format!("Scalar-array export admission: {said}")
            })
        })?;
    let source_parts = (0..zip.len())
        .map(|index| {
            zip.name_for_index(index)
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing source part name"))
                .map(str::to_owned)
        })
        .collect::<std::result::Result<std::collections::BTreeSet<_>, _>>()
        .map_err(err)?;
    let output = local(root, "export.xlsx")?;
    let temporary = local(root, "export.partial")?;
    let mut writer = ZipWriter::new(crate::quota::create(&temporary).map_err(err)?);
    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    let output_styles = metadata["replacements"]["xl/styles.xml"].as_str();
    for i in 0..zip.len() {
        let member = zip.by_index(i).map_err(err)?;
        let name = member.name().to_owned();
        if metadata["removed"]
            .as_array()
            .is_some_and(|names| names.iter().any(|item| item.as_str() == Some(&name)))
        {
            continue;
        }
        if let Some(replacement) = chart_caches.get(&name) {
            writer.start_file(&name, options).map_err(err)?;
            writer.write_all(replacement.as_bytes()).map_err(err)?;
            continue;
        }
        if let Some(replacement) = spill_parts.get(&name) {
            writer.start_file(&name, options).map_err(err)?;
            writer.write_all(replacement.as_bytes()).map_err(err)?;
            continue;
        }
        if let Some(replacement) = metadata["replacements"].get(&name).and_then(Value::as_str) {
            if ![
                "xl/workbook.xml",
                "xl/_rels/workbook.xml.rels",
                "[Content_Types].xml",
                "xl/styles.xml",
            ]
            .contains(&name.as_str())
            {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid calculation metadata member",
                ));
            }
            writer.start_file(&name, options).map_err(err)?;
            writer.write_all(replacement.as_bytes()).map_err(err)?;
            continue;
        }
        if !sheets.remove(&name) {
            writer.raw_copy_file(member).map_err(err)?;
            continue;
        }
        let metadata = index["sheets"]
            .as_array()
            .unwrap()
            .iter()
            .find(|s| s["path"] == name)
            .ok_or_else(|| crate::coded("workbook-bad-input", "Sheet is not indexed"))?;
        let mut source = File::open(local(root, field(metadata, "artifact")?)?).map_err(err)?;
        let normalization = if let Some(version) = input["version"].as_u64()
            && index["sourceFactSchema"] == crate::used_range::FACTS_SCHEMA
        {
            crate::used_range::saved_export_normalization(
                root,
                &db,
                version,
                metadata,
                output_styles,
            )?
        } else {
            None
        };
        let normalized_bounds = normalization.as_ref().map(|plan| plan.bounds);
        db.execute_batch("DELETE FROM export_patches; DELETE FROM export_new_cells; DELETE FROM export_removed_cells; BEGIN;")
            .map_err(err)?;
        // A missing dimension can share the self-closing sheetData offset.
        // Queue it before data expansion so equal-offset insertions stay at
        // worksheet level, ahead of the new sheetData opening and rows.
        let dimension = &metadata["dimension"];
        if let Some(bounds) = normalized_bounds {
            if let Some(attributes) = dimension["attributes"].as_object() {
                let tag = field(dimension, "name")?;
                let mut text = format!("<{tag}");
                for (k, v) in attributes {
                    let value = if k == "ref" {
                        bounds.range()
                    } else {
                        v.as_str()
                            .ok_or_else(|| {
                                crate::coded("workbook-malformed", "Invalid dimension attribute")
                            })?
                            .to_owned()
                    };
                    text.push_str(&format!(" {k}=\"{}\"", escape(&value)));
                }
                text.push_str("/>");
                add_patch(
                    &db,
                    Patch {
                        start: dimension["start"].as_u64().ok_or_else(|| {
                            crate::coded("workbook-malformed", "Invalid dimension span")
                        })?,
                        end: dimension["end"].as_u64().ok_or_else(|| {
                            crate::coded("workbook-malformed", "Invalid dimension span")
                        })?,
                        text: text.into_bytes(),
                    },
                )?;
            } else {
                let at = metadata["dimensionInsert"].as_u64().ok_or_else(|| {
                    crate::coded(
                        "internal-missing-field",
                        "Missing qualified worksheet dimension insertion point",
                    )
                })?;
                insertion(
                    &db,
                    at,
                    format!(
                        "<dimension xmlns=\"{}\" ref=\"{}\"/>",
                        String::from_utf8_lossy(crate::xml::NS),
                        escape(&bounds.range())
                    ),
                )?;
            }
        }
        if let Some(plan) = &normalization {
            plan.removable_cells(&db, input["version"].as_u64().unwrap(), |row, col| {
                if let Some(patch) = patch_cell(
                    &db,
                    &mut source,
                    &name,
                    row,
                    col,
                    &json!({"value":null,"styleId":0,"$removeBlank":true}),
                )? {
                    add_patch(&db, patch)?;
                }
                db.execute(
                    "INSERT INTO export_removed_cells VALUES(?1,?2)",
                    params![row, col],
                )
                .map_err(err)?;
                Ok(())
            })?;
        }
        let (mut max_row, mut max_col) = (0u32, 0u32);
        let (mut min_row, mut min_col) = (u32::MAX, u32::MAX);
        {
            let sql = if prepared {
                "SELECT e.r,e.c,e.data,s.style FROM export_changes e LEFT JOIN resolved_styles s ON s.id=e.style WHERE sheet=?1 ORDER BY r,c"
            } else {
                "SELECT r,c,data,NULL FROM export_changes WHERE sheet=?1 ORDER BY r,c"
            };
            let mut statement = db.prepare_cached(sql).map_err(err)?;
            let mut rows = statement.query([&name]).map_err(err)?;
            while let Some(record) = rows.next().map_err(err)? {
                let row: u32 = record.get(0).map_err(err)?;
                let col: u32 = record.get(1).map_err(err)?;
                let mut value: Value =
                    serde_json::from_str(&record.get::<_, String>(2).map_err(err)?).map_err(err)?;
                if let Some(style) = record.get::<_, Option<u64>>(3).map_err(err)? {
                    value["styleId"] = json!(style);
                }
                if value.as_object().is_some_and(|v| v.is_empty()) {
                    continue;
                }
                let removed: bool = db
                    .query_row(
                        "SELECT EXISTS(SELECT 1 FROM export_removed_cells WHERE r=?1 AND c=?2)",
                        params![row, col],
                        |r| r.get(0),
                    )
                    .map_err(err)?;
                if removed {
                    continue;
                }
                min_row = min_row.min(row);
                min_col = min_col.min(col);
                max_row = max_row.max(row);
                max_col = max_col.max(col);
                if let Some(patch) = patch_cell(&db, &mut source, &name, row, col, &value)? {
                    add_patch(&db, patch)?;
                    continue;
                }
                let existing: Option<(u64, u64)> = db
                    .prepare_cached("SELECT open,close FROM rows WHERE sheet=?1 AND r=?2")
                    .map_err(err)?
                    .query_row(params![name, row], |r| Ok((r.get(0)?, r.get(1)?)))
                    .optional()
                    .map_err(err)?;
                if let Some((open, close)) = existing.filter(|(open, close)| open != close) {
                    let next:Option<u64>=db.prepare_cached("SELECT start FROM cells WHERE sheet=?1 AND r=?2 AND c>?3 ORDER BY c LIMIT 1").map_err(err)?.query_row(params![name,row,col],|r|r.get(0)).optional().map_err(err)?;
                    let after: Option<u64> = db
                        .prepare_cached("SELECT max(end) FROM cells WHERE sheet=?1 AND r=?2")
                        .map_err(err)?
                        .query_row(params![name, row], |r| r.get(0))
                        .map_err(err)?;
                    let _ = close;
                    insertion(
                        &db,
                        next.or(after).unwrap_or(open),
                        new_cell(row, col, &value)?,
                    )?;
                } else {
                    db.prepare_cached("INSERT INTO export_new_cells VALUES(?1,?2,?3)")
                        .map_err(err)?
                        .execute(params![row, col, new_cell(row, col, &value)?])
                        .map_err(err)?;
                }
            }
        }
        let fresh_count: u64 = db
            .query_row("SELECT count(*) FROM export_new_cells", [], |r| r.get(0))
            .map_err(err)?;
        if fresh_count > 0 {
            let start = metadata["dataStart"].as_u64().ok_or_else(|| {
                crate::coded(
                    "needs-structural-writer",
                    "Missing sheetData requires the shared structural writer",
                )
            })?;
            let end = metadata["dataEnd"].as_u64().ok_or_else(|| {
                crate::coded(
                    "needs-structural-writer",
                    "Missing sheetData requires the shared structural writer",
                )
            })?;
            let open = metadata["dataOpen"].as_u64().ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing sheetData opening")
            })?;
            let close = metadata["dataClose"].as_u64().ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing sheetData closing")
            })?;
            let sheet_close = if open == close {
                let (opening, closing) = opening(&mut source, start, end)?;
                insertion(&db, start, opening)?;
                Some(closing)
            } else {
                None
            };
            let mut statement = db
                .prepare_cached("SELECT DISTINCT r FROM export_new_cells ORDER BY r")
                .map_err(err)?;
            let mut rows = statement.query([]).map_err(err)?;
            while let Some(record) = rows.next().map_err(err)? {
                let row: u32 = record.get(0).map_err(err)?;
                let existing: Option<(u64, u64)> = db
                    .prepare_cached("SELECT start,end FROM rows WHERE sheet=?1 AND r=?2")
                    .map_err(err)?
                    .query_row(params![name, row], |r| Ok((r.get(0)?, r.get(1)?)))
                    .optional()
                    .map_err(err)?;
                let (at, until, opening, closing) = if let Some((a, b)) = existing {
                    let (o, c) = opening(&mut source, a, b)?;
                    (a, b, o, c)
                } else {
                    let next: Option<u64> = db
                        .prepare_cached(
                            "SELECT start FROM rows WHERE sheet=?1 AND r>?2 ORDER BY r LIMIT 1",
                        )
                        .map_err(err)?
                        .query_row(params![name, row], |r| r.get(0))
                        .optional()
                        .map_err(err)?;
                    let at = if sheet_close.is_some() {
                        start
                    } else {
                        next.unwrap_or(close)
                    };
                    (
                        at,
                        at,
                        format!(
                            "<row xmlns=\"{}\" r=\"{}\">",
                            String::from_utf8_lossy(crate::xml::NS),
                            row + 1
                        ),
                        "</row>".into(),
                    )
                };
                insertion(&db, at, opening)?;
                let mut cells = db
                    .prepare_cached("SELECT text FROM export_new_cells WHERE r=?1 ORDER BY c")
                    .map_err(err)?;
                let mut cells = cells.query([row]).map_err(err)?;
                while let Some(cell) = cells.next().map_err(err)? {
                    insertion(&db, at, cell.get(0).map_err(err)?)?;
                }
                add_patch(
                    &db,
                    Patch {
                        start: at,
                        end: until,
                        text: closing.into_bytes(),
                    },
                )?;
            }
            if let Some(closing) = sheet_close {
                add_patch(
                    &db,
                    Patch {
                        start,
                        end,
                        text: closing.into_bytes(),
                    },
                )?;
            }
        }
        if let Some(plan) = &normalization {
            let mut statement = db
                .prepare_cached("SELECT r,start,end FROM rows WHERE sheet=?1 ORDER BY r")
                .map_err(err)?;
            let mut rows = statement.query([&name]).map_err(err)?;
            while let Some(record) = rows.next().map_err(err)? {
                crate::cancel::check()?;
                let row: u32 = record.get(0).map_err(err)?;
                if !plan.removable_row(&db, row)? {
                    continue;
                }
                let survives:bool=db.prepare_cached("SELECT EXISTS(SELECT 1 FROM cells c WHERE c.sheet=?1 AND c.r=?2 AND NOT EXISTS(SELECT 1 FROM export_removed_cells x WHERE x.r=c.r AND x.c=c.c)) OR EXISTS(SELECT 1 FROM export_changes c WHERE c.sheet=?1 AND c.r=?2 AND NOT EXISTS(SELECT 1 FROM export_removed_cells x WHERE x.r=c.r AND x.c=c.c)) OR EXISTS(SELECT 1 FROM export_new_cells WHERE r=?2)").map_err(crate::err)?.query_row(params![name,row],|r|r.get(0)).map_err(err)?;
                if !survives {
                    let start: u64 = record.get(1).map_err(err)?;
                    let end: u64 = record.get(2).map_err(err)?;
                    db.execute(
                        "DELETE FROM export_patches WHERE start>=?1 AND end<=?2 AND start<end",
                        params![start, end],
                    )
                    .map_err(err)?;
                    add_patch(
                        &db,
                        Patch {
                            start,
                            end,
                            text: Vec::new(),
                        },
                    )?;
                }
            }
        }
        if normalized_bounds.is_none()
            && min_row != u32::MAX
            && let Some(attributes) = dimension["attributes"].as_object()
            && let Some(reference) = attributes.get("ref").and_then(Value::as_str)
        {
            let mut bounds = reference.split(':');
            let first = position(
                bounds
                    .next()
                    .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid dimension"))?,
            )?;
            let last = bounds.next().map(position).transpose()?.unwrap_or(first);
            if min_row != u32::MAX
                && (max_row > last.0 || max_col > last.1 || min_row < first.0 || min_col < first.1)
            {
                let tag = field(dimension, "name")?;
                let mut text = format!("<{tag}");
                for (k, v) in attributes {
                    let value = if k == "ref" {
                        format!(
                            "{}:{}",
                            address(first.0.min(min_row), first.1.min(min_col)),
                            address(max_row.max(last.0), max_col.max(last.1))
                        )
                    } else {
                        v.as_str()
                            .ok_or_else(|| {
                                crate::coded("workbook-malformed", "Invalid dimension attribute")
                            })?
                            .to_owned()
                    };
                    text.push_str(&format!(" {k}=\"{}\"", escape(&value)));
                }
                text.push_str("/>");
                add_patch(
                    &db,
                    Patch {
                        start: dimension["start"].as_u64().unwrap(),
                        end: dimension["end"].as_u64().unwrap(),
                        text: text.into_bytes(),
                    },
                )?;
            }
        }
        db.execute_batch("COMMIT;").map_err(err)?;
        writer.start_file(&name, options).map_err(err)?;
        stream_patches(&db, &mut source, &mut writer)?;
    }
    if !sheets.is_empty() {
        return Err(crate::coded(
            "internal-unknown-reference",
            "Unknown sheet in export",
        ));
    }
    for (part, xml) in spill_parts {
        if !source_parts.contains(&part) {
            writer.start_file(&part, options).map_err(err)?;
            writer.write_all(xml.as_bytes()).map_err(err)?;
        }
    }
    writer.finish().map_err(err)?.sync_all().map_err(err)?;
    fs::rename(&temporary, &output).map_err(err)?;
    let revision = hash(&output)?;
    if input["publication"] == true {
        if input["calculationMetadata"] != true || !prepared {
            return Err(crate::coded(
                "formula-bad-input",
                "Publication export requires selected-version native calculation metadata",
            ));
        }
        crate::calculation_transfer::issue_publication_receipt(
            root,
            &db,
            input["version"].as_u64().ok_or_else(|| {
                crate::coded(
                    "internal-missing-field",
                    "Missing publication export version",
                )
            })?,
            &revision,
        )?;
    }
    let mut reply = json!({"revision":revision,"file":"export.xlsx"});
    if let Some(candidate) = rewrite_candidate {
        reply["rewriteInputId"] = json!(crate::used_range_transfer::finish_export(
            root,
            &db,
            candidate,
            "export.xlsx",
            &revision
        )?);
    }
    Ok(reply)
}

#[cfg(test)]
mod scalar_publication_tests {
    use super::*;
    const SHEET: &str = "xl/worksheets/main.xml";

    fn fixture(scalar_rows: u32, dependents: bool) -> tempfile::TempDir {
        let root = tempfile::tempdir().unwrap();
        let mut zip = ZipWriter::new(File::create(root.path().join("source.xlsx")).unwrap());
        let mut rows = String::new();
        for row in 1..=scalar_rows {
            let literals = if row <= 5 {
                let key = [1, 2, 2, 4, 5][(row - 1) as usize];
                let value = if row == 5 {
                    "<c r=\"B5\" t=\"b\"><v>0</v></c>".into()
                } else {
                    format!(
                        "<c r=\"B{row}\" t=\"inlineStr\"><is><t xml:space=\"preserve\">{}</t></is></c>",
                        ["One", "First two", "Last two", "#N/A"][(row - 1) as usize]
                    )
                };
                format!("<c r=\"A{row}\"><v>{key}</v></c>{value}")
            } else {
                String::new()
            };
            let formula = if row == 1 || row > 5 {
                "SUM(_xlfn.SEQUENCE(J1,1))".into()
            } else {
                format!(
                    "_xlfn.XLOOKUP({},A1:A5,B1:B5)",
                    [0, 2, 5, 9, 4][(row - 1) as usize]
                )
            };
            let dependent = if dependents && row <= 4 {
                format!(
                    "<c r=\"E{row}\"><f>{}</f></c>",
                    ["D1+1", "LEN(D2)", "IFERROR(D4,99)", "IF(D3,1,0)"][(row - 1) as usize]
                )
            } else {
                String::new()
            };
            rows.push_str(&format!("<row r=\"{row}\">{literals}<c r=\"D{row}\"><f>{formula}</f></c>{dependent}{}</row>", if row == 1 { "<c r=\"J1\"><v>3</v></c>" } else { "" }));
        }
        let sheet = format!(
            "<worksheet xmlns=\"{}\"><sheetData xmlns=\"{}\">{rows}</sheetData></worksheet>",
            String::from_utf8_lossy(crate::xml::NS),
            String::from_utf8_lossy(crate::xml::NS)
        );
        for (name, text) in [
            (
                "[Content_Types].xml",
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Override PartName=\"/xl/worksheets/main.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>",
            ),
            (
                "xl/workbook.xml",
                "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Main\" sheetId=\"1\" r:id=\"main\"/></sheets><calcPr calcMode=\"auto\" calcOnSave=\"0\" fullCalcOnLoad=\"0\" forceFullCalc=\"0\"/></workbook>",
            ),
            (
                "xl/_rels/workbook.xml.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"main\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/main.xml\"/></Relationships>",
            ),
            (SHEET, sheet.as_str()),
            (
                "xl/styles.xml",
                "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts count=\"1\"><font/></fonts><fills count=\"1\"><fill/></fills><borders count=\"1\"><border/></borders><cellXfs count=\"1\"><xf numFmtId=\"0\"/></cellXfs></styleSheet>",
            ),
        ] {
            zip.start_file(name, SimpleFileOptions::default()).unwrap();
            zip.write_all(text.as_bytes()).unwrap();
        }
        zip.finish().unwrap();
        crate::import::index(root.path()).unwrap();
        root
    }

    fn prepare(root: &Path, scalar_rows: u32, dependents: bool) -> u64 {
        let view = json!({"sheets":[{"name":"Main","path":SHEET,"storedFormulaCount":scalar_rows+if dependents {4} else {0},"merges":[]}],"namedRanges":[],"tables":[]});
        crate::calculation_policy::command(root, &json!({"version":0,"view":view,"method":"initialize","initialPolicy":{"mode":"auto","recalculateBeforeSave":false}})).unwrap();
        let db = database(root).unwrap();
        let mut changes = Vec::new();
        for first in (1..=scalar_rows).step_by(64) {
            let items = (first..=(first + 63).min(scalar_rows))
                .map(|row| json!({"sheet":SHEET,"ref":format!("D{row}")}))
                .collect::<Vec<_>>();
            let proofs = crate::spill::scalars::command(
                root,
                &json!({"version":0,"view":view,"items":items,"revision":"scalar-proof"}),
            )
            .unwrap();
            assert_eq!(proofs["refused"], json!([]));
            assert_eq!(proofs["results"].as_array().unwrap().len(), items.len());
            for result in proofs["results"].as_array().unwrap() {
                let mut cell = crate::range::read_db(
                    &db,
                    &json!({"version":0,"sheet":SHEET,"range":result["ref"],"limit":1}),
                )
                .unwrap()["cells"][0][1]
                    .clone();
                cell["calculation"] = result["calculation"].clone();
                changes.push(json!({"sheet":SHEET,"cell":result["ref"],"value":cell,"edits":[]}));
            }
        }
        if dependents {
            // Ordinary dependent results are accepted fixture records; the five
            // typed scalar results above come from actual native proof generation.
            for (row, value) in [(1, 7), (2, 9), (3, 99), (4, 0)] {
                let reference = format!("E{row}");
                let mut cell = crate::range::read_db(
                    &db,
                    &json!({"version":0,"sheet":SHEET,"range":reference,"limit":1}),
                )
                .unwrap()["cells"][0][1]
                    .clone();
                cell["calculation"] =
                    json!({"status":"calculated","value":value,"revision":"fixture-dependent"});
                changes.push(json!({"sheet":SHEET,"cell":reference,"value":cell,"edits":[]}));
            }
        }
        drop(db);
        let version =
            crate::range::patch(root, &json!({"version":0,"changes":changes})).unwrap()["version"]
                .as_u64()
                .unwrap();
        crate::calculation_policy::command(
            root,
            &json!({"version":version,"view":view,"method":"settle"}),
        )
        .unwrap();
        crate::compact::compact(root, &json!({"keep":[0,version]})).unwrap();
        crate::export_state::prepare(root, &json!({"version":version,"view":view})).unwrap();
        let mut zip = ZipArchive::new(File::open(root.join("source.xlsx")).unwrap()).unwrap();
        let mut workbook = String::new();
        zip.by_name("xl/workbook.xml")
            .unwrap()
            .read_to_string(&mut workbook)
            .unwrap();
        fs::write(
            root.join("export-metadata.json"),
            json!({"styleIds":[],"replacements":{"xl/workbook.xml":workbook}}).to_string(),
        )
        .unwrap();
        version
    }

    fn publication(root: &Path, version: u64) -> Result<Value> {
        let index: Value =
            serde_json::from_slice(&fs::read(root.join("index.json")).unwrap()).unwrap();
        crate::command(
            json!({"root":root.to_str().unwrap(),"operation":"export","version":version,"revision":index["revision"],"calculationMetadata":true,"publication":true}),
        )
    }

    #[test]
    fn typed_scalar_proofs_with_dependents_export_and_issue_publication_receipt() {
        let _scope = crate::shared::begin();
        let fixture = fixture(5, true);
        let root = fixture.path();
        let version = prepare(root, 5, true);
        let output = publication(root, version).unwrap();
        assert_eq!(output["revision"], hash(&root.join("export.xlsx")).unwrap());
        let db = database(root).unwrap();
        let receipt: String = db
            .query_row(
                "SELECT export_revision FROM publication_export_receipts WHERE version=?1",
                [version],
                |row| row.get(0),
            )
            .unwrap();
        assert_eq!(output["revision"], receipt);
        let mut saved = ZipArchive::new(File::open(root.join("export.xlsx")).unwrap()).unwrap();
        let mut xml = String::new();
        saved
            .by_name(SHEET)
            .unwrap()
            .read_to_string(&mut xml)
            .unwrap();
        assert!(xml.contains("First two"));
        assert!(xml.contains("#N/A"));
        assert!(xml.contains(">6.0</v>") || xml.contains(">6</v>"));
    }

    #[test]
    fn scalar_publication_pages_validate_and_refuse_forged_last_page() {
        let _scope = crate::shared::begin();
        let fixture = fixture(70, false);
        let root = fixture.path();
        let version = prepare(root, 70, false);
        publication(root, version).unwrap();
        let db = database(root).unwrap();
        let raw: String = db
            .query_row(
                "SELECT data FROM export_changes WHERE sheet=?1 AND r=69 AND c=3",
                [SHEET],
                |row| row.get(0),
            )
            .unwrap();
        let mut prepared: Value = serde_json::from_str(&raw).unwrap();
        prepared["calculation"]["value"] = json!(999);
        db.execute(
            "UPDATE export_changes SET data=?1 WHERE sheet=?2 AND r=69 AND c=3",
            params![prepared.to_string(), SHEET],
        )
        .unwrap();
        assert!(
            publication(root, version)
                .unwrap_err()
                .contains("Prepared scalar array formula/cache differs from accepted proof")
        );
        db.execute(
            "UPDATE export_changes SET data=?1 WHERE sheet=?2 AND r=69 AND c=3",
            params![raw, SHEET],
        )
        .unwrap();
        let raw: String = db
            .query_row(
                "SELECT data FROM changes WHERE version=?1 AND sheet=?2 AND r=69 AND c=3",
                params![version, SHEET],
                |row| row.get(0),
            )
            .unwrap();
        let mut accepted: Value = serde_json::from_str(&raw).unwrap();
        accepted["calculation"]["scalarArray"]["inputSha256"] = json!("0".repeat(64));
        db.execute(
            "UPDATE changes SET data=?1 WHERE version=?2 AND sheet=?3 AND r=69 AND c=3",
            params![accepted.to_string(), version, SHEET],
        )
        .unwrap();
        assert!(
            publication(root, version)
                .unwrap_err()
                .contains("Scalar array cache/input proof is stale or forged")
        );
    }
}
