//! Closed semantic DiagramML authoring. Reconstruct each part independently of the JS plan.
use super::managed_shape::{attribute, child, quote, same, scalar};
use super::print_settings::Tree;
use super::*;
const G: &str = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const DSP: &str = "http://schemas.microsoft.com/office/drawing/2008/diagram";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const D: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const P: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
fn invalid_xml_text(text: &str) -> bool {
    text.chars().any(|c| {
        (c < '\u{20}' && !['\t', '\n', '\r'].contains(&c)) || ['\u{fffe}', '\u{ffff}'].contains(&c)
    })
}
fn guid(id: &str, slot: u32) -> String {
    let tail = u32::from_str_radix(&id[28..], 16).expect("validated UUID") ^ slot;
    format!("{{{}{:08X}}}", id[..28].to_ascii_uppercase(), tail)
}
pub(super) fn planned_parts(id:&str,part:&str,_create:bool)->Result<Vec<String>>{
 let parts=parts(id);if part!=parts[0]{return Err(crate::coded("object-bad-input","Unknown diagram part identity"));}Ok(parts)
}

fn parts(id: &str) -> Vec<String> {
    let mut names = ["data", "layout", "style", "colors", "drawing"]
        .map(|role| {
            format!(workbook_format!(XL_DIAGRAMS_SMARTART_ID_ROLE_XML),
                id = id,
                role = role
            )
        })
        .to_vec();
    names.push(format!(workbook_format!(XL_DIAGRAMS_RELS_SMARTART_ID_DATA_XML_RELS),
        id = id
    ));
    names
}
fn fill(c: &str) -> String {
    format!("<a:solidFill><a:srgbClr val=\"{c}\"/></a:solidFill>")
}
fn text_body(text: &str) -> String {
    format!(
        "<a:bodyPr wrap=\"square\" anchor=\"ctr\"/><a:lstStyle/><a:p><a:pPr algn=\"ctr\"/><a:r><a:rPr lang=\"en-US\" sz=\"1400\">{}<a:latin typeface=\"Arial\"/></a:rPr><a:t xml:space=\"preserve\">{}</a:t></a:r></a:p>",
        fill("FFFFFF"),
        quote(text)
    )
}
fn point(id: &str, slot: u32, kind: &str, props: &str, text: &str, extra: &str) -> String {
    format!(
        "<dgm:pt modelId=\"{}\" type=\"{kind}\"{extra}><dgm:prSet{props}/><dgm:spPr/>{}</dgm:pt>",
        guid(id, slot),
        if kind == "node" {
            format!("<dgm:t>{}</dgm:t>", text_body(text))
        } else {
            String::new()
        }
    )
}
#[derive(Clone)]
struct BoxNode {
    x: f64,
    y: f64,
    w: f64,
    h: f64,
}
fn geometry(kind: &str, count: usize) -> (Vec<BoxNode>, Vec<(usize, usize)>) {
    let n = count as f64;
    let boxes = (0..count)
        .map(|i| {
            let i = i as f64;
            match kind {
                "process" => {
                    let step = 920.0 / n;
                    let w = 260.0_f64.min(step - 30.0);
                    BoxNode {
                        x: 40.0 + i * step + (step - w) / 2.0,
                        y: 220.0,
                        w,
                        h: 160.0,
                    }
                }
                "hierarchy" => {
                    let step = 920.0 / (n - 1.0);
                    let w = 260.0_f64.min(step - 24.0);
                    BoxNode {
                        x: if i == 0.0 {
                            370.0
                        } else {
                            40.0 + (i - 1.0) * step + (step - w) / 2.0
                        },
                        y: if i == 0.0 { 50.0 } else { 380.0 },
                        w: if i == 0.0 { 260.0 } else { w },
                        h: 150.0,
                    }
                }
                _ => {
                    let angle = -std::f64::consts::PI / 2.0 + 2.0 * std::f64::consts::PI * i / n;
                    BoxNode {
                        x: 500.0 + angle.cos() * 320.0 - 90.0,
                        y: 300.0 + angle.sin() * 180.0 - 55.0,
                        w: 180.0,
                        h: 110.0,
                    }
                }
            }
        })
        .collect::<Vec<_>>();
    let mut edges = (1..count)
        .map(|i| (if kind == "hierarchy" { 0 } else { i - 1 }, i))
        .collect::<Vec<_>>();
    if kind == "cycle" {
        edges.push((count - 1, 0));
    }
    (boxes, edges)
}
fn expected(
    id: &str,
    kind: &str,
    texts: &[String],
    width: u64,
    height: u64,
) -> BTreeMap<String, String> {
    let names = parts(id);
    let uid = format!(workbook_format!(URN_SMARTART_LAYOUT_KIND), kind = kind);
    let n = texts.len();
    let mut points = point(id, 0, "doc", "", "", "");
    let mut connections = String::new();
    for (i, text) in texts.iter().enumerate() {
        let i = i as u32;
        points += &point(id, i + 1, "node", " phldrT=\"Node\"", text, "");
        for (slot, kind) in [(200 + i, "parTrans"), (300 + i, "sibTrans")] {
            points += &point(
                id,
                slot,
                kind,
                "",
                "",
                &format!(" cxnId=\"{}\"", guid(id, 100 + i)),
            );
        }
        let parent = if kind == "hierarchy" && i > 0 { 1 } else { 0 };
        let order = if kind == "hierarchy" && i > 0 {
            i - 1
        } else {
            i
        };
        connections += &format!(
            "<dgm:cxn modelId=\"{}\" type=\"parOf\" srcId=\"{}\" destId=\"{}\" srcOrd=\"{order}\" destOrd=\"0\" parTransId=\"{}\" sibTransId=\"{}\"/>",
            guid(id, 100 + i),
            guid(id, parent),
            guid(id, i + 1),
            guid(id, 200 + i),
            guid(id, 300 + i)
        );
    }
    let (boxes, edges) = geometry(kind, n);
    let mut presentation = vec![(400, 0, "diagram")];
    presentation.extend((0..n as u32).map(|i| (401 + i, i + 1, "node")));
    presentation.extend((0..edges.len() as u32).map(|i| {
        (
            500 + i,
            if kind == "hierarchy" {
                201 + i
            } else {
                300 + i
            },
            "transition",
        )
    }));
    // The model records all presOf links interleaved with presParOf links in the declared order.
    for (slot, assoc, name) in presentation {
        points += &point(
            id,
            slot,
            "pres",
            &format!(
                " presAssocID=\"{}\" presName=\"{name}\" presStyleLbl=\"node1\"",
                guid(id, assoc)
            ),
            "",
            "",
        );
        connections += &format!(
            "<dgm:cxn modelId=\"{}\" type=\"presOf\" srcId=\"{}\" destId=\"{}\" srcOrd=\"0\" destOrd=\"0\" presId=\"{uid}\"/>",
            guid(id, 800 + slot),
            guid(id, assoc),
            guid(id, slot)
        );
        if slot != 400 {
            let (cxn, order) = if slot < 500 {
                (1500 + slot - 401, slot - 401)
            } else {
                (1600 + slot - 500, n as u32 + slot - 500)
            };
            connections += &format!(
                "<dgm:cxn modelId=\"{}\" type=\"presParOf\" srcId=\"{}\" destId=\"{}\" srcOrd=\"{order}\" destOrd=\"0\" presId=\"{uid}\"/>",
                guid(id, cxn),
                guid(id, 400),
                guid(id, slot)
            );
        }
    }
    let data = format!(workbook_format!(XML_F6A039DC2D),
        G = G,
        A = A,
        DSP = DSP,
        points = points,
        connections = connections
    );
    let label = match kind {
        "process" => "Process",
        "hierarchy" => "Hierarchy",
        _ => "Cycle",
    };
    let alg = match kind {
        "process" => "lin",
        "hierarchy" => "hierRoot",
        _ => "cycle",
    };
    let axis = if kind == "hierarchy" { "des" } else { "ch" };
    let hide = if kind == "cycle" { 0 } else { 1 };
    let layout = format!(workbook_format!(XML_2FD003E147),
        G = G,
        A = A,
        uid = uid,
        label = label,
        kind = kind,
        alg = alg,
        axis = axis,
        hide = hide
    );
    let refs = "<a:lnRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:lnRef><a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:fillRef><a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\"/></a:fontRef>";
    let style = format!(workbook_format!(XML_AE2E00CE8C), G = G, A = A, refs = refs);
    let colors = format!(workbook_format!(XML_7BF15A0DC6), G = G, A = A);
    // JS Math.round semantics (including half coordinates); dimensions stay unsigned.
    let rx = |x: f64| (x * width as f64 / 1000.0 + 0.5).floor() as i64;
    let ry = |y: f64| (y * height as f64 / 600.0 + 0.5).floor() as i64;
    let mut shapes = String::new();
    for (i, (from, to)) in edges.iter().enumerate() {
        let a = &boxes[*from];
        let b = &boxes[*to];
        let dx = b.x + b.w / 2.0 - a.x - a.w / 2.0;
        let dy = b.y + b.h / 2.0 - a.y - a.h / 2.0;
        let side = |b: &BoxNode, x: f64, y: f64| {
            let t = 1.0 / (x.abs() / (b.w / 2.0)).max(y.abs() / (b.h / 2.0));
            (b.x + b.w / 2.0 + x * t, b.y + b.h / 2.0 + y * t)
        };
        let (x1, y1) = side(a, dx, dy);
        let (x2, y2) = side(b, -dx, -dy);
        shapes += &format!(
            "<dsp:sp modelId=\"{}\"><dsp:nvSpPr><dsp:cNvPr id=\"{}\" name=\"n{}-n{}\"/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm{}{}><a:off x=\"{}\" y=\"{}\"/><a:ext cx=\"{}\" cy=\"{}\"/></a:xfrm><a:prstGeom prst=\"line\"><a:avLst/></a:prstGeom><a:noFill/><a:ln w=\"19050\">{}<a:tailEnd type=\"triangle\" w=\"med\" len=\"med\"/></a:ln></dsp:spPr></dsp:sp>",
            guid(id, 500 + i as u32),
            100 + i,
            from + 1,
            to + 1,
            if x2 < x1 { " flipH=\"1\"" } else { "" },
            if y2 < y1 { " flipV=\"1\"" } else { "" },
            rx(x1.min(x2)),
            ry(y1.min(y2)),
            rx((x2 - x1).abs()),
            ry((y2 - y1).abs()),
            fill("4472C4")
        );
    }
    for (i, b) in boxes.iter().enumerate() {
        shapes += &format!(
            "<dsp:sp modelId=\"{}\"><dsp:nvSpPr><dsp:cNvPr id=\"{}\" name=\"n{}\"/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x=\"{}\" y=\"{}\"/><a:ext cx=\"{}\" cy=\"{}\"/></a:xfrm><a:prstGeom prst=\"roundRect\"><a:avLst/></a:prstGeom>{}<a:ln w=\"12700\">{}</a:ln></dsp:spPr><dsp:txBody>{}</dsp:txBody></dsp:sp>",
            guid(id, 401 + i as u32),
            i + 1,
            i + 1,
            rx(b.x),
            ry(b.y),
            rx(b.w),
            ry(b.h),
            fill("4472C4"),
            fill("2F5597"),
            text_body(&texts[i])
        );
    }
    let drawing = format!(
        "<dsp:drawing xmlns:dsp=\"{DSP}\" xmlns:a=\"{A}\"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id=\"0\" name=\"SmartArt\"/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{width}\" cy=\"{height}\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"{width}\" cy=\"{height}\"/></a:xfrm></dsp:grpSpPr>{shapes}</dsp:spTree></dsp:drawing>"
    );
    let rels = format!(workbook_format!(XML_5961861C5F), P = P, id = id);
    names
        .into_iter()
        .zip([data, layout, style, colors, drawing, rels])
        .collect()
}
struct Qualified {
    number: u64,
    revision: String,
}
fn qualified(
    zip: &mut ZipArchive<crate::cancel::Reader<File>>,
    names: &BTreeSet<String>,
    replacements: &serde_json::Map<String, Value>,
    relations: &BTreeMap<String, BTreeMap<String, Edge>>,
    drawing: &str,
    id: &str,
) -> Result<Qualified> {
    let t = Tree::parse(&part_bytes(zip, drawing, replacements)?)?;
    t.root(DRAW_NS, "wsDr")?;
    let props = t
        .nodes
        .iter()
        .enumerate()
        .filter(|(_, n)| {
            n.ns == DRAW_NS
                && n.name == "cNvPr"
                && n.attrs
                    .get("name")
                    .is_some_and(|name| format_identity_matches!(name.as_str(), ID, id = id))
        })
        .map(|(i, _)| i)
        .collect::<Vec<_>>();
    if props.len() != 1 {
        return Err(crate::coded(
            "object-bad-input",
            "SmartArt needs one scoped identity",
        ));
    }
    let prop = props[0];
    let at = t.nodes[0]
        .children
        .iter()
        .copied()
        .find(|i| t.nodes[*i].start <= t.nodes[prop].start && t.nodes[*i].end >= t.nodes[prop].end)
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing SmartArt anchor"))?;
    let from = child(&t, at, DRAW_NS, "from")?;
    let col = scalar(&t, from, "col", 16383)?;
    let row = scalar(&t, from, "row", 1048575)?;
    let dx = scalar(&t, from, "colOff", 10000000)?;
    let dy = scalar(&t, from, "rowOff", 10000000)?;
    let ext = child(&t, at, DRAW_NS, "ext")?;
    let width = attribute(&t, ext, "cx", 1, 50000000)?;
    let height = attribute(&t, ext, "cy", 1, 50000000)?;
    let number = attribute(&t, prop, "id", 1, 2147483647)?;
    let alt = t.nodes[prop].attrs.get("descr").ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing SmartArt alternative text",
        )
    })?;
    if alt.encode_utf16().count() > 256 || invalid_xml_text(alt) {
        return Err(crate::coded(
            "object-limit",
            "SmartArt alternative text exceeds its bound",
        ));
    }
    let compact = id.replace('-', "");
    let prefix = format!(workbook_format!(DIAGRAM_COMPACT), compact = compact);
    let anchor = format!(workbook_format!(XML_D9E50A0C14),
        quote(alt),
        D = D,
        A = A,
        G = G,
        R = R,
        col = col,
        dx = dx,
        row = row,
        dy = dy,
        width = width,
        height = height,
        number = number,
        id = id,
        prefix = prefix
    );
    let expected_anchor = Tree::parse(anchor.as_bytes())?;
    if !same(&t, at, &expected_anchor, 0) {
        return Err(crate::coded(
            "object-protected",
            "Extended SmartArt anchor is protected",
        ));
    }
    let paths = parts(id);
    let data = Tree::parse(&part_bytes(zip, &paths[0], replacements)?)?;
    data.root(G.as_bytes(), "dataModel")?;
    let mut texts = Vec::new();
    for (i, node) in data.nodes.iter().enumerate().filter(|(_, n)| {
        n.ns == G.as_bytes() && n.name == "pt" && n.attrs.get("type").is_some_and(|s| s == "node")
    }) {
        let body = child(&data, i, G.as_bytes(), "t")?;
        let p = child(&data, body, A_NS, "p")?;
        let run = child(&data, p, A_NS, "r")?;
        let text = child(&data, run, A_NS, "t")?;
        let value = data.nodes[text].text.clone();
        if invalid_xml_text(&value)
            || value.trim().is_empty()
            || value.encode_utf16().count() > 128
            || value.contains(['\r', '\n', '\t'])
            || value.as_bytes().windows(7).any(|w| {
                w[0] == b'_'
                    && [b'x', b'X'].contains(&w[1])
                    && w[2..6].iter().all(u8::is_ascii_hexdigit)
                    && w[6] == b'_'
            })
        {
            return Err(crate::coded(
                "object-unsupported",
                "Unsupported SmartArt node text",
            ));
        }
        let _ = node;
        texts.push(value);
    }
    if !(2..=6).contains(&texts.len()) {
        return Err(crate::coded(
            "object-refused",
            "SmartArt supports two to six nodes",
        ));
    }
    let layout = Tree::parse(&part_bytes(zip, &paths[1], replacements)?)?;
    layout.root(G.as_bytes(), "layoutDef")?;
    let kind = layout.nodes[0]
        .attrs
        .get("uniqueId")
        .and_then(|s| format_prefix!(s, URN_SMARTART_LAYOUT))
        .filter(|s| ["process", "hierarchy", "cycle"].contains(s))
        .ok_or_else(|| crate::coded("object-unsupported", "Unsupported SmartArt layout"))?;
    let expected = expected(id, kind, &texts, width, height);
    let mut size = 0;
    let mut hash = Sha256::new();
    hash.update(&t.bytes[t.nodes[at].start..t.nodes[at].end]);
    hash.update([0]);
    for (path, xml) in &expected {
        if !names.contains(path) {
            return Err(crate::coded(
                "internal-missing-field",
                "Missing SmartArt graph part",
            ));
        }
        let bytes = part_bytes(zip, path, replacements)?;
        size += bytes.len();
        if bytes.len() > 256 * 1024 || size > 1024 * 1024 {
            return Err(crate::coded(
                "object-limit",
                "SmartArt graph exceeds its bound",
            ));
        }
        let actual = Tree::parse(&bytes)?;
        let tree = Tree::parse(xml.as_bytes())?;
        if !same(&actual, 0, &tree, 0) {
            return Err(crate::coded(
                "object-protected",
                format!("Extended SmartArt graph part is protected: {path}"),
            ));
        }
        hash.update(path.as_bytes());
        hash.update([0]);
        hash.update(&bytes);
        if path != expected.keys().last().expect("graph parts") {
            hash.update([0]);
        }
    }
    for (i, path) in paths[..5].iter().enumerate() {
        let incoming = relations
            .values()
            .flat_map(|m| m.values())
            .filter(|e| e.target.as_deref() == Some(path))
            .count();
        if incoming != 1 {
            return Err(crate::coded(
                "object-protected",
                "Shared SmartArt graph is protected",
            ));
        }
        if i != 0 && relations.get(path).is_some_and(|m| !m.is_empty()) {
            return Err(crate::coded(
                "object-protected",
                "Dependent SmartArt graph is protected",
            ));
        }
    }
    let edges = relations.get(drawing).ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing SmartArt drawing relationships",
        )
    })?;
    for (attr, role, suffix) in [
        ("dm", 0, "Data"),
        ("lo", 1, "Layout"),
        ("qs", 2, "QuickStyle"),
        ("cs", 3, "Colors"),
    ] {
        let edge = edges
            .get(&format!("{prefix}_{attr}"))
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing SmartArt edge"))?;
        if edge.target.as_deref() != Some(&paths[role])
            || edge.kind != format!("{R}/diagram{suffix}")
        {
            return Err(crate::coded("object-stale", "SmartArt edge changed"));
        }
    }
    let (dir, file) = drawing
        .rsplit_once('/')
        .ok_or_else(|| crate::coded("object-malformed", "Invalid SmartArt drawing path"))?;
    let drawing_links = Tree::parse(&part_bytes(
        zip,
        &format!("{dir}/_rels/{file}.rels"),
        replacements,
    )?)?;
    for (attr, role, suffix) in [
        ("dm", 0, "Data"),
        ("lo", 1, "Layout"),
        ("qs", 2, "QuickStyle"),
        ("cs", 3, "Colors"),
    ] {
        let rid = format!("{prefix}_{attr}");
        let matches = drawing_links
            .nodes
            .iter()
            .enumerate()
            .filter(|(_, n)| {
                n.ns == PACKAGE_REL_NS
                    && n.name == "Relationship"
                    && n.attrs.get("Id") == Some(&rid)
            })
            .map(|(i, _)| i)
            .collect::<Vec<_>>();
        let xml = format!(
            "<Relationship xmlns=\"{P}\" Id=\"{rid}\" Type=\"{R}/diagram{suffix}\" Target=\"../diagrams/{}\"/>",
            paths[role]
                .rsplit('/')
                .next()
                .ok_or_else(|| crate::coded("object-malformed", "Invalid SmartArt graph path"))?
        );
        let expected = Tree::parse(xml.as_bytes())?;
        if matches.len() != 1 || !same(&drawing_links, matches[0], &expected, 0) {
            return Err(crate::coded(
                "object-protected",
                "Extended SmartArt drawing edge is protected",
            ));
        }
    }
    let types = Tree::parse(&part_bytes(zip, "[Content_Types].xml", replacements)?)?;
    for (i, path) in paths[..5].iter().enumerate() {
        let mime = if i == 4 {
            "application/vnd.ms-office.drawingml.diagramDrawing+xml".to_string()
        } else {
            format!(
                "application/vnd.openxmlformats-officedocument.drawingml.diagram{}+xml",
                ["Data", "Layout", "Style", "Colors"][i]
            )
        };
        let matches = types
            .nodes
            .iter()
            .filter(|n| {
                n.ns == TYPE_NS
                    && n.name == "Override"
                    && n.attrs.get("PartName") == Some(&format!("/{path}"))
            })
            .collect::<Vec<_>>();
        if matches.len() != 1 || matches[0].attrs.get("ContentType") != Some(&mime) {
            return Err(crate::coded(
                "internal-mismatch",
                "SmartArt content type mismatch",
            ));
        }
    }
    Ok(Qualified {
        number,
        revision: format!("{:x}", hash.finalize()),
    })
}
pub(super) fn check(
    zip: &mut ZipArchive<crate::cancel::Reader<File>>,
    original: &BTreeSet<String>,
    names: &BTreeSet<String>,
    replacements: &serde_json::Map<String, Value>,
    before_relations: &BTreeMap<String, BTreeMap<String, Edge>>,
    relations: &BTreeMap<String, BTreeMap<String, Edge>>,
    plan: &Value,
) -> Result<()> {
    let id = field(plan, "objectId")?;
    let drawing = field(plan, "drawingPart")?;
    let sheet = field(plan, "sheet")?;
    let kind = field(plan, "kind")?;
    let create = kind == "createSmartArt";
    let delete = kind == "deleteSmartArt";

    let paths = parts(id);
    let empty = serde_json::Map::new();
    if field(plan, "objectPart")? != paths[0]
        || !plan["binaryAdditions"]
            .as_object()
            .is_some_and(|b| b.is_empty())
    {
        return Err(crate::coded(
            "object-unsupported",
            "SmartArt cannot claim binary or unrelated parts",
        ));
    }
    if create && paths.iter().any(|n| original.contains(n)) {
        return Err(crate::coded(
            "object-conflict",
            "SmartArt graph identity collides",
        ));
    }
    let old = if !create {
        Some(qualified(
            zip,
            original,
            &empty,
            before_relations,
            drawing,
            id,
        )?)
    } else {
        None
    };
    let next = if !delete {
        Some(qualified(zip, names, replacements, relations, drawing, id)?)
    } else {
        None
    };
    if old
        .as_ref()
        .is_some_and(|o| plan["priorObjectRevision"].as_str() != Some(&o.revision))
    {
        return Err(crate::coded(
            "object-stale",
            "Conflict: original SmartArt revision changed",
        ));
    }
    if let (Some(a), Some(b)) = (&old, &next) {
        if a.number != b.number {
            return Err(crate::coded(
                "object-stale",
                "SmartArt nonvisual identity changed",
            ));
        }
    }
    if !delete && paths.iter().any(|n| !replacements.contains_key(n)) {
        return Err(crate::coded(
            "object-bad-input",
            "SmartArt writes require the complete graph",
        ));
    }
    if delete
        && paths.iter().any(|n| {
            names.contains(n)
                || !plan["removed"]
                    .as_array()
                    .is_some_and(|a| a.contains(&json!(n)))
        })
    {
        return Err(crate::coded(
            "object-bad-input",
            "SmartArt deletion must retire the complete graph",
        ));
    }
    let incoming = |edges: &BTreeMap<String, BTreeMap<String, Edge>>| {
        edges
            .values()
            .flat_map(|m| m.values())
            .filter(|e| e.target.as_deref() == Some(drawing))
            .count()
    };
    if (!create && incoming(before_relations) != 1) || (!delete && incoming(relations) != 1) {
        return Err(crate::coded(
            "object-protected",
            "Shared SmartArt drawing is protected",
        ));
    }
    // Preserve every sibling and unrelated edge; only the six-part graph and sheet/drawing scope may change.
    for owner in before_relations
        .keys()
        .chain(relations.keys())
        .collect::<BTreeSet<_>>()
    {
        let keep = |edges: Option<&BTreeMap<String, Edge>>| {
            edges
                .into_iter()
                .flat_map(|m| m.iter())
                .filter(|(_, e)| {
                    !(owner == sheet && e.target.as_deref() == Some(drawing))
                        && !(owner == drawing
                            && e.target.as_ref().is_some_and(|p| paths[..4].contains(p)))
                })
                .map(|(k, e)| (k.clone(), e.clone()))
                .collect::<BTreeMap<_, _>>()
        };
        if owner == &paths[0] {
            continue;
        }
        if keep(before_relations.get(owner)) != keep(relations.get(owner)) {
            return Err(crate::coded(
                "object-stale",
                "Unrelated SmartArt relationship changed",
            ));
        }
    }
    let mut types =
        |replacements: &serde_json::Map<String, Value>| -> Result<BTreeMap<String, String>> {
            let t = Tree::parse(&part_bytes(zip, "[Content_Types].xml", replacements)?)?;
            Ok(t.nodes
                .iter()
                .filter(|n| n.ns == TYPE_NS && ["Override", "Default"].contains(&n.name.as_str()))
                .filter(|n| {
                    !n.attrs.get("PartName").is_some_and(|p| {
                        p == &format!("/{drawing}")
                            || paths[..5].iter().any(|n| p == &format!("/{n}"))
                    })
                })
                .map(|n| {
                    (
                        if n.name == "Override" {
                            format!("O:{}", n.attrs.get("PartName").cloned().unwrap_or_default())
                        } else {
                            format!(
                                "D:{}",
                                n.attrs
                                    .get("Extension")
                                    .cloned()
                                    .unwrap_or_default()
                                    .to_ascii_lowercase()
                            )
                        },
                        n.attrs.get("ContentType").cloned().unwrap_or_default(),
                    )
                })
                .collect())
        };
    if types(&empty)? != types(replacements)? {
        return Err(crate::coded(
            "object-stale",
            "Unrelated SmartArt content type changed",
        ));
    }
    let links = if let Some(fragment) = plan["fragments"]
        .get(sheet)
        .and_then(|p| p.get("drawing"))
        .and_then(Value::as_str)
    {
        fragment.as_bytes().to_vec()
    } else {
        part_bytes(zip, sheet, &empty)?
    };
    let mut ids = Vec::new();
    if !links.is_empty() {
        xml_events(&links, |ns, name, attrs| {
            if ns == NS && name == b"drawing" {
                ids.push(attrs.get("r:id").cloned().ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing sheet drawing ID")
                })?);
            }
            Ok(())
        })?;
    }
    if (!delete || names.contains(drawing))
        && (ids.len() != 1
            || relations
                .get(sheet)
                .and_then(|m| m.get(&ids[0]))
                .is_none_or(|e| {
                    e.target.as_deref() != Some(drawing) || e.kind != format!("{R}/drawing")
                }))
    {
        return Err(crate::coded(
            "internal-mismatch",
            "SmartArt sheet/drawing scope mismatch",
        ));
    }
    if delete && !names.contains(drawing) && !ids.is_empty() {
        return Err(crate::coded(
            "object-unsupported",
            "Deleted SmartArt drawing must have no worksheet link",
        ));
    }
    Ok(())
}
