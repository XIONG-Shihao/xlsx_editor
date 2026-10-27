//! Editable inline text/shapes and three closed vector icon profiles, independently checked against the source anchor.
use super::print_settings::Tree;
use super::*;
pub(super) fn child(t: &Tree, parent: usize, ns: &[u8], name: &str) -> Result<usize> {
    let found = t.nodes[parent]
        .children
        .iter()
        .copied()
        .filter(|i| t.nodes[*i].ns == ns && t.nodes[*i].name == name)
        .collect::<Vec<_>>();
    if found.len() != 1 {
        return Err(crate::coded(
            "object-ambiguous",
            format!("Ambiguous shape {name}"),
        ));
    }
    Ok(found[0])
}
pub(super) fn scalar(t: &Tree, parent: usize, name: &str, high: u64) -> Result<u64> {
    let i = child(t, parent, DRAW_NS, name)?;
    let s = &t.nodes[i].text;
    if s.is_empty() || !s.bytes().all(|c| c.is_ascii_digit()) {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid shape coordinate",
        ));
    }
    s.parse::<u64>()
        .map_err(err)?
        .checked_add(0)
        .filter(|v| *v <= high)
        .ok_or_else(|| crate::coded("object-limit", "Shape coordinate exceeds its bound"))
}
pub(super) fn attribute(t: &Tree, i: usize, key: &str, low: u64, high: u64) -> Result<u64> {
    let s = t.nodes[i]
        .attrs
        .get(key)
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing shape attribute"))?;
    if s.is_empty() || !s.bytes().all(|c| c.is_ascii_digit()) {
        return Err(crate::coded("object-malformed", "Invalid shape attribute"));
    }
    s.parse::<u64>()
        .map_err(err)?
        .checked_add(0)
        .filter(|v| *v >= low && *v <= high)
        .ok_or_else(|| crate::coded("object-limit", "Shape attribute exceeds its bound"))
}
pub(super) fn quote(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('\"', "&quot;")
        .replace('\r', "&#13;")
        .replace('\n', "&#10;")
        .replace('\t', "&#9;")
}
fn fill(color: &str) -> String {
    if color.is_empty() {
        "<a:noFill/>".into()
    } else {
        format!("<a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill>")
    }
}
pub(super) fn same(a: &Tree, i: usize, b: &Tree, j: usize) -> bool {
    let left = &a.nodes[i];
    let right = &b.nodes[j];
    let attrs = |t: &Tree, n: usize| {
        t.nodes[n]
            .attrs
            .iter()
            .filter(|(k, _)| !k.starts_with("xmlns"))
            .map(|(k, v)| (k.clone(), v.clone()))
            .collect::<BTreeMap<_, _>>()
    };
    left.ns == right.ns
        && left.name == right.name
        && attrs(a, i) == attrs(b, j)
        && left.children.len() == right.children.len()
        && if left.children.is_empty() {
            left.text == right.text
        } else {
            left.text.trim().is_empty()
                && right.text.trim().is_empty()
                && left
                    .children
                    .iter()
                    .zip(&right.children)
                    .all(|(x, y)| same(a, *x, b, *y))
        }
}
struct Qualified {
    number: u64,
    revision: String,
}
fn anchor(bytes: &[u8], id: &str) -> Result<Qualified> {
    let t = Tree::parse(bytes)?;
    t.root(DRAW_NS, "wsDr")?;
    let candidates = t.nodes[0]
        .children
        .iter()
        .copied()
        .filter(|i| {
            t.nodes.iter().any(|n| {
                n.start >= t.nodes[*i].start
                    && n.end <= t.nodes[*i].end
                    && n.ns == DRAW_NS
                    && n.name == "cNvPr"
                    && n.attrs
                        .get("name")
                        .is_some_and(|name| format_identity_matches!(name.as_str(), ID, id = id))
            })
        })
        .collect::<Vec<_>>();
    if candidates.len() != 1 {
        return Err(crate::coded(
            "object-bad-input",
            "Shape has no unique owned anchor",
        ));
    }
    let at = candidates[0];
    if t.nodes[at].name != "oneCellAnchor" || t.nodes[at].ns != DRAW_NS {
        return Err(crate::coded(
            "object-unsupported",
            "Only one-cell shape anchors are supported",
        ));
    }
    let from = child(&t, at, DRAW_NS, "from")?;
    let col = scalar(&t, from, "col", 16383)?;
    let row = scalar(&t, from, "row", 1048575)?;
    let dx = scalar(&t, from, "colOff", 10000000)?;
    let dy = scalar(&t, from, "rowOff", 10000000)?;
    let ext = child(&t, at, DRAW_NS, "ext")?;
    let width = attribute(&t, ext, "cx", 1, 50000000)?;
    let height = attribute(&t, ext, "cy", 1, 50000000)?;
    let sp = child(&t, at, DRAW_NS, "sp")?;
    let nv = child(&t, sp, DRAW_NS, "nvSpPr")?;
    let props = child(&t, nv, DRAW_NS, "cNvPr")?;
    let number = attribute(&t, props, "id", 1, 2147483647)?;
    let alt = t.nodes[props]
        .attrs
        .get("descr")
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing shape alternative text"))?;
    if alt.encode_utf16().count() > 256 {
        return Err(crate::coded(
            "object-limit",
            "Shape alternative text exceeds its bound",
        ));
    }
    // Icons are real filled custom paths with no text body or relationship.
    let shape_props = child(&t, sp, DRAW_NS, "spPr")?;
    if t.nodes[shape_props]
        .children
        .iter()
        .any(|i| t.nodes[*i].ns == A_NS && t.nodes[*i].name == "custGeom")
    {
        let solid = child(&t, shape_props, A_NS, "solidFill")?;
        let rgb = child(&t, solid, A_NS, "srgbClr")?;
        let color = t.nodes[rgb]
            .attrs
            .get("val")
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing icon color"))?;
        if color.len() != 6
            || !color
                .bytes()
                .all(|c| c.is_ascii_digit() || (b'A'..=b'F').contains(&c))
        {
            return Err(crate::coded(
                "object-unsupported",
                "Icon color is outside its sRGB profile",
            ));
        }
        let paths: &[&[(u16, u16)]] = &[
            &[(12, 50), (35, 73), (88, 20), (96, 28), (35, 89), (4, 58)],
            &[
                (50, 4),
                (61, 36),
                (96, 36),
                (68, 57),
                (79, 91),
                (50, 70),
                (21, 91),
                (32, 57),
                (4, 36),
                (39, 36),
            ],
            &[
                (40, 5),
                (60, 5),
                (60, 40),
                (95, 40),
                (95, 60),
                (60, 60),
                (60, 95),
                (40, 95),
                (40, 60),
                (5, 60),
                (5, 40),
                (40, 40),
            ],
        ];
        for points in paths {
            let path = points
                .iter()
                .enumerate()
                .map(|(i, (x, y))| {
                    let tag = if i == 0 { "moveTo" } else { "lnTo" };
                    format!("<a:{tag}><a:pt x=\"{x}\" y=\"{y}\"/></a:{tag}>")
                })
                .collect::<String>();
            let geometry = format!(
                "<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l=\"0\" t=\"0\" r=\"100\" b=\"100\"/><a:pathLst><a:path w=\"100\" h=\"100\" stroke=\"0\" extrusionOk=\"0\">{path}<a:close/></a:path></a:pathLst></a:custGeom>"
            );
            let xml = format!(workbook_format!(XML_BC0D1F4093),
                String::from_utf8_lossy(DRAW_NS),
                String::from_utf8_lossy(A_NS),
                quote(alt),
                fill(color),
                col = col,
                dx = dx,
                row = row,
                dy = dy,
                width = width,
                height = height,
                number = number,
                id = id,
                geometry = geometry
            );
            let expected = Tree::parse(xml.as_bytes())?;
            if same(&t, at, &expected, 0) {
                let mut raw = t.bytes[t.nodes[at].start..t.nodes[at].end].to_vec();
                raw.extend_from_slice(&[0, 0, 0]);
                return Ok(Qualified {
                    number,
                    revision: format!("{:x}", Sha256::digest(raw)),
                });
            }
        }
        return Err(crate::coded(
            "object-protected",
            "Imported or extended icon path is protected",
        ));
    }
    let body = child(&t, sp, DRAW_NS, "txBody")?;
    let mut texts = Vec::new();
    for paragraph in &t.nodes[body].children {
        let n = &t.nodes[*paragraph];
        if n.ns == A_NS && n.name == "p" {
            let run = child(&t, *paragraph, A_NS, "r")?;
            let text = child(&t, run, A_NS, "t")?;
            texts.push(t.nodes[text].text.clone());
        }
    }
    let text = texts.join("\n");
    let escape = text.as_bytes().windows(7).any(|w| {
        w[0] == b'_'
            && [b'x', b'X'].contains(&w[1])
            && w[2..6].iter().all(u8::is_ascii_hexdigit)
            && w[6] == b'_'
    });
    if texts.is_empty()
        || texts.len() > 8
        || text.encode_utf16().count() > 1024
        || text.contains(['\r', '\t'])
        || escape
    {
        return Err(crate::coded(
            "object-limit",
            "Shape text exceeds its ordinary paragraph profile",
        ));
    }
    let body_props = child(&t, body, A_NS, "bodyPr")?;
    if t.nodes[body_props].attrs.contains_key("fromWordArt") {
        if texts.len() > 4 || text.encode_utf16().count() > 256 {
            return Err(crate::coded(
                "object-limit",
                "WordArt text exceeds its short text profile",
            ));
        }
        for (color, outline, shadow) in [
            ("4472C4", "", false),
            ("C65911", "C65911", false),
            ("7030A0", "", true),
        ] {
            let line = format!("<a:ln w=\"12700\">{}</a:ln>", fill(outline));
            let effect = if shadow {
                "<a:outerShdw blurRad=\"19050\" dist=\"53882\" dir=\"2700000\" sx=\"100000\" sy=\"100000\" kx=\"0\" ky=\"0\" algn=\"ctr\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"35000\"/></a:srgbClr></a:outerShdw>"
            } else {
                ""
            };
            let run = format!(
                "<a:rPr lang=\"en-US\" sz=\"3200\" b=\"1\" i=\"0\">{line}{}<a:effectLst>{effect}</a:effectLst><a:latin typeface=\"Arial\"/></a:rPr>",
                fill(if outline.is_empty() { color } else { "" })
            );
            let paragraphs=texts.iter().map(|s|format!("<a:p><a:pPr algn=\"ctr\"/><a:r>{run}<a:t xml:space=\"preserve\">{}</a:t></a:r></a:p>",quote(s))).collect::<String>();
            let xml = format!(workbook_format!(XML_A5FAA3BD99),
                String::from_utf8_lossy(DRAW_NS),
                String::from_utf8_lossy(A_NS),
                quote(alt),
                col = col,
                dx = dx,
                row = row,
                dy = dy,
                width = width,
                height = height,
                number = number,
                id = id,
                paragraphs = paragraphs
            );
            let expected = Tree::parse(xml.as_bytes())?;
            if same(&t, at, &expected, 0) {
                let mut raw = t.bytes[t.nodes[at].start..t.nodes[at].end].to_vec();
                raw.extend_from_slice(&[0, 0, 0]);
                return Ok(Qualified {
                    number,
                    revision: format!("{:x}", Sha256::digest(raw)),
                });
            }
        }
        return Err(crate::coded(
            "object-protected",
            "Imported or extended WordArt profile is protected",
        ));
    }
    // Each declared sample has a real geometry/fill/text body. Unknown effects,
    // relationships/macros, custom paths, extra runs or fonts fail full tree equality.
    for (geometry, background, line, color, align, txbox) in [
        ("rect", "", "", "1F2937", "l", "1"),
        ("rect", "EAF2F8", "4472C4", "17365D", "l", "1"),
        ("rect", "DCE6F1", "4472C4", "17365D", "ctr", "0"),
        ("ellipse", "E2EFDA", "70AD47", "375623", "ctr", "0"),
        ("rightArrow", "FFF2CC", "BF9000", "7F6000", "ctr", "0"),
    ] {
        let paragraphs=texts.iter().map(|s|format!("<a:p><a:pPr algn=\"{align}\"/><a:r><a:rPr lang=\"en-US\" sz=\"1400\" b=\"0\" i=\"0\">{}<a:latin typeface=\"Arial\"/></a:rPr><a:t xml:space=\"preserve\">{}</a:t></a:r></a:p>",fill(color),quote(s))).collect::<String>();
        let xml = format!(workbook_format!(XML_853DF0B88B),
            String::from_utf8_lossy(DRAW_NS),
            String::from_utf8_lossy(A_NS),
            quote(alt),
            fill(background),
            fill(line),
            col = col,
            dx = dx,
            row = row,
            dy = dy,
            width = width,
            height = height,
            number = number,
            id = id,
            txbox = txbox,
            geometry = geometry,
            paragraphs = paragraphs
        );
        let expected = Tree::parse(xml.as_bytes())?;
        if same(&t, at, &expected, 0) {
            let mut raw = t.bytes[t.nodes[at].start..t.nodes[at].end].to_vec();
            raw.extend_from_slice(&[0, 0, 0]);
            return Ok(Qualified {
                number,
                revision: format!("{:x}", Sha256::digest(raw)),
            });
        }
    }
    Err(crate::coded(
        "object-protected",
        "Imported or extended shape profile is protected",
    ))
}
pub(super) fn check(
    before: Option<&[u8]>,
    after: Option<&[u8]>,
    id: &str,
    revision: Option<&str>,
) -> Result<()> {
    let old = before.map(|bytes| anchor(bytes, id)).transpose()?;
    let next = after.map(|bytes| anchor(bytes, id)).transpose()?;
    if let Some(original) = &old {
        if revision != Some(original.revision.as_str()) {
            return Err(crate::coded(
                "object-stale",
                "Conflict: original shape revision changed",
            ));
        }
    }
    if let (Some(old), Some(next)) = (&old, &next) {
        if old.number != next.number {
            return Err(crate::coded(
                "object-stale",
                "Shape nonvisual identity changed",
            ));
        }
    }
    if old.is_none() && next.is_none() {
        return Err(crate::coded(
            "object-bad-input",
            "Shape edit has no owned model",
        ));
    }
    Ok(())
}
