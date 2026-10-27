//! Bounded owned open DrawingML strokes. Source model and prior revision are checked independently.
use super::managed_shape::{attribute, child, quote, same, scalar};
use super::print_settings::Tree;
use super::*;

struct Stroke {
    number: u64,
    revision: String,
    model: Value,
}
fn anchor(bytes: &[u8], id: &str) -> Result<Stroke> {
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
            "Stroke has no unique owned anchor",
        ));
    }
    let at = candidates[0];
    if t.nodes[at].ns != DRAW_NS || t.nodes[at].name != "oneCellAnchor" {
        return Err(crate::coded(
            "object-unsupported",
            "Only one-cell strokes are supported",
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
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing stroke alternative text"))?;
    if alt.encode_utf16().count() > 256
        || alt.as_bytes().windows(7).any(|w| {
            w[0] == b'_'
                && [b'x', b'X'].contains(&w[1])
                && w[2..6].iter().all(u8::is_ascii_hexdigit)
                && w[6] == b'_'
        })
    {
        return Err(crate::coded(
            "object-limit",
            "Stroke alternative text exceeds its literal profile",
        ));
    }
    let pr = child(&t, sp, DRAW_NS, "spPr")?;
    let line = child(&t, pr, A_NS, "ln")?;
    let solid = child(&t, line, A_NS, "solidFill")?;
    let rgb = child(&t, solid, A_NS, "srgbClr")?;
    let alpha = child(&t, rgb, A_NS, "alpha")?;
    let color = t.nodes[rgb]
        .attrs
        .get("val")
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing stroke color"))?;
    let cap = t.nodes[line]
        .attrs
        .get("cap")
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing stroke cap"))?;
    let pen_width = attribute(&t, line, "w", 1, 114300)?;
    let opacity = attribute(&t, alpha, "val", 1, 100000)?;
    if ![
        ("000000", 19050, 100000, "rnd"),
        ("D32F2F", 19050, 100000, "rnd"),
        ("FFFF00", 114300, 30000, "flat"),
    ]
    .iter()
    .any(|(c, w, a, k)| color == c && pen_width == *w && opacity == *a && cap == k)
    {
        return Err(crate::coded(
            "object-unsupported",
            "Unsupported stroke pen profile",
        ));
    }
    let custom = child(&t, pr, A_NS, "custGeom")?;
    let paths = child(&t, custom, A_NS, "pathLst")?;
    let path = child(&t, paths, A_NS, "path")?;
    if attribute(&t, path, "w", 1, 50000000)? != width
        || attribute(&t, path, "h", 1, 50000000)? != height
        || t.nodes[path].children.len() < 2
        || t.nodes[path].children.len() > 512
    {
        return Err(crate::coded(
            "object-limit",
            "Stroke point count or extent exceeds its bound",
        ));
    }
    let pad = (pen_width + 1) / 2;
    let mut points = Vec::new();
    for (i, step) in t.nodes[path].children.iter().enumerate() {
        let n = &t.nodes[*step];
        if n.ns != A_NS || n.name != if i == 0 { "moveTo" } else { "lnTo" } {
            return Err(crate::coded(
                "object-unsupported",
                "Only open stroke polylines are qualified",
            ));
        }
        let point = child(&t, *step, A_NS, "pt")?;
        let x = attribute(&t, point, "x", pad, 50000000)?;
        let y = attribute(&t, point, "y", pad, 50000000)?;
        if x > width
            .checked_sub(pad)
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid stroke width"))?
            || y > height
                .checked_sub(pad)
                .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid stroke height"))?
            || points.last() == Some(&(x, y))
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid padded stroke points",
            ));
        }
        points.push((x, y));
    }
    if points.iter().map(|p| p.0).min() != Some(pad)
        || points.iter().map(|p| p.1).min() != Some(pad)
        || points.iter().map(|p| p.0).max() != Some(width - pad)
        || points.iter().map(|p| p.1).max() != Some(height - pad)
    {
        return Err(crate::coded(
            "object-limit",
            "Stroke points do not tightly match their bounds",
        ));
    }
    let steps = points
        .iter()
        .enumerate()
        .map(|(i, (x, y))| {
            let tag = if i == 0 { "moveTo" } else { "lnTo" };
            format!("<a:{tag}><a:pt x=\"{x}\" y=\"{y}\"/></a:{tag}>")
        })
        .collect::<String>();
    let xml = format!(workbook_format!(XML_6C801357F7),
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
        steps = steps,
        pen_width = pen_width,
        cap = cap,
        color = color,
        opacity = opacity
    );
    let expected = Tree::parse(xml.as_bytes())?;
    if !same(&t, at, &expected, 0) {
        return Err(crate::coded(
            "object-protected",
            "Imported or extended stroke is protected",
        ));
    }
    let mut raw = t.bytes[t.nodes[at].start..t.nodes[at].end].to_vec();
    raw.extend_from_slice(&[0, 0, 0]);
    Ok(Stroke {
        number,
        revision: format!("{:x}", Sha256::digest(raw)),
        model: json!([width, height, color, cap, pen_width, opacity, points]),
    })
}
pub(super) fn check(
    before: Option<&[u8]>,
    after: Option<&[u8]>,
    id: &str,
    revision: Option<&str>,
) -> Result<()> {
    let old = before.map(|bytes| anchor(bytes, id)).transpose()?;
    let next = after.map(|bytes| anchor(bytes, id)).transpose()?;
    if let Some(old) = &old {
        if revision != Some(old.revision.as_str()) {
            return Err(crate::coded(
                "object-stale",
                "Conflict: original stroke revision changed",
            ));
        }
    }
    if let (Some(old), Some(next)) = (&old, &next) {
        if old.model != next.model {
            return Err(crate::coded(
                "object-unsupported",
                "Stroke moves cannot change its points, dimensions or pen",
            ));
        }
        if old.number != next.number {
            return Err(crate::coded(
                "object-stale",
                "Stroke nonvisual identity changed",
            ));
        }
    }
    if old.is_none() && next.is_none() {
        return Err(crate::coded(
            "object-bad-input",
            "Stroke edit has no owned model",
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    const SOURCE: &str = include_str!("managed_ink_fixture.xml");
    const ID: &str = "11111111-1111-4111-8111-111111111111";
    #[test]
    fn accepts_three_closed_styles() {
        for id in [
            ID,
            "22222222-2222-4222-8222-222222222222",
            "33333333-3333-4333-8333-333333333333",
        ] {
            assert!(anchor(SOURCE.as_bytes(), id).is_ok());
            assert!(check(None, Some(SOURCE.as_bytes()), id, None).is_ok());
        }
    }
    #[test]
    fn moves_and_alt_text_keep_the_model_and_original_revision() {
        let old = anchor(SOURCE.as_bytes(), ID).unwrap();
        let next = SOURCE
            .replacen(
                "<xdr:colOff>180975</xdr:colOff>",
                "<xdr:colOff>200000</xdr:colOff>",
                1,
            )
            .replacen("Black zigzag", "Edited &amp; stroke", 1);
        assert_ne!(next, SOURCE);
        assert!(
            check(
                Some(SOURCE.as_bytes()),
                Some(next.as_bytes()),
                ID,
                Some(&old.revision)
            )
            .is_ok()
        );
        assert!(check(Some(SOURCE.as_bytes()), None, ID, Some(&old.revision)).is_ok());
        assert!(
            check(
                Some(SOURCE.as_bytes()),
                Some(next.as_bytes()),
                ID,
                Some(&"0".repeat(64))
            )
            .is_err()
        );
    }
    #[test]
    fn a_move_cannot_change_points_pen_or_nonvisual_identity() {
        let old = anchor(SOURCE.as_bytes(), ID).unwrap();
        for next in [
            SOURCE.replacen("y=\"57150\"", "y=\"57151\"", 1),
            SOURCE.replacen("val=\"000000\"", "val=\"D32F2F\"", 1),
            SOURCE.replacen(
                workbook_format!(XML_3E43FDA91C),
                workbook_format!(XML_870DB7DE39),
                1,
            ),
        ] {
            assert_ne!(next, SOURCE);
            assert!(
                check(
                    Some(SOURCE.as_bytes()),
                    Some(next.as_bytes()),
                    ID,
                    Some(&old.revision)
                )
                .is_err()
            );
        }
    }
    #[test]
    fn malformed_styles_effects_fill_and_closed_paths_do_not_gain_admission() {
        for next in [
            SOURCE.replacen("fill=\"none\"", "fill=\"norm\"", 1),
            SOURCE.replacen("val=\"100000\"", "val=\"99999\"", 1),
            SOURCE.replacen("<xdr:sp>", "<xdr:sp macro=\"run\">", 1),
            SOURCE.replacen("</a:path>", "<a:close/></a:path>", 1),
            SOURCE.replacen("<a:round/>", "<a:miter/>", 1),
        ] {
            assert!(anchor(next.as_bytes(), ID).is_err());
        }
    }
}
