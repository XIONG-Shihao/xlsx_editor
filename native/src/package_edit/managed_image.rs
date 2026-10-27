//! Placement-only edits: the PNG and picture structure remain immutable.
use super::print_settings::Tree;
use super::*;
fn child(tree: &Tree, parent: usize, ns: &[u8], name: &str) -> Result<usize> {
    let found = tree.nodes[parent]
        .children
        .iter()
        .copied()
        .filter(|i| tree.nodes[*i].ns == ns && tree.nodes[*i].name == name)
        .collect::<Vec<_>>();
    if found.len() != 1 {
        return Err(crate::coded(
            "object-ambiguous",
            format!("Ambiguous managed image {name}"),
        ));
    }
    Ok(found[0])
}
fn children(tree: &Tree, parent: usize, expected: &[(&[u8], &str)]) -> Result<()> {
    let nodes = &tree.nodes[parent].children;
    if nodes.len() != expected.len()
        || nodes
            .iter()
            .zip(expected)
            .any(|(i, (ns, name))| tree.nodes[*i].ns != *ns || tree.nodes[*i].name != *name)
    {
        return Err(crate::coded(
            "object-protected",
            "Extended managed image structure is protected",
        ));
    }
    Ok(())
}
fn attrs(tree: &Tree, index: usize, allowed: &[&str]) -> Result<()> {
    if tree.nodes[index]
        .attrs
        .keys()
        .any(|key| key != "xmlns" && !key.starts_with("xmlns:") && !allowed.contains(&key.as_str()))
    {
        return Err(crate::coded(
            "object-protected",
            "Extended managed image attributes are protected",
        ));
    }
    Ok(())
}
fn number(tree: &Tree, index: usize, low: u64, high: u64) -> Result<u64> {
    let value = &tree.nodes[index].text;
    if value.is_empty() || !value.bytes().all(|ch| ch.is_ascii_digit()) {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid managed image coordinate",
        ));
    }
    let value = value.parse::<u64>().map_err(err)?;
    if value < low || value > high {
        return Err(crate::coded(
            "object-limit",
            "Managed image coordinate exceeds its bound",
        ));
    }
    Ok(value)
}
fn anchor(tree: &Tree, id: &str, width: u64, height: u64) -> Result<Value> {
    tree.root(DRAW_NS, "wsDr")?;
    let candidates = tree.nodes[0]
        .children
        .iter()
        .copied()
        .filter(|index| {
            let node = &tree.nodes[*index];
            tree.nodes.iter().any(|item| {
                item.start >= node.start
                    && item.end <= node.end
                    && item.ns == DRAW_NS
                    && item.name == "cNvPr"
                    && item
                        .attrs
                        .get("name")
                        .is_some_and(|name| format_identity_matches!(name.as_str(), ID, id = id))
            })
        })
        .collect::<Vec<_>>();
    if candidates.len() != 1 {
        return Err(crate::coded(
            "object-bad-input",
            "Managed image has no unique anchor",
        ));
    }
    let root = candidates[0];
    if tree.nodes[root].ns != DRAW_NS || tree.nodes[root].name != "oneCellAnchor" {
        return Err(crate::coded(
            "object-unsupported",
            "Only one-cell PNG placement is qualified",
        ));
    }
    children(
        tree,
        root,
        &[
            (DRAW_NS, "from"),
            (DRAW_NS, "ext"),
            (DRAW_NS, "pic"),
            (DRAW_NS, "clientData"),
        ],
    )?;
    let from = child(tree, root, DRAW_NS, "from")?;
    children(
        tree,
        from,
        &[
            (DRAW_NS, "col"),
            (DRAW_NS, "colOff"),
            (DRAW_NS, "row"),
            (DRAW_NS, "rowOff"),
        ],
    )?;
    let positions = [
        ("col", 0, 16383),
        ("colOff", 0, 10000000),
        ("row", 0, 1048575),
        ("rowOff", 0, 10000000),
    ]
    .iter()
    .map(|(name, lo, hi)| child(tree, from, DRAW_NS, name).and_then(|i| number(tree, i, *lo, *hi)))
    .collect::<Result<Vec<_>>>()?;
    let _ = positions;
    let ext = child(tree, root, DRAW_NS, "ext")?;
    let extent = |name: &str| -> Result<u64> {
        let text = tree.nodes[ext]
            .attrs
            .get(name)
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing image extent"))?;
        if text.is_empty() || !text.bytes().all(|ch| ch.is_ascii_digit()) {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid image extent",
            ));
        }
        let value = text.parse::<u64>().map_err(err)?;
        if !(1..=50000000).contains(&value) {
            return Err(crate::coded(
                "object-limit",
                "Image extent exceeds its bound",
            ));
        }
        Ok(value)
    };
    let cx = extent("cx")?;
    let cy = extent("cy")?;
    if (cx * height).abs_diff(cy * width) > width.max(height) {
        return Err(crate::coded(
            "object-unsupported",
            "PNG placement must preserve its aspect ratio",
        ));
    }
    let pic = child(tree, root, DRAW_NS, "pic")?;
    children(
        tree,
        pic,
        &[
            (DRAW_NS, "nvPicPr"),
            (DRAW_NS, "blipFill"),
            (DRAW_NS, "spPr"),
        ],
    )?;
    let nv = child(tree, pic, DRAW_NS, "nvPicPr")?;
    children(tree, nv, &[(DRAW_NS, "cNvPr"), (DRAW_NS, "cNvPicPr")])?;
    let properties = child(tree, nv, DRAW_NS, "cNvPr")?;
    if !tree.nodes[properties]
        .attrs
        .get("name")
        .is_some_and(|name| format_identity_matches!(name.as_str(), ID, id = id))
        || tree.nodes[properties]
            .attrs
            .get("descr")
            .is_some_and(|text| {
                text.chars().count() > 256
                    || text
                        .chars()
                        .any(|ch| ch < ' ' && !['\t', '\n', '\r'].contains(&ch))
            })
    {
        return Err(crate::coded(
            "object-malformed",
            "Invalid image description/identity",
        ));
    }
    let fill = child(tree, pic, DRAW_NS, "blipFill")?;
    children(tree, fill, &[(A_NS, "blip"), (A_NS, "stretch")])?;
    let stretch = child(tree, fill, A_NS, "stretch")?;
    children(tree, stretch, &[(A_NS, "fillRect")])?;
    let shape = child(tree, pic, DRAW_NS, "spPr")?;
    children(tree, shape, &[(A_NS, "prstGeom")])?;
    let geometry = child(tree, shape, A_NS, "prstGeom")?;
    children(tree, geometry, &[(A_NS, "avLst")])?;
    if tree.nodes[geometry].attrs.get("prst").map(String::as_str) != Some("rect") {
        return Err(crate::coded(
            "object-protected",
            "Alternative PNG geometry is protected",
        ));
    }
    fn model(
        tree: &Tree,
        index: usize,
        from: usize,
        properties: usize,
        ext: usize,
    ) -> Result<Value> {
        let node = &tree.nodes[index];
        let allowed = if index == properties {
            vec!["id", "name", "descr"]
        } else if index == ext {
            vec!["cx", "cy"]
        } else if node.ns == A_NS && node.name == "blip" {
            vec!["r:embed"]
        } else if node.ns == A_NS && node.name == "prstGeom" {
            vec!["prst"]
        } else {
            vec![]
        };
        attrs(tree, index, &allowed)?;
        let values = node
            .attrs
            .iter()
            .filter(|(key, _)| {
                key.as_str() != "xmlns"
                    && !key.starts_with("xmlns:")
                    && !(index == properties && key.as_str() == "descr")
                    && index != ext
            })
            .collect::<BTreeMap<_, _>>();
        let coordinate = tree.nodes[from].children.contains(&index);
        if !coordinate && !node.text.trim().is_empty() {
            return Err(crate::coded(
                "object-protected",
                "Extended picture text is protected",
            ));
        }
        let children = node
            .children
            .iter()
            .map(|i| model(tree, *i, from, properties, ext))
            .collect::<Result<Vec<_>>>()?;
        Ok(json!([
            String::from_utf8_lossy(&node.ns),
            node.name,
            values,
            if coordinate { "#" } else { "" },
            children
        ]))
    }
    model(tree, root, from, properties, ext)
}
pub(super) fn check(before: &[u8], after: &[u8], asset: &[u8], id: &str, part: &str) -> Result<()> {
    if asset.len() < 24 {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid managed PNG",
        ));
    }
    let width = u32::from_be_bytes(asset[16..20].try_into().map_err(err)?) as u64;
    let height = u32::from_be_bytes(asset[20..24].try_into().map_err(err)?) as u64;
    let sha = format!("{:x}", Sha256::digest(asset));
    let asset_id = format_prefix!(part, XL_MEDIA_IMAGE)
        .and_then(|value| value.strip_suffix(&format!("-{sha}.png")));
    if !asset_id.is_some_and(uuid) {
        return Err(crate::coded(
            "internal-mismatch",
            "Managed PNG asset identity mismatch",
        ));
    }
    png(
        asset,
        &json!({"sha256":sha,"widthPx":width,"heightPx":height}),
    )?;
    let old = anchor(&Tree::parse(before)?, id, width, height)?;
    let new = anchor(&Tree::parse(after)?, id, width, height)?;
    if old != new {
        return Err(crate::coded(
            "object-stale",
            "Image placement changed immutable picture metadata",
        ));
    }
    Ok(())
}
