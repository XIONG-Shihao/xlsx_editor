//! Private GLB samples and Office model3d graphs; independent of renderer/JS admission.
use super::managed_shape::{attribute, child, quote, same, scalar};
use super::print_settings::Tree;
use super::*;
pub(super) const M: &str = "http://schemas.microsoft.com/office/drawing/2017/model3d";
pub(super) const REL: &str = "http://schemas.microsoft.com/office/2017/06/relationships/model3d";
const MIME: &str = "model/gltf-binary";
const PRIVATE: &str = workbook_format!(URN_MODEL3D_SAMPLES_V1);
const LIMIT: usize = 64 * 1024;
fn sample(name: &str) -> Result<(Vec<[f64; 3]>, Vec<[usize; 3]>, &'static str)> {
    Ok(match name {
        "cube" => (
            vec![
                [-1., -1., -1.],
                [1., -1., -1.],
                [1., 1., -1.],
                [-1., 1., -1.],
                [-1., -1., 1.],
                [1., -1., 1.],
                [1., 1., 1.],
                [-1., 1., 1.],
            ],
            vec![
                [4, 5, 6],
                [4, 6, 7],
                [1, 0, 3],
                [1, 3, 2],
                [5, 1, 2],
                [5, 2, 6],
                [0, 4, 7],
                [0, 7, 3],
                [7, 6, 2],
                [7, 2, 3],
                [0, 1, 5],
                [0, 5, 4],
            ],
            "Cube",
        ),
        "pyramid" => (
            vec![
                [-1., -1., -1.],
                [1., -1., -1.],
                [1., -1., 1.],
                [-1., -1., 1.],
                [0., 1., 0.],
            ],
            vec![
                [0, 1, 4],
                [1, 2, 4],
                [2, 3, 4],
                [3, 0, 4],
                [0, 3, 2],
                [0, 2, 1],
            ],
            "Pyramid",
        ),
        "octahedron" => (
            vec![
                [1., 0., 0.],
                [-1., 0., 0.],
                [0., 1., 0.],
                [0., -1., 0.],
                [0., 0., 1.],
                [0., 0., -1.],
            ],
            vec![
                [2, 0, 4],
                [2, 4, 1],
                [2, 1, 5],
                [2, 5, 0],
                [3, 4, 0],
                [3, 1, 4],
                [3, 5, 1],
                [3, 0, 5],
            ],
            "Octahedron",
        ),
        _ => return Err(crate::coded("object-unsupported", "Unsupported 3D sample")),
    })
}
fn matching_json(a: &Value, b: &Value, path: &str) -> bool {
    match (a, b) {
        (Value::Object(a), Value::Object(b)) => {
            a.len() == b.len()
                && a.iter().all(|(k, v)| {
                    b.get(k)
                        .is_some_and(|w| matching_json(v, w, &format!("{path}/{k}")))
                })
        }
        (Value::Array(a), Value::Array(b)) => {
            a.len() == b.len()
                && a.iter()
                    .zip(b)
                    .enumerate()
                    .all(|(i, (v, w))| matching_json(v, w, &format!("{path}/{i}")))
        }
        (Value::Number(a), Value::Number(b)) => a.as_f64().zip(b.as_f64()).is_some_and(|(x, y)| {
            x == y
                || ((path.starts_with("/nodes/0/rotation/")
                    || path.starts_with("/materials/0/pbrMetallicRoughness/baseColorFactor/"))
                    && (x - y).abs() <= 1e-12)
        }),
        _ => a == b,
    }
}
pub(super) fn glb(bytes: &[u8]) -> Result<Value> {
    if bytes.len() < 28 || bytes.len() > LIMIT {
        return Err(crate::coded("object-limit", "3D GLB exceeds its bound"));
    }
    let u32_at = |i: usize| -> u32 {
        u32::from_le_bytes(bytes[i..i + 4].try_into().expect("checked GLB boundary"))
    };
    if u32_at(0) != 0x46546c67 || u32_at(4) != 2 || u32_at(8) as usize != bytes.len() {
        return Err(crate::coded("internal-invalid-field", "Invalid GLB header"));
    }
    let jl = u32_at(12) as usize;
    let bh = 20usize
        .checked_add(jl)
        .ok_or_else(|| crate::coded("object-limit", "GLB overflow"))?;
    if jl < 4
        || jl % 4 != 0
        || bh.checked_add(8).is_none_or(|n| n > bytes.len())
        || u32_at(16) != 0x4e4f534a
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid GLB JSON chunk",
        ));
    }
    let bl = u32_at(bh) as usize;
    let start = bh + 8;
    if bl % 4 != 0 || start.checked_add(bl) != Some(bytes.len()) || u32_at(bh + 4) != 0x004e4942 {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid GLB binary chunk",
        ));
    }
    let json: Value = serde_json::from_slice(&bytes[20..bh]).map_err(err)?;
    let model = &json["extras"][workbook_format!(MODEL3D)];
    if !model.as_object().is_some_and(|m| m.len() == 3)
        || !model["rotation"].as_object().is_some_and(|m| m.len() == 3)
    {
        return Err(crate::coded(
            "internal-missing-field",
            "Missing closed 3D model",
        ));
    }
    let color = field(model, "color")?;
    if color.len() != 6
        || !color
            .bytes()
            .all(|b| b.is_ascii_digit() || (b'A'..=b'F').contains(&b))
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid 3D material color",
        ));
    }
    let angle = |axis: &str| -> Result<f64> {
        let n = model["rotation"][axis]
            .as_f64()
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid 3D angle"))?;
        if n.fract() != 0. || !(-180. ..=180.).contains(&n) {
            return Err(crate::coded("object-limit", "3D angle exceeds its bound"));
        }
        Ok(n * std::f64::consts::PI / 360.)
    };
    let (x, y, z) = (angle("x")?, angle("y")?, angle("z")?);
    let (sx, cx) = (x.sin(), x.cos());
    let (sy, cy) = (y.sin(), y.cos());
    let (sz, cz) = (z.sin(), z.cos());
    let rotation = [
        sx * cy * cz - cx * sy * sz,
        cx * sy * cz + sx * cy * sz,
        cx * cy * sz - sx * sy * cz,
        cx * cy * cz + sx * sy * sz,
    ];
    let (vertices, faces, label) = sample(field(model, "sample")?)?;
    let count = faces.len() * 3;
    let pb = count * 12;
    let ib = count * 2;
    let raw = pb * 2 + ib;
    if bl != (raw + 3) / 4 * 4 || bytes[start + raw..].iter().any(|b| *b != 0) {
        return Err(crate::coded(
            "internal-mismatch",
            "3D buffer size/padding mismatch",
        ));
    }
    let mut positions = Vec::new();
    let mut normals = Vec::new();
    for f in &faces {
        let mut p = f.map(|i| vertices[i]);
        let v = [p[1][0] - p[0][0], p[1][1] - p[0][1], p[1][2] - p[0][2]];
        let w = [p[2][0] - p[0][0], p[2][1] - p[0][1], p[2][2] - p[0][2]];
        let mut normal = [
            v[1] * w[2] - v[2] * w[1],
            v[2] * w[0] - v[0] * w[2],
            v[0] * w[1] - v[1] * w[0],
        ];
        let length = (normal.iter().map(|v| v * v).sum::<f64>()).sqrt();
        normal = normal.map(|n| n / length);
        let center = [0, 1, 2].map(|i| p.iter().map(|v| v[i]).sum::<f64>() / 3.);
        if normal.iter().zip(center).map(|(a, b)| a * b).sum::<f64>() < 0. {
            p.swap(1, 2);
            normal = normal.map(|n| -n);
        }
        positions.extend(p);
        normals.extend([normal; 3]);
    }
    for (base, vectors) in [(0, &positions), (pb, &normals)] {
        for (i, v) in vectors.iter().enumerate() {
            for (k, n) in v.iter().enumerate() {
                let at = start + base + i * 12 + k * 4;
                let actual = f32::from_le_bytes(bytes[at..at + 4].try_into().map_err(err)?);
                if !actual.is_finite() || actual != *n as f32 {
                    return Err(crate::coded(
                        "object-stale",
                        "Foreign/changed 3D geometry is protected",
                    ));
                }
            }
        }
    }
    for i in 0..count {
        let at = start + pb * 2 + i * 2;
        if u16::from_le_bytes(bytes[at..at + 2].try_into().map_err(err)?) as usize != i {
            return Err(crate::coded(
                "object-stale",
                "Changed 3D indices are protected",
            ));
        }
    }
    let min = [0, 1, 2].map(|i| vertices.iter().map(|p| p[i]).fold(f64::INFINITY, f64::min));
    let max = [0, 1, 2].map(|i| {
        vertices
            .iter()
            .map(|p| p[i])
            .fold(f64::NEG_INFINITY, f64::max)
    });
    let mut rgba = Vec::new();
    for i in [0, 2, 4] {
        let s = u8::from_str_radix(&color[i..i + 2], 16).map_err(err)? as f64 / 255.;
        rgba.push(if s <= 0.04045 {
            s / 12.92
        } else {
            ((s + 0.055) / 1.055).powf(2.4)
        });
    }
    rgba.push(1.);
    let expected = json!({"asset":{"version":"2.0","generator":workbook_format!(CLOSED_MODEL3D_SAMPLES_V1)},"scene":0,"scenes":[{"nodes":[0]}],"nodes":[{"mesh":0,"rotation":rotation}],"meshes":[{"name":label,"primitives":[{"attributes":{"POSITION":0,"NORMAL":1},"indices":2,"material":0,"mode":4}]}],"materials":[{"pbrMetallicRoughness":{"baseColorFactor":rgba,"metallicFactor":0,"roughnessFactor":1},"alphaMode":"OPAQUE","doubleSided":false}],"buffers":[{"byteLength":raw}],"bufferViews":[{"buffer":0,"byteOffset":0,"byteLength":pb,"target":34962},{"buffer":0,"byteOffset":pb,"byteLength":pb,"target":34962},{"buffer":0,"byteOffset":pb*2,"byteLength":ib,"target":34963}],"accessors":[{"bufferView":0,"componentType":5126,"count":count,"type":"VEC3","min":min,"max":max},{"bufferView":1,"componentType":5126,"count":count,"type":"VEC3"},{"bufferView":2,"componentType":5123,"count":count,"type":"SCALAR","min":[0],"max":[count-1]}],"extras":{workbook_format!(MODEL3D):{"sample":field(model,"sample")?,"color":color,"rotation":{"x":model["rotation"]["x"],"y":model["rotation"]["y"],"z":model["rotation"]["z"]}}}});
    if !matching_json(&json, &expected, "") {
        return Err(crate::coded(
            "object-protected",
            "Foreign/extended/textured 3D model is protected",
        ));
    }
    Ok(expected["extras"][workbook_format!(MODEL3D)].clone())
}
fn asset(
    zip: &mut ZipArchive<crate::cancel::Reader<File>>,
    part: &str,
    additions: &BTreeMap<String, Vec<u8>>,
) -> Result<Vec<u8>> {
    if let Some(b) = additions.get(part) {
        return Ok(b.clone());
    }
    let mut member = zip.by_name(part).map_err(err)?;
    if member.size() > LIMIT as u64 || member.compressed_size() > (LIMIT + 65536) as u64 {
        return Err(crate::coded(
            "object-limit",
            "3D source asset exceeds its bound",
        ));
    }
    let mut bytes = Vec::new();
    (&mut member)
        .take((LIMIT + 1) as u64)
        .read_to_end(&mut bytes)
        .map_err(err)?;
    if bytes.len() > LIMIT {
        return Err(crate::coded(
            "object-limit",
            "3D source asset inflated outside its bound",
        ));
    }
    Ok(bytes)
}
struct Qualified {
    number: u64,
    part: String,
    revision: String,
}
fn qualified(
    zip: &mut ZipArchive<crate::cancel::Reader<File>>,
    names: &BTreeSet<String>,
    replacements: &serde_json::Map<String, Value>,
    relations: &BTreeMap<String, BTreeMap<String, Edge>>,
    additions: &BTreeMap<String, Vec<u8>>,
    drawing: &str,
    id: &str,
) -> Result<Qualified> {
    let bytes = part_bytes(zip, drawing, replacements)?;
    let t = Tree::parse(&bytes)?;
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
        return Err(crate::coded("object-ambiguous", "Ambiguous 3D identity"));
    }
    let prop = props[0];
    let private = PRIVATE;
    let at = t.nodes[0]
        .children
        .iter()
        .copied()
        .find(|i| t.nodes[*i].start <= t.nodes[prop].start && t.nodes[*i].end >= t.nodes[prop].end)
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing 3D anchor"))?;
    let from = child(&t, at, DRAW_NS, "from")?;
    let col = scalar(&t, from, "col", 16383)?;
    let row = scalar(&t, from, "row", 1048575)?;
    let dx = scalar(&t, from, "colOff", 10000000)?;
    let dy = scalar(&t, from, "rowOff", 10000000)?;
    let ext = child(&t, at, DRAW_NS, "ext")?;
    let w = attribute(&t, ext, "cx", 1, 50000000)?;
    let h = attribute(&t, ext, "cy", 1, 50000000)?;
    let number = attribute(&t, prop, "id", 1, 2147483647)?;
    let alt = t.nodes[prop]
        .attrs
        .get("descr")
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing 3D alternative text"))?;
    if alt.encode_utf16().count() > 256
        || alt.chars().any(|c| {
            (c < '\u{20}' && !['\t', '\n', '\r'].contains(&c))
                || ['\u{fffe}', '\u{ffff}'].contains(&c)
        })
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid 3D alternative text",
        ));
    }
    let rid = format!(workbook_format!(MODEL3D_D80344), id.replace('-', ""));
    let edge = relations
        .get(drawing)
        .and_then(|r| r.get(&rid))
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing 3D edge"))?;
    let part = edge
        .target
        .as_ref()
        .ok_or_else(|| crate::coded("object-protected", "External 3D model is protected"))?;
    let prefix = format!(workbook_format!(XL_MEDIA_MODEL3D_ID), id = id);
    let digest = part
        .strip_prefix(&prefix)
        .and_then(|p| p.strip_suffix(".glb"))
        .filter(|s| valid_hash(s))
        .ok_or_else(|| crate::coded("object-protected", "Foreign 3D asset is protected"))?;
    if edge.kind != REL
        || !names.contains(part)
        || names.contains(&format!(
            "xl/media/_rels/{}.rels",
            part.rsplit('/')
                .next()
                .ok_or_else(|| crate::coded("object-malformed", "Invalid 3D asset path"))?
        ))
        || relations
            .values()
            .flat_map(|m| m.values())
            .filter(|e| e.target.as_ref() == Some(part))
            .count()
            != 1
    {
        return Err(crate::coded(
            "object-protected",
            "Shared/dependent 3D asset is protected",
        ));
    }
    let data = asset(zip, part, additions)?;
    if format!("{:x}", Sha256::digest(&data)) != digest {
        return Err(crate::coded("internal-mismatch", "3D source hash mismatch"));
    }
    let model = glb(&data)?;
    let sample = field(&model, "sample")?;
    let color = field(&model, "color")?;
    let (x, y, z) = (
        &model["rotation"]["x"],
        &model["rotation"]["y"],
        &model["rotation"]["z"],
    );
    let d = String::from_utf8_lossy(DRAW_NS);
    let a = String::from_utf8_lossy(A_NS);
    let r = String::from_utf8_lossy(REL_NS);
    let (ortho_n, ortho_d) = (w.min(h), 5 * h);
    let expected = format!(workbook_format!(XML_4C3CF49590),
        quote(alt),
        d = d,
        a = a,
        r = r,
        M = M,
        PRIVATE = private,
        col = col,
        dx = dx,
        row = row,
        dy = dy,
        w = w,
        h = h,
        number = number,
        id = id,
        rid = rid,
        ortho_n = ortho_n,
        ortho_d = ortho_d,
        sample = sample,
        color = color,
        x = x,
        y = y,
        z = z
    );
    if !same(&t, at, &Tree::parse(expected.as_bytes())?, 0) {
        return Err(crate::coded(
            "object-protected",
            "Foreign/extended 3D anchor is protected",
        ));
    }
    let file = drawing
        .rsplit('/')
        .next()
        .ok_or_else(|| crate::coded("object-malformed", "Invalid 3D drawing path"))?;
    let links = Tree::parse(&part_bytes(
        zip,
        &format!("xl/drawings/_rels/{file}.rels"),
        replacements,
    )?)?;
    let matches = links
        .nodes
        .iter()
        .enumerate()
        .filter(|(_, n)| {
            n.name == "Relationship" && n.ns == PACKAGE_REL_NS && n.attrs.get("Id") == Some(&rid)
        })
        .map(|(i, _)| i)
        .collect::<Vec<_>>();
    let target = format!(
        "../media/{}",
        part.rsplit('/')
            .next()
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid 3D asset name"))?
    );
    let expected_edge = format!(
        "<Relationship xmlns=\"{}\" Id=\"{rid}\" Type=\"{REL}\" Target=\"{target}\"/>",
        String::from_utf8_lossy(PACKAGE_REL_NS)
    );
    if matches.len() != 1
        || !same(
            &links,
            matches[0],
            &Tree::parse(expected_edge.as_bytes())?,
            0,
        )
    {
        return Err(crate::coded(
            "object-protected",
            "Extended 3D edge is protected",
        ));
    }
    let types = Tree::parse(&part_bytes(zip, "[Content_Types].xml", replacements)?)?;
    let ct = types
        .nodes
        .iter()
        .filter(|n| {
            n.ns == TYPE_NS
                && n.name == "Override"
                && n.attrs.get("PartName") == Some(&format!("/{part}"))
        })
        .collect::<Vec<_>>();
    if ct.len() != 1
        || ct[0].attrs.get("ContentType").map(String::as_str) != Some(MIME)
        || ct[0]
            .attrs
            .keys()
            .any(|k| !k.starts_with("xmlns") && !["PartName", "ContentType"].contains(&k.as_str()))
    {
        return Err(crate::coded(
            "internal-mismatch",
            "3D content type mismatch",
        ));
    }
    let mut hash = Sha256::new();
    hash.update(&bytes[t.nodes[at].start..t.nodes[at].end]);
    for value in [part, &rid, REL, &target] {
        hash.update([0]);
        hash.update(value.as_bytes());
    }
    Ok(Qualified {
        number,
        part: part.clone(),
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
    additions: &BTreeMap<String, Vec<u8>>,
    plan: &Value,
) -> Result<()> {
    let id = field(plan, "objectId")?;
    let drawing = field(plan, "drawingPart")?;
    let sheet = field(plan, "sheet")?;
    let kind = field(plan, "kind")?;
    let create = kind == "createModel3D";
    let delete = kind == "deleteModel3D";
    let empty = serde_json::Map::new();
    let none = BTreeMap::new();
    let old = if !create {
        Some(qualified(
            zip,
            original,
            &empty,
            before_relations,
            &none,
            drawing,
            id,
        )?)
    } else {
        None
    };
    let next = if !delete {
        Some(qualified(
            zip,
            names,
            replacements,
            relations,
            additions,
            drawing,
            id,
        )?)
    } else {
        None
    };
    if old
        .as_ref()
        .is_some_and(|o| plan["priorObjectRevision"].as_str() != Some(&o.revision))
    {
        return Err(crate::coded(
            "object-stale",
            "Conflict: 3D revision changed",
        ));
    }
    if let (Some(a), Some(b)) = (&old, &next) {
        if a.number != b.number {
            return Err(crate::coded(
                "object-stale",
                "3D nonvisual identity changed",
            ));
        }
    }
    let scoped = next
        .as_ref()
        .or(old.as_ref())
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing 3D object"))?;
    if field(plan, "objectPart")? != scoped.part {
        return Err(crate::coded("internal-mismatch", "3D asset scope mismatch"));
    }
    let oldpart = old.as_ref().map(|o| o.part.as_str());
    let newpart = next.as_ref().map(|o| o.part.as_str());
    if oldpart
        .into_iter()
        .chain(newpart)
        .any(|p| replacements.contains_key(p))
    {
        return Err(crate::coded(
            "object-unsupported",
            "3D assets cannot be replaced as XML",
        ));
    }
    let expected_add = if newpart != oldpart { newpart } else { None };
    if additions.len() != usize::from(expected_add.is_some())
        || expected_add.is_some_and(|p| !additions.contains_key(p))
    {
        return Err(crate::coded(
            "object-bad-input",
            "3D addition must match its scoped asset transition",
        ));
    }
    let removals = plan["removed"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing 3D removals"))?;
    let mut expected_removal = BTreeSet::new();
    if oldpart != newpart {
        if let Some(p) = oldpart {
            expected_removal.insert(p.to_string());
        }
    }
    if delete && !names.contains(drawing) {
        expected_removal.insert(drawing.to_string());
        let rel = format!(
            "xl/drawings/_rels/{}.rels",
            drawing
                .rsplit('/')
                .next()
                .ok_or_else(|| crate::coded("object-malformed", "Invalid drawing path"))?
        );
        if original.contains(&rel) {
            expected_removal.insert(rel);
        }
    }
    if removals
        .iter()
        .map(|v| {
            v.as_str()
                .map(str::to_string)
                .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid 3D removal"))
        })
        .collect::<Result<BTreeSet<_>>>()?
        != expected_removal
    {
        return Err(crate::coded(
            "object-unsupported",
            "3D removal claims unrelated parts",
        ));
    }
    let incoming = |r: &BTreeMap<String, BTreeMap<String, Edge>>| {
        r.values()
            .flat_map(|m| m.values())
            .filter(|e| e.target.as_deref() == Some(drawing))
            .count()
    };
    if (!create && incoming(before_relations) != 1) || (!delete && incoming(relations) != 1) {
        return Err(crate::coded(
            "object-protected",
            "Shared 3D drawing is protected",
        ));
    }
    for owner in before_relations
        .keys()
        .chain(relations.keys())
        .collect::<BTreeSet<_>>()
    {
        let keep = |edges: Option<&BTreeMap<String, Edge>>| {
            edges
                .into_iter()
                .flat_map(|e| e.iter())
                .filter(|(_, e)| {
                    !(owner == sheet && e.target.as_deref() == Some(drawing))
                        && !(owner == drawing
                            && e.target
                                .as_deref()
                                .is_some_and(|p| Some(p) == oldpart || Some(p) == newpart))
                })
                .map(|(k, e)| (k.clone(), e.clone()))
                .collect::<BTreeMap<_, _>>()
        };
        if keep(before_relations.get(owner)) != keep(relations.get(owner)) {
            return Err(crate::coded(
                "object-stale",
                "Unrelated 3D relationship changed",
            ));
        }
    }
    let mut types =
        |replacements: &serde_json::Map<String, Value>| -> Result<BTreeMap<String, String>> {
            let t = Tree::parse(&part_bytes(zip, "[Content_Types].xml", replacements)?)?;
            Ok(t.nodes
                .iter()
                .filter(|n| n.ns == TYPE_NS && ["Default", "Override"].contains(&n.name.as_str()))
                .filter(|n| {
                    !n.attrs.get("PartName").is_some_and(|p| {
                        p == &format!("/{drawing}")
                            || oldpart.is_some_and(|s| p == &format!("/{s}"))
                            || newpart.is_some_and(|s| p == &format!("/{s}"))
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
            "Unrelated 3D content type changed",
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
                    e.target.as_deref() != Some(drawing)
                        || e.kind != format!("{}/drawing", String::from_utf8_lossy(REL_NS))
                }))
    {
        return Err(crate::coded(
            "internal-mismatch",
            "3D sheet/drawing scope mismatch",
        ));
    }
    if delete && !names.contains(drawing) && !ids.is_empty() {
        return Err(crate::coded(
            "object-unsupported",
            "Deleted 3D drawing retains worksheet link",
        ));
    }
    Ok(())
}
#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn retained_samples_are_real_independently_admitted_glb() {
        for bytes in [
            include_bytes!("../../../core/assets/models/cube.glb").as_slice(),
            include_bytes!("../../../core/assets/models/pyramid.glb").as_slice(),
            include_bytes!("../../../core/assets/models/octahedron.glb").as_slice(),
        ] {
            assert!(glb(bytes).is_ok());
            let mut changed = bytes.to_vec();
            *changed.last_mut().unwrap() = 255;
            assert!(glb(&changed).is_err());
            assert!(glb(&bytes[..bytes.len() - 1]).is_err());
        }
    }
}
