//! Read-only print resolution of the closed owned theme catalog. Neither a
//! caller projection nor a matching font name qualifies an imported theme.
use super::{print_settings::Tree, theme_authoring, *};
const DRAWING: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
#[derive(Debug)]
pub(super) struct PrintTheme {
    pub sample: String,
    pub revision: String,
    pub major: String,
    pub minor: String,
    pub colors: Vec<String>,
}
fn direct(tree: &Tree, parent: usize, name: &str) -> Result<usize> {
    let nodes = tree.nodes[parent]
        .children
        .iter()
        .filter(|i| tree.nodes[**i].name == name)
        .copied()
        .collect::<Vec<_>>();
    if nodes.len() != 1 || tree.nodes[nodes[0]].ns != DRAWING {
        return Err(crate::coded(
            "print-refused",
            "PRINT_UNSUPPORTED_THEME: ambiguous theme element",
        ));
    }
    Ok(nodes[0])
}
pub(super) fn owned(
    zip: &mut ZipArchive<crate::cancel::Reader<File>>,
) -> Result<Option<PrintTheme>> {
    let empty = serde_json::Map::new();
    let rels = Tree::parse(&part_bytes(zip, "xl/_rels/workbook.xml.rels", &empty)?)?;
    let edges = theme_authoring::relationships(&rels)?;
    let themes = edges
        .iter()
        .filter(|i| rels.nodes[**i].attrs["Type"] == theme_authoring::THEME_EDGE)
        .copied()
        .collect::<Vec<_>>();
    if themes.is_empty() {
        return Ok(None);
    }
    if themes.len() != 1 {
        return Err(crate::coded(
            "print-refused",
            "PRINT_UNSUPPORTED_THEME: ambiguous relationship",
        ));
    }
    let edge = &rels.nodes[themes[0]];
    let theme_part = theme_authoring::target("xl/workbook.xml", &edge.attrs["Target"])?;
    if edge
        .attrs
        .get("TargetMode")
        .is_some_and(|mode| mode != "Internal")
        || !format_matches!(theme_part.as_str(), XL_THEME_PAGE_THEME_XML)
    {
        return Ok(None);
    }
    let size = zip.by_name(&theme_part).map_err(err)?.size();
    if size > 65536 {
        return Err(crate::coded(
            "print-limit",
            "PRINT_UNSUPPORTED_THEME: theme exceeds 64 KiB",
        ));
    }
    let bytes = part_bytes(zip, &theme_part, &empty)?;
    let Some(sample) = ["neutral", "forest", "plum"].into_iter().find(|id| {
        theme_authoring::sample(id)
            .is_ok_and(|xml| crate::workbook_format::page_theme_matches(&bytes, xml))
    }) else {
        return Ok(None);
    };
    let tree = Tree::parse(&bytes)?;
    tree.root(DRAWING, "theme")?;
    let elements = direct(&tree, 0, "themeElements")?;
    let scheme = direct(&tree, elements, "clrScheme")?;
    let fonts = direct(&tree, elements, "fontScheme")?;
    let font = |role: &str| -> Result<String> {
        let parent = direct(&tree, fonts, role)?;
        let latin = direct(&tree, parent, "latin")?;
        let name = tree.nodes[latin]
            .attrs
            .get("typeface")
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing theme Latin font"))?;
        if name.trim().is_empty()
            || name.encode_utf16().count() > 128
            || name.chars().any(char::is_control)
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid theme Latin font",
            ));
        }
        Ok(name.clone())
    };
    let colors = [
        "lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5",
        "accent6", "hlink", "folHlink",
    ]
    .into_iter()
    .map(|role| -> Result<String> {
        let slot = direct(&tree, scheme, role)?;
        let children = &tree.nodes[slot].children;
        if children.len() != 1 {
            return Err(crate::coded("print-ambiguous", "Ambiguous theme color"));
        }
        let color = &tree.nodes[children[0]];
        if color.ns != DRAWING {
            return Err(crate::coded("print-unsupported", "Foreign theme color"));
        }
        let value = match color.name.as_str() {
            "srgbClr" => color.attrs.get("val"),
            "sysClr" => color.attrs.get("lastClr"),
            _ => None,
        }
        .ok_or_else(|| crate::coded("print-unsupported", "Unresolved theme color"))?;
        if value.len() != 6 || !value.bytes().all(|c| c.is_ascii_hexdigit()) {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid theme color",
            ));
        }
        Ok(value.to_ascii_uppercase())
    })
    .collect::<Result<Vec<_>>>()?;
    Ok(Some(PrintTheme {
        sample: sample.into(),
        revision: format!("{:x}", Sha256::digest(&bytes)),
        major: font("majorFont")?,
        minor: font("minorFont")?,
        colors,
    }))
}
pub(super) fn color(
    tree: &Tree,
    index: usize,
    theme: Option<&PrintTheme>,
    black_only: bool,
) -> Result<()> {
    tree.clean(index, &["theme", "tint"])?;
    let node = &tree.nodes[index];
    let theme = theme.ok_or_else(|| {
        crate::coded(
            "print-unsupported",
            "PRINT_UNSUPPORTED_THEME: imported theme color",
        )
    })?;
    if let Some(tint) = node.attrs.get("tint") {
        let value = tint.parse::<f64>().map_err(err)?;
        if !value.is_finite() || value != 0.0 {
            return Err(crate::coded(
                "print-refused",
                "PRINT_UNSUPPORTED_THEME: tinted color",
            ));
        }
    }
    let slot = node
        .attrs
        .get("theme")
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing theme color index"))?
        .parse::<usize>()
        .map_err(err)?;
    let color = theme
        .colors
        .get(slot)
        .ok_or_else(|| crate::coded("print-unsupported", "Unresolved theme color index"))?;
    if black_only && color != "000000" {
        return Err(crate::coded(
            "print-unsupported",
            "Colored print border requires qualification",
        ));
    }
    Ok(())
}
pub(super) fn scheme(tree: &Tree, index: usize, theme: Option<&PrintTheme>) -> Result<()> {
    tree.clean(index, &["val"])?;
    match tree.nodes[index].attrs.get("val").map(String::as_str) {
        Some("none") => Ok(()),
        Some("major" | "minor") if theme.is_some() => Ok(()),
        _ => Err(crate::coded(
            "print-unsupported",
            "PRINT_UNSUPPORTED_THEME: unresolved font scheme",
        )),
    }
}
/// Independent source Normal-font projection for owned-theme geometry. An
/// imported Normal font scheme remains unsupported even with literal cell XFs.
pub(super) fn normal_font(tree: &Tree, theme: Option<&PrintTheme>) -> Result<Option<Value>> {
    let fonts = tree
        .direct(0, "fonts")?
        .map(|i| tree.nodes[i].children.clone())
        .unwrap_or_default();
    let mut base = 0usize;
    if let Some(styles) = tree.direct(0, "cellStyles")? {
        let normal = tree.nodes[styles]
            .children
            .iter()
            .filter(|i| {
                let n = &tree.nodes[**i];
                n.attrs.get("builtinId").is_some_and(|v| v == "0")
                    || n.attrs.get("name").is_some_and(|v| v == "Normal")
            })
            .copied()
            .collect::<Vec<_>>();
        if normal.len() > 1 {
            return Err(crate::coded(
                "print-refused",
                "PRINT_UNRESOLVED_GEOMETRY: ambiguous Normal style",
            ));
        }
        if let Some(i) = normal.first() {
            if tree.nodes[*i].ns != NS || tree.nodes[*i].name != "cellStyle" {
                return Err(crate::coded(
                    "print-unsupported",
                    "PRINT_UNRESOLVED_GEOMETRY: foreign Normal style",
                ));
            }
            base = tree.nodes[*i]
                .attrs
                .get("xfId")
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing Normal style base"))?
                .parse::<usize>()
                .map_err(err)?;
        }
    }
    let font_id = if let Some(bases) = tree.direct(0, "cellStyleXfs")? {
        let xf = *tree.nodes[bases]
            .children
            .get(base)
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing Normal style base"))?;
        if tree.nodes[xf].ns != NS || tree.nodes[xf].name != "xf" {
            return Err(crate::coded(
                "print-unsupported",
                "PRINT_UNRESOLVED_GEOMETRY: foreign Normal style base",
            ));
        }
        tree.nodes[xf]
            .attrs
            .get("fontId")
            .map(|v| v.parse::<usize>().map_err(err))
            .transpose()?
            .unwrap_or(0)
    } else {
        0
    };
    let Some(font) = fonts.get(font_id).copied() else {
        if theme.is_none() {
            return Ok(None);
        }
        return Err(crate::coded(
            "print-refused",
            "PRINT_UNRESOLVED_GEOMETRY: missing Normal font",
        ));
    };
    if tree.nodes[font].ns != NS || tree.nodes[font].name != "font" {
        return Err(crate::coded(
            "print-unsupported",
            "PRINT_UNRESOLVED_GEOMETRY: foreign Normal font",
        ));
    }
    // A theme role cannot hide ambiguous literal source-font declarations.
    let literal_name = tree.direct(font, "name")?;
    let scheme = tree.direct(font, "scheme")?;
    if theme.is_none()
        && scheme.is_none_or(|i| tree.nodes[i].attrs.get("val").is_some_and(|v| v == "none"))
    {
        return Ok(None);
    }
    if let Some(i) = scheme {
        self::scheme(tree, i, theme)?;
    }
    let name = match scheme.and_then(|i| tree.nodes[i].attrs.get("val").map(String::as_str)) {
        Some("major") => theme
            .ok_or_else(|| crate::coded("print-unsupported", "Unresolved Normal major font"))?
            .major
            .clone(),
        Some("minor") => theme
            .ok_or_else(|| crate::coded("print-unsupported", "Unresolved Normal minor font"))?
            .minor
            .clone(),
        _ => literal_name
            .and_then(|i| tree.nodes[i].attrs.get("val").cloned())
            .unwrap_or_else(|| "Calibri".into()),
    };
    if name.trim().is_empty() || name.len() > 256 || name.chars().any(char::is_control) {
        return Err(crate::coded(
            "print-refused",
            "PRINT_UNRESOLVED_GEOMETRY: invalid Normal font",
        ));
    }
    let size = tree
        .direct(font, "sz")?
        .and_then(|i| tree.nodes[i].attrs.get("val"))
        .map(|s| s.parse::<f64>().map_err(err))
        .transpose()?
        .unwrap_or(11.0);
    if !size.is_finite() || size <= 0.0 || size > 409.5 {
        return Err(crate::coded(
            "print-refused",
            "PRINT_UNRESOLVED_GEOMETRY: invalid Normal size",
        ));
    }
    let flag = |name: &str| -> Result<bool> {
        let Some(i) = tree.direct(font, name)? else {
            return Ok(false);
        };
        tree.clean(i, &["val"])?;
        match tree.nodes[i].attrs.get("val").map(String::as_str) {
            None | Some("1" | "true") => Ok(true),
            Some("0" | "false") => Ok(false),
            _ => Err(crate::coded(
                "print-refused",
                "PRINT_UNRESOLVED_GEOMETRY: invalid Normal font flag",
            )),
        }
    };
    Ok(Some(
        json!({"fontFamily":name,"fontSize":size,"bold":flag("b")?,"italic":flag("i")?}),
    ))
}

#[cfg(test)]
mod tests {
    use super::super::theme_authoring::tests::{fixture, mutate, read_part};
    use super::*;
    fn attach(root: &Path, id: &str, target: &str, bytes: &[u8]) {
        let rels = String::from_utf8(read_part(root, "source.xlsx", "xl/_rels/workbook.xml.rels"))
            .unwrap();
        let changed = rels.replace(
            "</Relationships>",
            &format!(
                "<Relationship Id=\"theme\" Type=\"{}\" Target=\"{}\"/></Relationships>",
                theme_authoring::THEME_EDGE,
                target
            ),
        );
        mutate(
            root,
            "source.xlsx",
            "xl/_rels/workbook.xml.rels",
            changed.as_bytes(),
        );
        mutate(root, "source.xlsx", id, bytes);
    }
    fn theme(root: &Path) -> Result<Option<PrintTheme>> {
        let mut zip = ZipArchive::new(crate::cancel::Reader(
            File::open(root.join("source.xlsx")).unwrap(),
        ))
        .unwrap();
        owned(&mut zip)
    }
    fn styles() -> Vec<u8> {
        format!("<styleSheet xmlns=\"{}\"><fonts><font><name val=\"Calibri\"/><sz val=\"11\"/><scheme val=\"minor\"/><color theme=\"4\"/></font><font><name val=\"Courier New\"/><sz val=\"12\"/><color rgb=\"FF119922\"/></font></fonts><fills><fill><patternFill patternType=\"none\"/></fill></fills><borders><border/></borders><cellStyleXfs><xf fontId=\"0\"/></cellStyleXfs><cellXfs><xf fontId=\"0\" fillId=\"0\" borderId=\"0\" numFmtId=\"0\"/><xf fontId=\"1\" fillId=\"0\" borderId=\"0\" numFmtId=\"0\"/></cellXfs></styleSheet>",String::from_utf8_lossy(NS)).into_bytes()
    }
    #[test]
    fn canonical_samples_qualify_fonts_colors_and_source_normal_geometry() {
        let _scope = crate::shared::begin();
        for (id, minor, major, color) in [
            ("neutral", "Arial", "Arial", "3366CC"),
            ("forest", "Verdana", "Georgia", "28744B"),
            ("plum", "Georgia", "Verdana", "86509C"),
        ] {
            let (_owner, root) = fixture();
            attach(
                &root,
                theme_authoring::THEME_PART,
                workbook_format!(THEME_PAGE_THEME_XML),
                theme_authoring::sample(id).unwrap().as_bytes(),
            );
            let resolved = theme(&root).unwrap().unwrap();
            assert_eq!(resolved.sample, id);
            assert_eq!(resolved.minor, minor);
            assert_eq!(resolved.major, major);
            assert_eq!(resolved.colors[4], color);
            assert_eq!(resolved.colors[0], "FFFFFF");
            assert_eq!(
                resolved.colors[1],
                match id {
                    "neutral" => "17212B",
                    "forest" => "182A20",
                    _ => "2D2034",
                }
            );
            let bytes = styles();
            assert_eq!(
                super::super::print_settings::qualified_styles(
                    Some(bytes.clone()),
                    Some(&resolved)
                )
                .unwrap(),
                vec![true, true]
            );
            assert_eq!(
                normal_font(&Tree::parse(&bytes).unwrap(), Some(&resolved))
                    .unwrap()
                    .unwrap(),
                json!({"fontFamily":minor,"fontSize":11.0,"bold":false,"italic":false})
            );
        }
    }
    #[test]
    fn imported_modified_external_and_overbound_themes_do_not_grant_print_roles() {
        let _scope = crate::shared::begin();
        for target in ["theme/imported.xml", workbook_format!(THEME_PAGE_THEME_XML)] {
            let (_owner, root) = fixture();
            let xml = if target.ends_with("imported.xml") {
                theme_authoring::sample("forest").unwrap().to_owned()
            } else {
                theme_authoring::sample("forest").unwrap().to_owned() + "\n"
            };
            attach(&root, &format!("xl/{target}"), target, xml.as_bytes());
            assert!(theme(&root).unwrap().is_none());
            assert!(normal_font(&Tree::parse(&styles()).unwrap(), None).is_err());
            assert_eq!(
                super::super::print_settings::qualified_styles(Some(styles()), None).unwrap(),
                vec![false, true]
            );
        }
        let (_owner, root) = fixture();
        attach(
            &root,
            theme_authoring::THEME_PART,
            workbook_format!(THEME_PAGE_THEME_XML),
            &vec![b'x'; 65537],
        );
        assert!(theme(&root).unwrap_err().contains("64 KiB"));
        let (_owner, root) = fixture();
        attach(
            &root,
            theme_authoring::THEME_PART,
            workbook_format!(THEME_PAGE_THEME_XML),
            theme_authoring::sample("neutral").unwrap().as_bytes(),
        );
        let rels = String::from_utf8(read_part(
            &root,
            "source.xlsx",
            "xl/_rels/workbook.xml.rels",
        ))
        .unwrap();
        mutate(
            &root,
            "source.xlsx",
            "xl/_rels/workbook.xml.rels",
            rels.replace("Id=\"theme\"", "Id=\"theme\" TargetMode=\"External\"")
                .as_bytes(),
        );
        assert!(theme(&root).unwrap().is_none());
    }
    #[test]
    fn tinted_ambiguous_colors_and_unknown_scheme_normal_font_refuse() {
        let _scope = crate::shared::begin();
        let (_owner, root) = fixture();
        attach(
            &root,
            theme_authoring::THEME_PART,
            workbook_format!(THEME_PAGE_THEME_XML),
            theme_authoring::sample("forest").unwrap().as_bytes(),
        );
        let resolved = theme(&root).unwrap().unwrap();
        for attrs in [
            "theme=\"4\" tint=\"0.2\"",
            "theme=\"12\"",
            "theme=\"4\" rgb=\"FF000000\"",
        ] {
            let xml = format!(
                "<styleSheet xmlns=\"{}\"><color {attrs}/></styleSheet>",
                String::from_utf8_lossy(NS)
            );
            assert!(
                color(
                    &Tree::parse(xml.as_bytes()).unwrap(),
                    1,
                    Some(&resolved),
                    false
                )
                .is_err()
            );
        }
        let bytes = String::from_utf8(styles()).unwrap();
        for candidate in [
            bytes.replace("val=\"minor\"", "val=\"unknown\""),
            bytes.replace(
                "<name val=\"Calibri\"/>",
                "<name val=\"Calibri\"/><name val=\"Arial\"/>",
            ),
        ] {
            assert!(
                normal_font(&Tree::parse(candidate.as_bytes()).unwrap(), Some(&resolved)).is_err()
            );
        }
    }
}
