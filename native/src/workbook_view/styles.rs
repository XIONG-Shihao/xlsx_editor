//! The workbook theme and the cell styles that follow from it: for each style
//! number, the format, font, fill, border and alignment a cell with it shows.
use super::{dom::Dom, package::Package};
use crate::{Result, coded};
use serde_json::{Map, Value, json};
use std::collections::HashMap;

const DRAWING: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
// Spreadsheet indices differ from the XML sequence (dk1, lt1, dk2, lt2).
const THEME_COLORS: [&str; 12] = [
    "lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
];

/// Problems found while resolving colors and fonts, each reported once, in the order met.
#[derive(Default)]
pub(super) struct Issues(pub Vec<String>);
impl Issues {
    pub fn add(&mut self, text: &str) {
        if !self.0.iter().any(|t| t == text) {
            self.0.push(text.to_string());
        }
    }
}

pub(super) struct Theme {
    pub part: String,
    pub colors: HashMap<&'static str, String>,
    pub fonts: HashMap<&'static str, String>,
}

/// JavaScript's `Number(text)` for the plain decimal forms workbooks use.
pub(super) fn js_number(text: &str) -> f64 {
    let t = text.trim();
    if t.is_empty() {
        return 0.0;
    }
    if !t.bytes().all(|b| b.is_ascii_digit() || b"+-.eE".contains(&b)) {
        return f64::NAN;
    }
    t.parse().unwrap_or(f64::NAN)
}
pub(super) fn number_value(n: f64) -> Value {
    if n.is_finite() && n.fract() == 0.0 && n.abs() < 9.0e15 && !(n == 0.0 && n.is_sign_negative()) {
        json!(n as i64)
    } else {
        json!(n)
    }
}

fn rgb(value: &str) -> Option<String> {
    (value.len() == 6 && value.bytes().all(|b| b.is_ascii_hexdigit())).then(|| value.to_ascii_uppercase())
}
fn family(value: &str) -> Option<String> {
    (!value.trim().is_empty() && value.encode_utf16().count() <= 128 && !value.chars().any(|c| c < ' ' || c == '\u{7f}'))
        .then(|| value.to_string())
}

fn single(dom: &Dom, of: usize, name: &str) -> Result<usize> {
    let found: Vec<usize> = dom
        .child_elements(of)
        .into_iter()
        .filter(|i| dom.nodes[*i].ns == DRAWING && dom.nodes[*i].local == name)
        .collect();
    if found.len() != 1 {
        return Err(coded("style-refused", format!("Theme has a missing or ambiguous {name}")));
    }
    Ok(found[0])
}

fn read_theme(package: &mut Package, theme_paths: &[Option<String>]) -> Result<(Theme, Vec<String>)> {
    let Some(Some(part)) = (theme_paths.len() == 1).then(|| theme_paths[0].clone()) else {
        return Err(coded("style-refused", "Theme relationship is external or ambiguous"));
    };
    let legal = !part.is_empty()
        && part.ends_with(".xml")
        && part[..part.len() - 4]
            .bytes()
            .all(|b| b.is_ascii_alphanumeric() || b"_-/".contains(&b))
        && !part.contains("..")
        && !part.contains("//");
    if !legal {
        return Err(coded("style-unsupported", "Theme part path is unsupported"));
    }
    let Some(size) = package.size(&part) else {
        return Err(coded("internal-missing-field", "Theme part is missing"));
    };
    if size > 64 * 1024 {
        return Err(coded("style-limit", "Theme XML exceeds 64 KiB"));
    }
    let bytes = package.bytes(&part)?;
    let lower = String::from_utf8_lossy(&bytes).to_ascii_lowercase();
    if lower.contains("<!doctype") || lower.contains("<!entity") {
        return Err(coded("style-malformed", "Theme XML entities are unsupported"));
    }
    let dom = Dom::parse(&bytes)?;
    if !dom.root_is(DRAWING, "theme") {
        return Err(coded("style-unsupported", "Theme namespace/root is unsupported"));
    }
    let elements = single(&dom, 0, "themeElements")?;
    let scheme = single(&dom, elements, "clrScheme")?;
    let fonts = single(&dom, elements, "fontScheme")?;
    let mut issues: Vec<String> = Vec::new();
    let mut note = |text: &str| {
        if !issues.iter().any(|t| t == text) {
            issues.push(text.to_string());
        }
    };
    let mut colors = HashMap::new();
    for key in THEME_COLORS {
        let slot = single(&dom, scheme, key)?;
        let values = dom.child_elements(slot);
        let color = (values.len() == 1
            && dom.nodes[values[0]].ns == DRAWING
            && dom.child_elements(values[0]).is_empty())
        .then(|| match dom.nodes[values[0]].local.as_str() {
            "srgbClr" => rgb(dom.attr(values[0], "val")),
            "sysClr" => rgb(dom.attr(values[0], "lastClr")),
            _ => None,
        })
        .flatten();
        match color {
            Some(color) => {
                colors.insert(key, color);
            }
            None => note("A theme color uses an unsupported value or transform; its cell color is unresolved"),
        }
    }
    let mut fonts_found = HashMap::new();
    for (key, tag) in [("major", "majorFont"), ("minor", "minorFont")] {
        let latin = single(&dom, single(&dom, fonts, tag)?, "latin")?;
        match family(dom.attr(latin, "typeface")) {
            Some(font) => {
                fonts_found.insert(key, font);
            }
            None => note("A theme Latin font is unresolved; the explicit cell font is used"),
        }
    }
    Ok((Theme { part, colors, fonts: fonts_found }, issues))
}

/// The theme the workbook names, and what was wrong with it if anything.
pub(super) fn theme(package: &mut Package, theme_paths: &[Option<String>]) -> (Option<Theme>, Vec<String>) {
    if theme_paths.is_empty() {
        return (None, Vec::new());
    }
    match read_theme(package, theme_paths) {
        Ok((theme, issues)) => (Some(theme), issues),
        Err(error) => {
            let said: String = crate::sentence(&error).chars().take(200).collect();
            (None, vec![format!("Workbook theme is unresolved: {said}")])
        }
    }
}

fn theme_color(dom: &Dom, node: Option<usize>, theme: &Option<Theme>, issues: &mut Issues) -> Option<String> {
    let node = node?;
    if let Some(tint) = dom.attribute(node, "tint") {
        if !tint.is_empty() && js_number(tint) != 0.0 {
            issues.add("Cell color tint is not qualified; its display color is unresolved");
            return None;
        }
    }
    if let Some(raw) = dom.attribute(node, "theme") {
        let index = (raw == "0" || (raw.starts_with(|c: char| ('1'..='9').contains(&c)) && raw.bytes().all(|b| b.is_ascii_digit())))
            .then(|| raw.parse::<usize>().ok())
            .flatten();
        let result = index
            .filter(|i| *i < THEME_COLORS.len())
            .and_then(|i| theme.as_ref().and_then(|t| t.colors.get(THEME_COLORS[i]).cloned()));
        if result.is_none() {
            issues.add("Cell theme color is unresolved");
        }
        return result;
    }
    let value = dom.attr(node, "rgb");
    if value.len() == 8 && value.bytes().all(|b| b.is_ascii_hexdigit()) {
        Some(value[2..].to_ascii_uppercase())
    } else {
        rgb(value)
    }
}
fn theme_font(
    explicit: Option<&str>,
    scheme: Option<&str>,
    theme: &Option<Theme>,
    issues: &mut Issues,
) -> Option<String> {
    let explicit = explicit.map(str::to_string);
    match scheme {
        None | Some("") | Some("none") => explicit,
        Some(scheme) => {
            if ["major", "minor"].contains(&scheme) {
                if let Some(font) = theme.as_ref().and_then(|t| t.fonts.get(scheme)) {
                    return Some(font.clone());
                }
            }
            issues.add("Cell theme font is unresolved; its explicit font is used");
            explicit
        }
    }
}

pub(super) struct Styles {
    pub cell_styles: Vec<Value>,
    pub normal_font: Value,
    /// Problems with the normal font that every sheet's geometry inherits.
    pub normal_issues: Vec<String>,
}

fn flag(dom: &Dom, font: usize, name: &str) -> bool {
    dom.elements(font, name)
        .iter()
        .any(|f| !["0", "false"].contains(&dom.attr(*f, "val")))
}

/// Resolve every cell style once. `issues` receives what could not be resolved.
pub(super) fn read(dom: Option<&Dom>, theme: &Option<Theme>, issues: &mut Issues) -> Styles {
    let Some(dom) = dom else {
        return Styles {
            cell_styles: Vec::new(),
            normal_font: json!({"fontFamily":"Calibri","fontSize":11,"bold":false,"italic":false}),
            normal_issues: Vec::new(),
        };
    };
    let formats: HashMap<String, String> = dom
        .elements(0, "numFmt")
        .into_iter()
        .map(|n| (dom.attr(n, "numFmtId").to_string(), dom.attr(n, "formatCode").to_string()))
        .collect();
    let cell_xfs = dom.elements(0, "cellXfs").first().copied();
    let xfs: Vec<usize> = cell_xfs.map(|c| dom.children(c, "xf")).unwrap_or_default();
    struct Font {
        bold: bool,
        italic: bool,
        underline: bool,
        family: Option<String>,
        size: Option<f64>,
        color: Option<String>,
    }
    let fonts: Vec<Font> = dom
        .elements(0, "font")
        .into_iter()
        .map(|f| {
            let name = dom.first(f, "name").map(|n| dom.attr(n, "val"));
            let scheme = dom.first(f, "scheme").map(|n| dom.attr(n, "val"));
            let size = dom.first(f, "sz").map_or(f64::NAN, |n| js_number(dom.attr(n, "val")));
            Font {
                bold: flag(dom, f, "b"),
                italic: flag(dom, f, "i"),
                underline: !dom.elements(f, "u").is_empty(),
                family: theme_font(name, scheme, theme, issues),
                size: (size != 0.0 && !size.is_nan()).then_some(size),
                color: theme_color(dom, dom.first(f, "color"), theme, issues),
            }
        })
        .collect();
    let cell_style_nodes = dom.elements(0, "cellStyle");
    let normal_style = cell_style_nodes
        .iter()
        .find(|s| dom.attr(**s, "builtinId") == "0" || dom.attr(**s, "name") == "Normal")
        .copied();
    let style_bases: Vec<usize> = dom
        .elements(0, "cellStyleXfs")
        .first()
        .map(|c| dom.children(*c, "xf"))
        .unwrap_or_default();
    let base_index = normal_style.map_or(0.0, |s| js_number(dom.attr(s, "xfId")));
    let normal_base = (base_index >= 0.0 && base_index.fract() == 0.0)
        .then(|| style_bases.get(base_index as usize).copied())
        .flatten();
    let font_index = normal_base.map_or(0.0, |b| js_number(dom.attr(b, "fontId")));
    let source = (font_index >= 0.0 && font_index.fract() == 0.0)
        .then(|| fonts.get(font_index as usize))
        .flatten()
        .or(fonts.first());
    let mut normal_issues = Vec::new();
    let mut family = source.and_then(|f| f.family.clone()).unwrap_or_else(|| "Calibri".into());
    let mut size = source.and_then(|f| f.size).unwrap_or(11.0);
    if family.trim().is_empty() || family.encode_utf16().count() > 256 || family.chars().any(|c| c < ' ' || c == '\u{7f}') {
        normal_issues.push("Normal font family is unresolved; its display uses Calibri".to_string());
        family = "Calibri".into();
    }
    if !size.is_finite() || size <= 0.0 || size > 409.5 {
        normal_issues.push("Normal font size is unresolved; its display uses 11 points".to_string());
        size = 11.0;
    }
    let normal_font = json!({
        "fontFamily": family, "fontSize": number_value(size),
        "bold": source.is_some_and(|f| f.bold), "italic": source.is_some_and(|f| f.italic),
    });
    let fills: Vec<Option<String>> = dom
        .elements(0, "fill")
        .into_iter()
        .map(|f| theme_color(dom, dom.first(f, "fgColor"), theme, issues))
        .collect();
    let borders: Vec<&str> = dom
        .elements(0, "border")
        .into_iter()
        .map(|b| {
            if ["left", "right", "top", "bottom"]
                .iter()
                .any(|side| dom.first(b, side).is_some_and(|n| dom.attribute(n, "style").is_some()))
            {
                "all"
            } else {
                "none"
            }
        })
        .collect();
    let index = |text: &str| -> Option<usize> {
        let n = js_number(text);
        (n >= 0.0 && n.fract() == 0.0 && n < 4.0e9).then_some(n as usize)
    };
    let cell_styles = xfs
        .iter()
        .map(|xf| {
            let id = js_number(dom.attr(*xf, "numFmtId"));
            let key = if id.is_finite() && id.fract() == 0.0 { format!("{}", id as i64) } else { "NaN".into() };
            let mut style = Map::new();
            match formats.get(&key) {
                Some(code) => {
                    if !code.is_empty() {
                        style.insert("format".into(), json!(code));
                    }
                }
                None => {
                    style.insert("format".into(), number_value(id));
                }
            }
            if let Some(font) = index(dom.attr(*xf, "fontId")).and_then(|i| fonts.get(i)) {
                style.insert("bold".into(), json!(font.bold));
                style.insert("italic".into(), json!(font.italic));
                style.insert("underline".into(), json!(font.underline));
                if let Some(name) = font.family.as_ref().filter(|n| !n.is_empty()) {
                    style.insert("fontFamily".into(), json!(name));
                }
                if let Some(size) = font.size {
                    style.insert("fontSize".into(), number_value(size));
                }
                if let Some(color) = &font.color {
                    style.insert("fontColor".into(), json!(color));
                }
            }
            if let Some(Some(fill)) = index(dom.attr(*xf, "fillId")).and_then(|i| fills.get(i)) {
                style.insert("fillColor".into(), json!(fill));
            }
            if let Some(border) = index(dom.attr(*xf, "borderId")).and_then(|i| borders.get(i)) {
                style.insert("border".into(), json!(border));
            }
            let alignment = dom.children(*xf, "alignment").first().copied();
            if let Some(a) = alignment {
                for (attribute, key) in [("horizontal", "horizontalAlign"), ("vertical", "verticalAlign")] {
                    let value = dom.attr(a, attribute);
                    if !value.is_empty() {
                        style.insert(key.into(), json!(value));
                    }
                }
            }
            let wrap = alignment.is_some_and(|a| ["1", "true"].contains(&dom.attr(a, "wrapText")));
            style.insert("wrapText".into(), json!(wrap));
            Value::Object(style)
        })
        .collect();
    // Colors in differential styles are resolved too: a problem with one is reported.
    if let Some(dxfs) = dom.elements(0, "dxfs").first() {
        for dxf in dom.children(*dxfs, "dxf") {
            let color = dom.first(dxf, "fgColor").or_else(|| dom.first(dxf, "bgColor"));
            theme_color(dom, color, theme, issues);
        }
    }
    Styles { cell_styles, normal_font, normal_issues }
}
