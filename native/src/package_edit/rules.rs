//! Closed fill/list/numeric rule authoring. Prior XML and untouched styles are
//! independently fenced; proposed arbitrary fragments are never trusted.
use super::print_settings::Tree;
use super::*;
fn allowed_fields(value: &Value, keys: &[&str]) -> Result<()> {
    if value
        .as_object()
        .is_none_or(|fields| fields.keys().any(|key| !keys.contains(&key.as_str())))
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid rule fields",
        ));
    }
    Ok(())
}
fn bool_field(value: &Value, key: &str) -> Result<bool> {
    value[key]
        .as_bool()
        .ok_or_else(|| crate::coded("internal-invalid-field", format!("Invalid rule {key}")))
}
fn string(value: &Value, key: &str, max: usize) -> Result<String> {
    let text = field(value, key)?;
    if text.encode_utf16().count() > max {
        return Err(crate::coded(
            "internal-invalid-field",
            format!("Invalid rule {key}"),
        ));
    }
    xml_text(text)?;
    Ok(text.into())
}
fn xml_text(text: &str) -> Result<()> {
    if text.chars().any(|c| {
        c < ' ' && !['\n', '\r', '\t'].contains(&c) || ['\u{fffe}', '\u{ffff}'].contains(&c)
    }) {
        return Err(crate::coded("rules-malformed", "Invalid rule XML text"));
    }
    Ok(())
}
fn number(value: &Value) -> Result<f64> {
    let n = value
        .as_f64()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid rule number"))?;
    if !n.is_finite() || n.abs() > 1e15 {
        return Err(crate::coded(
            "rules-limit",
            "Rule number exceeds precision bound",
        ));
    }
    Ok(if n == 0.0 { 0.0 } else { n })
}
fn numeric(n: f64) -> String {
    if n != 0.0 && n.abs() < 1e-6 {
        format!("{n:e}")
    } else {
        n.to_string()
    }
}
fn escape(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\r', "&#13;")
        .replace('\n', "&#10;")
        .replace('\t', "&#9;")
}
fn attrs(values: Vec<(&str, String)>) -> String {
    values
        .iter()
        .map(|(key, value)| format!(" {key}=\"{}\"", escape(value)))
        .collect()
}
fn prefix(tree: &Tree) -> String {
    tree.nodes[0]
        .qname
        .rsplit_once(':')
        .map(|(p, _)| format!("{p}:"))
        .unwrap_or_default()
}
fn range(value: &Value) -> Result<String> {
    let text = field(value, "range")?;
    let (a, b) = crate::bulk::bounds(text)?;
    let canonical = format!(
        "{}:{}",
        crate::xml::address(a.0, a.1),
        crate::xml::address(b.0, b.1)
    );
    if text != canonical || ((b.0 - a.0 + 1) as u64) * ((b.1 - a.1 + 1) as u64) > 10000 {
        return Err(crate::coded(
            "rules-limit",
            "Rule range is noncanonical or exceeds 10000 positions",
        ));
    }
    Ok(canonical)
}
fn formula(range: &str, text: &str) -> String {
    let literal = text
        .replace('~', "~~")
        .replace('*', "~*")
        .replace('?', "~?")
        .replace('"', "\"\"");
    format!(
        "NOT(ISERROR(SEARCH(\"{literal}\",{})))",
        range.split(':').next().unwrap()
    )
}
fn qualified_attrs(tree: &Tree, index: usize, allowed: &[&str]) -> Result<()> {
    let node = &tree.nodes[index];
    if node.ns != NS
        || node.attrs.keys().any(|key| {
            key != "xmlns" && !key.starts_with("xmlns:") && !allowed.contains(&key.as_str())
        })
    {
        return Err(crate::coded(
            "rules-protected",
            "Extended rule attributes are protected",
        ));
    }
    Ok(())
}
fn qualified_children(tree: &Tree, index: usize, allowed: &[&str]) -> Result<()> {
    let node = &tree.nodes[index];
    if !node.text.trim().is_empty()
        || node
            .children
            .iter()
            .any(|i| tree.nodes[*i].ns != NS || !allowed.contains(&tree.nodes[*i].name.as_str()))
    {
        return Err(crate::coded(
            "rules-protected",
            "Opaque rule XML is protected",
        ));
    }
    Ok(())
}
fn bit(value: Option<&String>, default: bool) -> Result<bool> {
    match value.map(String::as_str) {
        None | Some("") => Ok(default),
        Some("1" | "true") => Ok(true),
        Some("0" | "false") => Ok(false),
        _ => Err(crate::coded(
            "internal-invalid-field",
            "Invalid rule boolean",
        )),
    }
}
fn fill_qualified(styles: &Tree, id: usize) -> Result<()> {
    let dxfs = styles
        .direct(0, "dxfs")?
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing differential styles"))?;
    let dxf = *styles.nodes[dxfs]
        .children
        .get(id)
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing rule style"))?;
    qualified_attrs(styles, dxf, &[])?;
    qualified_children(styles, dxf, &["fill"])?;
    if styles.nodes[dxf].children.len() != 1 {
        return Err(crate::coded(
            "rules-protected",
            "Advanced differential style protected",
        ));
    }
    let fill = styles.nodes[dxf].children[0];
    qualified_attrs(styles, fill, &[])?;
    qualified_children(styles, fill, &["patternFill"])?;
    if styles.nodes[fill].children.len() != 1 {
        return Err(crate::coded("rules-protected", "Advanced fill protected"));
    }
    let pattern = styles.nodes[fill].children[0];
    qualified_attrs(styles, pattern, &["patternType"])?;
    qualified_children(styles, pattern, &["fgColor", "bgColor"])?;
    if styles.nodes[pattern]
        .attrs
        .get("patternType")
        .map(String::as_str)
        != Some("solid")
    {
        return Err(crate::coded(
            "rules-unsupported",
            "Only solid conditional fills qualified",
        ));
    }
    let fg = styles
        .direct(pattern, "fgColor")?
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing conditional foreground"))?;
    styles.clean(fg, &["rgb"])?;
    let rgb = styles.nodes[fg]
        .attrs
        .get("rgb")
        .ok_or_else(|| crate::coded("rules-unsupported", "Opaque conditional color"))?;
    if rgb.len() != 8 || !rgb.starts_with("FF") || !rgb.bytes().all(|b| b.is_ascii_hexdigit()) {
        return Err(crate::coded(
            "rules-unsupported",
            "Opaque conditional color",
        ));
    }
    // A solid fill's background is the automatic index or the same colour repeated.
    if let Some(bg) = styles.direct(pattern, "bgColor")? {
        if let Some(same) = styles.nodes[bg].attrs.get("rgb") {
            styles.clean(bg, &["rgb"])?;
            if !same.eq_ignore_ascii_case(rgb) {
                return Err(crate::coded(
                    "rules-unsupported",
                    "Opaque conditional background",
                ));
            }
        } else {
            styles.clean(bg, &["indexed"])?;
            if styles.nodes[bg].attrs.get("indexed").map(String::as_str) != Some("64") {
                return Err(crate::coded(
                    "rules-unsupported",
                    "Opaque conditional background",
                ));
            }
        }
    }
    Ok(())
}
fn selected(tree: &Tree, conditional: bool) -> Result<Vec<usize>> {
    let tag = if conditional {
        "conditionalFormatting"
    } else {
        "dataValidations"
    };
    let indexes = tree.nodes[0]
        .children
        .iter()
        .copied()
        .filter(|i| tree.nodes[*i].name == tag)
        .collect::<Vec<_>>();
    if indexes.iter().any(|i| tree.nodes[*i].ns != NS)
        || !conditional && indexes.len() > 1
        || tree.nodes[0]
            .children
            .iter()
            .any(|i| tree.nodes[*i].name == "extLst")
    {
        return Err(crate::coded(
            "rules-protected",
            "Extended/ambiguous worksheet rules protected",
        ));
    }
    Ok(indexes)
}
fn qualify_source(
    tree: &Tree,
    styles: Option<&Tree>,
    indexes: &[usize],
    conditional: bool,
) -> Result<()> {
    let mut priorities = BTreeSet::new();
    let mut count = 0;
    let mut duplicate_ranges = BTreeSet::new();
    let mut duplicate_positions = 0_u64;
    for index in indexes {
        qualified_attrs(
            tree,
            *index,
            if conditional { &["sqref"] } else { &["count"] },
        )?;
        qualified_children(
            tree,
            *index,
            if conditional {
                &["cfRule"]
            } else {
                &["dataValidation"]
            },
        )?;
        if !conditional
            && tree.nodes[*index].attrs.get("count").is_some_and(|value| {
                value.parse::<usize>().ok() != Some(tree.nodes[*index].children.len())
            })
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid validation count",
            ));
        }
        for child in &tree.nodes[*index].children {
            count += 1;
            if count > 64 {
                return Err(crate::coded("rules-limit", "Rule count exceeds 64"));
            }
            let node = &tree.nodes[*child];
            qualified_attrs(
                tree,
                *child,
                if conditional {
                    &[
                        "type",
                        "operator",
                        "dxfId",
                        "priority",
                        "stopIfTrue",
                        "text",
                    ]
                } else {
                    &[
                        "type",
                        "operator",
                        "allowBlank",
                        "showDropDown",
                        "showInputMessage",
                        "showErrorMessage",
                        "errorStyle",
                        "sqref",
                        "promptTitle",
                        "prompt",
                        "errorTitle",
                        "error",
                    ]
                },
            )?;
            qualified_children(
                tree,
                *child,
                if conditional {
                    &["formula"]
                } else {
                    &["formula1", "formula2"]
                },
            )?;
            let source_range = if conditional {
                tree.nodes[*index].attrs.get("sqref")
            } else {
                node.attrs.get("sqref")
            }
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing rule range"))?;
            let (a, b) = crate::bulk::bounds(source_range)?;
            if ((b.0 - a.0 + 1) as u64) * ((b.1 - a.1 + 1) as u64) > 10000 {
                return Err(crate::coded(
                    "rules-limit",
                    "Imported rule range exceeds qualification",
                ));
            }
            let formulas = node
                .children
                .iter()
                .map(|i| {
                    qualified_attrs(tree, *i, &[])?;
                    if !tree.nodes[*i].children.is_empty() {
                        return Err(crate::coded("rules-unsupported", "Nested rule formula"));
                    }
                    Ok(tree.nodes[*i].text.as_str())
                })
                .collect::<Result<Vec<_>>>()?;
            let kind = node
                .attrs
                .get("type")
                .map(String::as_str)
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing rule type"))?;
            if conditional {
                let priority = node
                    .attrs
                    .get("priority")
                    .and_then(|n| n.parse::<u64>().ok())
                    .filter(|n| *n > 0 && *n <= 9_007_199_254_740_991)
                    .ok_or_else(|| {
                        crate::coded("internal-invalid-field", "Invalid conditional priority")
                    })?;
                if !priorities.insert(priority) {
                    return Err(crate::coded(
                        "rules-ambiguous",
                        "Ambiguous conditional priority",
                    ));
                }
                bit(node.attrs.get("stopIfTrue"), false)?;
                let style = node
                    .attrs
                    .get("dxfId")
                    .and_then(|n| n.parse().ok())
                    .ok_or_else(|| {
                        crate::coded("internal-invalid-field", "Invalid differential style")
                    })?;
                fill_qualified(
                    styles.ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing conditional styles")
                    })?,
                    style,
                )?;
                match kind {
                    "duplicateValues"
                        if formulas.is_empty()
                            && !node.attrs.contains_key("operator")
                            && !node.attrs.contains_key("text") =>
                    {
                        // Count each normalized source rectangle once, matching
                        // the portable owner rather than physical XML spelling.
                        if duplicate_ranges.insert((a, b)) {
                            duplicate_positions +=
                                u64::from(b.0 - a.0 + 1) * u64::from(b.1 - a.1 + 1);
                            if duplicate_positions > 10000 {
                                return Err(crate::coded(
                                    "rules-duplicate",
                                    "Duplicate rules exceed 10000 source positions",
                                ));
                            }
                        }
                    }
                    "containsText"
                        if node.attrs.get("operator").map(String::as_str)
                            == Some("containsText")
                            && formulas.len() == 1
                            && node
                                .attrs
                                .get("text")
                                .is_some_and(|text| formulas[0] == formula(source_range, text)) => {
                    }
                    "cellIs" if !node.attrs.contains_key("text") => {
                        let op =
                            node.attrs
                                .get("operator")
                                .map(String::as_str)
                                .ok_or_else(|| {
                                    crate::coded("internal-missing-field", "Missing operator")
                                })?;
                        if !["greaterThan", "lessThan", "equal", "between"].contains(&op)
                            || formulas.len() != if op == "between" { 2 } else { 1 }
                        {
                            return Err(crate::coded(
                                "rules-protected",
                                "Dynamic conditional formula protected",
                            ));
                        }
                        let mut thresholds = Vec::with_capacity(formulas.len());
                        for value in formulas {
                            let n = value.parse::<f64>().map_err(err)?;
                            if !n.is_finite() || n.abs() > 1e15 {
                                return Err(crate::coded(
                                    "rules-unsupported",
                                    "Unqualified conditional threshold",
                                ));
                            }
                            thresholds.push(n);
                        }
                        if op == "between" && thresholds[0] > thresholds[1] {
                            return Err(crate::coded(
                                "rules-limit",
                                "Rule minimum exceeds maximum",
                            ));
                        }
                    }
                    _ => {
                        return Err(crate::coded(
                            "rules-protected",
                            "Imported conditional rule protected",
                        ));
                    }
                }
            } else {
                if node.attrs.get("errorStyle").is_some_and(|v| v != "stop")
                    || !bit(node.attrs.get("showErrorMessage"), false)?
                {
                    return Err(crate::coded(
                        "rules-protected",
                        "Warning/disabled validation protected",
                    ));
                }
                for key in ["allowBlank", "showDropDown", "showInputMessage"] {
                    bit(node.attrs.get(key), false)?;
                }
                match kind {
                    "list"
                        if formulas.len() == 1
                            && formulas[0].starts_with('"')
                            && formulas[0].ends_with('"')
                            && !node.attrs.contains_key("operator") =>
                    {
                        if node
                            .children
                            .iter()
                            .map(|i| tree.nodes[*i].name.as_str())
                            .collect::<Vec<_>>()
                            != vec!["formula1"]
                        {
                            return Err(crate::coded(
                                "rules-ambiguous",
                                "Ambiguous validation formulas",
                            ));
                        }
                        let value = &formulas[0][1..formulas[0].len() - 1];
                        if value.encode_utf16().count() > 255
                            || value.contains(['"', '\n', '\r'])
                            || value.split(',').any(str::is_empty)
                        {
                            return Err(crate::coded("rules-protected", "Imported list protected"));
                        }
                    }
                    "whole" | "decimal" | "textLength"
                        if formulas.len() == 2
                            && node.attrs.get("operator").is_none_or(|v| v == "between") =>
                    {
                        if node
                            .children
                            .iter()
                            .map(|i| tree.nodes[*i].name.as_str())
                            .collect::<Vec<_>>()
                            != vec!["formula1", "formula2"]
                        {
                            return Err(crate::coded(
                                "rules-ambiguous",
                                "Ambiguous validation formulas",
                            ));
                        }
                        for value in formulas {
                            let n = value.parse::<f64>().map_err(err)?;
                            if !n.is_finite() || n.abs() > 1e15 {
                                return Err(crate::coded(
                                    "rules-protected",
                                    "Dynamic validation protected",
                                ));
                            }
                        }
                    }
                    _ => {
                        return Err(crate::coded(
                            "rules-protected",
                            "Imported validation protected",
                        ));
                    }
                }
            }
        }
    }
    Ok(())
}
/// Print may use only the same independently checked, closed fill collection
/// admitted by rule authoring. No renderer-supplied evaluation grants this proof.
pub(super) fn qualify_print(tree: &Tree, styles: Option<&Tree>) -> Result<()> {
    qualify_source(tree, styles, &selected(tree, true)?, true)
}
fn generated(
    rules: &[Value],
    conditional: bool,
    p: &str,
    first_dxf: usize,
) -> Result<(String, String)> {
    if rules.len() > 64 {
        return Err(crate::coded("rules-limit", "Rule count exceeds 64"));
    }
    let mut fragment = String::new();
    let mut dxfs = String::new();
    let mut ranges = Vec::new();
    let mut duplicate_ranges = BTreeSet::new();
    let mut duplicate_positions = 0u64;
    for (index, rule) in rules.iter().enumerate() {
        crate::cancel::check()?;
        allowed_fields(
            rule,
            if conditional {
                &["range", "kind", "values", "text", "fillColor", "stopIfTrue"]
            } else {
                &[
                    "range",
                    "kind",
                    "values",
                    "minimum",
                    "maximum",
                    "allowBlank",
                    "showDropdown",
                    "promptTitle",
                    "prompt",
                    "errorTitle",
                    "error",
                ]
            },
        )?;
        let area = range(rule)?;
        let kind = field(rule, "kind")?;
        if conditional && kind == "duplicateValues" && duplicate_ranges.insert(area.clone()) {
            let (a, b) = crate::bulk::bounds(&area)?;
            duplicate_positions += (b.0 - a.0 + 1) as u64 * (b.1 - a.1 + 1) as u64;
            if duplicate_positions > 10000 {
                return Err(crate::coded(
                    "rules-duplicate",
                    "Duplicate rules exceed 10000 source positions",
                ));
            }
        }
        if conditional {
            let color = string(rule, "fillColor", 7)?;
            if color.len() != 7
                || !color.starts_with('#')
                || !color[1..]
                    .bytes()
                    .all(|b| b.is_ascii_digit() || (b'A'..=b'F').contains(&b))
            {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Invalid rule fill color",
                ));
            }
            let stop = bool_field(rule, "stopIfTrue")?;
            let (ty, operator, text, values) = match kind {
                "duplicateValues" if rule.get("values").is_none() && rule.get("text").is_none() => {
                    ("duplicateValues", None, None, Vec::new())
                }
                "textContains" if rule.get("values").is_none() => {
                    let text = string(rule, "text", 128)?;
                    if text.is_empty() {
                        return Err(crate::coded("rules-bad-input", "Empty contains text"));
                    }
                    (
                        "containsText",
                        Some("containsText"),
                        Some(text.clone()),
                        vec![formula(&area, &text)],
                    )
                }
                "greaterThan" | "lessThan" | "equal" | "between" if rule.get("text").is_none() => {
                    let values = rule["values"].as_array().ok_or_else(|| {
                        crate::coded("internal-missing-field", "Missing thresholds")
                    })?;
                    if values.len() != if kind == "between" { 2 } else { 1 } {
                        return Err(crate::coded("internal-invalid-field", "Invalid thresholds"));
                    }
                    let numbers = values.iter().map(number).collect::<Result<Vec<_>>>()?;
                    if numbers.len() == 2 && numbers[0] > numbers[1] {
                        return Err(crate::coded(
                            "internal-invalid-field",
                            "Invalid threshold order",
                        ));
                    }
                    (
                        "cellIs",
                        Some(kind),
                        None,
                        numbers.into_iter().map(numeric).collect(),
                    )
                }
                _ => {
                    return Err(crate::coded(
                        "rules-unsupported",
                        "Unsupported conditional rule",
                    ));
                }
            };
            let mut a = vec![("type", ty.into())];
            if let Some(operator) = operator {
                a.push(("operator", operator.into()))
            }
            if let Some(text) = text {
                a.push(("text", text))
            }
            a.extend([
                ("dxfId", (first_dxf + index).to_string()),
                ("priority", (index + 1).to_string()),
                ("stopIfTrue", if stop { "1" } else { "0" }.into()),
            ]);
            fragment += &format!(
                "<{p}conditionalFormatting sqref=\"{area}\"><{p}cfRule{}>{}</{p}cfRule></{p}conditionalFormatting>",
                attrs(a),
                values
                    .iter()
                    .map(|v| format!("<{p}formula>{}</{p}formula>", escape(v)))
                    .collect::<String>()
            );
            dxfs += &format!(
                "<{p}dxf><{p}fill><{p}patternFill patternType=\"solid\"><{p}fgColor rgb=\"FF{}\"/><{p}bgColor indexed=\"64\"/></{p}patternFill></{p}fill></{p}dxf>",
                &color[1..]
            );
        } else {
            let (a, b) = crate::bulk::bounds(&area)?;
            if ranges.iter().any(|(x, y): &((u32, u32), (u32, u32))| {
                a.0 <= y.0 && b.0 >= x.0 && a.1 <= y.1 && b.1 >= x.1
            }) {
                return Err(crate::coded("rules-conflict", "Validation ranges overlap"));
            }
            ranges.push((a, b));
            let allow = bool_field(rule, "allowBlank")?;
            let dropdown = bool_field(rule, "showDropdown")?;
            let title = string(rule, "promptTitle", 32)?;
            let prompt = string(rule, "prompt", 255)?;
            let error_title = string(rule, "errorTitle", 32)?;
            let error = string(rule, "error", 225)?;
            let values = if kind == "list"
                && rule.get("minimum").is_none()
                && rule.get("maximum").is_none()
            {
                let values = rule["values"].as_array().ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing validation list")
                })?;
                if values.is_empty() || values.len() > 128 {
                    return Err(crate::coded("internal-invalid-field", "Invalid list count"));
                }
                let mut unique = BTreeSet::new();
                let values = values
                    .iter()
                    .map(|v| {
                        let s = v.as_str().ok_or_else(|| {
                            crate::coded("internal-invalid-field", "Invalid list item")
                        })?;
                        xml_text(s)?;
                        if s.is_empty() || s.contains([',', '"', '\n', '\r']) || !unique.insert(s) {
                            return Err(crate::coded(
                                "internal-invalid-field",
                                "Invalid/duplicate list item",
                            ));
                        }
                        Ok(s)
                    })
                    .collect::<Result<Vec<_>>>()?;
                let joined = values.join(",");
                if joined.encode_utf16().count() > 255 {
                    return Err(crate::coded(
                        "rules-limit",
                        "Validation list exceeds 255 characters",
                    ));
                }
                vec![format!("\"{joined}\"")]
            } else if ["wholeNumber", "decimal", "textLength"].contains(&kind)
                && rule.get("values").is_none()
            {
                let lo = number(&rule["minimum"])?;
                let hi = number(&rule["maximum"])?;
                if lo > hi
                    || dropdown
                    || kind != "decimal" && (lo.fract() != 0.0 || hi.fract() != 0.0)
                    || kind == "textLength" && (lo < 0.0 || hi > 32767.0)
                {
                    return Err(crate::coded(
                        "internal-invalid-field",
                        "Invalid validation bounds",
                    ));
                }
                vec![numeric(lo), numeric(hi)]
            } else {
                return Err(crate::coded("rules-unsupported", "Unsupported validation"));
            };
            let mut a = vec![(
                "type",
                if kind == "wholeNumber" { "whole" } else { kind }.into(),
            )];
            if kind != "list" {
                a.push(("operator", "between".into()))
            }
            a.extend([
                ("allowBlank", if allow { "1" } else { "0" }.into()),
                (
                    "showDropDown",
                    if kind == "list" && !dropdown {
                        "1"
                    } else {
                        "0"
                    }
                    .into(),
                ),
                (
                    "showInputMessage",
                    if !title.is_empty() || !prompt.is_empty() {
                        "1"
                    } else {
                        "0"
                    }
                    .into(),
                ),
                ("showErrorMessage", "1".into()),
                ("errorStyle", "stop".into()),
                ("sqref", area),
                ("promptTitle", title),
                ("prompt", prompt),
                ("errorTitle", error_title),
                ("error", error),
            ]);
            fragment += &format!(
                "<{p}dataValidation{}>{}</{p}dataValidation>",
                attrs(a),
                values
                    .iter()
                    .enumerate()
                    .map(|(i, v)| format!(
                        "<{p}formula{}>{}</{p}formula{}>",
                        i + 1,
                        escape(v),
                        i + 1
                    ))
                    .collect::<String>()
            );
        }
    }
    if !conditional && !rules.is_empty() {
        fragment = format!(
            "<{p}dataValidations count=\"{}\">{fragment}</{p}dataValidations>",
            rules.len()
        )
    }
    Ok((fragment, dxfs))
}
fn append_styles(tree: &Tree, rules: &[Value]) -> Result<(Vec<u8>, usize)> {
    tree.root(NS, "styleSheet")?;
    let container = tree.direct(0, "dxfs")?;
    let old = container.map(|i| tree.nodes[i].children.len()).unwrap_or(0);
    if old + rules.len() > 4096 {
        return Err(crate::coded(
            "rules-limit",
            "Differential style bound exceeded",
        ));
    }
    let p = prefix(tree);
    let (_, added) = generated(rules, true, &p, old)?;
    if let Some(index) = container {
        qualified_attrs(tree, index, &["count"])?;
        qualified_children(tree, index, &["dxf"])?;
        let node = &tree.nodes[index];
        if node
            .attrs
            .get("count")
            .is_some_and(|v| v.parse::<usize>().ok() != Some(old))
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid differential style count",
            ));
        }
        let raw = std::str::from_utf8(tree.raw(index)).map_err(err)?;
        let end = raw
            .find('>')
            .ok_or_else(|| crate::coded("rules-malformed", "Invalid styles container"))?;
        let opening = &raw[..=end];
        let mut start = opening.to_string();
        if let Some(count) = node.attrs.get("count") {
            let mut found = false;
            for quote in ['\"', '\''] {
                let wanted = format!("count={quote}{count}{quote}");
                if let Some(at) = start.find(&wanted) {
                    let begin = at
                        .checked_sub(1)
                        .filter(|i| start.as_bytes()[*i].is_ascii_whitespace())
                        .ok_or_else(|| {
                            crate::coded("internal-invalid-field", "Noncanonical style count")
                        })?;
                    start.replace_range(begin..at + wanted.len(), "");
                    found = true;
                    break;
                }
            }
            if !found {
                return Err(crate::coded(
                    "internal-invalid-field",
                    "Noncanonical style count",
                ));
            }
        }
        let start = start
            .trim_end_matches('>')
            .trim_end()
            .trim_end_matches('/')
            .trim_end();
        let body = if opening.trim_end().ends_with("/>") {
            ""
        } else {
            &raw[end + 1
                ..raw.rfind("</").ok_or_else(|| {
                    crate::coded("internal-invalid-field", "Invalid styles close")
                })?]
        };
        let replacement = format!(
            "{start} count=\"{}\">{body}{added}</{}>",
            old + rules.len(),
            node.qname
        );
        let mut output = tree.bytes[..node.start].to_vec();
        output.extend(replacement.as_bytes());
        output.extend(&tree.bytes[node.end..]);
        Ok((output, old))
    } else {
        let at = tree.nodes[0]
            .children
            .iter()
            .find(|i| ["tableStyles", "colors", "extLst"].contains(&tree.nodes[**i].name.as_str()))
            .map(|i| tree.nodes[*i].start)
            .unwrap_or_else(|| {
                std::str::from_utf8(&tree.bytes)
                    .unwrap()
                    .rfind("</")
                    .unwrap()
            });
        let mut output = tree.bytes[..at].to_vec();
        output.extend(format!("<{p}dxfs count=\"{}\">{added}</{p}dxfs>", rules.len()).as_bytes());
        output.extend(&tree.bytes[at..]);
        Ok((output, old))
    }
}
pub(super) fn validate(root: &Path, plan: &Value, input: &Value) -> Result<()> {
    validate_files(
        &local(root, "metadata.xlsx")?,
        &local(root, field(input, "source")?)?,
        plan,
    )
}
pub(super) fn validate_files(metadata_path: &Path, source_path: &Path, plan: &Value) -> Result<()> {
    if !["setConditionalRules", "setValidationRules"].contains(&field(plan, "kind")?) {
        return Err(crate::coded(
            "rules-unsupported",
            "Unqualified rule collection operation",
        ));
    }
    allowed_fields(
        plan,
        &[
            "kind",
            "sheetPart",
            "sheetName",
            "scope",
            "priorRulesRevision",
            "rules",
            "replacements",
            "removed",
            "copies",
            "fragments",
        ],
    )?;
    let conditional = plan["kind"] == "setConditionalRules";
    let part = field(plan, "sheetPart")?;
    let name = field(plan, "sheetName")?;
    let scope = plan["scope"]
        .as_u64()
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid rules scope"))?;
    let rules = plan["rules"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing rules"))?;
    if plan["removed"].as_array().is_none_or(|v| !v.is_empty())
        || plan["copies"].as_object().is_none_or(|v| !v.is_empty())
    {
        return Err(crate::coded(
            "rules-unsupported",
            "Rules cannot remove/copy package parts",
        ));
    }
    let replacements = plan["replacements"]
        .as_object()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing rules replacements"))?;
    if replacements.len() != usize::from(conditional && !rules.is_empty())
        || replacements.keys().any(|k| k != "xl/styles.xml")
    {
        return Err(crate::coded(
            "rules-unsupported",
            "Rules may only append to existing styles",
        ));
    }
    let fragments = plan["fragments"]
        .as_object()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing rules fragment"))?;
    let tag = if conditional {
        "conditionalFormatting"
    } else {
        "dataValidations"
    };
    if fragments.len() != 1
        || fragments
            .get(part)
            .and_then(Value::as_object)
            .is_none_or(|v| v.len() != 1 || !v.contains_key(tag))
    {
        return Err(crate::coded(
            "rules-refused",
            "Rules touch exactly one worksheet collection",
        ));
    }
    let empty = serde_json::Map::new();
    let mut metadata = ZipArchive::new(crate::cancel::Reader(
        File::open(metadata_path).map_err(err)?,
    ))
    .map_err(err)?;
    let sheet = Tree::parse(&part_bytes(&mut metadata, part, &empty)?)?;
    sheet.root(NS, "worksheet")?;
    let workbook = Tree::parse(&part_bytes(&mut metadata, "xl/workbook.xml", &empty)?)?;
    workbook.root(NS, "workbook")?;
    let sheets = workbook
        .direct(0, "sheets")?
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing rule sheets"))?;
    let node = *workbook.nodes[sheets]
        .children
        .get(scope as usize)
        .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid rule scope"))?;
    if workbook.nodes[node].ns != NS
        || workbook.nodes[node].name != "sheet"
        || workbook.nodes[node].attrs.get("name").map(String::as_str) != Some(name)
    {
        return Err(crate::coded(
            "internal-mismatch",
            "Rules worksheet ownership mismatch",
        ));
    }
    let id = workbook.nodes[node].attrs.get("r:id").ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing rules worksheet relationship",
        )
    })?;
    let relationships = Tree::parse(&part_bytes(
        &mut metadata,
        "xl/_rels/workbook.xml.rels",
        &empty,
    )?)?;
    relationships.root(PACKAGE_REL_NS, "Relationships")?;
    let edges = relationships.nodes[0]
        .children
        .iter()
        .filter(|i| relationships.nodes[**i].attrs.get("Id") == Some(id))
        .collect::<Vec<_>>();
    if edges.len() != 1 {
        return Err(crate::coded(
            "rules-ambiguous",
            "Ambiguous rule relationship",
        ));
    }
    let edge = &relationships.nodes[*edges[0]];
    if edge.ns != PACKAGE_REL_NS
        || edge.name != "Relationship"
        || edge.attrs.contains_key("TargetMode")
        || edge.attrs.get("Type").map(String::as_str)
            != Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
        || resolve(
            "xl/workbook.xml",
            edge.attrs
                .get("Target")
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing rule target"))?,
        )? != part
    {
        return Err(crate::coded(
            "internal-invalid-field",
            "Invalid rule worksheet edge",
        ));
    }
    let has_styles = metadata.by_name("xl/styles.xml").is_ok();
    let styles = if has_styles {
        Some(Tree::parse(&part_bytes(
            &mut metadata,
            "xl/styles.xml",
            &empty,
        )?)?)
    } else {
        None
    };
    let selected = selected(&sheet, conditional)?;
    qualify_source(&sheet, styles.as_ref(), &selected, conditional)?;
    let revision = if selected.is_empty() {
        Value::Null
    } else {
        let kind = if conditional {
            "conditional"
        } else {
            "validation"
        };
        let mut pieces = vec![
            /* Saved V1 CAS identity: hash the original domain separator. */
            workbook_format!(RULES_V1).to_string(),
            kind.into(),
            part.into(),
            name.into(),
            scope.to_string(),
        ];
        for i in &selected {
            pieces.push(std::str::from_utf8(sheet.raw(*i)).map_err(err)?.into())
        }
        if conditional {
            if let Some(styles) = &styles {
                if let Some(dxfs) = styles.direct(0, "dxfs")? {
                    pieces.push(std::str::from_utf8(styles.raw(dxfs)).map_err(err)?.into())
                }
            }
        }
        Value::String(format!(
            "{:x}",
            Sha256::digest(pieces.join("\0").as_bytes())
        ))
    };
    if plan["priorRulesRevision"] != revision {
        return Err(crate::coded(
            "rules-stale",
            "Conflict: native rules definition changed",
        ));
    }
    let mut first = 0;
    if conditional && !rules.is_empty() {
        if styles.is_none() {
            return Err(crate::coded(
                "rules-unsupported",
                "Conditional authoring requires existing styles",
            ));
        }
        let style_edges=relationships.nodes[0].children.iter().map(|i|&relationships.nodes[*i]).filter(|node|node.attrs.get("Type").map(String::as_str)==Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")).collect::<Vec<_>>();
        if style_edges.len() != 1
            || style_edges[0].ns != PACKAGE_REL_NS
            || style_edges[0].name != "Relationship"
            || style_edges[0].attrs.contains_key("TargetMode")
            || resolve(
                "xl/workbook.xml",
                style_edges[0]
                    .attrs
                    .get("Target")
                    .ok_or_else(|| crate::coded("internal-missing-field", "Missing styles edge"))?,
            )? != "xl/styles.xml"
        {
            return Err(crate::coded(
                "rules-unsupported",
                "Unqualified conditional style ownership",
            ));
        }
        let types = Tree::parse(&part_bytes(&mut metadata, "[Content_Types].xml", &empty)?)?;
        types.root(TYPE_NS, "Types")?;
        let claims = types.nodes[0]
            .children
            .iter()
            .map(|i| &types.nodes[*i])
            .filter(|node| node.attrs.get("PartName").map(String::as_str) == Some("/xl/styles.xml"))
            .collect::<Vec<_>>();
        if claims.len() != 1
            || claims[0].ns != TYPE_NS
            || claims[0].name != "Override"
            || claims[0].attrs.get("ContentType").map(String::as_str)
                != Some("application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
        {
            return Err(crate::coded(
                "rules-unsupported",
                "Unqualified conditional styles content type",
            ));
        }
        let mut exported =
            ZipArchive::new(crate::cancel::Reader(File::open(source_path).map_err(err)?))
                .map_err(err)?;
        let current = Tree::parse(&part_bytes(&mut exported, "xl/styles.xml", &empty)?)?;
        let (expected, old) = append_styles(&current, rules)?;
        first = old;
        if replacements["xl/styles.xml"].as_str().map(str::as_bytes) != Some(expected.as_slice()) {
            return Err(crate::coded(
                "rules-stale",
                "Rules changed unrelated styles or injected DXFs",
            ));
        }
    }
    let (expected, _) = generated(rules, conditional, &prefix(&sheet), first)?;
    if fragments[part][tag].as_str() != Some(expected.as_str()) {
        return Err(crate::coded(
            "internal-mismatch",
            "Rules fragment differs from qualified command",
        ));
    }
    Ok(())
}
