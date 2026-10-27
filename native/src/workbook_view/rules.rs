//! The highlight and dropdown rules this editor wrote or can follow exactly:
//! solid-fill highlights, and dropdown lists or number bounds that stop bad
//! input. Anything richer is left in the file untouched and reads as no rules.
use super::dom::{Dom, MAIN};
use super::styles::js_number;
use serde_json::{Value, json};

/// `None`: not a rule this editor reads; the sheet then shows no rules.
type Read<T> = Option<T>;

fn clean_attrs(dom: &Dom, node: usize, allowed: &[&str]) -> Read<()> {
    let n = &dom.nodes[node];
    if n.ns != MAIN || n.attrs.iter().any(|a| !a.ns.is_empty() || !allowed.contains(&a.local.as_str())) {
        return None;
    }
    Some(())
}
fn clean_children(dom: &Dom, node: usize, allowed: &[&str]) -> Read<()> {
    if dom
        .child_elements(node)
        .iter()
        .any(|i| dom.nodes[*i].ns != MAIN || !allowed.contains(&dom.nodes[*i].local.as_str()))
        || dom.has_loose_text(node)
    {
        return None;
    }
    Some(())
}
fn boolean(value: &str, fallback: bool) -> Read<bool> {
    match value {
        "" => Some(fallback),
        "1" | "true" => Some(true),
        "0" | "false" => Some(false),
        _ => None,
    }
}
fn formulas(dom: &Dom, node: usize) -> Read<Vec<String>> {
    dom.child_elements(node)
        .into_iter()
        .map(|child| {
            clean_attrs(dom, child, &[])?;
            if !dom.child_elements(child).is_empty() {
                return None;
            }
            Some(dom.text(child))
        })
        .collect()
}
pub(super) fn plain_number(text: &str) -> bool {
    // ^[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?$
    let t = text.strip_prefix(['-', '+']).unwrap_or(text);
    let (mantissa, exponent) = match t.find(['e', 'E']) {
        Some(at) => (&t[..at], Some(&t[at + 1..])),
        None => (t, None),
    };
    let digits = |s: &str| !s.is_empty() && s.bytes().all(|b| b.is_ascii_digit());
    let mantissa_ok = match mantissa.split_once('.') {
        Some((whole, fraction)) => {
            (digits(whole) && (fraction.is_empty() || digits(fraction))) || (whole.is_empty() && digits(fraction))
        }
        None => digits(mantissa),
    };
    mantissa_ok
        && exponent.is_none_or(|e| digits(e.strip_prefix(['-', '+']).unwrap_or(e)))
}
fn xml_text_ok(value: &str, max: usize) -> Read<()> {
    if value.encode_utf16().count() > max
        || value
            .chars()
            .any(|c| matches!(c, '\u{0}'..='\u{8}' | '\u{b}' | '\u{c}' | '\u{e}'..='\u{1f}' | '\u{fffe}' | '\u{ffff}'))
    {
        return None;
    }
    Some(())
}
fn finite(value: f64) -> Read<f64> {
    if !value.is_finite() || value.abs() > 1e15 {
        return None;
    }
    Some(if value == 0.0 { 0.0 } else { value })
}
fn number_json(n: f64) -> Value {
    super::styles::number_value(n)
}

fn column_number(letters: &str) -> usize {
    letters.bytes().fold(0, |n, c| n * 26 + (c - b'A' + 1) as usize)
}
fn column_name(index: usize) -> String {
    let mut name = String::new();
    let mut n = index + 1;
    while n > 0 {
        name.insert(0, (b'A' + ((n - 1) % 26) as u8) as char);
        n = (n - 1) / 26;
    }
    name
}
/// `A1`, `$A$1`: (row, column) counted from zero.
fn cell(text: &str) -> Read<(usize, usize)> {
    let t = text.replace('$', "");
    let split = t.find(|c: char| c.is_ascii_digit())?;
    let (letters, digits) = t.split_at(split);
    if letters.is_empty()
        || letters.len() > 3
        || !letters.bytes().all(|b| b.is_ascii_uppercase())
        || digits.starts_with('0')
        || !digits.bytes().all(|b| b.is_ascii_digit())
    {
        return None;
    }
    let (row, column) = (digits.parse::<usize>().ok()?, column_number(letters));
    if column > 16384 || row > 1048576 {
        return None;
    }
    Some((row - 1, column - 1))
}
struct Bounds {
    first: (usize, usize),
    last: (usize, usize),
}
impl Bounds {
    fn size(&self) -> usize {
        (self.last.0 - self.first.0 + 1) * (self.last.1 - self.first.1 + 1)
    }
    fn overlaps(&self, other: &Bounds) -> bool {
        self.first.0 <= other.last.0
            && self.last.0 >= other.first.0
            && self.first.1 <= other.last.1
            && self.last.1 >= other.first.1
    }
}
fn bounds(range: &str) -> Read<Bounds> {
    let parts: Vec<&str> = range.split(':').collect();
    if parts.len() > 2 {
        return None;
    }
    let (a, b) = (cell(parts[0])?, cell(parts[parts.len() - 1])?);
    if a.0 > b.0 || a.1 > b.1 {
        return None;
    }
    Some(Bounds { first: a, last: b })
}
/// A rule range as written ("A1" or "A1:B2", dollar signs allowed), in plain "A1:B2" form.
fn rule_range(value: &str) -> Read<String> {
    let shape = |s: &str| {
        let s = s.strip_prefix('$').unwrap_or(s);
        let letters = s.bytes().take_while(u8::is_ascii_uppercase).count();
        let rest = s[letters..].strip_prefix('$').unwrap_or(&s[letters..]);
        (1..=3).contains(&letters)
            && !rest.is_empty()
            && !rest.starts_with('0')
            && rest.bytes().all(|b| b.is_ascii_digit())
    };
    let ok = match value.split_once(':') {
        Some((a, b)) => shape(a) && shape(b),
        None => shape(value),
    };
    if !ok {
        return None;
    }
    let b = bounds(value)?;
    if b.size() > 10000 {
        return None;
    }
    Some(format!(
        "{}{}:{}{}",
        column_name(b.first.1),
        b.first.0 + 1,
        column_name(b.last.1),
        b.last.0 + 1
    ))
}

fn differential_fill(styles: Option<&Dom>, index: f64) -> Read<String> {
    let dom = styles?;
    let containers: Vec<usize> = dom
        .child_elements(0)
        .into_iter()
        .filter(|i| dom.nodes[*i].local == "dxfs")
        .collect();
    if containers.len() != 1 || !(index >= 0.0 && index.fract() == 0.0 && index < 9.0e15) {
        return None;
    }
    let dxf = *dom.child_elements(containers[0]).get(index as usize)?;
    if dom.nodes[dxf].local != "dxf" {
        return None;
    }
    clean_attrs(dom, dxf, &[])?;
    clean_children(dom, dxf, &["fill"])?;
    let kids = dom.child_elements(dxf);
    if kids.len() != 1 {
        return None;
    }
    let fill = kids[0];
    clean_attrs(dom, fill, &[])?;
    clean_children(dom, fill, &["patternFill"])?;
    let kids = dom.child_elements(fill);
    if kids.len() != 1 {
        return None;
    }
    let pattern = kids[0];
    clean_attrs(dom, pattern, &["patternType"])?;
    clean_children(dom, pattern, &["fgColor", "bgColor"])?;
    if dom.attr(pattern, "patternType") != "solid" {
        return None;
    }
    let fg = dom.children(pattern, "fgColor");
    let bg = dom.children(pattern, "bgColor");
    if fg.len() != 1 || bg.len() > 1 {
        return None;
    }
    clean_attrs(dom, fg[0], &["rgb"])?;
    let rgb = dom.attr(fg[0], "rgb");
    let argb = rgb.len() == 8
        && rgb[..2].eq_ignore_ascii_case("FF")
        && rgb.bytes().all(|b| b.is_ascii_hexdigit());
    if !dom.child_elements(fg[0]).is_empty() || !argb {
        return None;
    }
    if let Some(bg) = bg.first() {
        let same = dom.attribute(*bg, "rgb").is_some();
        clean_attrs(dom, *bg, &[if same { "rgb" } else { "indexed" }])?;
        if !dom.child_elements(*bg).is_empty()
            || if same {
                !dom.attr(*bg, "rgb").eq_ignore_ascii_case(rgb)
            } else {
                dom.attr(*bg, "indexed") != "64"
            }
        {
            return None;
        }
    }
    Some(format!("#{}", rgb[2..].to_ascii_uppercase()))
}

struct Highlight {
    range: String,
    kind: String,
    fill: String,
    stop: bool,
    values: Option<Vec<f64>>,
    text: Option<String>,
}
fn check_highlights(rules: &[Highlight]) -> Read<()> {
    if rules.len() > 64 {
        return None;
    }
    let mut duplicates: Vec<&str> = Vec::new();
    for rule in rules {
        let range = rule_range(&rule.range)?;
        xml_text_ok(&rule.fill, 7)?;
        let fill = rule.fill.to_ascii_uppercase();
        if fill.len() != 7 || !fill.starts_with('#') || !fill[1..].bytes().all(|b| matches!(b, b'0'..=b'9' | b'A'..=b'F')) {
            return None;
        }
        let _ = range;
        match rule.kind.as_str() {
            "greaterThan" | "lessThan" | "equal" | "between" => {
                let values = rule.values.as_ref()?;
                if values.len() != if rule.kind == "between" { 2 } else { 1 } || rule.text.is_some() {
                    return None;
                }
                for v in values {
                    finite(*v)?;
                }
                if rule.kind == "between" && values[0] > values[1] {
                    return None;
                }
            }
            "textContains" => {
                let text = rule.text.as_ref()?;
                if rule.values.is_some() {
                    return None;
                }
                xml_text_ok(text, 128)?;
                if text.is_empty() {
                    return None;
                }
            }
            "duplicateValues" => {
                if rule.text.is_some() || rule.values.is_some() {
                    return None;
                }
                duplicates.push(&rule.range);
            }
            _ => return None,
        }
    }
    let mut ranges = Vec::new();
    for range in duplicates {
        let normal = rule_range(range)?;
        if !ranges.contains(&normal) {
            ranges.push(normal);
        }
    }
    let mut total = 0;
    for range in ranges {
        total += bounds(&range)?.size();
    }
    if total > 10000 {
        return None;
    }
    Some(())
}

/// The sheet's highlight rules in priority order, or an error when any is not one we follow exactly.
fn highlights(sheet: &Dom, styles: Option<&Dom>) -> Read<Vec<Value>> {
    let mut found: Vec<(i64, Highlight)> = Vec::new();
    let mut priorities = Vec::new();
    let containers: Vec<usize> = sheet
        .child_elements(0)
        .into_iter()
        .filter(|i| sheet.nodes[*i].local == "conditionalFormatting")
        .collect();
    for container in containers {
        clean_attrs(sheet, container, &["sqref"])?;
        clean_children(sheet, container, &["cfRule"])?;
        for node in sheet.child_elements(container) {
            clean_attrs(sheet, node, &["type", "operator", "dxfId", "priority", "stopIfTrue", "text"])?;
            clean_children(sheet, node, &["formula"])?;
            let priority = js_number(sheet.attr(node, "priority"));
            if !(priority.fract() == 0.0 && (1.0..9.0e15).contains(&priority)) || priorities.contains(&(priority as i64)) {
                return None;
            }
            priorities.push(priority as i64);
            let (kind, operator) = (sheet.attr(node, "type"), sheet.attr(node, "operator"));
            let values = formulas(sheet, node)?;
            let range = sheet.attr(container, "sqref").to_string();
            let mut rule = Highlight {
                kind: match kind {
                    "duplicateValues" => "duplicateValues".into(),
                    "containsText" => "textContains".into(),
                    "cellIs" => operator.into(),
                    _ => String::new(),
                },
                fill: differential_fill(styles, js_number(sheet.attr(node, "dxfId")))?,
                stop: boolean(sheet.attr(node, "stopIfTrue"), false)?,
                range,
                values: None,
                text: None,
            };
            match kind {
                "duplicateValues" => {
                    if !values.is_empty() || !operator.is_empty() || sheet.attribute(node, "text").is_some() {
                        return None;
                    }
                }
                "containsText" => {
                    let text = sheet.attr(node, "text").to_string();
                    let start = rule.range.split(':').next().unwrap_or("");
                    let literal = text.replace('~', "~~").replace('*', "~*").replace('?', "~?").replace('"', "\"\"");
                    let expected = format!("NOT(ISERROR(SEARCH(\"{literal}\",{start})))");
                    if operator != "containsText" || values.len() != 1 || values[0] != expected {
                        return None;
                    }
                    rule.text = Some(text);
                }
                "cellIs" => {
                    if sheet.attribute(node, "text").is_some() || values.iter().any(|v| !plain_number(v)) {
                        return None;
                    }
                    rule.values = Some(values.iter().map(|v| js_number(v)).collect());
                }
                _ => return None,
            }
            check_highlights(std::slice::from_ref(&rule))?;
            found.push((priority as i64, rule));
        }
    }
    found.sort_by_key(|(priority, _)| *priority);
    let rules: Vec<Highlight> = found.into_iter().map(|(_, rule)| rule).collect();
    check_highlights(&rules)?;
    Some(rules
        .into_iter()
        .enumerate()
        .map(|(at, rule)| {
            let mut out = serde_json::Map::new();
            out.insert("range".into(), json!(rule_range(&rule.range).unwrap_or_default()));
            out.insert("fillColor".into(), json!(rule.fill.to_ascii_uppercase()));
            out.insert("stopIfTrue".into(), json!(rule.stop));
            if let Some(values) = &rule.values {
                let values: Vec<Value> = values.iter().map(|v| number_json(if *v == 0.0 { 0.0 } else { *v })).collect();
                out.insert("threshold".into(), values[0].clone());
                out.insert("values".into(), Value::Array(values));
            }
            if let Some(text) = &rule.text {
                out.insert("text".into(), json!(text));
            }
            out.insert("rule".into(), json!(rule.kind));
            out.insert("priority".into(), json!(at + 1));
            Value::Object(out)
        })
        .collect())
}

struct Dropdown {
    range: String,
    kind: String,
    allow_blank: bool,
    show_dropdown: bool,
    texts: [String; 4],
    list: Option<Vec<String>>,
    bounds: Option<(f64, f64)>,
}
fn dropdowns(sheet: &Dom) -> Read<Vec<Value>> {
    let containers: Vec<usize> = sheet
        .child_elements(0)
        .into_iter()
        .filter(|i| sheet.nodes[*i].local == "dataValidations")
        .collect();
    if containers.len() > 1 {
        return None;
    }
    let Some(container) = containers.first().copied() else {
        return Some(Vec::new());
    };
    clean_attrs(sheet, container, &["count"])?;
    clean_children(sheet, container, &["dataValidation"])?;
    let nodes = sheet.child_elements(container);
    if sheet.attribute(container, "count").is_some() && js_number(sheet.attr(container, "count")) != nodes.len() as f64 {
        return None;
    }
    let mut rules: Vec<Dropdown> = Vec::new();
    for node in nodes {
        clean_attrs(
            sheet,
            node,
            &[
                "type", "operator", "allowBlank", "showDropDown", "showInputMessage", "showErrorMessage",
                "errorStyle", "sqref", "promptTitle", "prompt", "errorTitle", "error",
            ],
        )?;
        clean_children(sheet, node, &["formula1", "formula2"])?;
        let style = sheet.attr(node, "errorStyle");
        if (if style.is_empty() { "stop" } else { style }) != "stop" || !boolean(sheet.attr(node, "showErrorMessage"), false)? {
            return None;
        }
        let kind = match sheet.attr(node, "type") {
            "whole" => "wholeNumber",
            other => other,
        }
        .to_string();
        let texts = ["promptTitle", "prompt", "errorTitle", "error"].map(|a| sheet.attr(node, a).to_string());
        let values = formulas(sheet, node)?;
        let names: Vec<&str> = sheet.child_elements(node).iter().map(|i| sheet.nodes[*i].local.as_str()).collect();
        if names.join(",") != if kind == "list" { "formula1" } else { "formula1,formula2" } {
            return None;
        }
        if sheet.attribute(node, "showInputMessage").is_some()
            && boolean(sheet.attr(node, "showInputMessage"), true)? != !(texts[0].is_empty() && texts[1].is_empty())
        {
            return None;
        }
        let mut rule = Dropdown {
            range: sheet.attr(node, "sqref").to_string(),
            allow_blank: boolean(sheet.attr(node, "allowBlank"), false)?,
            show_dropdown: kind == "list" && !boolean(sheet.attr(node, "showDropDown"), false)?,
            texts,
            list: None,
            bounds: None,
            kind,
        };
        if rule.kind == "list" {
            let quoted = values.len() == 1
                && values[0].len() >= 3
                && values[0].starts_with('"')
                && values[0].ends_with('"')
                && !values[0][1..values[0].len() - 1].contains('"');
            if sheet.attribute(node, "operator").is_some() || !quoted {
                return None;
            }
            rule.list = Some(values[0][1..values[0].len() - 1].split(',').map(str::to_string).collect());
        } else {
            let operator = sheet.attr(node, "operator");
            if (if operator.is_empty() { "between" } else { operator }) != "between"
                || values.len() != 2
                || values.iter().any(|v| !plain_number(v))
            {
                return None;
            }
            rule.bounds = Some((js_number(&values[0]), js_number(&values[1])));
        }
        rules.push(rule);
    }
    // The shared limits every rule must meet.
    if rules.len() > 64 {
        return None;
    }
    let mut seen: Vec<Bounds> = Vec::new();
    let mut out = Vec::new();
    for rule in &rules {
        let range = rule_range(&rule.range)?;
        for (text, max) in rule.texts.iter().zip([32, 255, 32, 225]) {
            xml_text_ok(text, max)?;
        }
        let mut item = serde_json::Map::new();
        item.insert("range".into(), json!(range));
        item.insert("kind".into(), json!(rule.kind));
        item.insert("allowBlank".into(), json!(rule.allow_blank));
        item.insert("showDropdown".into(), json!(rule.show_dropdown));
        for (key, text) in ["promptTitle", "prompt", "errorTitle", "error"].iter().zip(&rule.texts) {
            item.insert((*key).into(), json!(text));
        }
        match (&rule.list, rule.bounds, rule.kind.as_str()) {
            (Some(list), _, "list") => {
                if list.is_empty() || list.len() > 128 {
                    return None;
                }
                for entry in list {
                    xml_text_ok(entry, 255)?;
                    if entry.is_empty() || entry.contains([',', '\r', '\n', '"']) {
                        return None;
                    }
                }
                let mut unique = list.clone();
                unique.sort();
                unique.dedup();
                if unique.len() != list.len() || list.join(",").encode_utf16().count() > 255 {
                    return None;
                }
                item.insert("values".into(), json!(list));
            }
            (_, Some((min, max)), "wholeNumber" | "decimal" | "textLength") => {
                let (min, max) = (finite(min)?, finite(max)?);
                let integral = |n: f64| n.fract() == 0.0 && n.abs() <= 9007199254740991.0;
                if min > max
                    || (rule.kind != "decimal" && (!integral(min) || !integral(max)))
                    || (rule.kind == "textLength" && (min < 0.0 || max > 32767.0))
                    || rule.show_dropdown
                {
                    return None;
                }
                item.insert("minimum".into(), number_json(min));
                item.insert("maximum".into(), number_json(max));
            }
            _ => return None,
        }
        let here = bounds(&range)?;
        if seen.iter().any(|other| here.overlaps(other)) {
            return None;
        }
        seen.push(here);
        out.push(Value::Object(item));
    }
    Some(out)
}

pub(super) fn conditional(sheet: &Dom, styles: Option<&Dom>) -> Vec<Value> {
    highlights(sheet, styles).unwrap_or_default()
}
pub(super) fn validation(sheet: &Dom) -> Vec<Value> {
    dropdowns(sheet).unwrap_or_default()
}
