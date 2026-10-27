//! Bounded section-aware display profile. Unsupported syntax retains the raw
//! display; this parser never changes stored values, styles or calculation state.
use crate::Result;
use serde::Serialize;
use serde_json::Value;
use std::borrow::Cow;

const CODE_UNITS: usize = 4096;
const OUTPUT_UNITS: usize = 65536;
const FRACTION_PLACES: usize = 100;
#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub(crate) struct CustomDisplay {
    pub text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    pub kind: &'static str,
    pub supported: bool,
}
#[derive(Debug, Clone)]
enum Token {
    Raw(char),
    Literal(String),
    Unsupported,
}
#[derive(Debug, Clone)]
struct Condition {
    op: String,
    threshold: f64,
}
impl Condition {
    fn matches(&self, value: f64) -> bool {
        match self.op.as_str() {
            "<" => value < self.threshold,
            "<=" => value <= self.threshold,
            ">" => value > self.threshold,
            ">=" => value >= self.threshold,
            "=" => value == self.threshold,
            "<>" => value != self.threshold,
            _ => false,
        }
    }
}
#[derive(Debug, Clone, Default)]
struct Section {
    body: Vec<Token>,
    color: Option<String>,
    condition: Option<Condition>,
}
fn color(text: &str) -> Option<&'static str> {
    match text.to_ascii_lowercase().as_str() {
        "black" => Some("000000"),
        "blue" => Some("0000FF"),
        "cyan" => Some("00FFFF"),
        "green" => Some("00FF00"),
        "magenta" => Some("FF00FF"),
        "red" => Some("FF0000"),
        "white" => Some("FFFFFF"),
        "yellow" => Some("FFFF00"),
        _ => None,
    }
}
fn condition(text: &str) -> Option<Condition> {
    let op = ["<=", ">=", "<>", "=", "<", ">"]
        .into_iter()
        .find(|op| text.starts_with(op))?;
    let number = &text[op.len()..];
    let number = number.strip_prefix(['+', '-']).unwrap_or(number);
    let mut parts = number.split('.');
    let integer = parts.next()?;
    let fraction = parts.next();
    if parts.next().is_some()
        || integer.is_empty() && fraction.is_none_or(str::is_empty)
        || integer.bytes().any(|b| !b.is_ascii_digit())
        || fraction.is_some_and(|digits| digits.bytes().any(|b| !b.is_ascii_digit()))
    {
        return None;
    }
    let number = &text[op.len()..];
    let threshold = number.parse::<f64>().ok()?;
    if !threshold.is_finite() {
        return None;
    }
    Some(Condition {
        op: op.into(),
        threshold,
    })
}
fn tokenize(code: &str) -> Option<Vec<Section>> {
    if code.is_empty() || code.encode_utf16().count() > CODE_UNITS || code.chars().any(|ch| matches!(ch, '\u{0}'..='\u{8}' | '\u{b}' | '\u{c}' | '\u{e}'..='\u{1f}' | '\u{fffe}' | '\u{ffff}')) { return None; }
    let chars: Vec<char> = code.chars().collect();
    let mut sections = vec![Section::default()];
    let mut at = 0;
    let mut prefix = true;
    while at < chars.len() {
        let ch = chars[at];
        at += 1;
        let section = sections.last_mut()?;
        match ch {
            ';' => {
                if sections.len() == 4 {
                    return None;
                }
                sections.push(Section::default());
                prefix = true;
            }
            '"' => {
                let start = at;
                while at < chars.len() && chars[at] != '"' {
                    at += 1;
                }
                if at == chars.len() {
                    return None;
                }
                section
                    .body
                    .push(Token::Literal(chars[start..at].iter().collect()));
                at += 1;
                prefix = false;
            }
            '\\' => {
                let literal = chars.get(at)?;
                section.body.push(Token::Literal(literal.to_string()));
                at += 1;
                prefix = false;
            }
            '_' | '*' => {
                chars.get(at)?;
                at += 1;
                section.body.push(Token::Unsupported);
                prefix = false;
            }
            '[' => {
                let start = at;
                while at < chars.len() && chars[at] != ']' {
                    if chars[at] == '[' {
                        return None;
                    }
                    at += 1;
                }
                if at == chars.len() {
                    return None;
                }
                let directive: String = chars[start..at].iter().collect();
                at += 1;
                if let Some(hex) = color(&directive) {
                    if !prefix || section.color.is_some() {
                        return None;
                    }
                    section.color = Some(hex.into());
                } else if let Some(condition) = condition(&directive) {
                    if !prefix || section.condition.is_some() {
                        return None;
                    }
                    section.condition = Some(condition);
                } else if let Some(currency) = directive.strip_prefix('$') {
                    if let Some((symbol, locale)) = currency.rsplit_once('-') {
                        if ["409", "0409"].contains(&locale) && !symbol.contains(['[', ']', '-']) {
                            if !symbol.is_empty() {
                                section.body.push(Token::Literal(symbol.into()));
                                prefix = false;
                            }
                        } else {
                            section.body.push(Token::Unsupported);
                            prefix = false;
                        }
                    } else {
                        section.body.push(Token::Unsupported);
                        prefix = false;
                    }
                } else {
                    section.body.push(Token::Unsupported);
                    prefix = false;
                }
            }
            ']' => return None,
            _ => {
                section.body.push(Token::Raw(ch));
                prefix = false;
            }
        }
    }
    Some(sections)
}
fn finish(text: String, section: &Section, kind: &'static str) -> Option<CustomDisplay> {
    if text.encode_utf16().count() > OUTPUT_UNITS {
        return None;
    }
    Some(CustomDisplay {
        text,
        color: section.color.clone(),
        kind,
        supported: true,
    })
}
fn literal_char(ch: char) -> bool {
    "$+-/(): !^&'~{}<>=".contains(ch)
}
fn literal(tokens: &[Token], text: Option<&str>, percentages: &mut usize) -> Option<String> {
    let mut out = String::new();
    let mut units = 0_usize;
    for token in tokens {
        let piece: Cow<'_, str> = match token {
            Token::Literal(value) => Cow::Borrowed(value),
            Token::Raw('@') if text.is_some() => Cow::Borrowed(text?),
            Token::Raw('%') => {
                *percentages += 1;
                Cow::Borrowed("%")
            }
            Token::Raw(ch) if literal_char(*ch) => Cow::Owned(ch.to_string()),
            _ => return None,
        };
        // Count this borrowed piece before extending the output. Repeated @
        // must refuse after a bounded prefix, without building a huge string.
        units = units.checked_add(piece.encode_utf16().count())?;
        if units > OUTPUT_UNITS {
            return None;
        }
        out.push_str(&piece);
    }
    Some(out)
}
fn raw_body(tokens: &[Token]) -> Option<String> {
    tokens
        .iter()
        .map(|token| match token {
            Token::Raw(ch) => Some(*ch),
            _ => None,
        })
        .collect()
}
fn iso_date(tokens: &[Token]) -> bool {
    if tokens.len() != 10 {
        return false;
    }
    tokens.iter().enumerate().all(|(index, token)| {
        let expected = "yyyy-mm-dd".as_bytes()[index] as char;
        match token {
            Token::Raw(ch) => ch.to_ascii_lowercase() == expected,
            Token::Literal(value) => expected == '-' && value == "-",
            _ => false,
        }
    })
}
fn text_section(value: &str, sections: &[Section]) -> Option<CustomDisplay> {
    let section = if sections.len() == 4 {
        &sections[3]
    } else if sections.len() == 1
        && sections[0]
            .body
            .iter()
            .any(|token| matches!(token, Token::Raw('@')))
    {
        &sections[0]
    } else {
        return finish(value.into(), &Section::default(), "text");
    };
    if section.condition.is_some() {
        return None;
    }
    let mut percentages = 0;
    let rendered = literal(&section.body, Some(value), &mut percentages)?;
    finish(rendered, section, "text")
}
fn select(sections: &[Section], value: f64) -> Option<(&Section, bool, bool)> {
    let numeric = &sections[..sections.len().min(3)];
    let conditional = numeric.iter().any(|section| section.condition.is_some());
    if conditional {
        if numeric[0].condition.is_none()
            || numeric
                .iter()
                .enumerate()
                .any(|(index, section)| index > 1 && section.condition.is_some())
        {
            return None;
        }
        for section in numeric {
            if let Some(condition) = &section.condition {
                if condition.matches(value) {
                    return Some((section, false, true));
                }
            } else {
                return Some((section, false, true));
            }
        }
        return None;
    }
    let index = if numeric.len() == 1 {
        0
    } else if value < 0.0 {
        1
    } else if numeric.len() >= 3 && value == 0.0 {
        2
    } else {
        0
    };
    Some((&numeric[index], numeric.len() > 1 && value < 0.0, false))
}
fn numeric_token(token: &Token) -> bool {
    matches!(token, Token::Raw('0' | '#' | '.' | ','))
}
fn placeholder(token: &Token) -> bool {
    matches!(token, Token::Raw('0' | '#'))
}
fn integer_pattern(value: &str) -> Option<(usize, bool)> {
    let chars: Vec<char> = value.chars().collect();
    let mut zeros = 0;
    let mut required = false;
    let mut grouped = false;
    for (index, ch) in chars.iter().enumerate() {
        match ch {
            '0' => {
                required = true;
                zeros += 1;
            }
            '#' if !required => (),
            ',' if index > 0
                && index + 1 < chars.len()
                && ['0', '#'].contains(&chars[index - 1])
                && ['0', '#'].contains(&chars[index + 1]) =>
            {
                grouped = true
            }
            _ => return None,
        }
    }
    Some((zeros, grouped))
}
fn grouping(digits: &str) -> String {
    let mut out = String::new();
    for (index, ch) in digits.chars().enumerate() {
        if index > 0 && (digits.len() - index) % 3 == 0 {
            out.push(',');
        }
        out.push(ch);
    }
    out
}
fn numeric(
    value: f64,
    section: &Section,
    absolute: bool,
    conditional: bool,
) -> Result<Option<CustomDisplay>> {
    if section
        .body
        .iter()
        .any(|token| matches!(token, Token::Unsupported))
    {
        return Ok(None);
    }
    if raw_body(&section.body)
        .is_some_and(|body| body.eq_ignore_ascii_case("General") || body == "@")
    {
        return Ok(finish(super::number(value), section, "number"));
    }
    let Some(first) = section.body.iter().position(placeholder) else {
        let mut percentages = 0;
        let Some(text) = literal(&section.body, None, &mut percentages) else {
            return Ok(None);
        };
        return Ok(if percentages == 0 {
            finish(text, section, "literal")
        } else {
            None
        });
    };
    if conditional && value < 0.0 {
        return Ok(None);
    }
    let last = section.body.iter().rposition(placeholder).unwrap();
    let mut start = first;
    while start > 0 && numeric_token(&section.body[start - 1]) {
        start -= 1;
    }
    let mut end = last + 1;
    while end < section.body.len() && numeric_token(&section.body[end]) {
        end += 1;
    }
    let Some(core) = raw_body(&section.body[start..end]) else {
        return Ok(None);
    };
    let mut pieces = core.split('.');
    let integer = pieces.next().unwrap();
    let fractional = pieces.next();
    if pieces.next().is_some() {
        return Ok(None);
    }
    let int_scaling = integer.len() - integer.trim_end_matches(',').len();
    let frac_scaling = fractional.map_or(0, |fraction| {
        fraction.len() - fraction.trim_end_matches(',').len()
    });
    if int_scaling > 0 && frac_scaling > 0 {
        return Ok(None);
    }
    let Some((minimum_integer, grouped)) = integer_pattern(integer.trim_end_matches(',')) else {
        return Ok(None);
    };
    let fraction = fractional.unwrap_or("").trim_end_matches(',');
    if fraction.len() > FRACTION_PLACES || fraction.bytes().any(|byte| !b"0#".contains(&byte)) {
        return Ok(None);
    }
    let mut percentages = 0;
    let Some(prefix) = literal(&section.body[..start], None, &mut percentages) else {
        return Ok(None);
    };
    let Some(suffix) = literal(&section.body[end..], None, &mut percentages) else {
        return Ok(None);
    };
    if percentages > 1 {
        return Ok(None);
    }
    let scaled = value.abs() * 100_f64.powi(percentages as i32)
        / 1000_f64.powi((int_scaling + frac_scaling) as i32);
    if !scaled.is_finite() {
        return Ok(None);
    }
    let fixed = super::fixed(scaled, fraction.len() as u32, grouped)?;
    if fixed.contains(['e', 'E']) {
        return Ok(None);
    }
    let (whole, rounded_fraction) = fixed.split_once('.').unwrap_or((fixed.as_str(), ""));
    let whole = whole.replace(',', "");
    let mut whole = if whole == "0" && minimum_integer == 0 {
        String::new()
    } else {
        whole
    };
    while whole.len() < minimum_integer {
        whole.insert(0, '0');
    }
    if grouped {
        whole = grouping(&whole);
    }
    let mut rounded_fraction = rounded_fraction.to_owned();
    while rounded_fraction.ends_with('0')
        && !rounded_fraction.is_empty()
        && fraction.as_bytes()[rounded_fraction.len() - 1] == b'#'
    {
        rounded_fraction.pop();
    }
    let point = if fractional.is_some() && (!rounded_fraction.is_empty() || fraction.is_empty()) {
        "."
    } else {
        ""
    };
    let sign = if value < 0.0 && !absolute { "-" } else { "" };
    Ok(finish(
        format!("{sign}{prefix}{whole}{point}{rounded_fraction}{suffix}"),
        section,
        "number",
    ))
}
pub(crate) fn format_custom_number(
    value: &Value,
    code: &str,
    date1904: bool,
) -> Result<Option<CustomDisplay>> {
    let Some(sections) = tokenize(code) else {
        return Ok(None);
    };
    if let Some(text) = value.as_str() {
        return Ok(text_section(text, &sections));
    }
    let Some(number) = value.as_f64().filter(|number| number.is_finite()) else {
        return Ok(None);
    };
    let Some((section, absolute, conditional)) = select(&sections, number) else {
        return Ok(None);
    };
    if iso_date(&section.body) {
        let selected = if absolute { number.abs() } else { number };
        if selected < 0.0 {
            return Ok(None);
        }
        let days = if date1904 {
            selected
        } else if selected < 60.0 {
            selected + 1.0
        } else {
            selected
        };
        let millis = (if date1904 {
            -2082844800000.0
        } else {
            -2209161600000.0
        }) + days * 86400000.0;
        if !millis.is_finite() || millis.abs() > 8640000000000000.0 {
            return Ok(None);
        }
        return Ok(finish(super::date(selected, date1904), section, "date"));
    }
    numeric(number, section, absolute, conditional)
}

#[cfg(test)]
mod tests;
