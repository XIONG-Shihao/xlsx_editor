//! Grid display contract, shared by storage-side search. No cell collection leaves SQLite.
use crate::{Result, err};
use num_bigint::BigInt;
use num_rational::BigRational;
use num_traits::Signed;
use serde_json::Value;

mod custom_number_format;
pub(crate) use custom_number_format::format_custom_number as custom;

pub(crate) fn number(n: f64) -> String {
    if n.is_nan() {
        return "NaN".into();
    }
    if n == f64::INFINITY {
        return "Infinity".into();
    }
    if n == f64::NEG_INFINITY {
        return "-Infinity".into();
    }
    if n == 0.0 {
        return "0".into();
    }
    let text = n.to_string();
    if (1e-6..1e21).contains(&n.abs()) {
        return text;
    }
    let negative = n < 0.0;
    let raw = text.trim_start_matches('-');
    let (coefficient, exponent) = raw
        .split_once(['e', 'E'])
        .map_or((raw, 0), |(a, b)| (a, b.parse::<i32>().unwrap()));
    let point = coefficient.find('.').unwrap_or(coefficient.len()) as i32;
    let digits = coefficient.replace('.', "");
    let leading = digits.len() - digits.trim_start_matches('0').len();
    let exp = exponent + point - leading as i32 - 1;
    let digits = digits.trim_start_matches('0').trim_end_matches('0');
    let first = &digits[..1];
    let rest = &digits[1..];
    format!(
        "{}{first}{}e{}{exp}",
        if negative { "-" } else { "" },
        if rest.is_empty() {
            String::new()
        } else {
            format!(".{rest}")
        },
        if exp >= 0 { "+" } else { "" }
    )
}
fn rational_decimal(n: f64) -> Result<BigRational> {
    let text = n.abs().to_string();
    let (base, exp) = text
        .split_once(['e', 'E'])
        .map_or((text.as_str(), 0), |(a, b)| (a, b.parse::<i32>().unwrap()));
    let decimals = base.split_once('.').map_or(0, |(_, b)| b.len() as i32);
    let v = base.replace('.', "").parse::<BigInt>().map_err(err)?;
    let p = exp - decimals;
    Ok(if p >= 0 {
        BigRational::from_integer(v * BigInt::from(10).pow(p as u32))
    } else {
        BigRational::new(v, BigInt::from(10).pow((-p) as u32))
    })
}
fn fixed(n: f64, digits: u32, locale: bool) -> Result<String> {
    if !n.is_finite() || (!locale && n.abs() >= 1e21) {
        return Ok(number(n));
    }
    let ratio = if locale {
        rational_decimal(n)?
    } else {
        BigRational::from_float(n.abs())
            .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid displayed number"))?
    };
    let a = ratio.numer().abs() * BigInt::from(10).pow(digits);
    let b = ratio.denom();
    let mut q = &a / b;
    if (a % b) * 2 >= *b {
        q += 1;
    }
    let mut text = q.to_string();
    while text.len() <= digits as usize {
        text.insert(0, '0');
    }
    if digits > 0 {
        text.insert(text.len() - digits as usize, '.');
    }
    if locale {
        let end = text.find('.').unwrap_or(text.len());
        let mut at = end as isize - 3;
        while at > 0 {
            text.insert(at as usize, ',');
            at -= 3;
        }
    }
    if n < 0.0 || locale && n.is_sign_negative() {
        text.insert(0, '-');
    }
    Ok(text)
}
fn date(value: f64, date1904: bool) -> String {
    if !date1904 && value.floor() == 60.0 {
        return "1900-02-29".into();
    }
    let days = if date1904 {
        value
    } else if value < 60.0 {
        value + 1.0
    } else {
        value
    };
    let millis = (if date1904 {
        -2082844800000.0
    } else {
        -2209161600000.0
    } + days * 86400000.0)
        .trunc();
    if !millis.is_finite() || millis.abs() > 8640000000000000.0 {
        return number(value);
    }
    let z = (millis / 86400000.0).floor() as i64 + 719468;
    let era = if z >= 0 { z } else { z - 146096 } / 146097;
    let doe = z - era * 146097;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let mut y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let day = doy - (153 * mp + 2) / 5 + 1;
    let month = mp + if mp < 10 { 3 } else { -9 };
    y += i64::from(month <= 2);
    let year = if (0..10000).contains(&y) {
        format!("{y:04}")
    } else if y < 0 {
        format!("-{v:06}", v = -y)
    } else {
        format!("+{y:06}")
    };
    format!("{year}-{month:02}-{day:02}")
        .chars()
        .take(10)
        .collect()
}
pub(crate) fn formatted(n: f64, format: &Value, date1904: bool) -> Result<String> {
    if !n.is_finite() {
        return Ok(number(n));
    }
    if let Some(code) = format.as_str() {
        return Ok(custom(&serde_json::json!(n), code, date1904)?
            .map_or_else(|| number(n), |result| result.text));
    }
    match format.as_u64() {
        Some(9 | 10) => Ok(format!(
            "{}%",
            fixed(
                n * 100.0,
                if format.as_u64() == Some(9) { 0 } else { 2 },
                false
            )?
        )),
        Some(14 | 15 | 16 | 17 | 22) => Ok(date(n, date1904)),
        Some(2) => fixed(n, 2, false),
        Some(1) => fixed(n, 0, false),
        Some(3) => fixed(n, 0, true),
        Some(4 | 44) => fixed(n, 2, true),
        _ => Ok(number(n)),
    }
}
fn text_display(value: &str, format: &Value, date1904: bool) -> Result<String> {
    if let Some(code) = format.as_str() {
        if let Some(result) = custom(&Value::String(value.into()), code, date1904)? {
            return Ok(result.text);
        }
    }
    Ok(value.into())
}
fn displayed_number(value: &Value) -> Option<f64> {
    value.as_f64().or_else(|| {
        value["$workbookNumber"]
            .as_str()
            .map(|marker| match marker {
                "-0" => -0.0,
                "Infinity" => f64::INFINITY,
                "-Infinity" => f64::NEG_INFINITY,
                _ => f64::NAN,
            })
    })
}
pub(crate) fn cell(cell: &Value, format: &Value, date1904: bool) -> Result<String> {
    if cell["calculation"]["pending"] == true {
        let cache = &cell["cachedResult"];
        if cache["status"] == "error" {
            let error = cache.get("error").filter(|value| !value.is_null());
            let text = error.map_or_else(
                || "#ERROR!".into(),
                |value| {
                    value
                        .as_str()
                        .map(str::to_owned)
                        .unwrap_or_else(|| value.to_string())
                },
            );
            return Ok(format!("Stored: {text}"));
        }
        let Some(value) = cache.get("value") else {
            return Ok("Pending calculation".into());
        };
        let rendered = if let Some(n) = displayed_number(value) {
            if let Some(raw) = cache["rawValue"].as_str()
                && crate::bulk::meaning(raw) != crate::bulk::meaning(&number(n))
            {
                raw.to_owned()
            } else {
                formatted(n, cell.get("format").unwrap_or(format), date1904)?
            }
        } else {
            match value {
                Value::Null => "null".into(),
                Value::String(s) => {
                    text_display(s, cell.get("format").unwrap_or(format), date1904)?
                }
                _ => value.to_string(),
            }
        };
        return Ok(format!("Stored: {rendered}"));
    }
    let special = cell.get("formula").is_some()
        || cell.get("formulaType").is_some()
        || cell["scenarioResult"] == true;
    let stored = special && cell["calculation"]["status"] == "stored";
    let value = if special {
        match cell["calculation"]["status"].as_str() {
            Some("calculated" | "stored") => cell["calculation"]
                .get("value")
                .filter(|value| !value.is_null())
                .or_else(|| cell["calculation"].get("error"))
                .unwrap_or(&Value::Null),
            Some("circular") => return Ok("#CYCLE!".into()),
            Some("error") => {
                return Ok(cell["calculation"]["error"]
                    .as_str()
                    .unwrap_or("#ERROR!")
                    .into());
            }
            None => return Ok("Calculating…".into()),
            _ => return Ok("Unresolved".into()),
        }
    } else {
        &cell["value"]
    };
    if !special && (cell["dataType"] == "error" || cell.get("error").is_some()) {
        let error = cell
            .get("error")
            .filter(|error| !error.is_null())
            .unwrap_or(value);
        return Ok(if error.is_null() {
            String::new()
        } else {
            error
                .as_str()
                .map(str::to_owned)
                .unwrap_or_else(|| error.to_string())
        });
    }
    if special && cell["calculation"]["value"].is_null() {
        return Ok(if value.is_null() {
            String::new()
        } else {
            value
                .as_str()
                .map(str::to_owned)
                .unwrap_or_else(|| value.to_string())
        });
    }
    // Stored source/cache fields still describe the same result. A current
    // calculation must ignore those older type/error/raw fields.
    if stored
        && (cell["calculation"].get("error").is_some()
            || cell["dataType"] == "error"
            || cell.get("error").is_some())
    {
        let error = cell["calculation"]
            .get("error")
            .filter(|value| !value.is_null())
            .or_else(|| cell.get("error").filter(|value| !value.is_null()))
            .or_else(|| {
                cell["calculation"]
                    .get("value")
                    .filter(|value| !value.is_null())
            });
        return Ok(error.map_or_else(String::new, |value| {
            value
                .as_str()
                .map(str::to_owned)
                .unwrap_or_else(|| value.to_string())
        }));
    }
    if let Some(n) = displayed_number(value) {
        let raw = if stored {
            cell["calculation"]
                .get("rawValue")
                .filter(|value| !value.is_null())
                .or_else(|| cell.get("rawValue"))
                .and_then(Value::as_str)
        } else if !special {
            cell["rawValue"].as_str()
        } else {
            None
        };
        if let Some(raw) = raw
            && crate::bulk::meaning(raw) != crate::bulk::meaning(&number(n))
        {
            return Ok(raw.into());
        }
        return formatted(n, cell.get("format").unwrap_or(format), date1904);
    }
    Ok(match value {
        Value::Null => String::new(),
        Value::String(s) => text_display(s, cell.get("format").unwrap_or(format), date1904)?,
        Value::Bool(b) => b.to_string(),
        _ => value.to_string(),
    })
}

#[cfg(test)]
mod search_display_tests {
    use super::*;
    use serde_json::json;
    #[test]
    fn pending_search_uses_exact_grid_labels_without_granting_value_authority() {
        assert_eq!(cell(&json!({"formula":"1","calculation":{"pending":true},"cachedResult":{"status":"calculated","value":5},"format":2}),&Value::Null,false).unwrap(),"Stored: 5.00");
        assert_eq!(cell(&json!({"formulaOutput":{"schema":1,"master":"A1","range":"A1:A2"},"calculation":{"pending":true}}),&Value::Null,false).unwrap(),"Pending calculation");
        assert_eq!(cell(&json!({"calculation":{"pending":true},"cachedResult":{"status":"error","error":"#DIV/0!"}}),&Value::Null,false).unwrap(),"Stored: #DIV/0!");
        assert_eq!(cell(&json!({"calculation":{"pending":true},"cachedResult":{"status":"calculated","value":9007199254740992_f64,"rawValue":"9007199254740993"}}),&Value::Null,false).unwrap(),"Stored: 9007199254740993");
    }
    #[test]
    fn nonpending_stored_status_matches_the_shared_display_contract() {
        assert_eq!(
            cell(
                &json!({"formula":"1","calculation":{"status":"stored","value":5}}),
                &Value::Null,
                false
            )
            .unwrap(),
            "5"
        );
        assert_eq!(
            cell(
                &json!({"formula":"1","calculation":{"status":"calculated","error":"#N/A"}}),
                &Value::Null,
                false
            )
            .unwrap(),
            "#N/A"
        );
    }
    #[test]
    fn stored_typed_error_and_precision_fences_do_not_override_current_results() {
        let text_format = json!("0;0;0;\"text:\"@");
        let cases = [
            (
                json!({"formula":"1/0","dataType":"error","error":"#DIV/0!","calculation":{"status":"stored","value":"#DIV/0!","error":"#DIV/0!"}}),
                "#DIV/0!",
            ),
            (
                json!({"formula":"1/0","calculation":{"status":"stored","value":"#DIV/0!","error":"#N/A"}}),
                "#N/A",
            ),
            (
                json!({"formula":"1/0","dataType":"error","calculation":{"status":"stored","value":"#DIV/0!"}}),
                "#DIV/0!",
            ),
            (
                json!({"formula":"1/0","error":"#N/A","calculation":{"status":"stored","value":"#DIV/0!"}}),
                "#N/A",
            ),
            (
                json!({"formula":"\"#DIV/0!\"","dataType":"text","calculation":{"status":"stored","value":"#DIV/0!"}}),
                "text:#DIV/0!",
            ),
            (
                json!({"formula":"1/0","dataType":"error","error":"#N/A","calculation":{"status":"stored","value":null}}),
                "",
            ),
            (
                json!({"formula":"9007199254740993","rawValue":"9007199254740993","calculation":{"status":"stored","value":9007199254740992_f64}}),
                "9007199254740993",
            ),
            (
                json!({"formula":"9007199254740993","rawValue":"9007199254740992","calculation":{"status":"stored","value":9007199254740992_f64,"rawValue":"9007199254740993"}}),
                "9007199254740993",
            ),
            (
                json!({"formula":"9007199254740993","rawValue":"9007199254740993","calculation":{"status":"stored","value":9007199254740992_f64,"rawValue":null}}),
                "9007199254740993",
            ),
            (
                json!({"formula":"\"#DIV/0!\"","dataType":"error","error":"#N/A","rawValue":"9007199254740993","calculation":{"status":"calculated","value":"#DIV/0!"}}),
                "text:#DIV/0!",
            ),
            (
                json!({"formula":"7","dataType":"error","error":"#N/A","rawValue":"9007199254740993","calculation":{"status":"calculated","value":7}}),
                "7",
            ),
        ];
        for (source, expected) in cases {
            let before = source.clone();
            assert_eq!(cell(&source, &text_format, false).unwrap(), expected);
            assert_eq!(source, before);
        }
        assert_eq!(cell(&json!({"formula":"-7","dataType":"error","error":"#N/A","rawValue":"9007199254740993","calculation":{"status":"calculated","value":-7}}), &json!("$#,##0.0000;[Red]($#,##0.0000)"), false).unwrap(), "($7.0000)");
    }
}
