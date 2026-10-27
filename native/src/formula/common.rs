//! Everyday typed criteria and exact lookup comparisons; no locale/array coercion.
use super::*;
fn numeric(value: &Value) -> Option<f64> {
    let n = value
        .as_f64()
        .or_else(|| value.as_str()?.trim().parse::<f64>().ok())?;
    n.is_finite().then_some(n)
}
pub(super) fn lookup_equal(a: &Value, b: &Value) -> bool {
    match (a.as_f64(), b.as_f64(), a.as_str(), b.as_str()) {
        (Some(a), Some(b), _, _) => a == b,
        (_, _, Some(a), Some(b)) => {
            a.eq_ignore_ascii_case(b) || a.to_lowercase() == b.to_lowercase()
        }
        _ => a == b,
    }
}
fn wildcard(pattern: &str, text: &str) -> Result<bool> {
    let mut tokens = Vec::new();
    let mut chars = pattern.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch == '~' && chars.peek().is_some_and(|ch| matches!(ch, '*' | '?' | '~')) {
            tokens.push((chars.next().unwrap(), true));
        } else {
            tokens.push((ch, false));
        }
    }
    let text = text.chars().collect::<Vec<_>>();
    if tokens.len().saturating_mul(text.len().max(1)) > 32768 {
        return Err(crate::coded(
            "formula-limit",
            "Criteria wildcard work exceeds the bounded common profile",
        ));
    }
    let mut prior = vec![false; text.len() + 1];
    prior[0] = true;
    for (ch, literal) in tokens {
        crate::cancel::check()?;
        let mut next = vec![false; prior.len()];
        if ch == '*' && !literal {
            next[0] = prior[0];
            for i in 1..next.len() {
                next[i] = prior[i] || next[i - 1];
            }
        } else {
            for i in 1..next.len() {
                next[i] = prior[i - 1]
                    && ((ch == '?' && !literal)
                        || ch.to_lowercase().eq(text[i - 1].to_lowercase()));
            }
        }
        prior = next;
    }
    Ok(prior[text.len()])
}
pub(super) fn criterion(criteria: &Value, value: &Value) -> Result<bool> {
    let text = criteria.as_str();
    let (operator, operand) = if let Some(text) = text {
        let op = ["<=", ">=", "<>", "<", ">", "="]
            .into_iter()
            .find(|op| text.starts_with(op))
            .unwrap_or("=");
        (
            op,
            if text.starts_with(op) {
                &text[op.len()..]
            } else {
                text
            },
        )
    } else {
        ("=", "")
    };
    if let Some(code) = value.get("$error").and_then(Value::as_str) {
        return Ok(text.is_some()
            && matches!(operator, "=" | "<>")
            && if operator == "=" {
                code.eq_ignore_ascii_case(operand)
            } else {
                !code.eq_ignore_ascii_case(operand)
            });
    }
    if criteria.is_boolean() {
        return Ok(value.as_bool() == criteria.as_bool());
    }
    let wanted = if criteria.is_null() {
        Some(0.0)
    } else if text.is_some() {
        operand.trim().parse::<f64>().ok().filter(|n| n.is_finite())
    } else {
        numeric(criteria)
    };
    if let Some(wanted) = wanted {
        let actual = if value.is_null() {
            Some(0.0)
        } else {
            numeric(value)
        };
        let Some(actual) = actual else {
            return Ok(operator == "<>");
        };
        return Ok(match operator {
            "=" => actual == wanted,
            "<>" => actual != wanted,
            "<" => actual < wanted,
            ">" => actual > wanted,
            "<=" => actual <= wanted,
            ">=" => actual >= wanted,
            _ => false,
        });
    }
    if text.is_none() {
        return Err(crate::coded(
            "formula-bad-input",
            "Criteria must be a scalar number, text or Boolean",
        ));
    }
    if matches!(operator, "=" | "<>") {
        let matched = if operand.is_empty() {
            value.is_null() || value.as_str() == Some("")
        } else if let Some(actual) = value.as_str() {
            wildcard(operand, actual)?
        } else if let Some(actual) = value.as_bool() {
            operand.eq_ignore_ascii_case(if actual { "TRUE" } else { "FALSE" })
        } else {
            false
        };
        return Ok(if operator == "=" { matched } else { !matched });
    }
    Err(crate::coded(
        "formula-bad-input",
        "Ordered text criteria require a wider comparison profile",
    ))
}
pub(super) fn selected(value: Value) -> Result<Value> {
    if let Some(code) = value.get("$error").and_then(Value::as_str) {
        return Err(format!("Excel error: {code}"));
    }
    Ok(if value.is_null() { json!(0) } else { value })
}
#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn ordinary_criteria_and_wildcard_escaping() {
        assert!(criterion(&json!(">=2.5"), &json!(3.0)).unwrap());
        assert!(criterion(&json!(2), &json!("2")).unwrap());
        assert!(criterion(&json!("a*?"), &json!("Abc")).unwrap());
        assert!(criterion(&json!("~*"), &json!("*")).unwrap());
        assert!(!criterion(&json!("*"), &json!(2)).unwrap());
        assert!(criterion(&json!("<>"), &json!("filled")).unwrap());
        assert!(criterion(&json!(true), &json!(true)).unwrap());
        assert!(!criterion(&json!(true), &json!(1)).unwrap());
        assert!(criterion(&json!(""), &Value::Null).unwrap());
    }
    #[test]
    fn typed_lookup_and_error_identity() {
        assert!(lookup_equal(&json!("Sku"), &json!("SKU")));
        assert!(lookup_equal(&json!(2), &json!(2.0)));
        assert!(!lookup_equal(&json!("2"), &json!(2)));
        assert!(!lookup_equal(&json!(false), &json!(0)));
        assert_eq!(selected(json!("#N/A")).unwrap(), json!("#N/A"));
        assert!(selected(json!({"$error":"#N/A"})).is_err());
    }
}
