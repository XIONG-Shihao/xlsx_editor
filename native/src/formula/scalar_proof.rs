//! Closed scalar transformations of bounded arrays. This is not spill support.
use super::{Expr, Parser};
use crate::{Result, spill::Sequence, xml::address};
use serde_json::{Value, json};
const MAX: i128 = 9_007_199_254_740_991;
fn function(name: &str) -> &str {
    name.strip_prefix("_XLFN.").unwrap_or(name)
}
fn number(value: &Value) -> Result<i128> {
    let n = value.as_f64().ok_or_else(|| {
        crate::coded(
            "formula-bad-input",
            "Scalar array requires numeric inputs without coercion",
        )
    })?;
    if !n.is_finite()
        || n.fract() != 0.0
        || n.abs() > MAX as f64
        || n == 0.0 && n.is_sign_negative()
    {
        return Err(crate::coded(
            "formula-unsupported",
            "Scalar array input is outside exact integers",
        ));
    }
    Ok(n as i128)
}
fn scalar(expr: &Expr, read: &mut impl FnMut(&str) -> Result<Value>) -> Result<Value> {
    match expr {
        Expr::Number(n) => {
            number(&json!(n))?;
            Ok(json!(n))
        }
        Expr::String(text) => Ok(json!(text)),
        Expr::Boolean(value) => Ok(json!(value)),
        Expr::Unary(op, arg) if matches!(op, b'+' | b'-') => {
            let n = number(&scalar(arg, read)?)?;
            Ok(json!(if *op == b'-' { -(n as f64) } else { n as f64 }))
        }
        Expr::Ref(r)
            if r.sheet.is_none() && r.last.is_none() && !r.whole_columns && !r.whole_rows =>
        {
            read(&address(r.first.0, r.first.1))
        }
        _ => Err(crate::coded(
            "formula-bad-input",
            "Scalar array argument requires a literal or same-sheet literal cell",
        )),
    }
}
fn column(expr: &Expr, read: &mut impl FnMut(&str) -> Result<Value>) -> Result<Vec<Value>> {
    let Expr::Ref(r) = expr else {
        return Err(crate::coded(
            "formula-bad-input",
            "Scalar array needs a direct one-column range",
        ));
    };
    let last = r.last.unwrap_or(r.first);
    if r.sheet.is_some()
        || r.whole_columns
        || r.whole_rows
        || r.first.1 != last.1
        || r.first.0 > last.0
        || last.0 - r.first.0 >= 1000
    {
        return Err(crate::coded(
            "formula-unsupported",
            "Scalar array range is outside one-column <=1000 positions",
        ));
    }
    (r.first.0..=last.0)
        .map(|row| read(&address(row, r.first.1)))
        .collect()
}
// Lookup performs selection, not integer arithmetic. Keep Sequence/Unique
// arithmetic on number(), while qualifying typed decimal/text/Boolean atoms.
fn lookup_atom(value: Value) -> Result<Value> {
    if let Some(n) = value.as_f64() {
        if !n.is_finite() || n.abs() > MAX as f64 || n == 0.0 && n.is_sign_negative() {
            return Err(crate::coded(
                "formula-limit",
                "Lookup numeric input exceeds its finite precision bound",
            ));
        }
    } else if !value.is_string() && !value.is_boolean() {
        return Err(crate::coded(
            "formula-bad-input",
            "Lookup requires an ordinary number, text or Boolean literal",
        ));
    }
    if value.as_str().is_some_and(|s| s.len() > 4096) {
        return Err(crate::coded("formula-limit", "Lookup text exceeds4 KiB"));
    }
    Ok(value)
}
fn lookup_scalar(expr: &Expr, read: &mut impl FnMut(&str) -> Result<Value>) -> Result<Value> {
    let value = match expr {
        Expr::Number(n) => json!(n),
        Expr::String(text) => json!(text),
        Expr::Boolean(value) => json!(value),
        Expr::Unary(op, arg) if matches!(op, b'+' | b'-') => {
            let value = lookup_scalar(arg, read)?;
            let n = value.as_f64().ok_or_else(|| {
                crate::coded("formula-bad-input", "Lookup unary input must be numeric")
            })?;
            json!(if *op == b'-' { -n } else { n })
        }
        Expr::Ref(r)
            if r.sheet.is_none() && r.last.is_none() && !r.whole_columns && !r.whole_rows =>
        {
            read(&address(r.first.0, r.first.1))?
        }
        _ => {
            return Err(crate::coded(
                "formula-bad-input",
                "Lookup scalar requires a literal or same-sheet literal cell",
            ));
        }
    };
    lookup_atom(value)
}
fn lookup_vector(
    expr: &Expr,
    read: &mut impl FnMut(&str) -> Result<Value>,
) -> Result<(Vec<Value>, (u32, u32))> {
    let Expr::Ref(r) = expr else {
        return Err(crate::coded(
            "formula-bad-input",
            "Lookup range must be a direct reference",
        ));
    };
    let last = r.last.unwrap_or(r.first);
    if r.sheet.is_some()
        || r.whole_columns
        || r.whole_rows
        || last.0 < r.first.0
        || last.1 < r.first.1
    {
        return Err(crate::coded(
            "formula-bad-input",
            "Lookup range must be a same-sheet vector",
        ));
    }
    let rows = last.0 - r.first.0 + 1;
    let columns = last.1 - r.first.1 + 1;
    if rows != 1 && columns != 1 || rows as u64 * columns as u64 > 1000 {
        return Err(crate::coded(
            "formula-limit",
            "Lookup range must be one row/column with at most1000 cells",
        ));
    }
    let values = (0..rows.max(columns))
        .map(|i| {
            crate::cancel::check()?;
            lookup_atom(read(&address(
                r.first.0 + if rows > 1 { i } else { 0 },
                r.first.1 + if columns > 1 { i } else { 0 },
            ))?)
        })
        .collect::<Result<Vec<_>>>()?;
    Ok((values, (rows, columns)))
}
use super::common::lookup_equal;
fn sum(values: impl IntoIterator<Item = Value>) -> Result<Value> {
    let mut total = 0i128;
    for value in values {
        total += number(&value)?;
        if total.abs() > MAX {
            return Err(crate::coded(
                "formula-bad-input",
                "Scalar array sum requires an inexact intermediate",
            ));
        }
    }
    Ok(json!(total as f64))
}
pub(crate) fn root(formula: &str) -> bool {
    match Parser::parse(formula) {
        Ok(Expr::Call(name, args)) => {
            function(&name) == "XLOOKUP"
                || function(&name) == "SUM"
                    && args.len() == 1
                    && matches!(&args[0],Expr::Call(producer,_) if ["SEQUENCE","UNIQUE"].contains(&function(producer)))
        }
        _ => false,
    }
}
/// Output type stays separate: a literal text '#N/A' must not become an error.
pub(crate) fn calculate(
    formula: &str,
    mut read: impl FnMut(&str) -> Result<Value>,
) -> Result<Option<(Value, &'static str, String)>> {
    let expr = match Parser::parse(formula) {
        Ok(expr) => expr,
        Err(_) => return Ok(None),
    };
    let Expr::Call(name, args) = expr else {
        return Ok(None);
    };
    let result = match function(&name) {
        "SUM" if args.len() == 1 => {
            let Expr::Call(producer, inputs) = &args[0] else {
                return Ok(None);
            };
            match function(producer) {
                "SEQUENCE" if (1..=4).contains(&inputs.len()) => {
                    let supplied = inputs
                        .iter()
                        .map(|arg| number(&scalar(arg, &mut read)?).map(|n| n as f64))
                        .collect::<Result<Vec<_>>>()?;
                    sum(Sequence::new((0, 0), &supplied)?.values)?
                }
                "UNIQUE" if inputs.len() == 1 => {
                    let values = column(&inputs[0], &mut read)?;
                    let mut seen = std::collections::HashSet::new();
                    let mut unique = Vec::new();
                    for value in values {
                        if seen.insert(number(&value)?) {
                            unique.push(value);
                        }
                    }
                    sum(unique)?
                }
                _ => return Ok(None),
            }
        }
        "XLOOKUP" if (3..=6).contains(&args.len()) => {
            let key = lookup_scalar(&args[0], &mut read)?;
            let (keys, lookup_shape) = lookup_vector(&args[1], &mut read)?;
            let (returns, return_shape) = lookup_vector(&args[2], &mut read)?;
            if lookup_shape != return_shape {
                return Err(crate::coded(
                    "formula-bad-input",
                    "XLOOKUP ranges must have matching rows and columns",
                ));
            }
            if args.len() >= 5 && number(&scalar(&args[4], &mut read)?)? != 0 {
                return Err(crate::coded(
                    "formula-unsupported",
                    "Only exact XLOOKUP match mode is qualified",
                ));
            }
            let mode = if args.len() == 6 {
                number(&scalar(&args[5], &mut read)?)?
            } else {
                1
            };
            if ![1, -1].contains(&mode) {
                return Err(crate::coded(
                    "formula-unsupported",
                    "Only first/last linear XLOOKUP search is qualified",
                ));
            }
            let found = if mode == 1 {
                keys.iter().position(|n| lookup_equal(n, &key))
            } else {
                keys.iter().rposition(|n| lookup_equal(n, &key))
            };
            match found {
                Some(index) => returns[index].clone(),
                None => {
                    if args.len() >= 4 {
                        lookup_scalar(&args[3], &mut read)?
                    } else {
                        return Ok(Some((json!("#N/A"), "error", canonical(formula))));
                    }
                }
            }
        }
        "XLOOKUP" => {
            return Err(crate::coded(
                "formula-unsupported",
                "XLOOKUP argument grammar is outside the qualified scalar slice",
            ));
        }
        _ => return Ok(None),
    };
    let kind = if result.is_number() {
        "number"
    } else if result.is_boolean() {
        "boolean"
    } else {
        "text"
    };
    Ok(Some((result, kind, canonical(formula))))
}
/// Prefix only parsed modern function tokens, preserving strings and references.
fn canonical(formula: &str) -> String {
    let bytes = formula.as_bytes();
    let mut out = String::new();
    let mut at = 0;
    let mut copied = 0;
    let mut quote = 0;
    while at < bytes.len() {
        let c = bytes[at];
        if quote != 0 {
            if c == quote {
                if bytes.get(at + 1) == Some(&quote) {
                    at += 1;
                } else {
                    quote = 0;
                }
            }
            at += 1;
            continue;
        }
        if c == b'"' || c == b'\'' {
            quote = c;
            at += 1;
            continue;
        }
        if c.is_ascii_alphabetic() || c == b'_' {
            let start = at;
            while bytes
                .get(at)
                .is_some_and(|c| c.is_ascii_alphanumeric() || matches!(c, b'_' | b'.'))
            {
                at += 1;
            }
            let mut next = at;
            while bytes.get(next).is_some_and(u8::is_ascii_whitespace) {
                next += 1;
            }
            let token = &formula[start..at];
            if bytes.get(next) == Some(&b'(')
                && ["SEQUENCE", "UNIQUE", "XLOOKUP"].contains(&token.to_ascii_uppercase().as_str())
            {
                out.push_str(&formula[copied..start]);
                out.push_str("_xlfn.");
                copied = start;
            }
            continue;
        }
        at += 1;
    }
    out.push_str(&formula[copied..]);
    out
}
#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn exact_scalar_arrays_and_closed_modes() {
        assert_eq!(
            calculate("SUM(SEQUENCE(3,1))", |_| unreachable!())
                .unwrap()
                .unwrap()
                .0,
            json!(6.0)
        );
        assert_eq!(
            calculate("SUM(SEQUENCE(2,2,10,2))", |_| unreachable!())
                .unwrap()
                .unwrap()
                .0,
            json!(52.0)
        );
        let vals = [1.0, 2.0, 2.0, 4.0, 5.0];
        assert_eq!(
            calculate("SUM(UNIQUE(A1:A5))", |r| Ok(json!(
                vals[crate::xml::position(r).unwrap().0 as usize]
            )))
            .unwrap()
            .unwrap()
            .0,
            json!(12.0)
        );
        for f in [
            "SUM(SEQUENCE(1001))",
            "SUM(SEQUENCE(2,1,9007199254740991,1))",
            "SUM(SEQUENCE(1.5))",
            "SUM(SEQUENCE(2,1,-0))",
            "XLOOKUP(1,A1:A2,B1:B2,0,1)",
            "XLOOKUP(1,A1:A2,B1:B2,0,0,2)",
        ] {
            assert!(calculate(f, |_| Ok(json!(1.0))).is_err(), "{f}");
        }
        for f in [
            "SEQUENCE(3,1)",
            "SUM(SEQUENCE(3,1))+1",
            "IF(TRUE,SUM(SEQUENCE(3)),0)",
            "SUM(UNIQUE(A1:B5))",
            "SUM(UNIQUE(A1:A5,FALSE,TRUE))",
        ] {
            assert!(
                !matches!(calculate(f, |_| Ok(json!(1.0))), Ok(Some(_))),
                "{f}"
            );
        }
    }
    #[test]
    fn xlookup_first_last_types_and_missing() {
        let read = |r: &str| {
            Ok(match r {
                "A1" | "A3" => json!(2.0),
                "A2" => json!(3.0),
                "B1" => json!("#N/A"),
                "B2" => json!(false),
                "B3" => json!(9.0),
                _ => Value::Null,
            })
        };
        let first = calculate("XLOOKUP(2,A1:A3,B1:B3)", read).unwrap().unwrap();
        assert_eq!((first.0, first.1), (json!("#N/A"), "text"));
        assert_eq!(
            calculate("XLOOKUP(2,A1:A3,B1:B3,0,0,-1)", read)
                .unwrap()
                .unwrap()
                .0,
            json!(9.0)
        );
        assert_eq!(
            calculate("XLOOKUP(4,A1:A3,B1:B3)", read)
                .unwrap()
                .unwrap()
                .1,
            "error"
        );
    }
    #[test]
    fn xlookup_text_decimal_and_horizontal_vectors() {
        let read = |r: &str| {
            Ok(match r {
                "A1" => json!("SKU"),
                "A2" => json!("other"),
                "A3" => json!("sku"),
                "B1" => json!(1.25),
                "B2" => json!(2.5),
                "B3" => json!(9.0),
                _ => Value::Null,
            })
        };
        assert_eq!(
            calculate("XLOOKUP(\"sku\",A1:A3,B1:B3)", read)
                .unwrap()
                .unwrap()
                .0,
            json!(1.25)
        );
        assert_eq!(
            calculate("XLOOKUP(\"missing\",A1:A3,B1:B3,0.75)", read)
                .unwrap()
                .unwrap()
                .0,
            json!(0.75)
        );
        let row = |r: &str| {
            Ok(match r {
                "A1" => json!(1.5),
                "B1" => json!(2.25),
                "C1" => json!(3.0),
                "A2" => json!("no"),
                "B2" => json!(false),
                "C2" => json!("yes"),
                _ => Value::Null,
            })
        };
        assert_eq!(
            calculate("XLOOKUP(2.25,A1:C1,A2:C2)", row)
                .unwrap()
                .unwrap()
                .0,
            json!(false)
        );
        assert!(calculate("XLOOKUP(2.25,A1:C1,A2:A4)", row).is_err());
        assert!(calculate("SUM(UNIQUE(A1:A2))", |_| Ok(json!(1.25))).is_err());
    }
}
