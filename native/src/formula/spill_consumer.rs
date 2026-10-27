//! Closed scalar spill grammar. Ordinary AST/portable array support stays closed.
use super::{Expr, Parser};
use crate::Result;
use std::ops::Range;

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) enum ConsumerKind {
    Sum,
    Index(u32, u32),
}
#[derive(Clone, Debug)]
pub(crate) struct Consumer {
    pub anchor: (u32, u32),
    pub kind: ConsumerKind,
    pub canonical: String,
}
fn call(text: &str) -> Result<(String, Vec<Range<usize>>)> {
    if text.len() > 8192 {
        return Err(crate::coded(
            "spill-limit",
            "Spill consumer formula exceeds the length bound",
        ));
    }
    let bytes = text.as_bytes();
    let mut at = usize::from(text.starts_with('='));
    while bytes.get(at).is_some_and(u8::is_ascii_whitespace) {
        at += 1;
    }
    let start = at;
    while bytes
        .get(at)
        .is_some_and(|c| c.is_ascii_alphanumeric() || matches!(c, b'_' | b'.'))
    {
        at += 1;
    }
    let name = text[start..at].to_ascii_uppercase();
    while bytes.get(at).is_some_and(u8::is_ascii_whitespace) {
        at += 1;
    }
    if bytes.get(at) != Some(&b'(') {
        return Err(crate::coded(
            "spill-bad-input",
            "Spill consumer requires a root function call",
        ));
    }
    at += 1;
    let mut start = at;
    let mut depth = 0;
    let mut quote = 0;
    let mut args = Vec::new();
    while let Some(&c) = bytes.get(at) {
        if quote != 0 {
            if c == quote {
                if bytes.get(at + 1) == Some(&quote) {
                    at += 1;
                } else {
                    quote = 0;
                }
            }
        } else {
            match c {
                b'"' | b'\'' => quote = c,
                b'(' => {
                    depth += 1;
                    if depth > 64 {
                        return Err(crate::coded("spill-limit", "Spill consumer nesting bound"));
                    }
                }
                b')' => {
                    if depth == 0 {
                        args.push(start..at);
                        at += 1;
                        while bytes.get(at).is_some_and(u8::is_ascii_whitespace) {
                            at += 1;
                        }
                        if at != bytes.len() {
                            return Err(crate::coded(
                                "spill-unsupported",
                                "Spill consumers cannot be nested in another expression",
                            ));
                        }
                        return Ok((name, args));
                    }
                    depth -= 1;
                }
                b',' if depth == 0 => {
                    args.push(start..at);
                    start = at + 1;
                    if args.len() > 3 {
                        return Err(crate::coded("spill-limit", "Spill consumer argument bound"));
                    }
                }
                _ => (),
            }
        }
        at += 1;
    }
    Err(crate::coded(
        "spill-malformed",
        "Unterminated spill consumer call",
    ))
}
fn direct_anchor(text: &str) -> Result<(u32, u32)> {
    let text = text.trim();
    if text.is_empty()
        || text
            .bytes()
            .any(|c| !c.is_ascii_alphanumeric() && c != b'$')
    {
        return Err(crate::coded(
            "spill-unsupported",
            "Spill anchor requires one unqualified same-sheet cell",
        ));
    }
    let Expr::Ref(r) = Parser::parse(text)? else {
        return Err(crate::coded(
            "spill-bad-input",
            "Spill anchor is not a cell reference",
        ));
    };
    if r.sheet.is_some() || r.last.is_some() || r.whole_columns || r.whole_rows {
        return Err(crate::coded(
            "spill-bad-input",
            "Spill anchor requires a direct same-sheet master",
        ));
    }
    Ok(r.first)
}
fn axis(text: &str) -> Result<u32> {
    let text = text.trim();
    if text.is_empty() || !text.bytes().all(|c| c.is_ascii_digit()) {
        return Err(crate::coded(
            "spill-bad-input",
            "Spill INDEX axes require positive integer literals",
        ));
    }
    text.parse::<u32>().ok().filter(|n| *n > 0).ok_or_else(|| {
        crate::coded(
            "spill-bad-input",
            "Spill INDEX axes require positive integer literals",
        )
    })
}
pub(crate) fn parse(formula: &str) -> Result<Option<Consumer>> {
    if !formula.contains('#') && !formula.to_ascii_uppercase().contains("ANCHORARRAY") {
        return Ok(None);
    }
    let text = formula.trim();
    let (name, args) = match call(text) {
        Ok(call) => call,
        Err(_) => return Ok(None),
    };
    if !["SUM", "INDEX"].contains(&name.as_str()) {
        return Ok(None);
    }
    let Some(first) = args.first() else {
        return Ok(None);
    };
    let spill = text[first.clone()].trim();
    let (anchor, raw) = if let Some(raw) = spill.strip_suffix('#') {
        (direct_anchor(raw)?, raw.trim())
    } else {
        let (function, inner) = match call(spill) {
            Ok(call) => call,
            Err(_) => return Ok(None),
        };
        if !["ANCHORARRAY", "_XLFN.ANCHORARRAY"].contains(&function.as_str()) {
            return Ok(None);
        }
        if inner.len() != 1 {
            return Err(crate::coded(
                "spill-bad-input",
                "ANCHORARRAY requires one direct master reference",
            ));
        }
        let raw = spill[inner[0].clone()].trim();
        (direct_anchor(raw)?, raw)
    };
    let kind = match name.as_str() {
        "SUM" if args.len() == 1 => ConsumerKind::Sum,
        "INDEX" if args.len() == 3 => {
            ConsumerKind::Index(axis(&text[args[1].clone()])?, axis(&text[args[2].clone()])?)
        }
        _ => {
            return Err(crate::coded(
                "spill-unsupported",
                "Only single-spill SUM and scalar INDEX are qualified",
            ));
        }
    };
    let canonical = format!(
        "{}_xlfn.ANCHORARRAY({raw}){}",
        &text[..first.start],
        &text[first.end..]
    );
    Ok(Some(Consumer {
        anchor,
        kind,
        canonical,
    }))
}
/// A print protection proof, not an evaluator. Range-valued expressions, names,
/// unknown providers and array constructors cannot certify absent followers.
pub(crate) fn scalar_print_formula(formula: &str) -> bool {
    fn scalar(expr: &Expr, range: bool) -> bool {
        match expr {
            Expr::Number(_)
            | Expr::Text
            | Expr::String(_)
            | Expr::Error(_)
            | Expr::QualifiedError(..)
            | Expr::Boolean(_) => true,
            Expr::Ref(reference) => {
                range
                    || reference.last.is_none() && !reference.whole_columns && !reference.whole_rows
            }
            Expr::Name(..) | Expr::QualifiedName(..) | Expr::Structured(..) | Expr::External => {
                false
            }
            Expr::Unary(_, value) => scalar(value, false),
            Expr::Binary(_, left, right) | Expr::Compare(_, left, right) => {
                scalar(left, false) && scalar(right, false)
            }
            Expr::Union(areas) => range && areas.iter().all(|area| scalar(area, true)),
            Expr::Call(name, args) => {
                if name == "INDEX" {
                    return args.len()==3&&scalar(&args[0],true)&&args[1..].iter().all(|arg|matches!(arg,Expr::Number(value) if value.is_finite()&&*value>=1.0&&value.fract()==0.0));
                }
                if !super::supported_function(name) {
                    return false;
                }
                let aggregate = [
                    "SUM",
                    "AVERAGE",
                    "MIN",
                    "MAX",
                    "COUNT",
                    "COUNTA",
                    "PRODUCT",
                    "AND",
                    "OR",
                    "ROWS",
                    "COLUMNS",
                    "COUNTBLANK",
                    "CONCAT",
                    "CONCATENATE",
                ]
                .contains(&name.as_str());
                args.iter().enumerate().all(|(index, arg)| {
                    scalar(
                        arg,
                        aggregate
                            || match name.as_str() {
                                "VLOOKUP" | "HLOOKUP" | "MATCH" => index == 1,
                                "SUMIF" | "COUNTIF" | "AVERAGEIF" => index == 0 || index == 2,
                                "SUMIFS" | "AVERAGEIFS" => index == 0 || index % 2 == 1,
                                "COUNTIFS" => index % 2 == 0,
                                _ => false,
                            },
                    )
                })
            }
        }
    }
    formula.len() <= 8192 && Parser::parse(formula).is_ok_and(|expr| scalar(&expr, false))
}
#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn closed_aliases_and_scalar_axes() {
        for formula in [
            "SUM(A1#)",
            "=SUM($A$1#)",
            "SUM(ANCHORARRAY(A1))",
            "SUM(_xlfn.ANCHORARRAY(A1))",
        ] {
            let parsed = parse(formula).unwrap().unwrap();
            assert_eq!(parsed.anchor, (0, 0));
            assert_eq!(parsed.kind, ConsumerKind::Sum);
            assert!(parsed.canonical.contains("_xlfn.ANCHORARRAY("));
        }
        assert_eq!(
            parse("INDEX(A1#,3,2)").unwrap().unwrap().kind,
            ConsumerKind::Index(3, 2)
        );
    }
    #[test]
    fn invalid_or_broader_contexts_never_prove_a_scalar() {
        for formula in [
            "ABS(A1#)",
            "IF(TRUE,SUM(A1#),0)",
            "SUM(A1#)+1",
            "SUM(A1#,2)",
            "SUM(Sheet1!A1#)",
            "SUM(A1:B2#)",
            "SUM(Name#)",
            "SUM(ANCHORARRAY(A1,A2))",
            "INDEX(A1#,0,1)",
            "INDEX(A1#,1)",
            "INDEX(A1#,1+1,2)",
            "INDEX(A1#,2.0,2)",
            "INDEX(A1#,J1,2)",
        ] {
            assert!(
                parse(formula).is_err() || parse(formula).unwrap().is_none(),
                "{formula}"
            );
        }
        for formula in [
            "SUM(A1:A2)",
            "LEN(\"A1#\")",
            "SUM(\"ANCHORARRAY\")",
            "INDEX(A1:B2,1,1)",
        ] {
            assert!(parse(formula).unwrap().is_none(), "{formula}");
        }
    }
    #[test]
    fn hidden_array_protection_is_not_top_left_scalar_certification() {
        for formula in [
            "SEQUENCE(3,2)",
            "FILTER(A1:B2,A1:A2)",
            "A1:B2",
            "IF(TRUE,A1:B2,0)",
            "SUM(TRANSPOSE(A1:B2))",
            "INDEX(A1:B2,0,1)",
            "Name",
            "UNKNOWN(1)",
            "SUM(A1#)",
        ] {
            assert!(!scalar_print_formula(formula), "{formula}");
        }
        for formula in [
            "SUM(A1:B2)",
            "INDEX(A1:B2,1,1)",
            "IF(TRUE,A1,0)",
            "D1+1",
            "DEC2HEX(D1,4)",
            "LEN(\"SEQUENCE(3,2)\")",
        ] {
            assert!(scalar_print_formula(formula), "{formula}");
        }
    }
}
