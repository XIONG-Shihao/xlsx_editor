//! Read-only static reference inspection. Evaluation/provider support is separate.
use super::{Expr, Parser, resolved, sheet_name, supported_function};
use crate::xml::address;
use serde_json::{Value, json};
use std::collections::BTreeSet;
pub(crate) struct References {
    pub ranges: Vec<Value>,
    pub markers: Vec<Value>,
    pub unknown: bool,
}
struct Visitor<'a> {
    view: &'a Value,
    sheet: &'a str,
    cell: &'a str,
    text: &'a str,
    owned_sequence: bool,
    result: References,
    seen: BTreeSet<String>,
}
impl Visitor<'_> {
    fn marker(&mut self, kind: &str, reason: impl AsRef<str>) {
        self.result.unknown = true;
        if self.result.markers.len() < 64 {
            let reason: String = reason.as_ref().chars().take(240).collect();
            let marker = json!({"sheet":self.sheet,"cell":self.cell,"kind":kind,"reason":reason});
            if !self.result.markers.contains(&marker) {
                self.result.markers.push(marker);
            }
        }
    }
    fn walk(&mut self, expr: &Expr, via: &str) {
        if crate::cancel::check().is_err() {
            return;
        }
        match expr {
            Expr::Ref(r) => {
                let target = match sheet_name(self.view, r.sheet.as_deref().unwrap_or(self.sheet)) {
                    Ok(s) => s,
                    Err(reason) => {
                        self.marker("external", reason);
                        return;
                    }
                };
                let last = r.last.unwrap_or(r.first);
                if r.first.0 > last.0 || r.first.1 > last.1 {
                    self.marker("unsupported", "Reversed reference range is not qualified");
                    return;
                }
                let first = address(r.first.0, r.first.1);
                let reference = if last == r.first {
                    first
                } else {
                    format!("{first}:{}", address(last.0, last.1))
                };
                let item = json!({"sheet":target,"range":reference,"via":via});
                if self.seen.insert(item.to_string()) {
                    if self.result.ranges.len() < 512 {
                        self.result.ranges.push(item);
                    } else {
                        self.marker(
                            "limit",
                            "One formula exceeds512 distinct static reference ranges",
                        );
                    }
                }
            }
            Expr::Name(_, start, end)
            | Expr::QualifiedName(_, _, start, end)
            | Expr::Structured(_, start, end) => {
                let kind = if matches!(expr, Expr::Structured(..)) {
                    "table"
                } else {
                    "name"
                };
                let Some(token) = self.text.get(*start..*end) else {
                    self.marker("unsupported", "Reference token has no authored identity");
                    return;
                };
                match resolved(token, self.view, self.sheet, self.cell) {
                    Ok(expanded) => self.walk_resolved(&expanded, kind),
                    Err(reason) => self.marker("unsupported", reason),
                }
            }
            Expr::Unary(_, child) => self.walk(child, via),
            Expr::Binary(_, a, b) | Expr::Compare(_, a, b) => {
                self.walk(a, via);
                self.walk(b, via);
            }
            Expr::Union(areas) => {
                for area in areas {
                    self.walk(area, via)
                }
            }
            Expr::Call(name, args) => {
                self.call(name);
                for arg in args {
                    self.walk(arg, via)
                }
            }
            Expr::Error(error) if error == "#REF!" => {
                self.marker("unsupported", "Formula contains a broken#REF! reference")
            }
            Expr::QualifiedError(..) => {
                self.marker("unsupported", "Formula contains a broken#REF! reference")
            }
            _ => {}
        }
    }
    fn walk_resolved(&mut self, expr: &Expr, via: &str) {
        match expr {
            Expr::Name(..) | Expr::QualifiedName(..) | Expr::Structured(..) => self.marker(
                "unsupported",
                "A name or table did not resolve to static references",
            ),
            Expr::Unary(_, a) => self.walk_resolved(a, via),
            Expr::Binary(_, a, b) | Expr::Compare(_, a, b) => {
                self.walk_resolved(a, via);
                self.walk_resolved(b, via)
            }
            Expr::Call(name, args) => {
                self.call(name);
                for a in args {
                    self.walk_resolved(a, via)
                }
            }
            Expr::Union(areas) => {
                for a in areas {
                    self.walk_resolved(a, via)
                }
            }
            _ => self.walk(expr, via),
        }
    }
    fn call(&mut self, name: &str) {
        if ["INDIRECT", "OFFSET"].contains(&name) {
            self.marker(
                "dynamic",
                "Dynamic reference targets cannot be certified by a static trace",
            );
        } else if !(supported_function(name)
            || self.owned_sequence && ["SEQUENCE", "_XLFN.SEQUENCE"].contains(&name))
        {
            self.marker(
                "unsupported",
                format!("Function {name} has no complete static auditing profile"),
            );
        }
    }
}
pub(crate) fn references(
    view: &Value,
    sheet: &str,
    reference: &str,
    cell: &Value,
    owned_sequence: bool,
) -> crate::Result<References> {
    let text = cell["formula"].as_str().unwrap_or("");
    let mut visitor = Visitor {
        view,
        sheet,
        cell: reference,
        text,
        owned_sequence,
        result: References {
            ranges: Vec::new(),
            markers: Vec::new(),
            unknown: false,
        },
        seen: BTreeSet::new(),
    };
    if (cell.get("formulaType").is_some() && !owned_sequence) || cell["scenarioResult"] == true {
        visitor.marker("special","Imported/shared/array/scenario formula ownership is not a complete static auditing profile");
    }
    if text.is_empty() {
        return Ok(visitor.result);
    }
    match Parser::references(text) {
        Ok(ast) => visitor.walk(&ast, "formula"),
        Err(reason) => visitor.marker(
            if reason.contains("External or3D") || reason.contains("External or 3D") {
                "external"
            } else {
                "unsupported"
            },
            reason,
        ),
    }
    crate::cancel::check()?;
    Ok(visitor.result)
}
#[cfg(test)]
mod tests {
    use super::*;
    fn view() -> Value {
        json!({"sheets":[{"name":"Inputs"},{"name":"Data"},{"name":"Report"}],"namedRanges":[{"name":"RateInput","formula":"Inputs!$A$1","sheet":null}],"tables":[{"name":"Rates","sheet":"Data","ref":"A1:B4","columns":["Key","Amount"],"headerRowCount":1,"totalsRowCount":0}]})
    }
    #[test]
    fn known_ranges_names_tables_and_partial_dynamic_are_real() {
        let v = view();
        for (f, range, via) in [
            ("SUM(Inputs!A1:A2)", "A1:A2", "formula"),
            ("RateInput*3", "A1", "name"),
            ("SUM(Rates[Amount])", "B2:B4", "table"),
        ] {
            let r = references(&v, "Report", "A1", &json!({"formula":f}), false).unwrap();
            assert!(!r.unknown);
            assert_eq!(r.ranges[0]["range"], range);
            assert_eq!(r.ranges[0]["via"], via);
        }
        let r = references(
            &v,
            "Report",
            "C1",
            &json!({"formula":"INDIRECT(\"Inputs!A1\")+Inputs!A2"}),
            false,
        )
        .unwrap();
        assert!(r.unknown);
        assert_eq!(r.ranges[0]["range"], "A2");
        assert_eq!(r.markers[0]["kind"], "dynamic");
    }
    #[test]
    fn unknown_siblings_and_literal_address_text_never_fabricate_edges() {
        let v = view();
        let r = references(
            &v,
            "Report",
            "C1",
            &json!({"formula":"UNKNOWN(Inputs!A1)+MissingName+LEN(\"Z999\")"}),
            false,
        )
        .unwrap();
        assert!(r.unknown);
        assert_eq!(r.ranges.len(), 1);
        assert_eq!(r.ranges[0]["range"], "A1");
        assert_eq!(r.markers.len(), 2);
    }
    #[test]
    fn qualified_error_references_never_disappear_from_auditing() {
        let v = view();
        for formula in ["Inputs!#REF!", "'Other Sheet'!#REF!"] {
            let r = references(&v, "Report", "C1", &json!({"formula":formula}), false).unwrap();
            assert!(r.unknown);
            assert!(r.ranges.is_empty());
            assert_eq!(r.markers.len(), 1);
            assert!(r.markers[0]["reason"].as_str().unwrap().contains("#REF!"));
        }
        let r = references(
            &v,
            "Report",
            "C1",
            &json!({"formula":"\"Inputs!#REF!\""}),
            false,
        )
        .unwrap();
        assert!(!r.unknown);
        assert!(r.markers.is_empty());
    }
}
