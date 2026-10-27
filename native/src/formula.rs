//! Deliberately bounded, fail-closed formula fast path. Unsupported syntax or
//! coercion returns a miss; the compatibility engine owns the complete language.
use crate::xml::{address, position};
use crate::{Result, field};
use serde_json::{Value, json};
use std::collections::{HashMap, HashSet};
pub(crate) mod auditing;
mod date;
mod financial;
pub(crate) use financial::profiles as financial_profiles;
mod common;
mod scalar_proof;
mod spill_consumer;
pub(crate) use scalar_proof::{calculate as scalar_array_result, root as scalar_array_root};
pub(crate) use spill_consumer::scalar_print_formula;
pub(crate) use spill_consumer::{ConsumerKind, parse as spill_consumer};
mod syntax_cache;
pub(crate) use syntax_cache::begin as syntax_cache_scope;
// Excel 16.113.3, fixtures/excel-union-reference.json: ordered unqualified
// same-sheet SUM areas retain overlaps/duplicates. Typed error inputs still
// refuse native evaluation; cross-sheet/name/table/scalar unions remain closed.

#[derive(Clone, Debug)]
struct Ref {
    sheet: Option<String>,
    prefix: String,
    first: (u32, u32),
    last: Option<(u32, u32)>,
    fixed: (bool, bool),
    last_fixed: (bool, bool),
    whole_columns: bool,
    whole_rows: bool,
    start: usize,
    end: usize,
}
#[derive(Clone, Debug)]
enum Expr {
    Number(f64),
    Text,
    String(String),
    Error(String),
    QualifiedError(String, usize, usize), // Sheet and exact qualifier span; #REF! only.
    Boolean(bool),
    Compare(String, Box<Expr>, Box<Expr>),
    Ref(Ref),
    Name(String, usize, usize),
    QualifiedName(String, String, usize, usize),
    Structured(String, usize, usize),
    /// A reference into another workbook ('[1]Sheet'!A1). Its cells are not in
    /// this workbook, so no edit here moves it and it cannot be calculated here.
    External,
    Unary(u8, Box<Expr>),
    Binary(u8, Box<Expr>, Box<Expr>),
    Call(String, Vec<Expr>),
    Union(Vec<Expr>),
}
fn union_context(expr: &Expr, aggregate: bool) -> Result<()> {
    match expr {
        Expr::Union(areas) => {
            if !aggregate {
                return Err(crate::coded(
                    "formula-unsupported",
                    "Union references are qualified only as SUM arguments",
                ));
            }
            for area in areas {
                if matches!(area, Expr::Ref(reference) if reference.sheet.is_some())
                    || matches!(area, Expr::Binary(b':', a, b)
                        if [a.as_ref(), b.as_ref()].iter().any(|e| matches!(e, Expr::Ref(r) if r.sheet.is_some())))
                {
                    return Err(crate::coded(
                        "formula-unsupported",
                        "Union areas require unqualified same-worksheet references",
                    ));
                }
                if !matches!(
                    area,
                    Expr::Ref(_) | Expr::Union(_) | Expr::Binary(b':', _, _)
                ) && !matches!(area, Expr::Error(error) if error == "#REF!")
                {
                    return Err(crate::coded(
                        "formula-bad-input",
                        "Union areas require direct static references",
                    ));
                }
                union_context(area, true)?;
            }
        }
        Expr::Unary(_, e) => union_context(e, false)?,
        Expr::Binary(_, a, b) | Expr::Compare(_, a, b) => {
            union_context(a, false)?;
            union_context(b, false)?;
        }
        Expr::Call(name, args) => {
            for arg in args {
                union_context(arg, name == "SUM")?;
            }
        }
        _ => {}
    }
    Ok(())
}
#[cfg(test)]
fn contains_union(expr: &Expr) -> bool {
    match expr {
        Expr::Union(_) => true,
        Expr::Unary(_, e) => contains_union(e),
        Expr::Binary(_, a, b) | Expr::Compare(_, a, b) => contains_union(a) || contains_union(b),
        Expr::Call(_, args) => args.iter().any(contains_union),
        _ => false,
    }
}
struct Parser<'a> {
    text: &'a str,
    at: usize,
    depth: usize,
    nodes: usize,
    node_limit: usize,
}
impl<'a> Parser<'a> {
    fn parse(text: &'a str) -> Result<Expr> {
        syntax_cache::parse(text, 256)
    }
    fn references(text: &'a str) -> Result<Expr> {
        syntax_cache::parse(text, 8192)
    }
    fn bounded(text: &'a str, node_limit: usize) -> Result<Expr> {
        if text.len() > 8192 {
            return Err(crate::coded(
                "formula-limit",
                "Formula fast-path length bound",
            ));
        }
        let mut p = Self {
            text,
            at: usize::from(text.starts_with('=')),
            depth: 0,
            nodes: 0,
            node_limit,
        };
        let result = p.expr(0)?;
        p.space();
        if p.at != text.len() {
            return Err(crate::coded(
                "formula-unsupported",
                "Unsupported formula syntax",
            ));
        }
        union_context(&result, false)?;
        Ok(result)
    }
    fn peek(&self) -> u8 {
        self.text.as_bytes().get(self.at).copied().unwrap_or(0)
    }
    fn space(&mut self) {
        while self.peek().is_ascii_whitespace() {
            self.at += 1
        }
    }
    fn expr(&mut self, min: u8) -> Result<Expr> {
        self.depth += 1;
        self.nodes += 1;
        if self.nodes > self.node_limit {
            return Err(crate::coded("formula-limit", "Formula node bound"));
        }
        if self.depth > 64 {
            return Err(crate::coded("formula-limit", "Formula nesting bound"));
        }
        self.space();
        let mut left = if matches!(self.peek(), b'+' | b'-') {
            let op = self.peek();
            self.at += 1;
            Expr::Unary(op, Box::new(self.expr(6)?))
        } else {
            self.atom()?
        };
        loop {
            self.space();
            if self.peek() == b'%' {
                self.nodes += 1;
                if self.nodes > self.node_limit {
                    return Err(crate::coded("formula-limit", "Formula node bound"));
                }
                self.at += 1;
                left = Expr::Unary(b'%', Box::new(left));
                continue;
            }
            let op = self.peek();
            let power = match op {
                b':' => 7,
                b'=' | b'<' | b'>' => 1,
                b'+' | b'-' => 2,
                b'*' | b'/' => 3,
                b'^' => 5,
                b'&' => 2,
                _ => 0,
            };
            if power == 0 || power < min {
                break;
            }
            self.at += 1;
            let mut comparison = String::from(op as char);
            if power == 1
                && ((op == b'<' && matches!(self.peek(), b'=' | b'>'))
                    || (op == b'>' && self.peek() == b'='))
            {
                comparison.push(self.peek() as char);
                self.at += 1;
            }
            let right = self.expr(power + 1)?;
            left = if op == b':' {
                // Ordinary A1 ranges are parsed as one Ref. A range with a
                // #REF! endpoint keeps its surviving endpoint's source span.
                fn endpoint(e: &Expr) -> bool {
                    matches!(e, Expr::Ref(r) if r.last.is_none())
                        || matches!(e, Expr::Error(error) if error == "#REF!")
                        || matches!(e, Expr::QualifiedError(..))
                }
                fn broken(e: &Expr) -> bool {
                    matches!(e, Expr::Error(error) if error == "#REF!")
                        || matches!(e, Expr::QualifiedError(..))
                }
                if !endpoint(&left) || !endpoint(&right) || !broken(&left) && !broken(&right) {
                    return Err(crate::coded(
                        "formula-unsupported",
                        "Unsupported reference range operator",
                    ));
                }
                Expr::Binary(op, Box::new(left), Box::new(right))
            } else if power == 1 {
                Expr::Compare(comparison, Box::new(left), Box::new(right))
            } else {
                Expr::Binary(op, Box::new(left), Box::new(right))
            };
        }
        self.depth -= 1;
        Ok(left)
    }
    fn whole_rows(
        &mut self,
        start: usize,
        sheet: Option<String>,
        ref_start: usize,
    ) -> Result<Option<Expr>> {
        let bytes = self.text.as_bytes();
        let mut at = self.at;
        let first_fixed = bytes.get(at) == Some(&b'$');
        if first_fixed {
            at += 1;
        }
        let a = at;
        while bytes.get(at).is_some_and(u8::is_ascii_digit) {
            at += 1;
        }
        if a == at || bytes.get(at) != Some(&b':') {
            return Ok(None);
        }
        let first = self.text[a..at].parse::<u32>().map_err(crate::err)?;
        at += 1;
        let last_fixed = bytes.get(at) == Some(&b'$');
        if last_fixed {
            at += 1;
        }
        let b = at;
        while bytes.get(at).is_some_and(u8::is_ascii_digit) {
            at += 1;
        }
        if b == at
            || bytes
                .get(at)
                .is_some_and(|c| c.is_ascii_alphanumeric() || matches!(c, b'_' | b'.' | b'$'))
        {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid whole-row reference",
            ));
        }
        let last = self.text[b..at].parse::<u32>().map_err(crate::err)?;
        if first == 0 || first > last || last > 1048576 {
            return Err(crate::coded(
                "internal-invalid-field",
                "Invalid whole-row reference",
            ));
        }
        self.at = at;
        Ok(Some(Expr::Ref(Ref {
            sheet,
            prefix: self.text[start..ref_start].into(),
            first: (first - 1, 0),
            last: Some((last - 1, 16383)),
            fixed: (first_fixed, true),
            last_fixed: (last_fixed, true),
            whole_columns: false,
            whole_rows: true,
            start,
            end: at,
        })))
    }
    fn atom(&mut self) -> Result<Expr> {
        self.space();
        if self.peek() == b'(' {
            self.at += 1;
            let e = self.expr(0)?;
            self.space();
            let mut areas = vec![e];
            while self.peek() == b',' {
                self.at += 1;
                areas.push(self.expr(0)?);
                self.space();
            }
            if self.peek() != b')' {
                return Err(crate::coded(
                    "internal-missing-field",
                    "Missing closing parenthesis",
                ));
            }
            self.at += 1;
            return Ok(if areas.len() == 1 {
                areas.pop().unwrap()
            } else {
                self.nodes += 1;
                if self.nodes > self.node_limit {
                    return Err(crate::coded("formula-limit", "Formula node bound"));
                }
                Expr::Union(areas)
            });
        }
        // Array constants contain no references. Keep them opaque to the
        // evaluator while permitting lossless reference-only transformations.
        if self.peek() == b'{' {
            self.at += 1;
            loop {
                let value = self.expr(0)?;
                fn constant(e: &Expr) -> bool {
                    match e {
                        Expr::Number(_)
                        | Expr::Text
                        | Expr::String(_)
                        | Expr::Boolean(_)
                        | Expr::Error(_) => true,
                        Expr::Unary(op, e) if *op == b'+' || *op == b'-' => constant(e),
                        _ => false,
                    }
                }
                if !constant(&value) {
                    return Err(crate::coded(
                        "formula-unsupported",
                        "Array constants cannot contain references",
                    ));
                }
                self.space();
                match self.peek() {
                    b'}' => {
                        self.at += 1;
                        break;
                    }
                    b',' | b';' => self.at += 1,
                    _ => {
                        return Err(crate::coded(
                            "internal-invalid-field",
                            "Invalid array constant separator",
                        ));
                    }
                }
            }
            return Ok(Expr::Text);
        }
        let start = self.at;
        if self.peek() == b'"' {
            self.at += 1;
            loop {
                match self.peek() {
                    0 => {
                        return Err(crate::coded(
                            "formula-malformed",
                            "Unterminated formula string",
                        ));
                    }
                    b'"' => {
                        self.at += 1;
                        if self.peek() != b'"' {
                            break;
                        }
                    }
                    _ => {}
                }
                self.at += 1;
            }
            return Ok(Expr::String(
                self.text[start + 1..self.at - 1].replace("\"\"", "\""),
            ));
        }
        if self.peek() == b'#' {
            for error in [
                "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#NULL!",
            ] {
                if self.text[self.at..].starts_with(error) {
                    self.at += error.len();
                    return Ok(Expr::Error(error.into()));
                }
            }
            return Err(crate::coded(
                "formula-unsupported",
                "Unsupported formula error",
            ));
        }
        if let Some(reference) = self.whole_rows(start, None, start)? {
            return Ok(reference);
        }
        if self.peek().is_ascii_digit() || self.peek() == b'.' {
            while self.peek().is_ascii_digit() {
                self.at += 1
            }
            if self.peek() == b'.' {
                self.at += 1;
                while self.peek().is_ascii_digit() {
                    self.at += 1
                }
            }
            if matches!(self.peek(), b'e' | b'E') {
                self.at += 1;
                if matches!(self.peek(), b'+' | b'-') {
                    self.at += 1
                }
                while self.peek().is_ascii_digit() {
                    self.at += 1
                }
            }
            let n = self.text[start..self.at].parse::<f64>().map_err(|_| {
                format!(
                    "Unsupported number {} in {}",
                    &self.text[start..self.at],
                    self.text.chars().take(80).collect::<String>()
                )
            })?;
            if !n.is_finite() {
                return Err(crate::coded("formula-unsupported", "Nonfinite number"));
            }
            return Ok(Expr::Number(n));
        }
        // Bracket selectors are opaque during structural rewrites, resolved for calculation.
        let mut probe = self.at;
        while self.text.as_bytes().get(probe).is_some_and(|b| word(*b)) {
            probe += 1;
        }
        if self.text.as_bytes().get(probe) == Some(&b'[') {
            self.at = probe;
            let mut brackets = 0;
            loop {
                match self.peek() {
                    b'[' => brackets += 1,
                    b']' => brackets -= 1,
                    0 => {
                        return Err(crate::coded(
                            "formula-malformed",
                            "Unterminated table selector",
                        ));
                    }
                    _ => {}
                }
                self.at += 1;
                if brackets == 0 {
                    break;
                }
                if brackets > 4 {
                    return Err(crate::coded(
                        "formula-limit",
                        "Table selector nesting bound",
                    ));
                }
            }
            return Ok(Expr::Structured(
                self.text[start..self.at].into(),
                start,
                self.at,
            ));
        }
        let mut sheet = None;
        if self.peek() == b'\'' {
            self.at += 1;
            let mut name = String::new();
            let mut part = self.at;
            loop {
                if self.peek() == 0 {
                    return Err(crate::coded("formula-malformed", "Unterminated sheet name"));
                }
                if self.peek() == b'\'' {
                    name.push_str(&self.text[part..self.at]);
                    self.at += 1;
                    if self.peek() == b'\'' {
                        name.push('\'');
                        self.at += 1;
                        part = self.at;
                        continue;
                    }
                    break;
                }
                self.at += 1;
            }
            if self.peek() != b'!' {
                return Err(crate::coded(
                    "formula-unsupported",
                    "Unsupported quoted token",
                ));
            }
            self.at += 1;
            if name.contains(':') {
                return Err(crate::coded(
                    "formula-bad-input",
                    "External or 3D reference requires compatibility engine",
                ));
            }
            if name.contains(['[', ']']) {
                // Only when adjusting references for an edit; calculation still declines.
                if self.node_limit <= 256 {
                    return Err(crate::coded(
                        "formula-bad-input",
                        "External or 3D reference requires compatibility engine",
                    ));
                }
                // Skip what it points at: a cell, a range, whole rows or columns, or a name.
                while word(self.peek()) || matches!(self.peek(), b'$' | b':' | b'#' | b'!') {
                    self.at += 1;
                }
                return Ok(Expr::External);
            }
            sheet = Some(name);
        } else {
            let begin = self.at;
            while word(self.peek()) {
                self.at += 1
            }
            if self.peek() == b'!' {
                sheet = Some(self.text[begin..self.at].into());
                self.at += 1
            } else {
                self.at = begin
            }
        }
        let ref_start = self.at;
        if sheet.is_some() && self.text[self.at..].starts_with("#REF!") {
            self.at += "#REF!".len();
            return Ok(Expr::QualifiedError(sheet.unwrap(), start, ref_start));
        }
        if let Some(reference) = self.whole_rows(start, sheet.clone(), ref_start)? {
            return Ok(reference);
        }
        // Whole-column references keep the finite XLSX row boundary in storage.
        // The evaluator scans occupied SQLite cells instead of allocating blanks.
        let mut probe = self.at;
        let first_fixed = self.peek() == b'$';
        if first_fixed {
            probe += 1;
        }
        let a = probe;
        while self
            .text
            .as_bytes()
            .get(probe)
            .is_some_and(u8::is_ascii_alphabetic)
        {
            probe += 1;
        }
        if probe > a && self.text.as_bytes().get(probe) == Some(&b':') {
            let first_column = &self.text[a..probe];
            probe += 1;
            let last_fixed = self.text.as_bytes().get(probe) == Some(&b'$');
            if last_fixed {
                probe += 1;
            }
            let b = probe;
            while self
                .text
                .as_bytes()
                .get(probe)
                .is_some_and(u8::is_ascii_alphabetic)
            {
                probe += 1;
            }
            if probe > b && !self.text.as_bytes().get(probe).is_some_and(|c| word(*c)) {
                let first = position(&format!("{}1", first_column.to_ascii_uppercase()))?.1;
                let last = position(&format!("{}1", self.text[b..probe].to_ascii_uppercase()))?.1;
                if first > last || last >= 16384 {
                    return Err(crate::coded(
                        "internal-invalid-field",
                        "Invalid whole-column reference",
                    ));
                }
                self.at = probe;
                return Ok(Expr::Ref(Ref {
                    sheet,
                    prefix: self.text[start..ref_start].into(),
                    first: (0, first),
                    last: Some((1048575, last)),
                    fixed: (true, first_fixed),
                    last_fixed: (true, last_fixed),
                    whole_columns: true,
                    whole_rows: false,
                    start,
                    end: probe,
                }));
            }
        }
        if let Ok((first, fixed)) = self.cell() {
            let mut last = None;
            let mut last_fixed = (false, false);
            if self.peek() == b':' {
                let colon = self.at;
                self.at += 1;
                if let Ok(parsed) = self.cell() {
                    last = Some(parsed.0);
                    last_fixed = parsed.1;
                } else {
                    // Let expr validate the restricted error-range grammar.
                    self.at = colon;
                }
            }
            if word(self.peek()) || self.peek() == b'(' {
                self.at = ref_start
            } else {
                if let Some(b) = last
                    && (first.0 > b.0 || first.1 > b.1)
                {
                    return Err(crate::coded(
                        "formula-malformed",
                        "Reversed reference requires compatibility engine",
                    ));
                }
                return Ok(Expr::Ref(Ref {
                    sheet,
                    prefix: self.text[start..ref_start].into(),
                    first,
                    last,
                    fixed,
                    last_fixed,
                    whole_columns: false,
                    whole_rows: false,
                    start,
                    end: self.at,
                }));
            }
        } else {
            self.at = ref_start
        }
        while word(self.peek()) {
            self.at += 1
        }
        if self.at == ref_start {
            return Err(crate::coded(
                "formula-unsupported",
                "Unsupported formula token",
            ));
        }
        let name = self.text[ref_start..self.at].to_ascii_uppercase();
        if let Some(scope) = sheet {
            return Ok(Expr::QualifiedName(scope, name, start, self.at));
        }
        self.space();
        if self.peek() != b'(' {
            if name == "TRUE" || name == "FALSE" {
                return Ok(Expr::Boolean(name == "TRUE"));
            }
            return Ok(Expr::Name(name, start, self.at));
        }
        self.at += 1;
        let mut args = Vec::new();
        loop {
            self.space();
            if self.peek() == b')' {
                self.at += 1;
                break;
            }
            if self.peek() == b',' {
                args.push(Expr::Text);
                self.at += 1;
                continue;
            }
            args.push(self.expr(0)?);
            self.space();
            if self.peek() == b')' {
                self.at += 1;
                break;
            }
            if self.peek() != b',' {
                return Err(crate::coded(
                    "formula-unsupported",
                    "Unsupported argument separator",
                ));
            }
            self.at += 1;
            if self.peek() == b')' {
                args.push(Expr::Text);
                self.at += 1;
                break;
            }
        }
        Ok(Expr::Call(name, args))
    }
    fn cell(&mut self) -> Result<((u32, u32), (bool, bool))> {
        let col = self.peek() == b'$';
        if col {
            self.at += 1
        }
        let a = self.at;
        while self.peek().is_ascii_alphabetic() {
            self.at += 1
        }
        let letters = &self.text[a..self.at];
        let row = self.peek() == b'$';
        if row {
            self.at += 1
        }
        let b = self.at;
        while self.peek().is_ascii_digit() {
            self.at += 1
        }
        if letters.is_empty() || b == self.at || self.text.as_bytes()[b] == b'0' {
            return Err(crate::coded("formula-bad-input", "Not an A1 reference"));
        }
        let p = position(&format!(
            "{}{}",
            letters.to_ascii_uppercase(),
            &self.text[b..self.at]
        ))?;
        if p.0 >= 1048576 || p.1 >= 16384 {
            return Err(crate::coded(
                "formula-unsupported",
                "Reference outside worksheet",
            ));
        }
        Ok((p, (row, col)))
    }
}
fn supported_function(name: &str) -> bool {
    [
        "SUM",
        "AVERAGE",
        "MIN",
        "MAX",
        "COUNT",
        "COUNTA",
        "ABS",
        "ROUND",
        "IF",
        "IFERROR",
        "INT",
        "SIGN",
        "AND",
        "OR",
        "NOT",
        "PRODUCT",
        "MOD",
        "SQRT",
        "POWER",
        "TRUNC",
        "DATE",
        "YEAR",
        "MONTH",
        "DAY",
        "EDATE",
        "EOMONTH",
        "DEC2HEX",
        "PMT",
        "PV",
        "FV",
        "LEN",
        "ROWS",
        "COLUMNS",
        "COUNTBLANK",
        "LEFT",
        "RIGHT",
        "MID",
        "UPPER",
        "LOWER",
        "TRIM",
        "CONCATENATE",
        "REPT",
        "SUBSTITUTE",
        "EXACT",
        "N",
        "T",
        "ISNUMBER",
        "ISTEXT",
        "ISLOGICAL",
        "ISBLANK",
        // Dependency proof only: evaluation still independently dispatches its
        // native subset. These pure functions cannot address cells outside args.
        "CONCAT",
        "SUMIF",
        "SUMIFS",
        "COUNTIF",
        "COUNTIFS",
        "AVERAGEIF",
        "AVERAGEIFS",
        "ROUNDUP",
        "ROUNDDOWN",
        "VLOOKUP",
        "XLOOKUP",
        "_XLFN.XLOOKUP",
        "HLOOKUP",
        "INDEX",
        "MATCH",
    ]
    .contains(&name)
}
fn visit<'a>(
    expr: &'a Expr,
    refs: &mut Vec<(&'a Ref, bool)>,
    aggregate: bool,
    allow_names: bool,
) -> Result<()> {
    match expr {
        Expr::Ref(r) => refs.push((r, aggregate)),
        Expr::Binary(b':', _, _) => {} // A broken range cannot read any cells.
        Expr::Union(areas) => {
            for area in areas {
                visit(area, refs, aggregate, allow_names)?;
            }
        }
        Expr::Name(..) | Expr::QualifiedName(..) | Expr::Structured(..) if !allow_names => {
            return Err(crate::coded(
                "formula-bad-input",
                "Defined name requires compatibility engine",
            ));
        }
        Expr::Unary(_, e) => visit(e, refs, false, allow_names)?,
        Expr::Binary(_, a, b) | Expr::Compare(_, a, b) => {
            visit(a, refs, false, allow_names)?;
            visit(b, refs, false, allow_names)?
        }
        Expr::Call(name, args) => {
            if ["INDIRECT", "OFFSET"].contains(&name.as_str())
                || (!allow_names && (!supported_function(name) || args.is_empty()))
            {
                return Err(crate::coded(
                    "formula-unsupported",
                    "Dynamic or unsupported function requires compatibility engine",
                ));
            }
            for arg in args {
                visit(
                    arg,
                    refs,
                    ["SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA", "PRODUCT"]
                        .contains(&name.as_str()),
                    allow_names,
                )?
            }
        }
        _ => {}
    }
    Ok(())
}
fn sheet_name<'a>(view: &'a Value, name: &str) -> Result<&'a str> {
    view["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheets"))?
        .iter()
        .filter_map(|s| s["name"].as_str())
        .find(|s| s.eq_ignore_ascii_case(name))
        .ok_or_else(|| crate::coded("internal-unknown-reference", "Unknown reference sheet"))
}
/// Resolve only static absolute names and qualified table columns. The same
/// expansion feeds dependency collection and evaluation, so invalidation agrees.
fn resolved(formula: &str, view: &Value, owner: &str, reference: &str) -> Result<Expr> {
    fn expand(
        expr: Expr,
        view: &Value,
        owner: &str,
        reference: &str,
        scope: Option<&str>,
        depth: usize,
        budget: &mut usize,
    ) -> Result<Expr> {
        crate::cancel::check()?;
        *budget += 1;
        if depth > 32 || *budget > 4096 {
            return Err(crate::coded(
                "formula-limit",
                "Name expansion bound or cycle",
            ));
        }
        Ok(match expr {
            Expr::QualifiedName(context, name, _, _) => {
                let context = sheet_name(view, &context)?;
                // Explicit sheet scope must name a local definition, not fall through globally.
                if !view["namedRanges"].as_array().is_some_and(|names| {
                    names.iter().any(|n| {
                        n["scopeInvalid"] != true
                            && n["name"]
                                .as_str()
                                .is_some_and(|n| n.eq_ignore_ascii_case(&name))
                            && n["sheet"]
                                .as_str()
                                .is_some_and(|s| s.eq_ignore_ascii_case(context))
                    })
                }) {
                    return Err(crate::coded(
                        "internal-unknown-reference",
                        "Unknown sheet-scoped name",
                    ));
                }
                expand(
                    Expr::Name(name, 0, 0),
                    view,
                    context,
                    reference,
                    scope,
                    depth + 1,
                    budget,
                )?
            }
            Expr::Name(name, _, _) => {
                let names = view["namedRanges"].as_array().ok_or_else(|| {
                    crate::coded("internal-unknown-reference", "Unknown defined name")
                })?;
                let matches = |n: &&Value| {
                    n["scopeInvalid"] != true
                        && n["name"]
                            .as_str()
                            .is_some_and(|s| s.eq_ignore_ascii_case(&name))
                };
                let local_count = names
                    .iter()
                    .filter(matches)
                    .filter(|n| {
                        n["sheet"]
                            .as_str()
                            .is_some_and(|s| s.eq_ignore_ascii_case(owner))
                    })
                    .count();
                let global_count = names
                    .iter()
                    .filter(matches)
                    .filter(|n| n["sheet"].is_null())
                    .count();
                if local_count > 1 || local_count == 0 && global_count > 1 {
                    return Err(crate::coded("formula-ambiguous", "Ambiguous defined name"));
                }
                let local = names.iter().filter(matches).find(|n| {
                    n["sheet"]
                        .as_str()
                        .is_some_and(|s| s.eq_ignore_ascii_case(owner))
                });
                let definition = local
                    .or_else(|| names.iter().filter(matches).find(|n| n["sheet"].is_null()))
                    .ok_or_else(|| {
                        crate::coded("internal-unknown-reference", "Unknown defined name")
                    })?;
                let context = definition["sheet"].as_str().unwrap_or("");
                expand(
                    Parser::parse(field(definition, "formula")?)?,
                    view,
                    if context.is_empty() { owner } else { context },
                    reference,
                    Some(context),
                    depth + 1,
                    budget,
                )?
            }
            Expr::Structured(token, _, _) => {
                let split = token.find('[').ok_or_else(|| {
                    crate::coded("internal-invalid-field", "Invalid table reference")
                })?;
                let table_name = &token[..split];
                let spec = &token[split..];
                if table_name.is_empty() || spec.starts_with("[@") {
                    return Err(crate::coded(
                        "formula-bad-input",
                        "Shorthand table syntax requires canonical serialization",
                    ));
                }
                let p = position(reference)?;
                let tables = view["tables"]
                    .as_array()
                    .ok_or_else(|| crate::coded("internal-unknown-reference", "Unknown table"))?;
                let mut candidates = Vec::new();
                for table in tables {
                    if table["headerRowCount"].as_u64().unwrap_or(1) != 1
                        || table["totalsRowCount"].as_u64().unwrap_or(0) != 0
                    {
                        continue;
                    }
                    let Expr::Ref(bounds) = Parser::parse(field(table, "ref")?)? else {
                        return Err(crate::coded(
                            "internal-invalid-field",
                            "Invalid table bounds",
                        ));
                    };
                    let end = bounds.last.unwrap_or(bounds.first);
                    if if table_name.is_empty() {
                        table["sheet"] == owner
                            && p.0 > bounds.first.0
                            && p.0 <= end.0
                            && p.1 >= bounds.first.1
                            && p.1 <= end.1
                    } else {
                        table["name"]
                            .as_str()
                            .is_some_and(|s| s.eq_ignore_ascii_case(table_name))
                    } {
                        candidates.push((table, bounds, end));
                    }
                }
                if candidates.len() != 1 {
                    return Err(crate::coded(
                        "formula-ambiguous",
                        "Ambiguous or unsupported table reference",
                    ));
                }
                let (table, bounds, end) = candidates.pop().unwrap();
                // Escaped/special header spellings and multicolumn selectors stay guarded.
                let flat = spec.replace(['[', ']'], "");
                if flat.contains(['\'', ':']) {
                    return Err(crate::coded(
                        "formula-bad-input",
                        "Complex table selector requires compatibility engine",
                    ));
                }
                let (column, current, all) = if let Some(c) = flat.strip_prefix('@') {
                    (c, true, false)
                } else if let Some(c) = flat.strip_prefix("#This Row,") {
                    (c, true, false)
                } else if let Some(c) = flat.strip_prefix("#Data,") {
                    (c, false, false)
                } else if let Some(c) = flat.strip_prefix("#All,") {
                    (c, false, true)
                } else if flat.starts_with('#') || flat.contains(',') {
                    return Err(crate::coded(
                        "formula-unsupported",
                        "Unsupported table selector",
                    ));
                } else {
                    (flat.as_str(), false, false)
                };
                let columns = table["columns"].as_array().ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing table columns")
                })?;
                let col = columns
                    .iter()
                    .position(|c| {
                        c.as_str()
                            .or_else(|| c["name"].as_str())
                            .is_some_and(|n| n.eq_ignore_ascii_case(column))
                    })
                    .ok_or_else(|| {
                        crate::coded("internal-unknown-reference", "Unknown table column")
                    })? as u32
                    + bounds.first.1;
                let first_row = if current {
                    if table["sheet"] != owner || p.0 <= bounds.first.0 || p.0 > end.0 {
                        return Err(crate::coded(
                            "formula-unsupported",
                            "Table row reference outside its data body",
                        ));
                    }
                    p.0
                } else {
                    bounds.first.0 + u32::from(!all)
                };
                if first_row > end.0 || col > end.1 {
                    return Err(crate::coded(
                        "formula-unsupported",
                        "Empty or inconsistent table range",
                    ));
                }
                Expr::Ref(Ref {
                    sheet: Some(field(table, "sheet")?.into()),
                    prefix: String::new(),
                    first: (first_row, col),
                    last: if current { None } else { Some((end.0, col)) },
                    fixed: (true, true),
                    last_fixed: (true, true),
                    whole_columns: false,
                    whole_rows: false,
                    start: 0,
                    end: 0,
                })
            }
            Expr::Ref(mut r) => {
                if let Some(context) = scope {
                    if r.fixed != (true, true) || r.last.is_some() && r.last_fixed != (true, true) {
                        return Err(crate::coded(
                            "formula-bad-input",
                            "Relative defined name requires compatibility engine",
                        ));
                    }
                    if r.sheet.is_none() {
                        if context.is_empty() {
                            return Err(crate::coded(
                                "formula-unsupported",
                                "Unqualified global name requires compatibility engine",
                            ));
                        }
                        r.sheet = Some(context.into());
                    }
                    // Authored one-cell names use A1:A1. In a scalar formula
                    // they have the same single-cell meaning as A1; leave real
                    // multi-cell rectangles and whole axes unchanged.
                    if !r.whole_rows && !r.whole_columns && r.last == Some(r.first) {
                        r.last = None;
                    }
                }
                Expr::Ref(r)
            }
            Expr::Unary(op, e) => Expr::Unary(
                op,
                Box::new(expand(
                    *e,
                    view,
                    owner,
                    reference,
                    scope,
                    depth + 1,
                    budget,
                )?),
            ),
            Expr::Binary(op, a, b) => Expr::Binary(
                op,
                Box::new(expand(
                    *a,
                    view,
                    owner,
                    reference,
                    scope,
                    depth + 1,
                    budget,
                )?),
                Box::new(expand(
                    *b,
                    view,
                    owner,
                    reference,
                    scope,
                    depth + 1,
                    budget,
                )?),
            ),
            Expr::Compare(op, a, b) => Expr::Compare(
                op,
                Box::new(expand(
                    *a,
                    view,
                    owner,
                    reference,
                    scope,
                    depth + 1,
                    budget,
                )?),
                Box::new(expand(
                    *b,
                    view,
                    owner,
                    reference,
                    scope,
                    depth + 1,
                    budget,
                )?),
            ),
            Expr::Call(name, args) => Expr::Call(
                name,
                args.into_iter()
                    .map(|a| expand(a, view, owner, reference, scope, depth + 1, budget))
                    .collect::<Result<Vec<_>>>()?,
            ),
            Expr::Union(areas) => Expr::Union(
                areas
                    .into_iter()
                    .map(|area| expand(area, view, owner, reference, scope, depth + 1, budget))
                    .collect::<Result<Vec<_>>>()?,
            ),
            other => other,
        })
    }
    expand(
        Parser::parse(formula)?,
        view,
        owner,
        reference,
        None,
        0,
        &mut 0,
    )
}

fn dynamic(expr: &Expr) -> bool {
    match expr {
        Expr::Call(name, args) => {
            ["INDIRECT", "OFFSET"].contains(&name.as_str()) || args.iter().any(dynamic)
        }
        Expr::Unary(_, e) => dynamic(e),
        Expr::Binary(_, a, b) | Expr::Compare(_, a, b) => dynamic(a) || dynamic(b),
        Expr::Union(areas) => areas.iter().any(dynamic),
        _ => false,
    }
}

pub(crate) fn dynamic_cell(view: &Value, sheet: &str, reference: &str, cell: &Value) -> bool {
    cell["formula"]
        .as_str()
        .and_then(|f| resolved(f, view, sheet, reference).ok())
        .is_some_and(|e| dynamic(&e))
}

/// An unchanged persisted result is reusable only when the entire expression
/// resolves against this exact metadata and contains proven deterministic calls.
/// The ordinary dependency visitor rejects unknown, volatile and dynamic calls,
/// including calls hidden in an unselected IF branch or a defined name.
pub(crate) fn cache_eligible(
    view: &Value,
    sheet: &str,
    reference: &str,
    cell: &Value,
) -> Result<()> {
    if cell.get("formulaType").is_some() || cell["scenarioResult"] == true {
        return Err(crate::coded(
            "formula-unsupported",
            "Special formula calculation is not reusable",
        ));
    }
    if cell["calculation"]["status"] != "calculated" {
        return Err(crate::coded(
            "formula-unsupported",
            "Unresolved or failed calculation is not reusable",
        ));
    }
    let text = field(cell, "formula")?;
    let ast = resolved(text, view, sheet, reference)?;
    let calls = financial::calls(text)?;
    if financial::count(&ast) != calls.len() {
        return Err(crate::coded(
            "formula-unsupported",
            format!("{}: financial name expansion", financial::REASON),
        ));
    }
    let mut references = Vec::new();
    visit(&ast, &mut references, false, false)?;
    for (reference, _) in references {
        sheet_name(view, reference.sheet.as_deref().unwrap_or(sheet))?;
    }
    Ok(())
}

fn analyze(view: &Value) -> Result<Value> {
    let mut items = Vec::new();
    let mut ranges = Vec::new();
    for s in view["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheets"))?
    {
        let name = field(s, "name")?;
        for (reference, cell) in s["cells"]
            .as_object()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing cells"))?
        {
            crate::cancel::check()?;
            if cell.get("formulaType").is_some() {
                return Err(crate::coded("formula-unsupported", "Special formula"));
            }
            let Some(formula) = cell["formula"].as_str() else {
                continue;
            };
            let ast = resolved(formula, view, name, reference)?;
            let calls = financial::calls(formula)?;
            if financial::count(&ast) != calls.len() {
                return Err(crate::coded(
                    "formula-unsupported",
                    format!("{}: financial name expansion", financial::REASON),
                ));
            }
            let mut refs = Vec::new();
            visit(&ast, &mut refs, false, false)?;
            let mut deps = Vec::new();
            for (r, _) in refs {
                let sheet = sheet_name(view, r.sheet.as_deref().unwrap_or(name))?;
                let d = if let Some(b) = r.last {
                    json!({"sheet":sheet,"startRow":r.first.0,"endRow":b.0,"startColumn":r.first.1,"endColumn":b.1})
                } else {
                    json!({"sheet":sheet,"ref":address(r.first.0,r.first.1)})
                };
                deps.push(d.clone());
                ranges.push(d);
            }
            items.push(json!({"sheet":name,"ref":reference,"ranges":deps}));
        }
    }
    Ok(json!({"items":items,"scope":{"mode":"static-references","ranges":ranges}}))
}
/// Closed preview admission visits every static branch before evaluating one root.
pub(crate) fn preview_dependencies(
    view: &Value,
    sheet: &str,
    reference: &str,
    formula: &str,
) -> Result<Value> {
    fn admit(expr: &Expr, depth: usize) -> Result<()> {
        if depth > 32 {
            return Err(crate::coded(
                "formula-limit",
                "Draft formula expression depth bound",
            ));
        }
        match expr {
            Expr::Text | Expr::Name(..) | Expr::QualifiedName(..) | Expr::Structured(..) => {
                return Err(crate::coded(
                    "formula-unsupported",
                    "Draft expression is outside the scalar profile",
                ));
            }
            Expr::Call(name, args) => {
                let count = args.len();
                let valid = match name.strip_prefix("_XLFN.").unwrap_or(name.as_str()) {
                    "SUM" | "AVERAGE" | "MIN" | "MAX" | "COUNT" | "COUNTA" | "PRODUCT" | "AND"
                    | "OR" | "CONCAT" | "CONCATENATE" => count >= 1 && count <= 128,
                    "ABS" | "INT" | "SIGN" | "SQRT" | "NOT" | "LEN" | "TRIM" | "UPPER"
                    | "LOWER" | "ISNUMBER" | "ISTEXT" | "ISLOGICAL" | "ISBLANK" | "N" | "T"
                    | "DEC2HEX" | "YEAR" | "MONTH" | "DAY" | "ROWS" | "COLUMNS" | "COUNTBLANK" => {
                        count == 1
                    }
                    "ROUND" | "ROUNDUP" | "ROUNDDOWN" | "MOD" | "POWER" | "EXACT" | "REPT"
                    | "IFERROR" | "EDATE" | "EOMONTH" => count == 2,
                    "IF" => count == 3,
                    "LEFT" | "RIGHT" => count == 1 || count == 2,
                    "MID" | "DATE" => count == 3,
                    "COUNTIF" => count == 2,
                    "SUMIF" => count == 2 || count == 3,
                    "VLOOKUP" => count == 4,
                    "XLOOKUP" => count >= 3 && count <= 6,
                    "SUBSTITUTE" => count == 3 || count == 4,
                    "PMT" | "PV" | "FV" => count >= 3 && count <= 5,
                    _ => false,
                };
                if !valid {
                    return Err(crate::coded(
                        "formula-unsupported",
                        "Draft function or argument count is outside the scalar profile",
                    ));
                }
                for arg in args {
                    admit(arg, depth + 1)?;
                }
            }
            Expr::Unary(_, a) => admit(a, depth + 1)?,
            Expr::Binary(_, a, b) | Expr::Compare(_, a, b) => {
                admit(a, depth + 1)?;
                admit(b, depth + 1)?;
            }
            Expr::Union(areas) => {
                for area in areas {
                    admit(area, depth + 1)?;
                }
            }
            Expr::Ref(r) if r.whole_rows || r.whole_columns => {
                return Err(crate::coded(
                    "formula-limit",
                    "Draft whole-axis references exceed the preview profile",
                ));
            }
            _ => (),
        }
        Ok(())
    }
    fn names(expr: &Expr, view: &Value, owner: &str) -> Result<()> {
        let choose = |name: &str, context: &str, explicit: bool| -> Result<()> {
            let records = view["namedRanges"]
                .as_array()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing preview names"))?;
            let valid = |n: &&Value| {
                n["scopeInvalid"] != true
                    && n["name"]
                        .as_str()
                        .is_some_and(|s| s.eq_ignore_ascii_case(name))
            };
            let local = records
                .iter()
                .filter(valid)
                .filter(|n| {
                    n["sheet"]
                        .as_str()
                        .is_some_and(|s| s.eq_ignore_ascii_case(context))
                })
                .collect::<Vec<_>>();
            let global = records
                .iter()
                .filter(valid)
                .filter(|n| n["sheet"].is_null())
                .collect::<Vec<_>>();
            let candidates = if explicit || !local.is_empty() {
                local
            } else {
                global
            };
            if candidates.len() != 1 {
                return Err(crate::coded(
                    "internal-unknown-reference",
                    "Unknown or ambiguous preview name",
                ));
            }
            fn static_range(e: &Expr) -> bool {
                match e {
                    Expr::Ref(r) => {
                        r.fixed == (true, true)
                            && r.last.is_none_or(|_| r.last_fixed == (true, true))
                            && !r.whole_rows
                            && !r.whole_columns
                    }
                    Expr::Union(areas) => areas.iter().all(static_range),
                    _ => false,
                }
            }
            if !static_range(&Parser::parse(field(candidates[0], "formula")?)?) {
                return Err(crate::coded(
                    "formula-bad-input",
                    "Draft names require static absolute ranges",
                ));
            }
            Ok(())
        };
        match expr {
            Expr::Name(name, ..) => choose(name, owner, false)?,
            Expr::QualifiedName(sheet, name, ..) => choose(name, sheet, true)?,
            Expr::Unary(_, a) => names(a, view, owner)?,
            Expr::Binary(_, a, b) | Expr::Compare(_, a, b) => {
                names(a, view, owner)?;
                names(b, view, owner)?;
            }
            Expr::Call(_, args) | Expr::Union(args) => {
                for a in args {
                    names(a, view, owner)?;
                }
            }
            _ => (),
        }
        Ok(())
    }
    names(&Parser::parse(formula)?, view, sheet)?;
    let ast = resolved(formula, view, sheet, reference)?;
    admit(&ast, 1)?;
    if financial::count(&ast) != financial::calls(formula)?.len() {
        return Err(crate::coded(
            "formula-unsupported",
            "Financial name expansion is outside the preview profile",
        ));
    }
    cell_dependencies(view, sheet, reference, &json!({"formula":formula}))
}

pub(crate) fn cell_dependencies(
    view: &Value,
    sheet: &str,
    reference: &str,
    cell: &Value,
) -> Result<Value> {
    if cell.get("formulaType").is_some() {
        return Err(crate::coded("formula-unsupported", "Special formula"));
    }
    let ast = resolved(field(cell, "formula")?, view, sheet, reference)?;
    // Runtime targets are unknown until the selector has been evaluated. This
    // initial graph remains conservative until successful transactional tracing.
    if dynamic(&ast) {
        let ranges: Vec<Value> = view["sheets"].as_array()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheets"))?.iter().map(|s| json!({"sheet":s["name"],"startRow":0,"endRow":1048575,"startColumn":0,"endColumn":16383})).collect();
        return Ok(json!({"sheet":sheet,"ref":reference,"ranges":ranges}));
    }
    let mut refs = Vec::new();
    visit(&ast, &mut refs, false, false)?;
    let mut ranges = Vec::new();
    for (r, _) in refs {
        let target = sheet_name(view, r.sheet.as_deref().unwrap_or(sheet))?;
        ranges.push(if let Some(b)=r.last{json!({"sheet":target,"startRow":r.first.0,"endRow":b.0,"startColumn":r.first.1,"endColumn":b.1})}else{json!({"sheet":target,"ref":address(r.first.0,r.first.1)})});
    }
    Ok(json!({"sheet":sheet,"ref":reference,"ranges":ranges}))
}
use crate::decimal::{
    add as exact_add, divide as exact_divide, multiply as exact_multiply, number as exact_number,
};
// Deliberately explicit: adding a supported function must not accidentally make
// context-sensitive or volatile expressions reusable across cell coordinates.
fn constant_expression(expr: &Expr) -> bool {
    match expr {
        Expr::Number(_) | Expr::String(_) | Expr::Boolean(_) => true,
        Expr::Unary(_, child) => constant_expression(child),
        Expr::Binary(_, a, b) | Expr::Compare(_, a, b) => {
            constant_expression(a) && constant_expression(b)
        }
        Expr::Call(name, args) => {
            matches!(
                name.as_str(),
                "SUM"
                    | "AVERAGE"
                    | "MIN"
                    | "MAX"
                    | "COUNT"
                    | "COUNTA"
                    | "PRODUCT"
                    | "ABS"
                    | "ROUND"
                    | "ROUNDUP"
                    | "ROUNDDOWN"
                    | "TRUNC"
                    | "INT"
                    | "SIGN"
                    | "MOD"
                    | "POWER"
                    | "SQRT"
                    | "IF"
                    | "IFERROR"
                    | "AND"
                    | "OR"
                    | "NOT"
                    | "TRUE"
                    | "FALSE"
                    | "LEN"
                    | "LOWER"
                    | "UPPER"
                    | "TRIM"
                    | "CONCAT"
                    | "CONCATENATE"
                    | "LEFT"
                    | "RIGHT"
                    | "MID"
                    | "DATE"
                    | "YEAR"
                    | "MONTH"
                    | "DAY"
                    | "EDATE"
                    | "EOMONTH"
                    | "DEC2HEX"
                    | "PMT"
                    | "PV"
                    | "FV"
            ) && args.iter().all(constant_expression)
        }
        _ => false,
    }
}
fn sequence_root(
    view: &Value,
    sheet: &str,
    reference: &str,
    formula: &str,
) -> Result<Option<Vec<Expr>>> {
    let ast = Parser::parse(formula)?;
    let Expr::Call(name, _) = &ast else {
        return Ok(None);
    };
    if !["SEQUENCE", "_XLFN.SEQUENCE"].contains(&name.as_str()) {
        return Ok(None);
    }
    let Expr::Call(_, args) = resolved(formula, view, sheet, reference)? else {
        unreachable!();
    };
    if !(1..=4).contains(&args.len()) {
        return Err(crate::coded(
            "formula-unsupported",
            "SEQUENCE constructor syntax is outside the qualified scalar slice",
        ));
    }
    // Reuse the existing deterministic scalar qualification; extraction still
    // independently requires a numeric result and refuses a direct range.
    // Array/volatile/unknown calls cannot hide in an unselected argument branch.
    for arg in &args {
        visit(arg, &mut Vec::new(), false, false)?;
    }
    Ok(Some(args))
}
pub(crate) fn sequence_dependencies(
    view: &Value,
    sheet: &str,
    reference: &str,
    formula: &str,
) -> Result<Option<Value>> {
    let Some(args) = sequence_root(view, sheet, reference, formula)? else {
        return Ok(None);
    };
    let mut refs = Vec::new();
    for arg in &args {
        visit(arg, &mut refs, false, false)?;
    }
    let mut ranges = Vec::new();
    for (r, _) in refs {
        let target = sheet_name(view, r.sheet.as_deref().unwrap_or(sheet))?;
        let first = address(r.first.0, r.first.1);
        let last = r.last.unwrap_or(r.first);
        let reference = if r.first == last {
            first
        } else {
            format!("{first}:{}", address(last.0, last.1))
        };
        ranges.push(json!({"sheet":target,"ref":reference}));
    }
    Ok(Some(json!({"sheet":sheet,"ref":reference,"ranges":ranges})))
}
pub(crate) struct Eval<'a> {
    db: Option<&'a rusqlite::Connection>,
    selected_only: bool,
    preview_limits: bool,
    fenced_spill_inputs: bool,
    overrides: HashMap<(String, String), Value>,
    view: &'a Value,
    trees: HashMap<(String, String), Expr>,
    memo: HashMap<(String, String), Value>,
    active: HashSet<(String, String)>,
    pub(crate) steps: usize,
    constants: HashMap<String, HashMap<String, Value>>,
    constant_bytes: usize,
    constant_entries: usize,
    pub(crate) constant_hits: u64,
}
impl<'a> Eval<'a> {
    pub(crate) fn fence_spill_inputs(&mut self) {
        self.fenced_spill_inputs = true;
    }
    fn guard_spill_input(&self, sheet: &str, reference: &str) -> Result<()> {
        if self.fenced_spill_inputs
            && let Some(db) = self.db
            && crate::spill::protected_input(db, &self.path(sheet)?, position(reference)?)?
        {
            return Err(crate::coded(
                "formula-unsupported",
                "Constructor input reads an array/spill output without qualified transformation ownership",
            ));
        }
        Ok(())
    }
    /// Spill constructors are evaluated by the separate complete-output owner.
    /// This helper proves numeric scalar arguments without enabling SEQUENCE in
    /// ordinary scalar dispatch or accepting an array's top-left value.
    pub(crate) fn sequence_arguments(
        &mut self,
        sheet: &str,
        reference: &str,
        formula: &str,
    ) -> Result<Option<Vec<f64>>> {
        let Some(args) = sequence_root(self.view, sheet, reference, formula)? else {
            return Ok(None);
        };
        let mut values = Vec::with_capacity(args.len());
        let key = (sheet.to_string(), reference.to_string());
        if !self.active.insert(key.clone()) {
            return Err(crate::coded(
                "formula-refused",
                "Circular spill constructor",
            ));
        }
        let result = (|| {
            for arg in &args {
                let value = self.argument(arg, sheet)?;
                values.push(exact_number(value.as_f64().ok_or_else(|| {
                    crate::coded(
                        "formula-unsupported",
                        "SEQUENCE requires qualified numeric scalar inputs",
                    )
                })?)?);
            }
            Ok(Some(values))
        })();
        self.active.remove(&key);
        result
    }
    /// A Manual source cache can retain an older referenced argument, but a
    /// literal/default argument cannot silently change under the same formula.
    pub(crate) fn sequence_cache_profile(
        &mut self,
        sheet: &str,
        reference: &str,
        formula: &str,
        cached: &[f64],
    ) -> Result<bool> {
        let Some(args) = sequence_root(self.view, sheet, reference, formula)? else {
            return Ok(false);
        };
        fn scalar_syntax(expr: &Expr, ranges: bool) -> bool {
            match expr {
                Expr::Ref(reference) => {
                    ranges || reference.last.is_none_or(|last| last == reference.first)
                }
                Expr::Unary(_, child) => scalar_syntax(child, false),
                Expr::Binary(_, a, b) | Expr::Compare(_, a, b) => {
                    scalar_syntax(a, false) && scalar_syntax(b, false)
                }
                Expr::Call(name, args) => args.iter().all(|arg| {
                    scalar_syntax(
                        arg,
                        ["SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA", "PRODUCT"]
                            .contains(&name.as_str()),
                    )
                }),
                Expr::Union(areas) => ranges && areas.iter().all(|area| scalar_syntax(area, true)),
                _ => true,
            }
        }
        for index in 0..4 {
            if let Some(arg) = args.get(index) {
                if !scalar_syntax(arg, false) {
                    return Ok(false);
                }
                if constant_expression(arg) {
                    let Some(value) = self.argument(arg, sheet)?.as_f64() else {
                        return Ok(false);
                    };
                    if value.to_bits() != cached[index].to_bits() {
                        return Ok(false);
                    }
                }
            } else if cached[index] != 1.0 {
                return Ok(false);
            }
        }
        Ok(true)
    }
    pub(crate) fn from_sqlite(
        view: &'a Value,
        db: &'a rusqlite::Connection,
        selected_only: bool,
    ) -> Self {
        Self {
            view,
            db: Some(db),
            selected_only,
            preview_limits: false,
            fenced_spill_inputs: false,
            overrides: HashMap::new(),
            trees: HashMap::new(),
            memo: HashMap::new(),
            active: HashSet::new(),
            steps: 0,
            constants: HashMap::new(),
            constant_bytes: 0,
            constant_entries: 0,
            constant_hits: 0,
        }
    }
    pub(crate) fn preview_limits(&mut self) {
        self.preview_limits = true;
    }
    pub(crate) fn constant_cache_size(&self) -> (usize, usize) {
        (self.constant_entries, self.constant_bytes)
    }
    fn path(&self, sheet: &str) -> Result<String> {
        Ok(field(
            self.view["sheets"]
                .as_array()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheets"))?
                .iter()
                .find(|s| s["name"] == sheet)
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheet"))?,
            "path",
        )?
        .to_string())
    }

    fn scalar_number(&self, n: f64) -> Result<f64> {
        if self.preview_limits && n.is_finite() {
            Ok(n)
        } else {
            exact_number(n)
        }
    }
    fn tick(&mut self) -> Result<()> {
        crate::cancel::check()?;
        self.steps += 1;
        if self.steps
            > if self.preview_limits {
                200_000
            } else if self.db.is_some() {
                50_000_000
            } else {
                200000
            }
            || self.active.len() > if self.preview_limits { 32 } else { 128 }
        {
            return Err(crate::coded(
                "formula-limit",
                "Native calculation work bound",
            ));
        }
        Ok(())
    }
    pub(crate) fn cell(&mut self, sheet: &str, reference: &str) -> Result<Option<f64>> {
        let value = self.result(sheet, reference)?;
        if value.is_null() {
            Ok(None)
        } else {
            Ok(Some(exact_number(value.as_f64().ok_or_else(|| {
                crate::coded(
                    "formula-bad-input",
                    "Input coercion requires compatibility engine",
                )
            })?)?))
        }
    }
    pub(crate) fn result(&mut self, sheet: &str, reference: &str) -> Result<Value> {
        self.guard_spill_input(sheet, reference)?;
        if let Some(value) = self.overrides.get(&(sheet.into(), reference.into())) {
            return Ok(value.clone());
        }
        if self.overrides.is_empty()
            && let Some(db) = self.db
        {
            let path = self.path(sheet)?;
            if let Some(value) = crate::calculation::memo(db, &path, reference)? {
                return Ok(value);
            }
        }
        if let Some(db) = self.db {
            use rusqlite::OptionalExtension;
            let path = self.path(sheet)?;
            let (r, c) = position(reference)?;
            let scenario:Option<crate::scenarios::Inputs>=db.prepare_cached("SELECT model,input1,value1,input2,value2 FROM calc_scenarios WHERE sheet=?1 AND r=?2 AND c=?3").map_err(|e|crate::rephrase("formula-refused", e, |said| format!("SQLite calculation: {said}")))?.query_row(rusqlite::params![path,r,c],|row|Ok((row.get(0)?,row.get(1)?,row.get(2)?,row.get(3)?,row.get(4)?))).optional().map_err(|e|crate::rephrase("formula-refused", e, |said| format!("SQLite calculation: {said}")))?;
            if let Some((model, input1, value1, input2, value2)) = scenario {
                if !self.overrides.is_empty() {
                    return Err(crate::coded(
                        "formula-unsupported",
                        "Nested What-If scenarios require compatibility",
                    ));
                }
                let key = (sheet.to_string(), reference.to_string());
                if !self.active.insert(key.clone()) {
                    return Err(crate::coded("formula-refused", "Circular What-If scenario"));
                }
                let result = (|| -> Result<Value> {
                    let first = self.result(sheet, &value1)?;
                    if !first.is_number() {
                        return Err(crate::coded(
                            "formula-bad-input",
                            "What-If inputs require numeric values",
                        ));
                    }
                    let second = if let Some(value2) = value2 {
                        let value = self.result(sheet, &value2)?;
                        if !value.is_number() {
                            return Err(crate::coded(
                                "formula-bad-input",
                                "What-If inputs require numeric values",
                            ));
                        }
                        Some(value)
                    } else {
                        None
                    };
                    self.overrides.insert((sheet.into(), input1), first);
                    if let (Some(input), Some(value)) = (input2, second) {
                        self.overrides.insert((sheet.into(), input), value);
                    }
                    let result = self.result(sheet, &model)?;
                    if !result.is_number() {
                        return Err(crate::coded(
                            "formula-bad-input",
                            "What-If outputs require numeric values",
                        ));
                    }
                    Ok(result)
                })();
                self.overrides.clear();
                self.active.remove(&key);
                let result = result.map_err(|e| {
                    crate::rephrase("formula-unsupported", e, |said| {
                        format!("What-If scenario calculation cannot publish: {said}")
                    })
                })?;
                crate::calculation::remember(db, &path, reference, &result)?;
                return Ok(result);
            }
        }
        let cell = if let Some(db) = self.db {
            crate::calculation::evaluation_cell(db, &self.path(sheet)?, reference)?
        } else {
            self.view["sheets"]
                .as_array()
                .unwrap()
                .iter()
                .find(|s| s["name"] == sheet)
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheet"))?["cells"]
                [reference]
                .clone()
        };
        self.typed_cell_value(sheet, reference, &cell)
    }
    fn typed_cell_value(&mut self, sheet: &str, reference: &str, cell: &Value) -> Result<Value> {
        self.guard_spill_input(sheet, reference)?;
        self.tick()?;
        if self.selected_only
            && self.overrides.is_empty()
            && let Some(db) = self.db
        {
            let path = self.path(sheet)?;
            let formula = cell.get("formulaType").is_none() && cell["formula"].is_string();
            let shared = cell["formulaType"] == "shared"
                && crate::calculation::proved_shared(db, &path, cell)?;
            if formula || shared {
                if let Some(value) = crate::calculation::memo(db, &path, reference)? {
                    return Ok(value);
                }
                if !crate::calculation::selected(db, &path, reference)? {
                    if cell["calculation"]["status"] != "calculated" {
                        return Err(crate::coded(
                            "formula-unsupported",
                            "Unreliable upstream formula",
                        ));
                    }
                    let value = cell["calculation"]["value"].clone();
                    if value.is_number() {
                        exact_number(value.as_f64().unwrap())?;
                    }
                    return Ok(value);
                }
            }
        }
        let expanded;
        let cell = if let Some(db) = self.db {
            expanded = crate::calculation::resolved_cell(db, &self.path(sheet)?, reference, cell)?;
            &expanded
        } else {
            cell
        };
        if cell.get("formulaType").is_some() {
            return Err(crate::coded("formula-unsupported", "Special formula"));
        }
        if let Some(formula) = cell["formula"].as_str() {
            let key = (sheet.to_string(), reference.to_string());
            if self.overrides.is_empty()
                && let Some(db) = self.db
            {
                let path = self.path(sheet)?;
                if let Some(value) = crate::calculation::memo(db, &path, reference)? {
                    return Ok(value);
                }
                if self.selected_only && !crate::calculation::selected(db, &path, reference)? {
                    if cell["calculation"]["status"] != "calculated" {
                        return Err(crate::coded(
                            "formula-unsupported",
                            "Unreliable upstream formula",
                        ));
                    }
                    let value = cell["calculation"]["value"].clone();
                    if value.is_number() {
                        exact_number(value.as_f64().unwrap())?;
                    }
                    return Ok(value);
                }
            } else if let Some(n) = self.memo.get(&key) {
                return Ok(n.clone());
            }
            if self.overrides.is_empty()
                && let Some(db) = self.db
                && let Some(value) = self
                    .constants
                    .get(sheet)
                    .and_then(|items| items.get(formula))
                    .cloned()
            {
                crate::calculation::remember(db, &self.path(sheet)?, reference, &value)?;
                self.constant_hits += 1;
                return Ok(value);
            }
            if !self.active.insert(key.clone()) {
                return Err(crate::coded(
                    "formula-bad-input",
                    "Circular formula requires compatibility engine",
                ));
            }
            let evaluated: Result<(Expr, Value, bool)> = (|| {
                let expr = if let Some(expr) = self.trees.remove(&key) {
                    expr
                } else {
                    resolved(formula, self.view, sheet, reference)?
                };
                let mut traced = Vec::new();
                let has_dynamic = dynamic(&expr);
                let cacheable = !has_dynamic && constant_expression(&expr);
                if has_dynamic && !self.overrides.is_empty() {
                    return Err(crate::coded(
                        "formula-bad-input",
                        "Dynamic What-If models require compatibility",
                    ));
                }
                let expr = if has_dynamic {
                    self.runtime_expr(expr, sheet, &mut traced)?
                } else {
                    expr
                };
                self.financial_preflight(formula, sheet, &expr)?;
                // Qualification covers every branch, not only the branch an
                // IF happens to execute on this revision. Dynamic references
                // and unknown functions cannot hide inside a native closure.
                visit(&expr, &mut Vec::new(), false, false)?;
                let n = if root_add_sub(formula)
                    && let Expr::Binary(op @ (b'+' | b'-'), a, b) = &expr
                {
                    let (a, b) = self.operator_pair(a, b, *op, sheet)?;
                    json!(exact_number(compensated(
                        a,
                        if *op == b'-' { -b } else { b }
                    ))?)
                } else {
                    self.scalar(&expr, sheet)?
                };
                if let Some(text) = n.as_str() {
                    Self::bounded_text(text.to_string())?;
                }
                if has_dynamic && let Some(db) = self.db {
                    for r in traced {
                        let target = sheet_name(self.view, r.sheet.as_deref().unwrap_or(sheet))?;
                        let end = r.last.unwrap_or(r.first);
                        db.execute(
                            "INSERT INTO calc_traces VALUES(?1,?2,?3,?4,?5,?6,?7)",
                            rusqlite::params![
                                sheet, reference, target, r.first.0, end.0, r.first.1, end.1
                            ],
                        )
                        .map_err(|e| {
                            crate::rephrase("formula-refused", e, |said| {
                                format!("SQLite calculation: {said}")
                            })
                        })?;
                    }
                }
                Ok((expr, n, cacheable))
            })();
            self.active.remove(&key);
            let (expr, n, cacheable) = evaluated?;
            if self.overrides.is_empty()
                && let Some(db) = self.db
            {
                crate::calculation::remember(db, &self.path(sheet)?, reference, &n)?;
                if cacheable && self.constant_entries < 256 {
                    let bytes = formula.len() + sheet.len() + n.to_string().len() + 128;
                    if self.constant_bytes + bytes <= 64 * 1024 {
                        self.constants
                            .entry(sheet.into())
                            .or_default()
                            .insert(formula.into(), n.clone());
                        self.constant_entries += 1;
                        self.constant_bytes += bytes;
                    }
                }
            } else if self.db.is_none() {
                self.trees.insert(key.clone(), expr);
                self.memo.insert(key, n.clone());
            }
            return Ok(n);
        }
        if cell["dataType"] == "error" || cell.get("error").is_some() {
            if self.preview_limits
                && let Some(code) = cell["error"].as_str().or_else(|| cell["value"].as_str())
                && crate::preview::error_code(code)
            {
                return Err(format!("Excel error: {code}"));
            }
            return Err(crate::coded(
                "formula-bad-input",
                "Error input requires compatibility engine",
            ));
        }
        let value = &cell["value"];
        if !(value.is_null() || value.is_string() || value.is_boolean() || value.is_number()) {
            return Err(crate::coded(
                "formula-bad-input",
                "Literal requires compatibility engine",
            ));
        }
        if let Some(n) = value.as_f64() {
            crate::bulk::literal(cell, false)?;
            exact_number(n)?;
        }
        Ok(value.clone())
    }
    /// Resolve all branches before evaluating. Selector inputs and actual
    /// targets both become dependencies, including currently unselected IF arms.
    fn runtime_expr(&mut self, expr: Expr, sheet: &str, traced: &mut Vec<Ref>) -> Result<Expr> {
        self.tick()?;
        Ok(match expr {
            Expr::Ref(r) => {
                traced.push(r.clone());
                Expr::Ref(r)
            }
            Expr::Binary(b':', a, b) => Expr::Binary(b':', a, b),
            Expr::Unary(op, e) => Expr::Unary(op, Box::new(self.runtime_expr(*e, sheet, traced)?)),
            Expr::Binary(op, a, b) => Expr::Binary(
                op,
                Box::new(self.runtime_expr(*a, sheet, traced)?),
                Box::new(self.runtime_expr(*b, sheet, traced)?),
            ),
            Expr::Compare(op, a, b) => Expr::Compare(
                op,
                Box::new(self.runtime_expr(*a, sheet, traced)?),
                Box::new(self.runtime_expr(*b, sheet, traced)?),
            ),
            Expr::Call(name, args) => {
                let args = args
                    .into_iter()
                    .map(|e| self.runtime_expr(e, sheet, traced))
                    .collect::<Result<Vec<_>>>()?;
                if name == "INDIRECT" || name == "OFFSET" {
                    for arg in &args {
                        visit(arg, &mut Vec::new(), false, false)?;
                    }
                }
                let target = if name == "INDIRECT" {
                    if !(args.len() == 1
                        || args.len() == 2 && matches!(args[1], Expr::Boolean(true)))
                    {
                        return Err(crate::coded(
                            "formula-unsupported",
                            "INDIRECT requires qualified A1 mode",
                        ));
                    }
                    let text = self.text_argument(&args[0], sheet)?;
                    let Expr::Ref(r) = Parser::parse(&text)? else {
                        return Err(crate::coded(
                            "formula-bad-input",
                            "INDIRECT target must be an A1 reference",
                        ));
                    };
                    r
                } else if name == "OFFSET" {
                    if !(3..=5).contains(&args.len()) {
                        return Err(crate::coded("formula-refused", "OFFSET argument count"));
                    }
                    let Expr::Ref(base) = &args[0] else {
                        return Err(crate::coded(
                            "formula-bad-input",
                            "OFFSET base must be a reference",
                        ));
                    };
                    let end = base.last.unwrap_or(base.first);
                    let mut integer = |e: &Expr| -> Result<i64> {
                        let n = self.eval(e, sheet)?;
                        if n.abs() > 1048576.0 {
                            return Err(crate::coded("formula-limit", "OFFSET displacement bound"));
                        }
                        Ok(n.trunc() as i64)
                    };
                    let row = i64::from(base.first.0) + integer(&args[1])?;
                    let col = i64::from(base.first.1) + integer(&args[2])?;
                    let height = if args.len() > 3 {
                        integer(&args[3])?
                    } else {
                        i64::from(end.0 - base.first.0) + 1
                    };
                    let width = if args.len() > 4 {
                        integer(&args[4])?
                    } else {
                        i64::from(end.1 - base.first.1) + 1
                    };
                    if row < 0
                        || col < 0
                        || height < 1
                        || width < 1
                        || row + height > 1048576
                        || col + width > 16384
                    {
                        return Err(crate::coded(
                            "formula-limit",
                            "OFFSET target outside qualified worksheet bounds",
                        ));
                    }
                    let mut r = base.clone();
                    r.first = (row as u32, col as u32);
                    r.last = if height == 1 && width == 1 {
                        None
                    } else {
                        Some(((row + height - 1) as u32, (col + width - 1) as u32))
                    };
                    r.whole_columns = false;
                    r.whole_rows = false;
                    r
                } else {
                    return Ok(Expr::Call(name, args));
                };
                sheet_name(self.view, target.sheet.as_deref().unwrap_or(sheet))?;
                traced.push(target.clone());
                Expr::Ref(target)
            }
            other => other,
        })
    }
    fn predicate(&mut self, e: &Expr, sheet: &str) -> Result<bool> {
        match e {
            Expr::Boolean(b) => Ok(*b),
            Expr::Call(name, args) if (name == "AND" || name == "OR") && !args.is_empty() => {
                let mut result = name == "AND";
                for arg in args {
                    let value = self.predicate(arg, sheet)?;
                    if name == "AND" {
                        result &= value;
                    } else {
                        result |= value;
                    }
                }
                Ok(result)
            }
            Expr::Call(name, args) if name == "NOT" && args.len() == 1 => {
                Ok(!self.predicate(&args[0], sheet)?)
            }
            Expr::Compare(op, a, b) => {
                let av = self.scalar(a, sheet)?;
                let bv = self.scalar(b, sheet)?;
                if let (Some(a), Some(b)) = (av.as_str(), bv.as_str()) {
                    if !a.is_ascii() || !b.is_ascii() || !matches!(op.as_str(), "=" | "<>") {
                        return Err(crate::coded(
                            "formula-bad-input",
                            "Text comparison requires compatibility engine",
                        ));
                    }
                    let equal = a.eq_ignore_ascii_case(b);
                    return Ok(if op == "=" { equal } else { !equal });
                }
                let a = av.as_f64().ok_or_else(|| {
                    crate::coded(
                        "formula-bad-input",
                        "Comparison coercion requires compatibility engine",
                    )
                })?;
                let b = bv.as_f64().ok_or_else(|| {
                    crate::coded(
                        "formula-bad-input",
                        "Comparison coercion requires compatibility engine",
                    )
                })?;
                Ok(match op.as_str() {
                    "=" => a == b,
                    "<>" => a != b,
                    "<" => a < b,
                    ">" => a > b,
                    "<=" => a <= b,
                    ">=" => a >= b,
                    _ => {
                        return Err(crate::coded(
                            "formula-unsupported",
                            "Unsupported comparison",
                        ));
                    }
                })
            }
            _ => {
                let v = self.scalar(e, sheet)?;
                if let Some(b) = v.as_bool() {
                    Ok(b)
                } else {
                    Ok(v.as_f64().ok_or_else(|| {
                        crate::coded(
                            "formula-bad-input",
                            "Predicate coercion requires compatibility engine",
                        )
                    })? != 0.0)
                }
            }
        }
    }
    fn count_blank(&mut self, r: &Ref, sheet: &str) -> Result<f64> {
        let target = sheet_name(self.view, r.sheet.as_deref().unwrap_or(sheet))?.to_string();
        let end = r.last.unwrap_or(r.first);
        let mut blanks = f64::from(end.0 - r.first.0 + 1) * f64::from(end.1 - r.first.1 + 1);
        let view = self.view;
        let db = self.db;
        let mut accept = |reference: &str, cell: &Value| -> Result<()> {
            self.guard_spill_input(&target, reference)?;
            self.tick()?;
            let nonblank = if self
                .overrides
                .contains_key(&(target.clone(), reference.into()))
            {
                true
            } else if cell["formula"].is_string() || cell.get("formulaType").is_some() {
                {
                    let v = self.typed_cell_value(&target, reference, cell)?;
                    !v.is_null() && v.as_str() != Some("")
                }
            } else {
                !cell["value"].is_null() && cell["value"].as_str() != Some("")
            };
            if nonblank {
                blanks -= 1.0;
            }
            Ok(())
        };
        if let Some(db) = db {
            let path = view["sheets"]
                .as_array()
                .unwrap()
                .iter()
                .find(|s| s["name"] == target)
                .unwrap()["path"]
                .as_str()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheet path"))?;
            let mut after = None;
            loop {
                let page = crate::calculation::input_page(db, path, r.first, end, after, false)?;
                if page.is_empty() {
                    break;
                }
                for (reference, cell) in &page {
                    accept(reference, cell)?;
                }
                after = Some(position(&page.last().unwrap().0)?);
            }
        } else {
            // Sparse projection: absent positions are already included in area.
            let cells = view["sheets"]
                .as_array()
                .unwrap()
                .iter()
                .find(|s| s["name"] == target)
                .unwrap()["cells"]
                .as_object()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing cells"))?;
            for (reference, cell) in cells {
                let p = position(reference)?;
                if p.0 >= r.first.0 && p.0 <= end.0 && p.1 >= r.first.1 && p.1 <= end.1 {
                    accept(reference, cell)?;
                }
            }
        }
        Ok(blanks)
    }
    fn aggregate(&mut self, name: &str, args: &[Expr], sheet: &str) -> Result<f64> {
        let mut total = if name == "PRODUCT" { 1.0 } else { 0.0 };
        let mut count = 0_u32;
        let mut low = f64::INFINITY;
        let mut high = f64::NEG_INFINITY;
        // Preserve multi-area source order and repeated/overlapping references.
        // SUM already uses one continuous binary64 accumulator across arguments.
        fn areas<'e>(expr: &'e Expr, out: &mut Vec<&'e Expr>) {
            if let Expr::Union(items) = expr {
                for item in items {
                    areas(item, out);
                }
            } else {
                out.push(expr);
            }
        }
        let has_union = args.iter().any(|arg| matches!(arg, Expr::Union(_)));
        let mut arguments = Vec::new();
        if has_union {
            for arg in args {
                areas(arg, &mut arguments);
            }
        }
        let mut ordinary = args.iter();
        let mut expanded = arguments.into_iter();
        while let Some(arg) = if has_union {
            expanded.next()
        } else {
            ordinary.next()
        } {
            let mut subtotal = if name == "PRODUCT" { 1.0 } else { 0.0 };
            let mut subcount = 0_u32;
            let mut sublow = f64::INFINITY;
            let mut subhigh = f64::NEG_INFINITY;
            let mut accept = |n: f64| -> Result<()> {
                if name == "SUM" {
                    total += n;
                }
                if name == "AVERAGE" {
                    subtotal = exact_add(subtotal, n)?;
                }
                if name == "PRODUCT" {
                    subtotal = exact_multiply(subtotal, n)?;
                }
                subcount += 1;
                sublow = sublow.min(n);
                subhigh = subhigh.max(n);
                Ok(())
            };
            if let Expr::Ref(r) = arg {
                let target =
                    sheet_name(self.view, r.sheet.as_deref().unwrap_or(sheet))?.to_string();
                let end = r.last.unwrap_or(r.first);
                if let Some(db) = self.db {
                    let path = self.path(&target)?;
                    let mut after = None;
                    loop {
                        let page =
                            crate::calculation::input_page(db, &path, r.first, end, after, false)?;
                        if page.is_empty() {
                            break;
                        }
                        for (reference, cell) in &page {
                            if let Some(n) = self.aggregate_cell(name, &target, reference, cell)? {
                                accept(n)?
                            }
                        }
                        after = Some(position(&page.last().unwrap().0)?);
                    }
                } else {
                    for row in r.first.0..=end.0 {
                        for col in r.first.1..=end.1 {
                            let reference = address(row, col);
                            let cell = self.view["sheets"]
                                .as_array()
                                .unwrap()
                                .iter()
                                .find(|s| s["name"] == target)
                                .unwrap()["cells"][&reference]
                                .clone();
                            if let Some(n) =
                                self.aggregate_cell(name, &target, &reference, &cell)?
                            {
                                accept(n)?
                            }
                        }
                    }
                }
            } else if name == "COUNTA"
                && matches!(
                    arg,
                    Expr::String(_) | Expr::Boolean(_) | Expr::Error(_) | Expr::QualifiedError(..)
                )
            {
                accept(1.0)?;
            } else {
                accept(self.eval(arg, sheet)?)?
            }
            if name == "AVERAGE" {
                total = exact_add(total, subtotal)?;
            }
            if name == "PRODUCT" && subcount > 0 {
                total = exact_multiply(total, subtotal)?;
            }
            count += subcount;
            low = low.min(sublow);
            high = high.max(subhigh);
        }
        Ok(match name {
            "SUM" => total,
            "PRODUCT" => {
                if count == 0 {
                    0.0
                } else {
                    total
                }
            }
            "COUNT" | "COUNTA" => f64::from(count),
            "AVERAGE" => {
                if count == 0 {
                    return Err("Excel error: #DIV/0!".into());
                }
                exact_divide(total, f64::from(count))?
            }
            "MIN" => {
                if count == 0 {
                    0.0
                } else {
                    low
                }
            }
            "MAX" => {
                if count == 0 {
                    0.0
                } else {
                    high
                }
            }
            _ => unreachable!(),
        })
    }
    fn aggregate_cell(
        &mut self,
        function: &str,
        sheet: &str,
        reference: &str,
        cell: &Value,
    ) -> Result<Option<f64>> {
        self.guard_spill_input(sheet, reference)?;
        if let Some(value) = self
            .overrides
            .get(&(sheet.into(), reference.into()))
            .cloned()
        {
            return Ok(Some(if ["COUNT", "COUNTA"].contains(&function) {
                1.0
            } else {
                value.as_f64().ok_or_else(|| {
                    crate::coded("formula-refused", "Numeric scenario input required")
                })?
            }));
        }
        if cell["scenarioResult"] == true || cell["formulaType"] == "dataTable" {
            let value = self.result(sheet, reference)?;
            return Ok(Some(if ["COUNT", "COUNTA"].contains(&function) {
                1.0
            } else {
                value.as_f64().ok_or_else(|| {
                    crate::coded("formula-refused", "Numeric scenario output required")
                })?
            }));
        }
        if !cell["formula"].is_string() && cell.get("formulaType").is_none() {
            if function == "COUNTA" {
                self.tick()?;
                return Ok((!cell["value"].is_null() || cell["dataType"] == "error").then_some(1.0));
            }
            if cell["dataType"] == "error" || cell.get("error").is_some() {
                if function == "COUNT" {
                    return Ok(None);
                }
                if self.preview_limits
                    && let Some(code) = cell["error"].as_str().or_else(|| cell["value"].as_str())
                    && crate::preview::error_code(code)
                {
                    return Err(format!("Excel error: {code}"));
                }
                return Err(crate::coded(
                    "formula-bad-input",
                    "Error input requires compatibility engine",
                ));
            }
            if cell["value"].is_string() || cell["value"].is_boolean() {
                self.tick()?;
                return Ok(None);
            }
        }
        let value = self.typed_cell_value(sheet, reference, cell)?;
        if function == "COUNTA" {
            return Ok((!value.is_null()).then_some(1.0));
        }
        Ok(value.as_f64())
    }
    fn argument(&mut self, e: &Expr, sheet: &str) -> Result<Value> {
        if let Expr::Ref(r) = e {
            if r.last.is_some() {
                return Err(crate::coded(
                    "formula-bad-input",
                    "Scalar argument range requires compatibility",
                ));
            }
            let name = sheet_name(self.view, r.sheet.as_deref().unwrap_or(sheet))?.to_string();
            self.result(&name, &address(r.first.0, r.first.1))
        } else {
            self.scalar(e, sheet)
        }
    }
    fn common_range(&mut self, expr: &Expr, sheet: &str) -> Result<(Vec<Value>, (u32, u32))> {
        let Expr::Ref(r) = expr else {
            return Err(crate::coded(
                "formula-bad-input",
                "Common function requires a direct bounded range",
            ));
        };
        let last = r.last.unwrap_or(r.first);
        if r.whole_columns || r.whole_rows || last.0 < r.first.0 || last.1 < r.first.1 {
            return Err(crate::coded(
                "formula-unsupported",
                "Common range is outside its bounded profile",
            ));
        }
        let rows = last.0 - r.first.0 + 1;
        let columns = last.1 - r.first.1 + 1;
        if rows as u64 * columns as u64 > 1000 {
            return Err(crate::coded(
                "formula-limit",
                "Common function range exceeds1000 cells",
            ));
        }
        let owner = sheet_name(self.view, r.sheet.as_deref().unwrap_or(sheet))?.to_owned();
        let mut values = Vec::new();
        for row in r.first.0..=last.0 {
            for col in r.first.1..=last.1 {
                self.tick()?;
                let value = match self.result(&owner, &address(row, col)) {
                    Ok(value) => value,
                    Err(error) if error.starts_with("Excel error: ") => {
                        json!({"$error":error.trim_start_matches("Excel error: ")})
                    }
                    Err(error) => return Err(error),
                };
                values.push(value);
            }
        }
        Ok((values, (rows, columns)))
    }
    fn common_function(&mut self, name: &str, args: &[Expr], sheet: &str) -> Result<Value> {
        if name == "COUNTIF" || name == "SUMIF" {
            let (criteria, shape) = self.common_range(&args[0], sheet)?;
            let wanted = self.argument(&args[1], sheet)?;
            let sums = if name == "SUMIF" {
                if args.len() == 3 {
                    let (values, other) = self.common_range(&args[2], sheet)?;
                    if shape != other {
                        return Err(crate::coded(
                            "formula-unsupported",
                            "SUMIF sum range shape requires wider qualification",
                        ));
                    }
                    values
                } else {
                    criteria.clone()
                }
            } else {
                Vec::new()
            };
            let mut total = 0.0;
            for (index, value) in criteria.iter().enumerate() {
                if common::criterion(&wanted, value)? {
                    if name == "COUNTIF" {
                        total += 1.0;
                    } else {
                        let value = common::selected(sums[index].clone())?;
                        if let Some(number) = value.as_f64() {
                            total += number;
                        }
                    }
                }
            }
            if !total.is_finite() {
                return Err(crate::coded(
                    "formula-unsupported",
                    "Nonfinite common aggregation result",
                ));
            }
            return Ok(json!(total));
        }
        let key = self.argument(&args[0], sheet)?;
        if name == "VLOOKUP" {
            let (table, shape) = self.common_range(&args[1], sheet)?;
            let column = self.argument(&args[2], sheet)?.as_f64().ok_or_else(|| {
                crate::coded("formula-bad-input", "Lookup column requires a number")
            })?;
            let mode = self.argument(&args[3], sheet)?;
            if mode.as_bool() != Some(false) && mode.as_f64() != Some(0.0) {
                return Err(crate::coded(
                    "formula-unsupported",
                    "Only exact VLOOKUP is qualified natively",
                ));
            }
            if column.fract() != 0.0 {
                return Err(crate::coded(
                    "formula-bad-input",
                    "VLOOKUP return column requires a whole-number profile",
                ));
            }
            if column < 1.0 || column > shape.1 as f64 {
                return Err("Excel error: #REF!".into());
            }
            for row in 0..shape.0 as usize {
                if common::lookup_equal(&key, &table[row * shape.1 as usize]) {
                    return common::selected(
                        table[row * shape.1 as usize + column as usize - 1].clone(),
                    );
                }
            }
        } else {
            let (keys, shape) = self.common_range(&args[1], sheet)?;
            let (returns, other) = self.common_range(&args[2], sheet)?;
            if shape != other || shape.0 != 1 && shape.1 != 1 {
                return Err(crate::coded(
                    "formula-bad-input",
                    "XLOOKUP needs matching one-row/column vectors",
                ));
            }
            if args.len() >= 5 && self.argument(&args[4], sheet)?.as_f64() != Some(0.0) {
                return Err(crate::coded(
                    "formula-unsupported",
                    "Only exact XLOOKUP is qualified natively",
                ));
            }
            let mode = if args.len() == 6 {
                self.argument(&args[5], sheet)?.as_f64().ok_or_else(|| {
                    crate::coded("formula-bad-input", "Lookup search mode must be numeric")
                })?
            } else {
                1.0
            };
            if mode != 1.0 && mode != -1.0 {
                return Err(crate::coded(
                    "formula-unsupported",
                    "Only linear lookup search is qualified",
                ));
            }
            let found = if mode == 1.0 {
                keys.iter()
                    .position(|value| common::lookup_equal(&key, value))
            } else {
                keys.iter()
                    .rposition(|value| common::lookup_equal(&key, value))
            };
            if let Some(index) = found {
                return common::selected(returns[index].clone());
            }
            if args.len() >= 4 {
                return self.argument(&args[3], sheet);
            }
        }
        Err("Excel error: #N/A".into())
    }
    fn scalar(&mut self, e: &Expr, sheet: &str) -> Result<Value> {
        self.tick()?;
        Ok(match e {
            Expr::String(text) => json!(text),
            Expr::Boolean(b) => json!(b),
            Expr::Compare(..) => json!(self.predicate(e, sheet)?),
            Expr::Ref(r) if r.last.is_none() => {
                let name = sheet_name(self.view, r.sheet.as_deref().unwrap_or(sheet))?.to_string();
                let value = self.result(&name, &address(r.first.0, r.first.1))?;
                if value.is_null() { json!(0) } else { value }
            }
            Expr::Binary(b'&', a, b) => json!(Self::bounded_text(format!(
                "{}{}",
                self.text_argument(a, sheet)?,
                self.text_argument(b, sheet)?
            ))?),
            Expr::Call(name, args)
                if (name == "COUNTIF" && args.len() == 2)
                    || (name == "SUMIF" && (args.len() == 2 || args.len() == 3))
                    || (name == "VLOOKUP" && args.len() == 4)
                    || (["XLOOKUP", "_XLFN.XLOOKUP"].contains(&name.as_str())
                        && (3..=6).contains(&args.len())) =>
            {
                if ["XLOOKUP", "_XLFN.XLOOKUP"].contains(&name.as_str()) && !self.preview_limits {
                    return Err(crate::coded(
                        "formula-bad-input",
                        "XLOOKUP requires its independent scalar source proof",
                    ));
                }
                self.common_function(name.strip_prefix("_XLFN.").unwrap_or(name), args, sheet)?
            }
            Expr::Call(name, args) if name == "IF" && args.len() == 3 => {
                let yes = self.predicate(&args[0], sheet)?;
                self.scalar(&args[if yes { 1 } else { 2 }], sheet)?
            }
            Expr::Call(name, args) if name == "IFERROR" && args.len() == 2 => {
                match self.scalar(&args[0], sheet) {
                    Ok(v) => v,
                    Err(e) if e.starts_with("Excel error:") => self.scalar(&args[1], sheet)?,
                    Err(e) => return Err(e),
                }
            }
            Expr::Call(name, args)
                if ["ISNUMBER", "ISTEXT", "ISLOGICAL", "ISBLANK"].contains(&name.as_str())
                    && args.len() == 1 =>
            {
                let value = match self.argument(&args[0], sheet) {
                    Ok(v) => Some(v),
                    Err(e) if e.starts_with("Excel error:") => None,
                    Err(e) => return Err(e),
                };
                json!(value.is_some_and(|v| match name.as_str() {
                    "ISNUMBER" => v.is_number(),
                    "ISTEXT" => v.is_string(),
                    "ISLOGICAL" => v.is_boolean(),
                    _ => v.is_null(),
                }))
            }
            Expr::Call(name, args) if name == "T" && args.len() == 1 => {
                let value = self.argument(&args[0], sheet)?;
                json!(value.as_str().unwrap_or(""))
            }
            Expr::Call(name, args) if name == "DEC2HEX" && args.len() == 1 => {
                // Excel 16.113.3, fixtures/excel-dec2hex-reference.json. This
                // numeric scalar slice returns text and performs no argument
                // coercion. argument() preserves referenced blanks as null.
                let value = self.argument(&args[0], sheet)?;
                let number = value.as_f64().ok_or_else(|| {
                    crate::coded(
                        "formula-unsupported",
                        "DEC2HEX requires a qualified numeric scalar input",
                    )
                })?;
                if !number.is_finite()
                    || number.fract() != 0.0
                    || !(0.0..=2147483647.0).contains(&number)
                {
                    // A subset miss is not an Excel error for IFERROR to hide.
                    return Err(crate::coded(
                        "formula-unsupported",
                        "DEC2HEX number is outside the native integer subset",
                    ));
                }
                json!(format!("{:X}", number as u32))
            }
            Expr::Call(name, args) if name == "EXACT" && args.len() == 2 => {
                json!(self.text_argument(&args[0], sheet)? == self.text_argument(&args[1], sheet)?)
            }
            Expr::Call(name, args) if name == "REPT" && args.len() == 2 => {
                let text = self.text_argument(&args[0], sheet)?;
                let count = self.eval(&args[1], sheet)?;
                if count < 0.0 {
                    return Err("Excel error: #VALUE!".into());
                }
                if !text.is_empty() && count.trunc() > 32767.0 / text.len() as f64 {
                    return Err(crate::coded("formula-limit", "Native text result bound"));
                }
                json!(if text.is_empty() {
                    String::new()
                } else {
                    text.repeat(count.trunc() as usize)
                })
            }
            Expr::Call(name, args)
                if name == "SUBSTITUTE" && (args.len() == 3 || args.len() == 4) =>
            {
                let text = self.text_argument(&args[0], sheet)?;
                let old = self.text_argument(&args[1], sheet)?;
                let new = self.text_argument(&args[2], sheet)?;
                let occurrence = if args.len() == 4 {
                    let n = self.eval(&args[3], sheet)?.trunc();
                    if n < 1.0 {
                        return Err("Excel error: #VALUE!".into());
                    }
                    Some(n as usize)
                } else {
                    None
                };
                let mut out = String::new();
                let mut start = 0;
                if !old.is_empty() {
                    for (i, (at, _)) in text.match_indices(&old).enumerate() {
                        self.tick()?;
                        let replace = occurrence.is_none_or(|n| n == i + 1);
                        let piece = if replace { &new } else { &old };
                        if out.len() + at - start + piece.len() > 32767 {
                            return Err(crate::coded("formula-limit", "Native text result bound"));
                        }
                        out.push_str(&text[start..at]);
                        out.push_str(piece);
                        start = at + old.len();
                    }
                }
                out.push_str(&text[start..]);
                json!(Self::bounded_text(out)?)
            }
            Expr::Call(name, _) if ["AND", "OR", "NOT"].contains(&name.as_str()) => {
                json!(self.predicate(e, sheet)?)
            }
            Expr::Call(name, args)
                if [
                    "LEFT",
                    "RIGHT",
                    "MID",
                    "UPPER",
                    "LOWER",
                    "TRIM",
                    "CONCATENATE",
                ]
                .contains(&name.as_str()) =>
            {
                if name == "CONCATENATE" {
                    let mut out = String::new();
                    for arg in args {
                        out.push_str(&self.text_argument(arg, sheet)?);
                        if out.len() > 32767 {
                            return Err(crate::coded("formula-limit", "Native text result bound"));
                        }
                    }
                    json!(out)
                } else {
                    if args.is_empty() {
                        return Err(crate::coded("formula-refused", "Text function arguments"));
                    }
                    let text = self.text_argument(&args[0], sheet)?;
                    // Preserve UTF-16 slicing boundaries until supplementary
                    // characters and workbook compatibility versions are proven.
                    if text.chars().any(|c| u32::from(c) > 0xffff) {
                        return Err(crate::coded(
                            "formula-bad-input",
                            "Supplementary text slicing requires compatibility",
                        ));
                    }
                    let chars: Vec<char> = text.chars().collect();
                    let out = match name.as_str() {
                        "LEFT" | "RIGHT" if args.len() <= 2 => {
                            let n = if args.len() == 2 {
                                self.eval(&args[1], sheet)?.trunc()
                            } else {
                                1.0
                            };
                            if n < 0.0 {
                                return Err("Excel error: #VALUE!".into());
                            }
                            let n = (n as usize).min(chars.len());
                            if name == "LEFT" {
                                chars[..n].iter().collect()
                            } else {
                                chars[chars.len() - n..].iter().collect()
                            }
                        }
                        "MID" if args.len() == 3 => {
                            let start = self.eval(&args[1], sheet)?.trunc();
                            let len = self.eval(&args[2], sheet)?.trunc();
                            if start < 1.0 || len < 0.0 {
                                return Err("Excel error: #VALUE!".into());
                            }
                            chars
                                .iter()
                                .skip((start as usize).saturating_sub(1))
                                .take(len as usize)
                                .collect()
                        }
                        "TRIM" if args.len() == 1 => text
                            .split(' ')
                            .filter(|s| !s.is_empty())
                            .collect::<Vec<_>>()
                            .join(" "),
                        "UPPER" | "LOWER" if args.len() == 1 && text.is_ascii() => {
                            if name == "UPPER" {
                                text.to_ascii_uppercase()
                            } else {
                                text.to_ascii_lowercase()
                            }
                        }
                        _ => {
                            return Err(crate::coded(
                                "formula-bad-input",
                                "Text function arguments or locale require compatibility",
                            ));
                        }
                    };
                    json!(Self::bounded_text(out)?)
                }
            }
            _ => json!(self.eval(e, sheet)?),
        })
    }
    fn bounded_text(text: String) -> Result<String> {
        if text.len() > 32767 {
            return Err(crate::coded("formula-limit", "Native text result bound"));
        }
        Ok(text)
    }
    fn text_argument(&mut self, e: &Expr, sheet: &str) -> Result<String> {
        let value = if let Expr::Ref(r) = e {
            if r.last.is_some() {
                return Err(crate::coded("formula-refused", "Text range coercion"));
            }
            let name = sheet_name(self.view, r.sheet.as_deref().unwrap_or(sheet))?.to_string();
            self.result(&name, &address(r.first.0, r.first.1))?
        } else {
            self.scalar(e, sheet)?
        };
        let text = match value {
            Value::Null => String::new(),
            Value::String(s) => s,
            Value::Bool(b) => if b { "TRUE" } else { "FALSE" }.into(),
            Value::Number(n) => {
                let n = n
                    .as_f64()
                    .ok_or_else(|| crate::coded("internal-invalid-field", "Invalid text number"))?;
                if n.fract() != 0.0 || n.abs() > 999_999_999_999_999.0 {
                    return Err(crate::coded(
                        "formula-bad-input",
                        "Numeric text spelling requires compatibility engine",
                    ));
                }
                format!("{n:.0}")
            }
            _ => {
                return Err(crate::coded(
                    "formula-bad-input",
                    "Text coercion requires compatibility engine",
                ));
            }
        };
        Self::bounded_text(text)
    }
    // Coercion is specific to arithmetic operators. Aggregates and functions
    // retain their own argument rules, including referenced text suppression.
    fn operator_number(&mut self, e: &Expr, sheet: &str) -> Result<(f64, bool)> {
        let value = self.argument(e, sheet)?;
        let coerced = value.is_string() || value.is_boolean();
        let n = match value {
            Value::Number(n) => n.as_f64().ok_or_else(|| {
                crate::coded("internal-invalid-field", "Invalid arithmetic number")
            })?,
            Value::Bool(b) => f64::from(u8::from(b)),
            Value::Null => 0.0,
            Value::String(text) => {
                let text = text.trim_matches(' ');
                if text.is_empty() || text.bytes().all(|c| c.is_ascii_alphabetic()) {
                    return Err("Excel error: #VALUE!".into());
                }
                let percent = text.ends_with('%');
                let mut text = if percent {
                    text[..text.len() - 1].to_string()
                } else {
                    text.to_string()
                };
                if !text
                    .bytes()
                    .all(|c| c.is_ascii_digit() || b"+-.eE,".contains(&c))
                    || text.contains(',') && (text.contains('e') || text.contains('E'))
                {
                    return Err(crate::coded(
                        "formula-bad-input",
                        "Arithmetic text conversion requires compatibility",
                    ));
                }
                if text.contains(',') {
                    let whole = text
                        .split('.')
                        .next()
                        .unwrap()
                        .trim_start_matches(['+', '-']);
                    let groups: Vec<_> = whole.split(',').collect();
                    if text
                        .split_once('.')
                        .is_some_and(|(_, fraction)| fraction.contains(','))
                        || groups[0].is_empty()
                        || groups[0].len() > 3
                        || groups
                            .iter()
                            .any(|s| !s.bytes().all(|c| c.is_ascii_digit()))
                        || groups[1..].iter().any(|s| s.len() != 3)
                    {
                        return Err("Excel error: #VALUE!".into());
                    }
                    text = text.replace(',', "");
                }
                if !text
                    .bytes()
                    .all(|c| c.is_ascii_digit() || b"+-.eE".contains(&c))
                    || text.bytes().filter(u8::is_ascii_digit).count() > 15
                {
                    return Err(crate::coded(
                        "formula-bad-input",
                        "Arithmetic text conversion requires compatibility",
                    ));
                }
                let value = match Parser::parse(&text)? {
                    Expr::Number(n) => n,
                    Expr::Unary(sign, inner) if matches!(sign, b'+' | b'-') => match *inner {
                        Expr::Number(n) => {
                            if sign == b'-' {
                                -n
                            } else {
                                n
                            }
                        }
                        _ => return Err("Excel error: #VALUE!".into()),
                    },
                    _ => {
                        return Err(crate::coded(
                            "formula-bad-input",
                            "Arithmetic text conversion requires compatibility",
                        ));
                    }
                };
                if percent { value / 100.0 } else { value }
            }
            _ => {
                return Err(crate::coded(
                    "formula-bad-input",
                    "Arithmetic coercion requires compatibility",
                ));
            }
        };
        if coerced && n.abs() > 999_999_999_999_999.0 {
            return Err(crate::coded(
                "formula-bad-input",
                "Fractional or large arithmetic coercion requires compatibility",
            ));
        }
        Ok((self.scalar_number(n)?, coerced))
    }
    fn operator_pair(&mut self, a: &Expr, b: &Expr, op: u8, sheet: &str) -> Result<(f64, f64)> {
        let (a, ca) = self.operator_number(a, sheet)?;
        let (b, cb) = self.operator_number(b, sheet)?;
        if ca || cb {
            if a.abs() > 999_999_999_999_999.0 || b.abs() > 999_999_999_999_999.0 {
                return Err(crate::coded(
                    "formula-bad-input",
                    "Large arithmetic coercion requires compatibility",
                ));
            }
            if op == b'^' && (a.fract() != 0.0 || b.fract() != 0.0) {
                return Err(crate::coded(
                    "formula-bad-input",
                    "Inexact power coercion requires compatibility",
                ));
            }
        }
        Ok((a, b))
    }
    fn eval(&mut self, e: &Expr, sheet: &str) -> Result<f64> {
        self.tick()?;
        let n = match e {
            Expr::Number(n) => *n,
            Expr::Text | Expr::String(_) => {
                return Err(crate::coded(
                    "formula-bad-input",
                    "Text coercion requires compatibility engine",
                ));
            }
            Expr::Error(error) => return Err(format!("Excel error: {error}")),
            Expr::QualifiedError(..) => return Err("Excel error: #REF!".into()),
            Expr::Boolean(_) | Expr::Compare(..) => {
                return Err(crate::coded(
                    "formula-bad-input",
                    "Boolean result requires compatibility engine",
                ));
            }
            Expr::Name(..) | Expr::QualifiedName(..) | Expr::Structured(..) => {
                return Err(crate::coded(
                    "formula-bad-input",
                    "Defined name requires compatibility engine",
                ));
            }
            Expr::External => {
                return Err(crate::coded(
                    "formula-bad-input",
                    "External or 3D reference requires compatibility engine",
                ));
            }
            Expr::Union(_) => {
                return Err(crate::coded(
                    "formula-unsupported",
                    "Scalar union semantics are not qualified",
                ));
            }
            Expr::Binary(b':', _, _) => return Err("Excel error: #REF!".into()),
            Expr::Ref(r) => {
                if r.last.is_some() {
                    return Err(crate::coded(
                        "formula-bad-input",
                        "Scalar range requires compatibility engine",
                    ));
                }
                let name = sheet_name(self.view, r.sheet.as_deref().unwrap_or(sheet))?.to_string();
                self.cell(&name, &address(r.first.0, r.first.1))?
                    .unwrap_or(0.0)
            }
            Expr::Unary(op, e) => {
                let (n, _) = self.operator_number(e, sheet)?;
                match op {
                    b'-' => -n,
                    b'%' => exact_divide(n, 100.0)?,
                    _ => n,
                }
            }
            Expr::Binary(op, a, b) => {
                let (a, b) = self.operator_pair(a, b, *op, sheet)?;
                match op {
                    b'+' => a + b,
                    b'-' => a - b,
                    b'*' => a * b,
                    b'/' => {
                        if b == 0.0 {
                            return Err("Excel error: #DIV/0!".into());
                        }
                        a / b
                    }
                    b'^' => power(a, b)?,
                    _ => {
                        return Err(crate::coded(
                            "formula-bad-input",
                            "Operator requires compatibility engine",
                        ));
                    }
                }
            }
            Expr::Call(name, args) => match name.as_str() {
                "N" if args.len() == 1 => match self.argument(&args[0], sheet)? {
                    Value::Number(n) => n.as_f64().ok_or_else(|| {
                        crate::coded("internal-invalid-field", "Invalid numeric argument")
                    })?,
                    Value::Bool(b) => f64::from(u8::from(b)),
                    Value::Null | Value::String(_) => 0.0,
                    _ => {
                        return Err(crate::coded(
                            "formula-bad-input",
                            "N coercion requires compatibility",
                        ));
                    }
                },
                "COUNTBLANK" if args.len() == 1 => {
                    let Expr::Ref(r) = &args[0] else {
                        return Err(crate::coded(
                            "formula-bad-input",
                            "COUNTBLANK requires a reference",
                        ));
                    };
                    self.count_blank(r, sheet)?
                }
                "ROWS" | "COLUMNS" if args.len() == 1 => {
                    let Expr::Ref(r) = &args[0] else {
                        return Err(crate::coded(
                            "formula-bad-input",
                            "ROWS/COLUMNS requires a reference",
                        ));
                    };
                    sheet_name(self.view, r.sheet.as_deref().unwrap_or(sheet))?;
                    let end = r.last.unwrap_or(r.first);
                    f64::from(if name == "ROWS" {
                        end.0 - r.first.0 + 1
                    } else {
                        end.1 - r.first.1 + 1
                    })
                }
                "LEN" if args.len() == 1 => {
                    let text = self.text_argument(&args[0], sheet)?;
                    if text.chars().any(|c| u32::from(c) > 0xffff) {
                        return Err(crate::coded(
                            "formula-unsupported",
                            "LEN supplementary Unicode requires compatibility qualification",
                        ));
                    }
                    text.encode_utf16().count() as f64
                }
                "POWER" if args.len() == 2 => {
                    power(self.eval(&args[0], sheet)?, self.eval(&args[1], sheet)?)?
                }
                "SQRT" if args.len() == 1 => {
                    let n = self.eval(&args[0], sheet)?;
                    if n < 0.0 {
                        return Err("Excel error: #NUM!".into());
                    }
                    n.sqrt()
                }
                "MOD" if args.len() == 2 => {
                    let a = self.eval(&args[0], sheet)?;
                    let b = self.eval(&args[1], sheet)?;
                    if b == 0.0 {
                        return Err("Excel error: #DIV/0!".into());
                    }
                    a - b * (a / b).floor()
                }
                "TRUNC" if args.len() == 1 => self.eval(&args[0], sheet)?.trunc(),
                "TRUNC" if args.len() == 2 => {
                    let value = self.eval(&args[0], sheet)?;
                    let digits = self.eval(&args[1], sheet)?;
                    if digits == 0.0 {
                        value.trunc()
                    } else {
                        let precision = digits.trunc();
                        if !(-12.0..=12.0).contains(&precision) {
                            return Err(crate::coded(
                                "formula-bad-input",
                                "TRUNC precision requires compatibility engine",
                            ));
                        }
                        let factor = 10_f64.powi(precision as i32);
                        let scaled = exact_multiply(value.abs(), factor)?;
                        if scaled > 9_000_000_000_000_000.0 {
                            return Err(crate::coded(
                                "formula-bad-input",
                                "TRUNC magnitude requires compatibility engine",
                            ));
                        }
                        let truncated = (scaled + f64::EPSILON * scaled.max(1.0)).trunc() / factor;
                        if truncated == 0.0 {
                            0.0
                        } else {
                            value.signum() * truncated
                        }
                    }
                }
                "PMT" | "PV" | "FV" => self.financial(name, args, sheet)?,
                "DATE" | "YEAR" | "MONTH" | "DAY" | "EDATE" | "EOMONTH"
                    if args.len()
                        == if name == "DATE" {
                            3
                        } else if ["EDATE", "EOMONTH"].contains(&name.as_str()) {
                            2
                        } else {
                            1
                        } =>
                {
                    let owner = self.view["sheets"]
                        .as_array()
                        .and_then(|sheets| sheets.iter().find(|s| s["name"] == sheet))
                        .ok_or_else(|| {
                            crate::coded(
                                "internal-missing-field",
                                "Date-function worksheet is missing",
                            )
                        })?;
                    let date1904 = match owner.get("date1904") {
                        None => false,
                        Some(Value::Bool(value)) => *value,
                        _ => {
                            return Err(crate::coded(
                                "internal-invalid-field",
                                "Invalid workbook date-system metadata",
                            ));
                        }
                    };
                    let values = args
                        .iter()
                        .map(|arg| {
                            self.scalar(arg, sheet)?.as_f64().ok_or_else(|| {
                                "Date functions require numeric scalar inputs".to_string()
                            })
                        })
                        .collect::<Result<Vec<_>>>()?;
                    date::calculate(name, &values, date1904)?
                }
                "INT" if args.len() == 1 => self.eval(&args[0], sheet)?.floor(),
                "SIGN" if args.len() == 1 => {
                    let n = self.eval(&args[0], sheet)?;
                    if n == 0.0 { 0.0 } else { n.signum() }
                }
                "ABS" if args.len() == 1 => self.eval(&args[0], sheet)?.abs(),
                "ROUND" if args.len() == 2 => {
                    let a = self.eval(&args[0], sheet)?;
                    let b = self.eval(&args[1], sheet)?;
                    crate::decimal::round(a, b)?
                }
                "IF" if args.len() == 3 => {
                    let yes = self.predicate(&args[0], sheet)?;
                    self.eval(&args[if yes { 1 } else { 2 }], sheet)?
                }
                "IFERROR" if args.len() == 2 => match self.eval(&args[0], sheet) {
                    Ok(n) => n,
                    Err(e) if e.starts_with("Excel error:") => self.eval(&args[1], sheet)?,
                    Err(e) => return Err(e),
                },
                "SUM" | "AVERAGE" | "MIN" | "MAX" | "COUNT" | "COUNTA" | "PRODUCT" => {
                    self.aggregate(name, args, sheet)?
                }
                _ => {
                    return Err(crate::coded(
                        "formula-bad-input",
                        "Function arguments require compatibility engine",
                    ));
                }
            },
        };
        self.scalar_number(n)
    }
}
fn evaluate(view: &Value) -> Result<Value> {
    let mut eval = Eval {
        db: None,
        selected_only: false,
        preview_limits: false,
        fenced_spill_inputs: false,
        overrides: HashMap::new(),
        view,
        trees: HashMap::new(),
        memo: HashMap::new(),
        active: HashSet::new(),
        steps: 0,
        constants: HashMap::new(),
        constant_bytes: 0,
        constant_entries: 0,
        constant_hits: 0,
    };
    let mut coordinates = Vec::new();
    for sheet in view["sheets"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheets"))?
    {
        for (reference, cell) in sheet["cells"]
            .as_object()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing cells"))?
        {
            if cell.get("formulaType").is_some() {
                return Err(crate::coded("formula-unsupported", "Special formula"));
            }
            if let Some(f) = cell["formula"].as_str() {
                let key = (field(sheet, "name")?.to_string(), reference.clone());
                eval.trees
                    .insert(key.clone(), resolved(f, view, &key.0, reference)?);
                coordinates.push(key);
            }
        }
    }
    let mut results = Vec::new();
    for (sheet, reference) in coordinates {
        match eval.result(&sheet,&reference) {
            Ok(value)=>results.push(json!({"sheet":sheet,"ref":reference,"value":value})),
            Err(error) if crate::sentence(&error).strip_prefix("Excel error: ").is_some_and(crate::preview::error_code)=>results.push(json!({"sheet":sheet,"ref":reference,"error":crate::sentence(&error).trim_start_matches("Excel error: ")})),
            Err(error)=>return Err(error),
        }
    }
    Ok(json!({"results":results,"steps":eval.steps}))
}
/// Excel compensates only an unwrapped final add/subtract, not a nested
/// expression. Retain authored grouping rather than guessing from a rounded result.
fn root_add_sub(formula: &str) -> bool {
    let text = formula.trim().trim_start_matches('=').trim();
    let mut depth = 0;
    let mut brackets = 0;
    let mut quote = 0;
    let mut candidate = false;
    let bytes = text.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        let c = bytes[i];
        if quote != 0 {
            if c == quote {
                if bytes.get(i + 1) == Some(&quote) {
                    i += 1;
                } else {
                    quote = 0;
                }
            }
        } else {
            match c {
                b'"' | b'\'' => quote = c,
                b'[' => brackets += 1,
                b']' => brackets -= 1,
                b'(' if brackets == 0 => depth += 1,
                b')' if brackets == 0 => depth -= 1,
                b'=' | b'<' | b'>' | b'&' if depth == 0 && brackets == 0 => return false,
                b'+' | b'-' if depth == 0 && brackets == 0 && i > 0 => {
                    let prev = text[..i].trim_end().as_bytes().last().copied().unwrap_or(0);
                    let exponent =
                        matches!(prev, b'e' | b'E') && i > 1 && bytes[i - 2].is_ascii_digit();
                    if !exponent && !matches!(prev, b'+' | b'-' | b'*' | b'/' | b'^' | b'(' | b',')
                    {
                        candidate = true;
                    }
                }
                _ => {}
            }
        }
        i += 1;
    }
    candidate && matches!(Parser::parse(text), Ok(Expr::Binary(b'+' | b'-', _, _)))
}
fn compensated(a: f64, b: f64) -> f64 {
    let n = a + b;
    if a.is_sign_positive() != b.is_sign_positive() && a != 0.0 && a.is_finite() {
        let magnitude = a.abs();
        let ulp = f64::from_bits(magnitude.to_bits() + 1) - magnitude;
        if ulp.is_finite() && n.abs() < ulp * 8.0 {
            return 0.0;
        }
    }
    n
}
fn power(a: f64, b: f64) -> Result<f64> {
    if a == 2.0 && b.fract() == 0.0 && (-1022.0..=1023.0).contains(&b) {
        return exact_number(2_f64.powi(b as i32));
    }
    if b.fract() != 0.0 || b.abs() > 32.0 {
        return Err(crate::coded(
            "formula-bad-input",
            "Power exponent requires compatibility engine",
        ));
    }
    if a == 0.0 && b <= 0.0 {
        return Err("Excel error: #NUM!".into());
    }
    // Pinned engine uses IEEE Math.pow. Restrict to exact integer outputs here.
    let n = a.powi(b as i32);
    if a.fract() != 0.0 || b < 0.0 || n.abs() > 9007199254740991.0 {
        return Err(crate::coded(
            "formula-bad-input",
            "Power precision requires compatibility engine",
        ));
    }
    exact_number(n)
}
fn spelling(p: (u32, u32), fixed: (bool, bool)) -> String {
    let a = address(p.0, p.1);
    let i = a.find(|c: char| c.is_ascii_digit()).unwrap();
    format!(
        "{}{}{}{}",
        if fixed.1 { "$" } else { "" },
        &a[..i],
        if fixed.0 { "$" } else { "" },
        &a[i..]
    )
}
fn column_spelling(column: u32, fixed: bool) -> String {
    let text = address(0, column);
    format!(
        "{}{}",
        if fixed { "$" } else { "" },
        text.trim_end_matches('1')
    )
}
/// Preserve the qualified shared sort policy: relative references within the
/// moved row, absolute rows elsewhere, and no dependents outside the selection.
pub(crate) fn sorted_formula(
    formula: &str,
    owner: &str,
    target: &str,
    row: u32,
    rows: (u32, u32),
    destination: Option<u32>,
) -> Result<String> {
    let ast = Parser::references(formula)?;
    let mut refs = Vec::new();
    // A name keeps pointing at the same cells after a sort, as in Excel, so a
    // formula that uses one is moved with its text unchanged. A table reference
    // is different: it follows the table.
    fn unbound(expr: &Expr) -> bool {
        match expr {
            Expr::Structured(..) => true,
            Expr::Unary(_, e) => unbound(e),
            Expr::Binary(_, a, b) | Expr::Compare(_, a, b) => unbound(a) || unbound(b),
            Expr::Call(_, args) => args.iter().any(unbound),
            Expr::Union(areas) => areas.iter().any(unbound),
            _ => false,
        }
    }
    if unbound(&ast) {
        return Err(crate::coded(
            "formula-unsupported",
            "Sort cannot safely update table references",
        ));
    }
    visit(&ast, &mut refs, false, true)?;
    // A formula on a sorted row moves the way Excel moves it: relative
    // references keep their distance from the formula's new row, absolute ones
    // stay. That covers a running balance (=F4+D5) and a rate held in one cell
    // (=C5*$B$1) alike. A formula outside the sorted rows keeps its text: a
    // total over the block is unchanged and is recalculated like any dependent.
    let _ = (owner, target, rows);
    if let Some(destination) = destination {
        let result = rewrite(
            &json!({"items":[{"formula":formula,"copyRows":i64::from(destination)-i64::from(row)}]}),
        )?;
        Ok(result["formulas"][0]
            .as_str()
            .ok_or_else(|| crate::coded("internal-missing-field", "Missing sorted formula"))?
            .into())
    } else {
        Ok(formula.into())
    }
}
fn rewrite_deleted_table_columns(formula: &str, input: &Value, item: &Value) -> Result<String> {
    let Some(changes) = input["tableChanges"].as_array() else {
        return Ok(formula.into());
    };
    if changes
        .iter()
        .all(|c| c["removed"].as_array().is_none_or(|v| v.is_empty()))
    {
        return Ok(formula.into());
    }
    let ast = Parser::references(formula)?;
    let mut tokens = Vec::new();
    fn collect<'a>(e: &'a Expr, out: &mut Vec<(&'a str, usize, usize)>) {
        match e {
            Expr::Structured(t, a, b) | Expr::Name(t, a, b) => out.push((t, *a, *b)),
            Expr::Unary(_, e) => collect(e, out),
            Expr::Binary(_, a, b) | Expr::Compare(_, a, b) => {
                collect(a, out);
                collect(b, out);
            }
            Expr::Call(_, args) | Expr::Union(args) => {
                for a in args {
                    collect(a, out);
                }
            }
            _ => {}
        }
    }
    fn bare_deleted(e: &Expr, changes: &[Value]) -> bool {
        match e {
            Expr::QualifiedName(_, name, _, _) => changes.iter().any(|c| {
                c["deleted"] == true
                    && c["name"]
                        .as_str()
                        .is_some_and(|n| n.eq_ignore_ascii_case(name))
            }),
            Expr::Unary(_, e) => bare_deleted(e, changes),
            Expr::Binary(_, a, b) | Expr::Compare(_, a, b) => {
                bare_deleted(a, changes) || bare_deleted(b, changes)
            }
            Expr::Call(_, args) => args.iter().any(|a| bare_deleted(a, changes)),
            Expr::Union(areas) => areas.iter().any(|a| bare_deleted(a, changes)),
            _ => false,
        }
    }
    if bare_deleted(&ast, changes) {
        return Err(crate::coded(
            "formula-unsupported",
            "Bare deleted-table names require reference qualification",
        ));
    }
    collect(&ast, &mut tokens);
    let mut out = formula.to_string();
    for (token, start, end) in tokens.into_iter().rev() {
        let Some(split) = token.find('[') else {
            if changes.iter().any(|c| {
                c["deleted"] == true
                    && c["name"]
                        .as_str()
                        .is_some_and(|n| n.eq_ignore_ascii_case(token))
            }) {
                out.replace_range(start..end, "#REF!");
            }
            continue;
        };
        let table = &token[..split];
        if table.is_empty() && position(item["reference"].as_str().unwrap_or("")).is_err() {
            return Err(crate::coded(
                "formula-unsupported",
                "Unbound table shorthand in metadata requires qualification",
            ));
        }
        let change = changes.iter().find(|c| {
            if !table.is_empty() {
                return c["name"]
                    .as_str()
                    .is_some_and(|n| n.eq_ignore_ascii_case(table));
            }
            if c["sheet"] != item["owner"] {
                return false;
            }
            let Ok((a, b)) = crate::bulk::bounds(c["ref"].as_str().unwrap_or("")) else {
                return false;
            };
            let Ok(p) = position(item["reference"].as_str().unwrap_or("")) else {
                return false;
            };
            p.0 > a.0 && p.0 <= b.0 && p.1 >= a.1 && p.1 <= b.1
        });
        let Some(change) = change else {
            continue;
        };
        if change["deleted"] == true {
            out.replace_range(start..end, "#REF!");
            continue;
        }
        let spec = token[split..].replace(['[', ']'], "");
        let column = spec
            .strip_prefix('@')
            .or_else(|| spec.strip_prefix("#This Row,"))
            .or_else(|| spec.strip_prefix("#Data,"))
            .or_else(|| spec.strip_prefix("#All,"))
            .or_else(|| spec.strip_prefix("#Headers,"))
            .unwrap_or(&spec);
        if ["#All", "#Data", "#Headers", "#Totals"].contains(&column) {
            continue;
        }
        if column.contains([':', '\'', ',']) {
            return Err(crate::coded(
                "formula-unsupported",
                "Complex deleted table selector requires qualification",
            ));
        }
        if change["removed"].as_array().is_some_and(|cols| {
            cols.iter()
                .any(|c| c.as_str().is_some_and(|n| n.eq_ignore_ascii_case(column)))
        }) {
            out.replace_range(start..end, "#REF!");
        }
    }
    Ok(out)
}
/// Conservative static consumers for named-range edit/removal. Text literals
/// cannot become names; dynamic/external/structured consumers stay protected.
pub(crate) fn static_name_use(
    formula: &str,
    target: &str,
    owner: &str,
    scope: Option<&str>,
    shadow: bool,
) -> Result<bool> {
    let ast = Parser::references(formula)?;
    fn inspect(
        e: &Expr,
        target: &str,
        owner: &str,
        scope: Option<&str>,
        shadow: bool,
    ) -> Result<bool> {
        Ok(match e {
            Expr::Name(name, _, _) => {
                name.eq_ignore_ascii_case(target)
                    && scope
                        .map(|s| s.eq_ignore_ascii_case(owner))
                        .unwrap_or(!shadow)
            }
            Expr::QualifiedName(sheet, name, _, _) => {
                name.eq_ignore_ascii_case(target)
                    && scope.is_some_and(|s| s.eq_ignore_ascii_case(sheet))
            }
            Expr::Structured(..) | Expr::QualifiedError(..) => {
                return Err(crate::coded(
                    "formula-unsupported",
                    "Structured/external name consumers require qualification",
                ));
            }
            Expr::Unary(_, a) => inspect(a, target, owner, scope, shadow)?,
            Expr::Binary(_, a, b) | Expr::Compare(_, a, b) => {
                let first = inspect(a, target, owner, scope, shadow)?;
                let second = inspect(b, target, owner, scope, shadow)?;
                first || second
            }
            Expr::Call(name, args) => {
                if !supported_function(name) && name != "XLOOKUP" {
                    return Err(crate::coded(
                        "formula-unsupported",
                        "Dynamic or unknown name consumers require qualification",
                    ));
                }
                let mut used = false;
                for a in args {
                    used |= inspect(a, target, owner, scope, shadow)?;
                }
                used
            }
            Expr::Union(args) => {
                let mut used = false;
                for a in args {
                    used |= inspect(a, target, owner, scope, shadow)?;
                }
                used
            }
            Expr::Ref(r)
                if r.sheet
                    .as_ref()
                    .is_some_and(|s| s.contains('[') || s.contains(']') || s.contains(':')) =>
            {
                return Err(crate::coded(
                    "formula-unsupported",
                    "External/3D name consumers require qualification",
                ));
            }
            _ => false,
        })
    }
    inspect(&ast, target, owner, scope, shadow)
}

/// Validate whether copying changes a cell/axis reference, without caching results.
pub(crate) fn copy_invariant(formula: &str) -> Result<bool> {
    fn invariant(expr: &Expr) -> bool {
        match expr {
            Expr::Ref(_) => false,
            Expr::Unary(_, child) => invariant(child),
            Expr::Binary(_, a, b) | Expr::Compare(_, a, b) => invariant(a) && invariant(b),
            Expr::Call(_, args) => args.iter().all(invariant),
            Expr::Union(areas) => areas.iter().all(invariant),
            _ => true,
        }
    }
    Ok(invariant(&Parser::references(formula)?))
}
pub(crate) fn rewrite(input: &Value) -> Result<Value> {
    let mut formulas = Vec::new();
    for item in input["items"]
        .as_array()
        .ok_or_else(|| crate::coded("internal-missing-field", "Missing rewrite items"))?
    {
        crate::cancel::check()?;
        let copy_offsets = if item.get("copyRows").is_some() || item.get("copyColumns").is_some() {
            Some((
                item.get("copyRows")
                    .map(|n| {
                        n.as_i64().ok_or_else(|| {
                            crate::coded("internal-invalid-field", "Invalid copy row offset")
                        })
                    })
                    .transpose()?
                    .unwrap_or(0),
                item.get("copyColumns")
                    .map(|n| {
                        n.as_i64().ok_or_else(|| {
                            crate::coded("internal-invalid-field", "Invalid copy column offset")
                        })
                    })
                    .transpose()?
                    .unwrap_or(0),
            ))
        } else {
            None
        };
        let rewritten = rewrite_deleted_table_columns(field(item, "formula")?, input, item)?;
        let formula = rewritten.as_str();
        let ast = Parser::references(formula)?;
        let mut refs = Vec::new();
        if item.get("copyRows").is_some() || item.get("copyColumns").is_some() {
            fn copy_refs<'a>(expr: &'a Expr, refs: &mut Vec<(&'a Ref, bool)>) {
                match expr {
                    Expr::Ref(r) => refs.push((r, false)),
                    Expr::Unary(_, e) => copy_refs(e, refs),
                    Expr::Binary(_, a, b) | Expr::Compare(_, a, b) => {
                        copy_refs(a, refs);
                        copy_refs(b, refs);
                    }
                    Expr::Call(_, args) | Expr::Union(args) => {
                        for e in args {
                            copy_refs(e, refs)
                        }
                    }
                    _ => {}
                }
            }
            copy_refs(&ast, &mut refs);
        } else {
            visit(&ast, &mut refs, false, true)?;
        }
        let mut output = formula.to_string();
        for (r, aggregate) in refs.into_iter().rev() {
            if (item["nameBoundary"] == true || item["name"] == true)
                && item["owner"].as_str().unwrap_or("").is_empty()
                && r.sheet.is_none()
            {
                return Err(crate::coded(
                    "formula-unsupported",
                    "Unqualified global name requires compatibility engine",
                ));
            }
            // A table-local column edit moves references inside this rectangle
            // only. Crossing ranges and whole axes need a separate union policy.
            if let Some(range) = input["tableRange"].as_str() {
                if !r
                    .sheet
                    .as_deref()
                    .unwrap_or(item["owner"].as_str().unwrap_or(""))
                    .eq_ignore_ascii_case(field(input, "target")?)
                {
                    continue;
                }
                let (a, b) = crate::bulk::bounds(range)?;
                let end = r.last.unwrap_or(r.first);
                let intersects =
                    r.first.0 <= b.0 && end.0 >= a.0 && r.first.1 <= b.1 && end.1 >= a.1;
                if !intersects {
                    continue;
                }
                if r.whole_rows
                    || r.whole_columns
                    || r.first.0 < a.0
                    || end.0 > b.0
                    || r.first.1 < a.1
                    || end.1 > b.1
                {
                    return Err(crate::coded(
                        "formula-refused",
                        "A reference crosses the table column movement boundary",
                    ));
                }
            }
            let mut first = r.first;
            let mut last = r.last;
            let columns = input["axis"] == "column"
                && item.get("copyRows").is_none()
                && item.get("copyColumns").is_none();
            let limit = if columns { 16384 } else { 1048576 };
            if columns {
                first = (first.1, first.0);
                last = last.map(|p| (p.1, p.0));
            }
            if let Some((offset, column_offset)) = copy_offsets {
                let shift = |p: &mut (u32, u32), fixed: (bool, bool)| -> bool {
                    if !fixed.0 {
                        let Some(row) = i64::from(p.0).checked_add(offset) else {
                            return false;
                        };
                        if !(0..1048576).contains(&row) {
                            return false;
                        }
                        p.0 = row as u32
                    }
                    if !fixed.1 {
                        let Some(col) = i64::from(p.1).checked_add(column_offset) else {
                            return false;
                        };
                        if !(0..16384).contains(&col) {
                            return false;
                        }
                        p.1 = col as u32;
                    }
                    true
                };
                let first_valid = shift(&mut first, r.fixed);
                let last_valid = last.as_mut().map(|b| shift(b, r.last_fixed));
                if !first_valid || last_valid == Some(false) {
                    // Excel Fill Up collapses the whole reference, including
                    // a range whose other endpoint is still valid/absolute.
                    output.replace_range(r.start..r.end, &format!("{}#REF!", r.prefix));
                    continue;
                }
            } else if !(r.whole_columns && !columns || r.whole_rows && columns)
                && r.sheet
                    .as_deref()
                    .unwrap_or(item["owner"].as_str().unwrap_or(""))
                    .eq_ignore_ascii_case(field(input, "target")?)
            {
                let index = input["index"]
                    .as_u64()
                    .filter(|n| *n < limit as u64)
                    .ok_or_else(|| crate::coded("formula-malformed", "Invalid structural index"))?
                    as u32;
                let count = input["count"]
                    .as_u64()
                    .filter(|n| *n > 0 && *n <= limit as u64)
                    .ok_or_else(|| {
                        crate::coded("internal-invalid-field", "Invalid structural count")
                    })? as u32;
                if input["remove"] == true {
                    let end = index + count;
                    if let Some(b) = &mut last {
                        if first.0 >= index && b.0 < end {
                            // Excel keeps the sheet name: Data!#REF!
                            output.replace_range(r.start..r.end, &format!("{}#REF!", r.prefix));
                            continue;
                        }
                        // Shrink the range to surviving rows, then close the gap.
                        first.0 = if first.0 >= end {
                            first.0 - count
                        } else if first.0 >= index {
                            index
                        } else {
                            first.0
                        };
                        // A range that runs to the last row or column still
                        // does afterwards: Excel brings new ones in at the edge.
                        b.0 = if b.0 == limit - 1 && b.0 >= end {
                            b.0
                        } else if b.0 >= end {
                            b.0 - count
                        } else if b.0 >= index {
                            index - 1
                        } else {
                            b.0
                        };
                    } else if first.0 >= index && first.0 < end {
                        output.replace_range(r.start..r.end, &format!("{}#REF!", r.prefix));
                        continue;
                    } else if first.0 >= end {
                        first.0 -= count;
                    }
                } else {
                    if first.0 >= index {
                        first.0 += count
                    }
                    if let Some(b) = &mut last
                        && (b.0 >= index
                            || (!columns
                                && b.0 + 1 == index
                                && (item["nameBoundary"] == true
                                    || item["boundary"] == true && aggregate)))
                    {
                        // What is pushed past the last row or column falls off.
                        b.0 = (b.0 + count).min(limit - 1);
                    }
                }
                if first.0 >= limit || last.is_some_and(|b| b.0 >= limit) {
                    return Err(crate::coded(
                        "formula-unsupported",
                        "Inserted reference outside worksheet",
                    ));
                }
            }
            if columns {
                first = (first.1, first.0);
                last = last.map(|p| (p.1, p.0));
            }
            let replacement = format!(
                "{}{}{}",
                r.sheet
                    .as_ref()
                    .filter(|s| !s.is_empty()
                        && s.bytes().all(|b| b.is_ascii_alphabetic() || b == b'_'))
                    .map(|s| format!("{s}!"))
                    .unwrap_or_else(|| r.prefix.clone()),
                if r.whole_columns {
                    column_spelling(first.1, r.fixed.1)
                } else if r.whole_rows {
                    format!("{}{}", if r.fixed.0 { "$" } else { "" }, first.0 + 1)
                } else {
                    spelling(first, r.fixed)
                },
                last.map(|b| format!(
                    ":{}",
                    if r.whole_columns {
                        column_spelling(b.1, r.last_fixed.1)
                    } else if r.whole_rows {
                        format!("{}{}", if r.last_fixed.0 { "$" } else { "" }, b.0 + 1)
                    } else {
                        spelling(b, r.last_fixed)
                    }
                ))
                .unwrap_or_default()
            );
            output.replace_range(r.start..r.end, &replacement);
        }
        formulas.push(output);
    }
    Ok(json!({"formulas":formulas}))
}
/// Rename a proved workbook table symbol, preserving every non-token byte.
/// The package owner independently proves no defined-name collision or opaque
/// reference-bearing parts before invoking this lexical reference operation.
pub(crate) fn rename_table(formula: &str, old: &str, new: &str) -> Result<String> {
    let ast = Parser::references(formula)?;
    let mut references = Vec::new();
    visit(&ast, &mut references, false, true)?; // Refuse INDIRECT/OFFSET addressing.
    let mut edits = Vec::new();
    fn collect(
        expr: &Expr,
        old: &str,
        new: &str,
        edits: &mut Vec<(usize, usize, String)>,
    ) -> Result<()> {
        match expr {
            Expr::Structured(token, start, _) => {
                let split = token.find('[').ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing table selector")
                })?;
                if token[..split].eq_ignore_ascii_case(old) {
                    edits.push((*start, *start + split, new.into()));
                }
            }
            Expr::Name(name, start, end) if name.eq_ignore_ascii_case(old) => {
                edits.push((*start, *end, new.into()))
            }
            Expr::QualifiedName(_, name, _, _) if name.eq_ignore_ascii_case(old) => {
                return Err(crate::coded(
                    "formula-unsupported",
                    "Qualified bare table names require separate rename qualification",
                ));
            }
            Expr::Unary(_, value) => collect(value, old, new, edits)?,
            Expr::Binary(_, left, right) | Expr::Compare(_, left, right) => {
                collect(left, old, new, edits)?;
                collect(right, old, new, edits)?;
            }
            Expr::Call(_, args) | Expr::Union(args) => {
                for arg in args {
                    collect(arg, old, new, edits)?;
                }
            }
            _ => (),
        }
        Ok(())
    }
    collect(&ast, old, new, &mut edits)?;
    let mut output = formula.to_owned();
    for (start, end, replacement) in edits.into_iter().rev() {
        output.replace_range(start..end, &replacement);
    }
    Ok(output)
}
/// Rename one ordinary table column with explicit shorthand ownership.
/// Returns edits only for parsed structured tokens, preserving all other bytes.
pub(crate) fn rename_table_column(
    formula: &str,
    table: &str,
    old: &str,
    new: &str,
    inside: bool,
) -> Result<String> {
    let ast = Parser::references(formula)?;
    visit(&ast, &mut Vec::new(), false, true)?;
    fn collect(
        expr: &Expr,
        table: &str,
        old: &str,
        new: &str,
        inside: bool,
        edits: &mut Vec<(usize, usize, String)>,
    ) -> Result<()> {
        match expr {
            Expr::Structured(token, start, end) => {
                let split = token.find('[').ok_or_else(|| {
                    crate::coded("internal-missing-field", "Missing column selector")
                })?;
                let owner = &token[..split];
                if owner.eq_ignore_ascii_case(table) || (owner.is_empty() && inside) {
                    let spec = &token[split..];
                    let (prefix, label, suffix) = if spec.starts_with("[[") {
                        let (prefix, rest) = ["[[#This Row],[", "[[#Data],[", "[[#All],["]
                            .iter()
                            .find_map(|prefix| {
                                spec.strip_prefix(prefix).map(|rest| (*prefix, rest))
                            })
                            .ok_or_else(|| {
                                crate::coded(
                                    "formula-unsupported",
                                    "Complex table column selector requires qualification",
                                )
                            })?;
                        (
                            prefix,
                            rest.strip_suffix("]]").ok_or_else(|| {
                                crate::coded("internal-invalid-field", "Invalid column selector")
                            })?,
                            "]]",
                        )
                    } else {
                        let prefix = if spec.starts_with("[@") { "[@" } else { "[" };
                        (
                            prefix,
                            spec.strip_prefix(prefix)
                                .and_then(|rest| rest.strip_suffix(']'))
                                .ok_or_else(|| {
                                    crate::coded(
                                        "internal-invalid-field",
                                        "Invalid column selector",
                                    )
                                })?,
                            "]",
                        )
                    };
                    if label.contains(['[', ']', ':', '\'', ',', '#', '@']) {
                        return Err(crate::coded(
                            "formula-unsupported",
                            "Complex table column selector requires qualification",
                        ));
                    }
                    if label.eq_ignore_ascii_case(old) {
                        edits.push((
                            *start,
                            *end,
                            if new.is_empty() {
                                "#REF!".into()
                            } else {
                                format!("{owner}{prefix}{new}{suffix}")
                            },
                        ));
                    }
                } else if owner.is_empty() {
                    return Err(crate::coded(
                        "formula-unsupported",
                        "Unbound shorthand table column requires qualification",
                    ));
                }
            }
            Expr::Unary(_, value) => collect(value, table, old, new, inside, edits)?,
            Expr::Binary(_, left, right) | Expr::Compare(_, left, right) => {
                collect(left, table, old, new, inside, edits)?;
                collect(right, table, old, new, inside, edits)?;
            }
            Expr::Call(_, args) | Expr::Union(args) => {
                for arg in args {
                    collect(arg, table, old, new, inside, edits)?;
                }
            }
            _ => (),
        }
        Ok(())
    }
    let mut edits = Vec::new();
    collect(&ast, table, old, new, inside, &mut edits)?;
    let mut output = formula.to_owned();
    for (start, end, replacement) in edits.into_iter().rev() {
        output.replace_range(start..end, &replacement);
    }
    Ok(output)
}
pub(crate) fn rename(formula: &str, old: &str, new: &str) -> Result<String> {
    let ast = Parser::references(formula)?;
    let mut patches = Vec::new();
    fn collect(
        e: &Expr,
        old: &str,
        new: &str,
        out: &mut Vec<(usize, usize, String)>,
    ) -> Result<()> {
        let prefix = format!("'{}'!", new.replace('\'', "''"));
        match e {
            Expr::Ref(r)
                if r.sheet
                    .as_deref()
                    .is_some_and(|s| s.eq_ignore_ascii_case(old)) =>
            {
                out.push((r.start, r.start + r.prefix.len(), prefix))
            }
            Expr::QualifiedName(sheet, name, a, b) if sheet.eq_ignore_ascii_case(old) => {
                out.push((*a, *b, format!("{prefix}{name}")))
            }
            Expr::QualifiedError(sheet, a, b) if sheet.eq_ignore_ascii_case(old) => {
                out.push((*a, *b, prefix))
            }
            Expr::Unary(_, e) => collect(e, old, new, out)?,
            Expr::Binary(_, a, b) | Expr::Compare(_, a, b) => {
                collect(a, old, new, out)?;
                collect(b, old, new, out)?;
            }
            Expr::Call(name, args) => {
                if ["INDIRECT", "OFFSET"].contains(&name.as_str()) {
                    return Err(crate::coded(
                        "formula-unsupported",
                        "Dynamic sheet rename requires qualification",
                    ));
                }
                for a in args {
                    collect(a, old, new, out)?;
                }
            }
            Expr::Union(areas) => {
                for area in areas {
                    collect(area, old, new, out)?;
                }
            }
            _ => {}
        }
        Ok(())
    }
    collect(&ast, old, new, &mut patches)?;
    patches.sort_by_key(|p| p.0);
    let mut out = formula.to_string();
    for (a, b, text) in patches.into_iter().rev() {
        out.replace_range(a..b, &text);
    }
    Ok(out)
}
pub fn command(input: &Value) -> Result<Value> {
    // Transport and AST bounds complement the request-frame limit. A miss never
    // writes a version or calculation cache, so fallback is atomic.
    if serde_json::to_vec(input).map_err(crate::err)?.len() > 2 * 1024 * 1024 {
        return Ok(json!({"supported":false,"reason":"Native projection byte bound"}));
    }
    let result = match field(input, "method")? {
        "analyze" => analyze(&input["view"]),
        "evaluate" => evaluate(&input["view"]),
        "rewrite" => rewrite(input),
        "rename" => (|| {
            let mut formulas = Vec::new();
            for item in input["items"]
                .as_array()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing rename items"))?
            {
                crate::cancel::check()?;
                formulas.push(rename(
                    field(item, "formula")?,
                    field(input, "old")?,
                    field(input, "new")?,
                )?);
            }
            Ok(json!({"formulas":formulas}))
        })(),
        _ => {
            return Err(crate::coded(
                "internal-unknown-reference",
                "Unknown formula method",
            ));
        }
    };
    crate::cancel::check()?;
    Ok(match result {
        Ok(value) if serde_json::to_vec(&value).map_err(crate::err)?.len() > 2 * 1024 * 1024 => {
            json!({"supported":false,"reason":"Native result byte bound"})
        }
        Ok(mut value) => {
            value["supported"] = json!(true);
            value
        }
        Err(reason) => json!({"supported":false,"reason":reason}),
    })
}
#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn copy_boundaries_keep_valid_axes_and_create_reference_errors() {
        for (formula, rows, columns, expected) in [
            ("A1", -1, 0, "#REF!"),
            ("=A1", 0, -1, "=#REF!"),
            ("$A1+A$1+$A$1", -1, 0, "#REF!+A$1+$A$1"),
            ("$A1+A$1+$A$1", 0, -1, "$A1+#REF!+$A$1"),
            ("'Other Sheet'!A1", -1, 0, "'Other Sheet'!#REF!"),
            ("IFERROR(A1,99)", -1, 0, "IFERROR(#REF!,99)"),
            ("\"A1\"&$A$1", -1, -1, "\"A1\"&$A$1"),
            ("SUM(A:A)", 0, -1, "SUM(#REF!)"),
            ("SUM(1:1)", -1, 0, "SUM(#REF!)"),
            ("SUM($A:$A)+SUM($1:$1)", -1, -1, "SUM($A:$A)+SUM($1:$1)"),
            ("SUM(A:A)", -1, 0, "SUM(A:A)"),
            ("SUM(1:1)", 0, -1, "SUM(1:1)"),
            ("XFD1048576", 1, 0, "#REF!"),
            ("XFD1048576", 0, 1, "#REF!"),
        ] {
            let copied = rewrite(
                &json!({"items":[{"formula":formula,"copyRows":rows,"copyColumns":columns}]}),
            )
            .unwrap();
            assert_eq!(copied["formulas"][0], expected, "{formula}");
            assert!(Parser::references(expected).is_ok(), "{expected}");
        }
        for offset in [i64::MIN, i64::MAX] {
            assert_eq!(
                rewrite(&json!({"items":[{"formula":"A2","copyRows":offset}]})).unwrap()["formulas"]
                    [0],
                "#REF!"
            );
        }
        for offset in [json!("-1"), json!(-0.5), Value::Null, json!(true)] {
            assert!(rewrite(&json!({"items":[{"formula":"1+1","copyRows":offset}]})).is_err());
        }
    }
    #[test]
    fn broken_reference_syntax_keeps_surviving_spans_for_copy() {
        for (formula, rows, expected) in [
            ("SUM(#REF!:A2)", 1, "SUM(#REF!:A3)"),
            ("SUM(A2:#REF!)", -1, "SUM(A1:#REF!)"),
            (
                "SUM('Other Sheet'!#REF!:A2)",
                1,
                "SUM('Other Sheet'!#REF!:A3)",
            ),
            (
                "IFERROR('Other Sheet'!#REF!,99)",
                1,
                "IFERROR('Other Sheet'!#REF!,99)",
            ),
            ("SUM(#REF!:#REF!)", 1, "SUM(#REF!:#REF!)"),
        ] {
            assert!(Parser::references(formula).is_ok(), "{formula}");
            assert_eq!(
                rewrite(&json!({"items":[{"formula":formula,"copyRows":rows}]})).unwrap()["formulas"]
                    [0],
                expected
            );
        }
        for formula in [
            "A1:#DIV/0!",
            "#REF!:SUM(A1)",
            "A1:MissingName",
            "A1:B2:#REF!",
            "SUM((#REF!:Other!A1,A2))",
        ] {
            assert!(Parser::references(formula).is_err(), "{formula}");
        }
        // A reference into another workbook is carried through an edit
        // unchanged, in any script, and is still declined for calculation.
        for formula in [
            "'[book.xlsx]Data'!#REF!",
            "SUM('[1]Sheet 1'!$A$1:$B$2)+A1",
            "'[2]资产负债表'!C5",
        ] {
            assert!(Parser::references(formula).is_ok(), "{formula}");
            assert!(Parser::parse(formula).is_err(), "{formula}");
        }
        assert_eq!(
            rewrite(&json!({
                "target":"S","axis":"row","index":0,"count":1,"remove":false,
                "items":[{"formula":"'[1]Sheet 1'!$A$1+A1+资产负债表!B2","owner":"S"}]
            }))
            .unwrap()["formulas"][0],
            "'[1]Sheet 1'!$A$1+A2+资产负债表!B2"
        );
    }
    #[test]
    fn broken_ranges_propagate_reference_errors_without_reading_endpoints() {
        for formula in [
            "#REF!",
            "'Other Sheet'!#REF!",
            "SUM(#REF!:A2)",
            "SUM(A2:#REF!)",
            "SUM(#REF!:#REF!)",
        ] {
            let view = json!({"sheets":[{"name":"Data","cells":{"A2":{"formula":"A2"},"B1":{"formula":format!("IFERROR({formula},99)")}}}]});
            let ast = resolved(&format!("IFERROR({formula},99)"), &view, "Data", "B1").unwrap();
            let mut refs = Vec::new();
            visit(&ast, &mut refs, false, false).unwrap();
            assert!(
                refs.is_empty(),
                "A broken range cannot acquire endpoint dependencies"
            );
            let mut isolated = view.clone();
            isolated["sheets"][0]["cells"]
                .as_object_mut()
                .unwrap()
                .remove("A2");
            let result = evaluate(&isolated).unwrap();
            assert_eq!(
                result["results"][0]["value"].as_f64(),
                Some(99.0),
                "{formula}"
            );
        }
    }
    #[test]
    fn partly_invalid_copies_collapse_complete_references_like_excel_fill_up() {
        // Actual owned Excel formula-bar observations: source B2:J2 filled up
        // to B1:J1, retaining the sheet prefix but no surviving endpoint.
        for (formula, expected) in [
            ("SUM(A1:A2)", "SUM(#REF!)"),
            ("SUM($A1:A$3)", "SUM(#REF!)"),
            ("SUM(Sheet1!A1:A2)", "SUM(Sheet1!#REF!)"),
            ("SUM($A$1:A1)", "SUM(#REF!)"),
            ("SUM(A:A)", "SUM(A:A)"),
            ("SUM(1:2)", "SUM(#REF!)"),
            ("IFERROR(A1,99)", "IFERROR(#REF!,99)"),
            ("\"A1\"&$A$1", "\"A1\"&$A$1"),
            ("A1", "#REF!"),
        ] {
            assert_eq!(
                rewrite(&json!({"items":[{"formula":formula,"copyRows":-1}]})).unwrap()["formulas"]
                    [0],
                expected
            );
        }
    }
    #[test]
    fn existing_broken_ranges_keep_colon_like_excel_fill_left() {
        // Actual owned Excel A4:B8 Fill Left formula-bar observations. The
        // simple Sheet1 prefix was already normalized at formula entry.
        for (formula, expected) in [
            ("SUM(A1:B1)", "SUM(#REF!)"),
            ("SUM(A:B)", "SUM(#REF!)"),
            ("Sheet1!A1", "Sheet1!#REF!"),
            ("SUM(#REF!:A2)", "SUM(#REF!:#REF!)"),
            ("SUM(A2:#REF!)", "SUM(#REF!:#REF!)"),
            ("SUM(#REF!:B2)", "SUM(#REF!:A2)"),
            ("SUM(B2:#REF!)", "SUM(A2:#REF!)"),
        ] {
            assert_eq!(
                rewrite(&json!({"items":[{"formula":formula,"copyColumns":-1}]})).unwrap()["formulas"]
                    [0],
                expected
            );
            assert!(Parser::references(expected).is_ok(), "{expected}");
        }
    }
    #[test]
    fn qualified_error_rename_keeps_sheet_identity_and_exact_prefix_span() {
        for (formula, old, new, expected) in [
            ("Sheet1!#REF!", "Sheet1", "New Sheet", "'New Sheet'!#REF!"),
            (
                "SUM(A2:Sheet1!#REF!)",
                "Sheet1",
                "New Sheet",
                "SUM(A2:'New Sheet'!#REF!)",
            ),
            (
                "SUM('O''Brien'!#REF!:A2)",
                "O'Brien",
                "Changed'Sheet",
                "SUM('Changed''Sheet'!#REF!:A2)",
            ),
            (
                "Sheet1!#REF!+Other!#REF!+#REF!+\"Sheet1!#REF!\"",
                "Sheet1",
                "New",
                "'New'!#REF!+Other!#REF!+#REF!+\"Sheet1!#REF!\"",
            ),
            (
                "#REF!+\"Sheet1!#REF!\"",
                "Sheet1",
                "New",
                "#REF!+\"Sheet1!#REF!\"",
            ),
        ] {
            assert_eq!(rename(formula, old, new).unwrap(), expected);
            assert!(Parser::references(expected).is_ok());
        }
        assert!(
            Parser::references("{Sheet1!#REF!}").is_err(),
            "A qualified reference cannot become an array constant"
        );
        let view =
            json!({"sheets":[{"name":"Data","cells":{"B1":{"formula":"COUNTA(Sheet1!#REF!)"}}}]});
        assert_eq!(
            evaluate(&view).unwrap()["results"][0]["value"].as_f64(),
            Some(1.0)
        );
    }
    #[test]
    fn invalid_name_scope_never_resolves_as_global_or_local() {
        let view = json!({"sheets":[{"name":"Inputs"},{"name":"Report"}],"namedRanges":[{"name":"BadScope","formula":"Inputs!$A$1","scopeInvalid":true},{"name":"ForgedLocal","formula":"Inputs!$A$1","sheet":"Report","scopeInvalid":true}],"tables":[]});
        for text in ["BadScope", "IFERROR(BadScope,99)", "Report!ForgedLocal"] {
            assert!(resolved(text, &view, "Report", "C1").is_err(), "{text}");
            assert!(
                cell_dependencies(&view, "Report", "C1", &json!({"formula":text})).is_err(),
                "{text}"
            );
        }
        let traced = auditing::references(
            &view,
            "Report",
            "C1",
            &json!({"formula":"BadScope+Inputs!A2"}),
            false,
        )
        .unwrap();
        assert!(traced.unknown);
        assert_eq!(traced.ranges.len(), 1);
        assert_eq!(traced.ranges[0]["sheet"], "Inputs");
        assert_eq!(traced.ranges[0]["range"], "A2");
    }
    #[test]
    fn names_keep_local_case_shadow_quoted_scope_and_ambiguity() {
        let mut view = json!({"sheets":[{"name":"Inputs"},{"name":"Report"},{"name":"O'Brien Data"}],"namedRanges":[{"name":"Rate","formula":"Inputs!$A$1"},{"name":"rate","formula":"Inputs!$A$2","sheet":"Report"},{"name":"QuoteValue","formula":"'O''Brien Data'!$A$1","sheet":"O'Brien Data"}],"tables":[]});
        for (text, owner, sheet, reference) in [
            ("RaTe", "Inputs", "Inputs", "A1"),
            ("RATE", "Report", "Inputs", "A2"),
            ("'O''Brien Data'!QuoteValue", "Report", "O'Brien Data", "A1"),
        ] {
            let proof = cell_dependencies(&view, owner, "C1", &json!({"formula":text})).unwrap();
            assert_eq!(proof["ranges"], json!([{"sheet":sheet,"ref":reference}]));
        }
        assert!(resolved("QuoteValue", &view, "Report", "C1").is_err());
        view["namedRanges"]
            .as_array_mut()
            .unwrap()
            .push(json!({"name":"RATE","formula":"Inputs!$A$3","sheet":"Report"}));
        assert!(
            resolved("rate", &view, "Report", "C1")
                .unwrap_err()
                .contains("Ambiguous")
        );
        assert!(resolved("rate", &view, "Inputs", "C1").is_ok());
    }
    #[test]
    fn table_symbol_rename_preserves_selectors_strings_and_arithmetic() {
        for (before, after) in [
            (
                "SUM(Prices[Value])+Prices[[#This Row],[Value]]",
                "SUM(Rates[Value])+Rates[[#This Row],[Value]]",
            ),
            (
                "IF(\"Prices[Value]\"=\"literal\",prices[Value],SUM(Prices))",
                "IF(\"Prices[Value]\"=\"literal\",Rates[Value],SUM(Rates))",
            ),
            (
                "SUM(Other[Value])+Prices[#All]",
                "SUM(Other[Value])+Rates[#All]",
            ),
            ("( Prices[Value] + 1 )*2", "( Rates[Value] + 1 )*2"),
        ] {
            assert_eq!(rename_table(before, "Prices", "Rates").unwrap(), after);
        }
        for formula in [
            "INDIRECT(\"Prices[Value]\")",
            "OFFSET(A1,1,1)",
            "Sheet1!Prices",
            "Prices[Value",
        ] {
            assert!(
                rename_table(formula, "Prices", "Rates").is_err(),
                "{formula}"
            );
        }
    }
    #[test]
    fn column_rename_preserves_literals_unrelated_tokens_and_selector_spelling() {
        for (before, after, inside) in [
            (
                "SUM(Prices[Value])+SUM(PRICES[[#Data],[value]])",
                "SUM(Prices[Unit Price])+SUM(PRICES[[#Data],[Unit Price]])",
                false,
            ),
            (
                "SUM(Prices[[#All],[Value]])+Prices[[#This Row],[Value]]",
                "SUM(Prices[[#All],[Unit Price]])+Prices[[#This Row],[Unit Price]]",
                true,
            ),
            (
                "IF(\"Prices[Value]\"=\"literal\",SUM(Prices[Value]),SUM(Other[Value]))",
                "IF(\"Prices[Value]\"=\"literal\",SUM(Prices[Unit Price]),SUM(Other[Value]))",
                false,
            ),
            (
                "SUM([Value])+[@Value]*2",
                "SUM([Unit Price])+[@Unit Price]*2",
                true,
            ),
            (
                "( SUM(Prices[Cost]) + SUM(Prices) ) * 2",
                "( SUM(Prices[Cost]) + SUM(Prices) ) * 2",
                false,
            ),
        ] {
            assert_eq!(
                rename_table_column(before, "Prices", "Value", "Unit Price", inside).unwrap(),
                after
            );
        }
    }
    #[test]
    fn column_rename_refuses_dynamic_unbound_and_complex_target_selectors() {
        for formula in [
            "INDIRECT(\"Prices[Value]\")",
            "OFFSET(A1,1,1)",
            "SUM([Value])",
            "SUM(Prices[[Value]:[Cost]])",
            "SUM(Prices[[#Headers],[Value]])",
            "SUM(Prices[Val'ue])",
        ] {
            assert!(
                rename_table_column(formula, "Prices", "Value", "Unit Price", false).is_err(),
                "{formula}"
            );
        }
    }
    #[test]
    fn table_local_column_references_shift_shrink_and_keep_outside_spelling() {
        for (remove, index, before, after) in [
            (
                false,
                1,
                "SUM(A2:B4)+$C$2+'Data'!H8",
                "SUM(A2:C4)+$D$2+'Data'!H8",
            ),
            (false, 0, "'Data'!A2+B8", "Data!B2+B8"),
            (true, 1, "SUM(A2:C4)+$B$2+C2", "SUM(A2:B4)+#REF!+B2"),
            (true, 0, "A2+'Data'!C2", "#REF!+Data!B2"),
        ] {
            let out=rewrite(&json!({"items":[{"formula":before,"owner":"Data"}],"target":"Data","index":index,"count":1,"axis":"column","remove":remove,"tableRange":"A1:C4"})).unwrap();
            assert_eq!(out["formulas"][0], after);
        }
        for formula in ["SUM(A2:D4)", "SUM(A:A)", "SUM(A1:A8)"] {
            assert!(rewrite(&json!({"items":[{"formula":formula,"owner":"Data"}],"target":"Data","index":1,"count":1,"axis":"column","remove":false,"tableRange":"A1:C4"})).is_err(),"{formula}");
        }
        assert_eq!(
            rename_table_column(
                "SUM(Prices[Cost])+\"Prices[Cost]\"+SUM(Other[Cost])",
                "Prices",
                "Cost",
                "",
                false
            )
            .unwrap(),
            "SUM(#REF!)+\"Prices[Cost]\"+SUM(Other[Cost])"
        );
    }
    fn raw_union_formula(text: &str) -> Expr {
        let mut parser = Parser {
            text,
            at: 0,
            depth: 0,
            nodes: 0,
            node_limit: 8192,
        };
        let expr = parser.expr(0).unwrap();
        parser.space();
        assert_eq!(parser.at, text.len());
        expr
    }
    #[test]
    fn union_areas_keep_order_duplicates_and_function_argument_boundaries() {
        let text = "SUM((A3,A1:A2,A1,($B$2,A1)),C1)";
        let ast = raw_union_formula(text);
        union_context(&ast, false).unwrap();
        let Expr::Call(name, args) = &ast else {
            panic!("Expected SUM")
        };
        assert_eq!(name, "SUM");
        assert_eq!(args.len(), 2);
        let Expr::Union(areas) = &args[0] else {
            panic!("Expected one union argument")
        };
        assert_eq!(areas.len(), 4);
        let mut refs = Vec::new();
        visit(&ast, &mut refs, false, true).unwrap();
        assert_eq!(
            refs.iter()
                .map(|(r, _)| &text[r.start..r.end])
                .collect::<Vec<_>>(),
            vec!["A3", "A1:A2", "A1", "$B$2", "A1", "C1"]
        );
        assert!(refs.iter().all(|(_, aggregate)| *aggregate));
        // Source spelling and parenthesis boundaries remain the rewriter's
        // authority; the AST never sorts or deduplicates area references.
        assert!(contains_union(&ast));
    }
    #[test]
    fn union_syntax_does_not_qualify_scalar_array_named_or_other_function_contexts() {
        for formula in [
            "(A1,A2)",
            "(A1,A2)+1",
            "SUM((A1,1))",
            "SUM((A1,A2*2))",
            "SUM((A1,Rate))",
            "SUM((A1,Table1[Cost]))",
            "AVERAGE((A1,A2))",
            "ABS((A1,A2))",
            "SUM(ABS((A1,A2)))",
            "SUM((A1,{1,2}))",
            "SUM((A1,'Other Sheet'!B2))",
        ] {
            assert!(
                union_context(&raw_union_formula(formula), false).is_err(),
                "{formula}"
            );
        }
        for formula in ["SUM((A1,A2))", "SUM((A:A,3:4))", "SUM((#REF!,A1))"] {
            let ast = raw_union_formula(formula);
            union_context(&ast, false).unwrap();
            assert!(Parser::parse(formula).is_ok(), "{formula}");
        }
        for formula in [
            "SUM((A1,))",
            "SUM((,A1))",
            "SUM((A1;A2))",
            "SUM(([book.xlsx]Data!A1,A2))",
            "SUM(('Data:Other'!A1,A2))",
        ] {
            assert!(Parser::parse(formula).is_err(), "{formula}");
        }
    }
    #[test]
    fn constant_cache_excludes_context_sensitive_syntax() {
        for formula in ["1+1", "SUM(1,2)", "IF(TRUE(),3,4)", "DATE(1900,2,29)"] {
            assert!(
                constant_expression(&Parser::parse(formula).unwrap()),
                "{formula}"
            );
        }
        for formula in [
            "A1+1",
            "$A$1",
            "Rate+1",
            "ROW()",
            "TODAY()",
            "RAND()",
            "INDIRECT(\"A1\")",
            "IF(TRUE(),1,ROW())",
        ] {
            assert!(
                !constant_expression(&Parser::parse(formula).unwrap()),
                "{formula}"
            );
        }
    }
    #[test]
    fn compensation_does_not_zero_large_finite_results() {
        assert_eq!(super::compensated(f64::MAX, -1.0), f64::MAX);
        assert_eq!(super::compensated(-f64::MAX, 1.0), -f64::MAX);
    }
    #[test]
    fn grammar_and_guards() {
        for f in [
            "SUM(A1:B4)+$C$1*2",
            "-'O''Brien'!$B2/4%",
            "1.2e-3+A1",
            "SUM1+1",
            "Table1[Cost]",
            "A:A",
            "SUM(A1,)", // Reference planning accepts omitted arguments; evaluation remains guarded.
        ] {
            assert!(Parser::parse(f).is_ok(), "{f}")
        }
        for f in ["A1 B1", "Table1[Cost", "[book]Sheet1!A1"] {
            assert!(Parser::parse(f).is_err(), "{f}")
        }
    }
    #[test]
    fn invariant_copy_is_syntax_only_and_matches_rewriter() {
        for formula in [
            "1+1",
            "SUM(1,2)",
            "IF(TRUE,1,2)",
            "INDIRECT(\"A1\")",
            "Table1[Cost]",
            "TaxRate",
            "\"A1\"",
        ] {
            assert!(copy_invariant(formula).unwrap(), "{formula}");
            let shifted =
                rewrite(&json!({"items":[{"formula":formula,"copyRows":10,"copyColumns":2}]}))
                    .unwrap();
            assert_eq!(shifted["formulas"][0], formula);
        }
        for formula in ["A1", "$A$1", "SUM(A1:A9)", "'Other Sheet'!A1", "SUM(1:1)"] {
            assert!(!copy_invariant(formula).unwrap(), "{formula}");
        }
        assert!(copy_invariant("A1 B1").is_err());
    }
    #[test]
    fn insertion_and_copy() {
        let r=rewrite(&json!({"target":"Data","index":4,"count":2,"items":[{"formula":"SUM($A$1:A4)+SUM(B1:B4)","owner":"Data","boundary":true},{"formula":"Other!A5+Data!$B$5","owner":"Data"},{"formula":"$A1+B$2+$C$3","copyRows":2}]})).unwrap();
        assert_eq!(
            r["formulas"],
            json!([
                "SUM($A$1:A6)+SUM(B1:B6)",
                "Other!A5+Data!$B$7",
                "$A3+B$2+$C$3"
            ])
        )
    }
}

#[cfg(test)]
mod named_edit_tests {
    use super::*;
    #[test]
    fn static_consumers_respect_scopes_and_do_not_use_strings_as_names() {
        assert!(!static_name_use("SUM(Inputs)", "Inputs", "Sheet1", None, true).unwrap());
        assert!(static_name_use("SUM(Inputs)", "Inputs", "Other", None, false).unwrap());
        assert!(
            static_name_use(
                "SUM(Sheet1!Inputs)",
                "Inputs",
                "Other",
                Some("Sheet1"),
                false
            )
            .unwrap()
        );
        assert!(!static_name_use("SUM(Sheet1!Inputs)", "Inputs", "Other", None, false).unwrap());
        assert!(
            !static_name_use("LEN(\"Inputs\")", "Inputs", "Sheet1", Some("Sheet1"), false).unwrap()
        );
        assert!(
            static_name_use(
                "IF(FALSE,Inputs,1)",
                "Inputs",
                "Sheet1",
                Some("Sheet1"),
                false
            )
            .unwrap()
        );
        assert!(static_name_use("INDIRECT(\"Inputs\")", "Unused", "Sheet1", None, false).is_err());
    }
    #[test]
    fn one_cell_authored_name_resolves_as_scalar_but_real_rectangles_do_not_collapse() {
        let view = json!({"sheets":[{"name":"Data"}],"namedRanges":[{"name":"Point","formula":"'Data'!$A$1:$A$1"},{"name":"Area","formula":"'Data'!$A$1:$B$2"}],"tables":[]});
        let Expr::Ref(point) = resolved("Point", &view, "Data", "C1").unwrap() else {
            panic!("Ref")
        };
        assert_eq!(point.last, None);
        assert_eq!(point.first, (0, 0));
        let Expr::Ref(area) = resolved("Area", &view, "Data", "C1").unwrap() else {
            panic!("Ref")
        };
        assert_eq!(area.last, Some((1, 1)));
    }
    #[test]
    fn per_formula_errors_do_not_convert_error_looking_text_or_mask_subset_misses() {
        let view = json!({"sheets":[{"name":"Sheet1","cells":{"A1":{"value":"#N/A","dataType":"text"},"B1":{"formula":"1/0"},"C1":{"formula":"IFERROR(B1,7)"},"D1":{"formula":"COUNTIF(A1:A1,\"*\")"}}}]});
        let result = evaluate(&view).unwrap();
        let rows = result["results"].as_array().unwrap();
        assert_eq!(
            rows.iter().find(|r| r["ref"] == "B1").unwrap()["error"],
            "#DIV/0!"
        );
        assert_eq!(
            rows.iter().find(|r| r["ref"] == "C1").unwrap()["value"],
            7.0
        );
        assert_eq!(
            rows.iter().find(|r| r["ref"] == "D1").unwrap()["value"],
            1.0
        );
        let mut unknown = view;
        unknown["sheets"][0]["cells"]["E1"] = json!({"formula":"VendorFunction(A1)"});
        assert!(evaluate(&unknown).is_err());
    }
}

/// A byte that can continue a sheet name or a defined name written without
/// quotes. Every byte of a non-ASCII character counts, so names in any script
/// (资产负债表!A1) are read whole; Excel quotes the names that need it.
fn word(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'.') || byte >= 0x80
}
