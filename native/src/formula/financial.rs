//! PRIVATE CANDIDATE. Numeric scalar arithmetic; vendor bits remain unqualified.
use super::*;
pub(crate) const REASON: &str = "Financial formula is outside the bounded numeric scalar profile";
pub(crate) fn calculate(name: &str, input: &[f64]) -> Result<f64> {
    if !["PMT", "PV", "FV"].contains(&name) || !(3..=5).contains(&input.len()) {
        return Err(crate::coded(
            "formula-unsupported",
            format!("{REASON}: argument count"),
        ));
    }
    let mut args = [0.0; 5];
    for (i, &n) in input.iter().enumerate() {
        if !n.is_finite()
            || n == 0.0 && n.is_sign_negative()
            || i == 0 && !(0.0..=1.0).contains(&n)
            || i == 1 && (n.fract() != 0.0 || !(1.0..=1200.0).contains(&n))
            || [2, 3].contains(&i) && n.abs() > 1e12
            || i == 4 && n != 0.0 && n != 1.0
        {
            return Err(crate::coded(
                "formula-unsupported",
                format!("{REASON}: input domain/type"),
            ));
        }
        args[i] = n;
    }
    let [rate, nper, flow, balance, timing] = args;
    let exponent = nper * rate.ln_1p();
    let (growth, annuity) = if rate == 0.0 {
        (1.0, nper)
    } else if rate == 1.0 {
        let growth = 2_f64.powi(nper as i32);
        (growth, growth - 1.0)
    } else {
        (exponent.exp(), exponent.exp_m1() / rate)
    };
    if !growth.is_finite() || !annuity.is_finite() {
        return Err(crate::coded(
            "formula-limit",
            format!("{REASON}: intermediate overflow"),
        ));
    }
    let due = 1.0 + rate * timing;
    let value = match name {
        "PMT" => -(flow * growth + balance) / (due * annuity),
        "PV" => -(flow * due * annuity + balance) / growth,
        _ => -(balance * growth + flow * due * annuity),
    };
    if !value.is_finite() || value.abs() > 9007199254740991.0 {
        return Err(crate::coded(
            "formula-unsupported",
            format!("{REASON}: result magnitude"),
        ));
    }
    Ok(if value == 0.0 { 0.0 } else { value })
}
fn operand_shape(expr: &Expr, depth: usize) -> Result<(usize, usize)> {
    if depth > 8 {
        return Err(crate::coded(
            "formula-unsupported",
            format!("{REASON}: operand depth"),
        ));
    }
    match expr {
        Expr::Number(_) => Ok((1, 0)),
        Expr::Ref(r) if r.sheet.is_none() && r.last.is_none() => Ok((1, 1)),
        Expr::Unary(op, e) if [b'+', b'-'].contains(op) => {
            let (n, r) = operand_shape(e, depth + 1)?;
            Ok((n + 1, r))
        }
        Expr::Binary(op, a, b) if [b'+', b'-', b'*', b'/'].contains(op) => {
            let (na, ra) = operand_shape(a, depth + 1)?;
            let (nb, rb) = operand_shape(b, depth + 1)?;
            Ok((na + nb + 1, ra + rb))
        }
        _ => Err(crate::coded(
            "formula-unsupported",
            format!("{REASON}: closed arithmetic operand grammar"),
        )),
    }
}
fn argument(token: &str) -> Result<Expr> {
    if token.len() > 1024 {
        return Err(crate::coded(
            "formula-limit",
            format!("{REASON}: operand byte bound"),
        ));
    }
    // Parser owns the authoritative grammar/tree. This independent lexical pass
    // retains original numeric spelling, which an f64 AST alone cannot prove.
    let expr = Parser::parse(token)?;
    let b = token.as_bytes();
    let mut at = 0;
    let mut nesting = 0;
    while at < b.len() {
        let c = b[at];
        if c.is_ascii_whitespace() || [b'+', b'-', b'*', b'/'].contains(&c) {
            at += 1;
            continue;
        }
        if c == b'(' {
            nesting += 1;
            if nesting > 8 {
                return Err(crate::coded(
                    "formula-unsupported",
                    format!("{REASON}: operand depth"),
                ));
            }
            at += 1;
            continue;
        }
        if c == b')' {
            nesting -= 1;
            at += 1;
            continue;
        }
        if c.is_ascii_digit() || c == b'.' {
            let start = at;
            while b.get(at).is_some_and(u8::is_ascii_digit) {
                at += 1;
            }
            if b.get(at) == Some(&b'.') {
                at += 1;
                while b.get(at).is_some_and(u8::is_ascii_digit) {
                    at += 1;
                }
            }
            if b.get(at).is_some_and(|c| [b'e', b'E'].contains(c)) {
                at += 1;
                if b.get(at).is_some_and(|c| [b'+', b'-'].contains(c)) {
                    at += 1;
                }
                while b.get(at).is_some_and(u8::is_ascii_digit) {
                    at += 1;
                }
            }
            let text = &token[start..at];
            let value = text.parse::<f64>().map_err(|_| {
                crate::coded("formula-unsupported", format!("{REASON}: numeric token"))
            })?;
            if !value.is_finite()
                || value.abs() > 9007199254740991.0
                || crate::bulk::meaning(text) != crate::bulk::meaning(&value.to_string())
            {
                return Err(crate::coded(
                    "formula-unsupported",
                    format!("{REASON}: numeric literal precision"),
                ));
            }
            continue;
        }
        if c.is_ascii_uppercase() || c == b'$' {
            let start = at;
            if c == b'$' {
                at += 1;
            }
            let letters = at;
            while b.get(at).is_some_and(u8::is_ascii_uppercase) {
                at += 1;
            }
            if at == letters || at - letters > 3 {
                return Err(crate::coded(
                    "formula-unsupported",
                    format!("{REASON}: reference token"),
                ));
            }
            if b.get(at) == Some(&b'$') {
                at += 1;
            }
            let digits = at;
            while b.get(at).is_some_and(u8::is_ascii_digit) {
                at += 1;
            }
            if at == digits || b[digits] == b'0' {
                return Err(crate::coded(
                    "formula-unsupported",
                    format!("{REASON}: reference token"),
                ));
            }
            let parsed = Parser::parse(&token[start..at])?;
            if !matches!(
                parsed,
                Expr::Ref(Ref {
                    sheet: None,
                    last: None,
                    ..
                })
            ) {
                return Err(crate::coded(
                    "formula-unsupported",
                    format!("{REASON}: reference token"),
                ));
            }
            continue;
        }
        return Err(crate::coded(
            "formula-unsupported",
            format!("{REASON}: closed arithmetic operand grammar"),
        ));
    }
    let (nodes, _) = operand_shape(&expr, 1)?;
    if nodes > 32 {
        return Err(crate::coded(
            "formula-limit",
            format!("{REASON}: operand node bound"),
        ));
    }
    Ok(expr)
}
/// Authored closed tokens are checked before name expansion or branch selection.
pub(crate) fn calls(formula: &str) -> Result<Vec<(String, Vec<Expr>)>> {
    let b = formula.as_bytes();
    let mut out = Vec::new();
    let mut i = 0;
    let mut quote = 0;
    let mut square = 0;
    while i < b.len() {
        let c = b[i];
        if quote != 0 {
            if c == quote {
                if b.get(i + 1) == Some(&quote) {
                    i += 1;
                } else {
                    quote = 0;
                }
            }
            i += 1;
            continue;
        }
        if c == b'"' || c == b'\'' {
            quote = c;
            i += 1;
            continue;
        }
        if c == b'[' {
            square += 1;
        }
        if c == b']' {
            square -= 1;
        }
        if square != 0 {
            i += 1;
            continue;
        }
        if !c.is_ascii_alphabetic() && c != b'_' {
            i += 1;
            continue;
        }
        let start = i;
        while b
            .get(i)
            .is_some_and(|c| c.is_ascii_alphanumeric() || [b'_', b'.'].contains(c))
        {
            i += 1;
        }
        let name = formula[start..i].to_ascii_uppercase();
        let mut open = i;
        while b.get(open).is_some_and(u8::is_ascii_whitespace) {
            open += 1;
        }
        if !["PMT", "PV", "FV", "_XLFN.PMT", "_XLFN.PV", "_XLFN.FV"].contains(&name.as_str())
            || b.get(open) != Some(&b'(')
        {
            continue;
        }
        if name.starts_with('_') {
            return Err(crate::coded(
                "formula-unsupported",
                format!("{REASON}: unqualified alias"),
            ));
        }
        let mut at = open + 1;
        let mut start = at;
        let mut depth = 1;
        let mut q = 0;
        let mut args = Vec::new();
        let mut close = None;
        while at < b.len() {
            let c = b[at];
            if q != 0 {
                if c == q {
                    if b.get(at + 1) == Some(&q) {
                        at += 1;
                    } else {
                        q = 0;
                    }
                }
                at += 1;
                continue;
            }
            if c == b'"' || c == b'\'' {
                q = c;
                at += 1;
                continue;
            }
            if c == b'(' {
                depth += 1;
            }
            if c == b')' {
                depth -= 1;
                if depth == 0 {
                    args.push(argument(&formula[start..at])?);
                    close = Some(at);
                    break;
                }
            }
            if c == b',' && depth == 1 {
                args.push(argument(&formula[start..at])?);
                start = at + 1;
            }
            at += 1;
        }
        i = close.ok_or_else(|| {
            crate::coded(
                "formula-unsupported",
                format!("{REASON}: unterminated call"),
            )
        })? + 1;
        if !(3..=5).contains(&args.len()) {
            return Err(crate::coded(
                "formula-unsupported",
                format!("{REASON}: argument count"),
            ));
        }
        let shapes = args
            .iter()
            .map(|arg| operand_shape(arg, 1))
            .collect::<Result<Vec<_>>>()?;
        if shapes.iter().map(|s| s.0).sum::<usize>() > 128
            || shapes.iter().map(|s| s.1).sum::<usize>() > 32
        {
            return Err(crate::coded(
                "formula-limit",
                format!("{REASON}: call node/reference bound"),
            ));
        }
        out.push((name, args));
    }
    Ok(out)
}
pub(crate) fn count(expr: &Expr) -> usize {
    match expr {
        Expr::Call(name, args) => {
            usize::from(["PMT", "PV", "FV"].contains(&name.as_str()))
                + args.iter().map(count).sum::<usize>()
        }
        Expr::Unary(_, e) => count(e),
        Expr::Binary(_, a, b) | Expr::Compare(_, a, b) => count(a) + count(b),
        Expr::Union(areas) => areas.iter().map(count).sum(),
        _ => 0,
    }
}
/// Independent version-bound source admission for the coordinator's mixed path.
/// Snapshot tables are TEMP; no calculation result or persistent version is written.
pub(crate) fn profiles(root: &std::path::Path, input: &Value) -> Result<Value> {
    let version = input["version"].as_u64().ok_or_else(|| {
        crate::coded(
            "internal-missing-field",
            "Missing financial profile version",
        )
    })?;
    let items = input["items"]
        .as_array()
        .filter(|items| items.len() <= 128)
        .ok_or_else(|| {
            crate::coded(
                "formula-limit",
                "Financial profile page exceeds128 formulas",
            )
        })?;
    let db = crate::database(root)?;
    crate::calculation::snapshot(&db, version)?;
    let view = &input["view"];
    let mut eval = Eval::from_sqlite(view, &db, false);
    let mut results = Vec::new();
    for item in items {
        crate::cancel::check()?;
        let path = field(item, "sheet")?;
        let reference = field(item, "ref")?;
        position(reference)?;
        let owner = view["sheets"]
            .as_array()
            .ok_or_else(|| {
                crate::coded("internal-missing-field", "Missing financial profile sheets")
            })?
            .iter()
            .find(|s| s["path"] == path)
            .ok_or_else(|| {
                crate::coded("internal-missing-field", "Financial worksheet is missing")
            })?;
        let name = field(owner, "name")?;
        let cell = crate::calculation::input_cell(&db, path, reference)?;
        let cell = crate::calculation::resolved_cell(&db, path, reference, &cell)?;
        let formula = field(&cell, "formula")?;
        let result = (|| {
            if cell.get("formulaType").is_some()
                || cell.get("spill").is_some()
                || cell["scenarioResult"] == true
            {
                return Err(crate::coded(
                    "formula-unsupported",
                    format!("{REASON}: special output"),
                ));
            }
            let expr = resolved(formula, view, name, reference)?;
            eval.financial_preflight(formula, name, &expr)
        })();
        results.push(match result{Ok(())=>json!({"sheet":path,"ref":reference,"formula":formula,"supported":true}),Err(reason)=>json!({"sheet":path,"ref":reference,"formula":formula,"supported":false,"reason":reason})});
    }
    Ok(json!({"version":version,"results":results}))
}
impl Eval<'_> {
    fn financial_reference_guard(
        &self,
        db: &rusqlite::Connection,
        sheet: &str,
        at: (u32, u32),
    ) -> Result<()> {
        // Financial profile pages and calculate own different connections. Build an
        // independent TEMP index from this accepted snapshot; never assume the spill
        // inspection connection's scratch table exists or borrow a follower cache.
        let exists:bool=db.prepare_cached("SELECT EXISTS(SELECT 1 FROM sqlite_temp_master WHERE type='table' AND name='financial_protected')").map_err(crate::err)?.query_row([],|row|row.get(0)).map_err(crate::err)?;
        if !exists {
            db.execute_batch("CREATE TEMP TABLE financial_protected(sheet TEXT,r0 INTEGER,c0 INTEGER,r1 INTEGER,c1 INTEGER);").map_err(crate::err)?;
            let version: u64 = db
                .prepare_cached("SELECT version FROM shared_context")
                .map_err(crate::err)?
                .query_row([], |row| row.get(0))
                .map_err(crate::err)?;
            for owner in self.view["sheets"]
                .as_array()
                .ok_or_else(|| crate::coded("internal-missing-field", "Missing financial sheets"))?
            {
                let path = field(owner, "path")?;
                for group in crate::spill::effective(db, version, path)? {
                    let g = group.geometry;
                    db.execute(
                        "INSERT INTO financial_protected VALUES(?1,?2,?3,?4,?5)",
                        rusqlite::params![path, g.first.0, g.first.1, g.last.0, g.last.1],
                    )
                    .map_err(crate::err)?;
                }
                let mut after = None;
                loop {
                    let page =
                        crate::calculation::page(db, path, (0, 0), (1048575, 16383), after, true)?;
                    for (reference, cell) in &page {
                        if cell.get("formulaType").is_some() {
                            let (a, b) = crate::bulk::bounds(
                                cell["formulaReference"].as_str().unwrap_or(reference),
                            )?;
                            db.execute(
                                "INSERT INTO financial_protected VALUES(?1,?2,?3,?4,?5)",
                                rusqlite::params![path, a.0, a.1, b.0, b.1],
                            )
                            .map_err(crate::err)?;
                        }
                    }
                    if page.len() < 512 {
                        break;
                    }
                    after = Some(position(&page.last().unwrap().0)?);
                }
            }
        }
        let protected:bool=db.prepare_cached("SELECT EXISTS(SELECT 1 FROM financial_protected WHERE sheet=?1 AND r0<=?2 AND r1>=?2 AND c0<=?3 AND c1>=?3)").map_err(crate::err)?.query_row(rusqlite::params![sheet,at.0,at.1],|row|row.get(0)).map_err(crate::err)?;
        if protected {
            return Err(crate::coded(
                "formula-unsupported",
                format!("{REASON}: special-group input"),
            ));
        }
        Ok(())
    }
    fn financial_input(&mut self, arg: &Expr, sheet: &str) -> Result<f64> {
        self.tick()?;
        let n = match arg {
            Expr::Number(n) => Ok(*n),
            Expr::Unary(op, e) if [b'+', b'-'].contains(op) => {
                let n = self.financial_input(e, sheet)?;
                Ok(if *op == b'-' { -n } else { n })
            }
            Expr::Binary(op, a, b) if [b'+', b'-', b'*', b'/'].contains(op) => {
                let a = self.financial_input(a, sheet)?;
                let b = self.financial_input(b, sheet)?;
                if *op == b'/' && b == 0.0 {
                    return Err(crate::coded(
                        "formula-unsupported",
                        format!("{REASON}: zero-divide operand"),
                    ));
                }
                let value = match op {
                    b'+' => a + b,
                    b'-' => a - b,
                    b'*' => a * b,
                    _ => a / b,
                };
                if [b'*', b'/'].contains(op) && a != 0.0 && b != 0.0 && value == 0.0 {
                    return Err(crate::coded(
                        "formula-unsupported",
                        format!("{REASON}: underflow operand"),
                    ));
                }
                Ok(value)
            }
            Expr::Ref(r) if r.sheet.is_none() && r.last.is_none() => {
                let reference = address(r.first.0, r.first.1);
                let cell = if let Some(db) = self.db {
                    let path = self.path(sheet)?;
                    self.financial_reference_guard(db, &path, r.first)?;
                    let root = std::path::Path::new(db.path().ok_or_else(|| {
                        crate::coded(
                            "internal-missing-field",
                            "Financial source database path is missing",
                        )
                    })?)
                    .parent()
                    .ok_or_else(|| {
                        crate::coded("internal-missing-field", "Financial source root is missing")
                    })?;
                    if crate::spill::source_attributes(root, db, &path, r.first)?.is_some_and(
                        |attrs| {
                            attrs
                                .keys()
                                .any(|key| !["r", "s", "t", "xmlns"].contains(&key.as_str()))
                                || attrs
                                    .get("xmlns")
                                    .is_some_and(|ns| ns.as_bytes() != crate::xml::NS)
                        },
                    ) {
                        return Err(crate::coded(
                            "formula-unsupported",
                            format!("{REASON}: opaque reference attributes"),
                        ));
                    }
                    crate::spill::source_content_proof(root, db, &path, r.first, true)?;
                    crate::calculation::input_cell(db, &path, &reference)?
                } else {
                    self.view["sheets"]
                        .as_array()
                        .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheets"))?
                        .iter()
                        .find(|s| s["name"] == sheet)
                        .ok_or_else(|| crate::coded("internal-missing-field", "Missing sheet"))?["cells"][&reference]
                        .clone()
                };
                if cell.get("formula").is_some()
                    || cell.get("formulaType").is_some()
                    || cell.get("spill").is_some()
                    || cell["scenarioResult"] == true
                    || cell["financialInputBlocked"] == true
                    || cell["dataType"] != "number"
                {
                    return Err(crate::coded(
                        "formula-unsupported",
                        format!("{REASON}: reference needs an exact ordinary numeric literal"),
                    ));
                }
                if cell["rawValue"]
                    .as_str()
                    .and_then(|s| s.parse::<f64>().ok())
                    .is_some_and(|n| n == 0.0 && n.is_sign_negative())
                {
                    return Err(crate::coded(
                        "formula-unsupported",
                        format!("{REASON}: negative-zero source"),
                    ));
                }
                crate::bulk::literal(&cell, false)?.as_f64().ok_or_else(|| {
                    crate::coded("formula-unsupported", format!("{REASON}: reference type"))
                })
            }
            _ => Err(crate::coded(
                "formula-unsupported",
                format!("{REASON}: argument grammar"),
            )),
        }?;
        if !n.is_finite() || n.abs() > 9007199254740991.0 || n == 0.0 && n.is_sign_negative() {
            return Err(crate::coded(
                "formula-unsupported",
                format!("{REASON}: nonfinite/unsafe/negative-zero operand"),
            ));
        }
        Ok(n)
    }
    pub(super) fn financial(&mut self, name: &str, args: &[Expr], sheet: &str) -> Result<f64> {
        let values = args
            .iter()
            .map(|arg| self.financial_input(arg, sheet))
            .collect::<Result<Vec<_>>>()?;
        calculate(name, &values)
    }
    pub(super) fn financial_preflight(
        &mut self,
        formula: &str,
        sheet: &str,
        expr: &Expr,
    ) -> Result<()> {
        let calls = calls(formula)?;
        if count(expr) != calls.len() {
            return Err(crate::coded(
                "formula-unsupported",
                format!("{REASON}: financial name expansion"),
            ));
        }
        for (name, args) in calls {
            self.financial(&name, &args, sheet)?;
        }
        Ok(())
    }
}
#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn exact_zero_and_rational_cases() {
        assert_eq!(calculate("PMT", &[0., 12., 1200.]).unwrap(), -100.);
        assert_eq!(calculate("PV", &[0., 12., -100.]).unwrap(), 1200.);
        assert_eq!(
            calculate("FV", &[0., 12., -100., 1200.]).unwrap().to_bits(),
            0
        );
        for (name, args, wanted) in [
            ("PMT", vec![0.5, 2., 100., 25., 0.], -100.),
            ("PV", vec![0.5, 2., -100., 25., 0.], 100.),
            ("FV", vec![0.5, 2., -100., 100., 1.], 150.),
        ] {
            assert_eq!(calculate(name, &args).unwrap(), wanted);
        }
    }
    #[test]
    fn unsupported_never_returns_excel_error() {
        for args in [
            vec![-0., 12., 100.],
            vec![-1., 12., 100.],
            vec![0., 2.5, 100.],
            vec![0., 12., 100., 0., 2.],
            vec![1., 1200., 100.],
        ] {
            let miss = calculate("PMT", &args).unwrap_err();
            assert!(!miss.starts_with("Excel error:"));
        }
    }
    #[test]
    fn authored_grammar_does_not_expand_or_coerce() {
        for f in [
            "PMT(0,12,1200)",
            "IFERROR(PMT($A$1,A2,A3),99)",
            "FV(0.5,2,-100,100,1)",
            "PMT(A1/12,A2*12,A3)",
        ] {
            assert_eq!(calls(f).unwrap().len(), 1);
        }
        for f in [
            "PMT(0,12,N)",
            "PMT(0,12,Other!A1)",
            "PMT(0,12,A1:A2)",
            "PMT(0,12,1200,,0)",
            "PMT(0.1234567890123456789012345,12,100)",
            "PMT(0,12,900719925474.0993)",
            "_xlfn.PMT(0,12,100)",
        ] {
            assert!(calls(f).is_err(), "{f}");
        }
        assert!(calls("\"PMT(0,12,100)\"").unwrap().is_empty());
    }
    #[test]
    fn arithmetic_operands_keep_order_and_closed_bounds() {
        let base = json!({"sheets":[{"name":"S","cells":{"A1":{"value":0,"dataType":"number"},"A2":{"value":1,"dataType":"number"},"A3":{"value":1200,"dataType":"number"},"B1":{"formula":"PMT(A1/12,A2*12,A3)"}}}]});
        assert_eq!(evaluate(&base).unwrap()["results"][0]["value"], -100.);
        let mut ordered = base.clone();
        ordered["sheets"][0]["cells"]["B1"]["formula"] = json!("PMT(0,12,0.3-0.2-0.1)");
        let expected = -(0.3_f64 - 0.2 - 0.1) / 12.0;
        assert_eq!(
            evaluate(&ordered).unwrap()["results"][0]["value"]
                .as_f64()
                .unwrap()
                .to_bits(),
            expected.to_bits()
        );
        for formula in [
            "PMT(A1^2,12,100)",
            "PMT(A1%,12,100)",
            "PMT(SUM(A1),12,100)",
            "PMT((A1,A2),12,100)",
            "PMT(0.1234567890123456789012345+0,12,100)",
            "PMT((((((((((A1)))))))))),12,100)",
        ] {
            assert!(calls(formula).is_err(), "{formula}");
        }
        assert!(calls(&format!("PMT({},12,100)", vec!["A1"; 18].join("+"))).is_err());
    }
    #[test]
    fn arithmetic_misses_precede_iferror_and_lazy_selection() {
        for operand in [
            "A1/0",
            "1e-200*1e-200",
            "-(A1-A1)",
            "9007199254740991+1",
            "9007199254740991*9007199254740991",
            "-0",
        ] {
            let formula = format!("IF(TRUE,7,IFERROR(PMT({operand},12,100),99))");
            let view = json!({"sheets":[{"name":"S","cells":{"A1":{"value":0,"dataType":"number"},"B1":{"formula":formula}}}]});
            let reason = evaluate(&view).unwrap_err();
            assert!(!reason.starts_with("Excel error:"), "{operand}: {reason}");
        }
        let view = json!({"sheets":[{"name":"S","cells":{"A1":{"formula":"0","value":0,"dataType":"number"},"B1":{"formula":"PMT(A1+0,12,100)"}}}]});
        assert!(evaluate(&view).is_err());
    }
}
