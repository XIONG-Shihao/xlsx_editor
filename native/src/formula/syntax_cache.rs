//! Bounded syntax reuse for one calculation. No values or resolved names live
//! here; cloned trees are resolved against each cell's own workbook context.
use super::{Expr, Parser};
use crate::Result;
use std::cell::RefCell;
use std::collections::HashMap;

#[derive(Default)]
struct Cache {
    formulas: HashMap<String, Expr>,
    references: HashMap<String, Expr>,
    bytes: usize,
    entries: usize,
    hits: u64,
}
thread_local! { static CACHE: RefCell<Option<Cache>> = const { RefCell::new(None) }; }
pub(crate) struct Guard(Option<Cache>);
pub(crate) fn begin() -> Guard {
    Guard(CACHE.with(|cache| cache.replace(Some(Cache::default()))))
}
impl Drop for Guard {
    fn drop(&mut self) {
        let current = CACHE.with(|cache| cache.replace(self.0.take()));
        if let Some(current) = current {
            crate::diagnostics::count("calculation.syntax_cache_hits", current.hits);
            crate::diagnostics::count("calculation.syntax_cache_entries", current.entries as u64);
            crate::diagnostics::count("calculation.syntax_cache_bytes", current.bytes as u64);
        }
    }
}
fn bytes(expr: &Expr) -> usize {
    std::mem::size_of::<Expr>()
        + match expr {
            Expr::String(s)
            | Expr::Error(s)
            | Expr::QualifiedError(s, ..)
            | Expr::Name(s, ..)
            | Expr::Structured(s, ..) => s.capacity(),
            Expr::QualifiedName(a, b, ..) => a.capacity() + b.capacity(),
            Expr::Ref(r) => r.prefix.capacity() + r.sheet.as_ref().map_or(0, String::capacity),
            Expr::Unary(_, child) => bytes(child) + 16,
            Expr::Binary(_, a, b) => bytes(a) + bytes(b) + 32,
            Expr::Compare(op, a, b) => op.capacity() + bytes(a) + bytes(b) + 32,
            Expr::Call(name, args) => {
                name.capacity()
                    + args.capacity() * std::mem::size_of::<Expr>()
                    + args
                        .iter()
                        .map(|e| bytes(e) - std::mem::size_of::<Expr>())
                        .sum::<usize>()
            }
            Expr::Union(areas) => {
                areas.capacity() * std::mem::size_of::<Expr>()
                    + areas
                        .iter()
                        .map(|e| bytes(e) - std::mem::size_of::<Expr>())
                        .sum::<usize>()
            }
            _ => 0,
        }
}
pub(super) fn parse(text: &str, limit: usize) -> Result<Expr> {
    let hit = CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        let cache = cache.as_mut()?;
        let entries = if limit == 8192 {
            &cache.references
        } else {
            &cache.formulas
        };
        let expr = entries.get(text)?.clone();
        cache.hits += 1;
        Some(expr)
    });
    if let Some(expr) = hit {
        return Ok(expr);
    }
    // Parse failures are never cached; each entry also has its parser limit.
    let expr = Parser::bounded(text, limit)?;
    CACHE.with(|cache| {
        if let Some(cache) = cache.borrow_mut().as_mut() {
            let cost = bytes(&expr) + text.len() + std::mem::size_of::<String>() + 64;
            if cache.entries < 128 && cache.bytes + cost <= 1024 * 1024 {
                let entries = if limit == 8192 {
                    &mut cache.references
                } else {
                    &mut cache.formulas
                };
                entries.insert(text.into(), expr.clone());
                cache.entries += 1;
                cache.bytes += cost;
            }
        }
    });
    Ok(expr)
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn limit_identity_and_scoped_cleanup_are_preserved() {
        let long = (0..300).map(|_| "1").collect::<Vec<_>>().join("+");
        {
            let _scope = begin();
            assert!(parse(&long, 8192).is_ok());
            assert!(parse(&long, 256).is_err());
            parse("A1+1", 256).unwrap();
            parse("A1+1", 256).unwrap();
            CACHE.with(|c| assert_eq!(c.borrow().as_ref().unwrap().hits, 1));
        }
        CACHE.with(|c| assert!(c.borrow().is_none()));
    }
    #[test]
    fn distinct_formulas_cannot_grow_the_cache_without_bound() {
        let _scope = begin();
        for i in 0..300 {
            parse(&format!("{i}+1"), 256).unwrap();
        }
        CACHE.with(|c| {
            let c = c.borrow();
            let c = c.as_ref().unwrap();
            assert_eq!(c.entries, 128);
            assert!(c.bytes <= 1024 * 1024)
        });
    }
}
