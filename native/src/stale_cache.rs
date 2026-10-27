//! Decides which stored results of uncalculated formulas are still true.
//!
//! A formula the engine could not calculate keeps the number it was imported
//! with. That number stays correct only while nothing the formula refers to has
//! changed. When an input did change, or the inputs cannot be read from the
//! formula text, the stored number is dropped on export: a program that does
//! not recalculate then shows an empty cell instead of a wrong value, and Excel
//! recalculates on open either way.
//!
//! The reader here is lexical on purpose. It must work on formulas the
//! calculator rejects, and it answers "unknown" for anything it does not
//! recognise rather than guessing.

use std::collections::{HashMap, HashSet};

/// An inclusive block of cells on one worksheet part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Area {
    pub(crate) sheet: String,
    pub(crate) first: (u32, u32),
    pub(crate) last: (u32, u32),
}

const LAST_ROW: u32 = 1_048_575;
const LAST_COLUMN: u32 = 16_383;

/// Functions whose result depends on more than the cells written in the formula.
const UNBOUNDED: &[&str] = &[
    "INDIRECT",
    "OFFSET",
    "NOW",
    "TODAY",
    "RAND",
    "RANDBETWEEN",
    "RANDARRAY",
    "CELL",
    "INFO",
    "GETPIVOTDATA",
    "RTD",
    "WEBSERVICE",
    "FILTERXML",
];

fn column(text: &str) -> Option<u32> {
    if text.is_empty() || text.len() > 3 || !text.bytes().all(|b| b.is_ascii_alphabetic()) {
        return None;
    }
    let index = text.bytes().fold(0u32, |n, b| {
        n * 26 + u32::from(b.to_ascii_uppercase() - b'A') + 1
    }) - 1;
    (index <= LAST_COLUMN).then_some(index)
}

fn row(text: &str) -> Option<u32> {
    if text.is_empty()
        || text.len() > 7
        || text.starts_with('0')
        || !text.bytes().all(|b| b.is_ascii_digit())
    {
        return None;
    }
    let index = text.parse::<u32>().ok()? - 1;
    (index <= LAST_ROW).then_some(index)
}

enum Endpoint {
    Cell(u32, u32),
    Column(u32),
    Row(u32),
}

/// `A1`, `$A$1`, `A`, `$A`, `1` or `$1`; anything else is not a reference end.
fn endpoint(token: &str) -> Option<Endpoint> {
    let token = token.strip_prefix('$').unwrap_or(token);
    let letters = token.bytes().take_while(u8::is_ascii_alphabetic).count();
    let (letters, rest) = token.split_at(letters);
    let digits = rest.strip_prefix('$').unwrap_or(rest);
    match (letters.is_empty(), digits.is_empty()) {
        (false, false) => Some(Endpoint::Cell(row(digits)?, column(letters)?)),
        (false, true) if rest.is_empty() => Some(Endpoint::Column(column(letters)?)),
        (true, false) if rest == digits => Some(Endpoint::Row(row(digits)?)),
        _ => None,
    }
}

fn area(sheet: &str, first: &str, second: Option<&str>) -> Option<Area> {
    let (first, last) = match (endpoint(first)?, second.map(endpoint)) {
        (Endpoint::Cell(r, c), None) => ((r, c), (r, c)),
        (Endpoint::Cell(r0, c0), Some(Some(Endpoint::Cell(r1, c1)))) => {
            ((r0.min(r1), c0.min(c1)), (r0.max(r1), c0.max(c1)))
        }
        (Endpoint::Column(c0), Some(Some(Endpoint::Column(c1)))) => {
            ((0, c0.min(c1)), (LAST_ROW, c0.max(c1)))
        }
        (Endpoint::Row(r0), Some(Some(Endpoint::Row(r1)))) => {
            ((r0.min(r1), 0), (r0.max(r1), LAST_COLUMN))
        }
        _ => return None,
    };
    Some(Area {
        sheet: sheet.into(),
        first,
        last,
    })
}

/// The cells a formula reads, or `None` when that cannot be told from its text.
///
/// `own` is the worksheet part holding the formula; `sheets` maps a lower-case
/// display name to its part.
pub(crate) fn precedents(
    formula: &str,
    own: &str,
    sheets: &HashMap<String, String>,
) -> Option<Vec<Area>> {
    let characters: Vec<char> = formula.chars().collect();
    let word = |c: char| c.is_alphanumeric() || matches!(c, '_' | '.' | '$' | '\\' | '?');
    let mut areas = Vec::new();
    let mut at = 0;
    while at < characters.len() {
        let c = characters[at];
        match c {
            '"' => {
                // A text literal; a doubled quote is an escaped quote.
                at += 1;
                loop {
                    match characters.get(at) {
                        None => return None,
                        Some('"') if characters.get(at + 1) == Some(&'"') => at += 2,
                        Some('"') => break,
                        Some(_) => at += 1,
                    }
                }
                at += 1;
            }
            // Structured, external and spilled references, and array constants.
            '[' | ']' | '#' | '{' | '}' | '@' => return None,
            '\'' => {
                let mut name = String::new();
                at += 1;
                loop {
                    match characters.get(at) {
                        None => return None,
                        Some('\'') if characters.get(at + 1) == Some(&'\'') => {
                            name.push('\'');
                            at += 2;
                        }
                        Some('\'') => break,
                        Some(other) => {
                            name.push(*other);
                            at += 1;
                        }
                    }
                }
                if characters.get(at + 1) != Some(&'!') {
                    return None;
                }
                at += 2;
                let sheet = sheets.get(&name.to_lowercase())?;
                at = reference(&characters, at, sheet, &mut areas, word)?;
            }
            c if word(c) => {
                let start = at;
                while at < characters.len() && word(characters[at]) {
                    at += 1;
                }
                let token: String = characters[start..at].iter().collect();
                match characters.get(at) {
                    Some('(') => {
                        if UNBOUNDED.contains(&token.to_ascii_uppercase().as_str()) {
                            return None;
                        }
                    }
                    Some('!') => {
                        let sheet = sheets.get(&token.to_lowercase())?;
                        at = reference(&characters, at + 1, sheet, &mut areas, word)?;
                    }
                    _ => {
                        let plain_number = token.parse::<f64>().is_ok()
                            && !(characters.get(at) == Some(&':')
                                || start > 0 && characters[start - 1] == ':');
                        let exponent = token.ends_with(['e', 'E'])
                            && token[..token.len() - 1].parse::<f64>().is_ok()
                            && matches!(characters.get(at), Some('+' | '-'));
                        if plain_number
                            || exponent
                            || ["TRUE", "FALSE"].contains(&token.to_ascii_uppercase().as_str())
                        {
                            continue;
                        }
                        at = reference(&characters, start, own, &mut areas, word)?;
                    }
                }
            }
            _ => at += 1,
        }
    }
    Some(areas)
}

/// Reads `A1`, `A1:B2`, `A:C` or `1:5` starting at `at`; returns the next index.
fn reference(
    characters: &[char],
    at: usize,
    sheet: &str,
    areas: &mut Vec<Area>,
    word: impl Fn(char) -> bool,
) -> Option<usize> {
    let token = |from: usize| {
        let mut end = from;
        while end < characters.len() && word(characters[end]) {
            end += 1;
        }
        (characters[from..end].iter().collect::<String>(), end)
    };
    let (first, mut end) = token(at);
    let mut second = None;
    if characters.get(end) == Some(&':') {
        let (text, next) = token(end + 1);
        // `Sheet1:Sheet3!A1` spans worksheets; its extent is not read here.
        if characters.get(next) == Some(&'!') {
            return None;
        }
        second = Some(text);
        end = next;
    }
    // A name followed by `(` or `!` here would be a function or another sheet.
    if matches!(characters.get(end), Some('(' | '!')) {
        return None;
    }
    areas.push(area(sheet, &first, second.as_deref())?);
    Some(end)
}

/// Which uncalculated formulas can no longer vouch for their stored result.
///
/// `changed` holds every cell whose saved content differs from the source.
/// `formulas` are the uncalculated formulas as (sheet part, row, column, text).
pub(crate) fn stale(
    changed: &HashSet<(String, u32, u32)>,
    formulas: &[(String, u32, u32, String)],
    sheets: &HashMap<String, String>,
) -> Vec<(String, u32, u32)> {
    if changed.is_empty() || formulas.is_empty() {
        return Vec::new();
    }
    let read: Vec<Option<Vec<Area>>> = formulas
        .iter()
        .map(|(sheet, _, _, text)| precedents(text, sheet, sheets))
        .collect();
    let mut is_stale: Vec<bool> = read.iter().map(Option::is_none).collect();
    type Blocks = HashMap<String, Vec<((u32, u32), (u32, u32))>>;
    let touches = |areas: &[Area], blocks: &Blocks| {
        areas.iter().any(|area| {
            blocks.get(&area.sheet).is_some_and(|blocks| {
                blocks.iter().any(|(first, last)| {
                    first.0 <= area.last.0
                        && last.0 >= area.first.0
                        && first.1 <= area.last.1
                        && last.1 >= area.first.1
                })
            })
        })
    };
    // An uncalculated formula may be a dynamic array whose results spill down and
    // to the right of it. Once its result is dropped, every cell in that
    // direction may have changed with it.
    let spill = |r: u32, c: u32| ((r, c), (LAST_ROW, LAST_COLUMN));
    // Start from the real changes and the formulas that could not be read. Each
    // round then checks the remaining formulas against what was dropped last.
    let mut pending: Blocks = HashMap::new();
    for (sheet, r, c) in changed {
        pending
            .entry(sheet.clone())
            .or_default()
            .push(((*r, *c), (*r, *c)));
    }
    for ((sheet, r, c, _), _) in formulas.iter().zip(&is_stale).filter(|(_, stale)| **stale) {
        pending
            .entry(sheet.clone())
            .or_default()
            .push(spill(*r, *c));
    }
    loop {
        let mut next: Blocks = HashMap::new();
        for (index, areas) in read.iter().enumerate() {
            if !is_stale[index]
                && areas
                    .as_deref()
                    .is_some_and(|areas| touches(areas, &pending))
            {
                is_stale[index] = true;
                let (sheet, r, c, _) = &formulas[index];
                next.entry(sheet.clone()).or_default().push(spill(*r, *c));
            }
        }
        if next.is_empty() {
            break;
        }
        pending = next;
    }
    formulas
        .iter()
        .zip(is_stale)
        .filter(|(_, stale)| *stale)
        .map(|((sheet, r, c, _), _)| (sheet.clone(), *r, *c))
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sheets() -> HashMap<String, String> {
        HashMap::from([
            ("sheet1".into(), "s1".into()),
            ("my data".into(), "s2".into()),
        ])
    }
    fn read(formula: &str) -> Option<Vec<((u32, u32), (u32, u32), String)>> {
        precedents(formula, "s1", &sheets()).map(|areas| {
            areas
                .into_iter()
                .map(|area| (area.first, area.last, area.sheet))
                .collect()
        })
    }

    #[test]
    fn reads_cells_ranges_and_other_sheets() {
        assert_eq!(read("A1+1").unwrap(), [((0, 0), (0, 0), "s1".to_string())]);
        assert_eq!(
            read("SUM($B$2:C10)*2.5e3").unwrap(),
            [((1, 1), (9, 2), "s1".to_string())]
        );
        assert_eq!(
            read("'My Data'!A:B").unwrap(),
            [((0, 0), (LAST_ROW, 1), "s2".to_string())]
        );
        assert_eq!(
            read("Sheet1!3:4").unwrap(),
            [((2, 0), (3, LAST_COLUMN), "s1".to_string())]
        );
        assert_eq!(read("IF(TRUE,\"A1 is \"\"text\"\"\",1E-3)").unwrap(), []);
        assert_eq!(read("LOG10(A1)+XFD1048576").unwrap().len(), 2);
    }

    #[test]
    fn answers_unknown_instead_of_guessing() {
        for formula in [
            "TaxRate*A1",
            "INDIRECT(\"A1\")",
            "OFFSET(A1,1,1)",
            "Table1[Total]",
            "[1]Sheet1!A1",
            "A1#",
            "Sheet1:Sheet3!A1",
            "Missing!A1",
            "{1,2}*A1",
            "TODAY()",
            "\"unterminated",
            "XFE1",
            "A1048577",
        ] {
            assert_eq!(read(formula), None, "{formula}");
        }
    }

    #[test]
    fn staleness_follows_changes_through_other_uncalculated_formulas() {
        let formula = |r, c, text: &str| ("s1".to_string(), r, c, text.to_string());
        let formulas = [
            formula(0, 5, "A1+1"),
            formula(1, 5, "F1*2"),
            formula(2, 5, "B9+1"),
            formula(3, 5, "SomeName+1"),
            formula(4, 5, "'My Data'!A1"),
            formula(5, 6, "G9*2"),
        ];
        let changed = HashSet::from([("s1".to_string(), 0u32, 0u32)]);
        let stale = stale(&changed, &formulas, &sheets());
        assert_eq!(
            stale.iter().map(|(_, r, _)| *r).collect::<Vec<_>>(),
            [0, 1, 3, 5],
            "a changed input, a chain through a dropped result, an unreadable formula, and a read of a possible spill"
        );
        assert!(
            super::stale(&HashSet::new(), &formulas, &sheets()).is_empty(),
            "nothing changed, nothing is stale"
        );
    }
}
