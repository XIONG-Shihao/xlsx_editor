---
status: evaluation
---

# One calculator: evaluation

## Outcome

Kept the current calculator. What changed after the evaluation below:

- **It is two routes, not three.** The "hand-written JavaScript" listed below turned out to be small
  patches fitted into the Univer engine so that dates and loan functions match Excel exactly (the
  1900 and 1904 calendars, PMT, PV, FV). They are part of the general route, not a separate
  calculator, and they are checked against workbooks Excel calculated. They stay.
- **A test now holds the routes to Excel.** `tests/calculator-routes.mjs` runs every workbook that
  Excel calculated through the general route and through the product route (Rust fast path first,
  general route for the rest). Either may decline a formula; neither may give a different answer
  from the one Excel stored, and they may not disagree with each other.

| | Excel results compared | Answered | Wrong |
|---|---:|---:|---:|
| General route (Univer with patches) | 620 | 523 | 0 |
| Product route (Rust fast path, then general) | 620 | 574 | 0 |

The 46 results the product route still declines are formulas the editor leaves for Excel to calculate.

## The problem

A formula in this editor can be calculated in these places.

| Calculator | Where | Covers |
|---|---|---|
| Rust engine | `native/src/formula.rs`, `calculation.rs` | About 21 functions (SUM, AVERAGE, IF, COUNTIF, SUMIF, VLOOKUP, XLOOKUP, DATE, EOMONTH, FV and a few text functions), with strict rules about which shapes qualify |
| Univer formula engine 0.25.2 | `core/calculation.mjs`, five `@univerjs` packages, 68 MB installed | Everything else it can, on a copy of the cells it needs |
| Patches inside Univer | `core/date-calculation.mjs`, `financial-calculation.mjs`, `array-calculation.mjs` | Six date functions, PMT/PV/FV, and array bookkeeping, made to match Excel |

A formula is routed by its shape. Every route has to agree with the others and with Excel, and each
hand-off is a place where a formula falls out as "cannot calculate". Results cross between
JavaScript and Rust for every dependent cell.

## What was tried

IronCalc 0.8.3 (Rust, MIT or Apache-2.0, reads and writes `.xlsx`, describes itself as work in
progress) was built into a small program that loads a workbook, calculates it and prints chosen
cells. Its results were compared with the values Excel had stored in this project's own Excel-made
reference workbooks (`fixtures/excel-*.xlsx`).

| Reference workbook | Formulas | Agree with Excel | Notes |
|---|---:|---:|---|
| b5-sum-order, b6-runtime, b7-add-sub, coercion23, dec2hex, native-text22, dynamic-reference | 354 | 354 | Arithmetic, coercion, DEC2HEX, text |
| b5-oracle | 13 | 11 | A table reference (`Prices[Value]`) gave `#NAME?`; `LEN("🙂")` gave 1 where Excel gives 2 |
| date-1900-reference | 257 | 192 | Dates at the start of the 1900 calendar gave `#NUM!` |
| date-1904-reference | 257 | 72 | The 1904 date system is not applied |
| union-reference | 11 | 0 | `SUM((B1,B2))` gave `#ERROR!` |
| data-table23, scenarios-coercion24 | — | — | The workbook would not load: data-table formulas are not implemented |
| **Total** | **892** | **629 (70.5%)** | 3 of 14 workbooks could not be evaluated at all |

These reference workbooks were built to probe hard cases, so 70% is not what an everyday workbook
would see. The failures are informative all the same: table references, the 1904 calendar and
union references are things real workbooks contain, and one unsupported feature made a whole
workbook unloadable.

The same reference workbooks through the editor's current general route (Univer behind this
project's admission checks, `scripts/audit/calculator-agreement.mjs`), without the Rust route:

| | Formulas compared | Agree with Excel | Wrong | Not calculated |
|---|---:|---:|---:|---:|
| Current general route | 620 | 523 (84.4%) | 0 | 97 |
| IronCalc 0.8.3 | 892 | 629 (70.5%) | 263 (wrong value or error) | 3 workbooks not loaded |

The important column is "Wrong". The current route gave no wrong answer: what it could not do, it
declined (indirect references, shared and array formulas, union references, a name conflict), and
the Rust route covers several of those in the product. IronCalc returned an error or a different
value for 263 formulas. A declined formula keeps its stored number and is recalculated by Excel; a
wrong number is shown as if it were right.

On one real 15 MB financial workbook with 2,133 formulas, IronCalc did not finish loading and
calculating within 7 minutes (and had used 15 minutes of processor time in an earlier attempt
before it was stopped).

Other candidates were ruled out on paper:

| Candidate | Why not |
|---|---|
| HyperFormula | GPLv3 or a paid licence; not compatible with shipping this product |
| Formula.js | A library of functions only; no dependency tracking or recalculation |
| Growing the Rust engine to cover everything | 21 functions today against several hundred in Excel; years of work to reach what exists elsewhere |

## Recommendation

Do not switch engines now. Make Univer the one calculator and shrink the Rust calculator to a
verified fast path.

1. **One source of truth.** Univer decides every formula result. The Rust functions stay only where
   a test proves they give the same answer as Univer for the same input, and they are used for
   speed, not for correctness.
2. ~~Delete the hand-written JavaScript calculators.~~ Not done, on purpose: they are patches
   inside Univer that make dates and loan functions match Excel. See the outcome above.
3. **Add a disagreement test.** Done: `tests/calculator-routes.mjs`.
4. **Look at IronCalc again when it reaches 1.0** or handles tables, the 1904 calendar and union
   references. A Rust engine inside the Rust process would remove the JavaScript hop and the 68 MB
   of packages, which is worth having, but not at 70% agreement and with workbooks it cannot load.

Why Univer rather than the Rust engine as the one calculator: it already covers the most functions,
it is already shipped, it gave no wrong answer in the trial, and its licence (Apache-2.0) is fine. What it costs is speed on large sheets
and the copy of cells handed to JavaScript; the fast path in step 1 keeps the common cases quick.

## What this does not settle

- The 84.4% is for the general route alone on hard reference cases; the product's combined routes were
  not measured as one. Step 3 does that.
- IronCalc was tried at one version on one machine with one real workbook.
- No timing comparison between the routes was made.
