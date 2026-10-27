---
status: working-notes
audience: the engineer (human or agent) doing the work; not the product owner
---

# Implementation plan for the goal

Goal and targets: `../README.md`. Owner-facing steps: `../ROADMAP.md`. This file is how to do them.
Keep it current; delete what is finished (history is in Git).

## Operating rules

1. A refusal is found by measurement, not by reading code. Run `scripts/audit/refusals.mjs` on real
   files, fix the most frequent message, re-measure. Do not hunt guards speculatively.
2. Every guard is rewritten as "what does this action touch?" (`BIDIRECTIONAL_COMPATIBILITY.md`,
   rules 1-3). A guard of the form `Object.keys(parts).some(name => /x/.test(name))` or
   `sheets.some(s => s.flag)` is suspect by construction.
3. Each relaxation needs: an ordinary-sheets test using the independent reader, and a case in
   `excel-check.mjs` if it changes what is written. Run the Excel check before saying it is safe.
4. Most edits are implemented twice: JS plans (`core/indexed-objects.mjs`,
   `core/workbook-package.mjs`) and Rust recomputes and compares
   (`native/src/package_edit/rewrite_admission.rs`, message "Insertion fragment differs from native
   coordinate/reference transformation"). Change both or the edit is refused at the comparison.
5. After any `core/` or `native/` change: rebuild the standalone app. State whether a change has reached the source only or an app.
6. Report measurements with their denominator and what was not run. Never describe a scripted check
   as a GUI check.

## Step 1: real-workbook corpus (done)

`corpus/`, `scripts/corpus/`. See `VERIFICATION.md` for mechanics.
To grow it: `fetch.mjs --discover` with more catalogues; add Windows-Excel-saved and
LibreOffice/WPS/Google-exported files (only one LibreOffice file today); raise the 6 MB size cap
once per-action time is down.

## Step 2: refusals

Scoreboard order as of the third run. Re-measure before starting one; some may be done. New from the set:

| Refusal | Where | Approach |
|---|---|---|
| Structure/rename on a pivot's source sheet | `linkedToSheet`, rename check in `workbook-package.mjs` | Rewrite `worksheetSource ref`/`sheet` in the cache definition (one attribute each) |
| Manual page breaks | `rowInsertionReason` `unmoved` | Shift `<brk id>`; needs a new fragment tag on both sides (`indexed-objects.mjs`, `rewrite_admission.rs`, `insertion.rs`) |
| Duplicate with table | `duplicateSheet` in `workbook-package.mjs`, `indexed-copy.mjs` | New table id/name/displayName; rewrite structured refs in the copy's cells (native) |
| Fill touching a table | bulk fill guard (native) | Judge against table ref; filling inside a table body is ordinary |
| Sort narrower than used columns | `transform.rs` row-XML move | Cell-level permutation within column bounds |
| Protected sheet -> whole workbook read only | open path (`readOnly`) | Honour `locked` per cell on that sheet; other sheets editable |
| `Calculation graph metadata differs from accepted source` | open/save check, 2 workbooks | Investigate (canada-quebec-land-registry, one other) |
| `Unqualified rule raw boundary` | rules reader, 1 workbook (humanitarian covid) | Investigate |

Older list (some now done):

| Refusal | Message | Where | Approach |
|---|---|---|---|
| Sort with link/note on sorted rows | `Linked workbook transform...` / `Referenced worksheet transform...` | `native/src/transform.rs` `sort_guard`, `blocks_sort` | Move hyperlink `ref` and comment `ref` + VML Row with the row permutation (`sort_destination`) |
| Sort narrower than used columns | `Sort must include every used column in each selected row` | `transform.rs` | Rows are moved as XML rows. Needs cell-level move within column bounds; larger change |
| Sort/merge/fill with a table | `Sorting a sheet that has a table...`, `table-or-protection`, `Fill source intersects a table` | `transform.rs`, `layout.rs`, bulk fill guard | Scope to intersection with the table ref first (cheap); table-aware sort later |
| Duplicate sheet with table / extLst | `Duplicating this worksheet would lose...` | `core/workbook-package.mjs` duplicateSheet, `core/indexed-copy.mjs` | Clone table part with new id/name/displayName; extLst needs unique ids |
| Filter off with imported dropdown | `Imported validation is protected for content editing` | `indexedValidationMutationGuard` | removeFilter touches no validation; exempt non-content ops |
| Notes on sheet with imported notes | `Imported/external note graphs stay protected` | `native/src/package_edit/note_authoring.rs` | Adopt foreign comments+VML: append to existing parts instead of requiring owned part names |
| Structure beside editor-applied filter | `special formulas or extended` | hidden-row provenance record | Move the provenance rows with the edit |
| Preserve mode beyond values/format/layout/merge | `Formula and structural changes remain protected...` | `core/preservation-editing.mjs`, `indexed-session.mjs` `layoutOnly` | Filter: transform path reindexes without the preserve policy (`Preservation export does not match native accepted policy`); pass `preserveCalculation` through `#reindexRewrite`. Structure: needs ref rewrite of formulas the engine cannot parse; carry-through rule applies |
| Array formulas block structure | `Row insertion cannot preserve special formulas...` (`special-kinds` notShared > 0) | `#expandShared` | Array `ref` moves as a rectangle; implement for non-intersecting edits first |
| extLst on the edited sheet | same | `#expandShared`, `rowInsertionReason` | Parse x14 CF / DV / sparkline `xm:sqref`, `xm:f` and shift them |

Preserve-mode triggers still in force (`indexed-session.mjs` ~l.567, `sharedSourceQualificationFailure`):
`protectedFormulaCaches > 0` (import.rs ~l.707), shared-formula attribute other than t/si/ref,
inconsistent follower, orphan follower, >1024 groups on a sheet, expansion byte budget, opaque
`calcPr`. Overlapping master ranges no longer trigger it. Measure how often each occurs in the corpus
before working on any of them.

## Step 3: round-trip proof

- `--round-trip` mode: open, one value edit in a free cell, export; run `describe-workbook.py` on
  source and output for every sheet; diff everything except the edited cell. Extend
  `describe-workbook.py` to all sheets, cell values + number formats + styles digest, column widths,
  drawings/pictures count, pivot/table inventory, and a list of package parts with sizes (catch a
  dropped part).
- Excel check: add cases for editor-authored chart, PNG, note, text box/shape, pivot. Add a
  no-edit open/save case per corpus file (sample of ~20) to detect repair on plain round trip.
- Stale results: rule lives in `native/src/stale_cache.rs`, wired in `export_state.rs`; not applied
  in preserve mode (there `fullCalcOnLoad` is set and results are kept). Decide: drop results of
  formulas depending on edited cells in preserve mode too (lexical precedents work without
  calculating). Likely yes; Excel recalculates anyway.

## Step 4: wider action list

Add to `refusals.mjs` `actions`: paste formulas/formats, fill series, find+replace, number format,
add validation, add chart, set print area, define name, hide/unhide row, autofit-equivalent resize,
clear contents, delete sheet, move sheet. Each needs a per-sheet target picked from the described
workbook (`describe()` in the tool), not fixed addresses.

## Step 5: calculator

- `tests/calculator-routes.mjs` is the invariant (declines allowed, wrong not).
- Function census over the corpus: tokenise formulas, count function names, mark declined by
  `calculateWorkbook`. Add by frequency. Univer already implements most; declines are mostly this
  project's admission checks in `core/calculation.mjs`, not missing functions.
- `date-calculation.mjs` / `financial-calculation.mjs` are overrides registered into Univer
  (`excelCompatibilityOverrides`), not a separate route. Keep; they encode Excel's 1900/1904 and
  PMT/PV/FV behaviour verified against `fixtures/excel-date-*`.

## Step 7: one implementation per edit

Order by churn: (1) object ranges on insert/delete (`indexed-objects.mjs` vs
`rewrite_admission.rs::rewrite_nodes`) — make Rust the writer and delete the JS planner, or make the
Rust side a structural validator only (well-formed, in schema order, refs in range) instead of a
recomputation; (2) merges; (3) rename. Do one, measure test and code delta, then decide the rest.

## Not doing

Anything in `../FEATURES.md` "What is left out"; byte-identical output; sample-library growth;
non-Excel consumers as a pass/fail reference.
