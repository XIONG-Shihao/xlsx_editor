---
status: working-notes
audience: the engineer (human or agent) doing the work; not the product owner
---

# How each claim is checked

Run from the product folder. `X` = `node --import ./scripts/ts-resolve-register.mjs`.

## Claims and their checks

| Claim (as worded in owner docs) | Check | Pass means | Does not show |
|---|---|---|---|
| "N% of everyday actions work on ordinary workbooks" | `X scripts/audit/refusals.mjs --json out.json` | Each of 30 actions applied to the first sheet of each of 14 generated workbooks (`ordinary-sheets.py`) without throwing | That the result is correct; that the UI offers the action; real files |
| "... on real workbook F" | `X scripts/audit/refusals.mjs F` | Same, first sheet of F | Other sheets of F |
| "edited files are right" | `X --test tests/ordinary-sheets.mjs` | openpyxl reads the saved file and sees the expected filter, merges, rules, validation, links, notes, tables, names, chart refs, formulas; `problems` empty | Styles, column widths, values beyond columns A:B, sheets other than the described one |
| "Excel accepts edited files" | `X scripts/audit/excel-check.mjs` | For each case Excel 16.x opens the file under its own path (a repaired file opens as a different name and fails), `calculate`, `save`; openpyxl description equal before/after for filter, merged, frozen, rules, validation, links, notes, tables, formulas; chart ref count equal; first `=SUM(` total equals the expected number | Visual rendering; Windows Excel; objects not in the cases |
| "calculator never wrong" | `X --test tests/calculator-routes.mjs` | For every formula with a stored result in `fixtures/excel-*.xlsx`: general route (Univer + patches) and product route (Rust fast path + general) each either decline or equal the stored value (1e-9 relative; `#SPILL!` matches stored `#VALUE!`), and agree with each other | Functions absent from the fixtures; real-file formulas; performance |
| "nothing unknown is dropped" | `tests/preserved-*.mjs`, `tests/preservation-content.mjs`, `tests/package-roundtrip.mjs` | Named opaque parts byte-equal after edits | Arbitrary real packages (step 3 of the plan) |
| "all tests pass" | `node scripts/test/run.mjs` | Every file exits 0, retried once alone on failure | GUI behaviour |

## Standing set

- `corpus/manifest.json` (committed) names each workbook: URL, source, licence, sha256, producer.
  `corpus/files/` (ignored) holds them; `fetch.mjs` re-downloads and verifies checksums; made ones
  are rebuilt by `build.py` + `excel-save.mjs` (their checksums change on each Excel save: timestamps).
- `measure.mjs` per workbook: `opening()`; 30 `attempt()`s (fresh open each; 150 s limit per
  attempt, 240 s budget per workbook, then "not tried"); `roundTrip()` = value 42 in an empty cell
  beyond the used and formatted columns, export, `compare-workbooks.py --ignore that cell`; with
  `--excel`: Excel open/calculate/save of the saved file, `compare --layout-may-change`; if that
  differs, Excel also re-saves the original and the two Excel copies are compared.
- `compare-workbooks.py` compares, per sheet: every cell's content (formula text normalised for
  spaces and quotes; numbers to 10 places) and format (number format, font, fill, alignment,
  borders, protection), links, notes; merges, freeze, filter, CF, validation, tables, print setup,
  header/footer, protection, tab colour, visibility, hidden/grouped rows and columns, widths and
  heights; workbook: sheet order, names; package: charts and their refs and kinds, picture hashes,
  pivots and sources, external links, embedded files, macros, threaded comments, custom XML,
  drawing anchors. NOT compared: sortState, filter button visibility, styles not used by a cell,
  theme, calcChain, stored formula results.
- Workbooks with `xl/externalLinks/` are not sent to Excel (it asks a person for file access).
  `excel.mjs` kills an Excel it started itself after a timeout, never one that was already open.
- `facts()` in `refusals.mjs` picks the visible sheet (of the first six) with most content, the
  header row (first row >=60% filled followed by three such rows) and the data block.

## Strict structure checks (`describe-workbook.py` -> `problems`)

1. Top-level worksheet records in CT_Worksheet order (list `ORDER` in the script).
2. No repeated singleton record.
3. Every internal relationship target of a sheet exists in the package.
4. Set of `<comment ref>` cells equals set of VML `ClientData ObjectType="Note"` (Row+1, Column+1).

Add a check here whenever Excel rejects something the reader accepted. Candidates not yet written:
`count` attributes equal child counts; `dimension` covers used cells; defined-name `localSheetId`
in range; `_xlnm._FilterDatabase` matches `autoFilter ref`; content types cover every part;
shared-string indexes in range; `dxfId`/`s` indexes in range.

## Excel check mechanics

- Excel for Mac is sandboxed: `open` of a file outside its container raises a Grant Access dialog
  that blocks AppleScript until a person clicks. The run folder is therefore
  `~/Library/Containers/com.microsoft.Excel/Data/Documents/workbook-excel-check-<time>/`. Removed on
  success unless `--keep`.
- First use needs macOS Automation permission for the terminal to control Excel (one prompt).
- Per-case timeouts: AppleScript 40 s, process 70 s. A repair prompt shows as a timeout.
- Failure "did not open the file under its own name" = Excel repaired it. Diff the package against
  the source part by part (`unzip` both, `cmp`), then look at record order and cross-part
  consistency first. "Parameter error" on open = Excel rejected the file outright.
- Do not kill Excel if the user has it open. `get name of every workbook` first.
- Excel legitimately changes: merges identical conditional-format rules into one `sqref`, renumbers
  `dxf`, rewrites VML and shared strings. Compare described content, not bytes.

## Fixtures

- `scripts/audit/ordinary-sheets.py` -> 01-plain, 02-filter, 03-merged-title, 04-table,
  05-highlight-rule, 06-colour-scale, 07-hyperlink, 08-dropdown, 09-frozen-formatted, 10-chart,
  11-names-two-sheets, 12-note, 13-typical-report, 15-other-language. (`14-excel-defaults` is made
  inside the test from 01.) openpyxl output: inline strings, `ns2:` VML prefix, absolute
  relationship targets. That variety is useful; keep it.
- Real, optional and not in the repository (neither is redistributable as a fixture):
  a supplied financial workbook named by `XLSX_EDITOR_FINANCIAL_WORKBOOK` (15 MB, 21 sheets, WPS-produced, 2,133 formulas, CJK names,
  external links, overlapping shared groups, OLE objects, `#REF!` names) and
  `samples/xlsx_sample.xlsx` (25 MB, data-only). Tests that need them
  (`tests/workbook-real-protected.mjs`) skip without them.
- Excel-calculated references: `fixtures/excel-*.xlsx` with `.json` answer keys.

## Python

`scripts/audit/python.mjs`: `XLSX_EDITOR_TEST_PYTHON` -> the `.build/python-audit` uv environment
(created on demand with `uv`, or `XLSX_EDITOR_UV`) -> without uv, `python3` if it imports openpyxl, or a
venv made from it. `scripts/test/run.mjs` calls `pythonOnPath()` so child `python3` resolves to it. A test run directly with `node --test` on a
machine without openpyxl fails with `ModuleNotFoundError` unless it imports `pythonOnPath` itself
(`ordinary-sheets.mjs` does) or PATH is set.

## Before saying "done"

1. `cargo test --release` in `native/`.
2. `scripts/test/run.mjs`.
3. `refusals.mjs` (ordinary) and on the real files; give the counts in the commit message.
4. `excel-check.mjs` if anything written to the package changed.
5. `npm run format`, `npm run test:format`.
6. Rebuild apps, or say they were not rebuilt.
