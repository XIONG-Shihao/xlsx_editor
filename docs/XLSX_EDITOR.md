---
status: current
---

# The spreadsheet editor

## What it is

A spreadsheet editor for `.xlsx` workbooks and `.csv`/`.tsv` files that runs on the user's own
computer. It opens, edits, calculates, formats, prints and saves without an account or a network
connection.

The standalone spreadsheet app is its only user; it is in no other app. How the app uses it is in
[XLSX_EDITOR_INTEGRATION.md](XLSX_EDITOR_INTEGRATION.md).

What it aims to be is on the [documentation front page](README.md#the-goal). What a person can do
with it is in [FEATURES.md](FEATURES.md).

## Files it handles

| File | Open | Edit and save | Notes |
|---|---|---|---|
| `.xlsx` | Yes | Yes | The main format. Parts the editor does not understand are kept as they were. |
| `.csv`, `.tsv` | Yes | Yes | Text only, no formulas. Encoding, delimiter, quoting and line endings are kept. |
| `.xlsm` | Yes | No | Read only. Macros can be inspected, not run. Save an exact copy, or export values to a new `.xlsx`. |
| `.xls`, `.xlsb` | Yes | No | Read only. Save an exact copy, or export values to a new `.xlsx`. |
| Protected or signed `.xlsx` | Yes | No | Read only; exact copy. |

## How it is built

Three layers, each with one job.

```text
   ui/        the grid, the ribbon, panels            React, TypeScript
     |          asks for the cells on screen, sends commands
   core/      the session: commands, undo, what to refuse    JavaScript
   adapters/  one worker per open document, recovery, save
     |          JSON messages
   native/    the engine: storage, calculation, writing      Rust, program "workbook-engine"
```

**The engine (`native/`).** A Rust program started once per session. It reads the workbook file
straight into an index held in SQLite on disk, so a large workbook is not held in memory. It answers
"give me these cells", applies edits as versions, calculates, and writes the saved file. It writes by
changing the parts of the original file that an edit touched and copying everything else, which is
how unknown content survives. See [compatibility](BIDIRECTIONAL_COMPATIBILITY.md).

**The session (`core/`, `adapters/`).** JavaScript. Turns what the user does into commands, decides
what is allowed, keeps undo history, recovery and save. It reads the small parts of a workbook
(names, tables, rules, drawings) itself and plans changes to them; the engine checks and writes them.

**The screen (`ui/`).** A React grid and ribbon. It only ever holds the cells being shown plus the
workbook's outline, never the whole workbook.

**Calculation.** Formulas are calculated by the Univer formula engine (version 0.25.2, open source),
with a small set of patches so dates and loan functions match Excel exactly. About 20 very common
functions also have a faster implementation inside the Rust engine, used when a formula's shape
allows. The two must give the same answers as Excel; a test checks them against results Excel
stored. A formula neither can calculate is saved as typed and left for Excel to calculate.
Background: [CALCULATOR_EVALUATION.md](CALCULATOR_EVALUATION.md).

**Runtime.** The JavaScript layers run on Bun inside the shipped apps (no Node.js and no Electron in
what ships). Development and tests use Node.js 24.

## Saving, undo and recovery

- **Save** writes a new file beside the original and swaps it in, after checking the original has
  not changed on disk since it was opened.
- **Undo** keeps the last 100 changes per workbook.
- **Recovery.** Each accepted change is written to a journal. After a crash the workbook reopens
  with the unsaved changes.
- **Refusals.** When a change cannot be made safely the workbook is left exactly as it was and the
  user is told why. Some refusals carry a short code (`frozen-area`, `filtered-column`,
  `table-or-protection`, and others) so the program can react to the kind of refusal without
  reading the sentence; the rest are being converted.

## Limits in the code today

| What | Limit |
|---|---|
| Open workbooks in the standalone app | 4 tabs |
| Undo | 100 steps per workbook |
| One paste, fill, find-and-replace or selection | 10,000 cells |
| Frozen panes | 100 rows, 30 columns |
| Chart | 8 series, 1,000 source cells |
| Picture | PNG, 128 KB |
| Print preview and PDF | 100 pages |
| Working index on disk, per session | 4 GB |

These are limits on single operations and on disk use. They are not a limit on workbook size: a
real 25 MB workbook opens and accepts all 30 of the measured everyday actions.

## Speed, measured

On an Apple-silicon Mac, with a generated workbook of 10,000 rows by 8 columns:

| Operation | Time |
|---|---|
| Save after one edit | 2.1 s |
| Second save | 2.4 s |

Measured without a window. Speed in the built app, on large real files, has not been measured;
that is [roadmap](ROADMAP.md) step 6.

## Known weak points

- **Real-file coverage is thinly measured.** Two real workbooks, plus generated ones. See
  [roadmap](ROADMAP.md) step 1.
- **Two implementations of many rules**, one in JavaScript and one in Rust, which must agree. This
  makes changes slow. See roadmap step 7.
- **The main screen file is very large** (`ui/SpreadsheetGrid.tsx`), which makes the interface hard
  to change safely.
- **Many older tests check internal details** rather than what a saved file contains. New tests use
  an independent reader; the old ones have not been rewritten.
- **The built app is checked by scripts, not by eye.** No one has yet walked the everyday list in
  the app window on this version.
