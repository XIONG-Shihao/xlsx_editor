---
status: current
---

# Developing the spreadsheet editor

Working notes for whoever changes the editor. Product-level documents are listed on the
[front page](README.md).

## Folders

| Folder | Holds |
|---|---|
| `ui/` | The grid, ribbon and panels (React, TypeScript). `ui/SpreadsheetGrid.tsx` holds the screen's state and wiring; each concern (a ribbon group, the find bar, a command family, keyboard movement) is a file in `ui/grid/` |
| `core/` | The session layer: commands, admission, planning changes to small workbook parts (JavaScript) |
| `adapters/` | One worker per document, recovery, saving, CSV, spelling, local queries |
| `native/` | The engine (Rust). Builds the program `workbook-engine` |
| `standalone/` | Shell and screen code the desktop app runs |
| `desktop/` | The desktop app: Tauri shell (`src-tauri/`, with its own process supervision in `src/host/`), the Bun coordinator (`backend/`), the window's code (`frontend/`), packaging and its helpers (`scripts/`), the pinned Bun (`bun-runtime.json`) |
| `fixtures/` | Small source workbooks and workbooks calculated by Excel (`excel-*.xlsx`) |
| `tests/` | Tests. `tests/support/` holds shared helpers |
| `scripts/audit/` | Measuring tools; see [Measuring](#measuring) |
| `scripts/bench/` | Timing tools |
| `scripts/test/` | Larger checks run as scripts, and the line-length guard |

## What you need

- Node.js 24 and npm, for building and testing. (Shipped apps run on Bun, not Node.)
- Rust and Cargo, for the engine.
- uv. Many tests read saved files with `openpyxl` (Python 3.10+) as an independent check. A Python
  named in `XLSX_EDITOR_TEST_PYTHON` is used as it is. Otherwise the tests and measuring tools make a
  private uv environment in `.build/python-audit` the first time, with `uv` on PATH or the uv named by `XLSX_EDITOR_UV` (this needs the network once). Without
  uv they fall back to a `python3` 3.10+ on PATH (macOS's own is 3.9).
- Microsoft Excel for Mac, only for the Excel check.

## Build

From this folder:

```sh
node desktop/scripts/tooling/bun.mjs   # the pinned Bun, into .build/runtimes/bun
.build/runtimes/bun/darwin-arm64/payload/bun install --frozen-lockfile --ignore-scripts --linker hoisted
node scripts/build-native.mjs     # the engine; installs the program atomically
npm run build                     # engine + UI library
npm run typecheck
```

The engine program must not be overwritten while a test is using it. `build-native.mjs` installs by
rename for that reason; do not copy the file into place by hand.

## Test

Run files several at a time:

```sh
node scripts/test/run.mjs        # npm run test:all
#   --jobs N   --only text   --list
npm run gates                    # builds, type checks, the desktop shell's Rust tests, every file, the coordinator check
```

A file that fails is run once more on its own, so a failure reported at the end failed twice.

Engine unit tests: `cd native && cargo test --release`.

One file: `node --import ./scripts/ts-resolve-register.mjs --test tests/NAME.mjs`.

The tests that matter most for the goal:

| Test | Checks |
|---|---|
| `tests/ordinary-sheets.mjs` | Everyday edits on ordinary workbooks, read back by an independent reader, including the strict structure checks |
| `tests/calculator-routes.mjs` | Neither calculation route disagrees with results Excel stored |
| `tests/preserved-layout.mjs` | What may change on a workbook kept as stored |
| `tests/refusal-codes.mjs` | Refusals carry codes |

`tests/standalone-bridge.mjs` checks the host-neutral API, and
`tests/saved-package-oracle.mjs` retains independent saved-style/cache oracles and
native footprint accounting. The desktop coordinator check is
`node desktop/scripts/verify-coordinator.mjs`, from source or, given its path, on a packaged app.
`tests/product-folder.mjs` fails when a file reaches outside this folder.

### Which test files run

Every file in `tests/` is run by `scripts/test/run.mjs` unless `tests/not-in-suite.json` lists it with the
reason. A new test file needs no registration. The list has two groups: tests that refuse to start
while another engine process is running, and tests that fail and need repair. A test must run without
arguments and must not write report files; put what it checks in assertions.

## Refusal codes

When the editor will not do something it refuses with a sentence and a stable code. The list of codes,
what each means and the rule for choosing one is at the top of `core/refusal-codes.mjs`. Tests assert
the code, not the sentence, so a sentence can be improved without breaking them.

- A new refusal is written as `refusal('code', 'sentence')` in JavaScript or
  `crate::coded("code", "sentence")` in Rust. Or write it as before and run
  `node scripts/refusal-codes/apply.mjs`, which gives every refusal that has no code the one that
  `scripts/refusal-codes/mapping.mjs` chooses (a rule per cause, not per sentence). If no rule covers it,
  add a rule or a new code to the list; `node scripts/refusal-codes/check.mjs --list` shows what is missing.
- `tests/refusal-sites.mjs` fails when a refusal has no code. The allow-list
  `scripts/refusal-codes/allow-list.json` names files that still have some; it only ever shrinks
  (`node scripts/refusal-codes/check.mjs --shrink`).
- `node scripts/refusal-codes/convert-tests.mjs --verify` turns test assertions on refusal wording into
  assertions on the code, runs each changed file and puts back what then fails.
- Rust code that reads what an error says uses `crate::sentence(&error)`; code that adds context uses
  `crate::rephrase(code, error, |said| …)`, so the code is kept in front.
- JavaScript that tests `error.code` to tell a system error (ENOENT) from other errors uses
  `isRefusal(error)` from the same file.

## The standing set of real workbooks

`corpus/` holds about 100 workbooks the editor is measured against: public files from open-data
catalogues and statistics offices, plus workbooks made here to cover features the public files lack.
The files themselves are not committed (`corpus/files/`); `corpus/manifest.json` records where each
came from, its licence and its checksum.

```sh
node scripts/corpus/fetch.mjs                 # download the public workbooks named in the manifest
python3 scripts/corpus/build.py corpus/made corpus/files   # make the everyday and merged workbooks
node --import ./scripts/ts-resolve-register.mjs scripts/corpus/excel-save.mjs   # pass the made ones through Excel
python3 scripts/corpus/inventory.py           # which features the set contains
node --import ./scripts/ts-resolve-register.mjs scripts/corpus/measure.mjs --excel   # the scoreboard
```

`measure.mjs` keeps its results in `corpus/reports/scoreboard.json` and can be stopped and resumed;
`--fresh` starts again, `--only text` limits it to matching names. A full run with Excel takes about
half an hour. `node scripts/corpus/fetch.mjs --discover` looks for more public workbooks.

Do not rebuild the engine while a measurement is running. To keep working during one, copy the
editor to a scratch folder (without `node_modules`, `native/target`, `tests`, `fixtures`, `ui`),
link `node_modules`, `corpus/files` and `.build/python-audit` into the copy, and run the measurement
there; the copy has its own engine program.

### The definition of done

A change to the editor is done when a fresh measurement is no worse than the recorded one:

```sh
node --import ./scripts/ts-resolve-register.mjs scripts/corpus/measure.mjs --fresh --excel-actions
node scripts/corpus/gate.mjs            # fails, naming each workbook and action that got worse
node scripts/corpus/gate.mjs --update   # records an equal or better result as the new baseline
```

`corpus/baseline.json` (committed) remembers, for each workbook, which actions work, whether the
round trip was clean, whether Excel accepted the file, and which actions matched the detailed check.
The gate compares workbook by workbook, so an improvement in one place cannot hide a loss in
another. An action a large workbook did not reach within its time allowance is not counted as lost.

`--excel-actions` adds the detailed check. An action that should change one place (typing, a format,
a merge) is compared with the original: only that place may differ. An action that moves things
(insert or delete a row or column, sort, rename, copy a sheet) is given to Excel as well, and the two
results are compared. `--detail` runs only the first kind and needs no Excel. A full run with Excel
takes about two hours and uses Excel the whole time; it quits Excel afterwards only if it started it.

## Format

```sh
npm run format        # Prettier and rustfmt
npm run test:format   # the line-length guard
```

The guard counts lines over 160 characters per folder and fails if the count grows. Never raise the
numbers in `scripts/test/line-length-baseline.json`; split the line.

## Measuring

All run from this folder with `node --import ./scripts/ts-resolve-register.mjs`. Real files given
to these tools are only read; a copy is edited.

| Tool | Answers |
|---|---|
| `scripts/audit/refusals.mjs [--json out.json] [FILE or FOLDER...]` | Of 30 everyday actions, how many work on each workbook, and what the refusals say. With no argument, uses the generated ordinary workbooks |
| `scripts/audit/excel-check.mjs [--keep]` | Does real Excel open, recalculate and save each edited file without change. Opens Excel on screen. Works in a folder named `workbook-excel-check-*` inside Excel's own container, the one place Excel opens files without asking |
| `scripts/audit/calculator-agreement.mjs FILE...` | How often the general calculator agrees with results stored in the file |
| `scripts/audit/save-edited.mjs SOURCE OUTPUT 'JSON commands'` | Apply commands and save, for looking at one case |
| `scripts/audit/describe-workbook.py FILE [SHEET]` | What an independent reader sees, plus problems Excel would reject |
| `scripts/audit/compare-workbooks.py BEFORE AFTER [--ignore 'Sheet!A1']` | Whether anything was lost or changed between two workbooks: cells, formats, objects, parts. Independent reader |
| `scripts/audit/excel.mjs` | Helpers for driving real Excel from a script (used by the Excel check and the scoreboard) |
| `scripts/audit/feature-coverage.py`, `feature-table.py` | What each row of `FEATURES.md` is checked by; the second writes the column |
| `scripts/audit/ordinary-sheets.py FOLDER` | Generate the ordinary workbooks |
| `scripts/bench/save-engine.mjs`, `save-document.mjs` | Save times |

## After changing the engine

Rebuild the standalone app to carry the change:

```sh
npm run desktop:package   # quit the app first
```

## Things that have gone wrong

- **A fix reported as done that was not in any app.** Tests passed; the app was not rebuilt. Say
  whether a change is in the source only or also in an app.
- **Standalone API drift.** The frontend uses `standalone/bridge.mjs`; the native coordinator
  verifier checks every bridge operation. Rebuild and verify the standalone app after changing
  coordinator or bridge contracts.
- **The engine program replaced while in use.** macOS killed the running one. Use
  `scripts/build-native.mjs`.
- **A check that could not fail.** Tests that read a saved file back with the editor's own reader
  passed files Excel rejected. Use the independent reader and, before a release, Excel.
- **Whole-workbook refusals.** A guard written as "the workbook contains X anywhere" refuses work
  that has nothing to do with X. Ask what the action touches.
