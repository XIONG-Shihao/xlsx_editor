---
status: current
---

# Features: everyday coverage checklist

What a person can see and do in the editor, tab by tab. This is the progress view for the
[goal](README.md#the-goal): cover the spreadsheet work people do every day.

Features that need large libraries of assets (shapes, icons, themes and similar) are listed
separately in [FEATURES_SAMPLES.md](FEATURES_SAMPLES.md). Those keep their controls and a few working
samples; the full libraries come later.

## How to read the status

| Status | Meaning |
|---|---|
| Works | The feature does what the row says. |
| Works, with limits | It works for the common case; the limit that matters is in the row. |
| Samples | Controls and a few examples exist. See [FEATURES_SAMPLES.md](FEATURES_SAMPLES.md). |
| Not verified | The control exists, but nothing that runs today checks it. |
| Not planned | Outside the goal. See [What is left out](#what-is-left-out). |

The **Checked by** column says what stands behind each status today:

- *N test files* are automated tests of the editor's engine that run on every change and pass.
- *x of y on real workbooks* is how the row's everyday actions did on the standing set of 101 real
  workbooks (see [What is measured](#what-is-measured)). This is the number to trust most.
- *not checked without the app window* means only a person using the app can confirm it.

Nothing in this column is a check in the app's window. No one has yet walked this list by hand in
the built app; that is [roadmap](ROADMAP.md) step 2.

## Summary

| Tab | Everyday features | Works | Works, with limits | Not verified |
|---|---:|---:|---:|---:|
| Home | 16 | 10 | 6 | 0 |
| Insert (everyday part) | 5 | 1 | 4 | 0 |
| Page Layout (everyday part) | 5 | 2 | 2 | 1 |
| Formulas | 8 | 3 | 5 | 0 |
| Data | 8 | 5 | 3 | 0 |
| Review | 4 | 2 | 1 | 1 |
| View | 7 | 5 | 2 | 0 |
| Pivot tables | 5 | 0 | 5 | 0 |
| Files and safety | 6 | 3 | 3 | 0 |
| **Total** | **64** | **31** | **31** | **2** |

Of the 64 rows, 62 are exercised by automated tests that run on every change and 17 are also
measured on real workbooks. Two (printing to paper, spelling) are checked by nothing today.

Nothing everyday is entirely missing a control. The work left is in the "with limits" column: making
each feature work on real workbooks, not only on simple ones. That is what the
[goal](README.md#the-goal) measures.

## Home

| Feature | What you can do | Limit that matters | Status | Checked by |
| --- | --- | --- | --- | --- |
| Copy, cut, paste, Paste Special | Paste all, values, formulas, formats, or as text | 10,000 cells per paste | Works | 5 test files; 89 of 94 on real workbooks (94%) |
| Font | Family, size, bold, italic, underline, font colour, fill | Fonts come from the computer; no embedded fonts | Works | 4 test files; 190 of 198 on real workbooks (95%) |
| Alignment and merge | Horizontal and vertical alignment, wrap, Merge & Center, Unmerge | Not inside a table (Excel refuses that too); beside a table is fine | Works | 3 test files; 281 of 293 on real workbooks (95%) |
| Number formats | General, Number, Percent, Date, Text, Currency, Accounting, decimals | Existing custom codes display; no editor for new custom codes | Works, with limits | 3 test files; 95 of 99 on real workbooks (95%) |
| Insert and delete rows and columns | Insert or delete; formulas, names, filters, rules, links, notes and charts follow | Refused on a sheet with manual page breaks, on a sheet that feeds or holds a pivot table, and when a formula on the sheet cannot be read | Works, with limits | 4 test files; 343 of 392 on real workbooks (87%) |
| Row height, column width | Resize, including on workbooks kept as stored | 8 to 200 points for rows | Works | 3 test files; 190 of 198 on real workbooks (95%) |
| Fill down, right, up, left | Copies values and formulas with correct relative references | Special formulas refuse | Works | 3 test files; 85 of 94 on real workbooks (90%) |
| Fill series: numbers | Preview a series from two starting cells, then apply | One row or column, 10,000 cells | Works | 1 test file |
| Fill series: dates | Day steps in both date systems | Whole days only; no month, year or workday steps | Works, with limits | 1 test file |
| Fill across worksheets | Copy contents, formats or both to chosen sheets | Visible sheets | Works | 2 test files |
| Clear | Contents, formats, all, or hyperlinks |  | Works | 3 test files; 75 of 79 on real workbooks (94%) |
| Find and Replace | Search values, displayed text or formulas; preview then replace | 10,000 matches | Works | 4 test files |
| Go To Special | Constants, formulas, blanks, visible cells, last cell, current region | 10,000 cells selected | Works, with limits | 4 test files |
| Undo and Redo | Every accepted change | 100 steps | Works | 3 test files |
| Conditional formatting | Greater, less, equal, between, text contains, duplicates, with a fill | No data bars or icon sets | Works, with limits | 4 test files; 89 of 94 on real workbooks (94%) |
| Cell styles | Good, Bad, Heading and a few number styles | Small fixed set | Works, with limits | 2 test files |

## Insert (everyday part)

| Feature | What you can do | Limit that matters | Status | Checked by |
| --- | --- | --- | --- | --- |
| Table | Make a table from a selection | Simple tables; totals rows limit some edits | Works, with limits | 3 test files |
| Chart | Column, bar or line chart: range, title, legend; move, resize, delete | Three chart types, eight series, 1,000 cells | Works, with limits | 5 test files |
| Picture | Insert a PNG; move, resize, alt text | PNG only, 128 KB | Works, with limits | 4 test files |
| Hyperlink | Web links and links to cells; edit, remove, open | One cell at a time | Works | 3 test files; 90 of 94 on real workbooks (95%) |
| Checkbox and symbols | A TRUE/FALSE checkbox cell; insert a symbol | Three symbols | Works, with limits | 3 test files |

Shapes, icons, text boxes, WordArt, SmartArt, 3D models and chart styles are in
[FEATURES_SAMPLES.md](FEATURES_SAMPLES.md).

## Page Layout (everyday part)

| Feature | What you can do | Limit that matters | Status | Checked by |
| --- | --- | --- | --- | --- |
| Margins, orientation, paper size, scaling | Set and save them | A4 and Letter | Works | 3 test files |
| Print area and repeated titles | Set from the selection, clear | They follow inserted and deleted rows and columns, renames and sheet copies | Works | 3 test files |
| Print preview and PDF | Preview pages, export a PDF | 100 pages, 10,000 cells; fonts may be substituted | Works, with limits | 4 test files |
| Print | Send to a printer | Not checked on paper | Not verified | not checked without the app window |
| Headers and footers | Left, centre and right text; three presets | Plain single-line text | Works, with limits | 1 test file |

Themes and sheet backgrounds are in [FEATURES_SAMPLES.md](FEATURES_SAMPLES.md).

## Formulas

| Feature | What you can do | Limit that matters | Status | Checked by |
| --- | --- | --- | --- | --- |
| Typing formulas | Any formula can be typed and saved | A formula the editor cannot calculate keeps no result; Excel calculates it on open | Works, with limits | 3 test files; 94 of 99 on real workbooks (94%) |
| Calculation | Automatic or manual; Calculate Now; Calculate Sheet | The function set is what the built-in calculator covers; see [calculator](CALCULATOR_EVALUATION.md) | Works, with limits | 4 test files |
| Insert Function | Guided entry for 26 common functions | 26 of several hundred | Works, with limits | 4 test files |
| AutoSum and recently used | SUM below a column; recent functions |  | Works | 2 test files |
| Names | Define, go to, update, delete; use in a formula; create from labels | Simple range names | Works, with limits | 5 test files |
| Trace, Watch, Error Checking | Arrows for precedents and dependents, a watch list, a list of errors | Not exhaustive on large dependency chains | Works, with limits | 3 test files |
| Show formulas | Show formulas instead of results |  | Works | 2 test files |
| Lookup and conditional sums | SUMIF, COUNTIF, VLOOKUP (exact), XLOOKUP | Exact match | Works | 3 test files |

## Data

| Feature | What you can do | Limit that matters | Status | Checked by |
| --- | --- | --- | --- | --- |
| Sort | Up to 16 columns, ascending or descending, with a header | The sorted range must take in every used column of its rows; refused when the sorted rows hold a merged cell, a link or a note. Rows inside a table can be sorted | Works, with limits | 2 test files; 76 of 97 on real workbooks (78%) |
| Filter | Choose values, reapply, clear, switch off | A filter applied here must be switched off before inserting or deleting rows | Works, with limits | 5 test files; 267 of 291 on real workbooks (91%) |
| Remove Duplicates | Review, then keep the first of each | Compares stored values, not displayed text | Works | 2 test files |
| Text to Columns | Split by delimiter or fixed width, with types | 1,000 rows | Works | 4 test files |
| Data validation | List, whole number, decimal, text length; prompts and error messages | 64 rules. A rule that came with the file and cannot be evaluated here is kept and does not block editing | Works | 3 test files |
| Import CSV or TSV | Choose columns and filters; save as a refreshable query | Local files only | Works | 5 test files |
| Refresh All | Preview then refresh every saved import | Local files only | Works | 4 test files |
| Table editing | Rename, resize, add or delete columns; references follow | Simple tables | Works, with limits | 2 test files |
| Outline, subtotals, consolidate | Not built |  | Not planned |  |

## Review

| Feature | What you can do | Limit that matters | Status | Checked by |
| --- | --- | --- | --- | --- |
| Notes | Add, edit, delete, list and jump to notes | Notes cannot yet be edited on a sheet that arrived with notes from Excel | Works, with limits | 5 test files; 80 of 94 on real workbooks (85%) |
| Spelling | English (US and UK) suggestions for text cells | English only | Not verified | tests exist but are not run (fails and needs repair) |
| Workbook statistics | Counts of cells, formulas and results |  | Works | 3 test files |
| Accessibility check | Missing alt text, blank or repeated table headers |  | Works | 2 test files |

## View

| Feature | What you can do | Limit that matters | Status | Checked by |
| --- | --- | --- | --- | --- |
| Freeze panes | At the selection, first row, first column; unfreeze | 100 rows, 30 columns | Works | 3 test files; 93 of 97 on real workbooks (95%) |
| Zoom | 10% to 400%, zoom to selection |  | Works | 3 test files |
| Gridlines, headings, formula bar | Show or hide | Kept per file on this computer | Works | 2 test files |
| Saved zoom and display from Excel | Opens with the zoom and gridlines Excel saved | First window only | Works | 2 test files |
| Local views | Three presets and your own named views | Kept on this computer, not in the file | Works | 3 test files |
| Page Layout view | See the printed pages | Read only | Works, with limits | 2 test files |
| Page Break Preview | See page outlines over the sheet | Read only; no manual breaks | Works, with limits | 1 test file |

## Pivot tables

| Feature | What you can do | Limit that matters | Status | Checked by |
| --- | --- | --- | --- | --- |
| Create | One row field, SUM or COUNT of one value field, on a new sheet | One pivot per workbook made here; 1,000 source rows | Works, with limits | 3 test files |
| Refresh | Refresh the pivot made here |  | Works, with limits | 2 test files |
| Fields | Change source range, row field, value field | Same limits as Create | Works, with limits | 3 test files |
| Go to source, field list | Inspect and jump to the source |  | Works, with limits | 3 test files |
| Pivots made in Excel | Kept exactly when the workbook is saved | Shown, not editable. Cells beside the pivot can be edited; rows and columns on its sheet, or on the sheet it summarises, cannot be inserted or deleted | Works, with limits | 2 test files |

Pivot styles are in [FEATURES_SAMPLES.md](FEATURES_SAMPLES.md).

## Files and safety

| Feature | What you can do | Limit that matters | Status | Checked by |
| --- | --- | --- | --- | --- |
| Open and save `.xlsx` | Edit and save; everything the editor does not understand is kept as it was | See [compatibility](BIDIRECTIONAL_COMPATIBILITY.md) | Works | 5 test files; 269 of 280 on real workbooks (96%) |
| Open and save `.csv`, `.tsv` | Edit text; encoding, delimiter and line endings kept | No formulas | Works | 5 test files |
| Open `.xlsm`, `.xls`, `.xlsb` | View, save an exact copy, export values to a new `.xlsx` | Read only | Works, with limits | 6 test files |
| Several workbooks | Four tabs in one window | Four | Works, with limits | 1 test file; 278 of 290 on real workbooks (95%) |
| Recovery | Unsaved work is restored after a crash |  | Works | 4 test files |
| Protected or signed workbooks | Open read only; save an exact copy | Not editable. A workbook with any protected sheet opens read only as a whole, which is stricter than Excel | Works, with limits | 2 test files |

## What is measured

Two numbers are recalculated by tools in the repository; they are the real progress figures.

| Measure | Latest | How |
|---|---|---|
| Everyday actions that work on 101 real workbooks | 2,669 of 2,883 (92.6%); 2,750 (95.4%) counting refusals Excel makes too | 30 actions on each workbook |
| Real workbooks that open for editing | 100 of 101 | The other one contains macros |
| Workbooks edited and saved with nothing lost or changed | 97 of 97 saved | A different program reads both files and compares cells, formats and objects |
| Saved files Excel opens, recalculates and saves without repair | 94 of 94 sent | Real Excel 16.113.4 |
| Actions that change only what they should | 2,068 of 2,068 checked | The saved file is compared with the original, apart from the place the action was meant to change |
| Moving actions with the same result as Excel doing the same action | 474 of 525 compared | Insert and delete rows and columns, sort, rename and copy a sheet |
| Everyday actions on generated ordinary workbooks | 411 of 420 (97.9%) | 30 actions on 14 workbooks |

The 101 workbooks are 75 public files (government statistics and open data from Australia, Canada,
Ireland, the United States and the UN humanitarian data exchange) and 26 made for this purpose to
cover what the public files lack, then saved by Excel.

These measure the engine without opening a window. They do not yet show that every row above works
in the app on real files; closing that gap is the [roadmap](ROADMAP.md).

## What is left out

Not planned, because they belong to Microsoft's or Google's own services, or are rare:

- Co-authoring, threaded comments, version history in the cloud, sharing links.
- Stocks, Currencies and other connected data types; Analyze Data; Maps; Forms; Office Scripts.
- Running macros. (Macro workbooks open read only and are kept intact.)
- Password protection, rights management, digital signing.
- Slicers, timelines and PivotCharts.
- Goal Seek, Scenario Manager, Solver, data tables as an editing feature.
- Extra windows, split panes, side-by-side view.
- Translation, Thesaurus, handwriting recognition.
- Windows and Linux builds, and public signed releases, until asked for.
