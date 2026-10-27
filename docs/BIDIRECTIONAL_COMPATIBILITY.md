---
status: current
---

# Working with Excel in both directions

## The promise

1. **Excel to the editor.** Open a workbook that Excel saved, edit it here, save it: nothing that was
   in the workbook is lost.
2. **The editor to Excel.** Open in Excel a workbook that the editor saved, edit it there, save it:
   nothing that was in the workbook is lost, and Excel does not have to repair the file.

"Excel" means Microsoft Excel. It is the only outside reference. Workbooks from WPS, LibreOffice or
Google Sheets are ordinary input: they must open and edit safely, but they are not what the promise
is measured against.

## What counts as lost

Loss is judged the way a person would judge it, not by comparing bytes.

| Lost | Not lost |
|---|---|
| A cell's value, formula or format is gone or different | The file's bytes differ |
| An object is gone or points at the wrong cells: chart, picture, note, link, filter, rule, dropdown, table, name, pivot | Formulas that Excel stored once for a whole column are written out cell by cell |
| A number is shown that is no longer true (a result calculated before an input changed) | A stored result is removed so that Excel calculates it again on open |
| Excel has to repair the file to open it (repair throws content away) | Attributes or records are written in a different but valid order |
| A part of the file the editor does not understand is dropped: macros, embedded files, custom data | Internal ids and timestamps change |

## How the editor keeps the promise

**The original file stays in charge.** The editor does not rebuild a workbook from its own picture of
it. It edits the parts of the file that an action changes and copies every other part as it was. A
chart type the editor cannot draw, a macro, an embedded document: these pass through untouched.

Three rules decide what an action may do.

### Rule 1: judge an action by what it touches

An action is checked against the things it actually affects, not against everything in the workbook.

- Inserting a row on sheet A is checked against sheet A, and against the formulas, names, charts and
  pivots elsewhere that point at sheet A.
- A pivot table on sheet C that summarises sheet D has no say in it.
- A link to another workbook has no say in it either: its cells are not in this file.

This rule replaced an earlier habit of locking the whole workbook when anything unfamiliar was found
anywhere in it. That habit kept files safe by making them unusable: one real 21-sheet financial
workbook refused 20 of 30 everyday actions for a reason that concerned none of them.

### Rule 2: what cannot be understood is carried, not guessed

If the editor cannot read something and the action does not touch it, it is copied through
unchanged. If the action does touch it, that one action is refused and says why. Nothing else is
blocked.

Examples that are carried through: a reference into another workbook, a name whose target was
deleted (`#REF!`), a formula using a function the editor does not know, an extension record from a
newer Excel.

### Rule 3: do what Excel does

Where Excel has a behaviour, the editor follows it instead of refusing. "What Excel does" is found by
watching Excel, not by remembering it: the same action is given to Excel and to the editor, and the
two results are compared (see [How it is checked](#how-it-is-checked)). That comparison corrected
several things the editor had been doing from memory.

- Delete the rows a merged range, a link or a dropdown covers, and that object is deleted with them.
- Insert rows in the middle of a column of filled-down formulas, and each cell gets its own formula.
- Merge cells on a sheet with a table: refused, because Excel refuses it too.
- Insert a row above the last frozen row and it is frozen too; insert just below the frozen rows and
  it is not; delete a frozen row and one fewer is frozen.
- A new row is empty (outside a table) and takes the height, format and outline group of the row
  above. A new column takes the formats of the column on its left.
- A highlight rule or dropdown that ends at the row above, or the column left of, the insertion
  takes in the new cells.
- A name is not stretched by a row inserted just below its range. A range that runs to the last row
  of the sheet still does after a row is deleted.
- A protected sheet refuses changes; the other sheets of the workbook do not.

**Where the editor knowingly differs from Excel.**

| Action | Excel | The editor | Why |
|---|---|---|---|
| Rename a sheet that a link inside the workbook points to | Leaves the link pointing at the old name (it breaks) | Updates the link | A broken link is never what the person wanted |
| Insert a row between bordered rows | Works out the new row's borders from both neighbours | Copies the borders of the row above | Not built yet; values and other formats match |
| Sort rows that have borders | Leaves borders where they are | Moves each row's borders with the row | Not built yet |
| Insert or delete a row in a filtered range | Leaves the "sorted by" record as it was | Moves it with the range | Harmless either way |
| Copy a sheet | Places the copy where the person chooses | Places it last | A choice of position is not offered yet |

### Results that may be out of date

A workbook stores the last result of every formula. If an input changes and the editor cannot
recalculate a formula that depends on it, the stored result is now wrong.

The editor never saves a result it knows may be wrong. It removes the stored result and marks the
workbook so that Excel calculates everything when it opens. Until then the editor shows the cell as
"not calculated", not as the old number.

A formula the editor cannot calculate is still saved exactly as typed.

### Refusals

When an action would break the promise, that action is refused, the workbook is unchanged, and the
message says what is in the way. A refusal is a defect to be removed when it stops everyday work, not
a safe resting place. Refusals on everyday actions are counted; see
[What is measured](FEATURES.md#what-is-measured).

## How it is checked

Four checks, each catching what the others miss. The editor reading back its own file is not one of
them: a writer and reader that share a mistake agree with each other.

| Check | What it does | When |
|---|---|---|
| Independent reader | A different program (openpyxl) reads the saved file and reports its filters, merges, rules, dropdowns, links, notes, tables, names, charts and formulas. Tests compare that report with what the edit should have produced. | Every test run |
| Strict structure | The saved file is checked for things Excel rejects and lenient readers accept: records out of order, a note whose box names another cell, a link to a part that is missing. | Every test run |
| Round trip and every action on real workbooks | Each of 101 real workbooks is opened, one cell is edited, and it is saved. Then each of 30 everyday actions is performed and its result compared with the original or with Excel doing the same. The independent reader then compares the saved file with the original: every cell's content and format, merges, rules, dropdowns, links, notes, tables, names, page setup, hidden and grouped rows, charts, pictures, pivots, embedded files. | With the scoreboard |
| Real Excel | Excel opens each saved file, recalculates it and saves it. The file must open without repair, and Excel's copy must have the same content. | With the scoreboard, on a Mac with Excel |

**Latest result.** 97 of 101 workbooks were saved, and in all 97 nothing was lost or
changed. (The other four were not saved: three have a protected first sheet, where the one-cell edit
is refused as it is in Excel, and one contains macros and opens read only.) Excel was given 94 of
them and opened, recalculated and saved all 94 without repair. Three were not sent because they link
to other workbooks, and Excel asks a person about that.

**Every action, not only the one-cell edit.** Each of the 30 everyday actions is now checked on each
workbook. An action that should change one place (typing, a format, a merge, a note) is compared
with the original: in 2,068 of 2,068 cases nothing else changed. An action that moves things (insert
or delete a row or column, sort, rename or copy a sheet) is also given to Excel, and the two results
are compared: 474 of 525 were the same. The 51 that were not are the differences in the table under
[Rule 3](#rule-3-do-what-excel-does), and the number Excel puts on a copied table's name.

**Telling Excel's changes from the editor's.** Excel tidies a file as it saves: it fills in default
fonts, drops formatting from unused cells, re-encodes pictures. In 18 workbooks Excel's copy differed
from the editor's file for that reason. To be sure the editor was not the cause, Excel also re-saved
the original of each, and the two Excel copies were compared: all 18 matched. What Excel does to the
original and to the editor's file alike is not a loss caused here.

The real-Excel check earns its place. On its first run it failed 5 of 12 generated files that every
other test had passed: merged cells written in the wrong position beside a filter, and notes whose
position was not moved with their rows. Both were fixed and both now have a strict-structure test.

Commands, the files used and the latest results are in [internal/VERIFICATION.md](internal/VERIFICATION.md).

## What is not yet true

Stated so nobody assumes otherwise.

- **Thirty actions, not every action.** The detailed check covers the 30 measured actions on 101
  workbooks. Other edits are checked on generated workbooks (29 kinds, all accepted by Excel).
- **The Excel comparison was not repeated in full on the last changes.** Changes made after the
  last full run were compared with Excel on the eight workbooks they affect.
- **Objects the editor creates itself** (charts, pictures, shapes and the other samples) were checked
  in Excel earlier in the project, on an older build. They are not yet part of these checks.
- **Excel on a Mac only.** Excel for Windows has not been used.
- **Workbooks that link to other workbooks** cannot be sent to Excel by a script, because Excel asks
  a person for access to the linked file. Three of the 101 are in that position.
- **Results that may be out of date.** In a workbook kept as stored (a formula form the editor does
  not understand), formulas are kept exactly and never recalculated here. Values, formats, widths,
  frozen panes and merges can be changed; Excel recalculates on open. One of the 101 opens that way.
- **Not editable at all:** password-protected, signed and macro workbooks. They open read only and
  can be saved as an exact copy. Excel does edit a workbook with macros, leaving the macros alone;
  one of the 101 is such a workbook. Allowing that here is a decision still to be taken.
- **A sheet with protection** refuses changes, as in Excel, but cannot be unprotected here.
- **Other programs.** Round trips through WPS, LibreOffice or Google Sheets are not checked.
