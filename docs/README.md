---
status: current
---

# Spreadsheet editor: documentation

## The goal

**A spreadsheet app that does the everyday work on real Excel files, and never loses anything going
to or from Excel.**

It is not a copy of Excel. Three sentences define it.

1. **Everyday coverage.** The things people do with spreadsheets most days work, on the workbooks
   people really have. Target: 95% of everyday actions succeed on real workbooks, and the rest say
   plainly why not.
2. **Safe both ways.** A workbook from Excel that is edited here and saved has lost nothing. A
   workbook saved here that is opened, edited and saved in Excel has lost nothing, and Excel never
   has to repair it.
3. **Deliberately small where Excel is large.** Features that are mostly a library of assets (shapes,
   icons, themes) keep their control, their interface and a few samples. Features that belong to
   Microsoft's or Google's own services, and rare edge cases, are left out.

### Where it stands

Measured on a standing set of 101 real workbooks: 75 public files (government statistics and open
data) and 26 made to cover what those lack.

| | Target | Now | First measured |
|---|---|---|---|
| Everyday actions that work | 95% | 92.6% (2,669 of 2,883) | 68.1% |
| ...or are refused exactly where Excel refuses too (a protected sheet) | | 95.4% (2,750 of 2,883) | |
| Workbooks that open for editing | all | 100 of 101 | 87 |
| Workbooks edited and saved with nothing lost or changed | all | 97 of 97 saved | 84 of 84 |
| Saved files that Excel opens, recalculates and saves without repair | all | 94 of 94 sent | 81 of 81 |
| Actions that change only what they should, checked against the original | all | 2,068 of 2,068 | not measured |
| Actions that move things, with the same result as Excel doing the same action | all | 474 of 525 | not measured |
| Wrong answers from the calculator, against results Excel stored | none | none in 620 | not measured |

The second half of the goal holds on every file tested: nothing was lost in either direction. The
first half is at 92.6% against a target of 95%, or 95.4% if a refusal that Excel also makes is
counted as the right answer. What stands between is listed, with counts, in the
[roadmap](ROADMAP.md). The last two rows are new: each action is now checked in detail, and where the
editor's result differs from Excel's the difference is listed in
[Working with Excel in both directions](BIDIRECTIONAL_COMPATIBILITY.md).

All of this is measured without opening the app's window. No one has yet walked the feature list by
hand in the built app.

## What to read

Written to be read by the product owner:

| Document | What it tells you |
|---|---|
| [ROADMAP.md](ROADMAP.md) | The next big steps toward the goal, in order |
| [FEATURES.md](FEATURES.md) | What a person can do, tab by tab, with status: the coverage checklist |
| [FEATURES_SAMPLES.md](FEATURES_SAMPLES.md) | The asset-heavy features kept at sample level |
| [BIDIRECTIONAL_COMPATIBILITY.md](BIDIRECTIONAL_COMPATIBILITY.md) | The promise about Excel in both directions, and how it is kept and checked |
| [XLSX_EDITOR.md](XLSX_EDITOR.md) | What the editor is, how it is built, its limits |
| [XLSX_EDITOR_INTEGRATION.md](XLSX_EDITOR_INTEGRATION.md) | How the standalone app uses it |
| [CALCULATOR_EVALUATION.md](CALCULATOR_EVALUATION.md) | Why the current calculator was kept |

Written for whoever is doing the engineering (working notes, not polished):

| Document | What it holds |
|---|---|
| [DEVELOPMENT.md](DEVELOPMENT.md) | Folders, build, test and measuring tools |
| [internal/PLAN.md](internal/PLAN.md) | The implementation plan for the goal |
| [internal/VERIFICATION.md](internal/VERIFICATION.md) | How each claim is checked: commands, files, what a pass means |
| [internal/feature-tests.json](internal/feature-tests.json) | Which tests and measured actions stand behind each feature row |
