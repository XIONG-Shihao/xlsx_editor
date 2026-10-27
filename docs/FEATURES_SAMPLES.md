---
status: current
---

# Features kept at sample level

Some Excel features are mostly a library: hundreds of shapes, icons, themes, chart styles. Building
the library is a lot of work and is not what makes a spreadsheet useful day to day.

For these features the editor keeps three things and stops there for now:

1. **The control** in the ribbon, in the place Excel has it.
2. **The programming interface** behind it (a typed command that inserts, edits and deletes the
   object, with Undo, Save and reopen).
3. **A few samples** that prove the control and the interface work end to end.

Adding more samples later is content work. It should not need new plumbing.

Objects of these kinds that arrive in a workbook from Excel are kept exactly when the workbook is
saved, whether or not the editor can show or edit them.

## The sample-level features

| Tab | Feature | Samples available | What works on a sample | Full library |
|---|---|---|---|---|
| Insert | Shapes | Rectangle, Ellipse, Right arrow | Insert, text, alt text, move, resize, delete | Later |
| Insert | Icons | Check, Star, Plus | Insert, colour, move, resize, delete | Later |
| Insert | Text box | Plain, Note | Insert, edit text, move, resize, delete | Later |
| Insert | WordArt | Blue Fill, Orange Outline, Purple Shadow | Insert, edit text, change style, move, resize | Later |
| Insert | SmartArt | Process, Hierarchy, Cycle | Insert, edit 2 to 6 labels, change layout, move, resize | Later |
| Insert | 3D models | Cube, Pyramid, Octahedron | Insert, colour, rotate, move, resize | Later |
| Insert | Recommended charts and chart styles | Three recommendations; Colorful, Monochrome, Outline | Preview, insert an editable chart, change style | Later |
| Insert | Symbols | Ω, ±, ° | Insert into a text cell | Later |
| Draw | Pens | Black pen, Red pen, Yellow highlighter | Draw, select, move, erase a stroke | Later |
| Page Layout | Themes | Neutral, Forest, Plum | Preview, apply, print with the theme | Later |
| Page Layout | Sheet background | Graph paper, Soft dots | Apply, remove | Later |
| Home | Fonts and cell styles | Arial, Georgia, Verdana; Good, Bad, Heading | Apply, Undo, Save, reopen | Later |
| Pivot tables | Pivot styles | Tabular Light, Compact Blue, Outline Dark | Preview, apply, grand totals on or off | Later |
| Automate | Saved actions | Freeze panes, Format range, Recalculate, Sort rows, Unhide all | Save an action, preview, run | Later |

All of these are fixed in size and style beyond what the row says (for example one font and size
inside a shape). That is deliberate at this level.

## Rules for this group

- A sample-level feature is never counted toward the everyday goal, in either direction.
- Do not grow a library until the everyday goal is met, unless a single missing item blocks real
  use.
- New samples must go through the same command as the existing ones. If a new sample needs new
  plumbing, the interface was wrong; fix the interface first.
- Excel must open a saved sample without repairing the file. That is checked the same way as
  everyday edits; see [compatibility](BIDIRECTIONAL_COMPATIBILITY.md).

## Not at sample level, and not planned

Pressure-sensitive ink, handwriting recognition, arbitrary SVG import, stock images, online
pictures, 3D model import from files, and theme or style galleries downloaded from a service.
