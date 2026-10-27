// Release check with real Microsoft Excel.
//
// Edits ordinary workbooks with the editor, then asks Excel to open each saved
// file, recalculate it and save it again. A file Excel has to repair, cannot
// open, or reads differently fails the check.
//
//   node --import ./scripts/ts-resolve-register.mjs scripts/audit/excel-check.mjs [--keep]
//
// Needs macOS, Microsoft Excel and permission for this program to control Excel
// (macOS asks once). openpyxl is found or prepared by ./python.mjs. It opens Excel on screen and
// works only on its own disposable files, in a folder named workbook-excel-check-*
// inside Excel's own container (the one place Excel opens without asking).
// Run it attended: if Excel shows a repair prompt, that case times out and is
// reported; dismiss the prompt to let the rest continue.
import fs from 'node:fs/promises';
import { existsSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync, spawnSync } from 'node:child_process';
import { IndexedWorkbook } from '../../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../../core/native-runtime.mjs';
import { editCalculatedWorkbookResult } from '@office-editors/xlsx-editor/engine';
import { pythonOnPath } from './python.mjs';
import { excelDone } from './excel.mjs';

pythonOnPath();

const component = new URL('../../', import.meta.url).pathname,
  scripts = new URL('./', import.meta.url).pathname,
  // Excel is sandboxed and asks a person for access to files anywhere else;
  // inside its own container it opens them without a prompt.
  container = path.join(os.homedir(), 'Library/Containers/com.microsoft.Excel/Data/Documents'),
  run = path.join(
    existsSync(container) ? container : path.join(component, '.build/runs'),
    'workbook-excel-check-' + new Date().toISOString().replace(/[:.]/g, '-')
  );
await fs.mkdir(path.join(run, 'index'), { recursive: true });
execFileSync('python3', [scripts + 'ordinary-sheets.py', path.join(run, 'source')], { stdio: 'ignore' });

const sales = (operation, extra) => ({ sheet: 'Sales', operation, ...extra });
/** name, source workbook, edits, and the total Excel must calculate (or null). */
const cases = [
  ['row inserted in a typical report', '13-typical-report', [sales('insertRow', { index: 2 })], 658.5],
  ['column deleted from a typical report', '13-typical-report', [sales('deleteColumn', { index: 1 })], 658.5],
  [
    'cells merged beside a filter, rule and note',
    '13-typical-report',
    [sales('mergeCells', { range: 'I5:J5' })],
    658.5,
  ],
  ['row inserted above a chart', '10-chart', [sales('insertRow', { index: 3 })], 658.5],
  ['column inserted beside a note', '12-note', [sales('insertColumn', { index: 1 })], 658.5],
  ['row deleted inside a filter', '02-filter', [sales('deleteRow', { index: 3 })], 568.5],
  [
    'rows sorted under a total',
    '01-plain',
    [sales('sortRows', { range: 'A2:E7', column: 'A', direction: 'asc' })],
    658.5,
  ],
  [
    'highlight rule and bold header added',
    '05-highlight-rule',
    [
      sales('addConditionalRule', { range: 'D2:D7', rule: 'greaterThan', threshold: 25, fillColor: 'C6EFCE' }),
      { type: 'formatRange', sheet: 'Sales', range: 'A1:E1', format: { bold: true, fillColor: 'FFF2CC' } },
    ],
    658.5,
  ],
  [
    'row inserted below a frozen header',
    '09-frozen-formatted',
    [sales('insertRow', { index: 3 }), sales('resizeColumn', { index: 1, width: 22 })],
    658.5,
  ],
  [
    'sheet renamed with a name on another sheet',
    '11-names-two-sheets',
    [sales('renameSheet', { newName: 'Orders' })],
    658.5,
  ],
  ['a value changed', '01-plain', [{ sheet: 'Sales', cell: 'C2', value: 13 }], 663],
  ['column inserted in a table', '04-table', [sales('insertColumn', { index: 6 })], 658.5],
  ['a typical report duplicated', '13-typical-report', [sales('duplicateSheet', { newName: 'Copy of sales' })], 658.5],
  ['a sheet with a note duplicated', '12-note', [sales('duplicateSheet', { newName: 'Copy of sales' })], 658.5],
  ['a chart sheet renamed', '10-chart', [sales('renameSheet', { newName: 'Orders 2' })], 658.5],
  ['the column holding a link deleted', '13-typical-report', [sales('deleteColumn', { index: 6 })], 658.5],
  ['the row holding a merged title deleted', '03-merged-title', [sales('deleteRow', { index: 0 })], 658.5],
  [
    'rows sorted under a merged title',
    '03-merged-title',
    [sales('sortRows', { range: 'A3:E8', column: 'A', direction: 'asc' })],
    658.5,
  ],
  ['row inserted inside the frozen header', '13-typical-report', [sales('insertRow', { index: 1 })], 658.5],
  [
    'table rows sorted and cells merged beside the table',
    '04-table',
    [sales('sortRows', { range: 'A2:E7', column: 'A', direction: 'asc' }), sales('mergeCells', { range: 'H3:I3' })],
    null,
  ],
  [
    'highlight rule added to a typical report',
    '13-typical-report',
    [sales('addConditionalRule', { range: 'D3:D8', rule: 'greaterThan', threshold: 25, fillColor: 'C6EFCE' })],
    658.5,
  ],
  [
    'row inserted on a sheet named in another script',
    '15-other-language',
    [{ sheet: '数据', operation: 'insertRow', index: 1 }],
    658.5,
  ],
  ['row inserted above manual page breaks', '16-page-breaks', [sales('insertRow', { index: 2 })], null],
  ['column inserted before a manual page break', '16-page-breaks', [sales('insertColumn', { index: 1 })], null],
  [
    'a filter applied where Excel had stored a sort and a hidden button',
    '17-sorted-filter',
    [sales('filterRows', { range: 'A1:E7', column: 'B', values: ['Widget'] })],
    null,
  ],
  ['a sheet edited beside a protected one', '18-protected-sheet', [{ sheet: 'Notes', cell: 'A2', value: 'Ok' }], null],
  ['a sheet with a table duplicated', '19-table-formula', [sales('duplicateSheet', { newName: 'Copy' })], null],
  ['column inserted after a formatted column', '13-typical-report', [sales('insertColumn', { index: 1 })], 658.5],
  ['a frozen row deleted', '13-typical-report', [sales('deleteRow', { index: 0 })], 658.5],
];

async function editorSave(source, commands, output) {
  const book = new IndexedWorkbook(
    native,
    async (file, edits, out) => {
      const result = await editCalculatedWorkbookResult(await fs.readFile(file), edits);
      await fs.writeFile(out, result.bytes);
      return { revision: result.view.revision, edits: result.edits };
    },
    (input) => calculateIndexedWorkbook(input, native)
  );
  try {
    await book.open(source, { storageDirectory: path.join(run, 'index') });
    let state = '0';
    for (const command of commands) state = (await book.apply(state, [command])).view.storage.state;
    await fs.copyFile((await book.export(state)).path, output);
  } finally {
    await book.close();
  }
}
const describe = (file, sheet) =>
  JSON.parse(
    execFileSync('python3', [scripts + 'describe-workbook.py', file, ...(sheet ? [sheet] : [])], { encoding: 'utf8' })
  );
const totalOf = (file) =>
  Number(
    execFileSync(
      'python3',
      [
        '-c',
        [
          'import sys',
          'from openpyxl import load_workbook',
          'formulas = load_workbook(sys.argv[1]).worksheets[0]',
          'values = load_workbook(sys.argv[1], data_only=True).worksheets[0]',
          "totals = [c for r in formulas.iter_rows() for c in r if isinstance(c.value, str) and c.value.startswith('=SUM(')]",
          "print(values[totals[0].coordinate].value if totals else 'nan')",
        ].join('\n'),
        file,
      ],
      { encoding: 'utf8' }
    )
  );

const excelScript = path.join(run, 'open-save.applescript');
await fs.writeFile(
  excelScript,
  `on run argv
 set ownedFile to item 1 of argv
 set ownedBook to missing value
 tell application "Microsoft Excel"
  with timeout of 40 seconds
   try
    open (POSIX file ownedFile)
    set matches to every workbook whose full name is ownedFile
    if (count of matches) is not 1 then error "Excel did not open the file under its own name (it may have repaired it)"
    set ownedBook to item 1 of matches
    calculate
    save ownedBook
    close ownedBook saving no
    return "saved"
   on error errText number errNumber
    if ownedBook is not missing value then
     try
      close ownedBook saving no
     end try
    end if
    error errText number errNumber
   end try
  end timeout
 end tell
end run
`
);
const version = spawnSync('osascript', ['-e', 'tell application "Microsoft Excel" to get version'], {
  encoding: 'utf8',
  timeout: 60000,
});
if (version.status !== 0) {
  console.error('Excel cannot be controlled from here: ' + (version.stderr || 'no answer').trim());
  process.exit(2);
}
console.log('Microsoft Excel ' + version.stdout.trim() + '\n');

const compared = ['filter', 'merged', 'frozen', 'rules', 'validation', 'links', 'notes', 'tables', 'formulas'];
const results = [];
for (const [index, [name, source, commands, total]] of cases.entries()) {
  const label = String(index + 1).padStart(2, '0'),
    saved = path.join(run, `${label}-editor.xlsx`),
    reopened = path.join(run, `${label}-excel.xlsx`),
    result = { name, problems: [] };
  try {
    await editorSave(path.join(run, 'source', source + '.xlsx'), commands, saved);
    await fs.copyFile(saved, reopened);
    const opened = spawnSync('osascript', [excelScript, reopened], { encoding: 'utf8', timeout: 70000 });
    if (opened.status !== 0)
      result.problems.push(
        'Excel: ' + (opened.stderr || 'timed out, possibly on a repair prompt').trim().slice(0, 200)
      );
    else {
      const before = describe(saved),
        after = describe(reopened);
      for (const problem of before.problems) result.problems.push('saved file: ' + problem);
      for (const key of compared)
        if (JSON.stringify(before[key]) !== JSON.stringify(after[key]))
          result.problems.push(
            `${key}: editor saved ${JSON.stringify(before[key])}, Excel read ${JSON.stringify(after[key])}`.slice(
              0,
              300
            )
          );
      if (before.charts.length !== after.charts.length)
        result.problems.push(`charts: ${before.charts.length} references saved, ${after.charts.length} after Excel`);
      const calculated = totalOf(reopened);
      if (total !== null && Math.abs(calculated - total) > 1e-9)
        result.problems.push(`Excel calculated a total of ${calculated}, expected ${total}`);
    }
  } catch (error) {
    result.problems.push('editor: ' + String(error?.message ?? error).slice(0, 200));
  }
  results.push(result);
  console.log((result.problems.length ? 'FAIL ' : 'ok   ') + name);
  for (const problem of result.problems) console.log('       ' + problem);
}
closeNativeWorkbookWorker();
// Leave Excel as it was found: quit it only if this check started it.
excelDone();
const failed = results.filter((result) => result.problems.length);
await fs.writeFile(path.join(run, 'report.json'), JSON.stringify({ excel: version.stdout.trim(), results }, null, 1));
console.log(
  `\n${results.length - failed.length} of ${results.length} files opened in Excel, recalculated and saved with the same content`
);
if (process.argv.includes('--keep') || failed.length) console.log('Files kept in ' + run);
else await fs.rm(run, { recursive: true, force: true });
process.exitCode = failed.length ? 1 : 0;
