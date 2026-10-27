// Everyday edits on ordinary sheets, checked by an independent reader.
//
// The workbooks come from scripts/audit/ordinary-sheets.py (openpyxl) and the
// saved result is described by scripts/audit/describe-workbook.py (openpyxl
// again), so nothing here trusts the editor's own view of what it saved.
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { editCalculatedWorkbookResult } from '@office-editors/xlsx-editor/engine';
import { pythonOnPath } from '../scripts/audit/python.mjs';

pythonOnPath();

const scripts = new URL('../scripts/audit/', import.meta.url).pathname;
const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-ordinary-'));
execFileSync('python3', [scripts + 'ordinary-sheets.py', path.join(root, 'source')], { stdio: 'ignore' });
after(async () => {
  closeNativeWorkbookWorker();
  await fs.rm(root, { recursive: true, force: true });
});

let run = 0;
async function edit(name, commands, describedSheet) {
  const directory = path.join(root, 'run-' + ++run),
    source = path.join(directory, 'source.xlsx');
  await fs.mkdir(directory);
  await fs.copyFile(path.join(root, 'source', name + '.xlsx'), source);
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
    await book.open(source, { storageDirectory: path.join(directory, 'index') });
    let state = '0';
    // A command given as a function is built from the view the previous command left.
    let view = book.view(state);
    for (const command of commands) {
      view = (await book.apply(state, [typeof command === 'function' ? command(view) : command])).view;
      state = view.storage.state;
    }
    const saved = path.join(directory, 'saved.xlsx');
    await fs.copyFile((await book.export(state)).path, saved);
    const described = JSON.parse(
      execFileSync('python3', [scripts + 'describe-workbook.py', saved, ...(describedSheet ? [describedSheet] : [])], {
        encoding: 'utf8',
      })
    );
    // Things Excel rejects that a lenient reader accepts: records out of order, a note box on the wrong cell.
    assert.deepEqual(described.problems, [], 'the saved file has a structure Excel would not accept');
    return described;
  } finally {
    await book.close();
  }
}
const sales = (operation, index) => ({ sheet: 'Sales', operation, index });

test('a row inserted in a typical report moves its filter, rule, dropdown, link, note and total', async () => {
  const saved = await edit('13-typical-report', [sales('insertRow', 2)]);
  assert.equal(saved.filter, 'A2:E9');
  assert.deepEqual(saved.rules, ['C4:C9']);
  assert.deepEqual(saved.validation, ['A4:A9']);
  assert.deepEqual(saved.links, ['G4=https://example.com/supplier']);
  assert.deepEqual(saved.notes, ['C4=Check this quantity']);
  assert.deepEqual(saved.merged, ['A1:E1']);
  assert.equal(saved.frozen, 'A3');
  assert.equal(saved.formulas.E10, '=SUM(E4:E9)');
  assert.equal(saved.formulas.E4, '=C4*D4');
});

test('a column deleted from a typical report narrows every range that crossed it', async () => {
  const saved = await edit('13-typical-report', [sales('deleteColumn', 1)]);
  assert.equal(saved.filter, 'A2:D8');
  assert.deepEqual(saved.rules, ['B3:B8']);
  assert.deepEqual(saved.validation, ['A3:A8']);
  assert.deepEqual(saved.links, ['F3=https://example.com/supplier']);
  assert.deepEqual(saved.notes, ['B3=Check this quantity']);
  assert.deepEqual(saved.merged, ['A1:D1']);
  assert.equal(saved.formulas.D9, '=SUM(D3:D8)');
});

// What Excel for Mac does was watched, not assumed (October 2026): a row
// inserted above the last frozen row joins the frozen rows, and deleting a
// frozen row leaves one fewer.
test('a row inserted inside the frozen header is frozen too, as Excel does', async () => {
  const saved = await edit('13-typical-report', [sales('insertRow', 1)]);
  assert.equal(saved.frozen, 'A4');
  assert.equal(saved.filter, 'A3:E9');
  assert.equal(saved.formulas.E10, '=SUM(E4:E9)');
});

test('a row inserted just below the frozen header is not frozen, and deleting a frozen row unfreezes it', async () => {
  assert.equal((await edit('13-typical-report', [sales('insertRow', 2)])).frozen, 'A3');
  assert.equal((await edit('13-typical-report', [sales('deleteRow', 0)])).frozen, 'A2');
  assert.equal((await edit('13-typical-report', [{ ...sales('deleteRow', 0), count: 2 }])).frozen, null);
});

test('manual page breaks follow their rows and columns', async () => {
  // Breaks sit after rows 4 and 7 and after column C.
  assert.deepEqual((await edit('16-page-breaks', [sales('insertRow', 2)])).breaks, { rows: [5, 8], columns: [3] });
  assert.deepEqual((await edit('16-page-breaks', [sales('deleteRow', 5)])).breaks, { rows: [4, 6], columns: [3] });
  assert.deepEqual((await edit('16-page-breaks', [sales('insertColumn', 1)])).breaks, { rows: [4, 7], columns: [4] });
  // The row that starts the second page is deleted: its break goes with it.
  assert.deepEqual((await edit('16-page-breaks', [sales('deleteRow', 4)])).breaks, { rows: [6], columns: [3] });
});

test('a new row is empty and a new column looks like the column on its left, as in Excel', async () => {
  // Excel does not continue formulas into an inserted row outside a table.
  const row = await edit('13-typical-report', [sales('insertRow', 4)]);
  assert.equal(row.formulas.E4, '=C4*D4');
  assert.equal(row.formulas.E5, undefined);
  assert.equal(row.formulas.E6, '=C6*D6');
  // The header row is bold; a column inserted after column A gets a bold header cell.
  const column = await edit('13-typical-report', [sales('insertColumn', 1)]);
  assert.equal(column.styled.A2, true);
  assert.equal(column.styled.B2, true);
});

test('print area, table rows and column formatting behave as in Excel on real-file shapes', async () => {
  const { unzipSync, zipSync, strFromU8, strToU8 } = await import('fflate');
  const rewrite = async (from, to, change) => {
    const parts = unzipSync(await fs.readFile(path.join(root, 'source', from + '.xlsx')));
    change(
      parts,
      (name) => strFromU8(parts[name]),
      (name, text) => (parts[name] = strToU8(text))
    );
    await fs.writeFile(path.join(root, 'source', to + '.xlsx'), zipSync(parts));
  };
  // A print area is a defined name and follows rows, columns and a rename.
  await rewrite('01-plain', '16-print-area', (parts, read, write) =>
    write(
      'xl/workbook.xml',
      read('xl/workbook.xml').replace(
        /<calcPr/,
        '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Sales!$A$1:$E$8</definedName></definedNames><calcPr'
      )
    )
  );
  const printed = await edit('16-print-area', [
    sales('insertRow', 2),
    sales('deleteColumn', 1),
    { sheet: 'Sales', operation: 'renameSheet', newName: 'Orders 2' },
  ]);
  assert.deepEqual(printed.sheets, ['Orders 2']);
  assert.match(String(printed.printArea), /^'Orders 2'!\$A\$1:\$D\$9$/);
  // Column formatting that runs to the last column stops at the edge on insert.
  await rewrite('01-plain', '17-wide-format', (parts, read, write) =>
    write(
      'xl/worksheets/sheet1.xml',
      read('xl/worksheets/sheet1.xml').replace(
        '<sheetData>',
        '<cols><col min="6" max="16384" width="9" customWidth="1"/></cols><sheetData>'
      )
    )
  );
  assert.equal((await edit('17-wide-format', [sales('insertColumn', 1)])).formulas.F8, '=SUM(F2:F7)');
  // Sorting the data rows of a table is the ordinary way to sort; merging beside a table is allowed.
  const sorted = await edit('04-table', [
    { sheet: 'Sales', operation: 'sortRows', range: 'A2:E7', column: 'A', direction: 'asc' },
  ]);
  assert.equal(sorted.values.A2, 'East');
  assert.deepEqual(sorted.tables, ['SalesTable=A1:E7']);
  assert.deepEqual((await edit('04-table', [{ sheet: 'Sales', operation: 'mergeCells', range: 'H3:I3' }])).merged, [
    'H3:I3',
  ]);
  await assert.rejects(edit('04-table', [{ sheet: 'Sales', operation: 'mergeCells', range: 'B3:C3' }]), {
    code: 'table-or-protection',
  });
});

test('a chart keeps pointing at its data when a row is inserted', async () => {
  const saved = await edit('10-chart', [sales('insertRow', 3)]);
  assert.deepEqual(saved.charts, ['Sales!E1', 'Sales!$A$2:$A$8', 'Sales!$E$2:$E$8']);
  assert.equal(saved.formulas.E9, '=SUM(E2:E8)');
});

test('a note follows its cell when a column is inserted, and a filter shrinks when a row is deleted', async () => {
  assert.deepEqual((await edit('12-note', [sales('insertColumn', 1)])).notes, ['D2=Check this quantity']);
  assert.equal((await edit('02-filter', [sales('deleteRow', 3)])).filter, 'A1:E6');
});

test('a filter the editor applied must be switched off before the structure changes', async () => {
  const filter = { sheet: 'Sales', operation: 'filterRows', range: 'A1:E7', column: 'B', values: ['Widget'] };
  // An applied filter records which rows it hid. Until that record can move
  // with the rows, inserting or deleting beside it is refused.
  await assert.rejects(
    edit('02-filter', [filter, sales('insertColumn', 0)]),
    /special formulas or extended|extended worksheet/
  );
});

test('a filter works on a sheet saved by Excel, which adds a prefixed attribute to the sheet defaults', async () => {
  // Excel writes x14ac:dyDescent on sheetFormatPr in nearly every workbook.
  const { unzipSync, zipSync, strFromU8, strToU8 } = await import('fflate');
  const file = path.join(root, 'source', '01-plain.xlsx'),
    excelLike = path.join(root, 'source', '14-excel-defaults.xlsx'),
    parts = unzipSync(await fs.readFile(file)),
    sheet = strFromU8(parts['xl/worksheets/sheet1.xml']);
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    sheet
      .replace('<worksheet ', '<worksheet xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" ')
      .replace(/<sheetFormatPr[^>]*\/>/, '<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>')
  );
  assert.match(strFromU8(parts['xl/worksheets/sheet1.xml']), /x14ac:dyDescent/);
  await fs.writeFile(excelLike, zipSync(parts));
  const saved = await edit('14-excel-defaults', [
    { sheet: 'Sales', operation: 'filterRows', range: 'A1:E7', column: 'B', values: ['Widget'] },
  ]);
  assert.equal(saved.filter, 'A1:E7');
  assert.deepEqual(saved.filterColumns, [1]);
});

test('filtering keeps the hidden dropdown button and the "sorted by" marker Excel had stored', async () => {
  const before = { filterButtons: [[0, true]], sortedBy: [['C2:C7', true]] };
  // Filter another column, the column whose button is hidden, then filter and show everything again.
  const steps = [
    [{ sheet: 'Sales', operation: 'filterRows', range: 'A1:E7', column: 'B', values: ['Widget'] }],
    [{ sheet: 'Sales', operation: 'filterRows', range: 'A1:E7', column: 'A', values: ['North'] }],
    [
      { sheet: 'Sales', operation: 'filterRows', range: 'A1:E7', column: 'B', values: ['Widget'] },
      { sheet: 'Sales', operation: 'filterRows', range: 'A1:E7', column: 'B', values: [] },
    ],
  ];
  for (const commands of steps) {
    const saved = await edit('17-sorted-filter', commands);
    assert.equal(saved.filter, 'A1:E7');
    assert.deepEqual(saved.filterButtons, before.filterButtons);
    assert.deepEqual(saved.sortedBy, before.sortedBy);
  }
  // The filter the editor wrote can be cleared by its revision afterwards.
  // (The gate on real workbooks caught this failing when the marks were first kept.)
  const cleared = await edit('17-sorted-filter', [
    steps[0][0],
    (view) => ({
      sheet: 'Sales',
      operation: 'clearFilter',
      expectedFilterRevision: view.sheets.find((sheet) => sheet.name === 'Sales').filter.revision,
    }),
  ]);
  // Only the column whose button is hidden is still listed; no criteria remain.
  assert.deepEqual(cleared.filterColumns, [0]);
  assert.deepEqual(cleared.sortedBy, before.sortedBy);
  const removed = await edit('17-sorted-filter', [
    steps[0][0],
    (view) => ({
      sheet: 'Sales',
      operation: 'removeFilter',
      expectedFilterRevision: view.sheets.find((sheet) => sheet.name === 'Sales').filter.revision,
    }),
  ]);
  assert.equal(removed.filter, null);
});

test('a protected sheet refuses changes and the other sheets stay editable, as in Excel', async () => {
  const saved = await edit(
    '18-protected-sheet',
    [
      { sheet: 'Notes', cell: 'A2', value: 'Approved' },
      { sheet: 'Notes', operation: 'insertRow', index: 0 },
    ],
    'Notes'
  );
  assert.equal(saved.formulas.B2, '=Sales!E8');
  for (const command of [
    { sheet: 'Sales', cell: 'H1', value: 1 },
    { sheet: 'Sales', operation: 'insertRow', index: 2 },
  ])
    await assert.rejects(edit('18-protected-sheet', [command]), (error) => error.code === 'sheet-protected');
  // The copy that was saved still has its protection.
  assert.equal((await edit('18-protected-sheet', [{ sheet: 'Notes', cell: 'A2', value: 1 }], 'Sales')).protected, true);
});

test('a copy of a sheet with a table gets its own table, and formulas on the copy use it', async () => {
  const copy = [{ sheet: 'Sales', operation: 'duplicateSheet', newName: 'Copy' }];
  const copied = await edit('19-table-formula', copy, 'Copy');
  assert.deepEqual(copied.tables, ['Orders2=A1:E7']);
  assert.equal(copied.formulas.H1, '=SUM(Orders2[Total])');
  const original = await edit('19-table-formula', copy, 'Sales');
  assert.deepEqual(original.tables, ['Orders=A1:E7']);
  assert.equal(original.formulas.H1, '=SUM(Orders[Total])');
});

test('sorting a block narrower than the sheet moves only that block, as in Excel', async () => {
  // Sort Region and Item (columns A and B); Quantity, Price and Total stay where they are.
  const saved = await edit('01-plain', [
    { sheet: 'Sales', operation: 'sortRows', range: 'A2:B7', column: 'A', direction: 'asc' },
  ]);
  const regions = [2, 3, 4, 5, 6, 7].map((row) => saved.values['A' + row]);
  assert.deepEqual(regions, ['East', 'North', 'North', 'South', 'South', 'West']);
  assert.deepEqual(
    [2, 3, 4, 5, 6, 7].map((row) => saved.values['B' + row]),
    ['Widget', 'Widget', 'Gizmo', 'Gadget', 'Widget', 'Gizmo']
  );
  assert.deepEqual(saved.thirdColumn, ['Quantity', 12, 7, 20, 3, 9, 15, null]);
  assert.equal(saved.formulas.E2, '=C2*D2');
  assert.equal(saved.formulas.E7, '=C7*D7');
});

test('sorting leaves a total outside the sorted rows alone and merging works beside formulas', async () => {
  const sorted = await edit('01-plain', [
    { sheet: 'Sales', operation: 'sortRows', range: 'A2:E7', column: 'A', direction: 'asc' },
  ]);
  assert.equal(sorted.values.A2, 'East');
  assert.equal(sorted.formulas.E8, '=SUM(E2:E7)');
  const merged = await edit('13-typical-report', [{ sheet: 'Sales', operation: 'mergeCells', range: 'I5:J5' }]);
  assert.deepEqual(merged.merged, ['A1:E1', 'I5:J5']);
  assert.equal(merged.filter, 'A2:E8');
  assert.deepEqual(merged.notes, ['C3=Check this quantity']);
});

test('sheet names in any script: formulas on another sheet follow inserted rows and a rename', async () => {
  const inserted = await edit('15-other-language', [{ sheet: '数据', operation: 'insertRow', index: 1 }], '汇总');
  assert.equal(inserted.formulas.B1, '=SUM(数据!E3:E8)');
  assert.equal(inserted.formulas.B2, '=数据!C3*2');
  const renamed = await edit(
    '15-other-language',
    [{ sheet: '数据', operation: 'renameSheet', newName: '明细' }],
    '汇总'
  );
  assert.deepEqual(renamed.sheets, ['明细', '汇总']);
  // Quoting a name that does not need it is valid and Excel reads it the same.
  assert.match(renamed.formulas.B1, /^=SUM\('?明细'?!E2:E7\)$/);
});

test('deleting the rows or columns an object covers deletes that object, as Excel does', async () => {
  // The link is in column G and the note in column C of the typical report.
  const withoutLink = await edit('13-typical-report', [sales('deleteColumn', 6)]);
  assert.deepEqual(withoutLink.links, []);
  assert.equal(withoutLink.filter, 'A2:E8');
  const withoutTitle = await edit('03-merged-title', [sales('deleteRow', 0)]);
  assert.deepEqual(withoutTitle.merged, []);
  assert.equal(withoutTitle.formulas.E8, '=SUM(E2:E7)');
});

test('a chart keeps pointing at its sheet when the sheet is renamed', async () => {
  const saved = await edit('10-chart', [{ sheet: 'Sales', operation: 'renameSheet', newName: 'Orders 2' }]);
  assert.deepEqual(saved.charts, ["'Orders 2'!E1", "'Orders 2'!$A$2:$A$7", "'Orders 2'!$E$2:$E$7"]);
});

test('a copy of a sheet carries its filter, rule, dropdown, link, note, merge and frozen header', async () => {
  const copy = await edit(
    '13-typical-report',
    [{ sheet: 'Sales', operation: 'duplicateSheet', newName: 'Copy of sales' }],
    'Copy of sales'
  );
  assert.deepEqual(copy.sheets, ['Sales', 'Copy of sales']);
  assert.equal(copy.filter, 'A2:E8');
  assert.deepEqual(copy.rules, ['C3:C8']);
  assert.deepEqual(copy.validation, ['A3:A8']);
  assert.deepEqual(copy.links, ['G3=https://example.com/supplier']);
  assert.deepEqual(copy.notes, ['C3=Check this quantity']);
  assert.deepEqual(copy.merged, ['A1:E1']);
  assert.equal(copy.frozen, 'A3');
  assert.equal(copy.formulas.E9, '=SUM(E3:E8)');
});

test('sorting is judged by the sorted rows: a merged title above them and a name elsewhere do not stop it', async () => {
  const titled = await edit('03-merged-title', [
    { sheet: 'Sales', operation: 'sortRows', range: 'A3:E8', column: 'A', direction: 'asc' },
  ]);
  assert.deepEqual(titled.merged, ['A1:E1']);
  assert.equal(titled.values.A3, 'East');
  // Column F multiplies by a name that points at another sheet; the name is unaffected by the sort.
  const named = await edit('11-names-two-sheets', [
    { sheet: 'Sales', operation: 'sortRows', range: 'A2:F7', column: 'A', direction: 'asc' },
  ]);
  assert.equal(named.values.A2, 'East');
  assert.equal(named.formulas.F2, '=E2*(1+TaxRate)');
  // A note sits on a sorted row here; moving it with its cell is not done yet, so this is refused.
  await assert.rejects(
    edit('12-note', [{ sheet: 'Sales', operation: 'sortRows', range: 'A2:E7', column: 'A', direction: 'asc' }])
  );
});
