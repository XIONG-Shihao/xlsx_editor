// How often does the editor say "cannot" to an everyday action?
//
// Tries thirty common actions on each workbook: open a fresh copy, perform the
// action, save, and reopen the saved file. Every refusal is recorded with the
// editor's own message. This is a measurement, not a pass/fail test.
//
//   node --import ./scripts/ts-resolve-register.mjs scripts/audit/refusals.mjs [--json report.json] FILE_OR_DIRECTORY...
//
// With no files it builds the ordinary workbooks from ordinary-sheets.py
// (needs python3 with openpyxl). Real files are only read; nothing is written
// next to them.
import fs from 'node:fs/promises';
import { pythonOnPath } from './python.mjs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { IndexedWorkbook } from '../../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../../core/indexed-calculation.mjs';
import { readWorkbook } from '../../core/workbook-package.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../../core/native-runtime.mjs';
import { editCalculatedWorkbookResult } from '@office-editors/xlsx-editor/engine';

const letters = (index) => {
  let name = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
};
const state = (result) => result.view.storage.state;

/** Each action receives the open workbook and facts about its main sheet. */
export const actions = [
  ['Type a number in an empty cell', (b, s) => b.apply('0', [{ sheet: s.name, cell: s.free, value: 42 }])],
  ['Type text in an empty cell', (b, s) => b.apply('0', [{ sheet: s.name, cell: s.free, value: 'Checked' }])],
  [
    'Change an existing number',
    (b, s) => s.number && b.apply('0', [{ sheet: s.name, cell: s.number.cell, value: s.number.value + 1 }]),
  ],
  [
    'Enter a SUM formula',
    (b, s) => b.apply('0', [{ sheet: s.name, cell: s.free, formula: `SUM(${s.column}:${s.column})` }]),
  ],
  [
    'Clear a cell',
    (b, s) =>
      s.number && b.apply('0', [{ type: 'clearRange', sheet: s.name, range: `${s.number.cell}:${s.number.cell}` }]),
  ],
  [
    'Make the header row bold',
    (b, s) => b.apply('0', [{ type: 'formatRange', sheet: s.name, range: s.header, format: { bold: true } }]),
  ],
  [
    'Fill cells with a colour',
    (b, s) => b.apply('0', [{ type: 'formatRange', sheet: s.name, range: s.header, format: { fillColor: 'FFF2CC' } }]),
  ],
  [
    'Show numbers as percent',
    (b, s) =>
      b.apply('0', [{ type: 'formatRange', sheet: s.name, range: s.body, format: { numberFormat: 'percent' } }]),
  ],
  [
    'Centre and wrap text',
    (b, s) =>
      b.apply('0', [
        { type: 'formatRange', sheet: s.name, range: s.header, format: { horizontalAlign: 'center', wrapText: true } },
      ]),
  ],
  ['Widen a column', (b, s) => b.apply('0', [{ sheet: s.name, operation: 'resizeColumn', index: 0, width: 24 }])],
  ['Make a row taller', (b, s) => b.apply('0', [{ sheet: s.name, operation: 'resizeRow', index: 0, height: 30 }])],
  ['Insert a row in the data', (b, s) => b.apply('0', [{ sheet: s.name, operation: 'insertRow', index: s.middle }])],
  ['Delete a row from the data', (b, s) => b.apply('0', [{ sheet: s.name, operation: 'deleteRow', index: s.middle }])],
  ['Insert a column in the data', (b, s) => b.apply('0', [{ sheet: s.name, operation: 'insertColumn', index: 1 }])],
  ['Delete a column from the data', (b, s) => b.apply('0', [{ sheet: s.name, operation: 'deleteColumn', index: 1 }])],
  [
    'Sort by the first column',
    (b, s) => b.apply('0', [{ sheet: s.name, operation: 'sortRows', range: s.rows, column: 'A', direction: 'asc' }]),
  ],
  [
    'Filter by a value',
    (b, s) =>
      b.apply('0', [{ sheet: s.name, operation: 'filterRows', range: s.table, column: 'A', values: [s.firstValue] }]),
  ],
  [
    'Filter, then show all rows again',
    async (b, s) => {
      const filtered = await b.apply('0', [
        { sheet: s.name, operation: 'filterRows', range: s.table, column: 'A', values: [s.firstValue] },
      ]);
      return b.apply(state(filtered), [
        { sheet: s.name, operation: 'filterRows', range: s.table, column: 'A', values: [] },
      ]);
    },
  ],
  [
    'Filter, then switch the filter off',
    async (b, s) => {
      const filtered = await b.apply('0', [
        { sheet: s.name, operation: 'filterRows', range: s.table, column: 'A', values: [s.firstValue] },
      ]);
      const revision = filtered.view.sheets.find((sheet) => sheet.name === s.name)?.filter?.revision;
      return b.apply(state(filtered), [{ sheet: s.name, operation: 'removeFilter', expectedFilterRevision: revision }]);
    },
  ],
  ['Merge two empty cells', (b, s) => b.apply('0', [{ sheet: s.name, operation: 'mergeCells', range: s.freePair }])],
  [
    'Merge, then unmerge',
    async (b, s) => {
      const merged = await b.apply('0', [{ sheet: s.name, operation: 'mergeCells', range: s.freePair }]);
      return b.apply(state(merged), [{ sheet: s.name, operation: 'unmergeCells', range: s.freePair }]);
    },
  ],
  ['Freeze the top row', (b, s) => b.apply('0', [{ sheet: s.name, operation: 'freeze', rows: 1, columns: 0 }])],
  ['Add a worksheet', (b, s) => b.apply('0', [{ sheet: 'Added sheet', operation: 'addSheet' }])],
  ['Rename the worksheet', (b, s) => b.apply('0', [{ sheet: s.name, operation: 'renameSheet', newName: 'Renamed' }])],
  [
    'Duplicate the worksheet',
    (b, s) => b.apply('0', [{ sheet: s.name, operation: 'duplicateSheet', newName: 'Copy of sheet' }]),
  ],
  [
    'Fill a value down',
    (b, s) => b.apply('0', [{ type: 'fillRange', sheet: s.name, source: s.lastCell, range: s.below }]),
  ],
  [
    'Copy and paste values',
    (b, s) => b.apply('0', [{ type: 'pasteValues', sheet: s.name, source: 'A1:B2', range: s.freeBlock }]),
  ],
  [
    'Add a web link',
    async (b, s) => {
      const current = await b.hyperlink('0', s.name, s.free);
      return b.apply('0', [
        {
          type: 'setHyperlink',
          sheet: s.name,
          cell: s.free,
          expectedHyperlinkRevision: current?.hyperlinkRevision ?? null,
          link: { kind: 'url', target: 'https://example.com/reference' },
        },
      ]);
    },
  ],
  [
    'Add a note to a cell',
    async (b, s) => {
      const current = await b.notes('0', s.name);
      return b.apply('0', [
        {
          type: 'setNote',
          sheet: s.name,
          cell: s.free,
          expectedNoteRevision: current?.noteRevision ?? null,
          note: { author: 'Reviewer', text: 'Please check' },
        },
      ]);
    },
  ],
  [
    'Highlight values above a number',
    (b, s) =>
      b.apply('0', [
        {
          sheet: s.name,
          operation: 'addConditionalRule',
          range: s.body,
          rule: 'greaterThan',
          threshold: 10,
          fillColor: 'FFC7CE',
        },
      ]),
  ],
];

async function facts(book, view) {
  // The view reports the grid size, not where the content ends; read the top
  // of a sheet and take the extent of what is actually there. Real workbooks
  // often open on a cover or notes sheet, so work on the visible sheet (of the
  // first six) with the most content: that is where a person would be working.
  let sheet = view.sheets[0],
    cells = [];
  for (const candidate of view.sheets.filter((item) => !item.hidden && item.state !== 'hidden').slice(0, 6)) {
    const page = await book.read('0', candidate.name, 'A1:AZ190').catch(() => ({ cells: {} }));
    const filled = Object.entries(page.cells ?? {}).filter(([, cell]) => cell.value != null || cell.formula);
    if (filled.length > cells.length * 1.5 || (!cells.length && filled.length)) {
      sheet = candidate;
      cells = filled;
    }
  }
  const position = (reference) => {
    const [, column, row] = /^([A-Z]+)(\d+)$/.exec(reference);
    return [Number(row), [...column].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0)];
  };
  // Stray notes to the right of a table are not part of it: take the block
  // that starts at column A and has no empty column inside.
  const used = new Set(cells.map(([reference]) => position(reference)[1]));
  let columns = 1;
  while (used.has(columns + 1)) columns += 1;
  columns = Math.max(2, columns);
  const rows = Math.max(
    3,
    ...cells.filter(([reference]) => position(reference)[1] <= columns).map(([reference]) => position(reference)[0])
  );
  // An empty column to work in: past everything the sheet uses, formatting
  // included, unless the sheet is formatted right to its last column.
  const grid = Number.isInteger(sheet.columns) && sheet.columns < 16000 ? sheet.columns : 0,
    freeIndex = Math.max(...used, columns, grid) + 1,
    last = letters(columns - 1),
    free = letters(freeIndex);
  // Real sheets often start with a title and notes. Work on the table a person
  // would: the first row that is mostly filled and is followed by filled rows
  // is the header, and the data runs until the rows stop being filled.
  const filledIn = new Map();
  for (const [reference] of cells) {
    const [row, column] = position(reference);
    if (column <= columns) filledIn.set(row, (filledIn.get(row) ?? 0) + 1);
  }
  const enough = Math.max(2, Math.ceil(columns * 0.6)),
    full = (row) => (filledIn.get(row) ?? 0) >= enough;
  let top = 1;
  for (let row = 1; row <= rows - 3; row++)
    if (full(row) && full(row + 1) && full(row + 2) && full(row + 3)) {
      top = row;
      break;
    }
  let end = top;
  while (end < rows && (filledIn.get(end + 1) ?? 0) >= 2) end += 1;
  end = Math.max(end, top + 2);
  const inBlock = ([reference]) => {
    const [row, column] = position(reference);
    return row > top && row <= end && column <= columns;
  };
  const number = cells.find(
    (entry) => inBlock(entry) && typeof entry[1].value === 'number' && !entry[1].formula && !entry[1].formulaType
  );
  const first = cells.find(
    ([reference, cell]) => /^A\d+$/.test(reference) && inBlock([reference]) && cell.value != null && !cell.formula
  );
  const fillColumn = letters(Math.min(2, columns - 1));
  return {
    name: sheet.name,
    tabled: (view.tables ?? []).some((table) => table.sheet === sheet.name),
    free: `${free}${top + 1}`,
    freePair: `${free}${top + 3}:${letters(freeIndex + 1)}${top + 3}`,
    freeBlock: `${free}${top + 5}:${letters(freeIndex + 1)}${top + 6}`,
    column: number ? number[0].replace(/\d+/g, '') : 'A',
    number: number ? { cell: number[0], value: number[1].value } : null,
    header: `A${top}:${last}${top}`,
    body: `B${top + 1}:${last}${end}`,
    // Sort the data rows only: people do not include a totals row in a sort.
    rows: `A${top + 1}:${last}${Math.max(top + 2, end - 1)}`,
    // Filter the range that is already filtered or tabled, as the screen would.
    table:
      sheet.filter?.range ?? view.tables?.find((table) => table.sheet === sheet.name)?.ref ?? `A${top}:${last}${end}`,
    // A data row in the middle of the table (rows are counted from zero here).
    middle: Math.min(top + 2, end - 1),
    firstValue: first ? first[1].value : 'x',
    lastCell: `${fillColumn}${Math.max(top + 1, end - 1)}`,
    below: `${fillColumn}${rows + 2}:${fillColumn}${rows + 4}`,
  };
}

const session = () =>
  new IndexedWorkbook(
    native,
    async (file, edits, out) => {
      const result = await editCalculatedWorkbookResult(await fs.readFile(file), edits);
      await fs.writeFile(out, result.bytes);
      return { revision: result.view.revision, edits: result.edits };
    },
    (input) => calculateIndexedWorkbook(input, native)
  );

/** How a workbook opens: editable, kept as stored, read only, or not at all. */
export async function opening(source, scratch) {
  const book = session();
  try {
    const view = await book.open(source, { storageDirectory: scratch });
    return {
      opens: true,
      readOnly: !!view.readOnly,
      keptAsStored: !!view.calculationPolicy?.preserveSource,
      reason: view.readOnly ? (view.passiveNotice ?? 'read only') : (view.calculationPolicy?.reason ?? null),
      sheets: view.sheets.length,
    };
  } catch (error) {
    return { opens: false, reason: String(error?.message ?? error) };
  } finally {
    await book.close().catch(() => {});
  }
}

/**
 * Open, type one number into an empty cell, save. Returns where the saved file
 * is and which cell was edited, so the caller can compare everything else with
 * the original.
 */
export async function roundTrip(source, scratch, output) {
  const book = session();
  try {
    const view = await book.open(source, { storageDirectory: scratch });
    if (view.readOnly) return { refused: view.passiveNotice ?? 'The workbook opened read-only.' };
    const sheet = await facts(book, view);
    const edited = await book.apply('0', [{ sheet: sheet.name, cell: sheet.free, value: 42 }]);
    await fs.copyFile((await book.export(state(edited))).path, output);
    return { sheet: sheet.name, cell: sheet.free };
  } catch (error) {
    return { refused: String(error?.message ?? error) };
  } finally {
    await book.close().catch(() => {});
  }
}

/**
 * How each action's result is checked in detail, given the facts about the
 * sheet it ran on. Either the saved file must equal the original except for
 * the cells the action was meant to touch (`ignore`) and the kinds of change it
 * was meant to make (`allow`); or, for actions that move things, the saved file
 * is compared with what Microsoft Excel produces for the same action (`excel`).
 */
export const checks = {
  'Type a number in an empty cell': (s) => ({ ignore: [s.free] }),
  'Type text in an empty cell': (s) => ({ ignore: [s.free] }),
  'Change an existing number': (s) => ({ ignore: [s.number.cell] }),
  'Enter a SUM formula': (s) => ({ ignore: [s.free] }),
  'Clear a cell': (s) => ({ ignore: [s.number.cell] }),
  'Make the header row bold': (s) => ({ ignore: [s.header] }),
  'Fill cells with a colour': (s) => ({ ignore: [s.header] }),
  'Show numbers as percent': (s) => ({ ignore: [s.body] }),
  'Centre and wrap text': (s) => ({ ignore: [s.header] }),
  'Widen a column': () => ({ allow: ['column widths'] }),
  'Make a row taller': () => ({ allow: ['row heights'] }),
  // Excel counts rows and columns from one; the editor's commands count from zero.
  'Insert a row in the data': (s) => ({ excel: ['insertRow', s.middle + 1] }),
  'Delete a row from the data': (s) => ({ excel: ['deleteRow', s.middle + 1] }),
  'Insert a column in the data': () => ({ excel: ['insertColumn', 2] }),
  'Delete a column from the data': () => ({ excel: ['deleteColumn', 2] }),
  // Inside a table a scripted sort in Excel takes the header row with it, which
  // a person sorting from the screen never gets; that case is left unchecked.
  'Sort by the first column': (s) =>
    s.tabled
      ? { skip: 'a table: Excel sorts differently by script' }
      : { excel: ['sort', s.rows, s.rows.split(':')[0], 'asc'] },
  'Filter by a value': () => ({ allow: ['filter', 'filter columns', 'filter buttons', 'hidden rows', 'names'] }),
  'Filter, then show all rows again': () => ({ allow: ['filter', 'filter columns', 'filter buttons', 'names'] }),
  'Filter, then switch the filter off': () => ({
    allow: ['filter', 'filter columns', 'filter buttons', 'sorted by', 'names'],
  }),
  'Merge two empty cells': (s) => ({ allow: ['merged'], ignore: [s.freePair] }),
  'Merge, then unmerge': (s) => ({ ignore: [s.freePair] }),
  'Freeze the top row': () => ({ allow: ['frozen'] }),
  'Add a worksheet': () => ({ allow: ['sheets'] }),
  'Rename the worksheet': () => ({ excel: ['rename', 'Renamed'] }),
  'Duplicate the worksheet': () => ({ excel: ['duplicate', 'Copy of sheet'] }),
  'Fill a value down': (s) => ({ ignore: [s.below] }),
  'Copy and paste values': (s) => ({ ignore: [s.freeBlock] }),
  'Add a web link': (s) => ({ ignore: [s.free] }),
  'Add a note to a cell': (s) => ({ ignore: [s.free] }),
  'Highlight values above a number': () => ({ allow: ['highlight rules'] }),
};

/** `keep`, when given, is where to leave the saved workbook for checking. */
export async function attempt(source, perform, scratch, keep) {
  const book = session();
  try {
    const view = await book.open(source, { storageDirectory: scratch });
    if (view.readOnly)
      return { outcome: 'refused', stage: 'open', message: view.passiveNotice ?? 'The workbook opened read-only.' };
    const sheet = await facts(book, view);
    let result;
    try {
      result = await perform(book, sheet);
    } catch (error) {
      return { outcome: 'refused', stage: 'action', message: String(error?.message ?? error), code: error?.code };
    }
    if (!result) return { outcome: 'skipped', message: 'The sheet has nothing suitable for this action.' };
    try {
      const saved = await book.export(state(result));
      readWorkbook(await fs.readFile(saved.path));
      if (keep) await fs.copyFile(saved.path, keep);
    } catch (error) {
      return { outcome: 'refused', stage: 'save', message: String(error?.message ?? error) };
    }
    return { outcome: 'ok', ...(keep ? { on: sheet } : {}) };
  } catch (error) {
    return { outcome: 'refused', stage: 'open', message: String(error?.message ?? error) };
  } finally {
    await book.close().catch(() => {});
  }
}

export { closeNativeWorkbookWorker };

if (process.argv[1] === new URL(import.meta.url).pathname) {
  pythonOnPath();
  const args = process.argv.slice(2),
    jsonAt = args.indexOf('--json'),
    report = jsonAt >= 0 ? args.splice(jsonAt, 2)[1] : null;
  const scratch = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-refusals-'));
  let inputs = args;
  if (!inputs.length) {
    execFileSync(
      'python3',
      [new URL('./ordinary-sheets.py', import.meta.url).pathname, path.join(scratch, 'ordinary')],
      {
        stdio: 'ignore',
      }
    );
    inputs = [path.join(scratch, 'ordinary')];
  }
  const files = [];
  for (const input of inputs) {
    const info = await fs.stat(input);
    if (info.isDirectory())
      for (const name of (await fs.readdir(input)).sort())
        if (/\.xls[xm]$/i.test(name) && !name.startsWith('~$')) files.push(path.join(input, name));
    if (info.isFile()) files.push(input);
  }
  const results = [];
  for (const file of files) {
    const copy = path.join(scratch, 'source-' + results.length + path.extname(file));
    await fs.copyFile(file, copy);
    const row = { file: path.basename(file), actions: [] };
    for (const [name, perform] of actions) {
      row.actions.push({ action: name, ...(await attempt(copy, perform, path.join(scratch, 'index'))) });
    }
    results.push(row);
    const refused = row.actions.filter((item) => item.outcome === 'refused').length;
    console.log(
      `${row.file.padEnd(30)} ${String(actions.length - refused).padStart(2)}/${actions.length} work, ${refused} refused`
    );
  }
  closeNativeWorkbookWorker();
  const total = results.length * actions.length,
    refusedAll = results.flatMap((row) =>
      row.actions.filter((item) => item.outcome === 'refused').map((item) => ({ file: row.file, ...item }))
    );
  console.log(
    `\n${total - refusedAll.length} of ${total} attempts worked; ${refusedAll.length} refused (${((refusedAll.length / total) * 100).toFixed(1)}%)`
  );
  const byAction = new Map();
  for (const item of refusedAll) byAction.set(item.action, [...(byAction.get(item.action) ?? []), item]);
  console.log('\nRefusals by action:');
  for (const [name, items] of [...byAction].sort((a, b) => b[1].length - a[1].length))
    console.log(
      `  ${String(items.length).padStart(2)} of ${results.length}  ${name}\n        e.g. ${items[0].file}: ${items[0].message.slice(0, 150)}`
    );
  if (report)
    await fs.writeFile(report, JSON.stringify({ files: results.length, actions: actions.length, results }, null, 1));
  await fs.rm(scratch, { recursive: true, force: true });
}
