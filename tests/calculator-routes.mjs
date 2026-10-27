// The calculation routes must agree with Excel and with each other.
//
// A formula is calculated either by the general route alone (the Univer engine
// with this project's Excel-compatibility patches) or by the product route (the
// Rust fast path first, the general route for the rest). The reference
// workbooks were calculated and saved by Microsoft Excel, so the result Excel
// stored for each formula is the answer. A route may decline a formula; it may
// never give a different answer.
import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { readWorkbook } from '../core/workbook-package.mjs';
import * as general from '../core/calculation-node.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { cellPosition, columnName } from '../core/workbook-commands.mjs';

const fixtures = new URL('../fixtures/', import.meta.url).pathname,
  scratch = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-routes-'));
after(async () => {
  closeNativeWorkbookWorker();
  await fs.rm(scratch, { recursive: true, force: true });
});

// Workbooks Excel calculated and saved. Each has formulas with stored results.
const references = (await fs.readdir(fixtures)).filter((name) => /^excel-.*\.xlsx$/.test(name)).sort();
// Excel saves a blocked spill (#SPILL!) as #VALUE! plus a marker that older
// programs ignore, so a stored #VALUE! also matches a calculated #SPILL!.
const same = (got, want) =>
  typeof got === 'number' && typeof want === 'number'
    ? Math.abs(got - want) <= 1e-9 * Math.max(1, Math.abs(want))
    : String(got) === String(want) || (want === '#VALUE!' && got === '#SPILL!');
const answer = (calculation) =>
  calculation?.status === 'calculated'
    ? { value: calculation.value }
    : calculation?.status === 'error'
      ? { value: calculation.error }
      : null;

/** Excel's stored result for every formula, by "Sheet!A1". */
function stored(view) {
  const cells = new Map();
  for (const sheet of view.sheets)
    for (const [reference, cell] of Object.entries(sheet.cells))
      if (cell.formula && cell.cachedValue !== undefined && cell.cachedValue !== null)
        cells.set(`${sheet.name}!${reference}`, cell.cachedValue);
  return cells;
}
async function generalRoute(view) {
  const answers = new Map();
  let calculated;
  try {
    calculated = await general.calculateWorkbook(view);
  } catch {
    return answers; // The whole workbook was declined.
  }
  for (const sheet of calculated.sheets)
    for (const [reference, cell] of Object.entries(sheet.cells)) {
      const result = cell.formula ? answer(cell.calculation) : null;
      if (result) answers.set(`${sheet.name}!${reference}`, result.value);
    }
  return answers;
}
async function productRoute(file, view, wanted) {
  const answers = new Map(),
    directory = await fs.mkdtemp(path.join(scratch, 'book-')),
    copy = path.join(directory, 'source.xlsx');
  await fs.copyFile(file, copy);
  const book = new IndexedWorkbook(
    native,
    async () => {
      throw Error('the general writer is not part of this test');
    },
    (input) => calculateIndexedWorkbook(input, native)
  );
  try {
    const opened = await book.open(copy, { storageDirectory: path.join(directory, 'index') });
    if (opened.calculationPolicy?.canRecalculate === false) return answers; // Kept as stored; nothing is calculated.
    let state = '0';
    try {
      state = (await book.apply('0', [{ type: 'recalculate' }])).view.storage.state;
    } catch {
      return answers; // The whole workbook was declined.
    }
    for (const sheet of view.sheets) {
      const cells = [...wanted.keys()]
        .filter((key) => key.startsWith(sheet.name + '!'))
        .map((key) => cellPosition(key.slice(sheet.name.length + 1)));
      if (!cells.length) continue;
      const lastColumn = Math.max(...cells.map((cell) => cell.col)),
        lastRow = Math.max(...cells.map((cell) => cell.row)),
        step = Math.max(1, Math.floor(5000 / (lastColumn + 1)));
      for (let row = 0; row <= lastRow; row += step) {
        const range = `A${row + 1}:${columnName(lastColumn)}${Math.min(lastRow, row + step - 1) + 1}`,
          page = await book.read(state, sheet.name, range);
        for (const [reference, cell] of Object.entries(page.cells)) {
          const result = cell.formula ? answer(cell.calculation) : null;
          if (result) answers.set(`${sheet.name}!${reference}`, result.value);
        }
      }
    }
    return answers;
  } finally {
    await book.close();
  }
}

const totals = { formulas: 0, general: 0, product: 0, both: 0 };
for (const name of references)
  test(`${name}: no route gives a different answer from Excel, and the routes agree`, async () => {
    const file = path.join(fixtures, name);
    let view;
    try {
      view = readWorkbook(await fs.readFile(file));
    } catch {
      return; // A workbook the reader declines has nothing to compare.
    }
    const wanted = stored(view);
    if (!wanted.size) return;
    const fromGeneral = await generalRoute(view),
      fromProduct = await productRoute(file, view, wanted),
      wrong = [];
    for (const [cell, excel] of wanted) {
      const a = fromGeneral.get(cell),
        b = fromProduct.get(cell);
      if (a !== undefined && !same(a, excel)) wrong.push(`${cell}: Excel ${excel}, general route ${a}`);
      if (b !== undefined && !same(b, excel)) wrong.push(`${cell}: Excel ${excel}, product route ${b}`);
      if (a !== undefined && b !== undefined && !same(a, b)) wrong.push(`${cell}: routes disagree, ${a} and ${b}`);
      totals.formulas += 1;
      totals.general += a !== undefined ? 1 : 0;
      totals.product += b !== undefined ? 1 : 0;
      totals.both += a !== undefined && b !== undefined ? 1 : 0;
    }
    assert.deepEqual(wrong.slice(0, 12), [], `${wrong.length} wrong answers in ${name}`);
  });

test('the comparison actually compared something', () => {
  // Guards against the test passing because every route declined everything.
  console.log(
    `calculator routes: ${totals.formulas} Excel results; general route answered ${totals.general}, ` +
      `product route ${totals.product}, both ${totals.both}`
  );
  assert(totals.formulas > 500, 'too few Excel reference results were found');
  assert(totals.general > 400, 'the general route answered too few formulas');
  assert(totals.product > 300, 'the product route answered too few formulas');
});
