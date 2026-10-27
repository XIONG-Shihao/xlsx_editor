/** Ordered direct-reference SUM unions; exact runtime evidence precedes enablement. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createHash } from 'node:crypto';
import { unzipSync, zipSync, strToU8 } from 'fflate';
import { createWorkbook, readWorkbook } from '../core/workbook-package.mjs';
import { calculateWorkbook } from '../core/calculation-node.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeFormula } from '../core/native-formula.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { checkXlsxPackage } from '../scripts/test/package-oracle.mjs';
after(() => closeNativeWorkbookWorker());

const view = (cells) => ({ revision: 'unions', sheets: [{ name: 'Data', rows: 8, columns: 8, cells }] });
const number = (value) => ({ value, dataType: 'number' });
async function fixture(t) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-unions-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  return root;
}

test('same-sheet SUM unions exactly match the independently captured Excel 16.113.3 oracle', async (t) => {
  const root = await fixture(t),
    oracle = JSON.parse(await fs.readFile(new URL('../fixtures/excel-union-reference.json', import.meta.url), 'utf8'));
  const sourceBytes = await fs.readFile(new URL('../fixtures/excel-union-source.xlsx', import.meta.url));
  const oracleBytes = await fs.readFile(new URL('../fixtures/excel-union-reference.xlsx', import.meta.url));
  const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
  assert.equal(hash(sourceBytes), oracle.sourceSha256);
  assert.equal(hash(oracleBytes), oracle.oracleWorkbookSha256);
  const source = readWorkbook(sourceBytes);
  for (const item of oracle.cases) {
    const input = {
      ...source,
      sheets: source.sheets.map((sheet) => ({
        ...sheet,
        cells: Object.fromEntries(
          Object.entries(sheet.cells).filter(
            ([ref, cell]) => !cell.formula || (sheet.name === oracle.sheet && ref === item.ref)
          )
        ),
      })),
    };
    const result = await nativeFormula(native, root, 'evaluate', { view: input });
    if (item.dataType === 'number') {
      assert.equal(result.supported, true, `${item.ref}: ${result.reason}`);
      assert.equal(result.results[0].value, item.expected, `${item.ref}: ${item.formula}`);
    } else {
      // Native error inputs and the captured cross-sheet #VALUE case remain
      // explicit misses. The portable owner must not guess a fresh error cache.
      assert.equal(result.supported, false, item.ref);
      assert.equal(result.results, undefined, item.ref);
    }
  }
});

test('direct-reference SUM union dependency analysis preserves area order and duplicates', async (t) => {
  const root = await fixture(t),
    input = view({ A1: number(2), A2: number(3), A3: number(5), C1: { formula: 'SUM((A3,A1:A2,A1),A2)' } });
  const analyzed = await nativeFormula(native, root, 'analyze', { view: input });
  assert.equal(analyzed.supported, true, analyzed.reason);
  assert.deepEqual(analyzed.items[0].ranges, [
    { sheet: 'Data', ref: 'A3' },
    { sheet: 'Data', startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 },
    { sheet: 'Data', ref: 'A1' },
    { sheet: 'Data', ref: 'A2' },
  ]);
  const evaluated = await nativeFormula(native, root, 'evaluate', { view: input });
  assert.equal(evaluated.supported, true, evaluated.reason);
  const equivalent = await calculateWorkbook(
    view({ ...input.sheets[0].cells, C1: { formula: 'SUM(A3,A1:A2,A1,A2)' } })
  );
  assert.equal(evaluated.results[0].value, equivalent.sheets[0].cells.C1.calculation.value);
});

test('union copies and structural rewrites preserve grouping, anchors and quoted sheet spelling', async (t) => {
  const root = await fixture(t),
    formula = '=SUM(( $A$1:A2 , B$3 , C4,C4),\'Other Sheet\'!D5)+("A1"="A1")';
  const copied = await nativeFormula(native, root, 'rewrite', { items: [{ formula, copyRows: 2, copyColumns: 1 }] });
  assert.equal(copied.supported, true, copied.reason);
  assert.equal(copied.formulas[0], '=SUM(( $A$1:B4 , C$3 , D6,D6),\'Other Sheet\'!E7)+("A1"="A1")');
  const inserted = await nativeFormula(native, root, 'rewrite', {
    target: 'Data',
    index: 1,
    count: 1,
    items: [{ formula: 'SUM((A1:A2,A2,A4))', owner: 'Data' }],
  });
  assert.equal(inserted.supported, true, inserted.reason);
  assert.equal(inserted.formulas[0], 'SUM((A1:A3,A3,A5))');
  const removed = await nativeFormula(native, root, 'rewrite', {
    target: 'Data',
    index: 1,
    count: 1,
    remove: true,
    items: [{ formula: 'SUM((A1:A2,A2,A4))', owner: 'Data' }],
  });
  assert.equal(removed.supported, true, removed.reason);
  assert.equal(removed.formulas[0], 'SUM((A1:A1,#REF!,A3))');
  const renamed = await nativeFormula(native, root, 'rename', {
    old: 'Other Sheet',
    new: "O'Brien",
    items: [{ formula }],
  });
  assert.equal(renamed.supported, true, renamed.reason);
  assert.equal(renamed.formulas[0], formula.replace("'Other Sheet'!", "'O''Brien'!"));
});

test('union unsupported contexts, malformed areas and cancellation never return partial rewrites', async (t) => {
  const root = await fixture(t);
  for (const formula of [
    '(A1,A2)+1',
    'SUM((A1,A2*2))',
    'AVERAGE((A1,A2))',
    'ABS((A1,A2))',
    'SUM((A1,Rate))',
    'SUM((A1,Table1[Cost]))',
    'SUM((A1,{1,2}))',
    'SUM((A1,))',
    "SUM((A1,'Other Sheet'!B2))",
    "SUM(('Data'!A1,'Data'!A2))",
    "SUM(('[external.xlsx]Data'!A1,A2))",
    "SUM(('Data:Other'!A1,A2))",
  ]) {
    const rewritten = await nativeFormula(native, root, 'rewrite', {
      items: [
        { formula: 'SUM((A1,A2))', copyRows: 1 },
        { formula, copyRows: 1 },
      ],
    });
    assert.equal(rewritten.supported, false, formula);
    assert.equal(rewritten.formulas, undefined, formula);
  }
  const boundary = await nativeFormula(native, root, 'rewrite', {
    items: [{ formula: 'SUM((A1048576,A1))', copyRows: 1 }],
  });
  assert.equal(boundary.supported, true);
  assert.deepEqual(boundary.formulas, ['SUM((#REF!,A2))']);
  await fs.writeFile(path.join(root, 'cancelled'), '1');
  await assert.rejects(
    nativeFormula(native, root, 'rewrite', { items: [{ formula: 'SUM((A1,A2))', copyRows: 1 }] }),
    { code: 'operation-cancelled' }
  );
  await fs.rm(path.join(root, 'cancelled'));
  assert.equal(
    (await nativeFormula(native, root, 'rewrite', { items: [{ formula: 'SUM((A1,A2))', copyRows: 1 }] })).supported,
    true
  );
});

test('indexed union fill, input edits, branches, insertion/deletion and export/reopen stay revision qualified', async (t) => {
  const root = await fixture(t),
    source = path.join(root, 'source.xlsx'),
    parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
      '<row r="1"><c r="A1"><v>2</v></c><c r="C1"><f>SUM((A1,A2))</f><v>999</v></c></row>' +
      '<row r="2"><c r="A2"><v>3</v></c></row><row r="3"><c r="A3"><v>5</v></c></row></sheetData></worksheet>'
  );
  parts['custom/opaque.bin'] = new Uint8Array([0, 1, 2, 255]);
  const original = Buffer.from(zipSync(parts));
  await fs.writeFile(source, original);
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('No whole-workbook writer'),
    (input) => calculateIndexedWorkbook(input, native)
  );
  await book.open(source);
  t.after(() => book.close());
  const read = (state, range = 'A1:C5') => book.read(state, 'Sheet1', range);
  assert.equal((await read('0')).cells.C1.calculation.value, 5);
  const filled = await book.apply('0', [{ type: 'fillRange', sheet: 'Sheet1', source: 'C1', range: 'C2:C4' }]);
  const filledState = filled.view.storage.state,
    cells = (await read(filledState)).cells;
  assert.equal(cells.C2.formula, 'SUM((A2,A3))');
  assert.equal(cells.C2.calculation.value, 8);
  assert.equal(cells.C3.calculation.value, 5);
  assert.equal(cells.C4.calculation.value, 0);
  const replayed = await book.apply('0', filled.edits);
  assert.equal((await read(replayed.view.storage.state)).cells.C2.formula, cells.C2.formula);
  assert.equal((await read(replayed.view.storage.state)).cells.C2.calculation.value, 8);
  const edited = await book.apply(filledState, [{ sheet: 'Sheet1', cell: 'A1', value: 7 }]);
  assert.equal((await read(edited.view.storage.state)).cells.C1.calculation.value, 10);
  assert.equal((await read(filledState)).cells.C1.calculation.value, 5);
  const inserted = await book.apply(filledState, [{ sheet: 'Sheet1', operation: 'insertRow', index: 1 }]);
  const moved = (await read(inserted.view.storage.state)).cells;
  assert.equal(moved.C1.formula, 'SUM((A1,A3))');
  assert.equal(moved.C1.calculation.value, 5);
  assert.equal(moved.C3.formula, 'SUM((A3,A4))');
  assert.equal(moved.C3.calculation.value, 8);
  const restored = await book.apply(inserted.view.storage.state, [
    { sheet: 'Sheet1', operation: 'deleteRow', index: 1 },
  ]);
  assert.equal((await read(restored.view.storage.state)).cells.C2.formula, 'SUM((A2,A3))');
  const output = (await book.export(edited.view.storage.state)).path;
  assert.equal(
    checkXlsxPackage(source, output, {
      targets: { Sheet1: ['A1', 'C2', 'C3', 'C4'] },
      expected: {
        'Sheet1!A1': { value: 7 },
        'Sheet1!C1': { formula: 'SUM((A1,A2))', calculated: 10 },
        'Sheet1!C2': { formula: 'SUM((A2,A3))', calculated: 8 },
      },
    }).passed,
    true
  );
  const reopened = new IndexedWorkbook(
    native,
    () => assert.fail('No writer'),
    (input) => calculateIndexedWorkbook(input, native)
  );
  await reopened.open(output);
  t.after(() => reopened.close());
  assert.equal((await reopened.read('0', 'Sheet1', 'C1:C2')).cells.C1.calculation.value, 10);
  assert.deepEqual(await fs.readFile(source), original);
});

test('qualified shared union masters and followers recalculate and normalize without changing old branches', async (t) => {
  const root = await fixture(t),
    source = path.join(root, 'shared.xlsx'),
    parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
      '<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f t="shared" si="0" ref="B1:B2">SUM((A1,A2))</f><v>999</v></c></row>' +
      '<row r="2"><c r="A2"><v>3</v></c><c r="B2"><f t="shared" si="0"/><v>999</v></c></row>' +
      '<row r="3"><c r="A3"><v>5</v></c></row></sheetData></worksheet>'
  );
  const original = Buffer.from(zipSync(parts));
  await fs.writeFile(source, original);
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('No writer'),
    (input) => calculateIndexedWorkbook(input, native)
  );
  await book.open(source);
  t.after(() => book.close());
  const read = (state) => book.read(state, 'Sheet1', 'A1:B3');
  const initial = (await read('0')).cells;
  assert.equal(initial.B1.calculation.value, 5);
  assert.equal(initial.B2.calculation.value, 8);
  const changed = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 7 }]);
  const current = (await read(changed.view.storage.state)).cells;
  assert.equal(current.B1.calculation.value, 10);
  assert.equal(current.B2.calculation.value, 8);
  const normalized = await book.apply(changed.view.storage.state, [{ sheet: 'Sheet1', cell: 'B1', value: 42 }]);
  const expanded = (await read(normalized.view.storage.state)).cells;
  assert.equal(expanded.B1.value, 42);
  assert.equal(expanded.B2.formula, 'SUM((A2,A3))');
  assert.equal(expanded.B2.formulaType, undefined);
  assert.equal(expanded.B2.calculation.value, 8);
  assert.equal((await read('0')).cells.B1.calculation.value, 5);
  assert.equal((await read(changed.view.storage.state)).cells.B1.calculation.value, 10);
  const output = (await book.export(normalized.view.storage.state)).path;
  assert.equal(readWorkbook(await fs.readFile(output)).sheets[0].cells.B2.formula, 'SUM((A2,A3))');
  assert.deepEqual(await fs.readFile(source), original);
});

test('native union results feed qualified compatibility dependents without portable union evaluation', async (t) => {
  const root = await fixture(t),
    source = path.join(root, 'mixed.xlsx'),
    parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
      '<row r="1"><c r="A1"><v>2</v></c><c r="C1"><f>SUM((A1,A2))</f><v>999</v></c>' +
      '<c r="D1"><f>VLOOKUP(C1,E1:F2,2,FALSE)</f><v>999</v></c><c r="E1"><v>5</v></c><c r="F1"><v>9</v></c></row>' +
      '<row r="2"><c r="A2"><v>3</v></c><c r="E2"><v>10</v></c><c r="F2"><v>11</v></c></row></sheetData></worksheet>'
  );
  await fs.writeFile(source, zipSync(parts));
  let diagnostics;
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('No writer'),
    async (input) => (diagnostics = await calculateIndexedWorkbook(input, native))
  );
  await book.open(source);
  t.after(() => book.close());
  assert(diagnostics.nativeSqlite);
  assert.equal(diagnostics.inputCells, 0);
  assert.equal((await book.read('0', 'Sheet1', 'C1:D1')).cells.D1.calculation.value, 9);
  const changed = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 7 }]);
  const cells = (await book.read(changed.view.storage.state, 'Sheet1', 'C1:D1')).cells;
  assert.equal(cells.C1.calculation.value, 10);
  assert.equal(cells.D1.calculation.value, 11);
  assert.equal((await book.read('0', 'Sheet1', 'C1:D1')).cells.D1.calculation.value, 9);
  const saved = readWorkbook(await fs.readFile((await book.export(changed.view.storage.state)).path));
  assert.equal(saved.sheets[0].cells.C1.cachedValue, 10);
  assert.equal(saved.sheets[0].cells.D1.cachedValue, 11);
});
