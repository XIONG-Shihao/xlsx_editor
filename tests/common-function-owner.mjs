/** Actual native preview/IndexedWorkbook results and saved facts, separate from UI proof. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { DOMParser } from '@xmldom/xmldom';
import { unzipSync, strFromU8 } from 'fflate';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { serializeFunction } from '../ui/functionCatalog.ts';
import { commonFunctionCases as cases, commonFunctionFixture } from './support/common-function-fixture.mjs';
after(() => closeNativeWorkbookWorker());
async function fixture(t) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'common-functions-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const source = path.join(directory, 'source.xlsx'),
    bytes = commonFunctionFixture();
  await fs.writeFile(source, bytes);
  const original = await fs.readFile(source),
    book = new IndexedWorkbook(
      native,
      () => assert.fail('No compatibility writer'),
      (input) => calculateIndexedWorkbook(input, native)
    );
  await book.open(source);
  t.after(() => book.close());
  return { directory, source, original, book };
}
function check(cell, c) {
  assert.equal(cell.formula.replace(/^_xlfn\./i, ''), c.formula, c.cell);
  if (c.value?.error) {
    assert.equal(cell.calculation.status, 'error', JSON.stringify(cell));
    assert.equal(cell.calculation.error, c.value.error);
  } else {
    assert.equal(cell.calculation.status, 'calculated', JSON.stringify(cell));
    assert.equal(cell.calculation.value, c.value, c.cell);
    assert.equal(typeof cell.calculation.value, typeof c.value);
  }
}
function saved(bytes) {
  const parts = unzipSync(bytes),
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    doc = new DOMParser().parseFromString(strFromU8(parts['xl/worksheets/sheet1.xml']), 'application/xml'),
    cells = Array.from(doc.getElementsByTagNameNS(ns, 'c'));
  for (const c of cases) {
    const node = cells.find((n) => n.getAttribute('r') === c.cell),
      formula = node.getElementsByTagNameNS(ns, 'f')[0].textContent,
      value = node.getElementsByTagNameNS(ns, 'v')[0]?.textContent ?? '',
      kind = node.getAttribute('t') || 'n';
    assert.equal(formula.replace(/^_xlfn\./i, ''), c.formula);
    if (c.value?.error) {
      assert.equal(kind, 'e');
      assert.equal(value, c.value.error);
    } else if (typeof c.value === 'number') {
      assert.equal(kind, 'n');
      assert.equal(Number(value), c.value);
    } else if (typeof c.value === 'boolean') {
      assert.equal(kind, 'b');
      assert.equal(value, c.value ? '1' : '0');
    } else {
      assert.equal(kind, 'str');
      assert.equal(value, c.value);
    }
  }
  assert.deepEqual(parts['custom/common-functions.bin'], Uint8Array.from([0, 18, 255, 9]));
}
const commands = () => [
  {
    type: 'setCells',
    sheet: 'Sheet1',
    cells: cases.map((c) => ({ cell: c.cell, formula: serializeFunction(c.id, c.args).formula })),
  },
];
test('all new guided profiles have private typed previews without accepting cells or changing history', async (t) => {
  const f = await fixture(t),
    b = f.book,
    before = b.diagnosticStats().retainedStates;
  for (const c of cases) {
    assert.equal(serializeFunction(c.id, c.args).formula, c.formula);
    const preview = await b.previewFormula('0', 'Sheet1', c.cell, c.formula, { requestId: randomUUID() });
    assert.equal(preview.authority, 'draft-preview');
    assert.equal(preview.result.status, 'value', c.formula + ' ' + JSON.stringify(preview));
    assert.equal(preview.result.type, c.value?.error ? 'error' : typeof c.value === 'string' ? 'text' : typeof c.value);
    assert.equal(preview.result.value, c.value?.error ?? c.value, c.formula);
  }
  assert.equal(b.diagnosticStats().retainedStates, before);
  assert.deepEqual((await b.read('0', 'Sheet1', 'H1:H20')).cells, {});
  assert.deepEqual(await fs.readFile(f.source), f.original);
});
test('Auto accepts twenty independent typed results, old-state purity, saved caches and cold reopen', async (t) => {
  const f = await fixture(t),
    b = f.book,
    made = await b.apply('0', commands()),
    state = made.view.storage.state,
    read = await b.read(state, 'Sheet1', 'H1:H20');
  for (const c of cases) check(read.cells[c.cell], c);
  assert.deepEqual((await b.read('0', 'Sheet1', 'H1:H20')).cells, {});
  const output = await b.export(state);
  saved(await fs.readFile(output.path));
  await fs.copyFile(output.path, path.join(f.directory, 'saved.xlsx'));
  const cold = new IndexedWorkbook(
    native,
    () => assert.fail('No compatibility writer'),
    (input) => calculateIndexedWorkbook(input, native)
  );
  t.after(() => cold.close());
  await cold.open(output.path);
  for (const c of cases) check((await cold.read('0', 'Sheet1', c.cell)).cells[c.cell], c);
  assert.deepEqual(await fs.readFile(f.source), f.original);
});
test('Manual previews remain private, insertion stays pending, and explicit calculation qualifies the same saved results', async (t) => {
  const f = await fixture(t),
    b = f.book,
    manual = await b.apply('0', [{ type: 'setCalculationOptions', mode: 'manual', recalculateBeforeSave: false }]),
    state = manual.view.storage.state;
  const preview = await b.previewFormula(state, 'Sheet1', 'H8', cases[7].formula, { requestId: randomUUID() });
  assert.equal(preview.result.value, 1.25);
  assert.equal(b.view(state).calculationPolicy.mode, 'manual');
  const inserted = await b.apply(state, commands()),
    pending = inserted.view.storage.state;
  for (const c of cases) {
    const cell = (await b.read(pending, 'Sheet1', c.cell)).cells[c.cell];
    assert.notEqual(cell.calculation.status, 'calculated');
    assert.equal(cell.calculation.value, undefined);
  }
  const calculated = await b.apply(pending, [{ type: 'recalculate' }]);
  for (const c of cases) check((await b.read(calculated.view.storage.state, 'Sheet1', c.cell)).cells[c.cell], c);
  saved(await fs.readFile((await b.export(calculated.view.storage.state)).path));
  assert.notEqual((await b.read(pending, 'Sheet1', 'H8')).cells.H8.calculation.status, 'calculated');
});
