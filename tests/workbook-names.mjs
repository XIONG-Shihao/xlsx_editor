import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { createWorkbook, readWorkbook } from '../core/workbook-package.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { indexedNamePlan } from '../core/indexed-names.mjs';
import { checkXlsxPackage } from '../scripts/test/package-oracle.mjs';
after(() => closeNativeWorkbookWorker());
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
async function fixture(t, { formula = false } = {}) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-names-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  const parts = unzipSync(createWorkbook());
  parts['custom/opaque.bin'] = Uint8Array.from([0, 1, 255, 2]);
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${ns}"><dimension ref="A1:${formula ? 'C' : 'B'}1"/><sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1"><v>3</v></c>${formula ? '<c r="C1"><f>SUM(Amount)</f><v>999</v></c>' : ''}</row></sheetData></worksheet>`
  );
  const source = path.join(root, 'source.xlsx'),
    original = Buffer.from(zipSync(parts));
  await fs.writeFile(source, original);
  const make = () =>
    new IndexedWorkbook(
      native,
      () => assert.fail('No compatibility writer'),
      (input) => calculateIndexedWorkbook(input, native)
    );
  const book = make();
  await book.open(source);
  t.after(() => book.close());
  return { root, source, original, book, make };
}
test('named-range creation rebinds calculation, preserves old state and independent package metadata', async (t) => {
  const { book, source, original, make } = await fixture(t, { formula: true });
  assert.equal((await book.read('0', 'Sheet1', 'C1')).cells.C1.calculation.status, 'unresolved');
  const command = {
    operation: 'defineName',
    sheet: 'Sheet1',
    name: 'Amount',
    range: 'A1:B1',
    nameScope: 'workbook',
    expectedNameFormula: null,
  };
  const next = await book.apply('0', [command]),
    state = next.view.storage.state;
  assert.deepEqual(next.view.namedRanges, [{ name: 'Amount', formula: "'Sheet1'!$A$1:$B$1" }]);
  assert.equal((await book.read(state, 'Sheet1', 'C1')).cells.C1.calculation.value, 5);
  assert.equal((await book.read('0', 'Sheet1', 'C1')).cells.C1.calculation.status, 'unresolved');
  const branch = await book.apply(state, [{ sheet: 'Sheet1', cell: 'A1', value: 7 }]);
  assert.equal((await book.read(branch.view.storage.state, 'Sheet1', 'C1')).cells.C1.calculation.value, 10);
  assert.equal((await book.read(state, 'Sheet1', 'C1')).cells.C1.calculation.value, 5);
  const saved = (await book.export(state)).path,
    bytes = await fs.readFile(saved),
    parts = unzipSync(bytes);
  assert.match(strFromU8(parts['xl/workbook.xml']), /<definedName name="Amount">'Sheet1'!\$A\$1:\$B\$1<\/definedName>/);
  assert.deepEqual(parts['custom/opaque.bin'], unzipSync(original)['custom/opaque.bin']);
  assert.equal(
    checkXlsxPackage(source, saved, {
      expectedNames: { Amount: "'Sheet1'!$A$1:$B$1" },
      expected: { 'Sheet1!C1': { formula: 'SUM(Amount)', value: 5 } },
    }).passed,
    true
  );
  const reopened = make();
  await reopened.open(saved);
  try {
    assert.equal((await reopened.read('0', 'Sheet1', 'C1')).cells.C1.calculation.value, 5);
  } finally {
    await reopened.close();
  }
  await assert.rejects(
    book.apply(state, [
      { operation: 'removeName', sheet: 'Sheet1', name: 'Amount', expectedNameFormula: "'Sheet1'!$A$1:$B$1" },
    ]),
    { code: 'name-unsupported' }
  );
  assert.deepEqual(await fs.readFile(source), original);
});
test('isolated scoped names support guarded removal, case collision refusal and cold reopen', async (t) => {
  const { book, source, original, make } = await fixture(t);
  const next = await book.apply('0', [
      {
        operation: 'defineName',
        sheet: 'Sheet1',
        name: 'Inputs',
        range: 'A1:A1048576',
        nameScope: 'worksheet',
        expectedNameFormula: null,
      },
    ]),
    state = next.view.storage.state;
  assert.deepEqual(next.view.namedRanges, [{ name: 'Inputs', formula: "'Sheet1'!$A$1:$A$1048576", sheet: 'Sheet1' }]);
  await assert.rejects(
    book.apply(state, [
      {
        operation: 'defineName',
        sheet: 'Sheet1',
        name: 'inputs',
        range: 'B1',
        nameScope: 'worksheet',
        expectedNameFormula: null,
      },
    ]),
    /already exists/
  );
  await assert.rejects(
    book.apply(state, [
      { operation: 'removeName', sheet: 'Sheet1', name: 'Inputs', nameScope: 'worksheet', expectedNameFormula: 'A1' },
    ]),
    /Conflict/
  );
  const saved = (await book.export(state)).path;
  const reopened = make();
  await reopened.open(saved);
  try {
    assert.deepEqual(reopened.view('0').namedRanges, next.view.namedRanges);
  } finally {
    await reopened.close();
  }
  const removed = await book.apply(state, [
    {
      operation: 'removeName',
      sheet: 'Sheet1',
      name: 'Inputs',
      nameScope: 'worksheet',
      expectedNameFormula: "'Sheet1'!$A$1:$A$1048576",
    },
  ]);
  assert.deepEqual(removed.view.namedRanges, []);
  assert.deepEqual(book.view(state).namedRanges, next.view.namedRanges);
  assert.equal(
    readWorkbook(await fs.readFile((await book.export(removed.view.storage.state)).path)).namedRanges.length,
    0
  );
  assert.deepEqual(await fs.readFile(source), original);
});
test('unsafe names, absent guards and mixed invalid batches never publish accepted state', async (t) => {
  const { book, source, original } = await fixture(t);
  const view = book.view('0'),
    bytes = original;
  const base = { operation: 'defineName', sheet: 'Sheet1', name: 'Safe', range: 'A1', expectedNameFormula: null };
  for (const name of ['A1', 'R1C1', 'C', '_xlnm.Print_Area', 'has space', 'Ω', '9abc', 'a'.repeat(256)])
    await assert.rejects(book.apply('0', [{ ...base, name }]), /name|addresses/);
  await assert.rejects(book.apply('0', [{ ...base, expectedNameFormula: undefined }]), { code: 'name-bad-input' });
  await assert.rejects(book.apply('0', [base, { sheet: 'Sheet1', cell: 'XFE1', value: 5 }]), /limits|address/);
  assert.deepEqual(book.view('0').namedRanges, []);
  assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.value, 2);
  assert.throws(() => indexedNamePlan(bytes, view, { ...base, range: 'B2:A1' }), /range/);
  assert.deepEqual(await fs.readFile(source), original);
});
