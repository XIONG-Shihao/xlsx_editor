import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { zipSync, unzipSync, strToU8 } from 'fflate';
import { createWorkbook } from '../core/workbook-package.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
const descriptor = {
  type: 'importCells',
  id: '11111111-1111-4111-8111-111111111111',
  sha256: 'a'.repeat(64),
  count: 2002,
  bytes: 30000,
  shape: { sheet: 'Sheet1', start: 'B2', rows: 1001, columns: 2 },
  policy: 'grid',
};
async function fixture(t, context) {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-owner-import-'));
  t.after(() => fs.rm(dir, { recursive: true, force: true }));
  const parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>SUM(B2:C2)</f><v>999</v></c></row><row r="2"><c r="B2"><v>2</v></c><c r="C2"><v>3</v></c></row></sheetData></worksheet>'
  );
  const source = path.join(dir, 'source.xlsx');
  await fs.writeFile(source, zipSync(parts));
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('No compatibility writer'),
    (input) => calculateIndexedWorkbook(input, native),
    context
  );
  await book.open(source);
  t.after(() => book.close());
  return { book, source };
}
async function* pages() {
  for (let start = 0, sequence = 0; start < descriptor.count; start += 1000, sequence++)
    yield {
      sequence,
      cells: Array.from({ length: Math.min(1000, descriptor.count - start) }, (_, i) => [
        start + i,
        String(start + i + 1),
      ]),
    };
}
test('trusted paged paste publishes one state/descriptor and calculates once from a rectangular summary', async (t) => {
  const progress = [],
    f = await fixture(t, {
      importPages: (value) => {
        assert.deepEqual(value, descriptor);
        return pages();
      },
      importProgress: (event) => progress.push(event),
    }),
    before = f.book.diagnosticStats(),
    accepted = await f.book.apply('0', [descriptor]),
    state = accepted.view.storage.state;
  assert.deepEqual(accepted.commands, [descriptor]);
  assert.equal(JSON.stringify(accepted.edits).length < 1000, true);
  assert.equal(f.book.diagnosticStats().retainedStates, 2);
  assert.equal(f.book.diagnosticStats().calculationCalls - before.calculationCalls, 1);
  assert.deepEqual(
    progress.map((event) => event.completed),
    [1000, 2000, 2002]
  );
  assert.equal((await f.book.read(state, 'Sheet1', 'A1')).cells.A1.calculation.value, 3);
  assert.equal((await f.book.read('0', 'Sheet1', 'A1')).cells.A1.calculation.value, 5);
  assert.equal((await f.book.read(state, 'Sheet1', 'B1002:C1002')).cells.C1002.value, 2002);
  const saved = await f.book.export(state);
  assert.ok((await fs.stat(saved.path)).size > 1000);
});
test('failed private pages preserve every accepted old/redo state and missing/mixed descriptors refuse atomically', async (t) => {
  let fail = true;
  const f = await fixture(t, {
      importPages: () =>
        fail
          ? (async function* () {
              yield { sequence: 0, cells: Array.from({ length: 1000 }, (_, i) => [i, String(i + 1)]) };
              throw Error('Attachment interrupted');
            })()
          : pages(),
    }),
    first = await f.book.apply('0', [{ sheet: 'Sheet1', cell: 'B2', value: 9 }]),
    state = first.view.storage.state,
    redo = await f.book.apply(state, [{ sheet: 'Sheet1', cell: 'A2', value: 30 }]),
    later = redo.view.storage.state,
    before = f.book.diagnosticStats().retainedStates;
  await assert.rejects(f.book.apply(state, [descriptor]), /interrupted/);
  assert.equal(f.book.diagnosticStats().retainedStates, before);
  assert.equal((await f.book.read('0', 'Sheet1', 'B2')).cells.B2.value, 2);
  assert.equal((await f.book.read(state, 'Sheet1', 'B2')).cells.B2.value, 9);
  assert.equal((await f.book.read(later, 'Sheet1', 'A2')).cells.A2.value, 30);
  await assert.rejects(f.book.apply(state, [descriptor, { sheet: 'Sheet1', cell: 'A2', value: 99 }]), { code: /^(fill-bad-input|session-bad-input)$/ });
  fail = false;
  const next = await f.book.apply(state, [descriptor]);
  assert.equal((await f.book.read(next.view.storage.state, 'Sheet1', 'A1')).cells.A1.calculation.value, 3);
  const noResolver = await fixture(t, {});
  await assert.rejects(noResolver.book.apply('0', [descriptor]), { code: /^(query-bad-input|session-bad-input)$/ });
  assert.equal(noResolver.book.diagnosticStats().retainedStates, 1);
});
