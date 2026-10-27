import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, realpath, writeFile, rm } from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { SpreadsheetWorker } from '../adapters/worker-client.mjs';
import { workbook } from '../fixtures/standalone.mjs';

test('readonly owner filterValues carries typed bounded pages and validates inputs', async (t) => {
  const root = await realpath(await mkdtemp(path.join(os.tmpdir(), 'xlsx-filter-transport-'))),
    filename = path.join(root, 'book.xlsx'),
    recoveryDirectory = path.join(root, 'recovery');
  await writeFile(filename, workbook());
  const owner = new SpreadsheetWorker({ filename, recoveryDirectory });
  t.after(async () => {
    await owner.close();
    await rm(root, { recursive: true, force: true });
  });
  const status = await owner.open(),
    request = {
      operation: 'filterValues',
      state: status.state,
      sheet: 'Sheet1',
      range: 'A1:A2',
      column: 'A',
      offset: 0,
      limit: 1,
    };
  const first = await owner.request(request);
  assert.equal(first.values.length, 1);
  assert.equal(first.values[0].value, 10);
  assert.equal(first.values[0].count, 1);
  assert.equal(first.total, 2);
  assert.equal(first.nextOffset, 1);
  const next = await owner.request({ ...request, offset: first.nextOffset });
  assert.equal(next.values[0].value, null);
  assert.equal(next.nextOffset, null);
  assert.equal((await owner.request({ operation: 'status' })).state, status.state);
  for (const change of [
    { sheet: 'missing' },
    { column: 'B' },
    { column: 'A1' },
    { range: 'A0:A2' },
    { offset: -1 },
    { offset: Number.MAX_SAFE_INTEGER + 1 },
    { limit: 0 },
    { limit: 101 },
    { query: '🙂'.repeat(8193) },
  ])
    await assert.rejects(owner.request({ ...request, ...change }));
});
test('CSV owners reject the XLSX-only readonly filter value transport', async (t) => {
  const root = await realpath(await mkdtemp(path.join(os.tmpdir(), 'csv-filter-transport-'))),
    filename = path.join(root, 'book.csv');
  await writeFile(filename, 'a\nb\n');
  const owner = new SpreadsheetWorker({ filename, recoveryDirectory: path.join(root, 'recovery') });
  t.after(async () => {
    await owner.close();
    await rm(root, { recursive: true, force: true });
  });
  const status = await owner.open();
  await assert.rejects(
    owner.request({ operation: 'filterValues', state: status.state, sheet: 'CSV', range: 'A1:A2', column: 'A' }),
    /XLSX/
  );
  assert.equal(
    (await owner.request({ operation: 'read', state: status.state, sheet: 'CSV', range: 'A1' })).cells.A1.value,
    'a'
  );
});
