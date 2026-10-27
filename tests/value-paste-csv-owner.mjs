/** CSV owner refuses XLSX typed errors while retaining literal error-looking strings. */
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import assert from 'node:assert/strict';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { record } from './support/file-record.mjs';
const output = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-value-paste-csv-owner-'));
const source = path.join(output, 'literal.csv'),
  bytes = Buffer.from('00123,=SUM(A1),original\r\n');
await fs.writeFile(source, bytes);
const recovery = path.join(output, 'recovery'),
  owner = new SpreadsheetDocument({ recoveryDirectory: recovery });
try {
  await owner.open(source);
  const before = owner.status(),
    cells = await owner.request({ operation: 'read', state: before.state, sheet: 'CSV', range: 'A1:C1' }),
    physicalBefore = await disk(recovery);
  const refusals = [];
  for (const commands of [
    [{ sheet: 'CSV', cell: 'A1', value: '#N/A', valueType: 'error' }],
    [{ type: 'setCells', sheet: 'CSV', cells: [{ cell: 'B1', value: '#N/A', valueType: 'error' }] }],
    [
      { sheet: 'CSV', cell: 'C1', value: 'must not apply' },
      { type: 'setCells', sheet: 'CSV', cells: [{ cell: 'A1', value: '#N/A', valueType: 'error' }] },
    ],
  ]) {
    let error;
    try {
      await owner.request({ operation: 'apply', state: before.state, commands });
    } catch (e) {
      error = String(e);
    }
    assert.match(error, /CSV.*literal text.*declared error/);
    assert.deepEqual(owner.status(), before);
    assert.deepEqual(
      await owner.request({ operation: 'read', state: before.state, sheet: 'CSV', range: 'A1:C1' }),
      cells
    );
    assert.deepEqual(await disk(recovery), physicalBefore);
    assert.deepEqual(await fs.readFile(source), bytes);
    refusals.push({ commands, error });
  }
  const accepted = await owner.request({
      operation: 'apply',
      state: before.state,
      commands: [{ type: 'setCells', sheet: 'CSV', cells: [{ cell: 'A1', value: '#N/A' }] }],
    }),
    head = owner.status().state,
    after = await owner.request({ operation: 'read', state: head, sheet: 'CSV', range: 'A1:C1' });
  assert.equal(after.cells.A1.value, '#N/A');
  assert.equal(after.cells.A1.dataType, 'text');
  assert.equal(after.cells.B1.value, '=SUM(A1)');
  assert.equal(after.cells.C1.value, 'original');
  await owner.request({ operation: 'undo' });
  assert.equal(owner.status().state, before.state);
  await owner.request({ operation: 'redo' });
  assert.equal(owner.status().state, head);
  assert.deepEqual(await fs.readFile(source), bytes);
} finally {
  const closing = await Promise.allSettled([owner.close()]);
  await fs.rm(output, { recursive: true, force: true });
  for (const result of closing) if (result.status === 'rejected') throw result.reason;
}
async function disk(directory) {
  const files = [];
  async function walk(dir) {
    for (const item of await fs.readdir(dir, { withFileTypes: true })) {
      assert(!item.isSymbolicLink());
      const file = path.join(dir, item.name);
      if (item.isDirectory()) await walk(file);
      else files.push(await record(file));
    }
  }
  await walk(directory);
  return files.sort((a, b) => a.path.localeCompare(b.path));
}
