import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { SpreadsheetWorker } from '../adapters/worker-client.mjs';
import { nativeWorkbookOperationAsync, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';

after(() => closeNativeWorkbookWorker());
const id = '11111111-1111-4111-8111-111111111111';
async function fixture(t, file = 'w7-managed-preview.xlsx') {
  const directory = await fs.realpath(await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-document-objects-'))),
    filename = path.join(directory, file),
    recoveryDirectory = path.join(directory, 'recovery'),
    owners = [];
  if (file.endsWith('.csv')) await fs.writeFile(filename, 'a,b\n1,2\n');
  else
    await fs.copyFile(
      path.resolve(file === 'w7-managed-preview.xlsx' ? 'tests/support' : 'tests/support/w7-reference-data', file),
      filename
    );
  const make = (worker) => {
    const owner = worker
      ? new SpreadsheetWorker({ filename, recoveryDirectory })
      : new SpreadsheetDocument({
          recoveryDirectory,
          native: (input) => nativeWorkbookOperationAsync(input),
          cacheIdentity: null,
        });
    owners.push(owner);
    return owner;
  };
  t.after(async () => {
    const failures = [];
    try {
      for (const owner of owners)
        try {
          await owner.close();
        } catch (error) {
          failures.push(error);
        }
    } finally {
      await fs.rm(directory, { recursive: true, force: true });
    }
    if (failures.length) throw new AggregateError(failures);
  });
  return { directory, filename, make };
}
const request = (owner, status, operation, extra = {}) =>
  owner.request({ operation, documentId: status.documentId, state: status.state, sheet: 'Sales', ...extra });
for (const worker of [false, true])
  test(`${worker ? 'worker' : 'Document'} object pages and preview stay on the exact accepted document/state with no private paths`, async (t) => {
    const f = await fixture(t),
      owner = f.make(worker);
    let status = worker ? await owner.open() : await owner.open(f.filename);
    const page = await request(owner, status, 'objects', { limit: 1 });
    assert(page.total >= 1);
    assert.equal(page.objects.length, 1);
    assert.equal(page.objects[0].id, id);
    assert.equal(page.nextCursor, page.total > 1 ? id : null);
    assert.equal(page.objects[0].canEdit, true);
    const preview = await request(owner, status, 'object', { objectId: id });
    assert.deepEqual(preview.data.categories, ['Jan', 'Feb', 'Mar']);
    assert.deepEqual(
      preview.data.series.map((series) => series.values),
      [
        [2, 3, 5],
        [1, 2, 3],
      ]
    );
    assert(!JSON.stringify(page).includes(f.directory));
    assert(!JSON.stringify(preview).includes(f.directory));
    assert(!JSON.stringify(preview).includes('xl/'));
    for (const extra of [
      { documentId: 'other' },
      { state: 'unknown' },
      { sheet: 'Missing' },
      { limit: 0 },
      { limit: 51 },
      { limit: 1.5 },
      { after: '../../private' },
      { after: '' },
      { objectId: id },
    ])
      await assert.rejects(request(owner, status, 'objects', extra), undefined, JSON.stringify(extra));
    for (const extra of [
      { objectId: '../../private' },
      { objectId: 'not-an-id' },
      { objectId: id, limit: 1 },
      { objectId: id, after: id },
      { objectId: id, documentId: 'other' },
    ])
      await assert.rejects(request(owner, status, 'object', extra));
    const before = status;
    status = await owner.request({
      operation: 'apply',
      state: status.state,
      commands: [{ sheet: 'Sales', cell: 'B2', value: 21 }],
    });
    await assert.rejects(request(owner, before, 'objects'), /stale/);
    await assert.rejects(request(owner, before, 'object', { objectId: id }), /stale/);
    assert.equal((await request(owner, status, 'object', { objectId: id })).data.series[0].values[0], 21);
  });

test('opaque object IDs and bounded cursors list protected drawings without granting authoring', async (t) => {
  const f = await fixture(t, 'opaque-shared.xlsx'),
    owner = f.make(false),
    status = await owner.open(f.filename),
    first = await request(owner, status, 'objects', { limit: 2 });
  assert.equal(first.objects.length, 2);
  assert(first.objects.every((object) => object.protectedReason));
  assert(first.nextCursor);
  const second = await request(owner, status, 'objects', { limit: 2, after: first.nextCursor });
  assert.equal(second.objects.length, 2);
  assert(!second.objects.some((object) => first.objects.some((old) => old.id === object.id)));
  const rest = second.nextCursor
    ? await request(owner, status, 'objects', { limit: 2, after: second.nextCursor })
    : { objects: [] };
  const opaque = [...first.objects, ...second.objects, ...rest.objects].find((object) =>
    object.id.startsWith('opaque-')
  );
  assert(opaque);
  const preview = await request(owner, status, 'object', { objectId: opaque.id });
  assert.equal(preview.canEdit, false);
  assert(preview.protectedReason);
  assert(!JSON.stringify(preview).includes(f.directory));
});

test('literal CSV owners refuse both object DTO operations before any object lookup', async (t) => {
  const f = await fixture(t, 'literal.csv'),
    owner = f.make(false),
    status = await owner.open(f.filename);
  await assert.rejects(request(owner, status, 'objects'), /XLSX/);
  await assert.rejects(request(owner, status, 'object', { objectId: id }), /XLSX/);
});
