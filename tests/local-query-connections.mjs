import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { LocalQueryConnections } from '../adapters/local-query-connections.mjs';
const draft = {
  name: 'Active contacts',
  target: { sheet: 'Sheet1', start: 'A1' },
  query: {
    columns: [1, 3],
    filters: [{ column: 2, op: 'equals', value: 'active', caseSensitive: true }],
    keepHeader: true,
  },
};
async function fixture(t) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-query-registry-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  const source = path.join(root, 'source.csv'),
    filename = path.join(root, 'queries.json');
  await fs.writeFile(source, 'name,state,code\na,active,01\n');
  return { root, source, filename, store: await LocalQueryConnections.open(filename) };
}
test('saved connection grants are scoped and opaque, persist, and exact prior definitions protect edits', async (t) => {
  const { store, source, filename } = await fixture(t),
    r = await store.create('document-A', draft, source);
  assert.equal(JSON.stringify(r).includes(source), false);
  assert.equal((await store.list('document-B')).entries.length, 0);
  await assert.rejects(store.granted('document-B', r.id), /access/);
  const changed = await store.update('document-A', r.id, 1, { ...draft, name: 'Current' });
  assert.equal(changed.revision, 2);
  await assert.rejects(store.update('document-A', r.id, 1, draft), /prior/);
  const reopened = await LocalQueryConnections.open(filename);
  assert.equal((await reopened.list('document-A')).entries[0].name, 'Current');
  assert.equal((await reopened.granted('document-A', r.id, 2)).source.filename, await fs.realpath(source));
});
test('corrupt and externally modified registry bytes are preserved and boot remains possible', async (t) => {
  const { store, source, filename } = await fixture(t);
  await store.create('doc', draft, source);
  const corrupt = Buffer.from('{wrong');
  await fs.writeFile(filename, corrupt);
  await assert.rejects(store.create('doc', draft, source), /changed/);
  assert.deepEqual(await fs.readFile(filename), corrupt);
  const reopened = await LocalQueryConnections.open(filename);
  assert.match((await reopened.list('doc')).warning, /preserved/);
  await assert.rejects(reopened.create('doc', draft, source), /preserved/);
  assert.deepEqual(await fs.readFile(filename), corrupt);
});
test('pending intent survives restart; absence is uncertainty, accepted Undo/Redo receipt finalizes exact descriptor', async (t) => {
  const { store, source, filename } = await fixture(t),
    r = await store.create('doc', draft, source),
    descriptor = {
      type: 'importCells',
      id: randomUUID(),
      sha256: 'a'.repeat(64),
      bytes: 20,
      count: 2,
      shape: { sheet: 'Sheet1', start: 'A1', rows: 1, columns: 2 },
      policy: 'text',
    },
    intent = {
      id: randomUUID(),
      descriptor,
      parentState: 'base',
      output: {
        shape: descriptor.shape,
        fingerprint: 'b'.repeat(64),
        sourceSha256: 'c'.repeat(64),
        sourceBytes: 30,
        outputRows: 1,
        outputColumns: 2,
        acceptedHead: 'base',
        updatedAt: Date.now(),
      },
    };
  await store.begin('doc', r.id, 1, intent);
  const recovered = await LocalQueryConnections.open(filename);
  assert.equal((await recovered.list('doc')).entries[0].pending, true);
  await assert.rejects(recovered.finalize('doc', r.id, intent.id, { found: false }), /receipt/);
  await assert.rejects(recovered.remove('doc', r.id, 1), /pending/);
  await assert.rejects(recovered.update('doc', r.id, 1, draft), /pending/);
  const finalized = await recovered.finalize('doc', r.id, intent.id, {
    found: true,
    parentState: 'base',
    head: 'accepted',
    descriptor: { ...descriptor, shape: { columns: 2, rows: 1, start: 'A1', sheet: 'Sheet1' } },
  });
  assert.equal(finalized.pending, undefined);
  assert.equal(finalized.last.acceptedHead, 'accepted');
  assert.equal(finalized.revision, 2);
});
test('per-document cap refuses without eviction; unknown keys/path injection refuse', async (t) => {
  const { store, source } = await fixture(t);
  for (let i = 0; i < 32; i++) await store.create('doc', { ...draft, name: 'Query ' + i }, source);
  await assert.rejects(store.create('doc', draft, source), { code: 'query-limit' });
  assert.equal((await store.list('doc')).entries.length, 32);
  await assert.rejects(store.create('other', { ...draft, filename: source }, source), /settings/);
});
