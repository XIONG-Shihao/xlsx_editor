import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { unzipSync } from 'fflate';
import { SMARTART_IDS, smartartFixture } from './support/smartart-fixture.mjs';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
const request = (o, operation, more = {}) =>
    o.request({ operation, documentId: o.status().documentId, state: o.status().state, ...more }),
  sha = (b) => createHash('sha256').update(b).digest('hex');
async function fixture(t) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'smartart-reader-owner-')),
    source = path.join(root, 'source.xlsx'),
    bytes = smartartFixture(),
    o = new SpreadsheetDocument({ recoveryDirectory: path.join(root, 'recovery'), cacheIdentity: null });
  await fs.writeFile(source, bytes);
  await o.open(source, { fresh: true });
  t.after(async () => {
    await o.close();
    await fs.rm(root, { recursive: true, force: true });
  });
  return { o, root, source, bytes };
}
test('native owner reads three semantic samples without data/head/history writes and copies their graph intact', async (t) => {
  const f = await fixture(t),
    before = structuredClone(f.o.status()),
    entries = (await request(f.o, 'objects', { sheet: 'Sheet1' })).objects;
  assert.equal(entries.length, 3);
  for (const id of SMARTART_IDS) {
    const p = await request(f.o, 'object', { sheet: 'Sheet1', objectId: id });
    assert.equal(p.kind, 'smartart');
    assert.equal(p.smartart.nodes.length, 3);
    assert.equal(p.canEdit, true);
    assert.equal(p.canDelete, true);
  }
  assert.deepEqual(f.o.status(), before);
  const copy = path.join(f.root, 'copy.xlsx');
  await request(f.o, 'copy', { destination: copy, expectedRevision: null });
  const old = unzipSync(f.bytes),
    saved = unzipSync(await fs.readFile(copy));
  for (const part of Object.keys(old).filter(
    (p) => p.startsWith('xl/diagrams/') || p.startsWith('xl/drawings/') || p === 'custom/smartart-opaque.bin'
  ))
    assert.equal(sha(saved[part]), sha(old[part]), part);
  assert.equal(sha(await fs.readFile(f.source)), sha(f.bytes));
});
test('generic deletion routes to the qualified SmartArt owner; structural transformations remain protected', async (t) => {
  const f = await fixture(t),
    p = await request(f.o, 'object', { sheet: 'Sheet1', objectId: SMARTART_IDS[0] });
  await request(f.o, 'apply', {
    commands: [{ operation: 'deleteObject', sheet: 'Sheet1', objectId: p.id, expectedObjectRevision: p.revision }],
  });
  assert.equal((await request(f.o, 'objects', { sheet: 'Sheet1' })).total, 2);
  await request(f.o, 'undo');
  assert.equal((await request(f.o, 'objects', { sheet: 'Sheet1' })).total, 3);
  const before = structuredClone(f.o.status());
  for (const command of [
    { operation: 'insertRow', sheet: 'Sheet1', row: 1 },
    { operation: 'insertColumn', sheet: 'Sheet1', column: 1 },
  ])
    await assert.rejects(request(f.o, 'apply', { commands: [command] }));
  assert.deepEqual(f.o.status(), before);
});
test('existing text authoring coexists with the immutable diagram graphs and unrelated source members', async (t) => {
  const f = await fixture(t),
    id = '44444444-4444-4444-8444-444444444444',
    before = unzipSync(f.bytes);
  await request(f.o, 'apply', {
    commands: [
      {
        operation: 'createShape',
        sheet: 'Sheet1',
        objectId: id,
        expectedObjectRevision: null,
        shape: { sample: 'text-note', text: 'Alongside SmartArt' },
        anchor: { cell: 'B4' },
      },
    ],
  });
  const entry = await request(f.o, 'object', { sheet: 'Sheet1', objectId: id });
  assert(entry.canEdit);
  assert.equal((await request(f.o, 'objects', { sheet: 'Sheet1' })).total, 4);
  for (const sid of SMARTART_IDS) assert((await request(f.o, 'object', { sheet: 'Sheet1', objectId: sid })).smartart);
  const copy = path.join(f.root, 'with-text.xlsx');
  await request(f.o, 'copy', { destination: copy, expectedRevision: null });
  const saved = unzipSync(await fs.readFile(copy));
  for (const part of Object.keys(before).filter(
    (p) => p.startsWith('xl/diagrams/') || p === 'custom/smartart-opaque.bin'
  ))
    assert.equal(sha(saved[part]), sha(before[part]), part);
  await request(f.o, 'apply', {
    commands: [{ operation: 'deleteShape', sheet: 'Sheet1', objectId: id, expectedObjectRevision: entry.revision }],
  });
  assert.equal((await request(f.o, 'objects', { sheet: 'Sheet1' })).total, 3);
});
