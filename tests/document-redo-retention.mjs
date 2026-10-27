import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { SpreadsheetDocument, listSpreadsheetRecovery } from '../adapters/document-session.mjs';
import { createWorkbook } from '../core/workbook-package.mjs';
import { nativeWorkbookOperationAsync, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { revision, listDraftRecords } from '../adapters/storage.mjs';

after(() => closeNativeWorkbookWorker());
const read = (owner) => owner.request({ operation: 'read', state: owner.status().state, sheet: 'Sheet1', range: 'A1' });
async function fixture(t) {
  const directory = await fs.realpath(await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-redo-retention-'))),
    recoveryDirectory = path.join(directory, 'recovery'),
    filename = path.join(directory, 'saved.xlsx'),
    owners = [];
  await fs.writeFile(filename, createWorkbook());
  const make = () => {
    const owner = new SpreadsheetDocument({
      recoveryDirectory,
      native: (input) => nativeWorkbookOperationAsync(input),
    });
    owners.push(owner);
    return owner;
  };
  t.after(async () => {
    for (const owner of owners) await owner.close();
    await fs.rm(directory, { recursive: true, force: true });
  });
  return { directory, recoveryDirectory, filename, make };
}

test('legacy ordinary literal Save edit Undo close Open restores its clean cursor and unsaved redo tip', async (t) => {
  const f = await fixture(t),
    owner = f.make();
  await owner.open(f.filename);
  await owner.request({ operation: 'save' });
  const savedHash = await revision(f.filename),
    base = owner.status().state;
  await owner.request({ operation: 'apply', state: base, commands: [{ sheet: 'Sheet1', cell: 'A1', value: 23 }] });
  const edited = owner.status().state;
  await owner.request({ operation: 'undo' });
  assert.equal(owner.status().dirty, false);
  assert.equal(owner.status().canRedo, true);
  await owner.close();
  const restored = f.make();
  await restored.open(f.filename);
  assert.equal(restored.status().state, base);
  assert.equal(restored.status().dirty, false);
  assert.equal(restored.status().canRedo, true);
  assert.equal(restored.status().recovered, true);
  assert.equal((await read(restored)).cells.A1, undefined);
  await restored.request({ operation: 'redo' });
  assert.equal(restored.status().state, edited);
  assert.equal((await read(restored)).cells.A1.value, 23);
  assert.equal(await revision(f.filename), savedHash);
  await restored.request({ operation: 'save' });
  assert.notEqual(await revision(f.filename), savedHash);
});

test('bound UUID redo-only draft is discovered and ordinary Open retains its first-save binding and redo', async (t) => {
  const f = await fixture(t),
    owner = f.make();
  await owner.create('first saved');
  await owner.request({
    operation: 'saveAs',
    destination: path.join(f.directory, 'first.xlsx'),
    expectedRevision: null,
  });
  const filename = (await owner.request({ operation: 'binding' })).filename,
    id = owner.status().documentId,
    base = owner.status().state,
    savedHash = await revision(filename);
  await owner.request({ operation: 'apply', state: base, commands: [{ sheet: 'Sheet1', cell: 'A1', value: 31 }] });
  const edited = owner.status().state;
  await owner.request({ operation: 'undo' });
  await owner.close();
  const listed = await listSpreadsheetRecovery(f.recoveryDirectory);
  assert.equal(listed.entries.length, 1);
  assert.equal(listed.entries[0].id, id);
  assert.equal(listed.entries[0].dirty, false);
  assert.equal(listed.entries[0].unsaved, false);
  assert.equal((await listDraftRecords(f.recoveryDirectory)).records.length, 1);
  const restored = f.make();
  await restored.open(filename);
  assert.equal(restored.status().documentId, id);
  assert.equal(restored.status().state, base);
  assert.equal(restored.status().canRedo, true);
  await restored.request({ operation: 'redo' });
  assert.equal(restored.status().state, edited);
  assert.equal((await read(restored)).cells.A1.value, 31);
  assert.equal(await revision(filename), savedHash);
});

test('unsaved private UUID redo-only literal survives close and recovery without creating a public file', async (t) => {
  const f = await fixture(t),
    owner = f.make();
  await owner.create('private redo');
  const id = owner.status().documentId,
    base = owner.status().state;
  await owner.request({
    operation: 'apply',
    state: base,
    commands: [{ sheet: 'Sheet1', cell: 'A1', value: 'private' }],
  });
  await owner.request({ operation: 'undo' });
  assert.equal(owner.status().dirty, false);
  await owner.close();
  const listed = await listSpreadsheetRecovery(f.recoveryDirectory);
  assert.equal(listed.entries.length, 1);
  assert.equal(listed.entries[0].id, id);
  assert.equal(listed.entries[0].unsaved, true);
  const restored = f.make();
  await restored.recover(id);
  assert.equal(restored.status().state, base);
  assert.equal(restored.status().canRedo, true);
  assert.equal(restored.status().unsaved, true);
  await restored.request({ operation: 'redo' });
  assert.equal((await read(restored)).cells.A1.value, 'private');
  assert.equal((await restored.request({ operation: 'binding' })).filename, null);
});

test('redo-only recovery still refuses external source publication conflicts and leaves older owner paths isolated', async (t) => {
  const f = await fixture(t),
    owner = f.make();
  await owner.open(f.filename);
  const originalHash = await revision(f.filename);
  await owner.request({
    operation: 'apply',
    state: owner.status().state,
    commands: [{ sheet: 'Sheet1', cell: 'A1', value: 'old accepted' }],
  });
  await owner.request({ operation: 'undo' });
  await owner.close();
  await fs.writeFile(f.filename, Buffer.from(createWorkbook()));
  // Give the public file a different valid package identity, preserving the old private snapshot.
  const other = f.make(),
    otherFile = path.join(f.directory, 'other.xlsx');
  await other.create('other owner');
  await other.request({
    operation: 'apply',
    state: other.status().state,
    commands: [{ sheet: 'Sheet1', cell: 'A1', value: 'external' }],
  });
  await other.request({ operation: 'saveAs', destination: otherFile, expectedRevision: null });
  await fs.copyFile(otherFile, f.filename);
  const externalHash = await revision(f.filename);
  assert.notEqual(externalHash, originalHash);
  const restored = f.make();
  await restored.open(f.filename);
  await restored.request({ operation: 'redo' });
  assert.equal((await read(restored)).cells.A1.value, 'old accepted');
  assert.equal((await read(other)).cells.A1.value, 'external');
  await assert.rejects(restored.request({ operation: 'save' }), /changed on disk/);
  assert.equal(await revision(f.filename), externalHash);
  const copied = path.join(f.directory, 'old-copy.xlsx');
  await restored.request({ operation: 'copy', destination: copied, expectedRevision: null });
  assert.equal(await revision(f.filename), externalHash);
  assert.notEqual(await revision(copied), externalHash);
});

test('discovery excludes clean bound base tips with absent or explicit null-base history and fully saved tips', async (t) => {
  const f = await fixture(t),
    owner = f.make();
  await owner.create('clean tip');
  const filename = path.join(f.directory, 'clean.xlsx');
  await owner.request({ operation: 'saveAs', destination: filename, expectedRevision: null });
  assert.equal((await listSpreadsheetRecovery(f.recoveryDirectory)).entries.length, 0);
  const record = path.join(f.recoveryDirectory, 'draft-' + owner.status().documentId, 'current.json'),
    manifest = JSON.parse(await fs.readFile(record, 'utf8'));
  assert.equal(manifest.history, undefined);
  await fs.writeFile(
    record,
    JSON.stringify({ ...manifest, history: { cursor: 0, frames: [{ head: null, hash: null }] } })
  );
  assert.equal((await listSpreadsheetRecovery(f.recoveryDirectory)).entries.length, 0);
  await owner.request({
    operation: 'apply',
    state: owner.status().state,
    commands: [{ sheet: 'Sheet1', cell: 'A1', value: 'saved' }],
  });
  await owner.request({ operation: 'save' });
  assert.equal((await listSpreadsheetRecovery(f.recoveryDirectory)).entries.length, 0);
  await owner.close();
  assert.equal((await listDraftRecords(f.recoveryDirectory)).records.length, 0);
});

for (const schema of [1, 2])
  test(`first Save refuses a destination with schema${schema} clean-cursor unsaved redo history`, async (t) => {
    const f = await fixture(t),
      owner = f.make();
    let filename = f.filename;
    if (schema === 1) await owner.open(filename);
    else {
      await owner.create('bound history');
      filename = path.join(f.directory, 'bound.xlsx');
      await owner.request({ operation: 'saveAs', destination: filename, expectedRevision: null });
    }
    const savedHash = await revision(filename);
    await owner.request({
      operation: 'apply',
      state: owner.status().state,
      commands: [{ sheet: 'Sheet1', cell: 'A1', value: 'redo kept' }],
    });
    await owner.request({ operation: 'undo' });
    await owner.close();
    const replacement = f.make();
    await replacement.create('replacement');
    await replacement.request({
      operation: 'apply',
      state: replacement.status().state,
      commands: [{ sheet: 'Sheet1', cell: 'A1', value: 'replacement' }],
    });
    await assert.rejects(
      replacement.request({ operation: 'saveAs', destination: filename, expectedRevision: savedHash }),
      { code: 'session-refused' }
    );
    assert.equal(await revision(filename), savedHash);
    const restored = f.make();
    await restored.open(filename);
    await restored.request({ operation: 'redo' });
    assert.equal((await read(restored)).cells.A1.value, 'redo kept');
  });
