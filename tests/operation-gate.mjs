import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, realpath, writeFile, readFile, rm, readdir } from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { operationGate } from '../adapters/operation-gate.mjs';
import { SpreadsheetWorker } from '../adapters/worker-client.mjs';
import { discardSpreadsheetRecovery, sha256 } from '../adapters/storage.mjs';
import { workbook } from '../fixtures/standalone.mjs';

test('a cancellation grant fences every later journal/publication boundary', () => {
  const events = [],
    gate = operationGate(7, (event) => events.push(event));
  gate.progress({ phase: 'applying', cancellable: true });
  assert.deepEqual(gate.authorize(), { granted: true });
  assert.throws(() => gate.enterCritical('saving-recovery'), { code: 'OPERATION_CANCELLED' });
  assert.throws(() => gate.check(), { code: 'OPERATION_CANCELLED' });
  assert(events.every((event) => event.requestId === 7));
  assert.equal(events.at(-1).cancellable, false);
});
test('worker critical state denies cancellation even when the last delivered UI phase was safe', () => {
  let last;
  const gate = operationGate(9, (event) => {
    if (event.cancellable) last = event;
  });
  gate.progress({ phase: 'exporting', cancellable: true });
  gate.enterCritical('writing-output');
  assert.equal(last.cancellable, true);
  assert.deepEqual(gate.authorize(), { granted: false, reason: 'finishing' });
  assert.doesNotThrow(() => gate.check());
  gate.finish();
  assert.deepEqual(gate.authorize(), { granted: false, reason: 'completed' });
});
test('progress has a bounded event count while omitted events still update the authoritative fence', () => {
  const events = [],
    gate = operationGate(1, (event) => events.push(event));
  for (let i = 0; i < 300; i++) gate.progress({ phase: i % 2 ? 'calculating' : 'applying', cancellable: true });
  assert.equal(events.length, 128);
  gate.enterCritical('saving-recovery');
  assert.equal(events.length, 128);
  assert.deepEqual(gate.authorize(), { granted: false, reason: 'finishing' });
});
test('a completed private initialization write releases the gate for subsequent preparation', () => {
  const gate = operationGate(1);
  gate.enterCritical('saving-recovery');
  assert.equal(gate.authorize().granted, false);
  gate.leaveCritical();
  gate.progress({ phase: 'indexing', cancellable: true });
  assert.equal(gate.authorize().granted, true);
});

async function fixture(t) {
  const root = await realpath(await mkdtemp(path.join(os.tmpdir(), 'xlsx-operation-gate-'))),
    filename = path.join(root, 'document.xlsx'),
    recoveryDirectory = path.join(root, 'recovery'),
    owners = [];
  await writeFile(filename, workbook());
  const open = async (options = {}) => {
    const owner = new SpreadsheetWorker({ filename, recoveryDirectory, ...options });
    owners.push(owner);
    return { owner, status: await owner.open() };
  };
  t.after(async () => {
    for (const owner of owners) await owner.close().catch(() => {});
    await rm(root, { recursive: true, force: true });
  });
  return { root, filename, recoveryDirectory, owners, open };
}
const read = (owner, status) => owner.request({ operation: 'read', state: status.state, sheet: 'Sheet1', range: 'A1' });
const edit = (owner, status, value) =>
  owner.request({ operation: 'apply', state: status.state, commands: [{ sheet: 'Sheet1', cell: 'A1', value }] });

for (const uuid of [false, true])
  test(
    (uuid ? 'UUID' : 'legacy') + ' resume retains saved undo/redo and accepted edits while the public file changes',
    async (t) => {
      const f = await fixture(t),
        { owner, status } = await f.open(uuid ? { fresh: true } : {});
      let changed = await edit(owner, status, 111);
      changed = await edit(owner, changed, 222);
      await owner.request({ operation: 'save' });
      const undone = await owner.request({ operation: 'undo' });
      assert.equal(undone.canRedo, true);
      await owner.cancel();
      await writeFile(f.filename, 'changed externally');
      const resumed = await f.open(uuid ? { filename: undefined, recoveryId: undone.documentId } : { resume: true });
      assert.equal((await read(resumed.owner, resumed.status)).cells.A1.value, 111);
      assert.equal(resumed.status.canRedo, true);
      assert.equal(resumed.status.canUndo, true);
      const redone = await resumed.owner.request({ operation: 'redo' });
      assert.equal((await read(resumed.owner, redone)).cells.A1.value, 222);
      assert.equal(redone.dirty, false);
      await assert.rejects(resumed.owner.request({ operation: 'save' }), { code: 'REVISION_CONFLICT' });
      assert.equal(await readFile(f.filename, 'utf8'), 'changed externally');
    }
  );

test('a real native operation is stopped only after its worker grant and recovers the accepted undo cursor', async (t) => {
  const f = await fixture(t),
    events = [],
    opened = await f.open({ onProgress: (event) => events.push(event) });
  const a = await edit(opened.owner, opened.status, 111),
    b = await edit(opened.owner, a, 222);
  await opened.owner.request({ operation: 'save' });
  const accepted = await opened.owner.request({ operation: 'undo' });
  let requestId, start;
  const began = new Promise((resolve) => {
    start = resolve;
  });
  const result = opened.owner
    .request(
      {
        operation: 'apply',
        state: accepted.state,
        commands: [{ type: 'formatRange', sheet: 'Sheet1', range: 'A2:A200000', format: { bold: true } }],
      },
      {
        onStart: (id) => {
          requestId = id;
        },
        onProgress: (event) => {
          events.push(event);
          if (event.phase === 'applying' && event.cancellable) start();
        },
      }
    )
    .then(
      (value) => ({ value }),
      (error) => ({ error })
    );
  await began;
  const storage = path.join(
    os.tmpdir(),
    `spreadsheet-workers-v1${process.platform === 'linux' ? `-${process.getuid()}` : ''}`
  );
  let observed = false;
  for (let i = 0; i < 1000 && !observed; i++) {
    for (const scope of await readdir(storage))
      try {
        const root = path.join(storage, scope),
          lease = JSON.parse(await readFile(path.join(root, 'worker-owner.json'), 'utf8'));
        if (lease.pid !== process.pid) continue;
        for (const child of await readdir(root))
          if (child.startsWith('owner-'))
            try {
              const active = JSON.parse(await readFile(path.join(root, child, 'native-active.json'), 'utf8'));
              observed = Number.isSafeInteger(active.pid);
            } catch (error) {
              if (error.code !== 'ENOENT' && !(error instanceof SyntaxError)) throw error;
            }
      } catch (error) {
        if (error.code !== 'ENOENT' && !(error instanceof SyntaxError)) throw error;
      }
    if (!observed) await new Promise((resolve) => setTimeout(resolve, 2));
  }
  assert(observed, 'must observe native work before asking its worker for cancellation');
  const grant = await opened.owner.authorizeCancel(requestId);
  assert.equal(grant.granted, true, 'must receive a worker cancellation grant before killing its owner');
  await opened.owner.cancel();
  assert((await result).error);
  assert(events.some((event) => event.phase === 'indexing'));
  assert(events.some((event) => event.phase === 'applying'));
  const resumed = await f.open({ resume: true });
  assert.equal(resumed.status.state, accepted.state);
  assert.equal((await read(resumed.owner, resumed.status)).cells.A1.value, 111);
  assert.equal(resumed.status.canRedo, true);
  const redone = await resumed.owner.request({ operation: 'redo' });
  assert.equal((await read(resumed.owner, redone)).cells.A1.value, 222);
  assert.deepEqual(await resumed.owner.authorizeCancel(requestId + 100), { granted: false, reason: 'completed' });
});

test('corrupt undo topology and corrupt discard identities preserve their private records', async (t) => {
  const f = await fixture(t),
    { owner, status } = await f.open({ fresh: true }),
    changed = await edit(owner, status, 42);
  await owner.cancel();
  const directory = path.join(f.recoveryDirectory, 'draft-' + changed.documentId),
    file = path.join(directory, 'current.json'),
    manifest = JSON.parse(await readFile(file));
  const corrupt = JSON.stringify({
    ...manifest,
    history: {
      cursor: 0,
      frames: [
        { head: manifest.head, hash: manifest.headHash },
        { head: null, hash: null },
      ],
    },
  });
  await writeFile(file, corrupt);
  const broken = new SpreadsheetWorker({ recoveryId: changed.documentId, recoveryDirectory: f.recoveryDirectory });
  f.owners.push(broken);
  await assert.rejects(broken.open(), { code: 'session-bad-input' });
  await broken.close();
  assert.equal(await readFile(file, 'utf8'), corrupt);
  await writeFile(file, JSON.stringify({ ...manifest, documentId: '00000000-0000-0000-0000-000000000000' }));
  await assert.rejects(discardSpreadsheetRecovery(f.recoveryDirectory, changed.documentId), /identity/);
  assert((await readdir(directory)).length > 1);
  await writeFile(file, JSON.stringify(manifest));
  await discardSpreadsheetRecovery(f.recoveryDirectory, changed.documentId);
  await assert.rejects(readFile(file), { code: 'ENOENT' });
  assert.equal(sha256(await readFile(f.filename)), manifest.baseRevision);
});
