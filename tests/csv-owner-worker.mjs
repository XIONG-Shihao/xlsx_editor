import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { Worker } from 'node:worker_threads';
import { CsvFileWorker } from '../adapters/csv-worker.mjs';

const storage = path.join(
  os.tmpdir(),
  `spreadsheet-workers-v1${process.platform === 'linux' ? `-${process.getuid()}` : ''}`
);
async function scopes() {
  const names = await fs.readdir(storage).catch((error) => {
    if (error.code === 'ENOENT') return [];
    throw error;
  });
  const result = [];
  for (const name of names) {
    try {
      if (JSON.parse(await fs.readFile(path.join(storage, name, 'worker-owner.json'), 'utf8')).pid === process.pid)
        result.push(name);
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
    }
  }
  return result.sort();
}
async function fixture(t, text = 'key,value\r\n001,"before\nline"\r\n002,tail\r\n') {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'csv-owner-worker-')),
    source = path.join(root, 'source.csv'),
    owners = [];
  await fs.writeFile(source, text);
  t.after(async () => {
    for (const owner of owners) await owner.close().catch(() => {});
    await fs.rm(root, { recursive: true, force: true });
  });
  const make = () => {
    const owner = new CsvFileWorker();
    owners.push(owner);
    return owner;
  };
  return { root, source, text, make };
}

test('CSV worker transfers metadata and bounded ranges while revisions/export stay isolated', async (t) => {
  const f = await fixture(t),
    before = await scopes(),
    a = f.make(),
    b = f.make();
  const initial = await a.open(f.source),
    second = await b.open(f.source);
  assert.equal(initial.storage.sessionId, a.id);
  assert.notEqual(initial.storage.sessionId, second.storage.sessionId);
  assert.deepEqual(initial.sheets[0].cells, {});
  const serialized = JSON.stringify(initial);
  for (const forbidden of ['rowOffsets', 'compressedText', 'blocks', f.root, 'before'])
    assert(!serialized.includes(forbidden));
  assert(serialized.length < 2048);
  const edited = await a.apply('0', [
    { sheet: 'CSV', cell: 'B2', value: '=literal,🙂', expectedValue: 'before\nline' },
  ]);
  assert.equal(edited.view.storage.sessionId, a.id);
  assert.equal((await a.read(edited.view.storage.state, 'CSV', 'B2')).cells.B2.value, '=literal,🙂');
  assert.equal((await a.read('0', 'CSV', 'B2')).cells.B2.value, 'before\nline');
  assert.equal((await b.read('0', 'CSV', 'B2')).cells.B2.value, 'before\nline');
  const saved = await a.export(edited.view.storage.state);
  assert.equal(await fs.readFile(saved.path, 'utf8'), f.text.replace('before\nline', '=literal,🙂'));
  assert.equal(await fs.readFile(f.source, 'utf8'), f.text);
  await a.close();
  assert.equal((await b.read('0', 'CSV', 'B3')).cells.B3.value, 'tail');
  await b.close();
  assert.deepEqual(await scopes(), before);
});

test('close fences new calls, drains a full request queue and awaits storage cleanup', async (t) => {
  const f = await fixture(t),
    before = await scopes(),
    owner = f.make();
  await owner.open(f.source);
  const reads = Array.from({ length: 16 }, () => owner.read('0', 'CSV', 'B2'));
  const first = owner.close();
  assert.equal(owner.close(), first);
  await assert.rejects(owner.read('0', 'CSV', 'B2'), /closed/);
  const results = await Promise.all(reads);
  assert(results.every((result) => result.cells.B2.value === 'before\nline'));
  await first;
  assert.deepEqual(await scopes(), before);
  assert.equal(owner.diagnosticStats().closed, 1);
  assert.equal(owner.diagnosticStats().retainedStates, 0);
  assert.equal(owner.diagnosticStats().storageRoots, 0);
  assert.equal(owner.diagnosticStats().pendingRequests, 0);
});

test('failed open and close racing initial collection cannot leave a live owner', async (t) => {
  const f = await fixture(t),
    before = await scopes(),
    failed = f.make();
  await assert.rejects(failed.open(path.join(f.root, 'missing.csv')), /ENOENT/);
  assert.equal(failed.diagnosticStats().closed, 1);
  assert.equal(failed.diagnosticStats().workerId, -1);
  assert.deepEqual(await scopes(), before);
  const early = f.make(),
    opening = assert.rejects(early.open(f.source), /closed/);
  await early.close();
  await opening;
  assert.deepEqual(await scopes(), before);
});

test('oversized/cloning and reply-budget failures preserve the accepted revision', async (t) => {
  const f = await fixture(t, 'key,value\n001,' + 'x'.repeat(4096) + '\n002,tail\n'),
    owner = f.make();
  await owner.open(f.source, { maxReplyBytes: 1024 });
  await assert.rejects(owner.read('0', 'CSV', 'B2'), /reply.*budget/);
  await assert.rejects(owner.query('0', 'CSV', 'x'.repeat(2 * 1024 ** 2), 'display'), {
    code: /^(csv-limit|session-limit)$/,
  });
  await assert.rejects(owner.apply('0', [{ sheet: 'CSV', cell: 'B3', value: () => {} }]), { name: 'DataCloneError' });
  await assert.rejects(owner.apply('0', [{ sheet: 'CSV', cell: 'B3', value: 'a'.repeat(2048) }]), /reply.*budget/);
  assert.equal(owner.diagnosticStats().retainedStates, 1);
  assert.equal((await owner.read('0', 'CSV', 'B3')).cells.B3.value, 'tail');
  const good = await owner.apply('0', [{ sheet: 'CSV', cell: 'B3', value: 'accepted' }]);
  assert.equal((await owner.read(good.view.storage.state, 'CSV', 'B3')).cells.B3.value, 'accepted');
});

for (const kind of ['exit', 'invalid-reply'])
  test(`${kind} rejects pending work and clears disposed-owner statistics`, async (t) => {
    const f = await fixture(t),
      before = await scopes(),
      owner = f.make();
    await owner.open(f.source);
    const original = Worker.prototype.postMessage;
    Worker.prototype.postMessage = function (message, ...args) {
      if (message.method === 'read') {
        if (kind === 'exit') void this.terminate();
        else this.emit('message', { id: message.id + 1000, result: {} });
        return;
      }
      return original.call(this, message, ...args);
    };
    try {
      await assert.rejects(owner.read('0', 'CSV', 'B2'), { code: /^(internal-invalid-field|session-closed)$/ });
    } finally {
      Worker.prototype.postMessage = original;
    }
    await owner.close();
    assert.deepEqual(await scopes(), before);
    assert.equal(owner.diagnosticStats().retainedStates, 0);
    assert.equal(owner.diagnosticStats().storageRoots, 0);
  });

test('opt-in worker heap counters are numeric, content-free and zero after disposal', async (t) => {
  const f = await fixture(t),
    previous = process.env.XLSX_EDITOR_TRACE_PATH;
  process.env.XLSX_EDITOR_TRACE_PATH = 'test-enables-counters-only';
  const owner = f.make();
  if (previous === undefined) delete process.env.XLSX_EDITOR_TRACE_PATH;
  else process.env.XLSX_EDITOR_TRACE_PATH = previous;
  await owner.open(f.source);
  const stats = owner.diagnosticStats();
  assert(stats.workerHeapUsed > 0);
  assert(stats.workerHeapTotal >= stats.workerHeapUsed);
  assert(Object.values(stats).every(Number.isFinite));
  assert(!JSON.stringify(stats).includes(f.root));
  await owner.close();
  assert.equal(owner.diagnosticStats().workerHeapUsed, 0);
});
