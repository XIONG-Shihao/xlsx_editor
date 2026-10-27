import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { Worker } from 'node:worker_threads';
import { spawn } from 'node:child_process';
import { SpreadsheetWorker } from '../adapters/worker-client.mjs';
import { allocateWorkerStorage, releaseWorkerStorage } from '../adapters/worker-storage.mjs';
import { workbook } from '../fixtures/standalone.mjs';
const storage = path.join(
  os.tmpdir(),
  `spreadsheet-workers-v1${process.platform === 'linux' ? `-${process.getuid()}` : ''}`
);
async function scopes(pid = process.pid) {
  const names = await fs.readdir(storage).catch((e) => {
      if (e.code === 'ENOENT') return [];
      throw e;
    }),
    owned = [];
  for (const name of names)
    try {
      if (JSON.parse(await fs.readFile(path.join(storage, name, 'worker-owner.json'), 'utf8')).pid === pid)
        owned.push(name);
    } catch (error) {
      if (error.code !== 'ENOENT' && !(error instanceof SyntaxError)) throw error;
    }
  return new Set(owned);
}
async function fixture(t) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-worker-lifecycle-')),
    filename = path.join(root, 'test.xlsx'),
    recoveryDirectory = path.join(root, 'recovery'),
    bytes = workbook();
  await fs.writeFile(filename, bytes);
  const workers = [];
  t.after(async () => {
    for (const owner of workers) await owner.close().catch(() => {});
    await fs.rm(root, { recursive: true, force: true });
  });
  async function open() {
    const before = await scopes(),
      owner = new SpreadsheetWorker({ filename, recoveryDirectory });
    workers.push(owner);
    const own = [...(await scopes())].filter((x) => !before.has(x));
    assert.equal(own.length, 1);
    const status = await owner.open();
    return { owner, status, scope: path.join(storage, own[0]) };
  }
  return { root, filename, bytes, open };
}
async function missing(file) {
  await assert.rejects(fs.stat(file), { code: 'ENOENT' });
}
const read = (owner, state) => owner.request({ operation: 'read', state, sheet: 'Sheet1', range: 'A1' });

test('cancel cleans only its own indexes and accepted edits recover while the host stays alive', async (t) => {
  const a = await fixture(t),
    b = await fixture(t),
    first = await a.open(),
    second = await b.open();
  const changed = await first.owner.request({
    operation: 'apply',
    state: first.status.state,
    commands: [{ sheet: 'Sheet1', cell: 'A1', value: 123 }],
  });
  assert.equal(changed.dirty, true);
  await first.owner.cancel();
  await missing(first.scope);
  await assert.rejects(read(first.owner, changed.state), /closed/);
  assert.equal((await read(second.owner, second.status.state)).cells.A1.value, 10);
  const recovered = await a.open();
  assert.equal(recovered.status.recovered, true);
  assert.equal((await read(recovered.owner, recovered.status.state)).cells.A1.value, 123);
  assert.deepEqual(new Uint8Array(await fs.readFile(a.filename)), a.bytes);
  await second.owner.close();
  await missing(second.scope);
});

test('unexpected worker termination rejects a pending call and close awaits automatic cleanup', async (t) => {
  const f = await fixture(t),
    { owner, status, scope } = await f.open();
  const original = Worker.prototype.postMessage;
  Worker.prototype.postMessage = function (message, ...args) {
    if (message.request?.simulateWorkerExit) {
      void this.terminate();
      return;
    }
    return original.call(this, message, ...args);
  };
  try {
    await assert.rejects(owner.request({ operation: 'status', simulateWorkerExit: true }), { code: 'session-closed' });
  } finally {
    Worker.prototype.postMessage = original;
  }
  await owner.close();
  await missing(scope);
  const next = await f.open();
  assert.equal((await read(next.owner, next.status.state)).cells.A1.value, 10);
});

test('native work in flight is cancelled before disposable storage is removed', async (t) => {
  const f = await fixture(t),
    { owner, status, scope } = await f.open();
  const changed = await owner.request({
    operation: 'apply',
    state: status.state,
    commands: [{ sheet: 'Sheet1', cell: 'A1', value: 321 }],
  });
  const outcome = owner
    .request({
      operation: 'apply',
      state: changed.state,
      commands: [{ type: 'formatRange', sheet: 'Sheet1', range: 'A2:A200000', format: { bold: true } }],
    })
    .then(
      (value) => ({ value }),
      (error) => ({ error })
    );
  let observed = false,
    nativePid;
  for (let i = 0; i < 1000 && !observed; i++) {
    for (const child of await fs.readdir(scope))
      if (child.startsWith('owner-')) {
        try {
          const lease = JSON.parse(await fs.readFile(path.join(scope, child, 'native-active.json'), 'utf8'));
          observed = Number.isSafeInteger(lease.pid);
          nativePid = lease.pid;
        } catch (error) {
          if (error.code !== 'ENOENT' && !(error instanceof SyntaxError)) throw error;
        }
      }
    if (!observed) await new Promise((resolve) => setTimeout(resolve, 5));
  }
  assert(observed, 'must witness real native work before cancellation');
  await owner.cancel();
  for (let i = 0; i < 500; i++) {
    try {
      process.kill(nativePid, 0);
    } catch (error) {
      if (error.code === 'ESRCH') {
        nativePid = null;
        break;
      }
      throw error;
    }
    await new Promise((resolve) => setTimeout(resolve, 10));
  }
  assert.equal(nativePid, null, 'terminated worker must not leave its native helper running');
  const result = await outcome;
  assert(result.error, 'pending caller must learn that its owner was cancelled');
  await missing(scope);
  assert.deepEqual(new Uint8Array(await fs.readFile(f.filename)), f.bytes);
  const recovered = await f.open();
  assert.equal((await read(recovered.owner, recovered.status.state)).cells.A1.value, 321);
});

test('early cancellation, oversized/cloning failures and bounded queues do not leak owners', async (t) => {
  const f = await fixture(t),
    before = await scopes();
  const early = new SpreadsheetWorker({ filename: f.filename, recoveryDirectory: path.join(f.root, 'early') });
  const opening = early.open().catch((e) => e);
  await early.cancel();
  assert((await opening) instanceof Error);
  assert.deepEqual(await scopes(), before);
  const { owner, status, scope } = await f.open();
  await assert.rejects(
    owner.request({ operation: 'query', query: 'x'.repeat(2 * 1024 * 1024), state: status.state, sheet: 'Sheet1' }),
    { code: /^(csv-limit|session-limit)$/ }
  );
  await assert.rejects(owner.request({ operation: 'status', uncloneable: () => {} }), { name: 'DataCloneError' });
  const results = await Promise.allSettled(Array.from({ length: 64 }, () => read(owner, status.state)));
  assert(results.some((r) => r.status === 'rejected' && /queue is full/.test(r.reason.message)));
  assert(results.some((r) => r.status === 'fulfilled' && r.value.cells.A1.value === 10));
  assert.equal((await read(owner, status.state)).cells.A1.value, 10);
  await owner.close();
  await missing(scope);
});

test('storage cleanup rejects a foreign token and never follows symlink directories', async (t) => {
  const lease = allocateWorkerStorage(),
    outside = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-worker-outside-'));
  t.after(async () => {
    await fs.rm(lease.root, { recursive: true, force: true });
    await fs.rm(outside, { recursive: true, force: true });
  });
  const marker = path.join(outside, 'keep');
  await fs.writeFile(marker, 'safe');
  await assert.rejects(releaseWorkerStorage({ ...lease, token: 'foreign' }), { code: 'session-refused' });
  await fs.symlink(outside, path.join(lease.root, 'owner-link'), 'dir');
  await releaseWorkerStorage(lease);
  assert.equal(await fs.readFile(marker, 'utf8'), 'safe');
});

// Abrupt host death exercises stale-process collection, not graceful close.
test('a new worker reclaims a dead host scope and recovers its accepted journal', async (t) => {
  const f = await fixture(t),
    before = await scopes();
  const source = `import {SpreadsheetWorker} from ${JSON.stringify(new URL('../adapters/worker-client.mjs', import.meta.url).href)};
const owner=new SpreadsheetWorker({filename:process.argv[1],recoveryDirectory:process.argv[2]});
const status=await owner.open();await owner.request({operation:'apply',state:status.state,commands:[{sheet:'Sheet1',cell:'A1',value:678}]});
process.exit(0);`;
  const child = spawn(
    process.execPath,
    ['--input-type=module', '-e', source, f.filename, path.join(f.root, 'recovery')],
    { stdio: ['ignore', 'ignore', 'pipe'] }
  );
  let errors = '';
  child.stderr.on('data', (c) => (errors += c));
  const code = await new Promise((resolve, reject) => {
    child.once('exit', resolve);
    child.once('error', reject);
  });
  assert.equal(code, 0, errors);
  const abandoned = [...(await scopes(child.pid))];
  assert.equal(abandoned.length, 1);
  const recovered = await f.open();
  assert.equal((await read(recovered.owner, recovered.status.state)).cells.A1.value, 678);
  await missing(path.join(storage, abandoned[0]));
  assert.deepEqual(new Uint8Array(await fs.readFile(f.filename)), f.bytes);
});

test('an unresponsive close handshake has a deadline and still reaps its native resources', async (t) => {
  const f = await fixture(t),
    { owner, scope } = await f.open(),
    original = Worker.prototype.postMessage;
  Worker.prototype.postMessage = function (message, ...args) {
    if (message.operation === 'close') return;
    return original.call(this, message, ...args);
  };
  const started = performance.now();
  try {
    await assert.rejects(owner.close(), { code: 'operation-timed-out' });
  } finally {
    Worker.prototype.postMessage = original;
  }
  assert(performance.now() - started < 8000, 'close must not wait indefinitely for the worker');
  await missing(scope);
  assert.deepEqual(new Uint8Array(await fs.readFile(f.filename)), f.bytes);
});
