/** Real host termination at observed Save boundaries; no manifest reconstruction. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createHash } from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath } from 'node:url';
import { workbook } from '../fixtures/standalone.mjs';
import { createWorkbook } from '../core/workbook-package.mjs';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { nativeWorkbookOperationAsync, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { checkXlsxPackage } from '../scripts/test/package-oracle.mjs';

after(() => closeNativeWorkbookWorker());
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const childScript = fileURLToPath(new URL('./support/document-save-crash-child.cjs', import.meta.url));
const register = fileURLToPath(new URL('../scripts/ts-resolve-register.mjs', import.meta.url));
const read = (owner) =>
  owner.request({ operation: 'read', state: owner.status().state, sheet: 'Sheet1', range: 'A1:C1' });

async function fixture(t) {
  const root = await fs.realpath(await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-save-crash-')));
  const filename = path.join(root, 'document.xlsx'),
    baseline = path.join(root, 'original.xlsx');
  const bytes = workbook(),
    originalHash = hash(bytes);
  await fs.writeFile(filename, bytes);
  await fs.writeFile(baseline, bytes);
  const recoveryDirectory = path.join(root, 'recovery'),
    indexDirectory = path.join(root, 'indexes');
  const owners = [],
    children = [];
  t.after(async () => {
    for (const child of children) {
      if (child.child.exitCode === null && child.child.signalCode === null) child.child.kill('SIGKILL');
      await child.exited;
    }
    for (const owner of owners) await owner.close();
    await fs.rm(root, { recursive: true, force: true });
  });
  async function open() {
    const owner = new SpreadsheetDocument({
      recoveryDirectory,
      indexDirectory,
      native: (input) => nativeWorkbookOperationAsync(input),
    });
    owners.push(owner);
    await owner.open(filename);
    return owner;
  }
  return { root, filename, baseline, bytes, originalHash, recoveryDirectory, indexDirectory, owners, children, open };
}

function runChild(f, boundary, action = 'pause', errorCode) {
  const config = {
    filename: f.filename,
    recoveryDirectory: f.recoveryDirectory,
    indexDirectory: f.indexDirectory,
    boundary,
    action,
    errorCode,
  };
  const child = spawn(process.execPath, ['--import', register, childScript, JSON.stringify(config)], {
    stdio: ['ignore', 'ignore', 'pipe', 'ipc'],
  });
  let stderr = '',
    ended = false;
  const messages = [],
    pending = [];
  child.stderr.on('data', (bytes) => {
    stderr += bytes;
  });
  child.on('message', (message) => {
    if (message.type === 'fatal') {
      for (const item of pending.splice(0)) item.reject(Error(message.error));
      return;
    }
    const index = pending.findIndex((item) => item.type === message.type);
    if (index >= 0) pending.splice(index, 1)[0].resolve(message);
    else messages.push(message);
  });
  child.on('error', (error) => {
    for (const item of pending.splice(0)) item.reject(error);
  });
  const exited = new Promise((resolve) =>
    child.once('exit', (code, signal) => {
      ended = true;
      for (const item of pending.splice(0)) item.reject(Error(`Save test host exited (${code}/${signal}): ${stderr}`));
      resolve({ code, signal });
    })
  );
  function next(type) {
    const index = messages.findIndex((message) => message.type === type);
    if (index >= 0) return Promise.resolve(messages.splice(index, 1)[0]);
    if (ended) return Promise.reject(Error(`Save test host already exited: ${stderr}`));
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        child.kill('SIGKILL');
        reject(Error(`Save boundary ${boundary}/${type} timed out: ${stderr}`));
      }, 15000);
      pending.push({
        type,
        resolve: (message) => {
          clearTimeout(timer);
          resolve(message);
        },
        reject: (error) => {
          clearTimeout(timer);
          reject(error);
        },
      });
    });
  }
  const result = { child, next, exited };
  f.children.push(result);
  return result;
}

async function journal(f) {
  const bucket = path.join(f.recoveryDirectory, hash(f.filename));
  const manifest = JSON.parse(await fs.readFile(path.join(bucket, 'current.json'), 'utf8'));
  assert.equal(hash(await fs.readFile(path.join(bucket, manifest.snapshot + '.xlsx'))), f.originalHash);
  let head = manifest.head,
    expected = manifest.headHash,
    count = 0;
  while (head !== null) {
    const bytes = await fs.readFile(path.join(bucket, head + '.json'));
    assert.equal(hash(bytes), expected);
    const node = JSON.parse(bytes);
    assert(Array.isArray(JSON.parse(node.commands)));
    head = node.parent;
    expected = node.parentHash;
    count++;
  }
  assert.equal(count, 2, 'Both accepted edits must survive in the actual journal');
  return manifest;
}

async function checkCandidate(f, checkpoint) {
  const bytes = await fs.readFile(checkpoint.exported);
  assert.equal(hash(bytes), checkpoint.candidateHash);
  assert.notEqual(checkpoint.candidateHash, f.originalHash);
  const candidate = path.join(f.root, 'candidate.xlsx');
  await fs.writeFile(candidate, bytes);
  const independent = checkXlsxPackage(f.baseline, candidate, {
    targets: { Sheet1: ['A1', 'B1'] },
    expected: {
      'Sheet1!A1': { value: 7 },
      'Sheet1!B1': { value: 4 },
      'Sheet1!C1': { formula: 'A1+B1', calculated: 11 },
    },
  });
  assert.equal(independent.passed, true);
}

async function checkRecovery(f, saved, expectedState) {
  const owner = await f.open(),
    status = owner.status();
  assert.equal(status.dirty, !saved);
  assert.equal(status.recovered, !saved);
  let cells = (await read(owner)).cells;
  assert.equal(cells.A1.value, 7);
  assert.equal(cells.B1.value, 4);
  assert.equal(cells.C1.calculation.value, 11);
  if (saved) {
    // Current reopen semantics start a clean base after publication; persisted
    // undo history is a different product feature, not a crash-test assumption.
    assert.equal(status.state, 'base');
    assert.equal(status.canUndo, false);
    const bucket = path.join(f.recoveryDirectory, hash(f.filename));
    const record = JSON.parse(await fs.readFile(path.join(bucket, 'current.json'), 'utf8'));
    assert.equal(record.pendingPublication, undefined);
    assert.equal(record.head, null);
    assert.equal(record.sourceRevision, hash(await fs.readFile(f.filename)));
    return;
  }
  assert.equal(status.state, expectedState);
  assert.equal(status.canUndo, true);
  await owner.request({ operation: 'undo' });
  cells = (await read(owner)).cells;
  assert.equal(cells.A1.value, 7);
  assert.equal(cells.B1.value, 20);
  assert.equal(cells.C1.calculation.value, 27);
  await owner.request({ operation: 'undo' });
  cells = (await read(owner)).cells;
  assert.equal(cells.A1.value, 10);
  assert.equal(cells.C1.calculation.value, 30);
  await owner.request({ operation: 'redo' });
  await owner.request({ operation: 'redo' });
  assert.equal(owner.status().dirty, !saved);
  assert.equal((await read(owner)).cells.C1.calculation.value, 11);
  assert.equal((await journal(f)).pendingPublication, undefined);
}

// Directory fsync is intentionally a no-op on Windows in production storage.
const boundaries = [
  ['export-complete', false],
  ['stage-file-sync', false],
  ['pending-manifest-file-sync', false],
  ['pending-receipt-sync', false],
  ['source-rename', true],
  ['publication-directory-sync', true],
  ['final-manifest-file-sync', true],
  ['final-manifest-sync', true],
];
for (const [boundary, saved] of boundaries)
  test(
    `killed host at ${boundary} ${saved ? 'reconciles a complete published workbook' : 'preserves source and recovers accepted undo journal'}`,
    {
      skip:
        process.platform === 'win32' &&
        ['pending-receipt-sync', 'publication-directory-sync', 'final-manifest-sync'].includes(boundary),
    },
    async (t) => {
      const f = await fixture(t),
        host = runChild(f, boundary),
        checkpoint = await host.next('checkpoint');
      assert.equal(checkpoint.boundary, boundary);
      assert.equal(checkpoint.accepted.dirty, true);
      assert.equal(checkpoint.accepted.canUndo, true);
      await checkCandidate(f, checkpoint);
      const manifest = await journal(f);
      assert.equal(manifest.head, checkpoint.accepted.state);
      if (
        ['pending-receipt-sync', 'source-rename', 'publication-directory-sync', 'final-manifest-file-sync'].includes(
          boundary
        )
      ) {
        const { savedSource, destinationIdentity, ...receipt } = manifest.pendingPublication;
        assert.deepEqual(receipt, { head: checkpoint.accepted.state, revision: checkpoint.candidateHash });
        assert.equal(savedSource.schema, 1);
        assert.equal(savedSource.head, checkpoint.accepted.state);
        assert.equal(savedSource.headHash, manifest.headHash);
        assert.equal(savedSource.revision, checkpoint.candidateHash);
        assert.match(destinationIdentity.dev, /^(?:0|[1-9]\d*)$/);
        assert.match(destinationIdentity.ino, /^(?:0|[1-9]\d*)$/);
        const savedBytes = await fs.readFile(
          path.join(f.recoveryDirectory, hash(f.filename), savedSource.snapshot + '.xlsx')
        );
        assert.equal(savedSource.bytes, savedBytes.length);
        assert.equal(hash(savedBytes), checkpoint.candidateHash);
      } else assert.equal(manifest.pendingPublication, undefined);
      assert.equal(manifest.savedHead, boundary === 'final-manifest-sync' ? checkpoint.accepted.state : null);
      assert.equal(checkpoint.sourceHash, saved ? checkpoint.candidateHash : f.originalHash);
      assert.equal(hash(await fs.readFile(f.filename)), saved ? checkpoint.candidateHash : f.originalHash);
      host.child.kill('SIGKILL');
      const exit = await host.exited;
      assert.equal(exit.signal, 'SIGKILL');
      await checkRecovery(f, saved, checkpoint.accepted.state);
      assert.equal(hash(await fs.readFile(f.filename)), saved ? checkpoint.candidateHash : f.originalHash);
      assert.equal(hash(await fs.readFile(f.baseline)), f.originalHash);
    }
  );

test(
  'external replacement after pending receipt refuses before destination rename and retains recoverable edits',
  { skip: process.platform === 'win32' },
  async (t) => {
    const f = await fixture(t),
      host = runChild(f, 'pending-receipt-sync'),
      checkpoint = await host.next('checkpoint');
    await checkCandidate(f, checkpoint);
    const external = createWorkbook(),
      externalHash = hash(external);
    await fs.writeFile(f.filename, external);
    host.child.send({ operation: 'continue' });
    const complete = await host.next('complete');
    assert.match(complete.error.message, /changed on disk/);
    assert.equal(complete.dirty, true);
    assert.equal(hash(await fs.readFile(f.filename)), externalHash);
    assert.equal((await host.exited).code, 0);
    await checkRecovery(f, false, checkpoint.accepted.state);
    assert.equal(hash(await fs.readFile(f.filename)), externalHash);
  }
);

for (const [boundary, errorCode] of [
  ['before-stage-copy', 'ENOSPC'],
  ['stage-file-sync', 'ENOSPC'],
  ['pending-manifest-file-sync', 'ENOSPC'],
  ['before-source-rename', 'EACCES'],
]) {
  test(`${errorCode} at ${boundary} leaves the source and accepted recovery unchanged`, async (t) => {
    const f = await fixture(t),
      host = runChild(f, boundary, 'fail', errorCode),
      checkpoint = await host.next('checkpoint');
    await checkCandidate(f, checkpoint);
    host.child.send({ operation: 'continue' });
    const complete = await host.next('complete');
    assert.equal(complete.hit, true);
    assert.equal(complete.error.code, errorCode);
    assert.equal(complete.dirty, true);
    assert.equal(complete.canUndo, true);
    assert.equal((await host.exited).code, 0);
    assert.equal(hash(await fs.readFile(f.filename)), f.originalHash);
    await assert.rejects(fs.stat(checkpoint.staged), { code: 'ENOENT' });
    await checkRecovery(f, false, checkpoint.accepted.state);
    assert.equal(hash(await fs.readFile(f.filename)), f.originalHash);
  });
}
