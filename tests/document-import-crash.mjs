import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { createWorkbook } from '../core/workbook-package.mjs';
import {
  nativeWorkbookBinary,
  nativeWorkbookOperationAsync,
  closeNativeWorkbookWorker,
} from '../core/native-runtime.mjs';
import { revision } from '../adapters/storage.mjs';

const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const sourceHelper = nativeWorkbookBinary(),
  helperBytes = await fs.readFile(sourceHelper),
  helperDirectory = await fs.realpath(await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-import-crash-helper-'))),
  helper = path.join(helperDirectory, path.basename(sourceHelper));
await fs.writeFile(helper, helperBytes, { mode: 0o700 });
const previous = process.env.XLSX_EDITOR_NATIVE_BIN;
process.env.XLSX_EDITOR_NATIVE_BIN = helper;
after(async () => {
  closeNativeWorkbookWorker();
  if (previous === undefined) delete process.env.XLSX_EDITOR_NATIVE_BIN;
  else process.env.XLSX_EDITOR_NATIVE_BIN = previous;
  await fs.rm(helperDirectory, { recursive: true, force: true });
});
const childScript = fileURLToPath(new URL('./support/document-import-crash-child.cjs', import.meta.url)),
  register = fileURLToPath(new URL('../scripts/ts-resolve-register.mjs', import.meta.url));

async function fixture(t) {
  const root = await fs.realpath(await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-import-crash-'))),
    filename = path.join(root, 'source.xlsx'),
    recoveryDirectory = path.join(root, 'recovery'),
    indexDirectory = path.join(root, 'indexes'),
    bytes = createWorkbook(),
    owners = [],
    children = [];
  await fs.writeFile(filename, bytes);
  const ownerDirectory = path.join(recoveryDirectory, hash(filename));
  t.after(async () => {
    for (const item of children) {
      if (item.child.exitCode === null && item.child.signalCode === null) item.child.kill('SIGKILL');
      await item.exited;
    }
    for (const owner of owners) await owner.close();
    await fs.rm(root, { recursive: true, force: true });
  });
  return { root, filename, recoveryDirectory, indexDirectory, ownerDirectory, bytes, owners, children };
}
function run(f, boundary) {
  const config = {
    filename: f.filename,
    recoveryDirectory: f.recoveryDirectory,
    indexDirectory: f.indexDirectory,
    ownerDirectory: f.ownerDirectory,
    boundary,
  };
  const child = spawn(process.execPath, ['--import', register, childScript, JSON.stringify(config)], {
      stdio: ['ignore', 'ignore', 'pipe', 'ipc'],
    }),
    messages = [],
    waiters = [];
  let ended = false,
    stderr = '';
  child.stderr.on('data', (bytes) => {
    stderr += bytes;
  });
  child.on('message', (message) => {
    if (message.type === 'fatal') {
      for (const waiter of waiters.splice(0)) waiter.reject(Error(message.error));
      return;
    }
    const index = waiters.findIndex((waiter) => waiter.type === message.type);
    if (index >= 0) waiters.splice(index, 1)[0].resolve(message);
    else messages.push(message);
  });
  const exited = new Promise((resolve) =>
    child.once('exit', (code, signal) => {
      ended = true;
      for (const waiter of waiters.splice(0)) waiter.reject(Error(`Import host stopped ${code}/${signal}: ${stderr}`));
      resolve({ code, signal });
    })
  );
  const next = (type) => {
    const index = messages.findIndex((message) => message.type === type);
    if (index >= 0) return Promise.resolve(messages.splice(index, 1)[0]);
    if (ended) return Promise.reject(Error(stderr));
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        child.kill('SIGKILL');
        reject(Error('Import boundary timed out: ' + boundary + ' ' + stderr));
      }, 30000);
      waiters.push({
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
  };
  const handle = { child, next, exited };
  f.children.push(handle);
  return handle;
}
const read = (owner) =>
  owner.request({ operation: 'read', state: owner.status().state, sheet: 'Sheet1', range: 'A1:A2' });
for (const [boundary, published] of [
  ['node-file-sync', false],
  ['node-rename', false],
  ['manifest-file-sync', false],
  ['manifest-rename', true],
  ['manifest-directory-sync', true],
])
  test(
    `killed import host at ${boundary} ${published ? 'replays one complete imported action' : 'preserves the old frame and removes unreferenced upload data'}`,
    { skip: process.platform === 'win32' && boundary === 'manifest-directory-sync' },
    async (t) => {
      const f = await fixture(t),
        host = run(f, boundary),
        checkpoint = await host.next('checkpoint');
      assert.equal(checkpoint.boundary, boundary);
      assert.equal(checkpoint.sourceHash, hash(f.bytes));
      const record = JSON.parse(await fs.readFile(path.join(f.ownerDirectory, 'current.json'), 'utf8'));
      assert.equal(record.head !== null, published);
      const attachment = path.join(f.ownerDirectory, 'paste-v1', checkpoint.descriptor.id + '.paste.jsonl');
      assert.equal(await revision(attachment), checkpoint.descriptor.sha256);
      host.child.kill('SIGKILL');
      assert.equal((await host.exited).signal, 'SIGKILL');
      const owner = new SpreadsheetDocument({
        recoveryDirectory: f.recoveryDirectory,
        indexDirectory: f.indexDirectory,
        native: (input) => nativeWorkbookOperationAsync(input),
      });
      f.owners.push(owner);
      await owner.open(f.filename);
      assert.equal(owner.status().canUndo, published);
      assert.equal(owner.status().dirty, published);
      assert.equal(owner.status().state, published ? record.head : 'base');
      assert.equal(await revision(f.filename), hash(f.bytes));
      if (published) {
        assert.equal((await read(owner)).cells.A1.value, 1);
        assert.equal(await revision(attachment), checkpoint.descriptor.sha256);
        await owner.request({ operation: 'undo' });
        assert.equal((await read(owner)).cells.A1, undefined);
        await owner.request({ operation: 'redo' });
        assert.equal((await read(owner)).cells.A2.value, 1);
      } else {
        assert.equal((await read(owner)).cells.A1, undefined);
        assert.deepEqual(await fs.readdir(path.join(f.ownerDirectory, 'paste-v1')), []);
      }
    }
  );
