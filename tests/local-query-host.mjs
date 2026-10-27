import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { LocalQueryConnections } from '../adapters/local-query-connections.mjs';
import { LocalQueryHost } from '../adapters/local-query-host.mjs';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
const draft = {
  name: 'Active',
  target: { sheet: 'Sheet1', start: 'A1' },
  query: {
    columns: [1, 3],
    filters: [{ column: 2, op: 'equals', value: 'active', caseSensitive: true }],
    keepHeader: true,
  },
};
async function fixture(t) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-query-owner-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const source = path.join(directory, 'source.csv'),
    recoveryDirectory = path.join(directory, 'recovery'),
    registry = path.join(directory, 'queries.json');
  await fs.writeFile(source, 'name,status,code\na,active,01\nb,inactive,02\nc,active,=1+2\n');
  let owner = new SpreadsheetDocument({ recoveryDirectory, cacheIdentity: null }),
    status = await owner.create('Query');
  t.after(() => owner.close());
  const store = await LocalQueryConnections.open(registry),
    document = status.documentId,
    connection = await store.create(document, draft, source);
  let frozen = status.state,
    loseReply = false;
  const context = {
    current(fenced = true) {
      if (status.documentId !== document || (fenced && status.state !== frozen)) throw Error('Scope changed');
      return { state: status.state, documentId: document };
    },
    request: (operation, input) => owner.request({ documentId: document, ...input, operation }),
    async apply(state, descriptor) {
      status = await owner.request({ operation: 'apply', state, commands: [descriptor] });
      if (loseReply) throw Error('Reply lost after durable acceptance');
      return status;
    },
  };
  const host = () => new LocalQueryHost({ store, document, directory: path.join(directory, 'staging'), context });
  return {
    source,
    registry,
    document,
    connection,
    host,
    store,
    get status() {
      return status;
    },
    read: (range) => owner.request({ operation: 'read', state: status.state, sheet: 'Sheet1', range }),
    async move(operation) {
      status = await owner.request({ operation, state: status.state });
      frozen = status.state;
    },
    async edit(value) {
      status = await owner.request({
        operation: 'apply',
        state: status.state,
        commands: [{ sheet: 'Sheet1', cell: 'B2', value }],
      });
      frozen = status.state;
    },
    freeze() {
      frozen = status.state;
    },
    lose() {
      loseReply = true;
    },
    async reopen() {
      await owner.close();
      owner = new SpreadsheetDocument({ recoveryDirectory, cacheIdentity: null });
      status = await owner.recover(document);
      frozen = status.state;
    },
  };
}
test('saved local query selects/filters literal fields, clears old tail, and is exactly one Undo/Redo/recovery action', async (t) => {
  const f = await fixture(t),
    preview = await f.host().preview(f.connection.id, 1);
  assert.deepEqual(preview.queryPreview.rows, [
    ['name', 'code'],
    ['a', '01'],
    ['c', '=1+2'],
  ]);
  let result = await f.host().refresh(f.connection.id, 1);
  assert.equal(result.refreshed, true);
  assert.equal(result.connection.revision, 2);
  assert.equal((await f.read('A1:B3')).cells.B3.value, '=1+2');
  f.freeze();
  await fs.writeFile(f.source, 'name,status,code\na,active,004\n');
  result = await f.host().refresh(f.connection.id, 2);
  assert.equal(result.outputRows, 2);
  let cells = (await f.read('A1:B3')).cells;
  assert.equal(cells.B2.value, '004');
  assert.equal(cells.B3?.value ?? '', '');
  await f.move('undo');
  assert.equal((await f.read('A1:B3')).cells.B3.value, '=1+2');
  await f.reopen();
  assert.equal((await f.read('A1:B3')).cells.B3.value, '=1+2');
  await f.move('redo');
  assert.equal((await f.read('A1:B3')).cells.B2.value, '004');
});
test('manual destination edit and source change during capture refuse without accepted state loss', async (t) => {
  const f = await fixture(t);
  await f.host().refresh(f.connection.id, 1);
  await f.edit('manual');
  const before = f.status.state;
  await assert.rejects(f.host().refresh(f.connection.id, 2), { code: /^(layout-stale|pivot-stale|query-stale)$/ });
  assert.equal(f.status.state, before);
  assert.equal((await f.read('B2')).cells.B2.value, 'manual');
  await f.move('undo');
  f.freeze();
  let changed = false;
  await assert.rejects(
    f.host().preview(f.connection.id, 2, {
      onProgress: (event) => {
        if (!changed && event.phase === 'copying-file') {
          changed = true;
          return fs.writeFile(f.source, 'changed');
        }
      },
    }),
    /changed/
  );
  assert.equal((await f.read('B2')).cells.B2.value, '01');
});
test('lost accepted reply keeps intent/attachment; restart and retained Redo receipt reconcile without another import', async (t) => {
  const f = await fixture(t);
  f.lose();
  await assert.rejects(f.host().refresh(f.connection.id, 1), /Reply lost/);
  const accepted = f.status.state;
  assert.equal((await f.store.list(f.document)).entries[0].pending, true);
  await f.move('undo');
  await f.reopen();
  const list = await f.host().list();
  assert.equal(list.entries[0].pending, undefined);
  assert.equal(list.entries[0].last.acceptedHead, accepted);
  await f.move('redo');
  assert.equal((await f.read('B2')).cells.B2.value, '01');
});
test('cancellation after durable intent but before dispatch rolls back only the unsubmitted intent', async (t) => {
  const f = await fixture(t),
    controller = new AbortController(),
    before = f.status.state;
  await assert.rejects(
    f.host().refresh(f.connection.id, 1, {
      signal: controller.signal,
      authorizeApply: async () => {
        controller.abort(Error('Cancelled before owner dispatch'));
      },
    }),
    /Cancelled/
  );
  assert.equal(f.status.state, before);
  const list = await f.store.list(f.document);
  assert.equal(list.entries[0].pending, undefined);
  assert.equal(list.entries[0].revision, 1);
  assert.deepEqual((await f.read('A1:B3')).cells, {});
});
