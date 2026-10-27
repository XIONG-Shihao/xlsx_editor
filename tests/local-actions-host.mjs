import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import {
  localActionRequest,
  invokeLocalActionRequest,
  LOCAL_ACTION_REQUEST_BYTES,
} from '../adapters/local-actions-host.mjs';
import { LocalActions } from '../adapters/local-actions.mjs';

const documentId = '84ebd540-35ef-4c60-93d0-fecdd8b59734',
  actionId = 'cbdfd441-8f29-40fd-8246-18e49d1ab4d1',
  head = '913c34df-4dc1-4839-9892-4aa5cc8e7a50',
  hash = 'ab'.repeat(32);
const binding = { documentId, state: 'base' };
const draft = () => ({
  name: 'Format selected cells',
  descriptor: {
    schema: 1,
    command: 'formatRange',
    parameters: { sheet: 'Data', range: '$a$1:$b$3', format: { bold: true, fontColor: '153C6D' } },
  },
});
const create = () => ({ operation: 'create', ...binding, expectedRevision: 0, draft: draft() });

test('canonical request keeps actual owner binding and independently freezes mutable nested draft', () => {
  const body = create(),
    captured = localActionRequest(body);
  assert.equal(captured.documentId, documentId);
  assert.equal(captured.state, 'base');
  assert.equal(captured.draft.descriptor.parameters.range, 'A1:B3');
  assert.equal(captured.expectedRevision, 0);
  body.draft.name = 'Retargeted';
  body.draft.descriptor.parameters.sheet = 'Other';
  body.draft.descriptor.parameters.format.bold = false;
  assert.equal(captured.draft.name, 'Format selected cells');
  assert.equal(captured.draft.descriptor.parameters.sheet, 'Data');
  assert.equal(captured.draft.descriptor.parameters.format.bold, true);
  assert(
    Object.isFrozen(captured) &&
      Object.isFrozen(captured.draft) &&
      Object.isFrozen(captured.draft.descriptor.parameters.format)
  );
  assert.throws(() => {
    captured.draft.descriptor.parameters.format.bold = false;
  }, TypeError);
  const sort = {
    operation: 'create',
    ...binding,
    expectedRevision: 1,
    draft: {
      name: 'Mixed sort',
      descriptor: {
        schema: 1,
        command: 'sortRows',
        parameters: {
          sheet: 'Data',
          range: 'A2:C4',
          keys: [
            { column: 'b', direction: 'desc' },
            { column: 'a', direction: 'asc' },
          ],
        },
      },
    },
  };
  const result = localActionRequest(sort);
  sort.draft.descriptor.parameters.keys[0].direction = 'asc';
  assert.deepEqual(result.draft.descriptor.parameters.keys, [
    { column: 'B', direction: 'desc' },
    { column: 'A', direction: 'asc' },
  ]);
  assert(Object.isFrozen(result.draft.descriptor.parameters.keys[0]));
});

test('each operation has an exact schema and safe numeric CAS revision', () => {
  const valid = [
    { operation: 'list', ...binding },
    create(),
    { operation: 'update', ...binding, expectedRevision: 2, actionId, draft: draft() },
    { operation: 'remove', ...binding, expectedRevision: 3, actionId },
  ];
  for (const body of valid) {
    assert.equal(localActionRequest(body).operation, body.operation);
    for (const key of Object.keys(body)) {
      const bad = { ...body };
      delete bad[key];
      assert.throws(() => localActionRequest(bad));
    }
  }
  for (const body of valid)
    for (const extra of [
      'actor',
      'author',
      'grants',
      'filename',
      'path',
      'version',
      'schema',
      'run',
      'expectedRegistryRevision',
    ])
      for (const value of ['forged', undefined])
        assert.throws(() => localActionRequest({ ...body, [extra]: value }), /fields|data/);
  assert.throws(() => localActionRequest({ ...valid[0], expectedRevision: 0 }), /fields/);
  assert.throws(() => localActionRequest({ ...valid[3], draft: draft() }), /fields/);
  assert.throws(() => localActionRequest({ ...create(), actionId }), /fields/);
  for (const revision of [-1, 1.5, Number.MAX_SAFE_INTEGER + 1, '0', false, null])
    assert.throws(() => localActionRequest({ ...create(), expectedRevision: revision }), /revision/);
  for (const id of [
    actionId.toUpperCase(),
    actionId.replace('-40fd-', '-70fd-'),
    'not-an-id',
    '/private/tmp/actions.json',
  ])
    assert.throws(() => localActionRequest({ ...valid[3], actionId: id }));
});

test('document/state admit only real canonical UUIDv4 or SHA256 heads and base', () => {
  for (const documentId of [binding.documentId, hash])
    for (const state of ['base', head, hash])
      assert.deepEqual(localActionRequest({ operation: 'list', documentId, state }), {
        operation: 'list',
        documentId,
        state,
      });
  for (const documentId of ['base', 'owner', hash.toUpperCase(), binding.documentId.toUpperCase(), '/tmp/book.xlsx'])
    assert.throws(() => localActionRequest({ operation: 'list', documentId, state: 'base' }), {
      code: /^(internal-invalid-field|query-malformed|session-stale)$/,
    });
  for (const state of [
    'current',
    '1',
    head.toUpperCase(),
    head.replace('-4839-', '-7839-'),
    hash.slice(1),
    '../source.xlsx',
  ])
    assert.throws(() => localActionRequest({ operation: 'list', ...binding, state }), /state/);
});

test('prototype aliases, getters, symbols, toJSON and sparse/custom arrays never enter dispatch', () => {
  let accessed = 0;
  const getter = create();
  Object.defineProperty(getter, 'draft', {
    enumerable: true,
    get() {
      accessed++;
      return draft();
    },
  });
  assert.throws(() => localActionRequest(getter), { code: 'query-refused' });
  assert.equal(accessed, 0);
  const inherited = Object.create({ operation: 'list' });
  Object.assign(inherited, binding);
  assert.throws(() => localActionRequest(inherited), /request/);
  const symbol = create();
  symbol[Symbol('actor')] = 'admin';
  assert.throws(() => localActionRequest(symbol), { code: 'query-refused' });
  const hidden = create();
  Object.defineProperty(hidden, 'actor', { value: 'admin' });
  assert.throws(() => localActionRequest(hidden), { code: /^(chart-unsupported|query-refused)$/ });
  const serializer = create();
  serializer.toJSON = () => {
    accessed++;
    return create();
  };
  assert.throws(() => localActionRequest(serializer), /data/);
  assert.equal(accessed, 0);
  const poison = JSON.parse(
    '{"operation":"list","documentId":"' + documentId + '","state":"base","__proto__":{"actor":"admin"}}'
  );
  assert.throws(() => localActionRequest(poison), /fields/);
  assert.equal(Object.prototype.actor, undefined);
  for (const operation of ['run', 'import', 'execute', 'constructor', 'toString', '__proto__'])
    assert.throws(() => localActionRequest({ operation, ...binding }), /fields/);
  const sort = {
    operation: 'create',
    ...binding,
    expectedRevision: 0,
    draft: {
      name: 'Sort',
      descriptor: { schema: 1, command: 'sortRows', parameters: { sheet: 'Data', range: 'A1:B2', keys: new Array(2) } },
    },
  };
  assert.throws(() => localActionRequest(sort), { code: /^(query-bad-input|query-stale)$/ });
  sort.draft.descriptor.parameters.keys = [{ column: 'A', direction: 'asc' }];
  Object.setPrototypeOf(sort.draft.descriptor.parameters.keys, { run() {} });
  assert.throws(() => localActionRequest(sort), { code: 'internal-invalid-field' });
  const nullPrototype = Object.assign(Object.create(null), { operation: 'list', ...binding });
  assert.deepEqual(localActionRequest(nullPrototype), { operation: 'list', ...binding });
});

test('16KiB budget measures escaped UTF8 input before draft normalization', () => {
  assert.equal(LOCAL_ACTION_REQUEST_BYTES, 16384);
  const oversized = create();
  oversized.draft.name = '🙂'.repeat(4096);
  assert(JSON.stringify(oversized).length < 16384);
  assert(Buffer.byteLength(JSON.stringify(oversized), 'utf8') > 16384);
  assert.throws(() => localActionRequest(oversized), { code: 'query-limit' });
  const escaped = create();
  escaped.draft.name = '\n'.repeat(8192);
  assert.throws(() => localActionRequest(escaped), { code: 'query-limit' });
  const below = create();
  below.draft.name = 'a'.repeat(200);
  assert(Buffer.byteLength(JSON.stringify(below)) < 16384);
  assert.throws(() => localActionRequest(below), { code: 'internal-invalid-field' });
  const cycle = create();
  cycle.draft.descriptor.parameters.format.self = cycle;
  assert.throws(() => localActionRequest(cycle), { code: 'query-bad-input' });
  assert.throws(() => localActionRequest({ ...create(), expectedRevision: NaN }), { code: 'query-bad-input' });
});

test('dispatcher calls exactly one CRUD method with copied canonical arguments and propagates CAS refusal', async () => {
  const calls = [],
    reply = { schema: 1, scope: 'private-profile', revision: 0, entries: [] };
  const store = {
    list: async (...args) => {
      calls.push(['list', args]);
      return reply;
    },
    create: async (...args) => {
      calls.push(['create', args]);
      return reply;
    },
    update: async (...args) => {
      calls.push(['update', args]);
      return reply;
    },
    remove: async (...args) => {
      calls.push(['remove', args]);
      throw Error('Stale local action registry revision');
    },
  };
  for (const name of ['run', 'apply', 'native', 'import', 'openWorkbook', 'shell'])
    Object.defineProperty(store, name, {
      get() {
        throw Error('Forbidden execution lookup');
      },
    });
  const body = create(),
    pending = invokeLocalActionRequest(store, body);
  body.draft.descriptor.parameters.format.bold = false;
  assert.equal(await pending, reply);
  assert.equal(calls[0][1][0].descriptor.parameters.format.bold, true);
  assert.equal(calls[0][1][1], 0);
  await invokeLocalActionRequest(store, { operation: 'list', ...binding });
  await invokeLocalActionRequest(store, {
    operation: 'update',
    ...binding,
    expectedRevision: 2,
    actionId,
    draft: draft(),
  });
  await assert.rejects(
    invokeLocalActionRequest(store, { operation: 'remove', ...binding, expectedRevision: 3, actionId }),
    /Stale/
  );
  assert.deepEqual(
    calls.map((c) => c[0]),
    ['create', 'list', 'update', 'remove']
  );
  assert.deepEqual(calls[1][1], []);
  assert.equal(calls[2][1][0], actionId);
  assert.equal(calls[2][1][2], 2);
  assert.deepEqual(calls[3][1], [actionId, 3]);
  const count = calls.length;
  assert.throws(() => invokeLocalActionRequest(store, { operation: 'run', ...binding }), /fields/);
  assert.equal(calls.length, count);
});

test('actual private store persists CRUD/CAS independently of workbook/source and reopens cold', async (t) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'workbook-local-actions-host-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const filename = path.join(directory, 'actions.json'),
    workbook = path.join(directory, 'source.xlsx'),
    original = Buffer.from('opaque independently authored workbook sentinel');
  await fs.writeFile(workbook, original);
  const store = await LocalActions.open(filename);
  const list = await invokeLocalActionRequest(store, { operation: 'list', ...binding });
  assert.deepEqual(list, { schema: 1, scope: 'private-profile', revision: 0, entries: [] });
  assert.equal(await fs.stat(filename).catch(() => null), null);
  const made = await invokeLocalActionRequest(store, create()),
    id = made.entries[0].id;
  assert.equal(made.revision, 1);
  assert.equal(made.entries[0].author, 'local-profile');
  assert.equal(made.entries[0].descriptor.parameters.range, 'A1:B3');
  const before = await fs.readFile(filename);
  await assert.rejects(invokeLocalActionRequest(store, create()), /Stale/);
  assert.deepEqual(await fs.readFile(filename), before);
  const updated = await invokeLocalActionRequest(store, {
    operation: 'update',
    ...binding,
    state: head,
    expectedRevision: 1,
    actionId: id,
    draft: { name: 'Recalculate definition', descriptor: { schema: 1, command: 'recalculate', parameters: {} } },
  });
  assert.equal(updated.revision, 2);
  assert.equal(updated.entries[0].revision, 2);
  assert.deepEqual(
    await invokeLocalActionRequest(await LocalActions.open(filename), { operation: 'list', ...binding, state: hash }),
    updated
  );
  const removed = await invokeLocalActionRequest(store, {
    operation: 'remove',
    ...binding,
    expectedRevision: 2,
    actionId: id,
  });
  assert.deepEqual(removed, { schema: 1, scope: 'private-profile', revision: 3, entries: [] });
  assert.deepEqual(
    await invokeLocalActionRequest(await LocalActions.open(filename), { operation: 'list', ...binding }),
    removed
  );
  assert.deepEqual(await fs.readFile(workbook), original);
  assert.deepEqual((await fs.readdir(directory)).sort(), ['actions.json', 'source.xlsx']);
});
