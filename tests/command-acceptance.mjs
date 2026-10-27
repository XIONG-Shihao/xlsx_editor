import './support/scratch-dir.mjs';
/** Executes the exported acceptance helper and the actual standalone renderer
 * functions/main request handler. IPC/registry boundaries are deterministic
 * substitutes here; actual Electron and durable provider checks are separate. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createHash } from 'node:crypto';
import ts from 'typescript';
import { transform } from 'esbuild';

const component = fileURLToPath(new URL('../', import.meta.url)),
  prefix = path.join(component, '.build/tests/m4-host-acceptance');
await fs.mkdir(path.dirname(prefix), { recursive: true });
const output = await fs.mkdtemp(prefix + '-');
const sha = (bytes) => createHash('sha256').update(bytes).digest('hex'),
  sources = {};
for (const name of ['ui/commandAcceptance.ts', 'standalone/renderer.tsx', 'standalone/main.mjs'])
  sources[name] = await fs.readFile(path.join(component, name), 'utf8');
const helperFile = path.join(output, 'commandAcceptance.mjs');
await fs.writeFile(
  helperFile,
  (await transform(sources['ui/commandAcceptance.ts'], { loader: 'ts', format: 'esm', target: 'es2022' })).code
);
const helper = await import(pathToFileURL(helperFile));
const deferred = () => {
  let resolve, reject;
  const promise = new Promise((a, b) => {
    resolve = a;
    reject = b;
  });
  return { promise, resolve, reject };
};
function extracted(source, name) {
  const tree = ts.createSourceFile('source.tsx', source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
  function find(node) {
    if (ts.isFunctionDeclaration(node) && node.name?.text === name) return node.getText(tree);
    return ts.forEachChild(node, find);
  }
  const result = find(tree);
  assert(result, 'Missing real source function ' + name);
  return result;
}
function renderer(records) {
  const all = { current: structuredClone(records) },
    activeDocument = { current: records[0].id },
    latest = { current: all.current[0] },
    ownerQueues = { current: new Map() },
    queue = { current: Promise.resolve() },
    pending = { current: 0 },
    calls = [];
  let request = async (id, input) => ({ ...all.current.find((record) => record.id === id), state: 'accepted' }),
    busy = {},
    errors = [];
  const context = {
    ...helper,
    crypto: globalThis.crypto,
    structuredClone,
    Set,
    Map,
    Promise,
    Error,
    console,
    pastePages: function* () {
      yield { sequence: 0, cells: [[0, 'value']] };
    },
    all,
    activeDocument,
    latest,
    ownerQueues,
    queue,
    pending,
    setDocuments() {},
    setActiveId() {},
    setOwnerBusy(value) {
      busy = typeof value === 'function' ? value(busy) : value;
    },
    setBusy() {},
    setError(value) {
      errors.push(value);
    },
    setNotice() {},
    errorMessage: String,
    desktop: {
      request: (id, input) => {
        calls.push({ id, input });
        return request(id, input);
      },
      documents: async () => ({ activeId: activeDocument.current, documents: all.current, limit: 4 }),
    },
  };
  vm.createContext(context);
  const code =
    ['accept', 'update', 'enqueue', 'edit', 'editCommand', 'historyCommand', 'importPaste']
      .map((name) => extracted(sources['standalone/renderer.tsx'], name))
      .join('\n') + '\nglobalThis.routes={accept,update,edit,editCommand,historyCommand,importPaste}';
  vm.runInContext(
    ts.transpileModule(code, { compilerOptions: { target: ts.ScriptTarget.ES2022 } }).outputText,
    context
  );
  return {
    ...context.routes,
    all,
    activeDocument,
    ownerQueues,
    calls,
    errors,
    desktop: context.desktop,
    setRequest: (fn) => (request = fn),
  };
}
const doc = (id = 'A', patch = {}) => ({
  id,
  tabId: 'tab-' + id,
  epoch: 1,
  documentId: 'grant-' + id,
  state: 'base',
  kind: 'xlsx',
  view: { readOnly: false },
  ...patch,
});
function mainRoute(record) {
  let records = [record],
    queued = [];
  const calls = [],
    observed = [];
  const registry = { active: record, contains: (value) => records.includes(value), records: () => records };
  const selected = (id) => {
    const value = records.find((record) => record.id === id);
    if (!value) throw Error('Document changed');
    return value;
  };
  record.owner = {
    closed: false,
    request: async (input) => {
      calls.push(input);
      return { ...record.status, state: 'accepted' };
    },
  };
  const context = {
    Buffer,
    Error,
    Object,
    registry,
    selected,
    closeTransaction: false,
    shell: (work) => {
      const wait = deferred();
      queued.push(() => Promise.resolve().then(work).then(wait.resolve, wait.reject));
      return wait.promise;
    },
    perform: (_label, _record, work) => work({}),
    observe: (operation) => {
      observed.push(operation);
      return {};
    },
    status: (record) =>
      record && {
        id: record.id,
        tabId: record.tabId,
        epoch: record.epoch,
        ...record.status,
        ...(record.operationCancelled ? { operationCancelled: true } : {}),
      },
    validId: () => true,
    keptRecovery: { remove: async () => {} },
    console,
  };
  const tree = ts.createSourceFile(
    'main.mjs',
    sources['standalone/main.mjs'],
    ts.ScriptTarget.Latest,
    true,
    ts.ScriptKind.JS
  );
  let code;
  function find(node) {
    if (
      ts.isCallExpression(node) &&
      node.expression.getText(tree) === 'handle' &&
      node.arguments[0]?.getText(tree) === "'request'"
    )
      code = node.arguments[1].getText(tree);
    ts.forEachChild(node, find);
  }
  find(tree);
  assert(code);
  vm.createContext(context);
  vm.runInContext('globalThis.request=' + code, context);
  return {
    registry,
    calls,
    observed,
    request: context.request,
    run: () => queued.shift()(),
    replace: (value) => {
      records = [value];
      registry.active = value;
    },
  };
}

test('binding snapshot validates fields; every runtime/head retirement refuses before enqueue', async () => {
  const original = {
    projectId: 'p',
    entryId: 'entry',
    documentId: 'grant',
    generation: 2,
    kind: 'rust-indexed-v1',
    sessionId: 'session',
    state: 'head',
  };
  const captured = helper.captureOwnerCommandBinding(original);
  original.state = 'changed';
  assert.equal(captured.state, 'head');
  assert(Object.isFrozen(captured));
  let queued = 0,
    applied = 0;
  for (const [key, value] of [
    ['projectId', 'q'],
    ['entryId', 'other'],
    ['documentId', 'new'],
    ['generation', 3],
    ['kind', 'csv-indexed-v1'],
    ['sessionId', 'other'],
    ['state', 'new'],
    ['active', false],
  ])
    await assert.rejects(
      helper.queueCapturedOwnerCommand(captured, {
        current: () => ({ ...captured, editable: true, [key]: value }),
        enqueue: (work) => {
          queued++;
          return work();
        },
        apply: async () => {
          applied++;
        },
      }),
      { code: 'OWNER_COMMAND_RETIRED' }
    );
  await assert.rejects(
    helper.queueCapturedOwnerCommand(captured, {
      current: () => ({ ...captured, editable: false }),
      enqueue: (work) => {
        queued++;
        return work();
      },
      apply: async () => {
        applied++;
      },
    }),
    /read-only/
  );
  assert.equal(queued, 0);
  assert.equal(applied, 0);
  for (const invalid of [
    { ...captured, generation: -1 },
    { ...captured, generation: 1.5 },
    { ...captured, documentId: '' },
  ])
    assert.throws(() => helper.captureOwnerCommandBinding(invalid));
});
test('exported queue checks the original head at execution; accepted truth survives subsequent scope change', async () => {
  const captured = helper.captureOwnerCommandBinding({
    projectId: 'p',
    entryId: 'e',
    documentId: 'g',
    generation: 0,
    kind: 'k',
    sessionId: 's',
    state: 'h',
  });
  let current = { ...captured, editable: true },
    run,
    calls = 0;
  const queue = deferred();
  const pending = helper.queueCapturedOwnerCommand(captured, {
    current: () => current,
    enqueue: (work) => {
      run = work;
      return queue.promise;
    },
    apply: async () => {
      calls++;
    },
  });
  current = { ...current, state: 'prior-edit' };
  await assert.rejects(run(), { code: 'OWNER_COMMAND_RETIRED' });
  queue.reject(new helper.OwnerCommandRetired());
  await assert.rejects(pending);
  assert.equal(calls, 0);
  current = { ...captured, editable: true };
  const accepted = deferred(),
    working = helper.queueCapturedOwnerCommand(captured, {
      current: () => current,
      enqueue: (work) => work(),
      apply: async () => {
        calls++;
        return accepted.promise;
      },
    });
  current = { ...current, sessionId: 'replacement' };
  accepted.resolve('durably accepted');
  assert.equal(await working, 'durably accepted');
  assert.equal(calls, 1);
});
test('standalone actual legacy FIFO advances heads; a strict batch queued behind it refuses', async () => {
  const host = renderer([doc()]),
    held = deferred();
  host.setRequest(async (id, input) => {
    await held.promise;
    return doc(id, { state: 'point-1' });
  });
  const legacy = host.edit('A', [{ sheet: 'Sheet1', cell: 'A1', value: 1 }]);
  await Promise.resolve();
  const strict = host.editCommand(doc(), [
    { type: 'setCells', sheet: 'Sheet1', cells: [{ cell: 'B1', formula: 'SUM(A1)' }] },
  ]);
  const refused = assert.rejects(strict, { code: 'OWNER_COMMAND_RETIRED' });
  held.resolve();
  await legacy;
  await refused;
  assert.equal(host.calls.length, 1);
  assert.equal(host.all.current[0].state, 'point-1');
  host.setRequest(async (id, input) => doc(id, { state: input.state === 'point-1' ? 'point-2' : 'point-3' }));
  await host.edit('A', [{ sheet: 'Sheet1', cell: 'A2', value: 2 }]);
  await host.edit('A', [{ sheet: 'Sheet1', cell: 'A3', value: 3 }]);
  assert.deepEqual(
    host.calls.map((call) => call.input.state),
    ['base', 'point-1', 'point-2']
  );
  assert(host.calls.every((call) => call.input.documentId === undefined));
});
test('standalone actual strict callback snapshots input and updates original inactive tab after durable reply', async () => {
  const host = renderer([doc(), doc('B')]),
    held = deferred();
  host.setRequest(async () => held.promise);
  const commands = [{ sheet: 'Sheet1', cell: 'B1', value: -0 }],
    pending = host.editCommand(doc(), commands);
  commands[0].cell = 'C1';
  commands[0].value = 100;
  await Promise.resolve();
  assert.equal(host.calls.length, 1);
  assert.equal(host.calls[0].input.documentId, 'grant-A');
  assert.equal(host.calls[0].input.commands[0].cell, 'B1');
  assert(Object.is(host.calls[0].input.commands[0].value, -0));
  host.activeDocument.current = 'B';
  held.resolve(doc('A', { state: 'accepted-A' }));
  await pending;
  assert.equal(host.all.current.find((record) => record.id === 'A').state, 'accepted-A');
  assert.equal(host.all.current.find((record) => record.id === 'B').state, 'base');
  assert.equal(host.activeDocument.current, 'B');
  assert.equal(host.calls.length, 1);
});
test('standalone cancellation adopts recovered replacement then rejects success; stale epoch cannot overwrite reload', async () => {
  const host = renderer([doc()]);
  host.setRequest(async () =>
    doc('recovered', { tabId: 'tab-A', epoch: 2, documentId: 'grant-A', state: 'base', operationCancelled: true })
  );
  await assert.rejects(host.editCommand(doc(), []), { code: 'OWNER_COMMAND_CANCELLED' });
  assert.equal(host.all.current[0].id, 'recovered');
  assert.equal(host.all.current[0].epoch, 2);
  assert.equal(host.calls.length, 1);
  const reload = renderer([doc()]),
    held = deferred();
  reload.setRequest(async () => held.promise);
  const accepted = reload.editCommand(doc(), []);
  await Promise.resolve();
  reload.accept({
    activeId: 'fresh',
    documents: [doc('fresh', { tabId: 'tab-A', epoch: 3, documentId: 'grant-A', state: 'fresh' })],
    limit: 4,
  });
  held.resolve(doc('A', { state: 'accepted-old' }));
  await accepted;
  assert.equal(reload.all.current[0].id, 'fresh');
  assert.equal(reload.all.current[0].state, 'fresh');
});
test('cancelled/null outcomes never report success or send a duplicate owner request', async () => {
  let updates = 0;
  for (const result of [null, undefined])
    assert.throws(() => helper.acceptRecoveredCommand(result, () => updates++), { code: 'OWNER_COMMAND_RETIRED' });
  assert.equal(updates, 0);
  assert.throws(() => helper.acceptRecoveredCommand({ cancelled: true }, () => updates++), {
    code: 'OWNER_COMMAND_CANCELLED',
  });
  assert.equal(updates, 1);
  const host = renderer([doc()]);
  host.setRequest(async () => null);
  await assert.rejects(host.editCommand(doc(), []), { code: 'OWNER_COMMAND_RETIRED' });
  assert.equal(host.calls.length, 1);
  assert.equal(host.all.current[0].state, 'base');
});
test('standalone main actual request route independently refuses stale or retargeted strict writes/history', async () => {
  for (const operation of ['apply', 'undo', 'redo'])
    for (const mutation of ['head', 'epoch', 'active', 'replacement', 'closed']) {
      const record = { id: 'A', tabId: 'tab-A', epoch: 1, status: { documentId: 'grant-A', state: 'base' } },
        route = mainRoute(record);
      const pending = route.request('A', { operation, documentId: 'grant-A', state: 'base', commands: [] }),
        refused = assert.rejects(pending, /initiating workbook/);
      if (mutation === 'head') record.status.state = 'prior-edit';
      if (mutation === 'epoch') record.epoch = 2;
      if (mutation === 'active') route.registry.active = { id: 'B' };
      if (mutation === 'replacement') route.replace({ ...record, owner: { ...record.owner } });
      if (mutation === 'closed') record.owner.closed = true;
      await route.run();
      await refused;
      assert.equal(route.calls.length, 0, mutation);
    }
  const route = mainRoute({ id: 'A', tabId: 'tab-A', epoch: 1, status: { documentId: 'grant-A', state: 'base' } });
  await assert.rejects(
    route.request('A', { operation: 'apply', documentId: 'other', state: 'base' }),
    /initiating workbook/
  );
  assert.equal(route.calls.length, 0);
});
test('standalone main actual accepted strict apply and unbound legacy route remain functional', async () => {
  const record = { id: 'A', tabId: 'tab-A', epoch: 1, status: { documentId: 'grant-A', state: 'base' } },
    route = mainRoute(record);
  const first = route.request('A', { operation: 'apply', documentId: 'grant-A', state: 'base', commands: [] });
  await route.run();
  assert.equal((await first).state, 'accepted');
  assert.equal(route.calls.length, 1);
  const legacy = route.request('A', { operation: 'apply', state: 'accepted', commands: [] });
  route.registry.active = { id: 'B' };
  await route.run();
  assert.equal((await legacy).state, 'accepted');
  assert.equal(route.calls.length, 2);
});
test('standalone main durable late response updates original owner without rechecking the new active tab', async () => {
  const record = { id: 'A', tabId: 'tab-A', epoch: 1, status: { documentId: 'grant-A', state: 'base' } },
    route = mainRoute(record),
    held = deferred(),
    entered = deferred();
  record.owner.request = async (input) => {
    route.calls.push(input);
    entered.resolve();
    await held.promise;
    return { documentId: 'grant-A', state: 'durable-A' };
  };
  const pending = route.request('A', { operation: 'apply', documentId: 'grant-A', state: 'base', commands: [] }),
    running = route.run();
  await entered.promise;
  const other = { id: 'B', status: { documentId: 'grant-B', state: 'base-B' } };
  route.registry.active = other;
  held.resolve();
  await running;
  assert.equal((await pending).state, 'durable-A');
  assert.equal(record.status.state, 'durable-A');
  assert.equal(other.status.state, 'base-B');
  assert.equal(route.calls.length, 1);
});
test('standalone strict history awaits acceptance and retains the original inactive owner', async () => {
  const initial = doc('A', { canUndo: true }),
    host = renderer([initial, doc('B')]),
    entered = deferred(),
    gate = deferred();
  host.setRequest(async () => {
    entered.resolve();
    return gate.promise;
  });
  let settled = false;
  const pending = host.historyCommand(initial, 'undo').then(() => {
    settled = true;
  });
  await entered.promise;
  assert(!settled);
  assert.equal(host.calls[0].input.operation, 'undo');
  assert.equal(host.calls[0].input.state, 'base');
  assert.equal(host.calls[0].input.documentId, 'grant-A');
  host.activeDocument.current = 'B';
  gate.resolve(doc('A', { state: 'undone', canRedo: true }));
  await pending;
  assert.equal(host.all.current[0].state, 'undone');
  assert.equal(host.all.current[1].state, 'base');
  assert.equal(host.activeDocument.current, 'B');
  await assert.rejects(host.historyCommand(initial, 'undo'), { code: 'OWNER_COMMAND_RETIRED' });
  assert.equal(host.calls.length, 1);
});
test('standalone refused original-owner command cannot publish an error into another tab', async () => {
  const initial = doc('A', { canUndo: true }),
    host = renderer([initial, doc('B')]),
    entered = deferred(),
    gate = deferred();
  host.setRequest(async () => {
    entered.resolve();
    await gate.promise;
    throw Error('Original-owner refusal');
  });
  const pending = host.historyCommand(initial, 'undo'),
    refused = assert.rejects(pending, /Original-owner refusal/);
  await entered.promise;
  host.activeDocument.current = 'B';
  gate.resolve();
  await refused;
  await host.ownerQueues.current.get('tab-A').tail;
  assert(!host.errors.some((error) => String(error).includes('Original-owner refusal')));
  assert.equal(host.all.current[0].state, 'base');
  assert.equal(host.all.current[1].state, 'base');
});
test('standalone staged paste refuses retirement during preparation before any apply', async () => {
  const host = renderer([doc(), doc('B')]),
    entered = deferred(),
    gate = deferred(),
    controller = new AbortController();
  host.setRequest(async (id, input) => {
    if (input.operation === 'importBegin') {
      entered.resolve();
      await gate.promise;
      return { id: 'upload' };
    }
    return { cancelled: true };
  });
  const pending = host.importPaste(
      doc(),
      'value',
      { shape: { sheet: 'Sheet1', start: 'A1', rows: 1, columns: 1 }, count: 1, policy: 'grid' },
      { signal: controller.signal, onProgress() {} }
    ),
    refused = assert.rejects(pending, { code: 'OWNER_COMMAND_RETIRED' });
  await entered.promise;
  host.activeDocument.current = 'B';
  gate.resolve();
  await refused;
  assert.deepEqual(
    host.calls.map((call) => call.input.operation),
    ['importBegin', 'importCancel']
  );
  assert(
    host.calls.every((call) => call.id === 'A' && call.input.documentId === 'grant-A' && call.input.state === 'base')
  );
  assert.equal(host.all.current[0].state, 'base');
  assert.equal(host.all.current[1].state, 'base');
});
test('standalone staged paste awaits bound apply and preserves late accepted original owner', async () => {
  const host = renderer([doc(), doc('B')]),
    entered = deferred(),
    gate = deferred(),
    controller = new AbortController();
  host.setRequest(async (id, input) => {
    if (input.operation === 'importBegin') return { id: 'upload' };
    if (input.operation === 'importUpload') return { sequence: 1, count: 1 };
    if (input.operation === 'importSeal') return { type: 'importCells', token: 'sealed' };
    if (input.operation === 'apply') {
      entered.resolve();
      await gate.promise;
      return doc('A', { state: 'paste-accepted' });
    }
    throw Error('Unexpected request');
  });
  const pending = host.importPaste(
    doc(),
    'value',
    { shape: { sheet: 'Sheet1', start: 'A1', rows: 1, columns: 1 }, count: 1, policy: 'grid' },
    { signal: controller.signal, onProgress() {} }
  );
  await entered.promise;
  const apply = host.calls.find((call) => call.input.operation === 'apply');
  assert.equal(apply.id, 'A');
  assert.equal(apply.input.documentId, 'grant-A');
  assert.equal(apply.input.state, 'base');
  host.activeDocument.current = 'B';
  gate.resolve();
  await pending;
  assert.equal(host.all.current[0].state, 'paste-accepted');
  assert.equal(host.all.current[1].state, 'base');
});
test('preparation abort never requests cancellation of another operation', async () => {
  const host = renderer([doc()]),
    entered = deferred(),
    gate = deferred(),
    controller = new AbortController();
  let cancelled = 0;
  host.desktop.operation = async () => {
    throw Error('Paste must not look up arbitrary operations');
  };
  host.desktop.cancelPaste = async () => {
    cancelled++;
    throw Error('No applying paste is active');
  };
  host.setRequest(async (id, input) => {
    if (input.operation === 'importBegin') {
      entered.resolve();
      await gate.promise;
      return { id: 'upload' };
    }
    if (input.operation === 'importCancel') return { cancelled: true };
    throw Error('Unexpected apply');
  });
  const pending = host.importPaste(
      doc(),
      'value',
      { shape: { sheet: 'Sheet1', start: 'A1', rows: 1, columns: 1 }, count: 1, policy: 'grid' },
      { signal: controller.signal, onProgress() {} }
    ),
    refused = assert.rejects(pending, /Paste cancelled/);
  await entered.promise;
  controller.abort();
  gate.resolve();
  await refused;
  assert.equal(cancelled, 0);
  assert.deepEqual(
    host.calls.map((c) => c.input.operation),
    ['importBegin', 'importCancel']
  );
});
test('paste abort carries its exact request identity and finishes before a queued successor', async () => {
  const host = renderer([doc()]),
    entered = deferred(),
    gate = deferred(),
    cancelGate = deferred(),
    controller = new AbortController(),
    cancellations = [];
  host.desktop.operation = async () => {
    throw Error('Unbound operation lookup is forbidden');
  };
  host.desktop.cancelPaste = async (id, documentId, requestId) => {
    cancellations.push({ id, documentId, requestId });
    return cancelGate.promise;
  };
  host.setRequest(async (id, input) => {
    if (input.operation === 'importBegin') return { id: 'upload' };
    if (input.operation === 'importUpload') return { sequence: 1, count: 1 };
    if (input.operation === 'importSeal') return { type: 'importCells', token: 'sealed' };
    if (input.operation === 'apply') {
      entered.resolve();
      await gate.promise;
      return doc('A', { state: 'paste-accepted' });
    }
    throw Error('Unexpected request');
  });
  const pending = host.importPaste(
    doc(),
    'value',
    { shape: { sheet: 'Sheet1', start: 'A1', rows: 1, columns: 1 }, count: 1, policy: 'grid' },
    { signal: controller.signal, onProgress() {} }
  );
  await entered.promise;
  controller.abort();
  const apply = host.calls.find((c) => c.input.operation === 'apply').input;
  assert.deepEqual(cancellations, [{ id: 'A', documentId: 'grant-A', requestId: apply.requestId }]);
  assert.match(apply.requestId, /^[a-f0-9-]{36}$/);
  let finished = false;
  pending.then(() => {
    finished = true;
  });
  gate.resolve();
  await Promise.resolve();
  await Promise.resolve();
  assert(!finished);
  cancelGate.resolve({ outcome: 'completed', document: null });
  await pending;
  const calls = host.calls.length;
  controller.abort();
  await Promise.resolve();
  assert.equal(cancellations.length, 1);
  assert.equal(host.calls.length, calls);
});
test('actual main paste cancellation refuses a following operation on the same workbook', async () => {
  const source = sources['standalone/main.mjs'],
    tree = ts.createSourceFile('main.mjs', source, ts.ScriptTarget.Latest, true, ts.ScriptKind.JS);
  let code;
  function find(node) {
    if (
      ts.isCallExpression(node) &&
      node.expression.getText(tree) === 'handle' &&
      node.arguments[0]?.getText(tree) === "'cancel-paste'"
    )
      code = node.arguments[1].getText(tree);
    ts.forEachChild(node, find);
  }
  find(tree);
  assert(code);
  const record = { id: 'A', status: { documentId: 'grant-A', state: 'base' } },
    calls = [],
    context = {
      Error,
      active: { id: 'following', documentId: 'A', record, pasteRequestId: 'paste-next' },
      selected: (id) => {
        assert.equal(id, 'A');
        return record;
      },
      validId: (id) => ['paste-old', 'paste-next'].includes(id),
      status: (r) => r.status,
      cancelOperation: async (id, documentId) => {
        calls.push({ id, documentId });
        return { outcome: 'cancelled', document: null };
      },
    };
  vm.createContext(context);
  vm.runInContext('globalThis.cancel=' + code, context);
  assert.equal((await context.cancel('A', 'grant-A', 'paste-old')).outcome, 'completed');
  assert.equal(calls.length, 0);
  assert.equal((await context.cancel('A', 'grant-A', 'paste-next')).outcome, 'cancelled');
  assert.deepEqual(calls, [{ id: 'following', documentId: 'A' }]);
  context.active = { id: 'foreign', documentId: 'B', record: { id: 'B' }, pasteRequestId: 'paste-old' };
  assert.equal((await context.cancel('A', 'grant-A', 'paste-old')).outcome, 'completed');
  assert.equal(calls.length, 1);
  assert.throws(() => context.cancel('A', 'wrong-grant', 'paste-old'), /access is not granted/);
});
test('actual main tags only the staged Apply operation with its paste identity', async () => {
  const requestId = globalThis.crypto.randomUUID(),
    record = { id: 'A', tabId: 'tab-A', epoch: 1, status: { documentId: 'grant-A', state: 'base' } },
    route = mainRoute(record);
  const pending = route.request('A', {
    operation: 'apply',
    documentId: 'grant-A',
    state: 'base',
    requestId,
    commands: [{ type: 'importCells', token: 'sealed' }],
  });
  await route.run();
  await pending;
  assert.equal(route.observed.length, 1);
  assert.equal(route.observed[0].pasteRequestId, requestId);
  const ordinary = route.request('A', {
    operation: 'apply',
    documentId: 'grant-A',
    state: 'accepted',
    requestId,
    commands: [{ sheet: 'Sheet1', cell: 'A1', value: 1 }],
  });
  await route.run();
  await ordinary;
  assert.equal(route.observed.length, 2);
  assert.equal(route.observed[1].pasteRequestId, undefined);
});
test('actual main clipboard route awaits OS write completion and exposes rejection', async () => {
  const tree = ts.createSourceFile(
    'main.mjs',
    sources['standalone/main.mjs'],
    ts.ScriptTarget.Latest,
    true,
    ts.ScriptKind.JS
  );
  let code;
  function find(node) {
    if (
      ts.isCallExpression(node) &&
      node.expression.getText(tree) === 'handle' &&
      node.arguments[0]?.getText(tree) === "'clipboard-write'"
    )
      code = node.arguments[1].getText(tree);
    ts.forEachChild(node, find);
  }
  find(tree);
  assert(code);
  const gate = deferred(),
    writes = [],
    context = {
      Error,
      clipboard: {
        writeText: (text) => {
          writes.push(text);
          return gate.promise;
        },
      },
    };
  vm.createContext(context);
  vm.runInContext('globalThis.write=' + code, context);
  let complete = false;
  const pending = context.write('00123\t=SUM(A1)').then(() => {
    complete = true;
  });
  await Promise.resolve();
  assert(!complete);
  assert.deepEqual(writes, ['00123\t=SUM(A1)']);
  gate.resolve();
  await pending;
  assert(complete);
  context.clipboard.writeText = async () => {
    throw Error('OS clipboard refusal');
  };
  await assert.rejects(context.write('text'), /OS clipboard refusal/);
  assert.throws(() => context.write({ text: 'wrong type' }), /exceeds its limit/);
});
after(async () => {
  const changes = [];
  for (const [name, before] of Object.entries(sources))
    if ((await fs.readFile(path.join(component, name), 'utf8')) !== before) changes.push(name);
});
