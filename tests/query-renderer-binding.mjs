/** Actual renderer query callback with deterministic queue/IPC boundaries. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import fs from 'node:fs/promises';
import path from 'node:path';
import ts from 'typescript';
import { createHash } from 'node:crypto';
const component = path.resolve(import.meta.dirname, '..'),
  sourceFile = path.join(component, 'standalone/renderer.tsx'),
  source = await fs.readFile(sourceFile, 'utf8'),
  hash = (s) => createHash('sha256').update(s).digest('hex'),
  before = hash(source);
const tree = ts.createSourceFile(sourceFile, source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
function extract(node) {
  if (ts.isFunctionDeclaration(node) && node.name?.text === 'queryCall') return node.getText(tree);
  return ts.forEachChild(node, extract);
}
const actual = extract(tree);
assert(actual);
const doc = () => ({ id: 'tab', epoch: 1, documentId: 'document', state: 'base' });
function setup() {
  const queued = [],
    calls = [],
    updates = [],
    all = { current: [doc()] },
    activeDocument = { current: 'tab' },
    context = {
      all,
      activeDocument,
      Error,
      Promise,
      structuredClone,
      update: (r) => updates.push(r),
      desktop: {
        queryConnection: async (id, input) => {
          calls.push({ id, input });
          return { queryPreview: {}, document: { ...doc(), state: 'accepted' } };
        },
      },
      enqueue: (run, id) =>
        new Promise((resolve, reject) =>
          queued.push({ run: () => Promise.resolve().then(run).then(resolve, reject), id })
        ),
    };
  vm.createContext(context);
  vm.runInContext(
    ts.transpileModule(actual + '\nglobalThis.queryCall=queryCall', {
      compilerOptions: { target: ts.ScriptTarget.ES2022 },
    }).outputText,
    context
  );
  return { ...context, queued, calls, updates };
}
test('queued query keeps the initiating head and refuses after a preceding edit', async () => {
  const f = setup(),
    pending = f.queryCall(doc(), 'refresh', {
      connectionId: 'saved',
      expectedRevision: 3,
      requestId: 'request',
      expectedConversionRevision: 'review',
      expectedReplacementRevision: 'review',
    }),
    refused = assert.rejects(pending, /changed/);
  assert.equal(f.queued.length, 1);
  f.all.current[0].state = 'preceding-edit';
  await f.queued.shift().run();
  await refused;
  assert.equal(f.calls.length, 0);
  assert.equal(f.updates.length, 0);
});
test('queued query snapshots its scope and options and forwards the review page exactly', async () => {
  const f = setup(),
    bound = doc(),
    fields = {
      connectionId: 'saved',
      expectedRevision: 3,
      requestId: 'request',
      replacementOffset: 100,
      expectedConversionRevision: 'review',
    },
    pending = f.queryCall(bound, 'preview', fields);
  bound.state = 'mutated';
  fields.replacementOffset = 200;
  fields.expectedConversionRevision = 'other';
  await f.queued.shift().run();
  await pending;
  assert.deepEqual(JSON.parse(JSON.stringify(f.calls)), [
    {
      id: 'tab',
      input: {
        operation: 'preview',
        documentId: 'document',
        state: 'base',
        connectionId: 'saved',
        expectedRevision: 3,
        requestId: 'request',
        replacementOffset: 100,
        expectedConversionRevision: 'review',
      },
    },
  ]);
  assert.equal(f.updates.length, 1);
});
test('tab departure and runtime replacement reject without dispatch; inventory remains a read', async () => {
  for (const mutate of [
    (f) => {
      f.activeDocument.current = 'other';
    },
    (f) => {
      f.all.current[0].epoch = 2;
    },
    (f) => {
      f.all.current[0].documentId = 'replacement';
    },
  ]) {
    const f = setup(),
      pending = f.queryCall(doc(), 'preview', { requestId: 'request' }),
      refused = assert.rejects(pending, /changed/);
    mutate(f);
    await f.queued.shift().run();
    await refused;
    assert.equal(f.calls.length, 0);
  }
  const f = setup();
  await f.queryCall(doc(), 'list');
  assert.equal(f.queued.length, 0);
  assert.equal(f.calls.length, 1);
});
after(async () => {
  assert.equal(hash(await fs.readFile(sourceFile)), before);
});
