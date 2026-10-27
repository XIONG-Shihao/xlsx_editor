import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, writeFile, readdir, rm } from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { unzipSync } from 'fflate';
import { workbook } from '../fixtures/standalone.mjs';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { SpreadsheetWorker } from '../adapters/worker-client.mjs';
import { createWorkbook, readWorkbook } from '../core/workbook-package.mjs';
import { encodeCsvText } from '../core/csv-codec.mjs';
import { nativeWorkbookOperationAsync, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { revision } from '../adapters/storage.mjs';
after(() => closeNativeWorkbookWorker());

async function fixture(t, extension = '.xlsx', bytes = workbook()) {
  const root = await mkdtemp(path.join(os.tmpdir(), 'xlsx-document-')),
    filename = path.join(root, 'document' + extension),
    recoveryDirectory = path.join(root, 'recovery');
  const owners = [];
  await writeFile(filename, bytes);
  t.after(async () => {
    for (const owner of owners) await owner.close();
    await rm(root, { recursive: true, force: true });
  });
  async function open(options = {}) {
    const owner = new SpreadsheetDocument({ recoveryDirectory, ...options });
    owners.push(owner);
    await owner.open(filename);
    return owner;
  }
  return { root, filename, recoveryDirectory, open, owners };
}
const read = (owner, range = 'A1:C1') =>
  owner.request({
    operation: 'read',
    state: owner.status().state,
    sheet: owner.status().kind === 'csv' ? 'CSV' : 'Sheet1',
    range,
  });
const apply = (owner, commands) => owner.request({ operation: 'apply', state: owner.status().state, commands });

test('native owner exposes bounded ranges and recalculates edits without a cell graph', async (t) => {
  const f = await fixture(t),
    owner = await f.open();
  assert.deepEqual(owner.status().view.sheets[0].cells, {});
  assert.equal((await read(owner)).cells.C1.calculation.value, 30);
  await assert.rejects(read(owner, 'A1:Z1000'), /10,000/);
  await apply(owner, [
    { sheet: 'Sheet1', cell: 'A1', value: 7 },
    { type: 'formatRange', sheet: 'Sheet1', range: 'A1:C1', format: { bold: true } },
  ]);
  assert.equal((await read(owner)).cells.C1.calculation.value, 27);
  assert.equal((await read(owner)).cells.A1.bold, true);
  assert.equal(owner.status().dirty, true);
  await owner.request({ operation: 'save' });
  const saved = readWorkbook(await readFile(f.filename));
  assert.equal(saved.sheets[0].cells.C1.cachedValue, 27);
  assert.equal(saved.sheets[0].cells.A1.bold, true);
  assert.deepEqual(unzipSync(await readFile(f.filename))['custom/preserved.bin'], new Uint8Array([0, 1, 2, 255]));
  assert.equal(owner.status().dirty, false);
});
test('undo/redo branches preserve saved-state identity and reject stale edits', async (t) => {
  const f = await fixture(t),
    owner = await f.open(),
    initial = owner.status().state;
  await apply(owner, [{ sheet: 'Sheet1', cell: 'A1', value: 2 }]);
  await owner.request({ operation: 'save' });
  await owner.request({ operation: 'undo' });
  assert.equal((await read(owner)).cells.A1.value, 10);
  assert.equal(owner.status().dirty, true);
  await owner.request({ operation: 'redo' });
  assert.equal(owner.status().dirty, false);
  await assert.rejects(
    owner.request({ operation: 'apply', state: initial, commands: [{ sheet: 'Sheet1', cell: 'B1', value: 9 }] }),
    /stale/
  );
  await owner.request({ operation: 'undo' });
  await apply(owner, [{ sheet: 'Sheet1', cell: 'B1', value: 4 }]);
  assert.equal(owner.status().canRedo, false);
  assert.equal((await read(owner)).cells.C1.calculation.value, 14);
});
test('restart replays compact commands; external changes require Save Copy', async (t) => {
  const f = await fixture(t),
    owner = await f.open(),
    original = await readFile(f.filename);
  await apply(owner, [
    { sheet: 'Sheet1', cell: 'A1', value: -0 },
    { sheet: 'Sheet1', cell: 'B1', value: 4 },
  ]);
  await owner.close();
  const restored = await f.open();
  assert.equal(restored.status().recovered, true);
  assert(Object.is((await read(restored)).cells.A1.value, -0));
  assert.deepEqual(await readFile(f.filename), original);
  await writeFile(f.filename, createWorkbook());
  await assert.rejects(restored.request({ operation: 'save' }), /changed on disk/);
  const copy = path.join(f.root, 'copy.xlsx');
  await restored.request({ operation: 'copy', destination: copy, expectedRevision: null });
  assert.equal(readWorkbook(await readFile(copy)).sheets[0].cells.B1.value, 4);
  assert.equal(restored.status().dirty, true);
  await assert.rejects(restored.request({ operation: 'copy', destination: copy, expectedRevision: null }), /exist/i);
});
test('interrupted save receipt recognizes a published candidate on restart', async (t) => {
  const f = await fixture(t),
    owner = await f.open();
  await apply(owner, [{ sheet: 'Sheet1', cell: 'A1', value: 66 }]);
  const bucket = path.join(
      f.recoveryDirectory,
      (await readdir(f.recoveryDirectory)).find((name) => /^[a-f0-9]{64}$/.test(name))
    ),
    file = path.join(bucket, 'current.json'),
    before = JSON.parse(await readFile(file, 'utf8'));
  await owner.request({ operation: 'save' });
  const savedRevision = await revision(f.filename);
  await owner.close();
  // Simulate a crash after the atomic file rename and before the final receipt.
  await writeFile(
    file,
    JSON.stringify({ ...before, pendingPublication: { head: before.head, revision: savedRevision } })
  );
  const restored = await f.open();
  assert.equal(restored.status().dirty, false);
  assert.equal((await read(restored)).cells.A1.value, 66);
});
test('failed native batch and oversized recovery command leave the accepted state unchanged', async (t) => {
  const f = await fixture(t),
    owner = await f.open(),
    before = owner.status();
  await assert.rejects(
    apply(owner, [
      { sheet: 'Sheet1', cell: 'A1', value: 7 },
      { sheet: 'Sheet1', cell: 'B1', value: 2, expectedValue: 999 },
    ])
  );
  assert.equal(owner.status().state, before.state);
  assert.equal((await read(owner)).cells.A1.value, 10);
  await assert.rejects(apply(owner, [{ sheet: 'Sheet1', cell: 'A1', value: 'x'.repeat(3 * 1024 * 1024) }]), /budget/);
  assert.equal(owner.status().state, before.state);
});
test('tampered recovery commands reject before changing the original', async (t) => {
  const f = await fixture(t),
    owner = await f.open();
  await apply(owner, [{ sheet: 'Sheet1', cell: 'A1', value: 9 }]);
  await owner.close();
  const original = await readFile(f.filename),
    bucket = path.join(
      f.recoveryDirectory,
      (await readdir(f.recoveryDirectory)).find((name) => /^[a-f0-9]{64}$/.test(name))
    ),
    manifest = JSON.parse(await readFile(path.join(bucket, 'current.json'), 'utf8'));
  await writeFile(path.join(bucket, manifest.head + '.json'), '{}');
  await assert.rejects(f.open(), { code: /^(paste-malformed|session-malformed)$/ });
  assert.deepEqual(await readFile(f.filename), original);
});
for (const encoding of ['utf-8', 'utf-16le', 'utf-16be'])
  test(`CSV ${encoding} sparse save preserves BOM, quoting, newlines and literal formulas`, async (t) => {
    const text = '"name";"value"\r\n"中文😀";"0012"\r\n',
      bytes = encodeCsvText(text, encoding, true);
    const f = await fixture(t, '.csv', bytes),
      owner = await f.open();
    await apply(owner, [{ sheet: 'CSV', cell: 'B2', value: '=literal' }]);
    await owner.request({ operation: 'save' });
    assert.deepEqual(
      new Uint8Array(await readFile(f.filename)),
      encodeCsvText(text.replace('0012', '=literal'), encoding, true)
    );
    assert.deepEqual(owner.status().view.sheets[0].cells, {});
    assert.equal((await read(owner, 'B2')).cells.B2.value, '=literal');
  });
test('CSV streamed append, search, undo, restart and explicit discard', async (t) => {
  const f = await fixture(t, '.tsv', Buffer.from('a\tb\n1\t2')),
    owner = await f.open();
  await apply(owner, [{ sheet: 'CSV', cell: 'B3', value: 'tail' }]);
  const result = await owner.request({
    operation: 'query',
    state: owner.status().state,
    sheet: 'CSV',
    query: 'tail',
    mode: 'display',
  });
  assert.equal(result.cells.B3.value, 'tail');
  await owner.request({ operation: 'undo' });
  assert.equal(owner.status().dirty, false);
  await owner.request({ operation: 'redo' });
  await owner.close();
  const restored = await f.open();
  assert.equal((await read(restored, 'B3')).cells.B3.value, 'tail');
  await restored.request({ operation: 'save' });
  assert.equal(await readFile(f.filename, 'utf8'), 'a\tb\n1\t2\n\ttail');
  await apply(restored, [{ sheet: 'CSV', cell: 'A1', value: 'discard' }]);
  await restored.request({ operation: 'discard' });
  await restored.close();
  const clean = await f.open();
  assert.equal((await read(clean, 'A1')).cells.A1.value, 'a');
  assert.equal(clean.status().dirty, false);
});
test('CSV literal Save and Copy preserve leading zero and formula-looking text with exact Undo/Redo state', async (t) => {
  const original = Buffer.from('Id,Expression\r\n00123,=SUM(A1)\r\nkeep,literal\r\n');
  const expected = Buffer.from('Id,Expression\r\n00123,=SUM(A1)\r\n00123,=SUM(A1)\r\n');
  const f = await fixture(t, '.csv', original),
    owner = await f.open(),
    initial = owner.status().state;
  const history = () => {
    const s = owner.status();
    return { state: s.state, dirty: s.dirty, canUndo: s.canUndo, canRedo: s.canRedo };
  };
  const literal = async (row) => {
    const cells = (await read(owner, `A${row}:B${row}`)).cells;
    assert.equal(cells['A' + row].value, '00123');
    assert.equal(cells['B' + row].value, '=SUM(A1)');
    assert.equal(cells['A' + row].formula, undefined);
    assert.equal(cells['B' + row].formula, undefined);
    return cells;
  };
  await literal(2);
  assert.deepEqual(history(), { state: initial, dirty: false, canUndo: false, canRedo: false });
  await apply(owner, [
    {
      type: 'setCells',
      sheet: 'CSV',
      cells: [
        { cell: 'A3', value: '00123' },
        { cell: 'B3', value: '=SUM(A1)' },
      ],
    },
  ]);
  const accepted = history();
  assert.notEqual(accepted.state, initial);
  assert.deepEqual(accepted, { state: accepted.state, dirty: true, canUndo: true, canRedo: false });
  await literal(3);
  await owner.request({ operation: 'undo' });
  assert.deepEqual(history(), { state: initial, dirty: false, canUndo: false, canRedo: true });
  assert.equal((await read(owner, 'A3:B3')).cells.A3.value, 'keep');
  await owner.request({ operation: 'redo' });
  assert.deepEqual(history(), accepted);
  const beforeCopy = await read(owner, 'A1:B3');
  const copy = path.join(f.root, 'literal-copy.csv');
  assert.deepEqual(await owner.request({ operation: 'copy', destination: copy, expectedRevision: null }), {
    copied: true,
  });
  assert.deepEqual(await readFile(copy), expected);
  assert.deepEqual(await readFile(f.filename), original);
  assert.deepEqual(history(), accepted);
  assert.deepEqual(await read(owner, 'A1:B3'), beforeCopy);
  await owner.request({ operation: 'save' });
  assert.deepEqual(await readFile(f.filename), expected);
  assert.deepEqual(history(), { ...accepted, dirty: false });
  await literal(2);
  await literal(3);
  await owner.request({ operation: 'undo' });
  assert.deepEqual(history(), { state: initial, dirty: true, canUndo: false, canRedo: true });
  assert.equal((await read(owner, 'A3:B3')).cells.B3.value, 'literal');
  assert.deepEqual(await readFile(f.filename), expected);
  await owner.request({ operation: 'redo' });
  assert.deepEqual(history(), { ...accepted, dirty: false });
  await literal(3);
  await owner.close();
  const cold = await f.open();
  assert.equal(cold.status().dirty, false);
  assert.equal(cold.status().canUndo, false);
  assert.equal(cold.status().canRedo, false);
  const coldCells = (await read(cold, 'A2:B3')).cells;
  for (const row of [2, 3]) {
    assert.equal(coldCells['A' + row].value, '00123');
    assert.equal(coldCells['B' + row].value, '=SUM(A1)');
    assert.equal(coldCells['B' + row].formula, undefined);
  }
  assert.deepEqual(await readFile(copy), expected);
  assert.deepEqual(await readFile(f.filename), expected);
  t.diagnostic(
    JSON.stringify({
      schema: 'workbook.xlsx.csv-literal-publication-case.v1',
      literalRows: 2,
      originalBytes: original.length,
      publishedBytes: expected.length,
      copyLeavesSourceAndAcceptedStateExact: true,
      saveKeepsAcceptedHead: true,
      oneUndoRedoBeforeAndAfterSave: true,
      coldLiteralReadback: true,
    })
  );
});
test('worker API uses the component helper and never returns the complete source', async (t) => {
  const f = await fixture(t),
    owner = new SpreadsheetWorker({ filename: f.filename, recoveryDirectory: f.recoveryDirectory });
  f.owners.push(owner);
  const status = await owner.open();
  assert.deepEqual(status.view.sheets[0].cells, {});
  assert(!JSON.stringify(status).includes(f.root));
  const result = await owner.request({ operation: 'read', state: status.state, sheet: 'Sheet1', range: 'A1' });
  assert.equal(result.cells.A1.value, 10);
});
test('large file-backed CSV streams exact UTF-8 byte spans across surrogate boundaries', async (t) => {
  const text = 'head,value\r\n"' + '😀x'.repeat(400000) + '",unchanged\r\nlater,target\r\n';
  const f = await fixture(t, '.csv', Buffer.from(text)),
    owner = await f.open();
  assert(JSON.stringify(owner.status()).length < 2048);
  await apply(owner, [{ sheet: 'CSV', cell: 'B3', value: 'with,comma' }]);
  await owner.request({ operation: 'save' });
  assert.equal(await readFile(f.filename, 'utf8'), text.replace('later,target', 'later,"with,comma"'));
});
test('CSV file-backed owner keeps recovery snapshot private across external changes and bounded reads', async (t) => {
  const original = Buffer.from('key,value\r\n001,"original\nvalue"\r\n002,tail\r\n'),
    f = await fixture(t, '.csv', original),
    owner = await f.open();
  const state = owner.status().state,
    status = JSON.stringify(owner.status());
  assert(!status.includes('rowOffsets'));
  assert(!status.includes('compressedText'));
  assert(!status.includes('blocks'));
  assert(status.length < 2048);
  await assert.rejects(read(owner, 'A1:Z1000'), /10,000/);
  await apply(owner, [{ sheet: 'CSV', cell: 'B2', expectedValue: 'original\nvalue', value: 'edited,🙂' }]);
  await writeFile(f.filename, 'external,content\n');
  const originalView = await owner.request({ operation: 'read', state, sheet: 'CSV', range: 'B2' });
  assert.equal(originalView.cells.B2.value, 'original\nvalue');
  assert.equal((await read(owner, 'B2')).cells.B2.value, 'edited,🙂');
  await assert.rejects(owner.request({ operation: 'save' }), /changed on disk/);
  const copy = path.join(f.root, 'saved-copy.csv');
  await owner.request({ operation: 'copy', destination: copy, expectedRevision: null });
  assert.equal(await readFile(copy, 'utf8'), 'key,value\r\n001,"edited,🙂"\r\n002,tail\r\n');
  assert.equal(await readFile(f.filename, 'utf8'), 'external,content\n');
  await owner.close();
  await assert.rejects(owner.request({ operation: 'read', state, sheet: 'CSV', range: 'B2' }), /closed/);
  const recovered = await f.open();
  assert.equal(recovered.status().recovered, true);
  assert.equal((await read(recovered, 'B2')).cells.B2.value, 'edited,🙂');
  assert.deepEqual(recovered.status().view.sheets[0].cells, {});
});
test('post-publication compaction failure reports accepted state instead of a failed edit', async (t) => {
  const f = await fixture(t);
  let fail = false;
  const owner = await f.open({
    native: (input) => {
      if (fail && input.operation === 'compact') throw Error('Injected index cleanup failure');
      return nativeWorkbookOperationAsync(input);
    },
  });
  fail = true;
  for (let i = 1; i <= 24; i++) await apply(owner, [{ sheet: 'Sheet1', cell: 'A1', value: i }]);
  assert.match(owner.status().warning, /Accepted edits are preserved/);
  assert.equal((await read(owner)).cells.A1.value, 24);
  await owner.close();
  const recovered = await f.open();
  assert.equal((await read(recovered)).cells.A1.value, 24);
});

test('standalone document cache requires the native implementation identity and reuses a verified index', async (t) => {
  const f = await fixture(t),
    calls = [],
    native = async (input) => {
      calls.push(input.operation);
      return nativeWorkbookOperationAsync(input);
    };
  const uncached = await f.open({ native });
  assert.equal((await read(uncached)).cells.C1.calculation.value, 30);
  await uncached.close();
  await assert.rejects(readdir(path.join(f.recoveryDirectory, '.native-index-cache')), { code: 'ENOENT' });
  const { indexedImplementationIdentity } = await import('../core/indexed-cache.mjs'),
    { nativeWorkbookBinary } = await import('../core/native-runtime.mjs');
  const cacheIdentity = await indexedImplementationIdentity(nativeWorkbookBinary());
  calls.length = 0;
  const first = await f.open({ native, cacheIdentity });
  assert(calls.includes('index'));
  await first.close();
  calls.length = 0;
  const second = await f.open({ native, cacheIdentity });
  assert(!calls.includes('index'), 'Verified native index should be reused');
  assert(!calls.includes('calculate'), 'Native-complete calculation should retain its proof');
  assert.equal((await read(second)).cells.C1.calculation.value, 30);
  await second.close();
  calls.length = 0;
  const changed = await f.open({ native, cacheIdentity: 'changed-implementation' });
  assert(calls.includes('index'));
  assert.equal((await read(changed)).cells.C1.calculation.value, 30);
});

test('standalone derived index cache has an independent default budget and safely misses when lowered', async (t) => {
  const { indexedCacheUsage } = await import('../core/indexed-cache.mjs');
  const f = await fixture(t),
    owner = await f.open({ maxRecoveryBytes: 1024 * 1024 });
  const directory = path.join(f.recoveryDirectory, '.native-index-cache');
  let usage = await indexedCacheUsage({ directory });
  assert.equal(usage.maxBytes, 2 * 1024 ** 3);
  assert.equal(usage.entries, 1, 'small durable recovery limit must not disable the derived cache');
  await apply(owner, [{ sheet: 'Sheet1', cell: 'A1', value: 17 }]);
  await owner.request({ operation: 'save' });
  await owner.close();
  usage = await indexedCacheUsage({ directory });
  assert(
    usage.entries >= 1 && usage.entries <= 2,
    'The import cache remains bounded; saved-source cache publication is optional'
  );
  assert.equal(readWorkbook(await readFile(f.filename)).sheets[0].cells.A1.value, 17);
  const reopened = await f.open({ maxRecoveryBytes: 1024 * 1024, maxIndexCacheBytes: 4096 });
  assert.equal((await read(reopened)).cells.A1.value, 17);
  usage = await indexedCacheUsage({ directory, maxBytes: 4096 });
  assert.equal(usage.entries, 0);
  assert(usage.bytes <= 4096);
});

test('document and worker cache budgets reject invalid bounds before starting owners', () => {
  for (const maxIndexCacheBytes of [0, 4095, 2 * 1024 ** 3 + 1, NaN, Infinity, 65536.5]) {
    assert.throws(() => new SpreadsheetDocument({ recoveryDirectory: 'unused', maxIndexCacheBytes }), {
      code: 'session-malformed',
    });
    assert.throws(
      () => new SpreadsheetWorker({ filename: 'unused.csv', recoveryDirectory: 'unused', maxIndexCacheBytes }),
      { code: 'session-malformed' }
    );
  }
});

test('standalone worker forwards a lower derived cache budget without changing workbook behavior', async (t) => {
  const { indexedCacheUsage } = await import('../core/indexed-cache.mjs'),
    f = await fixture(t),
    owner = new SpreadsheetWorker({
      filename: f.filename,
      recoveryDirectory: f.recoveryDirectory,
      maxIndexCacheBytes: 4096,
    });
  t.after(() => owner.close());
  const status = await owner.open();
  assert.equal(
    (await owner.request({ operation: 'read', state: status.state, sheet: 'Sheet1', range: 'C1' })).cells.C1.calculation
      .value,
    30
  );
  const usage = await indexedCacheUsage({
    directory: path.join(f.recoveryDirectory, '.native-index-cache'),
    maxBytes: 4096,
  });
  assert.equal(usage.entries, 0);
  assert(usage.bytes <= 4096);
});
