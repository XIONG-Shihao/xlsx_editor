import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { SpreadsheetDocument, listSpreadsheetRecovery } from '../adapters/document-session.mjs';
import { SpreadsheetWorker } from '../adapters/worker-client.mjs';
import { operationGate } from '../adapters/operation-gate.mjs';
import { createWorkbook } from '../core/workbook-package.mjs';
import { createWorkbookNativeClient } from '../core/native-client.mjs';
import { nativeWorkbookBinary } from '../core/native-runtime.mjs';
import { inspectPaste, pastePages } from '../ui/largePaste.ts';
import { decodedCommands } from '../core/transport.ts';
import { revision } from '../adapters/storage.mjs';

const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const sourceHelper = nativeWorkbookBinary(),
  helperBytes = await fs.readFile(sourceHelper),
  helperHash = hash(helperBytes);
const helperDirectory = await fs.realpath(await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-document-import-helper-'))),
  helper = path.join(helperDirectory, path.basename(sourceHelper));
await fs.writeFile(helper, helperBytes, { mode: 0o700 });
const priorHelper = process.env.XLSX_EDITOR_NATIVE_BIN;
process.env.XLSX_EDITOR_NATIVE_BIN = helper;
after(async () => {
  if (priorHelper === undefined) delete process.env.XLSX_EDITOR_NATIVE_BIN;
  else process.env.XLSX_EDITOR_NATIVE_BIN = priorHelper;
  await fs.rm(helperDirectory, { recursive: true, force: true });
});

async function fixture(t, { body, options = {} } = {}) {
  const directory = await fs.realpath(await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-document-import-'))),
    filename = path.join(directory, 'source.xlsx'),
    recoveryDirectory = path.join(directory, 'recovery'),
    owners = [];
  const parts = unzipSync(createWorkbook());
  parts['custom/opaque.bin'] = new Uint8Array([3, 0, 255, 19]);
  // Explicit accepted Auto policy keeps the strict unrelated-member oracle exact.
  parts['xl/workbook.xml'] = strToU8(
    strFromU8(parts['xl/workbook.xml']).replace(
      '</workbook>',
      '<calcPr calcId="162194" calcMode="auto" calcOnSave="1" fullCalcOnLoad="0" forceFullCalc="0"/></workbook>'
    )
  );
  if (body)
    parts['xl/worksheets/sheet1.xml'] = strToU8(`<worksheet xmlns="${ns}"><sheetData>${body}</sheetData></worksheet>`);
  const bytes = zipSync(parts);
  await fs.writeFile(filename, bytes);
  const client = createWorkbookNativeClient(helper),
    calls = [];
  const native = async (request) => {
    calls.push({
      operation: request.operation,
      changes: request.changes?.length,
      bytes: Buffer.byteLength(JSON.stringify(request)),
    });
    return client.call(request);
  };
  const make = (extra) => {
    const owner = new SpreadsheetDocument({ recoveryDirectory, native, cacheIdentity: null, ...options, ...extra });
    owners.push(owner);
    return owner;
  };
  const worker = (extra) => {
    const owner = new SpreadsheetWorker({ filename, recoveryDirectory, ...extra });
    owners.push(owner);
    return owner;
  };
  t.after(async () => {
    const failures = [];
    try {
      for (const owner of owners)
        try {
          await owner.close();
        } catch (error) {
          failures.push(error);
        }
    } finally {
      client.close();
      await fs.rm(directory, { recursive: true, force: true });
    }
    if (failures.length) throw new AggregateError(failures, 'Document test owner cleanup failed');
  });
  return { directory, filename, recoveryDirectory, bytes, make, worker, calls };
}
const status = (owner) => owner.request({ operation: 'status' });
const read = async (owner, range = 'A1') =>
  owner.request({ operation: 'read', state: (await status(owner)).state, sheet: 'Sheet1', range });
async function prepare(owner, text, options = {}) {
  const current = await status(owner),
    bound = { documentId: current.documentId, state: current.state },
    plan = inspectPaste(text, { sheet: 'Sheet1', start: 'A1', ...options });
  const begin = await owner.request({
    operation: 'importBegin',
    ...bound,
    shape: plan.shape,
    count: plan.count,
    policy: plan.policy,
  });
  let receipts = 0;
  for (const page of pastePages(text, plan)) {
    const reply = await owner.request({ operation: 'importUpload', ...bound, uploadId: begin.id, ...page });
    assert.equal(reply.sequence, ++receipts);
  }
  const descriptor = await owner.request({ operation: 'importSeal', ...bound, uploadId: begin.id });
  return { descriptor, bound, receipts };
}
const apply = async (owner, commands, options) =>
  owner.request({ operation: 'apply', state: (await status(owner)).state, commands }, options);
const applyImport = (owner, prepared, options) =>
  owner.request({ operation: 'apply', state: prepared.bound.state, commands: [prepared.descriptor] }, options);
function ownerDirectory(f, current) {
  return path.join(
    f.recoveryDirectory,
    current.documentId.length === 36 ? 'draft-' + current.documentId : current.documentId
  );
}
async function journal(f, current) {
  const directory = ownerDirectory(f, current),
    manifest = JSON.parse(await fs.readFile(path.join(directory, 'current.json'), 'utf8')),
    tip = manifest.history?.frames.at(-1) ?? { head: manifest.head, hash: manifest.headHash },
    nodes = [];
  let head = tip.head,
    expected = tip.hash;
  while (head !== null) {
    const bytes = await fs.readFile(path.join(directory, head + '.json'));
    assert.equal(hash(bytes), expected);
    const node = JSON.parse(bytes);
    nodes.push({ head, node, commands: decodedCommands({ encodedCommands: node.commands }), bytes: bytes.length });
    head = node.parent;
    expected = node.parentHash;
  }
  return { directory, manifest, nodes };
}

test('actual Document 100,000-cell paste publishes one compact undo frame and recovers its immutable attachment', async (t) => {
  const f = await fixture(t),
    owner = f.make();
  await owner.open(f.filename);
  const initial = await status(owner),
    sourceHash = await revision(f.filename),
    text = ('1\t'.repeat(99) + '1\n').repeat(1000);
  const prepared = await prepare(owner, text),
    events = [];
  assert.equal(prepared.receipts, 100);
  const accepted = await applyImport(owner, prepared, { onProgress: (event) => events.push(event) });
  assert.notEqual(accepted.state, initial.state);
  assert.equal(accepted.canUndo, true);
  assert.equal(accepted.canRedo, false);
  assert.deepEqual(accepted.view.sheets[0].cells, {});
  assert(!JSON.stringify(accepted).includes(f.directory));
  const summary = await owner.request({
    operation: 'selection',
    state: accepted.state,
    sheet: 'Sheet1',
    range: 'A1:CV1000',
  });
  assert.equal(summary.count, 100000);
  assert.equal(summary.numericCount, 100000);
  assert.equal(summary.sum, 100000);
  assert(events.some((event) => event.phase === 'applying-paste 100%'));
  assert(events.filter((event) => event.phase.startsWith('applying-paste ')).length <= 101);
  const saved = await journal(f, accepted);
  assert.equal(saved.nodes.length, 1);
  assert.deepEqual(saved.nodes[0].commands, [prepared.descriptor]);
  assert(saved.nodes[0].bytes < 2048);
  assert.equal(saved.manifest.history.frames.length, 2);
  const attachment = path.join(saved.directory, 'paste-v1', prepared.descriptor.id + '.paste.jsonl');
  assert.equal(await revision(attachment), prepared.descriptor.sha256);
  const copied = path.join(f.directory, 'copy.xlsx');
  await owner.request({ operation: 'copy', destination: copied, expectedRevision: null });
  const oracle = JSON.parse(
    execFileSync('python3', ['tests/support/staged-paste-oracle.py', f.filename, copied, '1000', '100'], {
      encoding: 'utf8',
      maxBuffer: 1024 * 1024,
    })
  );
  assert.equal(oracle.savedNumericOneCount, 100000);
  await owner.request({ operation: 'undo' });
  assert.equal((await status(owner)).state, initial.state);
  assert.equal((await read(owner, 'CV1000')).cells.CV1000, undefined);
  await owner.request({ operation: 'redo' });
  assert.equal((await status(owner)).state, accepted.state);
  assert.equal((await read(owner, 'CV1000')).cells.CV1000.value, 1);
  assert.equal(await revision(f.filename), sourceHash);
  await owner.close();
  const restored = f.make();
  await restored.open(f.filename);
  assert.equal((await status(restored)).state, accepted.state);
  assert.equal((await read(restored, 'CV1000')).cells.CV1000.value, 1);
  assert.equal(await revision(attachment), prepared.descriptor.sha256);
  await restored.request({ operation: 'undo' });
  assert.equal((await read(restored, 'A1')).cells.A1, undefined);
  assert.equal(await revision(f.filename), sourceHash);
});

test('staged import recalculates existing dependent formulas once after the complete candidate', async (t) => {
  const f = await fixture(t, { body: '<row r="3000"><c r="E3000"><f>SUM(A1:A2000)</f><v>0</v></c></row>' }),
    owner = f.make();
  await owner.open(f.filename);
  const prepared = await prepare(owner, '1\n'.repeat(2000)),
    before = f.calls.filter((call) => call.operation === 'calculate').length;
  await applyImport(owner, prepared);
  assert.equal(f.calls.filter((call) => call.operation === 'calculate').length - before, 1);
  assert.equal((await read(owner, 'E3000')).cells.E3000.calculation.value, 2000);
  await owner.request({ operation: 'undo' });
  assert.equal((await read(owner, 'E3000')).cells.E3000.calculation.value, 0);
  assert.equal(hash(await fs.readFile(f.filename)), hash(f.bytes));
});

for (const schema of [1, 2])
  test(`schema${schema} saved clean cursor keeps imported attachments for its unsaved redo tip`, async (t) => {
    const f = await fixture(t),
      owner = f.make();
    let filename = f.filename;
    if (schema === 1) await owner.open(filename);
    else {
      await owner.create('first saved import');
      filename = path.join(f.directory, 'first.xlsx');
      await owner.request({ operation: 'saveAs', destination: filename, expectedRevision: null });
    }
    const prepared = await prepare(owner, '1\n'.repeat(1001));
    await applyImport(owner, prepared);
    await owner.request({ operation: 'save' });
    const savedHash = await revision(filename),
      imported = await status(owner);
    await apply(owner, [{ sheet: 'Sheet1', cell: 'B1', value: 'redo literal' }]);
    const edited = (await status(owner)).state;
    await owner.request({ operation: 'undo' });
    assert.equal((await status(owner)).dirty, false);
    await owner.close();
    const retained = await journal(f, imported),
      attachment = path.join(retained.directory, 'paste-v1', prepared.descriptor.id + '.paste.jsonl');
    assert.equal(await revision(attachment), prepared.descriptor.sha256);
    if (schema === 2) {
      const list = await listSpreadsheetRecovery(f.recoveryDirectory);
      assert.equal(list.entries.length, 1);
      assert.equal(list.entries[0].dirty, false);
    }
    const restored = f.make();
    await restored.open(filename);
    assert.equal((await status(restored)).state, imported.state);
    assert.equal((await status(restored)).canRedo, true);
    assert.equal((await read(restored, 'A1001')).cells.A1001.value, 1);
    await restored.request({ operation: 'redo' });
    assert.equal((await status(restored)).state, edited);
    assert.equal((await read(restored, 'B1')).cells.B1.value, 'redo literal');
    assert.equal(await revision(filename), savedHash);
  });

test('unsaved UUID redo-only import retains its raw attachment and clean cursor across recovery', async (t) => {
  const f = await fixture(t),
    owner = f.make();
  await owner.create('private import');
  const initial = await status(owner),
    prepared = await prepare(owner, '1\n'.repeat(1001));
  await applyImport(owner, prepared);
  const accepted = await status(owner);
  await owner.request({ operation: 'undo' });
  assert.equal((await status(owner)).dirty, false);
  await owner.close();
  const list = await listSpreadsheetRecovery(f.recoveryDirectory);
  assert.equal(list.entries.length, 1);
  assert.equal(list.entries[0].id, initial.documentId);
  assert.equal(list.entries[0].unsaved, true);
  const restored = f.make();
  await restored.recover(initial.documentId);
  assert.equal((await status(restored)).state, initial.state);
  assert.equal((await status(restored)).canRedo, true);
  await restored.request({ operation: 'redo' });
  assert.equal((await status(restored)).state, accepted.state);
  assert.equal((await read(restored, 'A1001')).cells.A1001.value, 1);
  assert.equal((await restored.request({ operation: 'binding' })).filename, null);
});

test('document/state binding rejects stale uploads while cancellation cleans only the original partial handle', async (t) => {
  const f = await fixture(t),
    owner = f.make();
  await owner.open(f.filename);
  const initial = await status(owner),
    bound = { documentId: initial.documentId, state: initial.state },
    shape = { sheet: 'Sheet1', start: 'A1', rows: 2, columns: 1 };
  const handle = await owner.request({ operation: 'importBegin', ...bound, shape, count: 2, policy: 'grid' });
  await assert.rejects(
    owner.request({
      operation: 'importUpload',
      ...bound,
      documentId: 'other',
      uploadId: handle.id,
      sequence: 0,
      cells: [[0, '1']],
    }),
    { code: 'session-stale' }
  );
  await owner.request({ operation: 'importUpload', ...bound, uploadId: handle.id, sequence: 0, cells: [[0, '1']] });
  await apply(owner, [{ sheet: 'Sheet1', cell: 'B1', value: 9 }]);
  await assert.rejects(
    owner.request({ operation: 'importUpload', ...bound, uploadId: handle.id, sequence: 1, cells: [[1, '2']] }),
    /stale/
  );
  await assert.rejects(owner.request({ operation: 'importSeal', ...bound, uploadId: handle.id }), /stale/);
  await owner.request({ operation: 'importCancel', ...bound, uploadId: handle.id });
  assert.deepEqual(await fs.readdir(path.join(ownerDirectory(f, initial), 'paste-v1')), []);
  assert.equal((await read(owner, 'B1')).cells.B1.value, 9);
  assert.equal(hash(await fs.readFile(f.filename)), hash(f.bytes));
});

test('late protected target failure rejects one candidate, prunes unpublished attachments and keeps prior accepted undo', async (t) => {
  const body =
    '<row r="1001"><c r="A1001"><f t="array" ref="A1001:A1002">1+1</f><v>2</v></c></row><row r="1002"><c r="A1002"><v>2</v></c></row>';
  const f = await fixture(t, { body }),
    owner = f.make();
  await owner.open(f.filename);
  await apply(owner, [{ sheet: 'Sheet1', cell: 'B1', value: 'accepted' }]);
  const previous = await status(owner),
    prepared = await prepare(owner, '1\n'.repeat(1001));
  await assert.rejects(applyImport(owner, prepared), /array|special|qualified|partial/i);
  assert.equal((await status(owner)).state, previous.state);
  assert.equal((await read(owner, 'A1')).cells.A1, undefined);
  assert.equal((await read(owner, 'B1')).cells.B1.value, 'accepted');
  const saved = await journal(f, previous);
  assert.equal(saved.nodes.length, 1);
  assert.deepEqual(await fs.readdir(path.join(saved.directory, 'paste-v1')), []);
  await owner.request({ operation: 'undo' });
  assert.equal((await read(owner, 'B1')).cells.B1, undefined);
  assert.equal(hash(await fs.readFile(f.filename)), hash(f.bytes));
});

test('attachment bytes count toward recovery admission and partial budget failure remains cancellable', async (t) => {
  const f = await fixture(t, { options: { maxRecoveryBytes: 1024 * 1024 } }),
    owner = f.make();
  await owner.open(f.filename);
  const initial = await status(owner),
    bound = { documentId: initial.documentId, state: initial.state },
    text = ('x'.repeat(30) + '\n').repeat(30000),
    plan = inspectPaste(text, { sheet: 'Sheet1', start: 'A1' });
  const handle = await owner.request({
    operation: 'importBegin',
    ...bound,
    shape: plan.shape,
    count: plan.count,
    policy: plan.policy,
  });
  let failed = false;
  for (const page of pastePages(text, plan)) {
    try {
      await owner.request({ operation: 'importUpload', ...bound, uploadId: handle.id, ...page });
    } catch (error) {
      assert.match(error.message, /budget/);
      failed = true;
      break;
    }
  }
  assert.equal(failed, true);
  await owner.request({ operation: 'importCancel', ...bound, uploadId: handle.id });
  assert.equal((await status(owner)).state, initial.state);
  assert.equal((await journal(f, initial)).nodes.length, 0);
  assert.deepEqual(await fs.readdir(path.join(ownerDirectory(f, initial), 'paste-v1')), []);
  assert.equal(hash(await fs.readFile(f.filename)), hash(f.bytes));
});

test('confined attachment namespace and damaged replay refuse without changing source or accepted recovery files', async (t) => {
  const f = await fixture(t),
    owner = f.make();
  await owner.open(f.filename);
  const initial = await status(owner),
    privateDirectory = ownerDirectory(f, initial),
    foreign = path.join(f.directory, 'foreign');
  await fs.mkdir(foreign);
  await fs.writeFile(path.join(foreign, 'sentinel'), 'keep');
  await fs.symlink(foreign, path.join(privateDirectory, 'paste-v1'));
  await assert.rejects(prepare(owner, '1'), { code: /^(paste-bad-input|pivot-unsupported|session-bad-input)$/ });
  assert.equal(await fs.readFile(path.join(foreign, 'sentinel'), 'utf8'), 'keep');
  assert.equal((await status(owner)).state, initial.state);
  await fs.rm(path.join(privateDirectory, 'paste-v1'));
  await owner.close();
  const other = f.make();
  await other.open(f.filename);
  const prepared = await prepare(other, '1');
  await applyImport(other, prepared);
  const accepted = await status(other),
    record = await journal(f, accepted),
    attachment = path.join(record.directory, 'paste-v1', prepared.descriptor.id + '.paste.jsonl');
  await other.close();
  await fs.chmod(attachment, 0o600);
  const bytes = await fs.readFile(attachment);
  await fs.writeFile(attachment, Buffer.from(bytes.toString().replace('[0,"1"]', '[0,"2"]')));
  const damaged = await fs.readFile(attachment),
    manifest = await fs.readFile(path.join(record.directory, 'current.json')),
    restored = f.make();
  await assert.rejects(restored.open(f.filename), { code: 'paste-malformed' });
  await restored.close();
  assert.deepEqual(await fs.readFile(attachment), damaged);
  assert.deepEqual(await fs.readFile(path.join(record.directory, 'current.json')), manifest);
  assert.equal(hash(await fs.readFile(f.filename)), hash(f.bytes));
});

test('a direct cancellation grant between confirmed pages leaves the accepted frame and durable node unchanged', async (t) => {
  const f = await fixture(t),
    owner = f.make();
  await owner.open(f.filename);
  await apply(owner, [{ sheet: 'Sheet1', cell: 'B1', value: 'accepted' }]);
  const previous = await status(owner),
    prepared = await prepare(owner, '1\n'.repeat(10001));
  let granted;
  const gate = operationGate(1, (event) => {
    if (granted === undefined && event.phase.startsWith('applying-paste ')) granted = gate.authorize();
  });
  await assert.rejects(
    applyImport(owner, prepared, {
      onProgress: gate.progress,
      beforeCritical: gate.enterCritical,
      leaveCritical: gate.leaveCritical,
      checkCancelled: gate.check,
    }),
    /stopped|cancelled/
  );
  gate.finish();
  assert.equal(granted.granted, true);
  assert.equal((await status(owner)).state, previous.state);
  assert.equal((await read(owner, 'A1')).cells.A1, undefined);
  assert.equal((await read(owner, 'B1')).cells.B1.value, 'accepted');
  assert.equal((await journal(f, previous)).nodes.length, 1);
  assert.equal(hash(await fs.readFile(f.filename)), hash(f.bytes));
});

test('the real journal fence denies cancellation and publishes the complete imported frame', async (t) => {
  const f = await fixture(t),
    owner = f.make();
  await owner.open(f.filename);
  const prepared = await prepare(owner, '1\n'.repeat(1001));
  let denied;
  const gate = operationGate(1, (event) => {
    if (event.phase === 'saving-recovery') denied = gate.authorize();
  });
  const accepted = await applyImport(owner, prepared, {
    onProgress: gate.progress,
    beforeCritical: gate.enterCritical,
    leaveCritical: gate.leaveCritical,
    checkCancelled: gate.check,
  });
  gate.finish();
  assert.deepEqual(denied, { granted: false, reason: 'finishing' });
  assert.equal((await read(owner, 'A1001')).cells.A1001.value, 1);
  assert.equal((await journal(f, accepted)).nodes.length, 1);
  assert.equal(hash(await fs.readFile(f.filename)), hash(f.bytes));
});

test('worker partial upload termination cleans on restart while its previously accepted edit survives', async (t) => {
  const f = await fixture(t),
    worker = f.worker();
  let accepted = await worker.open();
  accepted = await apply(worker, [{ sheet: 'Sheet1', cell: 'B1', value: 'accepted' }]);
  const bound = { documentId: accepted.documentId, state: accepted.state };
  const handle = await worker.request({
    operation: 'importBegin',
    ...bound,
    shape: { sheet: 'Sheet1', start: 'A1', rows: 2000, columns: 1 },
    count: 2000,
    policy: 'grid',
  });
  const first = pastePages(
    '1\n'.repeat(2000),
    inspectPaste('1\n'.repeat(2000), { sheet: 'Sheet1', start: 'A1' })
  ).next().value;
  await worker.request({ operation: 'importUpload', ...bound, uploadId: handle.id, ...first });
  await worker.cancel();
  const restored = f.worker();
  const recovered = await restored.open();
  assert.equal(recovered.state, accepted.state);
  assert.equal((await read(restored, 'B1')).cells.B1.value, 'accepted');
  assert.equal((await read(restored, 'A1')).cells.A1, undefined);
  assert.deepEqual(await fs.readdir(path.join(ownerDirectory(f, recovered), 'paste-v1')), []);
  assert.equal(hash(await fs.readFile(f.filename)), hash(f.bytes));
});

test('granted worker cancellation during confirmed import pages recovers only the prior accepted frame', async (t) => {
  const f = await fixture(t),
    worker = f.worker();
  await worker.open();
  const accepted = await apply(worker, [{ sheet: 'Sheet1', cell: 'B1', value: 'accepted' }]),
    prepared = await prepare(worker, '1\n'.repeat(20001));
  let requestId, observe;
  const observed = new Promise((resolve) => (observe = resolve));
  const pending = applyImport(worker, prepared, {
    onStart: (id) => {
      requestId = id;
    },
    onProgress: (event) => {
      if (event.phase.startsWith('applying-paste ')) observe(event);
    },
  });
  const rejected = pending.catch((error) => error);
  await observed;
  const grant = await worker.authorizeCancel(requestId);
  assert.equal(grant.granted, true);
  await worker.cancel();
  assert.match((await rejected).message, /cancelled|stopped|reopen/);
  const restored = f.worker();
  assert.equal((await restored.open()).state, accepted.state);
  assert.equal((await read(restored, 'B1')).cells.B1.value, 'accepted');
  assert.equal((await read(restored, 'A1')).cells.A1, undefined);
  assert.equal((await journal(f, accepted)).nodes.length, 1);
  assert.deepEqual(await fs.readdir(path.join(ownerDirectory(f, accepted), 'paste-v1')), []);
  assert.equal(hash(await fs.readFile(f.filename)), hash(f.bytes));
});

test('larger compact fill uses one native request and one undo/replay command without a cell command array', async (t) => {
  const f = await fixture(t),
    owner = f.make();
  await owner.open(f.filename);
  await apply(owner, [{ sheet: 'Sheet1', cell: 'A1', value: 1 }]);
  const seed = (await status(owner)).state,
    calls = f.calls.length;
  const command = { type: 'fillRange', sheet: 'Sheet1', source: 'A1', range: 'A2:A100001' };
  const accepted = await apply(owner, [command]);
  assert.equal(f.calls.slice(calls).filter((call) => call.operation === 'bulk').length, 1);
  assert.equal((await read(owner, 'A100001')).cells.A100001.value, 1);
  const nodes = await journal(f, accepted);
  assert.deepEqual(nodes.nodes[0].commands, [command]);
  assert(nodes.nodes[0].bytes < 1024);
  await owner.request({ operation: 'undo' });
  assert.equal((await status(owner)).state, seed);
  assert.equal((await read(owner, 'A100001')).cells.A100001, undefined);
  await owner.request({ operation: 'redo' });
  await owner.close();
  const restored = f.make();
  await restored.open(f.filename);
  assert.equal((await read(restored, 'A100001')).cells.A100001.value, 1);
  await restored.request({ operation: 'undo' });
  assert.equal((await read(restored, 'A1')).cells.A1.value, 1);
  assert.equal(hash(await fs.readFile(f.filename)), hash(f.bytes));
});

for (const threaded of [false, true])
  test(`${threaded ? 'worker' : 'Document'} sealed upload abort prunes only unaccepted files and preserves clean-cursor Redo`, async (t) => {
    const f = await fixture(t),
      owner = threaded ? f.worker() : f.make();
    await owner.open(f.filename);
    const first = await prepare(owner, '1\n2'),
      accepted = await applyImport(owner, first);
    await owner.request({ operation: 'undo' });
    const clean = await status(owner);
    assert.equal(clean.canRedo, true);
    const abandoned = await prepare(owner, '99'),
      directory = path.join(ownerDirectory(f, clean), 'paste-v1');
    const retained = path.join(directory, first.descriptor.id + '.paste.jsonl'),
      unused = path.join(directory, abandoned.descriptor.id + '.paste.jsonl');
    assert.equal(await revision(retained), first.descriptor.sha256);
    assert.equal(await revision(unused), abandoned.descriptor.sha256);
    await assert.rejects(
      owner.request({ operation: 'importPrune', documentId: clean.documentId, state: accepted.state }),
      /stale/
    );
    await assert.rejects(owner.request({ operation: 'importPrune', documentId: 'foreign', state: clean.state }), {
      code: 'session-stale',
    });
    await owner
      .request({ operation: 'importCancel', ...first.bound, uploadId: first.descriptor.id })
      .catch((error) => assert.match(error.message, /upload|partial|sealed|active/i));
    assert.equal(await revision(retained), first.descriptor.sha256);
    assert.deepEqual(
      await owner.request({ operation: 'importPrune', documentId: clean.documentId, state: clean.state }),
      { pruned: true }
    );
    await assert.rejects(fs.lstat(unused), { code: 'ENOENT' });
    assert.equal(await revision(retained), first.descriptor.sha256);
    await owner.request({ operation: 'redo' });
    assert.equal((await read(owner, 'A1:A2')).cells.A2.value, 2);
    assert.equal((await journal(f, accepted)).nodes.length, 1);
    assert.equal(hash(await fs.readFile(f.filename)), hash(f.bytes));
  });
