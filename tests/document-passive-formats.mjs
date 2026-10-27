import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { SpreadsheetWorker } from '../adapters/worker-client.mjs';
const sha = (bytes) => createHash('sha256').update(bytes).digest('hex'),
  fixtures = path.resolve('fixtures/remaining-formats');
// Passive (read-only) formats are the old binary ones. A macro workbook (.xlsm)
// is edited like any other (tests/document-macro-workbooks.mjs).
async function fixture(t, format = 'xls') {
  const root = await fs.realpath(await fs.mkdtemp(path.join(os.tmpdir(), 'document-passive-'))),
    filename = path.join(root, 'source.' + format),
    source = path.join(fixtures, 'universal-content.' + format),
    bytes = await fs.readFile(source),
    recoveryDirectory = path.join(root, 'recovery'),
    indexDirectory = path.join(root, 'index'),
    owners = [];
  await fs.copyFile(source, filename);
  const make = () => {
    const owner = new SpreadsheetDocument({
      recoveryDirectory,
      indexDirectory,
      native: async () => {
        throw Error('Passive owner must never call native calculation or edit');
      },
      cacheIdentity: null,
    });
    owners.push(owner);
    return owner;
  };
  t.after(async () => {
    try {
      for (const owner of owners) await owner.close();
    } finally {
      await fs.rm(root, { recursive: true, force: true });
    }
  });
  return { root, filename, bytes, recoveryDirectory, indexDirectory, make, owners };
}
for (const format of ['xls', 'xlsb'])
  test(
    'Document ' + format + ' exposes stored readonly state, exact Copy and new values-only publication',
    async (t) => {
      const f = await fixture(t, format),
        owner = f.make(),
        opened = await owner.open(f.filename),
        bound = { state: opened.state, documentId: opened.documentId };
      assert.equal(opened.kind, format);
      assert.equal(opened.state, 'base');
      assert.equal(opened.view.readOnly, true);
      assert.equal(opened.dirty, false);
      assert.equal(opened.canUndo, false);
      assert.equal(opened.canRedo, false);
      assert.equal(opened.view.originalFormat, format);
      const read = await owner.request({ operation: 'read', state: opened.state, sheet: 'Sheet1', range: 'A1:C4' });
      assert.equal(read.cells.C1.calculation.status, 'stored');
      assert(!JSON.stringify(opened).includes(f.root));
      const copied = path.join(f.root, 'copy.' + format);
      await owner.request({ operation: 'copy', destination: copied, expectedRevision: null });
      assert.deepEqual(await fs.readFile(copied), f.bytes);
      await assert.rejects(owner.request({ operation: 'save' }), /read-only/);
      await assert.rejects(
        owner.request({
          operation: 'apply',
          state: opened.state,
          commands: [{ sheet: 'Sheet1', cell: 'A1', value: 7 }],
        }),
        /read-only/
      );
      assert.equal(owner.status().state, opened.state);
      let evaluation;
      const outputs = [];
      for (const target of ['xls', 'xlsb', 'xlsx']) {
        const destination = path.join(f.root, 'data.' + target),
          result = await owner.request({
            operation: 'dataOnlyExport',
            ...bound,
            format: target,
            destination,
            expectedRevision: null,
          });
        assert.equal(result.copied, true);
        assert.equal(result.revision, sha(await fs.readFile(destination)));
        assert.equal(result.sourceRevision, sha(f.bytes));
        assert(!JSON.stringify(result).includes(f.root));
        assert(!('path' in result));
        assert(result.omissions.some((value) => value.includes('VBA')));
        outputs.push({ format: target, ...result });
        await assert.rejects(
          owner.request({ operation: 'dataOnlyExport', ...bound, format: target, destination, expectedRevision: null }),
          /changed|revision|exists/
        );
      }
      assert.deepEqual(await fs.readFile(f.filename), f.bytes);
      assert.equal(owner.status().state, 'base');
      assert.equal(owner.status().dirty, false);
      assert.equal(owner.status().canUndo, false);
    }
  );
test('Document passive recovery resumes a missing original and guards snapshots, new destinations and revisions', async (t) => {
  const f = await fixture(t),
    owner = f.make(),
    opened = await owner.open(f.filename),
    bound = { documentId: opened.documentId, state: opened.state };
  await owner.close();
  await fs.rm(f.filename);
  const resumed = f.make(),
    restored = await resumed.resume(f.filename);
  assert(restored.recovered);
  assert.equal(restored.kind, 'xls');
  assert.equal(
    (await resumed.request({ operation: 'read', state: 'base', sheet: 'Sheet1', range: 'C1' })).cells.C1.calculation
      .value,
    6
  );
  const destination = path.join(f.root, 'recovered.xls');
  await resumed.request({ operation: 'copy', destination, expectedRevision: null });
  assert.deepEqual(await fs.readFile(destination), f.bytes);
  for (const request of [
    { state: 'old' },
    { documentId: 'other' },
    { destination: f.filename.replace('.xls', '.xlsx'), expectedRevision: '0'.repeat(64) },
    { format: 'xlsm' },
    { path: f.filename },
    { destination: path.join(f.recoveryDirectory, 'private.xlsx') },
  ])
    await assert.rejects(
      resumed.request({
        operation: 'dataOnlyExport',
        ...bound,
        format: 'xlsx',
        destination: path.join(f.root, 'data.xlsx'),
        expectedRevision: null,
        ...request,
      }),
      /stale|destination|fields/
    );
  await assert.rejects(
    resumed.request({ operation: 'copy', destination: path.join(f.root, 'wrong.xlsx'), expectedRevision: null }),
    { code: 'session-bad-input' }
  );
  await resumed.close();
  const directory = path.join(f.recoveryDirectory, opened.documentId),
    manifest = JSON.parse(await fs.readFile(path.join(directory, 'current.json'), 'utf8')),
    snapshot = path.join(directory, manifest.snapshot + manifest.extension);
  await fs.chmod(snapshot, 0o600);
  await fs.writeFile(snapshot, 'damaged');
  await assert.rejects(f.make().resume(f.filename), { code: 'session-malformed' });
  assert.deepEqual(await fs.readFile(destination), f.bytes);
});
test('cancel before publication retains passive accepted source and does not create a destination', async (t) => {
  const f = await fixture(t),
    owner = f.make(),
    opened = await owner.open(f.filename),
    destination = path.join(f.root, 'cancelled.xlsx');
  let cancelled = false;
  await assert.rejects(
    owner.request(
      {
        operation: 'dataOnlyExport',
        state: opened.state,
        documentId: opened.documentId,
        format: 'xlsx',
        destination,
        expectedRevision: null,
      },
      {
        onProgress: (event) => {
          if (event.phase === 'exporting-stored-values') cancelled = true;
        },
        checkCancelled: () => {
          if (cancelled) throw Error('cancelled before acceptance');
        },
      }
    ),
    /cancelled/
  );
  await assert.rejects(fs.stat(destination), /ENOENT/);
  assert.equal(owner.status().state, opened.state);
  assert.equal(owner.status().dirty, false);
  assert.deepEqual(await fs.readFile(f.filename), f.bytes);
  const result = await owner.request({
    operation: 'dataOnlyExport',
    state: opened.state,
    documentId: opened.documentId,
    format: 'xlsx',
    destination,
    expectedRevision: null,
  });
  assert(result.copied);
});
test('real owner worker routes bounded passive reading and export without native helper calls', async (t) => {
  const f = await fixture(t),
    worker = new SpreadsheetWorker({ filename: f.filename, recoveryDirectory: f.recoveryDirectory });
  f.owners.push(worker);
  const opened = await worker.open(),
    bound = { state: opened.state, documentId: opened.documentId };
  assert.equal(
    (await worker.request({ operation: 'read', state: opened.state, sheet: 'Sheet1', range: 'C1' })).cells.C1
      .calculation.value,
    6
  );
  const result = await worker.request({
    operation: 'dataOnlyExport',
    ...bound,
    format: 'xlsb',
    destination: path.join(f.root, 'worker.xlsb'),
    expectedRevision: null,
  });
  assert(result.copied);
  assert.equal(result.revision, sha(await fs.readFile(path.join(f.root, 'worker.xlsb'))));
  assert.deepEqual(await fs.readFile(f.filename), f.bytes);
  assert.equal((await worker.request({ operation: 'status' })).dirty, false);
});
test('oversized sparse passive source refuses before any recovery copy and recovery size refuses before hash parsing', async (t) => {
  const f = await fixture(t),
    limit = 64 * 1024 * 1024 + 1;
  await fs.chmod(f.filename, 0o600);
  const file = await fs.open(f.filename, 'r+');
  try {
    await file.truncate(limit);
  } finally {
    await file.close();
  }
  await assert.rejects(f.make().open(f.filename), { code: 'session-limit' });
  await assert.rejects(fs.stat(f.recoveryDirectory), /ENOENT/);
  await assert.rejects(fs.stat(f.indexDirectory), /ENOENT/);
  const valid = await fixture(t),
    owner = valid.make(),
    opened = await owner.open(valid.filename);
  await owner.close();
  const directory = path.join(valid.recoveryDirectory, opened.documentId),
    manifest = JSON.parse(await fs.readFile(path.join(directory, 'current.json'), 'utf8')),
    snapshot = path.join(directory, manifest.snapshot + manifest.extension);
  await fs.chmod(snapshot, 0o600);
  const handle = await fs.open(snapshot, 'r+');
  try {
    await handle.truncate(limit);
  } finally {
    await handle.close();
  }
  await assert.rejects(valid.make().resume(valid.filename), { code: 'session-limit' });
  assert.deepEqual(await fs.readFile(valid.filename), valid.bytes);
  assert.equal((await fs.stat(snapshot)).size, limit);
});
