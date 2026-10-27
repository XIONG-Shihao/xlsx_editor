import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { snapshotDelimitedFile, prepareDelimitedFileImport } from '../adapters/delimited-file-import.mjs';
import { inspectDelimitedSource, delimitedImportPages } from '../core/delimited-import.mjs';
import { createPasteStager } from '../adapters/large-paste.mjs';
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
async function fixture(t, text = '01\t=SUM(A1)\t\r\nragged') {
  const directory = await fs.realpath(await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-file-import-'))),
    filename = path.join(directory, 'selected.tsv'),
    privateDirectory = path.join(directory, 'private');
  await fs.writeFile(filename, text);
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  return { directory, filename, privateDirectory, bytes: Buffer.from(text) };
}
test('chooser snapshot is privately owned, immutable, hash checked and deterministically removed', async (t) => {
  const f = await fixture(t),
    snapshot = await snapshotDelimitedFile({ source: f.filename, directory: f.privateDirectory });
  assert.equal(snapshot.sha256, hash(f.bytes));
  assert(!JSON.stringify(snapshot).includes(f.directory));
  const roots = await fs.readdir(f.privateDirectory);
  assert.equal(roots.length, 1);
  const root = path.join(f.privateDirectory, roots[0]),
    privateFile = path.join(root, 'source.bytes');
  assert.equal((await fs.stat(root)).mode & 0o777, 0o700);
  assert.equal((await fs.stat(privateFile)).mode & 0o777, 0o400);
  const inspection = await inspectDelimitedSource(snapshot, { sheet: 'Sheet1', start: 'A1', delimiter: '\t' }),
    values = [];
  for await (const page of delimitedImportPages(snapshot, inspection)) values.push(...page.cells);
  assert.deepEqual(values, [
    [0, '01'],
    [1, '=SUM(A1)'],
    [2, ''],
    [3, 'ragged'],
  ]);
  await snapshot.dispose();
  assert.deepEqual(await fs.readdir(f.privateDirectory), []);
  await snapshot.dispose();
  await assert.rejects(snapshot.check(), /closed/);
  assert.equal(hash(await fs.readFile(f.filename)), hash(f.bytes));
});
test('original replacement and same-length edits invalidate the captured file before acceptance', async (t) => {
  for (const kind of ['replace', 'edit']) {
    const f = await fixture(t, 'old'),
      snapshot = await snapshotDelimitedFile({ source: f.filename, directory: f.privateDirectory });
    if (kind === 'replace') {
      await fs.rename(f.filename, f.filename + '.prior');
      await fs.writeFile(f.filename, 'new');
    } else await fs.writeFile(f.filename, 'new');
    await assert.rejects(snapshot.check(), /source changed/);
    await snapshot.dispose();
    assert.deepEqual(await fs.readdir(f.privateDirectory), []);
    assert.equal(await fs.readFile(f.filename, 'utf8'), 'new');
  }
});
test('snapshot mutation, private namespace symlinks and unexpected cleanup entries are never silently admitted', async (t) => {
  const f = await fixture(t),
    snapshot = await snapshotDelimitedFile({ source: f.filename, directory: f.privateDirectory }),
    root = path.join(f.privateDirectory, (await fs.readdir(f.privateDirectory))[0]),
    privateFile = path.join(root, 'source.bytes');
  await fs.chmod(privateFile, 0o600);
  await fs.writeFile(privateFile, 'changed');
  await assert.rejects(snapshot.check(), { code: 'csv-stale' });
  await fs.writeFile(path.join(root, 'foreign'), 'retain');
  await assert.rejects(snapshot.dispose(), /Unexpected/);
  assert.equal(await fs.readFile(path.join(root, 'foreign'), 'utf8'), 'retain');
  const other = await fixture(t),
    foreign = path.join(other.directory, 'foreign');
  await fs.mkdir(foreign, { mode: 0o700 });
  await fs.symlink(foreign, other.privateDirectory);
  await assert.rejects(snapshotDelimitedFile({ source: other.filename, directory: other.privateDirectory }), {
    code: /^(csv-unsupported|query-unsupported)$/,
  });
  assert.deepEqual(await fs.readdir(foreign), []);
});
test('cancellation while copying and unavailable chosen files leave no private partial snapshots', async (t) => {
  const f = await fixture(t, 'x'.repeat(600000)),
    controller = new AbortController();
  await assert.rejects(
    snapshotDelimitedFile({
      source: f.filename,
      directory: f.privateDirectory,
      signal: controller.signal,
      onProgress: (event) => {
        if (event.phase === 'copying-file') controller.abort(Error('cancel copy'));
      },
    }),
    /cancel copy/
  );
  assert.deepEqual(await fs.readdir(f.privateDirectory), []);
  assert.equal(hash(await fs.readFile(f.filename)), hash(f.bytes));
  await assert.rejects(
    snapshotDelimitedFile({ source: path.join(f.directory, 'missing.csv'), directory: f.privateDirectory }),
    /ENOENT/
  );
  assert.deepEqual(await fs.readdir(f.privateDirectory), []);
});
test('file preparation reuses durable staged pages, exact text policy and trusted history prune', async (t) => {
  const f = await fixture(t),
    stager = createPasteStager({
      directory: path.join(f.directory, 'attachments'),
      documentId: '11111111-1111-4111-8111-111111111111',
    });
  t.after(() => stager.close());
  const sink = {
    begin: (input) => stager.begin(input),
    upload: (input) => stager.upload(input),
    seal: (input) => stager.seal(input),
    cancel: (input) => stager.cancel(input),
    prune: () => stager.prune([]),
  };
  const ready = await prepareDelimitedFileImport({
    source: f.filename,
    directory: f.privateDirectory,
    sheet: 'Sheet1',
    start: 'C4',
    state: 'accepted',
    sink,
  });
  assert.equal(ready.descriptor.policy, 'text');
  assert.equal(ready.delimiter, '\t');
  assert.equal(ready.source.sha256, hash(f.bytes));
  assert.equal(ready.inspection.count, 4);
  const cells = [];
  for await (const page of stager.pages(ready.descriptor, { state: 'accepted' })) cells.push(...page.cells);
  assert.deepEqual(cells, [
    [0, '01'],
    [1, '=SUM(A1)'],
    [2, ''],
    [3, 'ragged'],
  ]);
  await ready.check();
  await ready.dispose();
  assert.deepEqual(await fs.readdir(f.privateDirectory), []);
  const retained = [ready.descriptor.id];
  await stager.prune(retained);
  assert.equal((await fs.readdir(path.join(f.directory, 'attachments'))).length, 1);
  await stager.prune([]);
  assert.deepEqual(await fs.readdir(path.join(f.directory, 'attachments')), []);
  assert.equal(hash(await fs.readFile(f.filename)), hash(f.bytes));
});
