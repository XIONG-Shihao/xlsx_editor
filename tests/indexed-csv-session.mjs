import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import syncFs from 'node:fs';
import { syncBuiltinESMExports } from 'node:module';
import path from 'node:path';
import { tmpdir } from 'node:os';
import { createHash } from 'node:crypto';
import { CsvFileSession } from '../core/indexed-csv-session.mjs';
import { readCsvFileIndex } from '../core/csv-file.mjs';
import { writeCsvIndex } from '../core/csv-file-save.mjs';
import { encodeCsvText } from '../core/csv-codec.mjs';

const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const state = (result) => result.view.storage.state;
async function fixture(t, bytes = 'id,value\r\n1,original\r\n2,tail\r\n', options = {}) {
  const root = await fs.mkdtemp(path.join(tmpdir(), 'workbook-indexed-csv-')),
    source = path.join(root, 'source.csv'),
    storageDirectory = path.join(root, 'owners');
  await fs.writeFile(source, bytes);
  const owner = new CsvFileSession();
  t.after(async () => {
    await owner.close();
    await fs.rm(root, { recursive: true, force: true });
  });
  const view = await owner.open(source, { storageDirectory, expectedRevision: hash(bytes), ...options });
  return { owner, root, source, storageDirectory, view };
}
const read = async (owner, version, range) => (await owner.read(version, 'CSV', range)).cells;

test('CSV owner returns bounded metadata and shares immutable source across branched revisions and saves', async (t) => {
  const { owner, source, storageDirectory, view } = await fixture(t);
  assert.equal(view.storage.kind, 'csv-indexed-v1');
  assert.deepEqual(view.sheets[0].cells, {});
  assert.equal(view.sheets[0].rows, 3);
  assert(Buffer.byteLength(JSON.stringify(view)) < 2048);
  assert(!JSON.stringify(view).includes(source));
  assert(!('text' in view));
  assert(!('rowOffsets' in view));
  const initial = owner.diagnosticStats();
  assert.equal(initial.retainedStates, 1);
  assert.equal(initial.indexBytes, 44);
  const first = await owner.apply('0', [{ sheet: 'CSV', cell: 'B2', expectedValue: 'original', value: 'first' }]);
  const second = await owner.apply(state(first), [{ sheet: 'CSV', cell: 'A3', value: '=1+1' }]);
  const branch = await owner.apply('0', [{ sheet: 'CSV', cell: 'B2', value: 'branch' }]);
  assert.equal((await read(owner, '0', 'B2')).B2.value, 'original');
  assert.equal((await read(owner, state(first), 'B2')).B2.value, 'first');
  assert.equal((await read(owner, state(branch), 'B2')).B2.value, 'branch');
  assert.equal((await read(owner, state(second), 'A3')).A3.value, '=1+1');
  const output = await owner.export(state(second)),
    bytes = await fs.readFile(output.path);
  assert.equal(bytes.toString(), 'id,value\r\n1,first\r\n=1+1,tail\r\n');
  assert.equal(output.revision, hash(bytes));
  assert.deepEqual(
    await owner.export(state(second)),
    output,
    'Repeated export of one revision reuses its verified candidate'
  );
  assert.equal(owner.diagnosticStats().indexBytes, initial.indexBytes);
  const branchOutput = await owner.export(state(branch));
  await owner.compact(['0', state(second)]);
  await assert.rejects(owner.read(state(branch), 'CSV', 'A1'), { code: 'internal-unknown-reference' });
  assert.equal(
    await fs.stat(branchOutput.path).then(
      () => true,
      () => false
    ),
    false
  );
  assert.equal((await read(owner, '0', 'B2')).B2.value, 'original');
  assert.equal(owner.diagnosticStats().retainedStates, 2);
  assert.equal((await fs.readFile(source)).toString(), 'id,value\r\n1,original\r\n2,tail\r\n');
  await owner.close();
  assert.deepEqual(await fs.readdir(storageDirectory), []);
  assert.equal(owner.diagnosticStats().closed, true);
  assert.equal(owner.diagnosticStats().storageRoots, 0);
  assert.equal(owner.diagnosticStats().indexBytes, 0);
  await assert.rejects(owner.read('0', 'CSV', 'A1'), /closed/);
});

for (const [encoding, bom] of [
  ['utf-8', false],
  ['utf-8', true],
  ['utf-16le', true],
  ['utf-16be', true],
])
  test(`CSV owner saves exact ${encoding} BOM=${bom} bytes with quoted newlines and sparse extension`, async (t) => {
    const original = 'id;value\r\n1;"🙂中文\r\nquote""x"\n2;keep\r3;tail',
      bytes = encodeCsvText(original, encoding, bom);
    const { owner, source } = await fixture(t, bytes, { delimiter: ';' });
    const edited = await owner.apply('0', [
      { sheet: 'CSV', cell: 'B2', value: 'new;🙂' },
      { sheet: 'CSV', cell: 'B4', value: '終' },
      { sheet: 'CSV', cell: 'C6', value: 'last' },
    ]);
    const output = await owner.export(state(edited)),
      actual = await fs.readFile(output.path);
    const expected = encodeCsvText('id;value\r\n1;"new;🙂"\n2;keep\r3;終\r\n\r\n;;last', encoding, bom);
    assert.deepEqual(actual, Buffer.from(expected));
    assert.equal(output.revision, hash(expected));
    const base = await owner.export('0');
    assert.deepEqual(await fs.readFile(base.path), Buffer.from(bytes));
    assert.deepEqual(await fs.readFile(source), Buffer.from(bytes));
  });

test('CSV owner reuses row metadata while enforcing state/history budgets before acceptance', async (t) => {
  const { owner } = await fixture(t, undefined, { maxStates: 3, maxHistoryBytes: 2048 });
  const first = await owner.apply('0', [{ sheet: 'CSV', cell: 'B2', value: 'first' }]);
  const second = await owner.apply(state(first), [{ sheet: 'CSV', cell: 'B2', value: 'second' }]);
  const before = owner.diagnosticStats();
  await assert.rejects(owner.apply(state(second), [{ sheet: 'CSV', cell: 'B2', value: 'third' }]), {
    code: 'csv-limit',
  });
  assert.deepEqual(owner.diagnosticStats(), before);
  assert.equal((await read(owner, state(second), 'B2')).B2.value, 'second');
  await owner.compact([state(second)]);
  const compacted = owner.diagnosticStats();
  assert(compacted.historyBytes < before.historyBytes);
  await assert.rejects(owner.read('0', 'CSV', 'A1'), { code: 'internal-unknown-reference' });
  await assert.rejects(owner.apply(state(second), [{ sheet: 'CSV', cell: 'B2', value: 'x'.repeat(1500) }]), {
    code: 'csv-limit',
  });
  assert.equal(owner.diagnosticStats().retainedStates, 1);
  const third = await owner.apply(state(second), [{ sheet: 'CSV', cell: 'B2', value: 'third' }]);
  assert.equal(state(third), '3');
  assert.equal(owner.diagnosticStats().indexBytes, before.indexBytes);
});

test('CSV owner rejects unsupported, stale, malformed and oversized edits without changing accepted views', async (t) => {
  const { owner } = await fixture(t, undefined, { maxProjectionBytes: 1024 });
  for (const commands of [
    [{ sheet: 'CSV', cell: 'B2', expectedValue: 'wrong', value: 'bad' }],
    [{ sheet: 'CSV', cell: 'B2', value: '\ud800' }],
    [{ sheet: 'CSV', cell: 'B2', value: 'x', bold: true }],
    [{ sheet: 'CSV', cell: 'B1000', value: 'large extension' }],
    [{ sheet: 'Other', cell: 'A1', value: 'wrong sheet' }],
    [{ sheet: 'CSV', cell: 'A1', value: 'x'.repeat(2 * 1024 ** 2) }],
    [{ type: 'setCells', sheet: 'CSV', cells: Array.from({ length: 10001 }, () => ({ cell: 'A1', value: 'x' })) }],
  ])
    await assert.rejects(owner.apply('0', commands));
  assert.equal(owner.diagnosticStats().retainedStates, 1);
  assert.equal((await read(owner, '0', 'B2')).B2.value, 'original');
  await assert.rejects(owner.apply('missing', [{ sheet: 'CSV', cell: 'A1', value: 'x' }]), {
    code: 'internal-unknown-reference',
  });
  await assert.rejects(owner.read('0', 'Other', 'A1'), { code: 'internal-unknown-reference' });
});

test('CSV row-index, row projection and reply budgets fail without materializing a complete source view', async (t) => {
  const root = await fs.mkdtemp(path.join(tmpdir(), 'workbook-csv-index-budget-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  const source = path.join(root, 'source.csv'),
    storageDirectory = path.join(root, 'owners');
  await fs.writeFile(source, 'a\nb\n');
  const rejected = new CsvFileSession();
  await assert.rejects(rejected.open(source, { storageDirectory, maxIndexBytes: 20 }), {
    code: /^(csv-limit|filter-limit)$/,
  });
  assert.deepEqual(await fs.readdir(storageDirectory), []);
  assert.equal(rejected.diagnosticStats().closed, true);
  const { owner } = await fixture(t, 'head,value\nrow,' + 'x'.repeat(3000) + '\n', { maxReplyBytes: 1024 });
  assert.equal(owner.view('0').sheets[0].rows, 2);
  await assert.rejects(owner.read('0', 'CSV', 'B2'), { code: /^(csv-limit|layout-limit|note-limit)$/ });
  assert.equal((await read(owner, '0', 'A1')).A1.value, 'head');
  const limited = await fixture(t, 'head,value\nrow,' + 'x'.repeat(3000) + '\n', { maxProjectionBytes: 1024 });
  await assert.rejects(limited.owner.read('0', 'CSV', 'B2'), { code: 'csv-limit' });
});

test('CSV owner scan/query pages, literal search and selection preserve bounded protocol behavior', async (t) => {
  const { owner } = await fixture(t, Array.from({ length: 600 }, (_, i) => `${i},row${i}`).join('\n'));
  const first = await owner.scan('0', 'CSV');
  assert.equal(Object.keys(first.cells).length, 1000);
  assert(first.nextCursor);
  const next = await owner.scan('0', 'CSV', first.nextCursor);
  assert.equal(Object.keys(next.cells).length, 200);
  assert.equal(next.nextCursor, null);
  const query = await owner.query('0', 'CSV', 'row', 'raw', undefined, { range: 'B3:B5', matchLimit: 1 });
  assert.deepEqual(Object.keys(query.cells), ['B3']);
  const second = await owner.query('0', 'CSV', 'row', 'raw', query.nextCursor, { range: 'B3:B5', matchLimit: 1 });
  assert.deepEqual(Object.keys(second.cells), ['B4']);
  const summary = await owner.selectionSummary('0', 'CSV', 'A1:B600');
  assert.equal(summary.count, 1200);
  assert.equal(summary.numericCount, 0);
  assert.equal(summary.sum, null);
  await assert.rejects(owner.query('0', 'CSV', 'row', 'raw', undefined, { range: 'B3:B5:C6' }), /range/);
  await assert.rejects(owner.query('0', 'CSV', 'row', 'raw', undefined, { matchLimit: 1001 }), /limit/);
  await assert.rejects(owner.selectionSummary('0', 'CSV', 'A1:A2:A3'), /range/);
  const edited = await owner.apply('0', [{ sheet: 'CSV', cell: 'A1', value: '=1+1' }]);
  assert.equal((await owner.query(state(edited), 'CSV', '=1+1', 'formula')).matches, 0);
  assert.equal((await owner.query(state(edited), 'CSV', '=1+1', 'raw')).matches, 1);
});

test('indexed grid clear commands remain literal, bounded and atomic across mixed batches', async (t) => {
  const { owner } = await fixture(t);
  const result = await owner.apply('0', [
    { type: 'clearRange', sheet: 'CSV', range: 'A2:B3' },
    { sheet: 'CSV', cell: 'A2', value: 'after clear' },
  ]);
  assert.equal((await read(owner, result.view.storage.state, 'A2')).A2.value, 'after clear');
  assert.equal((await read(owner, result.view.storage.state, 'B3')).B3.value, '');
  assert.equal((await read(owner, '0', 'B3')).B3.value, 'tail');
  await assert.rejects(
    owner.apply(result.view.storage.state, [{ type: 'clearRange', sheet: 'CSV', range: 'A1:A10001' }]),
    /10,000/
  );
  await assert.rejects(
    owner.apply(result.view.storage.state, [
      { type: 'clearRange', sheet: 'CSV', range: 'A1:B3' },
      { sheet: 'CSV', cell: 'A2', expectedValue: 'wrong', value: 'bad' },
    ]),
    /Conflict/
  );
  assert.equal((await read(owner, result.view.storage.state, 'A2')).A2.value, 'after clear');
});

test('CSV exports obey disk quotas and compaction releases only unretained candidates', async (t) => {
  const original = 'id,value\n1,' + 'x'.repeat(80 * 1024) + '\n';
  const { owner, storageDirectory } = await fixture(t, original, { maxSessionBytes: 128 * 1024 });
  const one = await owner.apply('0', [{ sheet: 'CSV', cell: 'A1', value: 'first' }]),
    two = await owner.apply('0', [{ sheet: 'CSV', cell: 'A1', value: 'second' }]);
  const output = await owner.export(state(one));
  const disk = owner.diagnosticStats().exportBytes;
  await assert.rejects(owner.export(state(two)), /disk budget/);
  assert.equal(owner.diagnosticStats().exportBytes, disk);
  assert.equal(
    (await fs.readdir(path.join(storageDirectory, (await fs.readdir(storageDirectory))[0]))).filter((n) =>
      n.startsWith('export-')
    ).length,
    1
  );
  assert.equal(hash(await fs.readFile(output.path)), output.revision);
  await owner.compact([state(two)]);
  assert.equal(owner.diagnosticStats().exportBytes, 0);
  const second = await owner.export(state(two));
  assert.equal(hash(await fs.readFile(second.path)), second.revision);
});

test('CSV cached reads and saves reject mutated or replaced snapshots, including identical-byte replacement', async (t) => {
  const { owner, source } = await fixture(t);
  await read(owner, '0', 'B2');
  const bytes = await fs.readFile(source);
  await fs.unlink(source);
  await fs.writeFile(source, bytes);
  await assert.rejects(owner.read('0', 'CSV', 'B2'), /source changed/);
  await assert.rejects(owner.export('0'), /source changed/);
  assert.equal(owner.diagnosticStats().exportBytes, 0);
  assert.equal(owner.diagnosticStats().retainedStates, 1);
});

test('copy-time stat guard rejects same-byte mutation and removes only its own failed candidate', async (t) => {
  const root = await fs.mkdtemp(path.join(tmpdir(), 'workbook-csv-copy-stamp-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  const source = path.join(root, 'source.csv'),
    target = path.join(root, 'target.csv'),
    bytes = Buffer.from('id,value\n1,original\n');
  await fs.writeFile(source, bytes);
  const index = readCsvFileIndex(source);
  const originalWrite = syncFs.writeSync;
  let changed = false;
  try {
    t.mock.method(syncFs, 'writeSync', (...args) => {
      const result = originalWrite(...args);
      if (!changed) {
        changed = true;
        syncFs.writeFileSync(source, bytes);
      }
      return result;
    });
    syncBuiltinESMExports();
    assert.throws(() => writeCsvIndex(index, source, target), /source changed/);
  } finally {
    t.mock.restoreAll();
    syncBuiltinESMExports();
  }
  assert.equal(
    await fs.stat(target).then(
      () => true,
      () => false
    ),
    false
  );
  assert.deepEqual(await fs.readFile(source), bytes);
  const fresh = readCsvFileIndex(source);
  await fs.writeFile(target, 'keep');
  assert.throws(() => writeCsvIndex(fresh, source, target), /exist/i);
  assert.equal(await fs.readFile(target, 'utf8'), 'keep');
});
