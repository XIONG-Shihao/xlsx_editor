import { preservedFixture } from './support/preserved-calculation-fixture.mjs';
import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook, readWorkbook } from '../core/workbook-package.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { createWorkbookNativeClient } from '../core/native-client.mjs';
import { closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
const client = createWorkbookNativeClient(new URL('../native/bin/darwin-arm64/workbook-engine', import.meta.url).pathname);
after(() => {
  client.close();
  closeNativeWorkbookWorker();
});
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

for (const kind of ['shared', 'cache', 'iterative', 'opaque'])
  test(kind + ' allows an unrelated ordinary edit, retains formulas/caches and blocks unsafe changes', async (t) => {
    const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-preserved-calc-')),
      file = path.join(dir, 'source.xlsx'),
      bytes = preservedFixture(kind);
    await fs.writeFile(file, bytes);
    const book = new IndexedWorkbook(
      client.call,
      () => assert.fail('No fallback'),
      (input) => calculateIndexedWorkbook(input, client.call)
    );
    t.after(async () => {
      await book.close();
      await fs.rm(dir, { recursive: true, force: true });
    });
    await book.open(file);
    assert.equal(book.view('0').readOnly, false);
    assert.equal(book.view('0').calculationPolicy.preserveSource, true);
    assert.equal(book.view('0').calculationPolicy.canRecalculate, false);
    assert.deepEqual(await fs.readFile((await book.export('0')).path), bytes);
    const next = await book.apply('0', [{ sheet: 'Sheet1', cell: 'B1', value: 7 }]);
    assert.equal((await book.read(next.view.storage.state, 'Sheet1', 'B1')).cells.B1.value, 7);
    const pending = (await book.read(next.view.storage.state, 'Sheet1', 'A1')).cells.A1;
    assert.equal(pending.calculation.pending, true);
    assert.equal(pending.cachedResult.value, 2);
    for (const command of [
      { sheet: 'Sheet1', cell: 'A1', value: 99 },
      { sheet: 'Sheet1', cell: 'C1', value: 'flatten' },
      { sheet: 'Sheet1', cell: 'B1', formula: '2+2' },
      { sheet: 'Sheet1', operation: 'insertRow', index: 0 },
      { type: 'recalculate' },
    ])
      await assert.rejects(book.apply(next.view.storage.state, [command]), /preserv|protect|existing ordinary/i);
    assert.equal((await book.read('0', 'Sheet1', 'B1')).cells.B1.value, 1);
    const saved = await fs.readFile((await book.export(next.view.storage.state)).path),
      a = unzipSync(bytes),
      b = unzipSync(saved);
    for (const part of Object.keys(a))
      if (!['xl/workbook.xml', 'xl/worksheets/sheet1.xml'].includes(part)) assert.deepEqual(b[part], a[part], part);
    assert(
      strFromU8(b['xl/worksheets/sheet1.xml']).includes(
        strFromU8(a['xl/worksheets/sheet1.xml']).match(/<c r="A1">.*?<\/c>/)[0]
      )
    );
    assert(strFromU8(b['xl/workbook.xml']).includes('fullCalcOnLoad="1"'));
    assert(strFromU8(b['xl/workbook.xml']).includes('forceFullCalc="1"'));
    if (kind === 'iterative') assert(strFromU8(b['xl/workbook.xml']).includes('iterateCount="23"'));
    if (kind === 'opaque') assert(strFromU8(b['xl/workbook.xml']).includes('<u:child value="&gt;"/>'));
    assert.deepEqual(await fs.readFile(file), bytes);
    assert.equal(readWorkbook(saved).sheets[0].cells.B1.value, 7);
  });
test('Document Save, Undo/Redo, second Save and cold reopen retain preserved calculation ownership', async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-preserved-document-')),
    file = path.join(dir, 'source.xlsx');
  await fs.writeFile(file, preservedFixture('cache'));
  const owner = new SpreadsheetDocument({ recoveryDirectory: path.join(dir, 'recovery'), cacheIdentity: null });
  t.after(async () => {
    await owner.close();
    await fs.rm(dir, { recursive: true, force: true });
  });
  let info = await owner.open(file);
  const apply = (value) =>
    owner.request({ operation: 'apply', state: info.state, commands: [{ sheet: 'Sheet1', cell: 'B1', value }] });
  info = await apply(7);
  await owner.request({ operation: 'save' });
  info = owner.status();
  assert.equal(info.view.calculationPolicy.preserveSource, true);
  assert.equal(info.dirty, false);
  await owner.request({ operation: 'undo' });
  await owner.request({ operation: 'redo' });
  info = owner.status();
  info = await apply(8);
  await owner.request({ operation: 'save' });
  await owner.close();
  const cold = new SpreadsheetDocument({ recoveryDirectory: path.join(dir, 'cold'), cacheIdentity: null });
  t.after(() => cold.close());
  const opened = await cold.open(file);
  assert.equal(opened.view.readOnly, false);
  assert.equal(opened.view.calculationPolicy.preserveSource, true);
  const read = await cold.request({ operation: 'read', state: opened.state, sheet: 'Sheet1', range: 'A1:B1' });
  assert.equal(read.cells.B1.value, 8);
  assert.equal(read.cells.A1.cachedResult.value, 2);
});
test('native export rejects forged formula authority and unrelated metadata replacement in preservation mode', async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-preserved-forge-')),
    file = path.join(dir, 'source.xlsx'),
    bytes = preservedFixture('cache');
  await fs.writeFile(file, bytes);
  let nativeRoot, nativeVersion;
  const call = async (request) => {
    if (request.operation === 'preservation-point-check' || request.method === 'preserve') nativeRoot = request.root;
    const reply = await client.call(request);
    if (request.method === 'preserve') nativeVersion = reply.version;
    return reply;
  };
  const book = new IndexedWorkbook(
    call,
    () => assert.fail('No fallback'),
    (i) => calculateIndexedWorkbook(i, call)
  );
  t.after(async () => {
    await book.close();
    await fs.rm(dir, { recursive: true, force: true });
  });
  await book.open(file);
  const { indexedCalculationMetadata } = await import('../core/workbook-package.mjs');
  const metadata = indexedCalculationMetadata(
    await fs.readFile(path.join(nativeRoot, 'metadata.xlsx')),
    true,
    book.view('0').calculationPolicy
  );
  metadata.styleIds = [];
  const point = await book.apply('0', [{ sheet: 'Sheet1', cell: 'B1', value: 7 }]);
  const goodVersion = nativeVersion;
  const baseline = await client.call({
    action: 'workbook',
    operation: 'export-plan',
    root: nativeRoot,
    version: goodVersion,
  });
  assert.equal(baseline.styleCount, 0);
  const extra = {
    ...metadata,
    replacements: {
      ...metadata.replacements,
      '[Content_Types].xml': strFromU8(unzipSync(bytes)['[Content_Types].xml']),
    },
  };
  await client.call({
    action: 'workbook',
    operation: 'artifact-write',
    root: nativeRoot,
    name: 'export-metadata.json',
    data: JSON.stringify(extra),
  });
  await assert.rejects(
    client.call({
      action: 'workbook',
      operation: 'export',
      root: nativeRoot,
      version: goodVersion,
      revision: book.view('0').revision,
      calculationMetadata: true,
      preservationExport: true,
    }),
    { code: /^(package-unsupported|pivot-stale)$/ }
  );
  await assert.rejects(
    client.call({
      action: 'workbook',
      operation: 'patch',
      root: nativeRoot,
      version: goodVersion,
      changes: [
        {
          sheet: 'xl/worksheets/sheet1.xml',
          cell: 'B1',
          value: { formula: '2+2', value: null, cachedValue: null, dataType: 'formula', styleId: 0 },
          edits: [{ formula: '2+2' }],
        },
      ],
    }),
    { code: 'formula-protected' }
  );
  assert.equal((await book.read(point.view.storage.state, 'Sheet1', 'B1')).cells.B1.value, 7);
  assert.deepEqual(await fs.readFile(file), bytes);
});
test('compaction retains preservation authority and pending results without keeping the initial state', async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-preserved-compact-')),
    file = path.join(dir, 'source.xlsx');
  await fs.writeFile(file, preservedFixture('shared'));
  const book = new IndexedWorkbook(
    client.call,
    () => assert.fail('No fallback'),
    (i) => calculateIndexedWorkbook(i, client.call)
  );
  t.after(async () => {
    await book.close();
    await fs.rm(dir, { recursive: true, force: true });
  });
  await book.open(file);
  const changed = await book.apply('0', [{ sheet: 'Sheet1', cell: 'B1', value: 7 }]),
    state = changed.view.storage.state;
  await book.compact([state], { retainBase: false });
  assert.throws(() => book.view('0'), { code: 'internal-unknown-reference' });
  assert.equal((await book.read(state, 'Sheet1', 'A1')).cells.A1.calculation.pending, true);
  const saved = await book.export(state);
  assert.equal(readWorkbook(await fs.readFile(saved.path)).sheets[0].cells.B1.value, 7);
});
