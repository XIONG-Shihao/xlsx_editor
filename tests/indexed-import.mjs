import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { randomUUID, createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook } from '../core/index.mjs';
import { readWorkbookMetadata, indexedCalculationMetadata } from '../core/workbook-package.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { createWorkbookNativeClient } from '../core/native-client.mjs';
import { nativeWorkbookBinary } from '../core/native-runtime.mjs';
import { applyIndexedImport } from '../core/indexed-import.mjs';
import { createPasteStager } from '../adapters/large-paste.mjs';
import { inspectPaste, pastePages } from '../ui/largePaste.ts';
import { rangeBounds } from '../core/workbook-commands.mjs';

const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const digest = (bytes) => createHash('sha256').update(bytes).digest('hex');
const installedHelper = nativeWorkbookBinary(),
  helperBytes = await fs.readFile(installedHelper),
  helperHash = digest(helperBytes);
const helperDirectory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-import-helper-')),
  helper = path.join(helperDirectory, path.basename(installedHelper));
await fs.writeFile(helper, helperBytes, { mode: 0o700 });
after(() => fs.rm(helperDirectory, { recursive: true, force: true }));
async function fixture(t, body = '', textStyle = false) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-indexed-import-')),
    client = createWorkbookNativeClient(helper);
  const parts = unzipSync(createWorkbook());
  parts['xl/workbook.xml'] = strToU8(
    strFromU8(parts['xl/workbook.xml']).replace(
      '</workbook>',
      '<calcPr calcId="162194" calcMode="auto" calcOnSave="1" fullCalcOnLoad="0" forceFullCalc="0"/></workbook>'
    )
  );
  parts['xl/worksheets/sheet1.xml'] = strToU8(`<worksheet xmlns="${ns}"><sheetData>${body}</sheetData></worksheet>`);
  parts['custom/opaque.bin'] = new Uint8Array([3, 0, 255, 19]);
  if (textStyle)
    parts['xl/styles.xml'] = strToU8(
      strFromU8(parts['xl/styles.xml'])
        .replace('<cellXfs count="1">', '<cellXfs count="2">')
        .replace('</cellXfs>', '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>')
    );
  const source = path.join(directory, 'source.xlsx'),
    bytes = zipSync(parts);
  await fs.writeFile(source, bytes);
  let root,
    index,
    peakRead = 0,
    peakPatch = 0,
    peakBytes = 0,
    patchCalls = 0;
  const native = async (request) => {
    if (request.operation === 'range') {
      peakRead = Math.max(peakRead, rangeBounds(request.range).size);
      assert(rangeBounds(request.range).size <= 1000);
    }
    if (request.operation === 'patch') {
      patchCalls++;
      peakPatch = Math.max(peakPatch, request.changes.length);
      peakBytes = Math.max(peakBytes, Buffer.byteLength(JSON.stringify(request.changes)));
      assert(request.changes.length <= 256);
      assert(request.changes.every((change) => Array.isArray(change.edits) && change.edits.length === 1));
    }
    const result = await client.call(request);
    if (request.operation === 'index') {
      root = request.root;
      index = result;
    }
    return result;
  };
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('Unexpected legacy replay'),
    async (input) => ({ version: input.version })
  );
  await book.open(source);
  const { styles } = readWorkbookMetadata(await fs.readFile(path.join(root, 'metadata.xlsx')));
  const current = { root, index, view: book.view('0'), styles, version: 0 };
  const stager = createPasteStager({ directory: path.join(directory, 'imports'), documentId: randomUUID() });
  t.after(async () => {
    await stager.close();
    await book.close();
    client.close();
    await fs.rm(directory, { recursive: true, force: true });
  });
  return {
    source,
    bytes,
    root,
    index,
    book,
    current,
    native,
    stager,
    stats: () => ({ peakRead, peakPatch, peakBytes, patchCalls }),
  };
}
async function prepare(f, text, options = {}) {
  const plan = inspectPaste(text, { sheet: 'Sheet1', start: 'A1', ...options }),
    { id } = await f.stager.begin({ ...plan, state: 'base' });
  for (const page of pastePages(text, plan)) await f.stager.upload({ id, state: 'base', ...page });
  const descriptor = await f.stager.seal({ id, state: 'base' });
  return { descriptor, pages: f.stager.pages(descriptor, { state: 'base' }) };
}
const decode = (cell) =>
  cell.value && typeof cell.value === 'object' && '$workbookNumber' in cell.value
    ? Number(cell.value.$workbookNumber)
    : cell.value;
const read = async (f, version, range) =>
  Object.fromEntries(
    (
      await f.native({
        action: 'workbook',
        operation: 'range',
        root: f.root,
        version,
        sheet: 'xl/worksheets/sheet1.xml',
        range,
        offset: 0,
        limit: 1000,
      })
    ).cells
  );

test('actual 100,000-cell paste stages and applies bounded native requests, retains source/old state and exports every value', async (t) => {
  const f = await fixture(t),
    accepted = structuredClone(f.current.view),
    text = ('1\t'.repeat(99) + '1\n').repeat(1000),
    prepared = await prepare(f, text),
    progress = [];
  const result = await applyIndexedImport({
    ...prepared,
    current: f.current,
    native: f.native,
    progress: (value) => progress.push(value.completed),
  });
  assert.equal(result.editCount, 100000);
  assert.deepEqual(result.changedRanges, [{ sheet: 'Sheet1', range: 'A1:CV1000' }]);
  assert.equal(result.view.sheets[0].rows, 1000);
  assert.equal(result.view.sheets[0].columns, 100);
  assert.equal(Object.keys(result.view.sheets[0].cells).length, 0);
  assert.deepEqual(result.journal, prepared.descriptor);
  assert(Buffer.byteLength(JSON.stringify(result.journal)) < 1024);
  assert.equal(result.diagnostics.peakPageCells, 1000);
  assert.equal(result.diagnostics.patchCalls, 400);
  assert.equal(progress.length, 100);
  assert.equal(progress.at(-1), 100000);
  assert(f.stats().peakRead <= 1000);
  assert.equal(f.stats().peakPatch, 256);
  assert(f.stats().peakBytes <= 1024 * 1024);
  const summary = await f.native({
    action: 'workbook',
    operation: 'selection',
    root: f.root,
    version: result.version,
    sheet: 'xl/worksheets/sheet1.xml',
    range: 'A1:CV1000',
  });
  assert.equal(summary.count, 100000);
  assert.equal(summary.numericCount, 100000);
  assert.equal(summary.sum, 100000);
  assert.equal((await f.book.selectionSummary('0', 'Sheet1', 'A1:CV1000')).count, 0);
  assert.equal(f.current.version, 0);
  assert.deepEqual(f.current.view, accepted);
  const exportPlan = await f.native({
    action: 'workbook',
    operation: 'export-plan',
    root: f.root,
    version: result.version,
  });
  assert.equal(exportPlan.special, false);
  assert.equal(exportPlan.styleCount, 0);
  await f.native({
    action: 'workbook',
    operation: 'artifact-write',
    root: f.root,
    name: 'export-metadata.json',
    data: JSON.stringify({
      ...indexedCalculationMetadata(
        await fs.readFile(path.join(f.root, 'metadata.xlsx')),
        exportPlan.incomplete,
        result.view.calculationPolicy
      ),
      styleIds: [],
    }),
  });
  const exported = await f.native({
    action: 'workbook',
    operation: 'export',
    root: f.root,
    version: result.version,
    revision: f.index.revision,
    calculationMetadata: true,
  });
  const savedPath = path.join(f.root, exported.file),
    savedBytes = await fs.readFile(savedPath);
  const oracle = JSON.parse(
    execFileSync('python3', ['tests/support/staged-paste-oracle.py', f.source, savedPath, '1000', '100'], {
      encoding: 'utf8',
      maxBuffer: 1024 * 1024,
    })
  );
  assert.equal(oracle.savedCellCount, 100000);
  assert.equal(oracle.savedNumericOneCount, 100000);
  assert.equal(oracle.allUnchangedPackageMembersPreserved, true);
  assert.equal(oracle.numericTypeVerified, true);
  assert.equal(digest(await fs.readFile(f.source)), digest(f.bytes));
});

test('target text style, signed zero, exact identifiers and ragged holes retain existing cell semantics', async (t) => {
  const f = await fixture(
    t,
    '<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>old</t></is></c></row><row r="2"><c r="B2"><v>99</v></c></row>',
    true
  );
  const prepared = await prepare(f, '001\tTRUE\t-0\t1234567890123456\t=A1\nshort');
  const result = await applyIndexedImport({ ...prepared, current: f.current, native: f.native }),
    cells = await read(f, result.version, 'A1:E2');
  assert.equal(cells.A1.value, '001');
  assert.equal(cells.A1.styleId, 1);
  assert.equal(cells.B1.value, true);
  assert(Object.is(decode(cells.C1), -0));
  assert.equal(cells.D1.value, '1234567890123456');
  assert.equal(cells.E1.formula, 'A1');
  assert.equal(cells.A2.value, 'short');
  assert.equal(cells.B2.value, 99);
  assert.equal(result.formulaDelta, 1);
  assert.equal(result.formulaChanged, true);
  assert.equal(result.view.sheets[0].storedFormulaCount, 1);
  assert.equal((await read(f, 0, 'A1:B2')).A1.value, 'old');
  assert.equal(digest(await fs.readFile(f.source)), digest(f.bytes));
});

test('a populated complete 1000-position target page is bounded despite the native continuation sentinel', async (t) => {
  const body = Array.from({ length: 1000 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>7</v></c></row>`).join('');
  const f = await fixture(t, body),
    prepared = await prepare(f, '2\n'.repeat(1000));
  const result = await applyIndexedImport({ ...prepared, current: f.current, native: f.native });
  assert.equal(result.editCount, 1000);
  assert.equal((await read(f, result.version, 'A1000')).A1000.value, 2);
  assert.equal((await read(f, 0, 'A1000')).A1000.value, 7);
  assert(f.stats().peakRead <= 1000);
});

test('target page duplication, out-of-range addresses and incomplete continuation never reach a native patch', async (t) => {
  for (const mutant of [
    (cells) => [...cells.slice(0, -1), cells[0]],
    (cells) => [...cells.slice(0, -1), ['B1000', cells.at(-1)[1]]],
    (cells) => cells.slice(0, -1),
  ]) {
    const f = await fixture(t),
      prepared = await prepare(f, '2\n'.repeat(1000));
    let patches = 0;
    const native = async (input) => {
      if (input.operation === 'patch') {
        patches++;
        throw Error('Unexpected patch');
      }
      if (input.operation === 'range')
        return {
          cells: mutant(Array.from({ length: 1000 }, (_, i) => ['A' + (i + 1), { value: 1 }])),
          nextOffset: 1000,
        };
      return f.native(input);
    };
    await assert.rejects(
      applyIndexedImport({ ...prepared, current: f.current, native }),
      /duplicate|out-of-range|declared bound/
    );
    assert.equal(patches, 0);
  }
});

test('a protected array found after the first 1,000 cells rejects without publishing earlier candidate writes', async (t) => {
  const body =
    '<row r="1001"><c r="A1001"><f t="array" ref="A1001:A1002">1+1</f><v>2</v></c></row><row r="1002"><c r="A1002"><v>2</v></c></row>';
  const f = await fixture(t, body),
    prepared = await prepare(f, '1\n'.repeat(1001));
  await assert.rejects(
    applyIndexedImport({ ...prepared, current: f.current, native: f.native }),
    /array|special|qualified|partial/i
  );
  assert(f.stats().patchCalls >= 4);
  assert.equal((await f.book.selectionSummary('0', 'Sheet1', 'A1:A1000')).count, 0);
  assert.equal((await read(f, 0, 'A1001:A1002')).A1002.value, 2);
  assert.equal(f.current.version, 0);
  assert.equal(digest(await fs.readFile(f.source)), digest(f.bytes));
  await f.native({ action: 'workbook', operation: 'compact', root: f.root, keep: [0] });
});

test('confirmed-page cancellation and malformed late replay preserve accepted metadata, source and prior versions', async (t) => {
  const f = await fixture(t),
    prepared = await prepare(f, '1\n'.repeat(1001));
  let cancelled = false;
  await assert.rejects(
    applyIndexedImport({
      ...prepared,
      current: f.current,
      native: f.native,
      checkCancelled: () => {
        if (cancelled) throw Error('cancelled');
      },
      progress: () => {
        cancelled = true;
      },
    }),
    /cancelled/
  );
  assert.equal((await f.book.selectionSummary('0', 'Sheet1', 'A1:A1000')).count, 0);
  assert.equal(f.current.version, 0);
  const malformed = {
    async *[Symbol.asyncIterator]() {
      yield { sequence: 0, cells: [[0, '1']] };
      yield { sequence: 2, cells: [[1, '2']] };
    },
  };
  await assert.rejects(
    applyIndexedImport({ descriptor: prepared.descriptor, pages: malformed, current: f.current, native: f.native }),
    { code: /^(csv-refused|package-bad-input|paste-bad-input)$/ }
  );
  assert.equal(digest(await fs.readFile(f.source)), digest(f.bytes));
  assert.equal(f.current.view.sheets[0].storedFormulaCount, 0);
});

test('the actual native patch guard protects unmarked array followers and permits unrelated ordinary targets', async (t) => {
  const body =
    '<row r="1001"><c r="A1001"><f t="array" ref="A1001:A1002">1+1</f><v>2</v></c></row><row r="1002"><c r="A1002"><v>2</v></c></row>';
  const f = await fixture(t, body),
    follower = await prepare(f, 'overwritten', { start: 'A1002' });
  await assert.rejects(
    applyIndexedImport({ ...follower, current: f.current, native: f.native }),
    /array|partial|special|qualified/i
  );
  assert.equal(f.stats().patchCalls, 1);
  const ordinary = await prepare(f, 'ordinary'),
    result = await applyIndexedImport({ ...ordinary, current: f.current, native: f.native });
  assert.equal((await read(f, result.version, 'A1')).A1.value, 'ordinary');
  assert.equal((await read(f, result.version, 'A1001:A1002')).A1002.value, 2);
  assert.equal((await read(f, 0, 'A1')).A1, undefined);
  assert.equal(digest(await fs.readFile(f.source)), digest(f.bytes));
});

test('actual native storage quota failure preserves the source and accepted owner revision', async (t) => {
  const f = await fixture(t),
    prepared = await prepare(f, '1\n'.repeat(1001));
  const lease = path.join(path.dirname(f.root), 'owner.json'),
    original = await fs.readFile(lease);
  try {
    await fs.writeFile(lease, JSON.stringify({ ...JSON.parse(original), maxSessionBytes: 128 * 1024 }));
    await assert.rejects(
      applyIndexedImport({ ...prepared, current: f.current, native: f.native }),
      /budget|disk|full|reserve/i
    );
  } finally {
    await fs.writeFile(lease, original);
  }
  assert.equal((await f.book.selectionSummary('0', 'Sheet1', 'A1:A1000')).count, 0);
  assert.equal(f.current.version, 0);
  assert.equal(digest(await fs.readFile(f.source)), digest(f.bytes));
});
