/** N6: complete current numeric spill proof, including off-page masters. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createHash } from 'node:crypto';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { createWorkbook } from '../core/workbook-package.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { defaultPrintSettings } from '../core/print-model.mjs';
import { composePrint, printDocumentHtml } from '../core/print-layout.mjs';
after(() => closeNativeWorkbookWorker());
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  part = 'xl/worksheets/sheet1.xml';
const settings = (area) => ({ ...defaultPrintSettings(), area }),
  sha = (bytes) => createHash('sha256').update(bytes).digest('hex');
async function owner(t, bytes) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-spill-print-')),
    source = path.join(directory, 'source.xlsx');
  await fs.writeFile(source, bytes);
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  let root, version;
  const calls = [];
  const bridge = async (input) => {
    const reply = await native(input);
    calls.push(input);
    if (input.operation === 'index') root = input.root;
    if (input.operation === 'range') version = input.version;
    return reply;
  };
  const book = new IndexedWorkbook(
    bridge,
    () => assert.fail('No legacy writer'),
    (input) => calculateIndexedWorkbook(input, bridge)
  );
  t.after(() => book.close());
  await book.open(source);
  return { book, source, bytes, directory, calls, root: () => root, version: () => version };
}
function authored({ formula = 'SEQUENCE($J$1,2)', blocker = '', height = 15 } = {}) {
  const parts = unzipSync(createWorkbook());
  parts[part] = strToU8(
    `<worksheet xmlns="${ns}"><sheetFormatPr defaultRowHeight="${height}"/><sheetData><row r="1"><c r="A1"><f>${formula}</f></c>${blocker}<c r="D1"><f>SUM(_xlfn.ANCHORARRAY(A1))</f></c><c r="J1"><v>3</v></c></row><row r="2"><c r="D2"><f>INDEX(_xlfn.ANCHORARRAY(A1),3,2)</f></c></row></sheetData></worksheet>`
  );
  parts['customXml/n6-opaque.xml'] = strToU8('<opaque> untouched </opaque>');
  return Buffer.from(zipSync(parts));
}
const cellTexts = (snapshot) => Object.fromEntries(snapshot.cells.map((cell) => [cell.ref, cell.text]));
test('owned off-page master prints exact followers, resize releases them and old snapshots stay immutable', async (t) => {
  const f = await owner(t, authored()),
    before = f.book.diagnosticStats().retainedStates,
    snapshot = await f.book.printSnapshot('0', 'Sheet1', settings('A2:B3'));
  assert.deepEqual(cellTexts(snapshot), { A2: '3', B2: '4', A3: '5', B3: '6' });
  assert.equal(f.book.diagnosticStats().retainedStates, before);
  const frozen = structuredClone(snapshot);
  const changed = await f.book.apply('0', [{ sheet: 'Sheet1', cell: 'J1', value: 2 }]),
    state = changed.view.storage.state,
    small = await f.book.printSnapshot(state, 'Sheet1', settings('A1:B3'));
  assert.deepEqual(cellTexts(small), { A1: '1', B1: '2', A2: '3', B2: '4', A3: '', B3: '' });
  assert.deepEqual(snapshot, frozen);
  assert.deepEqual(cellTexts(await f.book.printSnapshot('0', 'Sheet1', settings('D1:D2'))), { D1: '21', D2: '6' });
  const cleared = await f.book.apply(state, [{ sheet: 'Sheet1', cell: 'A1', value: null }]);
  assert.ok(
    (await f.book.printSnapshot(cleared.view.storage.state, 'Sheet1', settings('A1:B3'))).cells.every(
      (cell) => cell.text === ''
    )
  );
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
  assert.ok(f.calls.filter((call) => call.operation === 'print-check').every((call) => call.view && call.range));
});
test('retained actual source singleton consumers and owned arrays print, export and reopen with exact21/6', async (t) => {
  const oracle = JSON.parse(await fs.readFile('fixtures/excel-dynamic-reference.json', 'utf8')),
    source = await fs.readFile('fixtures/excel-dynamic-source.xlsx');
  assert.equal(sha(source), oracle.sourceSha256);
  const parts = unzipSync(source),
    document = new DOMParser().parseFromString(strFromU8(parts[part]), 'application/xml'),
    keep = new Set(['A1', 'D1', 'D2']);
  // Deterministic source minimization only: raw formulas/cm and every other
  // package member are retained; no generated expected caches are inserted.
  for (const cell of Array.from(document.getElementsByTagNameNS(ns, 'c')))
    if (!keep.has(cell.getAttribute('r'))) cell.parentNode.removeChild(cell);
  for (const row of Array.from(document.getElementsByTagNameNS(ns, 'row')))
    if (!row.childNodes.length) row.parentNode.removeChild(row);
  parts[part] = strToU8(new XMLSerializer().serializeToString(document));
  const f = await owner(t, Buffer.from(zipSync(parts))),
    snapshot = await f.book.printSnapshot('0', 'Cases', settings('A1:D3'));
  for (const ref of ['A1', 'B1', 'A2', 'B2', 'A3', 'B3', 'D1', 'D2'])
    assert.equal(cellTexts(snapshot)[ref], String(oracle.cells[ref].expected));
  const output = await f.book.export('0'),
    saved = unzipSync(await fs.readFile(output.path));
  assert.ok(strFromU8(saved[part]).includes('cm="1"'));
  const reopened = new IndexedWorkbook(
    native,
    () => assert.fail(),
    (input) => calculateIndexedWorkbook(input, native)
  );
  t.after(() => reopened.close());
  await reopened.open(output.path);
  assert.deepEqual((await reopened.printSnapshot('0', 'Cases', settings('A1:D3'))).cells, snapshot.cells);
  assert.equal(sha(await fs.readFile('fixtures/excel-dynamic-source.xlsx')), oracle.sourceSha256);
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
});
test('page cuts retain complete group proof and every numeric cell occurs once in physical pages', async (t) => {
  const f = await owner(t, authored({ height: 300 })),
    snapshot = await f.book.printSnapshot('0', 'Sheet1', settings('A1:B3')),
    layout = composePrint(snapshot);
  assert.equal(layout.pages.length, 2);
  const cells = layout.pages.flatMap((page) => page.cells);
  assert.equal(new Set(cells.map((cell) => cell.ref)).size, 6);
  assert.deepEqual(Object.fromEntries(cells.map((cell) => [cell.ref, cell.text])), {
    A1: '1',
    B1: '2',
    A2: '3',
    B2: '4',
    A3: '5',
    B3: '6',
  });
  assert.deepEqual(
    layout.pages.map((page) => page.rowIndices),
    [[0, 1], [2]]
  );
  const html = printDocumentHtml(layout);
  assert.equal((html.match(/data-xlsx-print-page=/g) ?? []).length, 2);
});
test('unowned/blocked and imported legacy off-page masters refuse instead of printing cached blank followers', async (t) => {
  for (const bytes of [
    authored({ formula: 'SEQUENCE(10)' }),
    authored({ blocker: '<c r="B1"><v>0</v></c>' }),
    (() => {
      const parts = unzipSync(createWorkbook());
      parts[part] = strToU8(
        `<worksheet xmlns="${ns}"><sheetData><row r="1"><c r="A1"><f t="array" ref="A1:A5">ROW(A1:A5)</f><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row></sheetData></worksheet>`
      );
      return Buffer.from(zipSync(parts));
    })(),
  ]) {
    const f = await owner(t, bytes),
      count = f.book.diagnosticStats().retainedStates;
    await f.book.read('0', 'Sheet1', 'A2:B3');
    await assert.rejects(
      f.book.printSnapshot('0', 'Sheet1', settings('A2:B3')),
      /PRINT_SPECIAL_FORMULA|PRINT_UNRESOLVED_CALCULATION|Special/
    );
    await assert.rejects(
      native({
        action: 'workbook',
        operation: 'print-check',
        root: f.root(),
        version: f.version(),
        sheet: part,
        range: 'A2:B3',
        view: f.book.view('0'),
      }),
      { code: 'print-unsupported' }
    );
    assert.equal(f.book.diagnosticStats().retainedStates, count);
    assert.deepEqual(await fs.readFile(f.source), f.bytes);
  }
});
test('native accepted-version proof rejects stale constructor and forged consumer cache without touching prior state', async (t) => {
  const f = await owner(t, authored()),
    cells = (await f.book.read('0', 'Sheet1', 'A1:J3')).cells,
    version = f.version(),
    view = f.book.view('0'),
    base = await f.book.printSnapshot('0', 'Sheet1', settings('A1:B3'));
  for (const [reference, cell] of [
    ['J1', { value: 2, dataType: 'number' }],
    ['D1', { ...cells.D1, calculation: { ...cells.D1.calculation, value: 999 } }],
  ]) {
    const patched = await native({
      action: 'workbook',
      operation: 'patch',
      root: f.root(),
      version,
      changes: [{ sheet: part, cell: reference, value: cell }],
    });
    await assert.rejects(
      native({
        action: 'workbook',
        operation: 'print-check',
        root: f.root(),
        version: patched.version,
        sheet: part,
        range: 'A2:B3',
        view,
      }),
      /stale|forged/
    );
    assert.deepEqual((await f.book.printSnapshot('0', 'Sheet1', settings('A1:B3'))).cells, base.cells);
  }
  await assert.rejects(
    native({
      action: 'workbook',
      operation: 'print-check',
      root: f.root(),
      version,
      sheet: part,
      range: 'A1:XFD1',
      view,
    }),
    /1,000/
  );
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
  // Consumer admission happened to reject omitted metadata in the fixture
  // above. A producer-only worksheet must independently fence the same request.
  const parts = unzipSync(authored());
  parts[part] = strToU8(strFromU8(parts[part]).replace(/<c r="D[12]">[\s\S]*?<\/c>/g, ''));
  const g = await owner(t, Buffer.from(zipSync(parts)));
  await g.book.read('0', 'Sheet1', 'A1:J3');
  const patched = await native({
    action: 'workbook',
    operation: 'patch',
    root: g.root(),
    version: g.version(),
    changes: [{ sheet: part, cell: 'J1', value: { value: 2, dataType: 'number' } }],
  });
  await assert.rejects(
    native({
      action: 'workbook',
      operation: 'print-check',
      root: g.root(),
      version: patched.version,
      sheet: part,
      range: 'A2:B3',
      view: { ...g.book.view('0'), sheets: [] },
    }),
    { code: 'print-unsupported' }
  );
  assert.deepEqual(await fs.readFile(g.source), g.bytes);
  assert.equal(g.book.diagnosticStats().retainedStates, 1);
});
