/** Scalar array proofs: actual retained scalar cache + primary-spec bounded vectors. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { defaultPrintSettings } from '../core/print-model.mjs';
import { composePrint, printDocumentHtml } from '../core/print-layout.mjs';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { createWorkbook, editWorkbook, readWorkbook, indexedCalculationMetadata } from '../core/workbook-package.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
const sha = (b) => createHash('sha256').update(b).digest('hex');
async function owner(t, bytes) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-scalar-array-')),
    source = path.join(directory, 'source.xlsx');
  await fs.writeFile(source, bytes);
  let root, version;
  const bridge = async (input) => {
    const result = await native(input);
    if (input.operation === 'index') root = input.root;
    if (input.operation === 'range') {
      root = input.root;
      version = input.version;
    }
    return result;
  };
  const make = () =>
    new IndexedWorkbook(
      bridge,
      () => assert.fail('No shared formula writer'),
      (input) => calculateIndexedWorkbook(input, bridge)
    );
  const book = make();
  t.after(async () => {
    await book.close();
    await fs.rm(directory, { recursive: true, force: true });
  });
  await book.open(source);
  return { book, make, source, bytes, root: () => root, version: () => version, directory };
}
function fixture() {
  let bytes = editWorkbook(
    createWorkbook(),
    [
      ['A1', 1],
      ['A2', 2],
      ['A3', 2],
      ['A4', 4],
      ['A5', 5],
      ['B1', 'One'],
      ['B2', 'First two'],
      ['B3', 'Last two'],
      ['B4', '#N/A'],
      ['B5', false],
      ['C1', 10],
      ['C2', 20],
      ['C3', 30],
      ['C4', 40],
      ['C5', 50],
      ['J1', 3],
    ].map(([cell, value]) => ({ sheet: 'Sheet1', cell, value }))
  );
  return editWorkbook(
    bytes,
    [
      ['D1', 'SUM(SEQUENCE(3,1))'],
      ['D2', 'SUM(SEQUENCE(2,2,10,2))'],
      ['D3', 'SUM(UNIQUE(A1:A5))'],
      ['D4', 'XLOOKUP(2,A1:A5,C1:C5)'],
      ['D5', 'XLOOKUP(2,A1:A5,C1:C5,0,0,-1)'],
      ['D6', 'XLOOKUP(2,A1:A5,B1:B5)'],
      ['D7', 'XLOOKUP(9,A1:A5,B1:B5,"Missing")'],
      ['D8', 'XLOOKUP(9,A1:A5,B1:B5)'],
      ['D9', 'XLOOKUP(4,A1:A5,B1:B5)'],
      ['D10', 'XLOOKUP(5,A1:A5,B1:B5)'],
      ['D11', 'SUM(SEQUENCE(J1,1))'],
      ['E1', 'D1+D3'],
      ['E2', 'LEN(D6)'],
      ['E3', 'IFERROR(D8,99)'],
      ['E4', 'DEC2HEX(D4,4)'],
      ['F1', 'FILTER(A1:A5,A1:A5>1)'],
    ].map(([cell, formula]) => ({ sheet: 'Sheet1', cell, formula }))
  );
}
const expected = {
  D1: 6,
  D2: 52,
  D3: 12,
  D4: 20,
  D5: 30,
  D6: 'First two',
  D7: 'Missing',
  D9: '#N/A',
  D10: false,
  D11: 6,
  E1: 18,
  E2: 9,
  E3: 99,
  E4: '0014',
};
async function check(book, state, values = expected) {
  const cells = (await book.read(state, 'Sheet1', 'D1:F11')).cells;
  for (const [ref, value] of Object.entries(values)) assert.equal(cells[ref].calculation.value, value, ref);
  assert.equal(cells.D8.calculation.status, 'error');
  assert.equal(cells.D8.calculation.error, '#N/A');
  assert.equal(cells.D9.calculation.scalarArray.valueType, 'text');
  assert.equal(cells.F1.calculation.status, 'unresolved');
  for (const ref of ['D1', 'D2', 'D3', 'D4', 'D5', 'D6', 'D7', 'D8', 'D9', 'D10', 'D11'])
    assert.equal(cells[ref].calculation.scalarArray.schema, 1, ref);
  return cells;
}
test('bounded scalar roots retain typed results through mixed dependents, edit, old states, replay/export/reopen', async (t) => {
  const f = await owner(t, fixture());
  await check(f.book, '0');
  const changed = await f.book.apply('0', [
      { sheet: 'Sheet1', cell: 'A3', value: 7 },
      { sheet: 'Sheet1', cell: 'J1', value: 4 },
    ]),
    state = changed.view.storage.state;
  await check(f.book, state, { ...expected, D3: 19, D5: 20, D11: 10, E1: 25 });
  await check(f.book, '0');
  assert.deepEqual(new Uint8Array(await fs.readFile(f.source)), new Uint8Array(f.bytes));
  const output = await f.book.export(state),
    saved = await fs.readFile(output.path);
  const disk = readWorkbook(saved);
  assert.equal(disk.sheets[0].cells.D3.cachedValue, 19);
  assert.match(disk.sheets[0].cells.D1.formula, /_xlfn.SEQUENCE/);
  assert.match(disk.sheets[0].cells.D4.formula, /^_xlfn.XLOOKUP/);
  const opened = f.make();
  t.after(() => opened.close());
  await opened.open(output.path);
  await check(opened, '0', { ...expected, D3: 19, D5: 20, D11: 10, E1: 25 });
  const replay = f.make();
  t.after(() => replay.close());
  await replay.open(f.source);
  await check(replay, (await replay.apply('0', changed.edits)).view.storage.state, {
    ...expected,
    D3: 19,
    D5: 20,
    D11: 10,
    E1: 25,
  });
  const readback = spawnSync(
    'python3',
    [
      '-c',
      `import sys,zipfile,xml.etree.ElementTree as E
n='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
with zipfile.ZipFile(sys.argv[1]) as z:
 c={x.get('r'):x for x in E.fromstring(z.read('xl/worksheets/sheet1.xml')).iter(n+'c')}
 for r,v in [('D1',6),('D2',52),('D3',19),('D4',20),('D5',20),('D11',10)]:assert float(c[r].find(n+'v').text)==v and c[r].get('t','n')=='n'
 assert c['D8'].get('t')=='e' and c['D8'].find(n+'v').text=='#N/A'
 assert c['D9'].get('t')=='str' and c['D9'].find(n+'v').text=='#N/A'
 assert c['D10'].get('t')=='b' and c['D10'].find(n+'v').text=='0'
`,
      output.path,
    ],
    { encoding: 'utf8' }
  );
  assert.equal(readback.status, 0, readback.stderr);
});
test('retained Excel singleton SUM(SEQUENCE) adopts source cm without changing authored formula/metadata geometry', async (t) => {
  const json = JSON.parse(
      await fs.readFile(new URL('../fixtures/excel-dynamic-reference.json', import.meta.url), 'utf8')
    ),
    bytes = await fs.readFile(new URL('../fixtures/excel-dynamic-reference.xlsx', import.meta.url));
  assert.equal(sha(bytes), json.oracleWorkbookSha256);
  const f = await owner(t, bytes);
  const cell = (await f.book.read('0', 'Cases', 'A47')).cells.A47;
  assert.equal(cell.calculation.status, 'calculated');
  assert.equal(cell.calculation.value, 6);
  assert.equal(cell.calculation.scalarArray.schema, 1);
  assert.equal(cell.formulaType, 'array');
  assert.equal(cell.formulaReference, 'A47');
  const out = await fs.readFile((await f.book.export('0')).path),
    parts = unzipSync(out),
    old = unzipSync(bytes);
  assert.deepEqual(parts['xl/metadata.xml'], old['xl/metadata.xml']);
  const before = readWorkbook(bytes).sheets[0].cells.A47,
    after = readWorkbook(out).sheets[0].cells.A47;
  assert.equal(after.formula, before.formula);
  assert.equal(after.formulaReference, 'A47');
  assert.equal(after.cachedValue, 6);
  assert.deepEqual(new Uint8Array(await fs.readFile(f.source)), new Uint8Array(bytes));
});
test('unsupported array/coercion/search/overflow contexts remain unresolved with original source preserved', async (t) => {
  const formulas = [
    'SUM(SEQUENCE(1001))',
    'SUM(SEQUENCE(1.5))',
    'SUM(SEQUENCE(2,1,9007199254740991,1))',
    'SUM(UNIQUE(A1:B5))',
    'SUM(UNIQUE(A1:A5,FALSE,TRUE))',
    'XLOOKUP(2,A1:A5,C1:C5,0,1)',
    'XLOOKUP(2,A1:A5,C1:C5,0,0,2)',
    'XLOOKUP(A1:A2,A1:A5,C1:C5)',
    'XLOOKUP(2,A1:A5,B1:C5)',
    'IF(TRUE,SUM(SEQUENCE(3)),0)',
  ];
  const f = await owner(
    t,
    editWorkbook(
      fixture(),
      formulas.map((formula, i) => ({ sheet: 'Sheet1', cell: 'H' + (i + 1), formula }))
    )
  );
  const cells = (await f.book.read('0', 'Sheet1', 'H1:H10')).cells;
  for (let i = 0; i < formulas.length; i++) {
    assert.equal(cells['H' + (i + 1)].calculation.status, 'unresolved', formulas[i]);
    assert.equal(cells['H' + (i + 1)].calculation.scalarArray, undefined);
    assert.equal(cells['H' + (i + 1)].formula, formulas[i]);
  }
  const edited = await f.book.apply('0', [{ sheet: 'Sheet1', cell: 'A2', value: '2' }]);
  const newCells = (await f.book.read(edited.view.storage.state, 'Sheet1', 'D1:D11')).cells;
  assert.equal(newCells.D3.calculation.status, 'unresolved');
  assert.equal(newCells.D4.calculation.status, 'calculated');
  assert.equal(newCells.D4.calculation.value, 30);
  assert.equal(newCells.D4.calculation.scalarArray.schema, 1);
  await check(f.book, '0');
  assert.deepEqual(new Uint8Array(await fs.readFile(f.source)), new Uint8Array(f.bytes));
});
test('native export refuses stale proof values, changed literal dependencies and prepared-cache substitution atomically', async (t) => {
  const f = await owner(t, fixture()),
    cell = (await f.book.read('0', 'Sheet1', 'D1')).cells.D1,
    root = f.root(),
    version = f.version();
  const bad = structuredClone(cell);
  bad.calculation.value = 7;
  const forged = await native({
    action: 'workbook',
    operation: 'patch',
    root,
    version,
    changes: [{ sheet: 'xl/worksheets/sheet1.xml', cell: 'D1', value: bad }],
  });
  await assert.rejects(native({ action: 'workbook', operation: 'export-plan', root, version: forged.version }), {
    code: 'spill-stale',
  });
  await check(f.book, '0');
  assert.deepEqual(new Uint8Array(await fs.readFile(f.source)), new Uint8Array(f.bytes));
  const stale = await native({
    action: 'workbook',
    operation: 'patch',
    root,
    version,
    changes: [
      {
        sheet: 'xl/worksheets/sheet1.xml',
        cell: 'A2',
        value: { value: 8, dataType: 'number', rawValue: '8', cachedValue: null, styleId: 0 },
        edits: [{ value: 8 }],
      },
    ],
  });
  await assert.rejects(native({ action: 'workbook', operation: 'export-plan', root, version: stale.version }), {
    code: 'spill-stale',
  });
  const missing = structuredClone(cell);
  delete missing.calculation.scalarArray;
  const noProof = await native({
    action: 'workbook',
    operation: 'patch',
    root,
    version,
    changes: [{ sheet: 'xl/worksheets/sheet1.xml', cell: 'D1', value: missing }],
  });
  await assert.rejects(native({ action: 'workbook', operation: 'export-plan', root, version: noProof.version }), {
    code: 'spill-bad-input',
  });
  const output = await f.book.export('0'),
    accepted = await fs.readFile(output.path),
    revision = f.book.view('0').revision;
  await f.book.read('0', 'Sheet1', 'D1');
  const streamRoot = f.root(),
    streamVersion = f.version();
  const exportPlan = await native({
    action: 'workbook',
    operation: 'export-plan',
    root: streamRoot,
    version: streamVersion,
  });
  await native({
    action: 'workbook',
    operation: 'artifact-write',
    root: streamRoot,
    name: 'export-metadata.json',
    data: JSON.stringify({
      ...indexedCalculationMetadata(
        await fs.readFile(path.join(streamRoot, 'metadata.xlsx')),
        exportPlan.incomplete,
        f.book.view('0').calculationPolicy
      ),
      styleIds: [],
    }),
  });
  const altered = spawnSync(
    'python3',
    [
      '-c',
      `import sqlite3,json,sys
db=sqlite3.connect(sys.argv[1]);part='xl/worksheets/sheet1.xml';v=json.loads(db.execute('SELECT data FROM export_changes WHERE sheet=? AND r=0 AND c=3',(part,)).fetchone()[0]);v['calculation']['value']=999;db.execute('UPDATE export_changes SET data=? WHERE sheet=? AND r=0 AND c=3',(json.dumps(v),part));db.commit()
`,
      path.join(streamRoot, 'cells.sqlite'),
    ],
    { encoding: 'utf8' }
  );
  assert.equal(altered.status, 0, altered.stderr);
  await assert.rejects(
    native({
      action: 'workbook',
      operation: 'export',
      root: streamRoot,
      version: streamVersion,
      revision,
      calculationMetadata: true,
    }),
    { code: /^(internal-mismatch|spill-bad-input)$/ }
  );
  assert.deepEqual(new Uint8Array(await fs.readFile(output.path)), new Uint8Array(accepted));
  assert.deepEqual(new Uint8Array(await fs.readFile(f.source)), new Uint8Array(f.bytes));
  const plan = await native({ action: 'workbook', operation: 'export-plan', root: streamRoot, version: streamVersion });
  assert.equal(typeof plan.incomplete, 'boolean');
});

test('imported singleton scalar-array source content clear refuses while ordinary authored roots clear coherently', async (t) => {
  const bytes = await fs.readFile(new URL('../fixtures/excel-dynamic-reference.xlsx', import.meta.url)),
    f = await owner(t, bytes);
  const view = f.book.view('0'),
    count = f.book.diagnosticStats().retainedStates;
  for (const range of ['A47', 'A47:B47'])
    await assert.rejects(f.book.apply('0', [{ type: 'clearRange', sheet: 'Cases', range }]), {
      code: 'fill-unsupported',
    });
  assert.deepEqual(f.book.view('0'), view);
  assert.equal(f.book.diagnosticStats().retainedStates, count);
  assert.equal((await f.book.read('0', 'Cases', 'A47')).cells.A47.calculation.value, 6);
  assert.deepEqual(new Uint8Array(await fs.readFile(f.source)), new Uint8Array(bytes));
  const ordinary = await owner(t, fixture()),
    cleared = await ordinary.book.apply('0', [{ type: 'clearRange', sheet: 'Sheet1', range: 'D1' }]);
  assert.equal((await ordinary.book.read(cleared.view.storage.state, 'Sheet1', 'D1')).cells.D1.formula, undefined);
  assert.equal((await ordinary.book.read('0', 'Sheet1', 'D1')).cells.D1.calculation.value, 6);
});

test('1000-position lookup proofs remain compact and support repeated paged reads/edits with exact source guards', async (t) => {
  const input = [];
  for (let i = 1; i <= 1000; i++) {
    input.push({ sheet: 'Sheet1', cell: 'A' + i, value: i }, { sheet: 'Sheet1', cell: 'B' + i, value: i * 2 });
  }
  let bytes = editWorkbook(createWorkbook(), input);
  bytes = editWorkbook(
    bytes,
    Array.from({ length: 65 }, (_, i) => ({
      sheet: 'Sheet1',
      cell: 'D' + (i + 1),
      formula: 'XLOOKUP(999,A1:A1000,B1:B1000)',
    }))
  );
  const f = await owner(t, bytes),
    cells = (await f.book.read('0', 'Sheet1', 'D1:D65')).cells;
  for (const cell of Object.values(cells)) {
    assert.equal(cell.calculation.value, 1998);
    assert.equal(cell.calculation.scalarArray.inputCells, 2000);
    assert.match(cell.calculation.scalarArray.inputSha256, /^[0-9a-f]{64}$/);
    assert.ok(JSON.stringify(cell.calculation.scalarArray).length < 1024);
    assert.equal(cell.calculation.scalarArray.inputs, undefined);
  }
  const edited = await f.book.apply('0', [{ sheet: 'Sheet1', cell: 'B999', value: 2998 }]);
  assert.equal((await f.book.read(edited.view.storage.state, 'Sheet1', 'D65')).cells.D65.calculation.value, 2998);
  assert.equal((await f.book.read('0', 'Sheet1', 'D65')).cells.D65.calculation.value, 1998);
  assert.deepEqual(new Uint8Array(await fs.readFile(f.source)), new Uint8Array(bytes));
});

test('native scalar-array print proof revalidates typed roots/inputs and preserves old snapshots; imported singleton cm prints without follower claims', async (t) => {
  const input = [
    ['A1', 1],
    ['A2', 2],
    ['A3', 2],
    ['A4', 4],
    ['A5', 5],
    ['B1', 'One'],
    ['B2', 'First two'],
    ['B3', 'Last two'],
    ['B4', '#N/A'],
    ['B5', false],
  ].map(([cell, value]) => ({ sheet: 'Sheet1', cell, value }));
  let bytes = editWorkbook(createWorkbook(), input);
  bytes = editWorkbook(
    bytes,
    [
      ['D1', 'SUM(SEQUENCE(3,1))'],
      ['D2', 'SUM(UNIQUE(A1:A5))'],
      ['D3', 'XLOOKUP(2,A1:A5,B1:B5)'],
      ['D4', 'XLOOKUP(5,A1:A5,B1:B5)'],
    ].map(([cell, formula]) => ({ sheet: 'Sheet1', cell, formula }))
  );
  const printable = unzipSync(bytes);
  printable['xl/worksheets/sheet1.xml'] = strToU8(
    strFromU8(printable['xl/worksheets/sheet1.xml']).replace(
      '<sheetData',
      '<sheetFormatPr defaultRowHeight="24"/><cols><col min="4" max="4" width="14" customWidth="1"/></cols><sheetData'
    )
  );
  bytes = Buffer.from(zipSync(printable));
  const f = await owner(t, bytes),
    settings = { ...defaultPrintSettings(), area: 'D1:D4' },
    snapshot = await f.book.printSnapshot('0', 'Sheet1', settings);
  assert.deepEqual(
    snapshot.cells.map((cell) => cell.text),
    ['6', '12', 'First two', 'false']
  );
  const frozen = structuredClone(snapshot),
    changed = await f.book.apply('0', [{ sheet: 'Sheet1', cell: 'B2', value: 'Changed' }]),
    state = changed.view.storage.state;
  assert.equal((await f.book.printSnapshot(state, 'Sheet1', settings)).cells[2].text, 'Changed');
  assert.deepEqual(snapshot, frozen);
  assert.deepEqual((await f.book.printSnapshot('0', 'Sheet1', settings)).cells, snapshot.cells);
  await f.book.read('0', 'Sheet1', 'D1:D4');
  const root = f.root(),
    version = f.version(),
    view = f.book.view('0');
  const stale = await native({
    action: 'workbook',
    operation: 'patch',
    root,
    version,
    changes: [
      {
        sheet: 'xl/worksheets/sheet1.xml',
        cell: 'B2',
        value: { value: 'Forged input', dataType: 'text', rawValue: null, cachedValue: null, styleId: 0 },
        edits: [{ value: 'Forged input' }],
      },
    ],
  });
  await assert.rejects(
    native({
      action: 'workbook',
      operation: 'print-check',
      root,
      version: stale.version,
      sheet: 'xl/worksheets/sheet1.xml',
      range: 'D1:D4',
      view,
    }),
    { code: 'spill-stale' }
  );
  assert.deepEqual(new Uint8Array(await fs.readFile(f.source)), new Uint8Array(bytes));
  const actual = await fs.readFile(new URL('../fixtures/excel-dynamic-reference.xlsx', import.meta.url)),
    parts = unzipSync(actual),
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    part = 'xl/worksheets/sheet1.xml',
    doc = new DOMParser().parseFromString(strFromU8(parts[part]), 'application/xml');
  for (const cell of Array.from(doc.getElementsByTagNameNS(ns, 'c')))
    if (cell.getAttribute('r') !== 'A47') cell.parentNode.removeChild(cell);
  for (const row of Array.from(doc.getElementsByTagNameNS(ns, 'row')))
    if (!row.childNodes.length) row.parentNode.removeChild(row);
  parts[part] = strToU8(new XMLSerializer().serializeToString(doc));
  const g = await owner(t, Buffer.from(zipSync(parts))),
    imported = await g.book.printSnapshot('0', 'Cases', { ...defaultPrintSettings(), area: 'A47:A47' });
  assert.deepEqual(
    imported.cells.map((cell) => cell.text),
    ['6']
  );
  assert.deepEqual(new Uint8Array(await fs.readFile(g.source)), new Uint8Array(g.bytes));
  const layout = composePrint(snapshot);
  assert.equal(layout.pages.length, 1);
  assert.ok(printDocumentHtml(layout).includes('First two'));
});
