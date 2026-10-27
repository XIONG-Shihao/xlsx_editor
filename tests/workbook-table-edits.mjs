/** Ordinary table definition changes, compact replay and exact native admission. */
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook, editWorkbook, readWorkbook } from '../core/workbook-package.mjs';
import { indexedTablePlan } from '../core/indexed-tables.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
const sha = (bytes) => createHash('sha256').update(bytes).digest('hex'),
  tablePart = 'xl/tables/table1.xml';
function fixture() {
  let bytes = editWorkbook(
    createWorkbook(),
    [
      ['A1', 'Value'],
      ['B1', 'Cost'],
      ['A2', 2],
      ['B2', 10],
      ['A3', 3],
      ['B3', 20],
      ['A4', 4],
      ['B4', 30],
    ].map(([cell, value]) => ({ sheet: 'Sheet1', cell, value }))
  );
  bytes = editWorkbook(bytes, [
    { operation: 'createTable', sheet: 'Sheet1', range: 'A1:B4', tableName: 'Prices' },
    { operation: 'addSheet', sheet: 'Other' },
  ]);
  bytes = editWorkbook(bytes, [
    { sheet: 'Sheet1', cell: 'D1', formula: 'SUM(Prices[Value])' },
    { sheet: 'Sheet1', cell: 'D2', formula: 'SUM(Prices[Cost])' },
    { sheet: 'Sheet1', cell: 'C2', formula: 'prices[[#This Row],[Value]]*2' },
    { sheet: 'Sheet1', cell: 'E1', formula: 'D1+1' },
    { sheet: 'Sheet1', cell: 'F1', formula: 'DEC2HEX(D1,4)' },
    { sheet: 'Sheet1', cell: 'G1', formula: '"Prices[Value]"' },
    { sheet: 'Other', cell: 'A1', formula: 'SUM(PRICES[Value])' },
  ]);
  const parts = unzipSync(bytes);
  parts['customXml/table-opaque.xml'] = strToU8('<opaque> unchanged </opaque>');
  return Buffer.from(zipSync(parts));
}
async function owner(t, bytes = fixture()) {
  bytes = Buffer.from(bytes);
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-table-edit-')),
    source = path.join(directory, 'source.xlsx');
  await fs.writeFile(source, bytes);
  let root, version;
  const calls = [];
  const bridge = async (input) => {
    const result = await native(input);
    calls.push(input);
    if (input.operation === 'index') root = input.root;
    if (input.operation === 'range') {
      version = input.version;
      root = input.root;
    }
    return result;
  };
  const make = () =>
    new IndexedWorkbook(
      bridge,
      () => assert.fail('No shared writer'),
      (input) => calculateIndexedWorkbook(input, bridge)
    );
  const book = make();
  t.after(async () => {
    await book.close();
    await fs.rm(directory, { recursive: true, force: true });
  });
  await book.open(source);
  return { book, make, bytes, directory, source, calls, root: () => root, version: () => version };
}
const rename = (tableName = 'Prices', newName = 'Rates', expectedTableRef = 'A1:B4') => ({
  operation: 'renameTable',
  sheet: 'Sheet1',
  tableName,
  newName,
  expectedTableRef,
});
const resize = (range, tableName = 'Prices', expectedTableRef = 'A1:B4') => ({
  operation: 'resizeTable',
  sheet: 'Sheet1',
  tableName,
  range,
  expectedTableRef,
});
async function values(book, state, sum, cost) {
  const cells = (await book.read(state, 'Sheet1', 'C1:G4')).cells;
  assert.equal(cells.D1.calculation.value, sum);
  assert.equal(cells.D2.calculation.value, cost);
  assert.equal(cells.C2.calculation.value, 4);
  assert.equal(cells.E1.calculation.value, sum + 1);
  assert.equal(cells.F1.calculation.value, sum.toString(16).toUpperCase().padStart(4, '0'));
  assert.equal(cells.G1.calculation.value, 'Prices[Value]');
  assert.equal((await book.read(state, 'Other', 'A1')).cells.A1.calculation.value, sum);
  return cells;
}
test('table rename updates workbook-wide parsed references, preserves strings/cells and replays one compact action', async (t) => {
  const f = await owner(t);
  await values(f.book, '0', 9, 60);
  const result = await f.book.apply('0', [rename()]),
    state = result.view.storage.state,
    cells = await values(f.book, state, 9, 60);
  assert.deepEqual(result.commands, [rename()]);
  assert.ok(JSON.stringify(result.commands).length < 256);
  assert.equal(result.view.tables[0].name, 'Rates');
  assert.equal(cells.D1.formula, 'SUM(Rates[Value])');
  assert.equal(cells.C2.formula, 'Rates[[#This Row],[Value]]*2');
  assert.equal(cells.G1.formula, '"Prices[Value]"');
  assert.equal((await f.book.read(state, 'Other', 'A1')).cells.A1.formula, 'SUM(Rates[Value])');
  assert.equal((await f.book.read('0', 'Sheet1', 'D1')).cells.D1.formula, 'SUM(Prices[Value])');
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
  const output = await f.book.export(state),
    saved = await fs.readFile(output.path),
    disk = readWorkbook(saved);
  assert.equal(disk.tables[0].name, 'Rates');
  assert.equal(disk.sheets[0].cells.D1.formula, cells.D1.formula);
  const parts = unzipSync(saved),
    old = unzipSync(f.bytes);
  assert.deepEqual(parts['customXml/table-opaque.xml'], old['customXml/table-opaque.xml']);
  for (const ref of ['A1', 'B1', 'A2', 'B2', 'A3', 'B3', 'A4', 'B4'])
    assert.deepEqual(disk.sheets[0].cells[ref], readWorkbook(f.bytes).sheets[0].cells[ref]);
  const replay = f.make();
  t.after(() => replay.close());
  await replay.open(f.source);
  const again = await replay.apply('0', result.commands);
  await values(replay, again.view.storage.state, 9, 60);
  const reopened = f.make();
  t.after(() => reopened.close());
  await reopened.open(output.path);
  await values(reopened, '0', 9, 60);
  const inspect = spawnSync(
    'python3',
    [
      '-c',
      `import sys,zipfile,xml.etree.ElementTree as E
n='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
with zipfile.ZipFile(sys.argv[1]) as z:
 t=E.fromstring(z.read('xl/tables/table1.xml'));assert t.attrib['name']==t.attrib['displayName']=='Rates' and t.attrib['ref']=='A1:B4'
 for part,ref,formula,value in [('xl/worksheets/sheet1.xml','D1','SUM(Rates[Value])','9'),('xl/worksheets/sheet2.xml','A1','SUM(Rates[Value])','9')]:
  c=next(c for c in E.fromstring(z.read(part)).iter(n+'c') if c.attrib['r']==ref);assert c.find(n+'f').text==formula and float(c.find(n+'v').text)==float(value)
`,
      output.path,
    ],
    { encoding: 'utf8' }
  );
  assert.equal(inspect.status, 0, inspect.stderr);
});
test('table row shrink recalculates structured dependents and preserves removed membership cells/old states', async (t) => {
  const f = await owner(t),
    result = await f.book.apply('0', [resize('A1:B3')]),
    state = result.view.storage.state;
  await values(f.book, state, 5, 30);
  await values(f.book, '0', 9, 60);
  assert.equal(result.view.tables[0].ref, 'A1:B3');
  assert.deepEqual(result.commands, [resize('A1:B3')]);
  const body = (await f.book.read(state, 'Sheet1', 'A4:B4')).cells;
  assert.equal(body.A4.value, 4);
  assert.equal(body.B4.value, 30);
  const saved = unzipSync(await fs.readFile((await f.book.export(state)).path));
  assert.match(strFromU8(saved[tablePart]), /ref="A1:B3"/);
  assert.match(strFromU8(saved[tablePart]), /<autoFilter ref="A1:B3"/);
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
  const replay = f.make();
  t.after(() => replay.close());
  await replay.open(f.source);
  await values(replay, (await replay.apply('0', result.commands)).view.storage.state, 5, 30);
});
test('blank row growth preserves style-only cells and rebuilds dependencies for later literal input', async (t) => {
  const parts = unzipSync(fixture());
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    strFromU8(parts['xl/worksheets/sheet1.xml']).replace(
      '</sheetData>',
      '<row r="5"><c r="A5" s="0"/></row></sheetData>'
    )
  );
  const f = await owner(t, Buffer.from(zipSync(parts))),
    grown = await f.book.apply('0', [resize('A1:B6')]),
    state = grown.view.storage.state;
  await values(f.book, state, 9, 60);
  assert.equal(grown.view.tables[0].ref, 'A1:B6');
  const populated = await f.book.apply(state, [
    { sheet: 'Sheet1', cell: 'A5', value: 5 },
    { sheet: 'Sheet1', cell: 'B5', value: 40 },
  ]);
  await values(f.book, populated.view.storage.state, 14, 100);
  await values(f.book, state, 9, 60);
  assert.equal((await f.book.read(state, 'Sheet1', 'A5')).cells.A5.styleId, 0);
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
});
test('stale refs, absorption, changed columns/origin, dynamic/special dependencies and native delta forgeries refuse atomically', async (t) => {
  const f = await owner(t);
  const shrunk = await f.book.apply('0', [resize('A1:B3')]),
    state = shrunk.view.storage.state,
    count = f.book.diagnosticStats().retainedStates;
  for (const command of [
    rename('Prices', 'OtherName', 'A1:B2'),
    resize('A1:C4'),
    resize('A2:B4'),
    resize('A1:B1'),
    rename('Prices', 'A1'),
  ])
    await assert.rejects(f.book.apply('0', [command]));
  await assert.rejects(f.book.apply(state, [resize('A1:B4', 'Prices', 'A1:B3')]), { code: /^(formula-bad-input|table-bad-input|table-unsupported)$/ });
  assert.equal(f.book.diagnosticStats().retainedStates, count);
  await values(f.book, '0', 9, 60);
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
  const metadata = await fs.readFile(path.join(f.root(), 'metadata.xlsx')),
    plan = await indexedTablePlan(metadata, f.book.view('0'), rename());
  await f.book.read('0', 'Sheet1', 'A1');
  const version = f.version(),
    source = await f.book.export('0');
  for (const mutate of [
    (p) => (p.priorTableSha256 = '0'.repeat(64)),
    (p) => (p.replacements['customXml/table-opaque.xml'] = '<changed/>'),
    (p) => (p.sheets = []),
    (p) => (p.replacements[tablePart] = p.replacements[tablePart].replace('TableStyleMedium2', 'TableStyleMedium9')),
  ]) {
    const bad = structuredClone(plan);
    mutate(bad);
    await native({
      action: 'workbook',
      operation: 'artifact-write',
      root: f.root(),
      name: 'package-metadata.json',
      data: JSON.stringify(bad),
    });
    await assert.rejects(
      native({
        action: 'workbook',
        operation: 'package-edit',
        root: f.root(),
        version,
        source: path.basename(source.path),
      }),
      /guard|only|scope|metadata/
    );
  }
  const dynamic = await owner(
    t,
    editWorkbook(fixture(), [{ sheet: 'Sheet1', cell: 'H1', formula: 'INDIRECT("Prices[Value]")' }])
  );
  await assert.rejects(dynamic.book.apply('0', [rename()]), /Dynamic/);
  assert.deepEqual(await fs.readFile(dynamic.source), dynamic.bytes);
});

test('ordinary table admission refuses opaque growth/header cells and malformed owner graphs without publishing', async (t) => {
  const sheetPart = 'xl/worksheets/sheet1.xml';
  const cases = [
    {
      name: 'growth opaque child',
      command: resize('A1:B5'),
      part: sheetPart,
      mutate: (xml) => xml.replace('</sheetData>', '<row r="5"><c r="A5"><notQualified/></c></row></sheetData>'),
      reason: /Opaque.*children/,
    },
    {
      name: 'growth cm',
      command: resize('A1:B5'),
      part: sheetPart,
      mutate: (xml) => xml.replace('</sheetData>', '<row r="5"><c r="A5" cm="1"/></row></sheetData>'),
      reason: /Opaque.*attributes/,
    },
    {
      name: 'header cm',
      command: rename(),
      part: sheetPart,
      mutate: (xml) => xml.replace('<c r="A1"', '<c cm="1" r="A1"'),
      reason: /Opaque.*attributes/,
    },
    {
      name: 'header mismatch',
      command: rename(),
      part: sheetPart,
      mutate: (xml) => xml.replace('>Value</t>', '>Wrong</t>'),
      reason: /headers/,
    },
    {
      name: 'owner count',
      command: rename(),
      part: sheetPart,
      mutate: (xml) => xml.replace('<tableParts count="1"', '<tableParts count="2"'),
      reason: /count/,
    },
    {
      name: 'duplicate default type',
      command: rename(),
      part: '[Content_Types].xml',
      mutate: (xml) => xml.replace('</Types>', '<Default Extension="xml" ContentType="application/xml"/></Types>'),
      reason: /Duplicate|content type|read-only/,
    },
  ];
  for (const entry of cases) {
    const parts = unzipSync(fixture()),
      before = strFromU8(parts[entry.part]),
      after = entry.mutate(before);
    assert.notEqual(after, before, entry.name + ' mutant applied');
    parts[entry.part] = strToU8(after);
    const f = await owner(t, Buffer.from(zipSync(parts))),
      view = f.book.view('0'),
      stats = f.book.diagnosticStats();
    await assert.rejects(f.book.apply('0', [entry.command]), entry.reason, entry.name);
    assert.deepEqual(f.book.view('0'), view);
    assert.equal(f.book.diagnosticStats().retainedStates, stats.retainedStates);
    assert.equal(sha(await fs.readFile(f.source)), sha(f.bytes));
  }
});

test('resizing refuses merge intersections, formula detach and name collisions while ordinary literals remain intact', async (t) => {
  const withFormula = await owner(t, editWorkbook(fixture(), [{ sheet: 'Sheet1', cell: 'A4', formula: '1+3' }]));
  await assert.rejects(withFormula.book.apply('0', [resize('A1:B3')]), { code: 'table-unsupported' });
  assert.equal((await withFormula.book.read('0', 'Sheet1', 'A4')).cells.A4.formula, '1+3');
  assert.deepEqual(await fs.readFile(withFormula.source), withFormula.bytes);
  const parts = unzipSync(fixture());
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    strFromU8(parts['xl/worksheets/sheet1.xml']).replace(
      '</worksheet>',
      '<mergeCells count="1"><mergeCell ref="A5:B5"/></mergeCells></worksheet>'
    )
  );
  const merged = await owner(t, Buffer.from(zipSync(parts)));
  await assert.rejects(merged.book.apply('0', [resize('A1:B5')]), /merge/);
  assert.equal(merged.book.view('0').tables[0].ref, 'A1:B4');
  let bytes = editWorkbook(fixture(), [
    { sheet: 'Sheet1', cell: 'J1', value: 'Other' },
    { sheet: 'Sheet1', cell: 'J2', value: 1 },
  ]);
  bytes = editWorkbook(bytes, [{ operation: 'createTable', sheet: 'Sheet1', range: 'J1:J2', tableName: 'Rates' }]);
  const collision = await owner(t, bytes);
  await assert.rejects(collision.book.apply('0', [rename()]), /collision/);
  assert.equal(collision.book.view('0').tables.find((table) => table.name === 'Prices').ref, 'A1:B4');
});

const renameColumn = (columnName = 'Value', newName = 'Unit Price', expectedTableRef = 'A1:B4') => ({
  operation: 'renameTableColumn',
  sheet: 'Sheet1',
  tableName: 'Prices',
  expectedTableRef,
  columnName,
  newName,
});
test('column rename updates its literal header and parsed workbook references as one compact replayable action', async (t) => {
  const f = await owner(t),
    before = await values(f.book, '0', 9, 60),
    result = await f.book.apply('0', [renameColumn()]),
    state = result.view.storage.state;
  const c = await values(f.book, state, 9, 60);
  assert.equal(c.D1.formula, 'SUM(Prices[Unit Price])');
  assert.equal(c.C2.formula, 'prices[[#This Row],[Unit Price]]*2');
  assert.equal(c.D2.formula, before.D2.formula);
  assert.equal(c.G1.formula, '"Prices[Value]"');
  assert.equal((await f.book.read(state, 'Other', 'A1')).cells.A1.formula, 'SUM(PRICES[Unit Price])');
  assert.equal((await f.book.read(state, 'Sheet1', 'A1')).cells.A1.value, 'Unit Price');
  assert.equal((await f.book.read('0', 'Sheet1', 'A1')).cells.A1.value, 'Value');
  assert.equal(result.view.tables[0].columns[0].name, 'Unit Price');
  assert.deepEqual(result.commands, [renameColumn()]);
  assert(JSON.stringify(result.commands).length < 256);
  await values(f.book, '0', 9, 60);
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
  const output = await f.book.export(state),
    disk = readWorkbook(await fs.readFile(output.path));
  assert.equal(disk.tables[0].columns[0].name, 'Unit Price');
  assert.equal(disk.sheets[0].cells.A1.value, 'Unit Price');
  for (const ref of ['B1', 'A2', 'B2', 'A3', 'B3', 'A4', 'B4'])
    assert.deepEqual(disk.sheets[0].cells[ref], readWorkbook(f.bytes).sheets[0].cells[ref]);
  const replay = f.make();
  t.after(() => replay.close());
  await replay.open(f.source);
  await values(replay, (await replay.apply('0', result.commands)).view.storage.state, 9, 60);
  const fresh = f.make();
  t.after(() => fresh.close());
  await fresh.open(output.path);
  await values(fresh, '0', 9, 60);
  const inspect = spawnSync(
    'python3',
    [
      '-c',
      `import sys,zipfile,xml.etree.ElementTree as E
n='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
with zipfile.ZipFile(sys.argv[1]) as z:
 t=E.fromstring(z.read('xl/tables/table1.xml'));assert t.attrib['name']=='Prices' and t.attrib['ref']=='A1:B4';assert [c.attrib['name'] for c in t.find(n+'tableColumns')]==['Unit Price','Cost']
 for part,ref,formula,value in [('xl/worksheets/sheet1.xml','D1','SUM(Prices[Unit Price])',9),('xl/worksheets/sheet2.xml','A1','SUM(PRICES[Unit Price])',9)]:
  c=next(c for c in E.fromstring(z.read(part)).iter(n+'c') if c.attrib['r']==ref);assert c.find(n+'f').text==formula and float(c.find(n+'v').text)==value
 c=next(c for c in E.fromstring(z.read('xl/worksheets/sheet1.xml')).iter(n+'c') if c.attrib['r']=='A1');assert c.attrib['t']=='inlineStr' and c.find(n+'is/'+n+'t').text=='Unit Price'
 assert z.read('customXml/table-opaque.xml')==b'<opaque> unchanged </opaque>'
`,
      output.path,
    ],
    { encoding: 'utf8' }
  );
  assert.equal(inspect.status, 0, inspect.stderr);
});
test('column rename captures caller input before queue and preserves styled headers and XML-special labels', async (t) => {
  const bytes = editWorkbook(fixture(), [
      { sheet: 'Sheet1', cell: 'A1', bold: true, fontColor: '11AA44', fillColor: 'FFEEDD' },
    ]),
    f = await owner(t, bytes),
    command = renameColumn('Value', 'Price & Tax'),
    pending = f.book.apply('0', [command]);
  command.newName = 'Changed after call';
  const result = await pending,
    state = result.view.storage.state;
  assert.equal(result.commands[0].newName, 'Price & Tax');
  const cell = (await f.book.read(state, 'Sheet1', 'A1')).cells.A1;
  assert.equal(cell.value, 'Price & Tax');
  assert.equal(cell.bold, true);
  assert.equal(cell.fontColor, '11AA44');
  assert.equal(cell.fillColor, 'FFEEDD');
  const disk = readWorkbook(await fs.readFile((await f.book.export(state)).path));
  assert.equal(disk.tables[0].columns[0].name, 'Price & Tax');
  assert.equal(disk.sheets[0].cells.A1.value, 'Price & Tax');
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
});
test('column rename refusals and a later batch failure do not publish partial metadata or formula/header edits', async (t) => {
  const f = await owner(t),
    baseline = f.book.view('0'),
    stats = f.book.diagnosticStats();
  for (const command of [
    renameColumn('Value', 'cost'),
    renameColumn('Missing', 'New'),
    renameColumn('Value', 'Value'),
    renameColumn('Value', 'Bad[Name]'),
    renameColumn('Value', ' Valid'),
    renameColumn('Value', 'New', 'A1:B3'),
    { ...renameColumn(), range: 'A1:B4' },
  ])
    await assert.rejects(f.book.apply('0', [command]));
  await assert.rejects(f.book.apply('0', [renameColumn(), { operation: 'unsupported', sheet: 'Sheet1' }]));
  assert.deepEqual(f.book.view('0'), baseline);
  assert.equal(f.book.diagnosticStats().retainedStates, stats.retainedStates);
  await values(f.book, '0', 9, 60);
  assert.equal((await f.book.read('0', 'Sheet1', 'A1')).cells.A1.value, 'Value');
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
  const opaque = await owner(
    t,
    editWorkbook(fixture(), [{ sheet: 'Sheet1', cell: 'H1', formula: 'INDIRECT("Prices[Value]")' }])
  );
  await assert.rejects(opaque.book.apply('0', [renameColumn()]), /Dynamic/);
  assert.equal((await opaque.book.read('0', 'Sheet1', 'A1')).cells.A1.value, 'Value');
  assert.deepEqual(await fs.readFile(opaque.source), opaque.bytes);
});
test('native column plan forgeries cannot change another header, style, table name or sheet scope', async (t) => {
  const f = await owner(t);
  await f.book.read('0', 'Sheet1', 'A1');
  const metadata = await fs.readFile(path.join(f.root(), 'metadata.xlsx')),
    plan = await indexedTablePlan(metadata, f.book.view('0'), renameColumn()),
    version = f.version(),
    source = await f.book.export('0');
  for (const mutate of [
    (p) => (p.renameTableColumn.index = 1),
    (p) => (p.renameTableColumn.new = 'Cost'),
    (p) => (p.sheets = []),
    (p) => (p.replacements[tablePart] = p.replacements[tablePart].replace('name="Cost"', 'name="Other"')),
    (p) => (p.replacements[tablePart] = p.replacements[tablePart].replace('TableStyleMedium2', 'TableStyleMedium9')),
    (p) => (p.renameTable = { old: 'Prices', new: 'Rates' }),
  ]) {
    const bad = structuredClone(plan);
    mutate(bad);
    await native({
      action: 'workbook',
      operation: 'artifact-write',
      root: f.root(),
      name: 'package-metadata.json',
      data: JSON.stringify(bad),
    });
    await assert.rejects(
      native({
        action: 'workbook',
        operation: 'package-edit',
        root: f.root(),
        version,
        source: path.basename(source.path),
      })
    );
  }
  assert.equal((await f.book.read('0', 'Sheet1', 'A1')).cells.A1.value, 'Value');
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
});

test('Document column rename has one Undo/Redo, captures queued input, durably recovers and cold-reopens saved facts', async (t) => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-column-document-')),
    source = path.join(root, 'source.xlsx'),
    recovery = path.join(root, 'recovery'),
    bytes = fixture();
  await fs.writeFile(source, bytes);
  const owners = [];
  t.after(async () => {
    for (const o of owners) await o.close();
    await fs.rm(root, { recursive: true, force: true });
  });
  const open = async (file = source, dir = recovery) => {
    const o = new SpreadsheetDocument({ recoveryDirectory: dir });
    owners.push(o);
    await o.open(file, { fresh: true });
    return o;
  };
  const o = await open(),
    request = (owner, operation, more = {}) =>
      owner.request({ documentId: owner.status().documentId, state: owner.status().state, operation, ...more }),
    read = async (owner) => (await request(owner, 'read', { sheet: 'Sheet1', range: 'A1:G4' })).cells;
  const start = o.status().state,
    command = renameColumn(),
    pending = request(o, 'apply', { commands: [command] });
  command.newName = 'Caller mutation';
  await pending;
  const accepted = o.status().state;
  assert.notEqual(accepted, start);
  assert.equal((await read(o)).A1.value, 'Unit Price');
  assert.equal((await read(o)).D1.formula, 'SUM(Prices[Unit Price])');
  assert.equal((await read(o)).D1.calculation.value, 9);
  await request(o, 'undo');
  assert.equal(o.status().state, start);
  assert.equal((await read(o)).A1.value, 'Value');
  await request(o, 'redo');
  assert.equal(o.status().state, accepted);
  assert.equal((await read(o)).A1.value, 'Unit Price');
  const manifest = JSON.parse(
    await fs.readFile(path.join(recovery, 'draft-' + o.status().documentId, 'current.json'), 'utf8')
  );
  assert.equal(manifest.history.frames.length, 2);
  const id = o.status().documentId;
  await o.close();
  const recovered = new SpreadsheetDocument({ recoveryDirectory: recovery });
  owners.push(recovered);
  await recovered.recover(id);
  assert.equal((await read(recovered)).D1.formula, 'SUM(Prices[Unit Price])');
  await request(recovered, 'undo');
  assert.equal((await read(recovered)).A1.value, 'Value');
  await request(recovered, 'redo');
  assert.equal((await read(recovered)).A1.value, 'Unit Price');
  const saved = path.join(root, 'saved.xlsx');
  await request(recovered, 'copy', { destination: saved, expectedRevision: null });
  await recovered.close();
  const cold = await open(saved, path.join(root, 'cold-recovery'));
  assert.equal(cold.status().canUndo, false);
  assert.equal(cold.status().dirty, false);
  assert.equal((await read(cold)).A1.value, 'Unit Price');
  assert.equal((await read(cold)).D1.calculation.value, 9);
  assert.deepEqual(await fs.readFile(source), bytes);
});

test('Manual column rename keeps pending calculation authority and saved policy explicit', async (t) => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-column-manual-')),
    source = path.join(root, 'source.xlsx'),
    bytes = fixture();
  await fs.writeFile(source, bytes);
  const owners = [];
  t.after(async () => {
    for (const o of owners) await o.close();
    await fs.rm(root, { recursive: true, force: true });
  });
  const open = async (file = source) => {
    const o = new SpreadsheetDocument({ recoveryDirectory: path.join(root, 'recovery-' + owners.length) });
    owners.push(o);
    await o.open(file, { fresh: true });
    return o;
  };
  const o = await open(),
    request = (owner, operation, more = {}) =>
      owner.request({ documentId: owner.status().documentId, state: owner.status().state, operation, ...more });
  await request(o, 'apply', {
    commands: [{ type: 'setCalculationOptions', mode: 'manual', recalculateBeforeSave: false }],
  });
  await request(o, 'apply', { commands: [{ sheet: 'Sheet1', cell: 'A2', value: 20 }] });
  const before = (await request(o, 'read', { sheet: 'Sheet1', range: 'D1' })).cells.D1;
  assert(before.calculation.pending);
  await request(o, 'apply', { commands: [renameColumn()] });
  const after = (await request(o, 'read', { sheet: 'Sheet1', range: 'A1:D2' })).cells;
  assert.equal(after.A1.value, 'Unit Price');
  assert.equal(after.D1.formula, 'SUM(Prices[Unit Price])');
  assert(after.D1.calculation.pending);
  assert.notEqual(after.D1.calculation.status, 'calculated');
  assert.equal(o.status().view.calculationPolicy.mode, 'manual');
  assert.equal(o.status().view.calculationPolicy.recalculateBeforeSave, false);
  const saved = path.join(root, 'manual-saved.xlsx');
  await request(o, 'copy', { destination: saved, expectedRevision: null });
  await o.close();
  const fresh = await open(saved);
  assert.equal(fresh.status().view.calculationPolicy.mode, 'manual');
  assert((await request(fresh, 'read', { sheet: 'Sheet1', range: 'D1' })).cells.D1.calculation.pending);
  assert.deepEqual(await fs.readFile(source), bytes);
});
