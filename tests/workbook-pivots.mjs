import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { spawnSync } from 'node:child_process';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook, editWorkbook } from '../core/workbook-package.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { indexedPivotPlan } from '../core/indexed-pivots.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
const command = {
  operation: 'createPivot',
  sheet: 'Sheet1',
  name: 'SalesPivot',
  range: 'A1:C5',
  rowField: 0,
  valueField: 1,
  aggregate: 'sum',
  pivotSheet: 'Summary',
  expectedPivotRevision: null,
};
function fixture() {
  const data = [
    ['Category', 'Amount', 'Flag'],
    ['A', 2, true],
    ['B', 5, false],
    ['A', -1, true],
    ['A', null, false],
  ];
  const edits = data.flatMap((row, r) =>
    row.map((value, c) => ({ sheet: 'Sheet1', cell: String.fromCharCode(65 + c) + (r + 1), value }))
  );
  const parts = unzipSync(editWorkbook(createWorkbook(), edits));
  parts['customXml/opaque.bin'] = new Uint8Array([0, 255, 17]);
  return Buffer.from(zipSync(parts));
}
async function owner(t, bytes = fixture()) {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-pivot-')),
    source = path.join(dir, 'source.xlsx');
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
      () => assert.fail('Legacy writer forbidden'),
      (input) => calculateIndexedWorkbook(input, bridge)
    );
  const book = make();
  t.after(async () => {
    await book.close();
    await fs.rm(dir, { recursive: true, force: true });
  });
  await book.open(source);
  return { book, make, source, bytes, dir, root: () => root, version: () => version };
}
async function values(book, state, expected) {
  const cells = (await book.read(state, 'Summary', 'A1:B4')).cells;
  assert.deepEqual(
    [
      cells.A1.value,
      cells.B1.value,
      cells.A2.value,
      cells.B2.value,
      cells.A3.value,
      cells.B3.value,
      cells.A4.value,
      cells.B4.value,
    ],
    expected
  );
}
test('pivot-only managed and foreign packages fence structural edits without a chart masking the guard', async (t) => {
  const seed = await owner(t),
    created = await seed.book.apply('0', [command]),
    saved = await seed.book.export(created.view.storage.state),
    parts = unzipSync(await fs.readFile(saved.path));
  assert(!Object.keys(parts).some((name) => name.startsWith('xl/charts/') || name.startsWith('xl/drawings/')));
  const table = Object.keys(parts).find((name) => name.startsWith('xl/pivotTables/') && name.endsWith('.xml'));
  assert(table);
  const foreign = { ...parts };
  foreign[table] = strToU8(strFromU8(parts[table]).replace('createdVersion="6"', 'createdVersion="5"'));
  assert.notDeepEqual(foreign[table], parts[table]);
  for (const [kind, bytes] of [
    ['managed', Buffer.from(zipSync(parts))],
    ['foreign', Buffer.from(zipSync(foreign))],
  ]) {
    const f = await owner(t, bytes);
    assert.equal(f.book.view('0').pivots.length, 1);
    assert.equal(f.book.view('0').pivots[0].canEdit, kind === 'managed');
    const original = await fs.readFile((await f.book.export('0')).path);
    for (const sheet of ['Sheet1', 'Summary'])
      for (const command of [
        { operation: 'insertRow', sheet, index: 0 },
        { operation: 'deleteColumn', sheet, index: 0 },
        ...(sheet === 'Summary' ? [{ operation: 'duplicateSheet', sheet, newName: 'Copy' }] : []),
      ]) {
        await assert.rejects(
          f.book.apply('0', [{ sheet: 'Sheet1', cell: 'B2', value: 99 }, command]),
          /pivot|linked|qualified|protected/i,
          kind + ' ' + sheet + ' ' + command.operation
        );
        assert.equal((await f.book.read('0', 'Sheet1', 'B2')).cells.B2.value, 2);
        await values(f.book, '0', ['Category', 'Sum of Amount', 'A', 1, 'B', 5, 'Grand Total', 6]);
        assert.equal(f.book.diagnosticStats().retainedStates, 1);
        assert.deepEqual(await fs.readFile((await f.book.export('0')).path), original);
        assert.deepEqual(await fs.readFile(f.source), bytes);
      }
    // Renaming the sheet a pivot table summarises is allowed: the pivot's
    // record of where its data is follows the new name, as in Excel.
    const renamed = await f.book.apply('0', [{ operation: 'renameSheet', sheet: 'Sheet1', newName: 'Renamed' }]),
      renamedParts = unzipSync(await fs.readFile((await f.book.export(renamed.view.storage.state)).path)),
      sources = Object.keys(renamedParts)
        .filter((name) => /pivotCache\/[^/]+\.xml$/.test(name))
        .flatMap((name) => strFromU8(renamedParts[name]).match(/<worksheetSource\b[^>]*>/g) ?? []);
    assert.equal(sources.length, 1);
    assert.match(sources[0], / sheet="Renamed"/);
    assert.equal(renamed.view.pivots.length, 1);
    assert.equal((await f.book.read('0', 'Sheet1', 'B2')).cells.B2.value, 2);
    const copied = await f.book.apply('0', [{ operation: 'duplicateSheet', sheet: 'Sheet1', newName: 'Copy' }]),
      copyState = copied.view.storage.state,
      output = await f.book.export(copyState),
      copyParts = unzipSync(await fs.readFile(output.path));
    assert.equal((await f.book.read(copyState, 'Copy', 'B2')).cells.B2.value, 2);
    assert.equal((await f.book.read(copyState, 'Sheet1', 'B2')).cells.B2.value, 2);
    await values(f.book, copyState, ['Category', 'Sum of Amount', 'A', 1, 'B', 5, 'Grand Total', 6]);
    for (const name of Object.keys(parts).filter((name) => /xl\/pivot(?:Tables|Cache)\//.test(name)))
      assert.deepEqual(copyParts[name], unzipSync(bytes)[name]);
    const cold = f.make();
    t.after(() => cold.close());
    await cold.open(output.path);
    assert.equal(cold.view('0').pivots[0].canEdit, kind === 'managed');
    await values(cold, '0', ['Category', 'Sum of Amount', 'A', 1, 'B', 5, 'Grand Total', 6]);
  }
});
test('real SUM pivot package, compact replay, protected output and independent saved cache/display', async (t) => {
  const f = await owner(t),
    created = await f.book.apply('0', [command]),
    state = created.view.storage.state;
  assert.equal(created.view.pivots.length, 1);
  assert(created.view.pivots[0].canEdit);
  assert.deepEqual(created.commands, [command]);
  assert(JSON.stringify(created.commands).length < 512);
  await values(f.book, state, ['Category', 'Sum of Amount', 'A', 1, 'B', 5, 'Grand Total', 6]);
  assert.equal(f.book.view('0').sheets.length, 1);
  await assert.rejects(f.book.apply(state, [{ sheet: 'Summary', cell: 'B2', value: 99 }]), /Pivot output|pivot table/);
  const output = await f.book.export(state),
    parts = unzipSync(await fs.readFile(output.path));
  assert.deepEqual(parts['customXml/opaque.bin'], unzipSync(f.bytes)['customXml/opaque.bin']);
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
  const probe = spawnSync(
    process.env.XLSX_EDITOR_PYTHON ?? 'python3',
    [
      '-c',
      `import sys,zipfile,xml.etree.ElementTree as E,json
from openpyxl import load_workbook
book=load_workbook(sys.argv[1]);p=book['Summary']._pivots[0]
assert p.name=='SalesPivot' and p.cacheId==1 and p.cache.recordCount==4
assert p.cache.cacheSource.worksheetSource.ref=='A1:C5'
assert len(p.cache.records.r)==4 and len(p.cache.cacheFields)==3
assert list(book['Summary'].values)==[('Category','Sum of Amount'),('A',1),('B',5),('Grand Total',6)]
with zipfile.ZipFile(sys.argv[1]) as z:
 n='{http://schemas.openxmlformats.org/package/2006/relationships}';links=E.fromstring(z.read('xl/_rels/workbook.xml.rels'));assert len([r for r in links if r.attrib['Type'].endswith('/pivotCacheDefinition')])==1
print(json.dumps({'sheets':book.sheetnames,'pivot':p.name,'records':p.cache.recordCount,'values':list(book['Summary'].values)}))`,
      output.path,
    ],
    { encoding: 'utf8' }
  );
  assert.equal(probe.status, 0, probe.stderr);
  console.log(probe.stdout.trim());
  const replay = f.make();
  t.after(() => replay.close());
  await replay.open(f.source);
  const again = await replay.apply('0', created.commands);
  await values(replay, again.view.storage.state, ['Category', 'Sum of Amount', 'A', 1, 'B', 5, 'Grand Total', 6]);
  const cold = f.make();
  t.after(() => cold.close());
  await cold.open(output.path);
  assert(cold.view('0').pivots[0].canEdit);
  await values(cold, '0', ['Category', 'Sum of Amount', 'A', 1, 'B', 5, 'Grand Total', 6]);
});
test('explicit refresh changes records/aggregate, keeps old snapshot and fences stale cache', async (t) => {
  const f = await owner(t),
    created = await f.book.apply('0', [command]),
    first = created.view.storage.state,
    pivot = created.view.pivots[0];
  const edited = await f.book.apply(first, [{ sheet: 'Sheet1', cell: 'B2', value: 7 }]),
    second = edited.view.storage.state;
  await values(f.book, second, ['Category', 'Sum of Amount', 'A', 1, 'B', 5, 'Grand Total', 6]);
  const refresh = {
    operation: 'refreshPivot',
    sheet: 'Summary',
    name: 'SalesPivot',
    expectedPivotRevision: pivot.revision,
  };
  const result = await f.book.apply(second, [refresh]),
    third = result.view.storage.state;
  await values(f.book, third, ['Category', 'Sum of Amount', 'A', 6, 'B', 5, 'Grand Total', 11]);
  await values(f.book, first, ['Category', 'Sum of Amount', 'A', 1, 'B', 5, 'Grand Total', 6]);
  assert.deepEqual(result.commands, [refresh]);
  assert.notEqual(result.view.pivots[0].revision, pivot.revision);
  await assert.rejects(f.book.apply(third, [refresh]), /Conflict/);
  const saved = await f.book.export(third),
    cold = f.make();
  t.after(() => cold.close());
  await cold.open(saved.path);
  assert(cold.view('0').pivots[0].canEdit);
  await values(cold, '0', ['Category', 'Sum of Amount', 'A', 6, 'B', 5, 'Grand Total', 11]);
});
test('numeric COUNT, protected literals and forged native cache/display/graph refuse', async (t) => {
  const f = await owner(t),
    created = await f.book.apply('0', [{ ...command, aggregate: 'count' }]);
  await values(f.book, created.view.storage.state, ['Category', 'Count of Amount', 'A', 2, 'B', 1, 'Grand Total', 3]);
  await f.book.read('0', 'Sheet1', 'A1');
  const root = f.root(),
    version = f.version(),
    metadata = await fs.readFile(path.join(root, 'metadata.xlsx')),
    prepared = await indexedPivotPlan({
      metadataBytes: metadata,
      view: f.book.view('0'),
      command,
      readRange: (sheet, range) => f.book.read('0', sheet, range),
    }),
    source = await f.book.export('0');
  for (const mutate of [
    (p) =>
      (p.replacements[p.parts.table] = p.replacements[p.parts.table].replace(
        '<colItems count="1"><i/></colItems>',
        ''
      )),
    (p) =>
      (p.replacements[p.parts.table] = p.replacements[p.parts.table]
        .replace('<items count="2">', '<items count="3">')
        .replace('</items>', '<item t="default"/></items>')),
    (p) =>
      (p.replacements[p.parts.definition] = p.replacements[p.parts.definition].replace(
        'containsSemiMixedTypes="1"',
        'containsSemiMixedTypes="0"'
      )),
    (p) =>
      (p.replacements[p.targetPath] = p.replacements[p.targetPath].replace(
        '</worksheet>',
        '<pivotTableParts count="1"><pivotTablePart r:id="rIdPivot"/></pivotTableParts></worksheet>'
      )),
    (p) =>
      (p.replacements['xl/pivotCache/workbookCacheRecords1.xml'] = p.replacements[
        'xl/pivotCache/workbookCacheRecords1.xml'
      ].replace('<x v="0"/>', '<x v="1"/>')),
    (p) => (p.replacements[p.targetPath] = p.replacements[p.targetPath].replace('<v>1</v>', '<v>99</v>')),
    (p) =>
      (p.replacements['xl/_rels/workbook.xml.rels'] = p.replacements['xl/_rels/workbook.xml.rels'].replace(
        'pivotCacheDefinition',
        'worksheet'
      )),
    (p) =>
      (p.replacements['[Content_Types].xml'] = p.replacements['[Content_Types].xml'].replace(
        'pivotTable+xml',
        'worksheet+xml'
      )),
    (p) =>
      (p.replacements[p.targetPath] = p.replacements[p.targetPath].replace(
        'defaultRowHeight="15"',
        'defaultRowHeight="99"'
      )),
  ]) {
    const bad = structuredClone(prepared.plan);
    mutate(bad);
    await native({
      action: 'workbook',
      operation: 'artifact-write',
      root,
      name: 'package-metadata.json',
      data: JSON.stringify(bad),
    });
    await assert.rejects(
      native({ action: 'workbook', operation: 'package-edit', root, version, source: path.basename(source.path) })
    );
  }
  assert.equal(f.book.view('0').sheets.length, 1);
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
  const formulas = await owner(t, editWorkbook(f.bytes, [{ sheet: 'Sheet1', cell: 'B2', formula: '1+1' }]));
  const formulaCreated = await formulas.book.apply('0', [command]);
  await values(formulas.book, formulaCreated.view.storage.state, [
    'Category',
    'Sum of Amount',
    'A',
    1,
    'B',
    5,
    'Grand Total',
    6,
  ]);
  assert(formulaCreated.view.pivots[0].canEdit);
  const manual = await owner(t),
    policy = await manual.book.apply('0', [
      { type: 'setCalculationOptions', mode: 'manual', recalculateBeforeSave: false },
    ]),
    pending = await manual.book.apply(policy.view.storage.state, [{ sheet: 'Sheet1', cell: 'B2', formula: '1+1' }]);
  await assert.rejects(manual.book.apply(pending.view.storage.state, [command]), /ordinary current|current numeric/);
  const unsafe = await owner(t, editWorkbook(f.bytes, [{ sheet: 'Sheet1', cell: 'B2', value: 0.1 }]));
  await assert.rejects(unsafe.book.apply('0', [command]), /integers/);
  await assert.rejects(f.book.apply('0', [{ ...command, pivotSheet: 'Sheet1' }]), /distinct/);
});

test('legacy nonstandard pivot markup remains preserved and protected from refresh', async (t) => {
  const f = await owner(t),
    made = await f.book.apply('0', [command]),
    saved = await f.book.export(made.view.storage.state),
    parts = unzipSync(await fs.readFile(saved.path));
  const key = made.view.sheets.find((s) => s.name === 'Summary').path,
    xml = new TextDecoder().decode(parts[key]);
  parts[key] = strToU8(
    xml.replace(
      '</worksheet>',
      '<pivotTableParts count="1"><pivotTablePart r:id="rIdPivot"/></pivotTableParts></worksheet>'
    )
  );
  const bytes = Buffer.from(zipSync(parts)),
    legacy = await owner(t, bytes),
    pivot = legacy.book.view('0').pivots[0];
  assert.equal(pivot.canEdit, false);
  await assert.rejects(
    legacy.book.apply('0', [
      { operation: 'refreshPivot', sheet: 'Summary', name: 'SalesPivot', expectedPivotRevision: pivot.revision },
    ]),
    /protected|Conflict/
  );
  const output = await legacy.book.export('0');
  assert.deepEqual(unzipSync(await fs.readFile(output.path))[key], parts[key]);
  assert.deepEqual(await fs.readFile(legacy.source), bytes);
});
