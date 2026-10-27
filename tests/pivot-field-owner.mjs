/** Actual owned-pivot field/cache/output authoring; UI and copied delivery separate. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { spawnSync } from 'node:child_process';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { pivotFieldFixture, initialPivotCommand } from './support/pivot-field-fixture.mjs';
after(() => closeNativeWorkbookWorker());
async function fixture(t) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'pivot-fields-')),
    source = path.join(root, 'source.xlsx'),
    bytes = pivotFieldFixture(),
    owners = [],
    calls = [];
  await fs.writeFile(source, bytes);
  let fault = null;
  const bridge = async (input) => {
    calls.push(input);
    if (fault && input.operation === 'artifact-write' && input.name === 'package-metadata.json') {
      const p = JSON.parse(input.data);
      fault(p);
      input = { ...input, data: JSON.stringify(p) };
    }
    return native(input);
  };
  const open = async (file = source, dir = path.join(root, 'recovery')) => {
    const o = new SpreadsheetDocument({ native: bridge, recoveryDirectory: dir, cacheIdentity: null });
    owners.push(o);
    await o.open(file, { fresh: true });
    return o;
  };
  t.after(async () => {
    for (const o of owners) await o.close();
    await fs.rm(root, { recursive: true, force: true });
  });
  const o = await open();
  await req(o, 'apply', { commands: [initialPivotCommand] });
  return { root, source, bytes, o, open, owners, fault: (f) => (fault = f), calls };
}
const req = (o, operation, more = {}) =>
  o.request({ operation, documentId: o.status().documentId, state: o.status().state, ...more });
const definition = (o) => o.status().view.pivots[0];
const command = (o, patch = {}) => ({
  operation: 'updatePivot',
  sheet: 'Summary',
  name: 'SalesPivot',
  sourceSheet: 'Sheet1',
  range: 'A1:D6',
  rowField: 2,
  valueField: 3,
  aggregate: 'count',
  expectedPivotRevision: definition(o).revision,
  ...patch,
});
const read = async (o, sheet = 'Summary', range = 'A1:B8') => (await req(o, 'read', { sheet, range })).cells;
async function rows(o, expected) {
  const cells = await read(o);
  assert.deepEqual(
    expected.map((row, i) => ['A', 'B'].map((c) => cells[c + (i + 1)]?.value)),
    expected
  );
  return cells;
}
const countRows = [
  ['Region', 'Count of Units'],
  ['North', 2],
  ['South', 1],
  ['West', 0],
  ['East', 1],
  ['Grand Total', 4],
];
test('source/row/value/COUNT edit regenerates real parts/output in one Undo and saved cold facts', async (t) => {
  const f = await fixture(t),
    before = f.o.status().state,
    c = command(f.o);
  await req(f.o, 'apply', { commands: [c] });
  const accepted = f.o.status().state;
  await rows(f.o, countRows);
  assert.equal(definition(f.o).sourceRange, 'A1:D6');
  assert.equal(definition(f.o).rowField, 2);
  assert.equal(definition(f.o).valueField, 3);
  assert.equal(definition(f.o).aggregate, 'count');
  assert.notEqual(definition(f.o).revision, c.expectedPivotRevision);
  await req(f.o, 'undo');
  assert.equal(f.o.status().state, before);
  await rows(f.o, [
    ['Category', 'Sum of Amount'],
    ['A', 1],
    ['B', 5],
    ['C', 4],
    ['Grand Total', 10],
  ]);
  await req(f.o, 'redo');
  assert.equal(f.o.status().state, accepted);
  await rows(f.o, countRows);
  const saved = path.join(f.root, 'count.xlsx');
  await req(f.o, 'copy', { destination: saved, expectedRevision: null });
  const check = spawnSync(
    'python3',
    [
      '-c',
      `import sys,zipfile,xml.etree.ElementTree as E
n='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
with zipfile.ZipFile(sys.argv[1]) as z:
 t=E.fromstring(z.read('xl/pivotTables/workbookPivot1.xml'));d=E.fromstring(z.read('xl/pivotCache/workbookCacheDefinition1.xml'));r=E.fromstring(z.read('xl/pivotCache/workbookCacheRecords1.xml'));assert t.find(n+'rowFields/'+n+'field').get('x')=='2';v=t.find(n+'dataFields/'+n+'dataField');assert v.get('fld')=='3' and v.get('subtotal')=='count';assert d.find(n+'cacheSource/'+n+'worksheetSource').attrib=={'ref':'A1:D6','sheet':'Sheet1'};assert len(list(r))==5;assert z.read('custom/opaque.bin')==bytes([1,3,5,255])
`,
      saved,
    ],
    { encoding: 'utf8' }
  );
  assert.equal(check.status, 0, check.stderr);
  await f.o.close();
  const cold = await f.open(saved, path.join(f.root, 'cold'));
  assert.equal(cold.status().canUndo, false);
  await rows(cold, countRows);
  assert(definition(cold).canEdit);
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
});
test('SUM change, shrink and alternate worksheet/rectangle reuse the same owned graph', async (t) => {
  const f = await fixture(t);
  await req(f.o, 'apply', { commands: [command(f.o, { aggregate: 'sum' })] });
  await rows(f.o, [
    ['Region', 'Sum of Units'],
    ['North', 40],
    ['South', 20],
    ['West', 0],
    ['East', 40],
    ['Grand Total', 100],
  ]);
  await req(f.o, 'apply', { commands: [command(f.o, { range: 'A1:D4', aggregate: 'sum' })] });
  const shrunk = await rows(f.o, [
    ['Region', 'Sum of Units'],
    ['North', 40],
    ['South', 20],
    ['Grand Total', 60],
  ]);
  assert.equal(shrunk.A5, undefined);
  assert.equal(shrunk.A6, undefined);
  await req(f.o, 'apply', {
    commands: [
      command(f.o, { sourceSheet: 'Alternate', range: 'B2:E6', rowField: 0, valueField: 1, aggregate: 'sum' }),
    ],
  });
  await rows(f.o, [
    ['Category', 'Sum of Amount'],
    ['Z', 10],
    ['Y', 5],
    ['Grand Total', 15],
  ]);
  assert.equal(definition(f.o).sourceSheet, 'Alternate');
  assert.equal(definition(f.o).sourceRange, 'B2:E6');
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
});
test('captured update survives caller mutation and durable recovery with one frame per edit', async (t) => {
  const f = await fixture(t),
    c = command(f.o),
    pending = req(f.o, 'apply', { commands: [c] });
  c.rowField = 0;
  c.aggregate = 'sum';
  await pending;
  await rows(f.o, countRows);
  const id = f.o.status().documentId,
    recovery = path.join(f.root, 'recovery');
  await f.o.close();
  const recovered = new SpreadsheetDocument({ recoveryDirectory: recovery, cacheIdentity: null });
  f.owners.push(recovered);
  await recovered.recover(id);
  await rows(recovered, countRows);
  await req(recovered, 'undo');
  await rows(recovered, [
    ['Category', 'Sum of Amount'],
    ['A', 1],
    ['B', 5],
    ['C', 4],
    ['Grand Total', 10],
  ]);
  await req(recovered, 'redo');
  await rows(recovered, countRows);
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
});
test('stale/missing/colliding fields, wrong identity and forged refresh retargeting cannot publish a partial pivot', async (t) => {
  const f = await fixture(t),
    before = structuredClone(f.o.status());
  const missing = command(f.o);
  delete missing.aggregate;
  for (const c of [
    command(f.o, { expectedPivotRevision: '0'.repeat(64) }),
    command(f.o, { rowField: 3, valueField: 3 }),
    command(f.o, { sourceSheet: 'Summary' }),
    command(f.o, { rowField: 8 }),
    missing,
  ]) {
    await assert.rejects(req(f.o, 'apply', { commands: [c] }));
    assert.deepEqual(f.o.status(), before);
  }
  for (const fault of [
    (p) => (p.name = 'Wrong'),
    (p) => (p.priorRevision = '0'.repeat(64)),
    (p) =>
      (p.replacements['xl/pivotCache/workbookCacheRecords1.xml'] = p.replacements[
        'xl/pivotCache/workbookCacheRecords1.xml'
      ].replace('<x v="0"/>', '<x v="1"/>')),
  ]) {
    f.fault(fault);
    await assert.rejects(req(f.o, 'apply', { commands: [command(f.o)] }));
    assert.deepEqual(f.o.status(), before);
  }
  f.fault((p) => (p.kind = 'refreshPivot'));
  await assert.rejects(req(f.o, 'apply', { commands: [command(f.o)] }), { code: 'pivot-unsupported' });
  f.fault(null);
  await assert.rejects(req(f.o, 'apply', { commands: [command(f.o), { sheet: 'missing', cell: 'A1', value: 9 }] }));
  assert.deepEqual(f.o.status(), before);
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
});
test('output growth collision is explicit and permitted neighbors survive source shrink', async (t) => {
  const f = await fixture(t),
    saved = path.join(f.root, 'neighbor.xlsx');
  await req(f.o, 'copy', { destination: saved, expectedRevision: null });
  await f.o.close();
  const parts = unzipSync(await fs.readFile(saved));
  parts['xl/worksheets/sheet3.xml'] = strToU8(
    strFromU8(parts['xl/worksheets/sheet3.xml'])
      .replace('dimension ref="A1:B5"', 'dimension ref="A1:B6"')
      .replace('</sheetData>', '<row r="6"><c r="A6" t="inlineStr"><is><t>Neighbor</t></is></c></row></sheetData>')
  );
  await fs.writeFile(saved, zipSync(parts));
  const o = await f.open(saved, path.join(f.root, 'neighbors')),
    before = structuredClone(o.status());
  assert(definition(o).canEdit);
  await assert.rejects(req(o, 'apply', { commands: [command(o)] }), /occupied|growth/);
  assert.deepEqual(o.status(), before);
  await req(o, 'apply', { commands: [command(o, { range: 'A1:D4', rowField: 0, valueField: 1, aggregate: 'sum' })] });
  assert.equal((await read(o)).A6.value, 'Neighbor');
});
test('pending required numeric formula refuses update without creating calculation authority', async (t) => {
  const f = await fixture(t);
  await req(f.o, 'apply', { commands: [{ sheet: 'Sheet1', cell: 'B3', formula: 'B2+1' }] });
  await req(f.o, 'apply', {
    commands: [{ type: 'setCalculationOptions', mode: 'manual', recalculateBeforeSave: false }],
  });
  await req(f.o, 'apply', { commands: [{ sheet: 'Sheet1', cell: 'B2', value: 20 }] });
  const before = structuredClone(f.o.status()),
    cells = await read(f.o, 'Sheet1', 'B2:B3');
  assert(cells.B3.calculation.pending);
  await assert.rejects(
    req(f.o, 'apply', { commands: [command(f.o, { valueField: 1, aggregate: 'sum' })] }),
    /current|pending|ordinary/
  );
  assert.deepEqual(f.o.status(), before);
  assert.deepEqual(await read(f.o, 'Sheet1', 'B2:B3'), cells);
});
