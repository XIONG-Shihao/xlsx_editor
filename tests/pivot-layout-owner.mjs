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
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'pivot-layout-')),
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

const layoutCommand = (o, patch = {}) => ({
  operation: 'setPivotLayout',
  sheet: 'Summary',
  name: 'SalesPivot',
  sample: 'compact-blue',
  grandTotals: true,
  expectedPivotRevision: definition(o).revision,
  ...patch,
});
const baseRows = [
  ['Category', 'Sum of Amount'],
  ['A', 1],
  ['B', 5],
  ['C', 4],
  ['Grand Total', 10],
];
const saved = async (o, file) => {
  await req(o, 'copy', { destination: file, expectedRevision: null });
  return unzipSync(await fs.readFile(file));
};
const cacheFiles = ['xl/pivotCache/workbookCacheDefinition1.xml', 'xl/pivotCache/workbookCacheRecords1.xml'];
test('three saved presentation samples with totals on/off have real flags/output and independent cold facts', async (t) => {
  const f = await fixture(t),
    original = path.join(f.root, 'original.xlsx'),
    initial = await saved(f.o, original);
  for (const sample of ['tabular-light', 'compact-blue', 'outline-dark'])
    for (const grandTotals of [false, true]) {
      const before = f.o.status().state;
      await req(f.o, 'apply', { commands: [layoutCommand(f.o, { sample, grandTotals })] });
      const after = f.o.status().state;
      const expected = baseRows.slice(0, grandTotals ? 5 : 4).map((row) => [...row]);
      if (sample === 'compact-blue') expected[0][0] = 'Row Labels';
      const c = await rows(f.o, expected);
      if (!grandTotals) assert.equal(c.A5, undefined);
      assert.equal(definition(f.o).sample, sample);
      assert.equal(definition(f.o).grandTotals, grandTotals);
      assert(definition(f.o).canEdit);
      await req(f.o, 'undo');
      assert.equal(f.o.status().state, before);
      await req(f.o, 'redo');
      assert.equal(f.o.status().state, after);
      const file = path.join(f.root, sample + '-' + grandTotals + '.xlsx'),
        parts = await saved(f.o, file);
      for (const part of cacheFiles) assert.deepEqual(parts[part], initial[part]);
      const check = spawnSync(
        process.env.XLSX_EDITOR_PYTHON ?? 'python3',
        [path.join(process.cwd(), 'tests/support/check-pivot-layout.py'), original, file, sample, String(grandTotals)],
        { encoding: 'utf8' }
      );
      assert.equal(check.status, 0, check.stderr);
      console.log(check.stdout.trim());
      await f.o.close();
      f.o = await f.open(file, path.join(f.root, 'cold-' + sample + '-' + grandTotals));
      assert.equal(f.o.status().canUndo, false);
      assert(definition(f.o).canEdit);
      await rows(f.o, expected);
    }
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
});
test('layout does not refresh a changed source or authorize pending formula results', async (t) => {
  const f = await fixture(t);
  await req(f.o, 'apply', { commands: [{ sheet: 'Sheet1', cell: 'B3', formula: 'B2+1' }] });
  await req(f.o, 'apply', {
    commands: [{ type: 'setCalculationOptions', mode: 'manual', recalculateBeforeSave: false }],
  });
  await req(f.o, 'apply', { commands: [{ sheet: 'Sheet1', cell: 'B2', value: 20 }] });
  const input = await read(f.o, 'Sheet1', 'B2:B3');
  assert(input.B3.calculation.pending);
  const prior = await saved(f.o, path.join(f.root, 'pending-before.xlsx'));
  await req(f.o, 'apply', { commands: [layoutCommand(f.o, { grandTotals: false })] });
  await rows(f.o, [
    ['Row Labels', 'Sum of Amount'],
    ['A', 1],
    ['B', 5],
    ['C', 4],
  ]);
  const after = await read(f.o, 'Sheet1', 'B2:B3');
  assert.equal(after.B2.value, 20);
  assert.equal(after.B3.formula, input.B3.formula);
  assert.equal(after.B3.calculation.pending, true);
  assert.notEqual(after.B3.calculation.status, 'calculated');
  const next = await saved(f.o, path.join(f.root, 'pending-after.xlsx'));
  assert.deepEqual(next['xl/worksheets/sheet1.xml'], prior['xl/worksheets/sheet1.xml']);
  for (const part of cacheFiles) assert.deepEqual(next[part], prior[part]);
});
test('refresh and field editing preserve presentation and totals while actual aggregates change', async (t) => {
  const f = await fixture(t);
  await req(f.o, 'apply', { commands: [layoutCommand(f.o, { grandTotals: false })] });
  await req(f.o, 'apply', { commands: [{ sheet: 'Sheet1', cell: 'B2', value: 20 }] });
  await req(f.o, 'apply', {
    commands: [
      {
        operation: 'refreshPivot',
        sheet: 'Summary',
        name: 'SalesPivot',
        expectedPivotRevision: definition(f.o).revision,
      },
    ],
  });
  await rows(f.o, [
    ['Row Labels', 'Sum of Amount'],
    ['A', 19],
    ['B', 5],
    ['C', 4],
  ]);
  await req(f.o, 'apply', { commands: [command(f.o)] });
  await rows(f.o, [
    ['Row Labels', 'Count of Units'],
    ['North', 2],
    ['South', 1],
    ['West', 0],
    ['East', 1],
  ]);
  assert.equal(definition(f.o).sample, 'compact-blue');
  assert.equal(definition(f.o).grandTotals, false);
});
test('closed commands and forged cache/flags/output/kind cannot publish partial layout state', async (t) => {
  const f = await fixture(t),
    before = structuredClone(f.o.status()),
    missing = layoutCommand(f.o);
  delete missing.grandTotals;
  for (const c of [
    layoutCommand(f.o, { sample: 'missing' }),
    layoutCommand(f.o, { sample: undefined }),
    layoutCommand(f.o, { grandTotals: 'false' }),
    layoutCommand(f.o, { expectedPivotRevision: '0'.repeat(64) }),
    missing,
    layoutCommand(f.o, { range: 'A1:D6' }),
  ]) {
    await assert.rejects(req(f.o, 'apply', { commands: [c] }));
    assert.deepEqual(f.o.status(), before);
  }
  const table = 'xl/pivotTables/workbookPivot1.xml',
    records = cacheFiles[1],
    target = 'xl/worksheets/sheet3.xml';
  for (const fault of [
    (p) => (p.replacements[records] = p.replacements[records].replace('<x v="0"/>', '<x v="1"/>')),
    (p) => (p.sourceRange = 'A1:D6'),
    (p) => (p.replacements[table] = p.replacements[table].replace('compact="1"', 'compact="0"')),
    (p) => (p.replacements[target] = p.replacements[target].replace('<v>10</v>', '<v>99</v>')),
    (p) => (p.kind = 'refreshPivot'),
  ]) {
    f.fault(fault);
    await assert.rejects(req(f.o, 'apply', { commands: [layoutCommand(f.o)] }));
    assert.deepEqual(f.o.status(), before);
  }
  f.fault(null);
  await assert.rejects(
    req(f.o, 'apply', { commands: [layoutCommand(f.o), { sheet: 'missing', cell: 'A1', value: 99 }] })
  );
  assert.deepEqual(f.o.status(), before);
});
test('captured layout/totals survives caller mutation and durable recovery without source changes', async (t) => {
  const f = await fixture(t),
    c = layoutCommand(f.o, { sample: 'outline-dark', grandTotals: false }),
    pending = req(f.o, 'apply', { commands: [c] });
  c.sample = 'tabular-light';
  c.grandTotals = true;
  await pending;
  const id = f.o.status().documentId;
  await f.o.close();
  const o = new SpreadsheetDocument({ recoveryDirectory: path.join(f.root, 'recovery'), cacheIdentity: null });
  f.owners.push(o);
  await o.recover(id);
  assert.equal(definition(o).sample, 'outline-dark');
  assert.equal(definition(o).grandTotals, false);
  await rows(o, baseRows.slice(0, 4));
  await req(o, 'undo');
  await rows(o, baseRows);
  await req(o, 'redo');
  await rows(o, baseRows.slice(0, 4));
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
});

test('re-enabling totals refuses occupied growth; a style change keeps outside cells', async (t) => {
  const f = await fixture(t);
  await req(f.o, 'apply', { commands: [layoutCommand(f.o, { grandTotals: false })] });
  const file = path.join(f.root, 'neighbors.xlsx'),
    parts = await saved(f.o, file);
  await f.o.close();
  const neighbor = '<row r="5"><c r="A5" t="inlineStr"><is><t>Neighbor</t></is></c></row>';
  let targetXml = strFromU8(parts['xl/worksheets/sheet3.xml']);
  targetXml = /<row r="5"(?:\/>|><\/row>)/.test(targetXml)
    ? targetXml.replace(/<row r="5"(?:\/>|><\/row>)/, neighbor)
    : targetXml.replace('</sheetData>', neighbor + '</sheetData>');
  targetXml = targetXml.replace('<dimension ref="A1:B4"/>', '<dimension ref="A1:B5"/>');
  parts['xl/worksheets/sheet3.xml'] = strToU8(targetXml);
  assert(strFromU8(parts['xl/worksheets/sheet3.xml']).includes('Neighbor'));
  await fs.writeFile(file, zipSync(parts));
  const o = await f.open(file, path.join(f.root, 'neighbors')),
    before = structuredClone(o.status());
  assert(definition(o).canEdit);
  await assert.rejects(req(o, 'apply', { commands: [layoutCommand(o, { grandTotals: true })] }), /occupied|growth/);
  assert.deepEqual(o.status(), before);
  await req(o, 'apply', { commands: [layoutCommand(o, { sample: 'outline-dark', grandTotals: false })] });
  assert.equal((await read(o)).A5.value, 'Neighbor');
});
