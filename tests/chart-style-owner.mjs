import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { unzipSync } from 'fflate';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
const id = '11111111-1111-4111-8111-111111111111',
  req = (o, operation, more = {}) =>
    o.request({ operation, documentId: o.status().documentId, state: o.status().state, ...more }),
  preview = (o) => req(o, 'object', { sheet: 'Sales', objectId: id }),
  create = (kind, style) => ({
    operation: 'createChart',
    sheet: 'Sales',
    objectId: id,
    expectedObjectRevision: null,
    chart: { kind, style, range: 'A1:C4', title: 'Styled sales' },
  });
async function fixture(t) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'chart-style-owner-')),
    source = path.join(root, 'source.xlsx'),
    bytes = await fs.readFile('tests/support/w7-reference-data/source.xlsx'),
    owners = [];
  await fs.writeFile(source, bytes);
  let fault = null;
  const bridge = async (input) => {
    if (fault && input.operation === 'artifact-write' && input.name === 'package-metadata.json') {
      const plan = JSON.parse(input.data);
      fault(plan);
      input = { ...input, data: JSON.stringify(plan) };
    }
    return native(input);
  };
  const open = async (file = source, recovery = path.join(root, 'recovery')) => {
    const o = new SpreadsheetDocument({ native: bridge, recoveryDirectory: recovery, cacheIdentity: null });
    owners.push(o);
    await o.open(file, { fresh: true });
    return o;
  };
  const o = await open();
  t.after(async () => {
    for (const owner of owners) await owner.close();
    await fs.rm(root, { recursive: true, force: true });
  });
  return { root, source, bytes, o, owners, open, fault: (f) => (fault = f) };
}
async function saved(f, rows, name) {
  const file = path.join(f.root, name + '.xlsx'),
    expected = path.join(f.root, name + '.json');
  await req(f.o, 'copy', { destination: file, expectedRevision: null });
  await fs.writeFile(expected, JSON.stringify(rows));
  const check = spawnSync(
    process.env.XLSX_ORACLE_PYTHON ?? 'python3',
    [path.resolve('tests/support/check-chart-styles.py'), f.source, file, expected],
    { encoding: 'utf8' }
  );
  assert.equal(check.status, 0, check.stderr);
  console.log(check.stdout.trim());
  return file;
}
test('three real styles across column/bar/line save series properties and reopen the exact editable model', async (t) => {
  for (const kind of ['column', 'bar', 'line'])
    for (const style of ['colorful', 'monochrome', 'outline']) {
      const f = await fixture(t);
      await req(f.o, 'apply', { commands: [create(kind, style)] });
      const p = await preview(f.o);
      assert(p.canEdit && p.canDelete);
      assert.equal(p.chart.kind, kind);
      assert.equal(p.chart.style, style);
      const file = await saved(
        f,
        [{ id, kind, style, values: p.data.series.map((s) => s.values) }],
        kind + '-' + style
      );
      await f.o.close();
      const cold = await f.open(file, path.join(f.root, 'cold'));
      assert.deepEqual((await preview(cold)).chart, p.chart);
      assert.equal(cold.status().canUndo, false);
      assert.deepEqual(await fs.readFile(f.source), f.bytes);
    }
});
test('style edits are captured, retain immutable old/redo models, recover and preserve formatting through cache refresh', async (t) => {
  const f = await fixture(t),
    c = create('column', 'colorful'),
    pending = req(f.o, 'apply', { commands: [c] });
  c.chart.style = 'outline';
  await pending;
  let p = await preview(f.o);
  assert.equal(p.chart.style, 'colorful');
  const initial = f.o.status().state,
    command = {
      operation: 'updateChart',
      sheet: 'Sales',
      objectId: id,
      expectedObjectRevision: p.revision,
      patch: { style: 'monochrome' },
    },
    edit = req(f.o, 'apply', { commands: [command] });
  command.patch.style = 'outline';
  await edit;
  assert.equal((await preview(f.o)).chart.style, 'monochrome');
  const edited = f.o.status().state;
  await req(f.o, 'undo');
  assert.equal(f.o.status().state, initial);
  assert.equal((await preview(f.o)).chart.style, 'colorful');
  await req(f.o, 'redo');
  assert.equal(f.o.status().state, edited);
  p = await preview(f.o);
  const doc = f.o.status().documentId;
  await f.o.close();
  const recovered = new SpreadsheetDocument({ recoveryDirectory: path.join(f.root, 'recovery'), cacheIdentity: null });
  f.owners.push(recovered);
  await recovered.recover(doc);
  f.o = recovered;
  assert.equal((await preview(f.o)).chart.style, 'monochrome');
  await req(f.o, 'apply', { commands: [{ sheet: 'Sales', cell: 'B2', value: 19 }] });
  p = await preview(f.o);
  assert.equal(p.chart.style, 'monochrome');
  assert.equal(p.data.series[0].values[0], 19);
  const file = await saved(
    f,
    [{ id, kind: 'column', style: 'monochrome', values: p.data.series.map((s) => s.values) }],
    'refreshed'
  );
  const cold = await f.open(file, path.join(f.root, 'cold'));
  assert.equal((await preview(cold)).chart.style, 'monochrome');
  assert.equal((await preview(cold)).data.series[0].values[0], 19);
});
test('forged colors/fills/lines/mixed presets fail native admission and invalid/undefined styles cannot publish', async (t) => {
  const f = await fixture(t);
  await req(f.o, 'apply', { commands: [create('bar', 'colorful')] });
  const p = await preview(f.o),
    state = structuredClone(f.o.status()),
    command = {
      operation: 'updateChart',
      sheet: 'Sales',
      objectId: id,
      expectedObjectRevision: p.revision,
      patch: { style: 'outline' },
    };
  const faults = [
    (plan) => (plan.replacements[plan.objectPart] = plan.replacements[plan.objectPart].replace('2563EB', 'FF0000')),
    (plan) => (plan.replacements[plan.objectPart] = plan.replacements[plan.objectPart].replace('19050', '19051')),
    (plan) =>
      (plan.replacements[plan.objectPart] = plan.replacements[plan.objectPart].replace(
        '<a:noFill/>',
        '<a:solidFill><a:srgbClr val="2563EB"/></a:solidFill>'
      )),
    (plan) =>
      (plan.replacements[plan.objectPart] = plan.replacements[plan.objectPart].replace('val="dash"', 'val="dot"')),
    (plan) =>
      (plan.replacements[plan.objectPart] = plan.replacements[plan.objectPart].replace(
        '<c:spPr>',
        '<c:spPr macro="run">'
      )),
  ];
  for (const fault of faults) {
    f.fault(fault);
    await assert.rejects(req(f.o, 'apply', { commands: [command] }));
    assert.deepEqual(f.o.status(), state);
  }
  f.fault(null);
  for (const style of ['unknown', null, undefined]) {
    await assert.rejects(req(f.o, 'apply', { commands: [{ ...command, patch: { style } }] }));
    assert.deepEqual(f.o.status(), state);
  }
  await assert.rejects(req(f.o, 'apply', { commands: [command, { sheet: 'missing', cell: 'A1', value: 99 }] }));
  assert.deepEqual(f.o.status(), state);
});
test('legacy unstyled charts remain unstyled, and normal source/type/title updates retain an explicit preset', async (t) => {
  const f = await fixture(t),
    c = create('column', undefined);
  delete c.chart.style;
  await req(f.o, 'apply', { commands: [c] });
  let p = await preview(f.o);
  assert.equal(p.chart.style, undefined);
  let file = path.join(f.root, 'legacy.xlsx');
  await req(f.o, 'copy', { destination: file, expectedRevision: null });
  const parts = unzipSync(await fs.readFile(file)),
    chart = Object.keys(parts).find((p) => p.startsWith('xl/charts/'));
  assert(!new TextDecoder().decode(parts[chart]).includes('<c:spPr>'));
  await req(f.o, 'apply', {
    commands: [
      {
        operation: 'updateChart',
        sheet: 'Sales',
        objectId: id,
        expectedObjectRevision: p.revision,
        patch: { style: 'outline' },
      },
    ],
  });
  p = await preview(f.o);
  await req(f.o, 'apply', {
    commands: [
      {
        operation: 'updateChart',
        sheet: 'Sales',
        objectId: id,
        expectedObjectRevision: p.revision,
        patch: { kind: 'line', title: 'Keep preset' },
      },
    ],
  });
  p = await preview(f.o);
  assert.equal(p.chart.style, 'outline');
  assert.equal(p.chart.kind, 'line');
  await saved(f, [{ id, kind: 'line', style: 'outline', values: p.data.series.map((s) => s.values) }], 'preserved');
});
