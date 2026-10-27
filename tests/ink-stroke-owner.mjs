/** Actual native/indexed owner storage, guarded movement/removal, history/recovery and saved models. */
import childProcess from 'node:child_process';
import { syncBuiltinESMExports } from 'node:module';
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import {
  nativeWorkbookOperationAsync as native,
  closeNativeWorkbookWorker,
  nativeWorkbookBinary,
} from '../core/native-runtime.mjs';
import { pageThemePrintFixture } from './support/page-theme-print-fixture.mjs';
import { inkFoundationCases } from './support/ink-foundation-fixture.mjs';
import { assertDesktopLaunchAllowed } from './support/xlsx-desktop-launch-guard.mjs';
const foreignFixtures = assertDesktopLaunchAllowed(),
  roots = [],
  helper = nativeWorkbookBinary(),
  hash = (b) => createHash('sha256').update(b).digest('hex'),
  helperHash = hash(await fs.readFile(helper));
const children = [],
  spawn = childProcess.spawn;
childProcess.spawn = (file, ...args) => {
  const c = spawn(file, ...args);
  if (file === helper)
    children.push({
      pid: c.pid,
      exit: new Promise((resolve) => c.once('exit', (code, signal) => resolve({ code, signal }))),
    });
  return c;
};
syncBuiltinESMExports();
after(async () => {
  await closeNativeWorkbookWorker();
  childProcess.spawn = spawn;
  syncBuiltinESMExports();
  const lifecycle = [];
  for (const c of children) {
    let timer;
    const end = await Promise.race([
      c.exit,
      new Promise(
        (_, reject) => (timer = setTimeout(() => reject(Error('Native stroke helper did not retire')), 5000))
      ),
    ]).finally(() => clearTimeout(timer));
    assert.throws(
      () => process.kill(c.pid, 0),
      (e) => e.code === 'ESRCH'
    );
    lifecycle.push({ pid: c.pid, ...end, confirmedAbsent: true });
  }
  assert.equal(children.length, 1);
  for (const root of roots) await fs.rm(root, { recursive: true, force: true });
});
const req = (o, operation, more = {}) =>
    o.request({ operation, documentId: o.status().documentId, state: o.status().state, ...more }),
  apply = (o, c) => req(o, 'apply', { commands: Array.isArray(c) ? c : [c] }),
  cases = inkFoundationCases(),
  command = (c) => ({
    operation: 'createInkStroke',
    sheet: 'Sheet1',
    objectId: c.id,
    expectedObjectRevision: null,
    stroke: structuredClone(c.stroke),
    anchor: { ...c.anchor },
    altText: c.label,
  }),
  preview = (o, id) => req(o, 'object', { sheet: 'Sheet1', objectId: id });
async function fixture(t) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'ink-owner-'));
  roots.push(root);
  const bytes = pageThemePrintFixture({ theme: null }),
    filename = path.join(root, 'source.xlsx'),
    owners = [];
  await fs.writeFile(filename, bytes);
  let fault = null;
  const bridge = async (input) => {
    if (fault && input.operation === 'artifact-write' && input.name === 'package-metadata.json') {
      const plan = JSON.parse(input.data);
      fault(plan);
      input = { ...input, data: JSON.stringify(plan) };
    }
    return native(input);
  };
  const open = async (file = filename, dir = path.join(root, 'recovery')) => {
    const o = new SpreadsheetDocument({ native: bridge, recoveryDirectory: dir, cacheIdentity: null });
    owners.push(o);
    await o.open(file, { fresh: true });
    return o;
  };
  const recover = async (id) => {
    const o = new SpreadsheetDocument({
      native: bridge,
      recoveryDirectory: path.join(root, 'recovery'),
      cacheIdentity: null,
    });
    owners.push(o);
    await o.recover(id);
    return o;
  };
  const o = await open();
  t.after(async () => {
    for (const o of owners) await o.close();
  });
  return { root, filename, bytes, o, open, recover, fault: (v) => (fault = v) };
}
async function facts(o) {
  const r = (await req(o, 'read', { sheet: 'Sheet1', range: 'A1:F1' })).cells;
  for (const v of Object.values(r)) {
    delete v.revision;
    if (v.calculation) delete v.calculation.revision;
    if (v.cachedResult) delete v.cachedResult.revision;
    if (v.formula) {
      if (v.calculation?.status === 'calculated' && !v.calculation.pending) v.value = v.calculation.value;
      delete v.rawValue;
      delete v.cachedValue;
    }
  }
  return r;
}
async function save(f, o, name, strokes, more = {}) {
  const filename = path.join(f.root, name + '.xlsx');
  await req(o, 'copy', { destination: filename, expectedRevision: null });
  const expected = path.join(f.root, name + '-expected.json');
  await fs.writeFile(
    expected,
    JSON.stringify(
      {
        strokes: strokes.map((p) => ({ id: p.id, stroke: p.stroke, anchor: p.anchor, altText: p.altText })),
        formulaCaches: { B1: 6 },
        ...more,
      },
      null,
      2
    )
  );
  const proof = JSON.parse(
    execFileSync(
      process.env.XLSX_ORACLE_PYTHON ?? 'python3',
      [path.resolve('tests/support/check-ink-owner.py'), f.filename, filename, expected],
      { encoding: 'utf8' }
    )
  );
  assert(proof.passed);
  return { filename, proof };
}
test('three actual native vector styles attach to captured cells, preserve source/cell truth and survive independent Save/cold reopen', async (t) => {
  const f = await fixture(t),
    before = await facts(f.o),
    models = [];
  for (const c of cases) {
    const state = f.o.status().state;
    await apply(f.o, command(c));
    assert.notEqual(f.o.status().state, state);
    const p = await preview(f.o, c.id);
    assert.equal(p.kind, 'ink');
    assert.deepEqual(p.stroke, c.stroke);
    assert.deepEqual(p.anchor, c.anchor);
    assert.equal(p.altText, c.label);
    assert.equal(p.canEdit, true);
    assert.equal(p.canDelete, true);
    models.push(p);
    assert.deepEqual(await facts(f.o), before);
  }
  const saved = await save(f, f.o, 'three-strokes', models);
  assert.equal(hash(await fs.readFile(f.filename)), hash(f.bytes));
  await f.o.close();
  const cold = await f.open(saved.filename, path.join(f.root, 'cold'));
  assert.equal(cold.status().canUndo, false);
  for (const p of models) {
    const actual = await preview(cold, p.id);
    assert.deepEqual(actual.stroke, p.stroke);
    assert.deepEqual(actual.anchor, p.anchor);
  }
});
test('whole-stroke moves/alt/remove use guarded history and durable replay while siblings and original path/style remain intact', async (t) => {
  const f = await fixture(t);
  for (const c of cases.slice(0, 2)) await apply(f.o, command(c));
  const first = await preview(f.o, cases[0].id),
    sibling = await preview(f.o, cases[1].id),
    before = f.o.status().state;
  await apply(f.o, {
    operation: 'updateInkStroke',
    sheet: 'Sheet1',
    objectId: first.id,
    expectedObjectRevision: first.revision,
    anchor: { cell: 'Z80' },
    altText: 'Moved <&é stroke',
  });
  const accepted = f.o.status().state,
    p = await preview(f.o, first.id);
  assert.deepEqual(p.stroke, first.stroke);
  assert.equal(p.anchor.cell, 'Z80');
  await req(f.o, 'undo');
  assert.equal(f.o.status().state, before);
  assert.deepEqual(await preview(f.o, first.id), first);
  await req(f.o, 'redo');
  assert.equal(f.o.status().state, accepted);
  assert.deepEqual(await preview(f.o, sibling.id), sibling);
  const document = f.o.status().documentId;
  await f.o.close();
  const recovered = await f.recover(document);
  assert.deepEqual(await preview(recovered, first.id), p);
  const saved = await save(f, recovered, 'moved-recovered', [p, sibling]);
  await apply(recovered, {
    operation: 'deleteInkStroke',
    sheet: 'Sheet1',
    objectId: first.id,
    expectedObjectRevision: p.revision,
  });
  assert.deepEqual(await preview(recovered, sibling.id), sibling);
  await req(recovered, 'undo');
  assert.deepEqual(await preview(recovered, first.id), p);
  await req(recovered, 'redo');
  await apply(recovered, {
    operation: 'deleteObject',
    sheet: 'Sheet1',
    objectId: sibling.id,
    expectedObjectRevision: sibling.revision,
  });
  assert.equal((await req(recovered, 'objects', { sheet: 'Sheet1' })).total, 0);
  const empty = await save(f, recovered, 'empty', [], { noDrawings: true });
});
test('captured models, stale revision, forbidden rescale/restyle, forged native plans and late mixed-batch failure cannot publish a wrong stroke', async (t) => {
  const f = await fixture(t),
    c = command(cases[0]),
    run = apply(f.o, c);
  c.stroke.points[0][0] = 123;
  await run;
  const p = await preview(f.o, cases[0].id);
  assert.deepEqual(p.stroke, cases[0].stroke);
  const before = structuredClone(f.o.status()),
    cells = await facts(f.o),
    move = {
      operation: 'updateInkStroke',
      sheet: 'Sheet1',
      objectId: p.id,
      expectedObjectRevision: p.revision,
      anchor: { cell: 'D7' },
    };
  for (const bad of [
    { ...move, expectedObjectRevision: '0'.repeat(64) },
    { ...move, stroke: cases[1].stroke },
    { ...move, patch: { pen: 'red' } },
    { ...move, anchor: { widthEmu: p.stroke.widthEmu + 1 } },
    { ...command(cases[1]), anchor: undefined },
  ])
    await assert.rejects(apply(f.o, bad));
  f.fault(
    (plan) =>
      (plan.replacements[plan.drawingPart] = plan.replacements[plan.drawingPart].replace(
        'x="200025" y="628650"',
        'x="200026" y="628650"'
      ))
  );
  await assert.rejects(apply(f.o, command(cases[1])), /captured command/);
  assert.deepEqual(f.o.status(), before);
  f.fault(null);
  for (const fault of [
    (plan) => (plan.priorObjectRevision = '0'.repeat(64)),
    (plan) =>
      (plan.replacements[plan.drawingPart] = plan.replacements[plan.drawingPart].replace('fill="none"', 'fill="norm"')),
    (plan) =>
      (plan.replacements[plan.drawingPart] = plan.replacements[plan.drawingPart].replace(
        'val="000000"',
        'val="D32F2F"'
      )),
    (plan) =>
      (plan.replacements[plan.drawingPart] = plan.replacements[plan.drawingPart].replace('y="57150"', 'y="57151"')),
    (plan) =>
      (plan.replacements[plan.drawingPart] = plan.replacements[plan.drawingPart].replace(
        '<xdr:sp>',
        '<xdr:sp macro="run">'
      )),
  ]) {
    f.fault(fault);
    await assert.rejects(apply(f.o, move));
    assert.deepEqual(f.o.status(), before);
  }
  f.fault(null);
  await assert.rejects(apply(f.o, [move, { sheet: 'Missing', cell: 'A1', value: 99 }]));
  assert.deepEqual(f.o.status(), before);
  assert.deepEqual(await facts(f.o), cells);
  assert.equal(hash(await fs.readFile(f.filename)), hash(f.bytes));
});
test('existing chart/shape neighbors remain intact during stroke create, move and deletion', async (t) => {
  const f = await fixture(t),
    chart = '44444444-4444-4444-8444-444444444444',
    shape = '55555555-5555-4555-8555-555555555555';
  await apply(f.o, [
    {
      operation: 'createChart',
      sheet: 'Sheet1',
      objectId: chart,
      expectedObjectRevision: null,
      chart: { kind: 'column', range: 'D1:F2' },
      anchor: { cell: 'G5' },
    },
    {
      operation: 'createShape',
      sheet: 'Sheet1',
      objectId: shape,
      expectedObjectRevision: null,
      shape: { sample: 'icon-star', text: '' },
      anchor: { cell: 'H10' },
    },
  ]);
  const chartBefore = await preview(f.o, chart),
    shapeBefore = await preview(f.o, shape);
  await apply(f.o, command(cases[0]));
  let p = await preview(f.o, cases[0].id);
  await apply(f.o, {
    operation: 'updateInkStroke',
    sheet: 'Sheet1',
    objectId: p.id,
    expectedObjectRevision: p.revision,
    anchor: { cell: 'C9' },
  });
  p = await preview(f.o, p.id);
  await apply(f.o, {
    operation: 'deleteInkStroke',
    sheet: 'Sheet1',
    objectId: p.id,
    expectedObjectRevision: p.revision,
  });
  assert.deepEqual(await preview(f.o, chart), chartBefore);
  assert.deepEqual(await preview(f.o, shape), shapeBefore);
  assert.equal((await req(f.o, 'objects', { sheet: 'Sheet1' })).total, 2);
});
test('Auto current and Manual pending formula authority survive purely graphical stroke writes and saved reopen', async (t) => {
  const f = await fixture(t),
    auto = await facts(f.o);
  assert.equal(auto.B1.calculation.value, 6);
  await apply(f.o, command(cases[0]));
  assert.deepEqual(await facts(f.o), auto);
  await apply(f.o, { type: 'setCalculationOptions', mode: 'manual', recalculateBeforeSave: false });
  await apply(f.o, { sheet: 'Sheet1', cell: 'D1', value: 9 });
  const pending = await facts(f.o);
  assert(pending.B1.calculation.pending);
  await apply(f.o, command(cases[1]));
  assert.deepEqual(await facts(f.o), pending);
  const saved = path.join(f.root, 'pending.xlsx');
  await req(f.o, 'copy', { destination: saved, expectedRevision: null });
  await f.o.close();
  const cold = await f.open(saved, path.join(f.root, 'pending-cold'));
  assert((await facts(cold)).B1.calculation.pending);
  assert.deepEqual((await preview(cold, cases[1].id)).stroke, cases[1].stroke);
});

test('read-only imported protection and frozen authoring refuse while owned stroke data remains readable', async (t) => {
  const f = await fixture(t);
  await apply(f.o, command(cases[0]));
  const saved = path.join(f.root, 'before-protection.xlsx');
  await req(f.o, 'copy', { destination: saved, expectedRevision: null });
  await f.o.close();
  const parts = unzipSync(await fs.readFile(saved));
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    strFromU8(parts['xl/worksheets/sheet1.xml']).replace('<drawing', '<sheetProtection sheet="1" objects="1"/><drawing')
  );
  const filename = path.join(f.root, 'protected.xlsx');
  await fs.writeFile(filename, zipSync(parts));
  const ro = await f.open(filename, path.join(f.root, 'ro'));
  assert.equal(ro.status().view.readOnly, true);
  const p = await preview(ro, cases[0].id),
    state = structuredClone(ro.status());
  assert.deepEqual(p.stroke, cases[0].stroke);
  await assert.rejects(
    apply(ro, { operation: 'deleteInkStroke', sheet: 'Sheet1', objectId: p.id, expectedObjectRevision: p.revision })
  );
  assert.deepEqual(ro.status(), state);
  await ro.close();
  const open = await f.open(f.filename, path.join(f.root, 'frozen'));
  await apply(open, { operation: 'freeze', sheet: 'Sheet1', rows: 1, columns: 1 });
  const frozen = structuredClone(open.status());
  await assert.rejects(apply(open, command(cases[0])), /frozen/);
  assert.deepEqual(open.status(), frozen);
});
test('unmanaged drawing and opaque InkML bytes are preserved beside new owned strokes without an imported edit grant', async (t) => {
  const f = await fixture(t);
  const { inkFoundationFixture } = await import('./support/ink-foundation-fixture.mjs');
  const made = await inkFoundationFixture(path.join(f.root, 'import-constructor')),
    parts = unzipSync(await fs.readFile(made.saved)),
    drawing = 'xl/drawings/workbook-ink-fixture.xml';
  parts[drawing] = strToU8(strFromU8(parts[drawing]).replace('workbook:' + cases[0].id, 'Imported ink trace'));
  parts['custom/imported.inkml'] = strToU8(
    '<ink xmlns="http://www.w3.org/2003/InkML"><trace>0 0,10 20,30 10</trace></ink>'
  );
  parts['[Content_Types].xml'] = strToU8(
    strFromU8(parts['[Content_Types].xml']).replace(
      '</Types>',
      '<Override PartName="/custom/imported.inkml" ContentType="application/inkml+xml"/></Types>'
    )
  );
  const file = path.join(f.root, 'imported.xlsx');
  await fs.writeFile(file, zipSync(parts));
  await f.o.close();
  const owner = await f.open(file, path.join(f.root, 'imported-recovery')),
    list = await req(owner, 'objects', { sheet: 'Sheet1' }),
    foreign = list.objects.find((o) => !o.managed);
  assert(foreign && !foreign.canEdit && !foreign.canDelete && foreign.protectedReason);
  const bytes = strFromU8(parts[drawing]),
    raw = bytes.slice(
      bytes.indexOf('<xdr:oneCellAnchor'),
      bytes.indexOf('</xdr:oneCellAnchor>') + '</xdr:oneCellAnchor>'.length
    );
  await apply(owner, command(cases[0]));
  const output = path.join(f.root, 'with-owned-stroke.xlsx');
  await req(owner, 'copy', { destination: output, expectedRevision: null });
  const saved = unzipSync(await fs.readFile(output));
  assert(strFromU8(saved[drawing]).includes(raw));
  assert.deepEqual(saved['custom/imported.inkml'], parts['custom/imported.inkml']);
  assert.deepEqual(saved['custom/ink-opaque.bin'], parts['custom/ink-opaque.bin']);
  assert.equal((await req(owner, 'objects', { sheet: 'Sheet1' })).total, 4);
});
test('32 managed drawing identities cap stroke growth atomically with unchanged accepted head and source', async (t) => {
  const f = await fixture(t);
  for (let i = 0; i < 32; i++) {
    const id = (i + 100).toString(16).padStart(8, '0') + '-1111-4111-8111-111111111111';
    await apply(f.o, { ...command(cases[0]), objectId: id });
  }
  const before = structuredClone(f.o.status());
  assert.equal((await req(f.o, 'objects', { sheet: 'Sheet1' })).total, 32);
  await assert.rejects(apply(f.o, command(cases[1])), /limit/);
  assert.deepEqual(f.o.status(), before);
  assert.equal(hash(await fs.readFile(f.filename)), hash(f.bytes));
});
