/** Actual native/Document content checks. No renderer control completion claim. */
import childProcess from 'node:child_process';
import { syncBuiltinESMExports } from 'node:module';
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createHash } from 'node:crypto';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { closeNativeWorkbookWorker, nativeWorkbookBinary } from '../core/native-runtime.mjs';
import { accessibilityFixture } from './support/accessibility-fixture.mjs';
import { validateAccessibility } from '../core/accessibility.mjs';
import { assertDesktopLaunchAllowed } from './support/xlsx-desktop-launch-guard.mjs';
assertDesktopLaunchAllowed();
const roots = [],
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
      new Promise((_, reject) => (timer = setTimeout(() => reject(Error('Native helper did not retire')), 5000))),
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
  receipt = (o) => {
    const s = o.status();
    return { documentId: s.documentId, ownerId: s.view.storage.sessionId, state: s.state, revision: s.view.revision };
  };
async function fixture(t, options = {}) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'accessibility-owner-'));
  roots.push(root);
  const { bytes } = accessibilityFixture(options),
    filename = path.join(root, 'source.xlsx'),
    owners = [];
  await fs.writeFile(filename, bytes);
  const open = async (file = filename) => {
    const o = new SpreadsheetDocument({ recoveryDirectory: path.join(root, 'recovery'), cacheIdentity: null });
    owners.push(o);
    await o.open(file, { fresh: true });
    return o;
  };
  const o = await open();
  t.after(async () => {
    for (const o of owners) await o.close();
  });
  return { root, filename, bytes, o, open };
}
const inspect = async (o) => {
  const v = (await req(o, 'accessibility')).accessibility;
  return validateAccessibility(v, receipt(o));
};
test('real native headers on two worksheets yield exact literal findings with zero accepted-head, history, recovery or source mutation', async (t) => {
  const f = await fixture(t),
    o = f.o,
    before = structuredClone(o.status()),
    journal = path.join(f.root, 'recovery', 'draft-' + before.documentId, 'current.json'),
    journalBytes = await fs.readFile(journal),
    cells = await req(o, 'read', { sheet: 'Sheet1', range: 'A1:C3' }),
    v = await inspect(o);
  assert.equal(v.status, 'complete');
  assert.deepEqual(v.coverage.tables, { total: 2, checked: 2, skipped: 0 });
  assert.deepEqual(
    v.findings.map((f) => [f.rule, f.sheet, f.cell, f.relatedCells]),
    [
      ['blank-table-header', 'Sheet1', 'C1', []],
      ['duplicate-table-header', 'Sheet1', 'A1', ['A1', 'B1']],
    ]
  );
  for (let i = 0; i < 3; i++) assert.deepEqual(await inspect(o), v);
  assert.deepEqual(o.status(), before);
  assert.deepEqual(await fs.readFile(journal), journalBytes);
  assert.deepEqual(await req(o, 'read', { sheet: 'Sheet1', range: 'A1:C3' }), cells);
  assert.equal(hash(await fs.readFile(f.filename)), hash(f.bytes));
});
test('real owned chart and shape alternative text findings disappear after a separate accepted object edit, while inspection preserves Redo', async (t) => {
  const f = await fixture(t),
    o = f.o,
    chart = '11111111-1111-4111-8111-111111111111',
    shape = '22222222-2222-4222-8222-222222222222';
  await req(o, 'apply', {
    commands: [
      {
        operation: 'createChart',
        sheet: 'Sheet1',
        objectId: chart,
        expectedObjectRevision: null,
        chart: { kind: 'column', range: 'A1:B3' },
        anchor: { cell: 'E2' },
      },
      {
        operation: 'createShape',
        sheet: 'Details',
        objectId: shape,
        expectedObjectRevision: null,
        shape: { sample: 'icon-star', text: '', color: '2563EB' },
        anchor: { cell: 'Z80' },
        altText: '   ',
      },
    ],
  });
  let v = await inspect(o);
  assert.deepEqual(v.coverage.objects, { total: 2, checked: 2, skipped: 0 });
  assert.deepEqual(
    v.findings.filter((f) => f.rule === 'missing-alt-text').map((f) => [f.sheet, f.cell, f.objectId]),
    [
      ['Sheet1', 'E2', chart],
      ['Details', 'Z80', shape],
    ]
  );
  const object = await req(o, 'object', { sheet: 'Details', objectId: shape });
  await req(o, 'apply', {
    commands: [
      {
        operation: 'updateShape',
        sheet: 'Details',
        objectId: shape,
        expectedObjectRevision: object.revision,
        patch: {},
        altText: 'A star marks important information',
      },
    ],
  });
  assert.equal((await inspect(o)).findings.filter((f) => f.rule === 'missing-alt-text').length, 1);
  await req(o, 'undo');
  const before = structuredClone(o.status());
  assert.equal(before.canRedo, true);
  v = await inspect(o);
  assert.equal(v.findings.filter((f) => f.rule === 'missing-alt-text').length, 2);
  assert.deepEqual(o.status(), before);
  await req(o, 'redo');
  assert.equal((await inspect(o)).findings.filter((f) => f.rule === 'missing-alt-text').length, 1);
});
test('protected worksheet tables still inspected; extended and formula-header tables have explicit partial coverage without guessed clean results', async (t) => {
  for (const options of [{ protectedSheet: true }, { extended: true }, { formula: true }]) {
    const f = await fixture(t, options),
      before = structuredClone(f.o.status()),
      v = await inspect(f.o);
    assert.deepEqual(f.o.status(), before);
    if (options.protectedSheet) {
      assert(before.view.protectedSheets.includes('Sheet1'));
      await assert.rejects(req(f.o, 'apply', { commands: [{ sheet: 'Sheet1', cell: 'A1', value: 'Blocked' }] }));
      assert.deepEqual(f.o.status(), before);
      assert.equal(v.status, 'complete');
      assert.equal(v.findings.length, 2);
    } else {
      assert.equal(v.status, 'partial');
      assert.equal(v.findings.length, 0);
      assert.deepEqual(v.coverage.tables, { total: 2, checked: 1, skipped: 1 });
      assert.equal(v.coverage.skipped[0].name, 'Data1');
    }
    assert.equal(hash(await fs.readFile(f.filename)), hash(f.bytes));
    await f.o.close();
  }
});
test('Document read fields and original receipt reject stale owners/heads and caller mutation before queue admission', async (t) => {
  const f = await fixture(t),
    o = f.o,
    before = structuredClone(o.status()),
    captured = { operation: 'accessibility', documentId: before.documentId, state: before.state },
    run = o.request(captured);
  captured.documentId = 'caller mutation';
  captured.state = 'caller mutation';
  const v = (await run).accessibility;
  validateAccessibility(v, receipt(o));
  for (const patch of [{ documentId: 'other' }, { state: 'other' }, { sheet: 'Sheet1' }, { repair: true }])
    await assert.rejects(o.request({ ...captured, documentId: before.documentId, state: before.state, ...patch }));
  await req(o, 'apply', { commands: [{ sheet: 'Details', cell: 'D5', value: 1 }] });
  await assert.rejects(o.request({ operation: 'accessibility', documentId: before.documentId, state: before.state }));
  assert.throws(() => validateAccessibility(v, receipt(o)), /retired/);
});

test('imported unmanaged drawings are explicitly skipped and readable protected owned objects still receive content findings', async (t) => {
  const f = await fixture(t),
    o = f.o,
    id = '33333333-3333-4333-8333-333333333333';
  await req(o, 'apply', {
    commands: [
      {
        operation: 'createShape',
        sheet: 'Sheet1',
        objectId: id,
        expectedObjectRevision: null,
        shape: { sample: 'icon-star', text: '' },
        anchor: { cell: 'E9' },
        altText: '',
      },
    ],
  });
  const file = path.join(f.root, 'owned.xlsx');
  await req(o, 'copy', { destination: file, expectedRevision: null });
  await o.close();
  const original = unzipSync(await fs.readFile(file)),
    protectedParts = { ...original };
  protectedParts['xl/worksheets/sheet1.xml'] = strToU8(
    strFromU8(original['xl/worksheets/sheet1.xml']).replace(
      '<drawing',
      '<sheetProtection sheet="1" objects="1"/><drawing'
    )
  );
  const protectedFile = path.join(f.root, 'protected-object.xlsx');
  await fs.writeFile(protectedFile, zipSync(protectedParts));
  const ro = await f.open(protectedFile),
    before = structuredClone(ro.status()),
    v = await inspect(ro);
  assert(before.view.protectedSheets.includes('Sheet1'));
  await assert.rejects(req(ro, 'apply', { commands: [{ sheet: 'Sheet1', cell: 'A1', value: 'Blocked' }] }));
  assert.deepEqual(ro.status(), before);
  assert.equal(v.coverage.objects.checked, 1);
  assert.equal(v.findings.filter((f) => f.rule === 'missing-alt-text').length, 1);
  assert.deepEqual(ro.status(), before);
  await ro.close();
  const drawing = Object.keys(original).find(
    (p) => p.startsWith('xl/drawings/') && p.endsWith('.xml') && !p.includes('_rels')
  );
  original[drawing] = strToU8(strFromU8(original[drawing]).replace('workbook:' + id, 'Imported star'));
  const foreign = path.join(f.root, 'foreign-object.xlsx');
  await fs.writeFile(foreign, zipSync(original));
  const imported = await f.open(foreign),
    state = structuredClone(imported.status()),
    partial = await inspect(imported);
  assert.equal(partial.status, 'partial');
  assert.deepEqual(partial.coverage.objects, { total: 1, checked: 0, skipped: 1 });
  assert.match(partial.coverage.skipped[0].reason, /Unmanaged/);
  assert.equal(partial.findings.filter((f) => f.rule === 'missing-alt-text').length, 0);
  assert.equal(partial.findings.length, 2);
  assert.deepEqual(imported.status(), state);
});
