import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { unzipSync, zipSync, strToU8 } from 'fflate';
import { createWorkbook } from '../core/workbook-package.mjs';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { normalizeModel3D } from '../core/model3d-model.mjs';
import { captureModel3DCommands } from '../core/model3d-authoring.mjs';
import { smartartSampleModel } from '../core/smartart-model.mjs';
after(() => closeNativeWorkbookWorker());
const id = '11111111-1111-4111-8111-111111111111',
  other = '22222222-2222-4222-8222-222222222222',
  sha = (b) => createHash('sha256').update(b).digest('hex');
const req = (o, operation, more = {}) =>
    o.request({ operation, documentId: o.status().documentId, state: o.status().state, ...more }),
  preview = (o, objectId = id) => req(o, 'object', { sheet: 'Sheet1', objectId }),
  create = (sample = 'cube', more = {}) => ({
    operation: 'createModel3D',
    sheet: 'Sheet1',
    objectId: id,
    expectedObjectRevision: null,
    model3d: normalizeModel3D({ sample }),
    altText: 'Real mesh',
    anchor: { cell: 'E2' },
    ...more,
  }),
  update = (p, patch, more = {}) => ({
    operation: 'updateModel3D',
    sheet: 'Sheet1',
    objectId: p.id,
    expectedObjectRevision: p.revision,
    patch,
    ...more,
  });
async function fixture(t) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'model3d-owner-')),
    source = path.join(root, 'source.xlsx'),
    parts = unzipSync(createWorkbook()),
    owners = [];
  parts['custom/model3d-opaque.bin'] = Uint8Array.from([0, 3, 255, 6]);
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Keep cells</t></is></c><c r="B1"><v>42</v></c></row><row r="2"><c r="A2"><f>B1+1</f><v>43</v></c></row></sheetData></worksheet>'
  );
  const bytes = Buffer.from(zipSync(parts));
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
  const open = async (file = source, recovery = path.join(root, 'recovery'), options = {}) => {
    const o = new SpreadsheetDocument({ native: bridge, recoveryDirectory: recovery, cacheIdentity: null });
    owners.push(o);
    await o.open(file, { fresh: true, ...options });
    return o;
  };
  const o = await open();
  t.after(async () => {
    for (const owner of owners) await owner.close();
    await fs.rm(root, { recursive: true, force: true });
  });
  return { root, source, bytes, o, owners, open, fault: (f) => (fault = f) };
}
const expected = (p) => ({ id: p.id, model: p.model3d, anchor: p.anchor, alt: p.altText });
async function checkSaved(f, rows, name) {
  const file = path.join(f.root, name + '.xlsx'),
    json = path.join(f.root, name + '.json');
  await req(f.o, 'copy', { destination: file, expectedRevision: null });
  await fs.writeFile(json, JSON.stringify(rows));
  const check = spawnSync(
    process.env.XLSX_ORACLE_PYTHON ?? 'python3',
    [path.resolve('tests/support/check-model3d-owner.py'), f.source, file, json],
    { encoding: 'utf8' }
  );
  assert.equal(check.status, 0, check.stderr);
  console.log(check.stdout.trim());
  if (process.env.MODEL3D_OWNER_EVIDENCE) {
    await fs.mkdir(process.env.MODEL3D_OWNER_EVIDENCE, { recursive: true });
    for (const [src, dst] of [
      [f.source, name + '-original.xlsx'],
      [file, name + '.xlsx'],
      [json, name + '.json'],
    ])
      await fs.copyFile(src, path.join(process.env.MODEL3D_OWNER_EVIDENCE, dst));
  }
  return file;
}
const facts = (cells) =>
  Object.fromEntries(
    Object.entries(cells).map(([k, c]) => {
      const v = structuredClone(c);
      if (v.calculation) delete v.calculation.revision;
      if (v.cachedResult) {
        delete v.cachedResult.revision;
        if (v.dataType === 'number' && v.cachedResult.rawValue != null)
          v.cachedResult.rawValue = Number(v.cachedResult.rawValue);
      }
      if (v.dataType === 'number' && v.rawValue !== null) v.rawValue = Number(v.rawValue);
      return [k, v];
    })
  );
test('all three real GLB samples create, rotate/recolor/reposition and save/fresh reopen without changing cells', async (t) => {
  for (const sample of ['cube', 'pyramid', 'octahedron']) {
    const f = await fixture(t),
      cells = facts((await req(f.o, 'read', { sheet: 'Sheet1', range: 'A1:B2' })).cells);
    await req(f.o, 'apply', { commands: [create(sample)] });
    let p = await preview(f.o);
    assert.equal(p.kind, 'model3d');
    assert.equal(p.canEdit, true);
    assert.equal(p.canDelete, true);
    assert.equal(p.asset3d.model.sample, sample);
    assert(p.asset3d.mesh.triangles.length >= 6);
    await req(f.o, 'apply', {
      commands: [
        update(
          p,
          { color: 'CC3333', rotation: { x: -35, y: 70, z: 15 } },
          {
            anchor: { cell: 'B4', dxEmu: 100, dyEmu: 200, widthEmu: 3810000, heightEmu: 2286000 },
            altText: 'Edited <&é "mesh"',
          }
        ),
      ],
    });
    p = await preview(f.o);
    assert.deepEqual(
      p.asset3d.model,
      normalizeModel3D({ sample, color: 'CC3333', rotation: { x: -35, y: 70, z: 15 } })
    );
    assert.equal(p.anchor.cell, 'B4');
    assert.deepEqual(facts((await req(f.o, 'read', { sheet: 'Sheet1', range: 'A1:B2' })).cells), cells);
    const file = await checkSaved(f, [expected(p)], 'edited-' + sample);
    await f.o.close();
    const cold = await f.open(file, path.join(f.root, 'cold'));
    assert.deepEqual((await preview(cold)).asset3d, p.asset3d);
    assert.deepEqual((await preview(cold)).anchor, p.anchor);
    assert.equal((await preview(cold)).altText, p.altText);
    assert.equal(cold.status().canUndo, false);
    assert.deepEqual(await fs.readFile(f.source), f.bytes);
  }
});
test('captured create/update commands, history/recovery and sibling/final GLB deletion remain reversible', async (t) => {
  const f = await fixture(t),
    base = f.o.status().state,
    c = create(),
    pending = req(f.o, 'apply', { commands: [c] });
  c.model3d.rotation.y = 180;
  c.model3d.color = '000000';
  c.anchor.cell = 'A100';
  await pending;
  let p = await preview(f.o);
  assert.equal(p.model3d.rotation.y, -30);
  assert.equal(p.anchor.cell, 'E2');
  const accepted = f.o.status().state;
  await req(f.o, 'undo');
  assert.equal(f.o.status().state, base);
  assert.equal((await req(f.o, 'objects', { sheet: 'Sheet1' })).total, 0);
  await req(f.o, 'redo');
  assert.equal(f.o.status().state, accepted);
  await req(f.o, 'apply', { commands: [create('pyramid', { objectId: other })] });
  const sibling = await preview(f.o, other),
    cmd = update(p, { sample: 'octahedron', rotation: { x: 0, y: 90, z: 180 } }),
    edit = req(f.o, 'apply', { commands: [cmd] });
  cmd.patch.rotation.y = 0;
  await edit;
  p = await preview(f.o);
  assert.equal(p.model3d.rotation.y, 90);
  assert.deepEqual(await preview(f.o, other), sibling);
  const documentId = f.o.status().documentId;
  await f.o.close();
  const recovered = new SpreadsheetDocument({ recoveryDirectory: path.join(f.root, 'recovery'), cacheIdentity: null });
  f.owners.push(recovered);
  await recovered.recover(documentId);
  f.o = recovered;
  p = await preview(f.o);
  assert.equal(p.model3d.sample, 'octahedron');
  await req(f.o, 'apply', {
    commands: [{ operation: 'deleteModel3D', sheet: 'Sheet1', objectId: id, expectedObjectRevision: p.revision }],
  });
  assert.deepEqual(await preview(f.o, other), sibling);
  await req(f.o, 'undo');
  assert.equal((await preview(f.o)).model3d.sample, 'octahedron');
  await req(f.o, 'redo');
  await req(f.o, 'apply', {
    commands: [
      { operation: 'deleteModel3D', sheet: 'Sheet1', objectId: other, expectedObjectRevision: sibling.revision },
    ],
  });
  assert.equal((await req(f.o, 'objects', { sheet: 'Sheet1' })).total, 0);
  await checkSaved(f, [], 'empty');
});
test('placement edits reuse the immutable GLB, preserve actual SmartArt/shape siblings and remain structurally protected and support generic deletion', async (t) => {
  const f = await fixture(t);
  await req(f.o, 'apply', { commands: [create()] });
  await req(f.o, 'apply', {
    commands: [
      {
        operation: 'createShape',
        sheet: 'Sheet1',
        objectId: other,
        expectedObjectRevision: null,
        shape: { sample: 'text-note', text: 'Keep text' },
      },
    ],
  });
  const third = '33333333-3333-4333-8333-333333333333';
  await req(f.o, 'apply', {
    commands: [
      {
        operation: 'createSmartArt',
        sheet: 'Sheet1',
        objectId: third,
        expectedObjectRevision: null,
        smartart: smartartSampleModel('process'),
      },
    ],
  });
  const text = await preview(f.o, other),
    diagram = await preview(f.o, third);
  let p = await preview(f.o);
  const before = path.join(f.root, 'before.xlsx');
  await req(f.o, 'copy', { destination: before, expectedRevision: null });
  await req(f.o, 'apply', {
    commands: [update(p, undefined, { anchor: { cell: 'H7', widthEmu: 3000000, heightEmu: 2000000 } })],
  });
  p = await preview(f.o);
  assert.deepEqual(await preview(f.o, other), text);
  assert.deepEqual(await preview(f.o, third), diagram);
  const after = path.join(f.root, 'after.xlsx');
  await req(f.o, 'copy', { destination: after, expectedRevision: null });
  const a = unzipSync(await fs.readFile(before)),
    b = unzipSync(await fs.readFile(after));
  for (const name of Object.keys(a).filter(
    (n) => n.endsWith('.glb') || n.startsWith('xl/diagrams/') || n.endsWith('.rels') || n === '[Content_Types].xml'
  ))
    assert.deepEqual(b[name], a[name], name);
  await req(f.o, 'apply', {
    commands: [{ operation: 'deleteObject', sheet: 'Sheet1', objectId: id, expectedObjectRevision: p.revision }],
  });
  assert.equal((await req(f.o, 'objects', { sheet: 'Sheet1' })).total, 2);
  await req(f.o, 'undo');
  assert.deepEqual(await preview(f.o), p);
  const state = structuredClone(f.o.status());
  for (const command of [
    { operation: 'insertRow', sheet: 'Sheet1', row: 1 },
    { operation: 'insertColumn', sheet: 'Sheet1', column: 1 },
  ]) {
    await assert.rejects(req(f.o, 'apply', { commands: [command] }));
    assert.deepEqual(f.o.status(), state);
  }
});
test('native admission rejects forged meshes/metadata/hash/camera/type/edge/removal plans and failed batches without publication', async (t) => {
  const f = await fixture(t);
  await req(f.o, 'apply', { commands: [create()] });
  const p = await preview(f.o),
    state = structuredClone(f.o.status()),
    command = update(p, { color: '3366CC', rotation: { x: 25, y: -90, z: 30 } });
  const before = path.join(f.root, 'before.xlsx');
  await req(f.o, 'copy', { destination: before, expectedRevision: null });
  const parts = unzipSync(await fs.readFile(before)),
    rels = Object.keys(parts).find((n) => n.startsWith('xl/drawings/_rels/'));
  const faults = [
    (plan) => (plan.priorObjectRevision = '0'.repeat(64)),
    (plan) =>
      (plan.replacements[plan.drawingPart] = plan.replacements[plan.drawingPart].replace('z="180000000"', 'z="99"')),
    (plan) =>
      (plan.replacements[plan.drawingPart] = plan.replacements[plan.drawingPart].replace(
        'color="3366CC"',
        'color="FFFFFF"'
      )),
    (plan) => delete plan.binaryAdditions[plan.objectPart],
    (plan) => (plan.binaryAdditions[plan.objectPart].sha256 = '0'.repeat(64)),
    (plan) => {
      const desc = plan.binaryAdditions[plan.objectPart],
        bytes = Buffer.from(desc.base64, 'base64'),
        jl = bytes.readUInt32LE(12);
      bytes.writeFloatLE(9, 28 + jl);
      const digest = sha(bytes),
        old = plan.objectPart,
        next = old.replace(/-[0-9a-f]{64}\.glb$/, '-' + digest + '.glb');
      plan.binaryAdditions[next] = { base64: bytes.toString('base64'), sha256: digest };
      delete plan.binaryAdditions[old];
      plan.objectPart = next;
      for (const key of Object.keys(plan.replacements))
        plan.replacements[key] = plan.replacements[key]
          .replaceAll(old, next)
          .replaceAll(path.basename(old), path.basename(next));
    },
    (plan) =>
      (plan.replacements['[Content_Types].xml'] = plan.replacements['[Content_Types].xml'].replace(
        'model/gltf-binary',
        'image/png'
      )),
    (plan) =>
      (plan.replacements[rels] = plan.replacements[rels].replace(
        'Type="http://schemas.microsoft.com/office/2017/06/relationships/model3d"',
        'TargetMode="Internal" Type="http://schemas.microsoft.com/office/2017/06/relationships/model3d"'
      )),
    (plan) => plan.removed.push('custom/model3d-opaque.bin'),
    (plan) => (plan.replacements['custom/forged.xml'] = '<secret/>'),
  ];
  for (const fault of faults) {
    f.fault(fault);
    await assert.rejects(req(f.o, 'apply', { commands: [command] }));
    assert.deepEqual(f.o.status(), state);
  }
  f.fault((plan) => (plan.replacements[plan.objectPart] = '<forged/>'));
  await assert.rejects(req(f.o, 'apply', { commands: [update(p, undefined, { anchor: { cell: 'F3' } })] }));
  assert.deepEqual(f.o.status(), state);
  f.fault(null);
  for (const invalid of [
    update(p, { rotation: { x: 181, y: 0, z: 0 } }),
    update(p, { sample: 'foreign' }),
    update(p, { color: 'aa0000' }),
    { ...command, expectedObjectRevision: '0'.repeat(64) },
  ]) {
    await assert.rejects(req(f.o, 'apply', { commands: [invalid] }));
    assert.deepEqual(f.o.status(), state);
  }
  await assert.rejects(req(f.o, 'apply', { commands: [command, { sheet: 'missing', cell: 'A1', value: 999 }] }));
  assert.deepEqual(f.o.status(), state);
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
});
test('corrupt source bytes and shared/foreign/readonly graphs cannot be edited or silently reconstructed', async (t) => {
  const f = await fixture(t);
  await req(f.o, 'apply', { commands: [create()] });
  const p = await preview(f.o),
    good = path.join(f.root, 'good.xlsx');
  await req(f.o, 'copy', { destination: good, expectedRevision: null });
  const parts = unzipSync(await fs.readFile(good)),
    asset = Object.keys(parts).find((n) => n.endsWith('.glb'));
  parts[asset][parts[asset].length - 1] = 255;
  const bad = path.join(f.root, 'bad.xlsx');
  await fs.writeFile(bad, zipSync(parts));
  const o = await f.open(bad, path.join(f.root, 'bad-recovery')),
    state = structuredClone(o.status());
  await assert.rejects(preview(o));
  await assert.rejects(req(o, 'apply', { commands: [update(p, { color: 'FFFFFF' })] }));
  assert.deepEqual(o.status(), state);
  const clean = unzipSync(await fs.readFile(good)),
    rels = Object.keys(clean).find((n) => n.startsWith('xl/drawings/_rels/'));
  clean[rels] = strToU8(
    new TextDecoder()
      .decode(clean[rels])
      .replace(
        '</Relationships>',
        '<Relationship Id="shared" Type="http://schemas.microsoft.com/office/2017/06/relationships/model3d" Target="../media/' +
          path.basename(asset) +
          '"/></Relationships>'
      )
  );
  const shared = path.join(f.root, 'shared.xlsx');
  await fs.writeFile(shared, zipSync(clean));
  const s = await f.open(shared, path.join(f.root, 'shared-recovery')),
    sp = await preview(s);
  assert.equal(sp.model3d, undefined);
  assert.match(sp.protectedReason, /Shared/);
  await assert.rejects(req(s, 'apply', { commands: [update(sp, { color: 'FFFFFF' })] }));
  const protectedParts = unzipSync(await fs.readFile(good));
  protectedParts['xl/worksheets/sheet1.xml'] = strToU8(
    new TextDecoder()
      .decode(protectedParts['xl/worksheets/sheet1.xml'])
      .replace('</sheetData>', '</sheetData><sheetProtection sheet="1"/>')
  );
  const protectedFile = path.join(f.root, 'protected.xlsx');
  await fs.writeFile(protectedFile, zipSync(protectedParts));
  const readOnly = await f.open(protectedFile, path.join(f.root, 'readonly'));
  // Protection belongs to the sheet: the workbook stays editable, that sheet refuses.
  assert.equal(readOnly.status().view.protectedSheets.length, 1);
  await assert.rejects(req(readOnly, 'apply', { commands: [update(p, { color: 'FFFFFF' })] }));
});

test('Manual pending calculation stays pending through 3D authoring and source/copy stay independent', async (t) => {
  const f = await fixture(t);
  await req(f.o, 'apply', {
    commands: [{ type: 'setCalculationOptions', mode: 'manual', recalculateBeforeSave: false }],
  });
  await req(f.o, 'apply', { commands: [{ sheet: 'Sheet1', cell: 'B1', value: 100 }] });
  const cells = facts((await req(f.o, 'read', { sheet: 'Sheet1', range: 'A1:B2' })).cells),
    policy = structuredClone(f.o.status().view.calculationPolicy);
  assert(cells.A2.calculation.pending);
  await req(f.o, 'apply', { commands: [create('octahedron')] });
  let p = await preview(f.o);
  const before = structuredClone(f.o.status());
  await preview(f.o);
  assert.deepEqual(f.o.status(), before);
  await req(f.o, 'apply', { commands: [update(p, { rotation: { x: -180, y: 0, z: 180 }, color: 'FFFFFF' })] });
  assert.deepEqual(facts((await req(f.o, 'read', { sheet: 'Sheet1', range: 'A1:B2' })).cells), cells);
  assert.deepEqual(f.o.status().view.calculationPolicy, policy);
  const file = path.join(f.root, 'manual.xlsx');
  await req(f.o, 'copy', { destination: file, expectedRevision: null });
  assert.deepEqual(f.o.status().view.calculationPolicy, policy);
  const cold = await f.open(file, path.join(f.root, 'manual-cold'));
  assert.equal(cold.status().view.calculationPolicy.mode, 'manual');
  assert.equal((await preview(cold)).model3d.rotation.x, -180);
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
});
test('internal command capture detaches nested rotations and rejects undeclared command fields', () => {
  const command = create(),
    copy = captureModel3DCommands([command]);
  command.model3d.rotation.z = 99;
  command.anchor.cell = 'XFD1048576';
  assert.equal(copy[0].model3d.rotation.z, 0);
  assert.equal(copy[0].anchor.cell, 'E2');
  assert.throws(() => captureModel3DCommands([{ ...command, runScript: true }]));
});
