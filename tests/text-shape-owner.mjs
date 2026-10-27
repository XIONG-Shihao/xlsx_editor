/** Native-owned editable text/basic-shape models; UI availability is separate. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { spawnSync } from 'node:child_process';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook } from '../core/workbook-package.mjs';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { SHAPE_SAMPLES, normalizeShape } from '../core/shape-model.mjs';
after(() => closeNativeWorkbookWorker());
const id = '11111111-1111-4111-8111-111111111111',
  imageId = '22222222-2222-4222-8222-222222222222',
  chartId = '33333333-3333-4333-8333-333333333333';
const req = (o, operation, more = {}) =>
  o.request({ operation, documentId: o.status().documentId, state: o.status().state, ...more });
async function fixture(t) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'text-shape-owner-')),
    source = path.join(root, 'source.xlsx'),
    parts = unzipSync(createWorkbook()),
    owners = [];
  parts['custom/shape-opaque.bin'] = Uint8Array.from([1, 4, 8, 255]);
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B3"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Category</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>A</t></is></c><c r="B2"><v>7</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>B</t></is></c><c r="B3"><v>8</v></c></row></sheetData></worksheet>'
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
  const open = async (file = source, dir = path.join(root, 'recovery')) => {
    const o = new SpreadsheetDocument({ native: bridge, recoveryDirectory: dir, cacheIdentity: null });
    owners.push(o);
    await o.open(file, { fresh: true });
    return o;
  };
  const o = await open();
  t.after(async () => {
    for (const owner of owners) await owner.close();
    await fs.rm(root, { recursive: true, force: true });
  });
  return { root, source, bytes, o, owners, open, fault: (value) => (fault = value) };
}
const create = (sample = 'text-plain', patch = {}) => ({
  operation: 'createShape',
  sheet: 'Sheet1',
  objectId: id,
  expectedObjectRevision: null,
  shape: { sample, text: 'Text <&é\nSecond line' },
  altText: 'Example text/shape',
  anchor: { cell: 'D2' },
  ...patch,
});
const preview = (o) => req(o, 'object', { sheet: 'Sheet1', objectId: id });
const saved = async (o, file) => {
  await req(o, 'copy', { destination: file, expectedRevision: null });
  return unzipSync(await fs.readFile(file));
};
test('two text and three basic-shape samples save real DrawingML, retain cells and reopen their owned model', async (t) => {
  for (const sample of SHAPE_SAMPLES) {
    const f = await fixture(t);
    await req(f.o, 'apply', { commands: [create(sample.id)] });
    let object = await preview(f.o);
    assert.equal(object.kind, 'shape');
    assert(object.canEdit && object.canDelete);
    assert.deepEqual(object.shape, { sample: sample.id, text: 'Text <&é\nSecond line' });
    assert(!JSON.stringify(object).includes('xl/drawings'));
    const baseline = (await req(f.o, 'read', { sheet: 'Sheet1', range: 'A1:B3' })).cells;
    const original = object.revision;
    await req(f.o, 'apply', {
      commands: [
        {
          operation: 'updateShape',
          sheet: 'Sheet1',
          objectId: id,
          expectedObjectRevision: original,
          patch: { text: 'Edited "&é\nNext' },
          anchor: { cell: 'B4', dxEmu: 100, dyEmu: 200, widthEmu: 1143000, heightEmu: 571500 },
          altText: 'Edited alternative text',
        },
      ],
    });
    object = await preview(f.o);
    assert.equal(object.shape.text, 'Edited "&é\nNext');
    assert.equal(object.anchor.cell, 'B4');
    assert.equal(object.anchor.widthEmu, 1143000);
    assert.deepEqual((await req(f.o, 'read', { sheet: 'Sheet1', range: 'A1:B3' })).cells, baseline);
    const file = path.join(f.root, 'saved.xlsx');
    await saved(f.o, file);
    const check = spawnSync(
      process.env.XLSX_EDITOR_PYTHON ?? 'python3',
      [path.join(process.cwd(), 'tests/support/check-text-shape.py'), f.source, file, sample.id],
      { encoding: 'utf8' }
    );
    assert.equal(check.status, 0, check.stderr);
    console.log(check.stdout.trim());
    await f.o.close();
    const cold = await f.open(file, path.join(f.root, 'cold'));
    const loaded = await preview(cold);
    assert.equal(loaded.shape.sample, sample.id);
    assert.equal(loaded.shape.text, object.shape.text);
    assert.equal(loaded.altText, 'Edited alternative text');
    assert(loaded.canEdit);
    assert.deepEqual(await fs.readFile(f.source), f.bytes);
  }
});
test('shape edit/move/delete has one Undo, captured commands and durable recovery', async (t) => {
  const f = await fixture(t),
    before = f.o.status().state,
    c = create('shape-arrow'),
    pending = req(f.o, 'apply', { commands: [c] });
  c.shape.text = 'Retargeted';
  c.anchor.cell = 'A100';
  await pending;
  assert.equal((await preview(f.o)).anchor.cell, 'D2');
  assert.equal((await preview(f.o)).shape.text, 'Text <&é\nSecond line');
  const accepted = f.o.status().state;
  await req(f.o, 'undo');
  assert.equal(f.o.status().state, before);
  assert.equal((await req(f.o, 'objects', { sheet: 'Sheet1' })).total, 0);
  await req(f.o, 'redo');
  assert.equal(f.o.status().state, accepted);
  const documentId = f.o.status().documentId;
  await f.o.close();
  const o = new SpreadsheetDocument({ recoveryDirectory: path.join(f.root, 'recovery'), cacheIdentity: null });
  f.owners.push(o);
  await o.recover(documentId);
  let p = await preview(o);
  await req(o, 'apply', {
    commands: [{ operation: 'deleteObject', sheet: 'Sheet1', objectId: id, expectedObjectRevision: p.revision }],
  });
  assert.equal((await req(o, 'objects', { sheet: 'Sheet1' })).total, 0);
  await req(o, 'undo');
  p = await preview(o);
  assert.equal(p.shape.sample, 'shape-arrow');
  await req(o, 'redo');
  assert.equal((await req(o, 'objects', { sheet: 'Sheet1' })).total, 0);
  const empty = await saved(o, path.join(f.root, 'empty.xlsx'));
  assert(!Object.keys(empty).some((k) => k.startsWith('xl/drawings/')));
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
});
test('new shape/text mutations preserve chart/PNG siblings and existing object editing still works', async (t) => {
  const f = await fixture(t);
  await req(f.o, 'apply', {
    commands: [
      {
        operation: 'createChart',
        sheet: 'Sheet1',
        objectId: chartId,
        expectedObjectRevision: null,
        chart: { kind: 'column', range: 'A1:B3', title: 'Amounts' },
      },
    ],
  });
  const png = (await fs.readFile(path.resolve('tests/support/w7-reference-data/marker.png'))).toString('base64');
  await req(f.o, 'apply', {
    commands: [
      {
        operation: 'insertImage',
        sheet: 'Sheet1',
        objectId: imageId,
        expectedObjectRevision: null,
        asset: { mime: 'image/png', base64: png },
      },
    ],
  });
  const first = await req(f.o, 'object', { sheet: 'Sheet1', objectId: chartId }),
    image = await req(f.o, 'object', { sheet: 'Sheet1', objectId: imageId });
  await req(f.o, 'apply', { commands: [create('shape-ellipse')] });
  assert.deepEqual(await req(f.o, 'object', { sheet: 'Sheet1', objectId: chartId }), first);
  assert.deepEqual(await req(f.o, 'object', { sheet: 'Sheet1', objectId: imageId }), image);
  await req(f.o, 'apply', {
    commands: [
      {
        operation: 'updateImage',
        sheet: 'Sheet1',
        objectId: imageId,
        expectedObjectRevision: image.revision,
        anchor: { cell: 'J4' },
      },
    ],
  });
  let shape = await preview(f.o);
  await req(f.o, 'apply', {
    commands: [
      {
        operation: 'updateShape',
        sheet: 'Sheet1',
        objectId: id,
        expectedObjectRevision: shape.revision,
        patch: { text: 'Coexisting' },
      },
    ],
  });
  const chart = await req(f.o, 'object', { sheet: 'Sheet1', objectId: chartId });
  await req(f.o, 'apply', {
    commands: [
      {
        operation: 'updateChart',
        sheet: 'Sheet1',
        objectId: chartId,
        expectedObjectRevision: chart.revision,
        patch: { title: 'New title' },
      },
    ],
  });
  shape = await preview(f.o);
  await req(f.o, 'apply', {
    commands: [{ operation: 'deleteShape', sheet: 'Sheet1', objectId: id, expectedObjectRevision: shape.revision }],
  });
  assert.equal((await req(f.o, 'objects', { sheet: 'Sheet1' })).total, 2);
  assert.equal((await req(f.o, 'object', { sheet: 'Sheet1', objectId: chartId })).chart.title, 'New title');
  assert.equal((await req(f.o, 'object', { sheet: 'Sheet1', objectId: imageId })).anchor.cell, 'J4');
});
test('stale/invalid/forged shape profiles and mixed-batch failure cannot publish partial state', async (t) => {
  const f = await fixture(t);
  await req(f.o, 'apply', { commands: [create()] });
  const p = await preview(f.o),
    before = structuredClone(f.o.status()),
    command = {
      operation: 'updateShape',
      sheet: 'Sheet1',
      objectId: id,
      expectedObjectRevision: p.revision,
      patch: { text: 'Accepted' },
    };
  for (const c of [
    { ...command, expectedObjectRevision: '0'.repeat(64) },
    { ...command, patch: { text: 'Too many\n'.repeat(9) } },
    { ...command, patch: { sample: 'foreign' } },
    { ...command, patch: { fontSize: 99 } },
    { ...command, sheet: 'missing' },
  ]) {
    await assert.rejects(req(f.o, 'apply', { commands: [c] }));
    assert.deepEqual(f.o.status(), before);
  }
  for (const fault of [
    (p) => (p.priorObjectRevision = '0'.repeat(64)),
    (p) => {
      const file = p.drawingPart;
      p.replacements[file] = p.replacements[file].replace('sz="1400"', 'sz="9900"');
    },
    (p) => {
      const file = p.drawingPart;
      p.replacements[file] = p.replacements[file].replace('<xdr:sp>', '<xdr:sp macro="unqualified">');
    },
    (p) => {
      const file = p.drawingPart;
      p.replacements[file] = p.replacements[file].replace(
        '<a:latin typeface="Arial"/>',
        '<a:latin typeface="Arial"/><a:hlinkClick/>'
      );
    },
  ]) {
    f.fault(fault);
    await assert.rejects(req(f.o, 'apply', { commands: [command] }));
    assert.deepEqual(f.o.status(), before);
  }
  f.fault(null);
  await assert.rejects(req(f.o, 'apply', { commands: [command, { sheet: 'missing', cell: 'A1', value: 99 }] }));
  assert.deepEqual(f.o.status(), before);
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
});
