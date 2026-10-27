/** Icons: native owner, real custom paths, colors, recovery and forged profile refusal. */
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
import { ICON_SAMPLES } from '../core/icon-samples.mjs';
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

const iconCreate = (sample = 'icon-check', patch = {}) => ({
  operation: 'createShape',
  sheet: 'Sheet1',
  objectId: id,
  expectedObjectRevision: null,
  shape: { sample, text: '', color: '217346' },
  altText: 'Example icon',
  ...patch,
});
test('three vector assets have stable paths and three native-owned color/placement models save and reopen', async (t) => {
  for (const [i, s] of ICON_SAMPLES.entries()) {
    const svg = await fs.readFile(s.asset, 'utf8');
    assert(svg.includes('viewBox="0 0 100 100"'));
    assert(svg.includes('points="' + s.points.map((p) => p.join(',')).join(' ') + '"'));
    assert(!/<image|script|foreignObject|href/.test(svg));
    const f = await fixture(t);
    await req(f.o, 'apply', { commands: [iconCreate(s.id)] });
    let p = await preview(f.o);
    assert.deepEqual(p.shape, { sample: s.id, text: '', color: '217346' });
    assert.equal(p.anchor.widthEmu, 914400);
    assert.equal(p.anchor.heightEmu, 914400);
    const color = ['2563EB', 'DC2626', '7C3AED'][i];
    await req(f.o, 'apply', {
      commands: [
        {
          operation: 'updateShape',
          sheet: 'Sheet1',
          objectId: id,
          expectedObjectRevision: p.revision,
          patch: { color },
          anchor: { cell: 'B4', dxEmu: 100, dyEmu: 200, widthEmu: 1143000, heightEmu: 571500 },
          altText: 'Edited icon alternative text',
        },
      ],
    });
    p = await preview(f.o);
    const file = path.join(f.root, 'saved.xlsx');
    await saved(f.o, file);
    const check = spawnSync(
      process.env.XLSX_EDITOR_PYTHON ?? 'python3',
      [path.resolve('tests/support/check-icon.py'), f.source, file, s.id, color],
      { encoding: 'utf8' }
    );
    assert.equal(check.status, 0, check.stderr);
    console.log(check.stdout.trim());
    await f.o.close();
    const cold = await f.open(file, path.join(f.root, 'cold'));
    assert.deepEqual((await preview(cold)).shape, p.shape);
    assert.equal((await preview(cold)).altText, 'Edited icon alternative text');
    assert.deepEqual(await fs.readFile(f.source), f.bytes);
  }
});
test('captured icon color edits have one Undo/Redo, recovery, sibling retention and complete deletion', async (t) => {
  const f = await fixture(t);
  await req(f.o, 'apply', { commands: [create('text-note')] });
  const other = '44444444-4444-4444-8444-444444444444',
    c = iconCreate('icon-star', { objectId: other });
  const pending = req(f.o, 'apply', { commands: [c] });
  c.shape.color = 'DC2626';
  await pending;
  let p = await req(f.o, 'object', { sheet: 'Sheet1', objectId: other });
  assert.equal(p.shape.color, '217346');
  const before = f.o.status().state;
  await req(f.o, 'apply', {
    commands: [
      {
        operation: 'updateShape',
        sheet: 'Sheet1',
        objectId: other,
        expectedObjectRevision: p.revision,
        patch: { color: '2563EB' },
      },
    ],
  });
  const after = f.o.status().state;
  await req(f.o, 'undo');
  assert.equal(f.o.status().state, before);
  assert.equal((await req(f.o, 'object', { sheet: 'Sheet1', objectId: other })).shape.color, '217346');
  await req(f.o, 'redo');
  assert.equal(f.o.status().state, after);
  const documentId = f.o.status().documentId;
  await f.o.close();
  const o = new SpreadsheetDocument({ recoveryDirectory: path.join(f.root, 'recovery'), cacheIdentity: null });
  f.owners.push(o);
  await o.recover(documentId);
  p = await req(o, 'object', { sheet: 'Sheet1', objectId: other });
  assert.equal(p.shape.color, '2563EB');
  await req(o, 'apply', {
    commands: [{ operation: 'deleteShape', sheet: 'Sheet1', objectId: other, expectedObjectRevision: p.revision }],
  });
  assert.equal((await req(o, 'objects', { sheet: 'Sheet1' })).total, 1);
  assert.equal((await preview(o)).shape.sample, 'text-note');
  p = await preview(o);
  await req(o, 'apply', {
    commands: [{ operation: 'deleteObject', sheet: 'Sheet1', objectId: id, expectedObjectRevision: p.revision }],
  });
  assert.equal((await req(o, 'objects', { sheet: 'Sheet1' })).total, 0);
  const empty = await saved(o, path.join(f.root, 'empty.xlsx'));
  assert(!Object.keys(empty).some((n) => n.startsWith('xl/drawings/')));
});
test('invalid colors/text/paths and forged native plans cannot publish partial icon or sibling changes', async (t) => {
  const f = await fixture(t);
  await req(f.o, 'apply', { commands: [iconCreate()] });
  const p = await preview(f.o),
    before = structuredClone(f.o.status()),
    c = {
      operation: 'updateShape',
      sheet: 'Sheet1',
      objectId: id,
      expectedObjectRevision: p.revision,
      patch: { color: '2563EB' },
    };
  for (const bad of [
    { ...c, patch: { color: '#123456' } },
    { ...c, patch: { color: 'abc123' } },
    { ...c, patch: { text: 'Icon text' } },
    { ...c, patch: { path: 'foreign' } },
    { ...c, expectedObjectRevision: '0'.repeat(64) },
  ]) {
    await assert.rejects(req(f.o, 'apply', { commands: [bad] }));
    assert.deepEqual(f.o.status(), before);
  }
  for (const fault of [
    (p) => (p.priorObjectRevision = '0'.repeat(64)),
    (p) => {
      p.replacements[p.drawingPart] = p.replacements[p.drawingPart].replace('x="12" y="50"', 'x="13" y="50"');
    },
    (p) => {
      p.replacements[p.drawingPart] = p.replacements[p.drawingPart].replace('val="2563EB"', 'val="abc123"');
    },
    (p) => {
      p.replacements[p.drawingPart] = p.replacements[p.drawingPart].replace('<xdr:sp>', '<xdr:sp macro="run">');
    },
    (p) => {
      p.replacements[p.drawingPart] = p.replacements[p.drawingPart].replace('</xdr:sp>', '<xdr:txBody/></xdr:sp>');
    },
  ]) {
    f.fault(fault);
    await assert.rejects(req(f.o, 'apply', { commands: [c] }));
    assert.deepEqual(f.o.status(), before);
  }
  f.fault(null);
  await assert.rejects(req(f.o, 'apply', { commands: [c, { sheet: 'missing', cell: 'A1', value: 99 }] }));
  assert.deepEqual(f.o.status(), before);
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
});
