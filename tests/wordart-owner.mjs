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
import { WORDART_SAMPLES } from '../core/wordart-samples.mjs';
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

const art = (sample = 'wordart-blue', patch = {}) => ({
  operation: 'createShape',
  sheet: 'Sheet1',
  objectId: id,
  expectedObjectRevision: null,
  shape: { sample, text: 'Hello <&é\nSecond' },
  altText: 'WordArt title',
  ...patch,
});
test('three editable WordArt fill/outline/shadow models preserve cells and save/reopen real text effects', async (t) => {
  for (const s of WORDART_SAMPLES) {
    const f = await fixture(t);
    await req(f.o, 'apply', { commands: [art(s.id)] });
    let p = await preview(f.o);
    assert.deepEqual(p.shape, { sample: s.id, text: 'Hello <&é\nSecond' });
    assert.equal(p.anchor.widthEmu, 320 * 9525);
    assert.equal(p.anchor.heightEmu, 120 * 9525);
    await req(f.o, 'apply', {
      commands: [
        {
          operation: 'updateShape',
          sheet: 'Sheet1',
          objectId: id,
          expectedObjectRevision: p.revision,
          patch: { text: 'Edited "&é\nNext' },
          anchor: { cell: 'B4', dxEmu: 100, dyEmu: 200, widthEmu: 3048000, heightEmu: 1143000 },
          altText: 'Edited WordArt alternative text',
        },
      ],
    });
    p = await preview(f.o);
    const file = path.join(f.root, 'saved.xlsx');
    await saved(f.o, file);
    const check = spawnSync(
      process.env.XLSX_EDITOR_PYTHON ?? 'python3',
      [path.resolve('tests/support/check-wordart.py'), f.source, file, s.id],
      { encoding: 'utf8' }
    );
    assert.equal(check.status, 0, check.stderr);
    console.log(check.stdout.trim());
    await f.o.close();
    const cold = await f.open(file, path.join(f.root, 'cold'));
    assert.deepEqual((await preview(cold)).shape, p.shape);
    assert.equal((await preview(cold)).altText, 'Edited WordArt alternative text');
    assert.deepEqual(await fs.readFile(f.source), f.bytes);
  }
});
test('captured WordArt edits/style switch have Undo/Redo, recovery and retain vector icon sibling', async (t) => {
  const f = await fixture(t),
    c = art(),
    run = req(f.o, 'apply', { commands: [c] });
  c.shape.text = 'Caller mutation';
  await run;
  assert.equal((await preview(f.o)).shape.text, 'Hello <&é\nSecond');
  const icon = '44444444-4444-4444-8444-444444444444';
  await req(f.o, 'apply', {
    commands: [
      {
        operation: 'createShape',
        sheet: 'Sheet1',
        objectId: icon,
        expectedObjectRevision: null,
        shape: { sample: 'icon-star', text: '', color: '2563EB' },
      },
    ],
  });
  const sibling = await req(f.o, 'object', { sheet: 'Sheet1', objectId: icon }),
    p = await preview(f.o),
    before = f.o.status().state;
  await req(f.o, 'apply', {
    commands: [
      {
        operation: 'updateShape',
        sheet: 'Sheet1',
        objectId: id,
        expectedObjectRevision: p.revision,
        patch: { sample: 'wordart-shadow', text: 'Restyled' },
      },
    ],
  });
  const accepted = f.o.status().state;
  await req(f.o, 'undo');
  assert.equal(f.o.status().state, before);
  assert.equal((await preview(f.o)).shape.sample, 'wordart-blue');
  await req(f.o, 'redo');
  assert.equal(f.o.status().state, accepted);
  assert.deepEqual(await req(f.o, 'object', { sheet: 'Sheet1', objectId: icon }), sibling);
  const documentId = f.o.status().documentId;
  await f.o.close();
  const o = new SpreadsheetDocument({ recoveryDirectory: path.join(f.root, 'recovery'), cacheIdentity: null });
  f.owners.push(o);
  await o.recover(documentId);
  const recovered = await preview(o);
  assert.deepEqual(recovered.shape, { sample: 'wordart-shadow', text: 'Restyled' });
  await req(o, 'apply', {
    commands: [{ operation: 'deleteShape', sheet: 'Sheet1', objectId: id, expectedObjectRevision: recovered.revision }],
  });
  assert.equal((await req(o, 'objects', { sheet: 'Sheet1' })).total, 1);
  const other = await req(o, 'object', { sheet: 'Sheet1', objectId: icon });
  await req(o, 'apply', {
    commands: [{ operation: 'deleteObject', sheet: 'Sheet1', objectId: icon, expectedObjectRevision: other.revision }],
  });
  assert.equal((await req(o, 'objects', { sheet: 'Sheet1' })).total, 0);
});
test('WordArt text/effect bounds, stale state and forged runs refuse without partial publication', async (t) => {
  const f = await fixture(t);
  await req(f.o, 'apply', { commands: [art('wordart-shadow')] });
  const p = await preview(f.o),
    before = structuredClone(f.o.status()),
    c = {
      operation: 'updateShape',
      sheet: 'Sheet1',
      objectId: id,
      expectedObjectRevision: p.revision,
      patch: { text: 'Accepted' },
    };
  for (const bad of [
    { ...c, patch: { text: 'x'.repeat(257) } },
    { ...c, patch: { text: 'a\nb\nc\nd\ne' } },
    { ...c, patch: { fontSize: 99 } },
    { ...c, patch: { color: '2563EB' } },
    { ...c, expectedObjectRevision: '0'.repeat(64) },
  ]) {
    await assert.rejects(req(f.o, 'apply', { commands: [bad] }));
    assert.deepEqual(f.o.status(), before);
  }
  for (const fault of [
    (p) => (p.priorObjectRevision = '0'.repeat(64)),
    (p) => {
      p.replacements[p.drawingPart] = p.replacements[p.drawingPart].replace('sz="3200"', 'sz="9900"');
    },
    (p) => {
      p.replacements[p.drawingPart] = p.replacements[p.drawingPart].replace('val="35000"', 'val="90000"');
    },
    (p) => {
      p.replacements[p.drawingPart] = p.replacements[p.drawingPart].replace('dist="53882"', 'dist="999999"');
    },
    (p) => {
      p.replacements[p.drawingPart] = p.replacements[p.drawingPart].replace(
        '<a:latin typeface="Arial"/>',
        '<a:latin typeface="Arial"/><a:hlinkClick/>'
      );
    },
    (p) => {
      p.replacements[p.drawingPart] = p.replacements[p.drawingPart].replace('<xdr:sp>', '<xdr:sp macro="run">');
    },
  ]) {
    f.fault(fault);
    await assert.rejects(req(f.o, 'apply', { commands: [c] }));
    assert.deepEqual(f.o.status(), before);
  }
  f.fault(null);
  await assert.rejects(req(f.o, 'apply', { commands: [c, { sheet: 'missing', cell: 'A1', value: 7 }] }));
  assert.deepEqual(f.o.status(), before);
});
