import { defaultPrintSettings } from '../core/print-model.mjs';
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createHash } from 'node:crypto';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook, readWorkbook } from '../core/workbook-package.mjs';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { worksheetBackgroundSamples } from '../core/page-style-samples.mjs';
import { backgroundCommand } from '../core/background-command.mjs';
const roots = [],
  hash = (b) => createHash('sha256').update(b).digest('hex');
after(async () => {
  await closeNativeWorkbookWorker();
  for (const root of roots) await fs.rm(root, { recursive: true, force: true });
});
async function fixture(t, { foreign = false } = {}) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'background-owner-'));
  roots.push(root);
  const parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Body</t></is></c><c r="B1"><f>SUM(D1,1)</f><v>777</v></c><c r="C1"><v>14</v></c><c r="D1"><v>5</v></c></row></sheetData>' +
      (foreign ? '<picture r:id="foreign"/>' : '') +
      '</worksheet>'
  );
  parts['custom/opaque.bin'] = new Uint8Array([3, 9, 20, 255]);
  if (foreign) {
    parts['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="foreign" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/imported.png"/></Relationships>'
    );
    parts['xl/media/imported.png'] = Buffer.from(worksheetBackgroundSamples[0].base64, 'base64');
  }
  const bytes = zipSync(parts),
    filename = path.join(root, 'source.xlsx'),
    recoveryDirectory = path.join(root, 'recovery');
  await fs.writeFile(filename, bytes);
  const owners = [];
  const open = async (file = filename) => {
    const o = new SpreadsheetDocument({ recoveryDirectory });
    owners.push(o);
    await o.open(file, { fresh: true });
    return o;
  };
  const recover = async (id) => {
    const o = new SpreadsheetDocument({ recoveryDirectory });
    owners.push(o);
    await o.recover(id);
    return o;
  };
  t.after(async () => {
    for (const o of owners) await o.close();
  });
  return { root, filename, parts, bytes, open, recover };
}
const request = (o, operation, rest = {}) => o.request({ operation, state: o.status().state, ...rest });
const read = async (o) => (await request(o, 'read', { sheet: 'Sheet1', range: 'A1:D1' })).cells;
const apply = (o, command) => request(o, 'apply', { commands: Array.isArray(command) ? command : [command] });
const command = (o, id) =>
  backgroundCommand({
    type: id ? 'setSheetBackground' : 'removeSheetBackground',
    sheet: 'Sheet1',
    expectedBackgroundRevision: o.status().view.sheets[0].background?.revision ?? null,
    ...(id ? { sample: id } : {}),
  });
const calculation = (cell) =>
  Object.fromEntries(Object.entries(cell.calculation ?? {}).filter(([key]) => key !== 'revision'));
test('two actual PNG backgrounds use worksheet pictures, save/reopen, Undo/Redo and durable recovery', async (t) => {
  for (const sample of worksheetBackgroundSamples) {
    const f = await fixture(t),
      o = await f.open(),
      start = o.status().state,
      before = await read(o);
    await apply(o, command(o, sample.id));
    assert.equal(o.status().view.sheets[0].background.sample, sample.id);
    assert.deepEqual(calculation((await read(o)).B1), calculation(before.B1));
    const printBefore = structuredClone(o.status()),
      printRequest = { documentId: o.status().documentId, state: o.status().state, sheet: 'Sheet1' },
      job = await o.request({
        operation: 'printPrepare',
        ...printRequest,
        settings: { ...defaultPrintSettings(), area: 'A1:D1' },
      });
    const page = job.pages[0];
    assert.equal(page.objects.length, 0);
    const html = await o.request({ operation: 'printDocument', ...printRequest, jobId: job.id });
    assert(!html.html.includes(sample.base64));
    assert.deepEqual(o.status(), printBefore);
    await o.request({ operation: 'printRelease', ...printRequest, jobId: job.id });
    const accepted = o.status().state;
    await request(o, 'undo');
    assert.equal(o.status().state, start);
    assert.equal(o.status().view.sheets[0].background, undefined);
    await request(o, 'redo');
    assert.equal(o.status().state, accepted);
    const saved = path.join(f.root, sample.id + '.xlsx');
    await request(o, 'copy', { destination: saved, expectedRevision: null });
    const actual = unzipSync(await fs.readFile(saved)),
      images = Object.keys(actual).filter((name) => name.startsWith('xl/media/'));
    assert.equal(images.length, 1);
    assert.equal(hash(actual[images[0]]), sample.sha256);
    const xml = strFromU8(actual['xl/worksheets/sheet1.xml']);
    assert(xml.includes('<picture'));
    assert(!xml.includes('<drawing'));
    assert.equal(readWorkbook(await fs.readFile(saved)).sheets[0].cells.B1.formula, 'SUM(D1,1)');
    assert.deepEqual(actual['xl/styles.xml'], f.parts['xl/styles.xml']);
    assert.deepEqual(actual['custom/opaque.bin'], f.parts['custom/opaque.bin']);
    assert.equal(hash(await fs.readFile(f.filename)), hash(f.bytes));
    const id = o.status().documentId;
    await o.close();
    const recovered = await f.recover(id);
    assert.equal(recovered.status().view.sheets[0].background.sample, sample.id);
    await request(recovered, 'undo');
    assert.equal(recovered.status().view.sheets[0].background, undefined);
    await request(recovered, 'redo');
    assert.equal(recovered.status().view.sheets[0].background.sample, sample.id);
    await recovered.close();
    const cold = await f.open(saved);
    assert.equal(cold.status().view.sheets[0].background.sample, sample.id);
    await cold.close();
  }
});
test('switch/remove retire only owned assets; repeated apply/absent removal preserve head and Redo', async (t) => {
  const f = await fixture(t),
    o = await f.open();
  const absent = structuredClone(o.status());
  await apply(o, command(o, null));
  assert.deepEqual(o.status(), absent);
  await apply(o, command(o, 'graph-paper'));
  const first = o.status().state;
  await apply(o, command(o, 'soft-dots'));
  await request(o, 'undo');
  assert.equal(o.status().state, first);
  const before = structuredClone(o.status()),
    manifestPath = path.join(f.root, 'recovery', 'draft-' + before.documentId, 'current.json'),
    manifest = await fs.readFile(manifestPath);
  await apply(o, command(o, 'graph-paper'));
  assert.deepEqual(o.status(), before);
  assert.deepEqual(await fs.readFile(manifestPath), manifest);
  await apply(o, command(o, 'soft-dots'));
  const switched = path.join(f.root, 'switched.xlsx');
  await request(o, 'copy', { destination: switched, expectedRevision: null });
  const parts = unzipSync(await fs.readFile(switched));
  assert.equal(Object.keys(parts).filter((name) => name.startsWith('xl/media/')).length, 1);
  await apply(o, command(o, null));
  assert.equal(o.status().view.sheets[0].background, undefined);
  const removed = path.join(f.root, 'removed.xlsx');
  await request(o, 'copy', { destination: removed, expectedRevision: null });
  const empty = unzipSync(await fs.readFile(removed));
  assert(!strFromU8(empty['xl/worksheets/sheet1.xml']).includes('<picture'));
  assert.equal(Object.keys(empty).filter((name) => name.startsWith('xl/media/')).length, 0);
  assert.deepEqual(empty['custom/opaque.bin'], f.parts['custom/opaque.bin']);
});
test('Manual pending formulas and subsequent ordinary cell edits work with a saved background', async (t) => {
  const f = await fixture(t),
    o = await f.open();
  await apply(o, { type: 'setCalculationOptions', mode: 'manual', recalculateBeforeSave: false });
  await apply(o, { sheet: 'Sheet1', cell: 'D1', value: 9 });
  const before = await read(o);
  assert(before.B1.calculation.pending);
  await apply(o, command(o, 'soft-dots'));
  assert.deepEqual(calculation((await read(o)).B1), calculation(before.B1));
  await apply(o, { sheet: 'Sheet1', cell: 'A1', value: 'Ordinary edit still works' });
  assert.equal((await read(o)).A1.value, 'Ordinary edit still works');
  assert.equal(o.status().view.sheets[0].background.sample, 'soft-dots');
  const saved = path.join(f.root, 'manual.xlsx');
  await request(o, 'copy', { destination: saved, expectedRevision: null });
  await o.close();
  const cold = await f.open(saved);
  assert((await read(cold)).B1.calculation.pending);
  assert.equal(cold.status().view.sheets[0].background.sample, 'soft-dots');
});
test('stale/closed/late batch failure and imported backgrounds retain accepted owner and source', async (t) => {
  const f = await fixture(t),
    o = await f.open();
  await apply(o, command(o, 'graph-paper'));
  const before = structuredClone(o.status()),
    cells = await read(o);
  for (const c of [
    { ...command(o, 'soft-dots'), expectedBackgroundRevision: null },
    { ...command(o, 'soft-dots'), bytes: 'caller' },
    { ...command(o, 'soft-dots'), sample: 'imported' },
  ])
    await assert.rejects(apply(o, c));
  await assert.rejects(apply(o, [command(o, 'soft-dots'), { sheet: 'missing', cell: 'A1', value: 'late' }]));
  assert.deepEqual(o.status(), before);
  assert.deepEqual(await read(o), cells);
  const foreign = await fixture(t, { foreign: true }),
    imported = await foreign.open();
  assert.equal(imported.status().view.sheets[0].background.canEdit, false);
  const unchanged = structuredClone(imported.status());
  await assert.rejects(apply(imported, command(imported, 'graph-paper')));
  assert.deepEqual(imported.status(), unchanged);
  assert.equal(hash(await fs.readFile(foreign.filename)), hash(foreign.bytes));
});
