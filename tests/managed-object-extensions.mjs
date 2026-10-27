import test, { after } from 'node:test';
import { objectSourceMetadata } from './support/object-source-metadata.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { createHash } from 'node:crypto';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { indexedObjectPlan } from '../core/indexed-object-authoring.mjs';
import { chartPreviewGeometry } from '../core/chart-layout.mjs';
import { defaultPrintSettings } from '../core/print-model.mjs';
import { composePrint, printDocumentHtml } from '../core/print-layout.mjs';
const chartId = '11111111-1111-4111-8111-111111111111',
  imageId = '22222222-2222-4222-8222-222222222222';
const reference = path.resolve('tests/support/w7-reference-data/source.xlsx');
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
after(() => closeNativeWorkbookWorker());
async function fixture(t, source = reference) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'objects-extension-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const input = path.join(directory, 'source.xlsx');
  await fs.writeFile(input, objectSourceMetadata(await fs.readFile(source)));
  const original = await fs.readFile(input);
  let root;
  const bridge = async (request) => {
    if (request.operation === 'index') root = request.root;
    return native(request);
  };
  const book = new IndexedWorkbook(
    bridge,
    () => assert.fail('No compatibility writer'),
    (input) => calculateIndexedWorkbook(input, bridge)
  );
  await book.open(input);
  t.after(() => book.close());
  const save = async (state) => {
    const result = await book.export(state),
      dest = path.join(directory, `saved-${state}.xlsx`);
    await fs.copyFile(result.path, dest);
    return dest;
  };
  return { book, input, original, directory, root: () => root, save };
}
const chart = {
  operation: 'createChart',
  sheet: 'Sales',
  objectId: chartId,
  expectedObjectRevision: null,
  chart: { kind: 'bar', range: 'A1:C4', title: 'Monthly sales' },
};
const image = async () => ({
  operation: 'insertImage',
  sheet: 'Sales',
  objectId: imageId,
  expectedObjectRevision: null,
  asset: {
    mime: 'image/png',
    base64: (await fs.readFile('tests/support/w7-reference-data/marker.png')).toString('base64'),
  },
  altText: 'Logo',
});
test('horizontal bar owner preserves old states, native accepted cache and reversed axes on Save/reopen', async (t) => {
  const f = await fixture(t),
    made = await f.book.apply('0', [chart]),
    state = made.view.storage.state,
    entry = (await f.book.objects(state, 'Sales')).objects[0];
  const edited = await f.book.apply(state, [{ sheet: 'Sales', cell: 'B2', value: -21 }]),
    later = edited.view.storage.state;
  assert.equal((await f.book.objects('0', 'Sales')).total, 0);
  assert.equal((await f.book.object(state, 'Sales', chartId)).data.series[0].values[0], 2);
  assert.equal((await f.book.object(later, 'Sales', chartId)).data.series[0].values[0], -21);
  const saved = await f.save(later),
    parts = unzipSync(await fs.readFile(saved)),
    xml = strFromU8(parts[`xl/charts/workbook-chart-${chartId}.xml`]);
  assert.match(xml, /<c:barDir val="bar"\/>/);
  assert.match(xml, /<c:catAx>[\s\S]*?<c:axPos val="l"\/>/);
  assert.match(xml, /<c:valAx>[\s\S]*?<c:axPos val="b"\/>/);
  assert.match(xml, /<c:pt idx="0"><c:v>-21<\/c:v>/);
  for (const part of ['custom/opaque.bin', 'xl/worksheets/sheet2.xml', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels'])
    assert.deepEqual(parts[part], unzipSync(f.original)[part]);
  const g = await fixture(t, saved);
  assert.equal((await g.book.object('0', 'Sales', chartId)).chart.kind, 'bar');
  await assert.rejects(
    f.book.apply(later, [
      {
        operation: 'updateChart',
        sheet: 'Sales',
        objectId: chartId,
        expectedObjectRevision: 'a'.repeat(64),
        patch: { kind: 'column' },
      },
    ]),
    /changed|protected/
  );
  const changed = await f.book.apply(later, [
    {
      operation: 'updateChart',
      sheet: 'Sales',
      objectId: chartId,
      expectedObjectRevision: entry.revision,
      patch: { kind: 'column' },
    },
  ]);
  assert.equal((await f.book.object(changed.view.storage.state, 'Sales', chartId)).chart.kind, 'column');
  assert.equal((await f.book.object(later, 'Sales', chartId)).chart.kind, 'bar');
  assert.deepEqual(await fs.readFile(f.input), f.original);
});
test('image move/resize is compact and asset-preserving across old states, replay and saved graph', async (t) => {
  const f = await fixture(t),
    made = await f.book.apply('0', [chart, await image()]),
    state = made.view.storage.state,
    entry = (await f.book.objects(state, 'Sales')).objects.find((item) => item.id === imageId),
    originalPreview = await f.book.object(state, 'Sales', imageId);
  const command = {
    operation: 'updateImage',
    sheet: 'Sales',
    objectId: imageId,
    expectedObjectRevision: entry.revision,
    anchor: { cell: 'G8', widthEmu: 3810000, heightEmu: 3810000 },
    altText: 'Sales logo moved',
  };
  const changed = await f.book.apply(state, [command]),
    later = changed.view.storage.state;
  assert.equal(changed.commands[0].asset, undefined);
  assert.ok(Buffer.byteLength(JSON.stringify(changed.commands)) < 1024);
  const now = await f.book.object(later, 'Sales', imageId);
  assert.equal(now.anchor.cell, 'G8');
  assert.equal(now.anchor.widthEmu, 3810000);
  assert.equal(now.altText, 'Sales logo moved');
  assert.deepEqual(now.image, originalPreview.image);
  assert.equal((await f.book.object(state, 'Sales', imageId)).anchor.cell, 'E2');
  const replay = await f.book.apply(state, changed.commands);
  assert.deepEqual(await f.book.object(replay.view.storage.state, 'Sales', imageId), now);
  await assert.rejects(f.book.apply(later, [command]), /changed|protected/);
  await assert.rejects(
    f.book.apply(state, [{ ...command, anchor: { cell: 'G8', widthEmu: 3810000, heightEmu: 952500 } }]),
    { code: 'object-unsupported' }
  );
  const old = unzipSync(await fs.readFile(await f.save(state))),
    parts = unzipSync(await fs.readFile(await f.save(later)));
  const images = Object.keys(parts).filter((part) => part.startsWith('xl/media/'));
  assert.equal(images.length, 1);
  assert.equal(hash(parts[images[0]]), originalPreview.image.sha256);
  for (const part of Object.keys(old))
    if (!part.startsWith('xl/drawings/')) assert.deepEqual(parts[part], old[part], part);
  const g = await fixture(t, await f.save(later));
  assert.equal((await g.book.object('0', 'Sales', imageId)).anchor.cell, 'G8');
  assert.deepEqual(await fs.readFile(f.input), f.original);
});
test('moving one managed PNG alias preserves the shared asset and other anchor exactly', async (t) => {
  const f = await fixture(t),
    made = await f.book.apply('0', [await image()]),
    saved = await f.save(made.view.storage.state),
    parts = unzipSync(await fs.readFile(saved)),
    drawing = Object.keys(parts).find(
      (part) => part.startsWith('xl/drawings/') && part.endsWith('.xml') && !part.includes('/_rels/')
    ),
    rels = drawing.replace(/([^/]+)$/, '_rels/$1.rels'),
    alias = '33333333-3333-4333-8333-333333333333';
  const xml = strFromU8(parts[drawing]),
    anchor = /<xdr:oneCellAnchor[\s\S]*?<\/xdr:oneCellAnchor>/.exec(xml)[0],
    oldRel = 'workbookObject_' + imageId.replaceAll('-', ''),
    newRel = 'workbookObject_' + alias.replaceAll('-', ''),
    copy = anchor
      .replace('id="1"', 'id="2"')
      .replace('workbook:' + imageId, 'workbook:' + alias)
      .replace(oldRel, newRel);
  parts[drawing] = strToU8(xml.replace('</xdr:wsDr>', copy + '</xdr:wsDr>'));
  const edge = /<Relationship[^>]+\/>/.exec(strFromU8(parts[rels]))[0];
  parts[rels] = strToU8(
    strFromU8(parts[rels]).replace('</Relationships>', edge.replace(oldRel, newRel) + '</Relationships>')
  );
  const shared = path.join(f.directory, 'shared.xlsx');
  await fs.writeFile(shared, zipSync(parts));
  const g = await fixture(t, shared),
    before = (await g.book.objects('0', 'Sales')).objects,
    entry = before.find((item) => item.id === alias),
    preview = await g.book.object('0', 'Sales', alias),
    changed = await g.book.apply('0', [
      {
        operation: 'updateImage',
        sheet: 'Sales',
        objectId: alias,
        expectedObjectRevision: entry.revision,
        anchor: { cell: 'G9', widthEmu: 1905000, heightEmu: 1905000 },
      },
    ]);
  assert.equal((await g.book.object(changed.view.storage.state, 'Sales', alias)).anchor.cell, 'G9');
  assert.deepEqual(
    await g.book.object(changed.view.storage.state, 'Sales', imageId),
    await g.book.object('0', 'Sales', imageId)
  );
  assert.deepEqual((await g.book.object(changed.view.storage.state, 'Sales', alias)).image, preview.image);
  const output = unzipSync(await fs.readFile(await g.save(changed.view.storage.state)));
  assert.equal(Object.keys(output).filter((part) => part.startsWith('xl/media/')).length, 1);
  for (const part of Object.keys(parts)) if (part !== drawing) assert.deepEqual(output[part], parts[part], part);
});
test('native placement admission rejects asset replacement, distorted ratio and immutable picture edits', async (t) => {
  const f = await fixture(t),
    made = await f.book.apply('0', [await image()]),
    saved = await f.save(made.view.storage.state),
    g = await fixture(t, saved),
    entry = (await g.book.objects('0', 'Sales')).objects[0],
    root = g.root();
  const prepared = await indexedObjectPlan(
    {
      metadataBytes: await fs.readFile(path.join(root, 'metadata.xlsx')),
      view: g.book.view('0'),
      readMedia: async (part) => unzipSync(await fs.readFile(saved))[part],
    },
    {
      operation: 'updateImage',
      sheet: 'Sales',
      objectId: imageId,
      expectedObjectRevision: entry.revision,
      anchor: { cell: 'F6', widthEmu: 1905000, heightEmu: 1905000 },
    }
  );
  const exported = await g.book.export('0');
  const attempt = async (plan) => {
    await fs.writeFile(path.join(root, 'package-metadata.json'), JSON.stringify(plan));
    return native({
      action: 'workbook',
      operation: 'package-edit',
      root,
      version: 0,
      source: path.basename(exported.path),
    });
  };
  await attempt(prepared.plan);
  for (const mutate of [
    (p) => {
      p.replacements[p.objectPart] = 'forged bytes';
    },
    (p) => {
      p.replacements[p.drawingPart] = p.replacements[p.drawingPart].replace('cy="1905000"', 'cy="952500"');
    },
    (p) => {
      p.replacements[p.drawingPart] = p.replacements[p.drawingPart].replace('prst="rect"', 'prst="ellipse"');
    },
    (p) => {
      p.replacements[p.drawingPart] = p.replacements[p.drawingPart].replace('r:embed="', 'r:embed="forged');
    },
    (p) => {
      p.binaryAdditions[p.objectPart] = {};
    },
    (p) => {
      p.replacements[p.drawingPart] = p.replacements[p.drawingPart].replace(
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
        ''
      );
    },
  ]) {
    const plan = structuredClone(prepared.plan);
    mutate(plan);
    await assert.rejects(attempt(plan), /mutation|PNG|picture|placement|relationship|addition|geometry|edge|namespace/);
  }
  assert.equal((await g.book.object('0', 'Sales', imageId)).anchor.cell, 'E2');
  assert.deepEqual(await fs.readFile(g.input), g.original);
});
test('grid and PDF use identical finite bar geometry with gaps, hidden rows and opposite signs', () => {
  const object = {
    id: chartId,
    name: 'Bar',
    kind: 'chart',
    chart: { kind: 'bar', title: 'Bar', legend: 'none' },
    data: {
      categories: ['Positive', 'Negative', 'Gap', 'Hidden'],
      series: [{ name: 'Value', values: [1e308, -1e308, null, 3] }],
      hidden: [false, false, false, true],
    },
  };
  const geometry = chartPreviewGeometry(object);
  assert.equal(geometry.bars.length, 2);
  assert.ok(geometry.bars[0].y > geometry.bars[1].y);
  assert.ok(geometry.bars[0].x > geometry.bars[1].x);
  assert.ok(
    geometry.bars.every((bar) =>
      Object.values(bar)
        .filter((value) => typeof value === 'number')
        .every(Number.isFinite)
    )
  );
  const snapshot = {
    schema: 1,
    fingerprint: 'a'.repeat(64),
    revision: 'r1',
    sheet: 'Sales',
    settings: { ...defaultPrintSettings(), area: 'A1:A1' },
    range: 'A1:A1',
    rows: [{ index: 0, heightPt: 400 }],
    columns: [{ index: 0, widthPt: 400 }],
    cells: [{ ref: 'A1', text: '', style: {} }],
    objects: [{ ...object, anchor: { cell: 'A1' }, boxPt: [0, 0, 350, 220] }],
    merges: [],
  };
  const layout = composePrint(snapshot),
    html = printDocumentHtml(layout);
  for (const bar of geometry.bars)
    assert.ok(html.includes(`x="${bar.x}" y="${bar.y}" width="${bar.width}" height="${bar.height}"`));
  assert.ok(!html.includes('NaN'));
  assert.ok(!html.includes('Infinity'));
});
