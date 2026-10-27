import { test, after } from 'node:test';
import { objectSourceMetadata } from './support/object-source-metadata.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createHash, randomBytes } from 'node:crypto';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import {
  indexedObjectPlan,
  indexedObjectList,
  indexedObjectPreview,
  indexedObjectCachePlan,
  indexedObjectReferenceGuard,
  indexedPngPayload,
  indexedObjectHasManagedCharts,
} from '../core/indexed-object-authoring.mjs';
import { checkObjectPackage, inventoryObjectPackage } from '../scripts/test/w7-object-oracle.mjs';
after(() => closeNativeWorkbookWorker());
const references = path.resolve('tests/support/w7-reference-data');
const chartId = '11111111-1111-4111-8111-111111111111',
  pngId = '22222222-2222-4222-8222-222222222222',
  extraId = '55555555-5555-4555-8555-555555555555';
const chart = {
  operation: 'createChart',
  sheet: 'Sales',
  objectId: chartId,
  expectedObjectRevision: null,
  chart: { kind: 'column', range: 'A1:C4', title: 'Monthly sales' },
};
const pattern = [
  '[[]Content_Types].xml',
  'xl/worksheets/sheet1.xml',
  'xl/worksheets/_rels/sheet1.xml.rels',
  'xl/drawings/*',
  'xl/charts/*',
  'xl/media/*',
];
const digest = (bytes) => createHash('sha256').update(bytes).digest('hex');
async function fixture(t, file = path.join(references, 'source.xlsx')) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-objects-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const source = path.join(directory, 'input.xlsx');
  await fs.writeFile(source, objectSourceMetadata(await fs.readFile(file)));
  const original = await fs.readFile(source);
  let root;
  const bridge = async (input) => {
    if (input.operation === 'index') root = input.root;
    return native(input);
  };
  const book = new IndexedWorkbook(
    bridge,
    () => assert.fail('Object authoring must not invoke compatibility writer'),
    (input) => calculateIndexedWorkbook(input, bridge)
  );
  await book.open(source);
  t.after(() => book.close());
  const prepare = async (command, state = '0') =>
    indexedObjectPlan(
      {
        metadataBytes: await fs.readFile(path.join(root, 'metadata.xlsx')),
        view: book.view(state),
        readRange: (sheet, range) => book.read(state, sheet, range),
      },
      command
    );
  const write = async (plan, state = '0') => {
    const exported = await book.export(state);
    await fs.writeFile(path.join(root, 'package-metadata.json'), JSON.stringify(plan));
    const result = await native({
      action: 'workbook',
      operation: 'package-edit',
      root,
      version: 0,
      source: path.basename(exported.path),
    });
    const saved = path.join(directory, 'result-' + Date.now() + '-' + Math.random().toString(16).slice(2) + '.xlsx');
    await fs.copyFile(path.join(root, result.file), saved);
    return saved;
  };
  const list = async (sheet = 'Sales', options) =>
    indexedObjectList(await fs.readFile(path.join(root, 'metadata.xlsx')), book.view('0'), sheet, options);
  const preview = async (id, state = '0') =>
    indexedObjectPreview(
      {
        metadataBytes: await fs.readFile(path.join(root, 'metadata.xlsx')),
        view: book.view(state),
        readRange: (sheet, range) => book.read(state, sheet, range),
        readMedia: async (part, { maxBytes }) => {
          const bytes = unzipSync(await fs.readFile((await book.export(state)).path))[part];
          assert.ok(bytes.length <= maxBytes);
          return bytes;
        },
      },
      'Sales',
      id
    );
  return { directory, source, original, root: () => root, book, prepare, write, list, preview };
}
async function oracle(f, saved, actions, objects = []) {
  const expectation = path.join(f.directory, 'expected-' + Math.random() + '.json');
  await fs.writeFile(
    expectation,
    JSON.stringify({
      actions,
      objects,
      allowChangedParts: pattern,
      preservedParts: [
        'custom/opaque.bin',
        'xl/worksheets/sheet2.xml',
        'xl/workbook.xml',
        'xl/_rels/workbook.xml.rels',
      ],
    })
  );
  const report = checkObjectPackage(f.source, saved, expectation);
  return report;
}
const action = (kind, id) => ({ kind, sheet: 'Sales', id: 'workbook:' + id });
async function expectedChart(kind = 'column', id = chartId) {
  const base = JSON.parse(await fs.readFile(path.join(references, 'expect-created.json'), 'utf8')).objects[0];
  return { ...base, id: 'workbook:' + id, chart: { ...base.chart, kind } };
}
async function insertedPng() {
  return {
    operation: 'insertImage',
    sheet: 'Sales',
    objectId: pngId,
    expectedObjectRevision: null,
    asset: { mime: 'image/png', base64: (await fs.readFile(path.join(references, 'marker.png'))).toString('base64') },
    altText: 'Sales logo',
  };
}

test('actual column chart package contains independent series, formulas, caches, axes and placement', async (t) => {
  const f = await fixture(t),
    prepared = await f.prepare(chart);
  assert.equal(prepared.command.chart.blanks, 'gap');
  assert.equal(prepared.command.chart.includeHidden, false);
  const saved = await f.write(prepared.plan);
  assert.equal((await oracle(f, saved, [action('create', chartId)], [await expectedChart()])).passed, true);
  const reopened = await fixture(t, saved),
    listing = await reopened.list();
  assert.equal(listing.objects[0].canEdit, true);
  assert.equal(listing.objects[0].chart.range, 'A1:C4');
  const preview = await reopened.preview(chartId);
  assert.deepEqual(preview.data.categories, ['Jan', 'Feb', 'Mar']);
  assert.deepEqual(
    preview.data.series.map((item) => item.values),
    [
      [2, 3, 5],
      [1, 2, 3],
    ]
  );
  assert.deepEqual(await f.list(), { objects: [], total: 0, nextCursor: null });
  assert.deepEqual(await fs.readFile(f.source), f.original);
});
test('guarded chart update changes real type/cache/placement and preserves the original accepted object', async (t) => {
  const f = await fixture(t),
    created = await f.write((await f.prepare(chart)).plan),
    g = await fixture(t, created),
    before = (await g.list()).objects[0];
  const prepared = await g.prepare({
    operation: 'updateChart',
    sheet: 'Sales',
    objectId: chartId,
    expectedObjectRevision: before.revision,
    patch: { kind: 'line', title: 'Trend' },
    anchor: { cell: 'F5' },
  });
  const saved = await g.write(prepared.plan),
    expected = await expectedChart('line');
  expected.chart.title = 'Trend';
  expected.anchor.from = { row: 4, col: 5, rowOff: 0, colOff: 0 };
  assert.equal((await oracle(g, saved, [action('update', chartId)], [expected])).passed, true);
  const reopened = await fixture(t, saved),
    after = (await reopened.list()).objects[0];
  assert.equal(after.chart.kind, 'line');
  assert.notEqual(after.revision, before.revision);
  await assert.rejects(
    reopened.prepare({
      operation: 'deleteObject',
      sheet: 'Sales',
      objectId: chartId,
      expectedObjectRevision: before.revision,
    }),
    /changed|protected/
  );
  assert.equal((await g.list()).objects[0].chart.kind, 'column');
  assert.deepEqual(await fs.readFile(g.source), g.original);
});
test('PNG replay payload is bounded and self contained; saved media/relationships cold reopen and final deletion are checked', async (t) => {
  const f = await fixture(t),
    prepared = await f.prepare(await insertedPng());
  assert.equal(prepared.command.asset.bytes, 73);
  assert.equal(prepared.command.asset.widthPx, 2);
  assert.equal(prepared.command.asset.sha256, digest(await fs.readFile(path.join(references, 'marker.png'))));
  assert.ok(Buffer.byteLength(JSON.stringify(prepared.command)) < 200000);
  const saved = await f.write(prepared.plan),
    expected = {
      sheet: 'Sales',
      id: 'workbook:' + pngId,
      kind: 'image',
      protected: [],
      altText: 'Sales logo',
      media: { mime: 'image/png', bytes: 73, sha256: prepared.command.asset.sha256 },
    };
  assert.equal((await oracle(f, saved, [action('create', pngId)], [expected])).passed, true);
  const g = await fixture(t, saved),
    entry = (await g.list()).objects[0],
    preview = await g.preview(pngId);
  assert.equal(preview.image.base64, prepared.command.asset.base64);
  await assert.rejects(
    indexedObjectPreview(
      {
        metadataBytes: await fs.readFile(path.join(g.root(), 'metadata.xlsx')),
        view: g.book.view('0'),
        readMedia: async () => Buffer.alloc(128 * 1024 + 1),
      },
      'Sales',
      pngId
    ),
    /byte bound/
  );
  const removed = await g.write(
    (
      await g.prepare({
        operation: 'deleteObject',
        sheet: 'Sales',
        objectId: pngId,
        expectedObjectRevision: entry.revision,
      })
    ).plan
  );
  assert.equal((await oracle(g, removed, [action('delete', pngId)])).passed, true);
  assert.deepEqual(inventoryObjectPackage(removed).objects, []);
  const parts = unzipSync(await fs.readFile(removed));
  assert.ok(!Object.keys(parts).some((name) => name.startsWith('xl/drawings/') || name.startsWith('xl/media/')));
  assert.ok(!strFromU8(parts['xl/worksheets/sheet1.xml']).includes('<drawing'));
});
test('opaque/unmanaged chart, images and shape bytes survive a scoped creation; paged inventory refuses authoring them', async (t) => {
  const f = await fixture(t, path.join(references, 'opaque-shared.xlsx')),
    before = await f.list();
  assert.equal(before.total, 5);
  assert.ok(before.objects.every((item) => item.protectedReason));
  const first = await f.list('Sales', { limit: 2 }),
    second = await f.list('Sales', { after: first.nextCursor, limit: 2 });
  assert.equal(first.objects.length, 2);
  assert.equal(second.objects.length, 2);
  assert.ok(!second.objects.some((item) => first.objects.some((old) => old.id === item.id)));
  await assert.rejects(f.list('Sales', { after: 'expired' }), { code: 'object-stale' });
  await assert.rejects(
    f.prepare({
      operation: 'deleteObject',
      sheet: 'Sales',
      objectId: chartId,
      expectedObjectRevision: before.objects.find((item) => item.id === chartId).revision,
    }),
    /protected/
  );
  const metadataBytes = await fs.readFile(path.join(f.root(), 'metadata.xlsx'));
  assert.equal(indexedObjectHasManagedCharts(metadataBytes), false);
  assert.deepEqual(
    await indexedObjectCachePlan({
      metadataBytes,
      view: f.book.view('0'),
      readRange: () => assert.fail('Foreign/opaque chart caches must stay untouched'),
    }),
    { replacements: {} }
  );
  const saved = await f.write((await f.prepare({ ...chart, objectId: extraId })).plan);
  assert.equal(
    (await oracle(f, saved, [action('create', extraId)], [await expectedChart('column', extraId)])).passed,
    true
  );
});
test('accepted numeric formulas and blanks qualify; typed errors, unresolved, raw unsafe and special results refuse before a package write', async (t) => {
  const f = await fixture(t),
    metadataBytes = await fs.readFile(path.join(f.root(), 'metadata.xlsx')),
    view = f.book.view('0'),
    base = (await f.book.read('0', 'Sales', 'A1:C4')).cells;
  for (const bad of [
    { value: '#DIV/0!', dataType: 'error' },
    { value: null, formula: 'UNKNOWN(1)', calculation: { status: 'unresolved' } },
    { value: 9007199254740992, dataType: 'number', rawValue: '9007199254740993' },
    { value: 0.1, dataType: 'number', rawValue: '0.10000000000000001' },
    { value: 2, formula: 'ROW(B2:B3)', formulaType: 'array', calculation: { status: 'calculated', value: 2 } },
    { value: true, dataType: 'boolean' },
  ])
    await assert.rejects(
      indexedObjectPlan({ metadataBytes, view, readRange: async () => ({ cells: { ...base, B2: bad } }) }, chart),
      /error|special|unresolved|unsafe|numeric/
    );
  const prepared = await indexedObjectPlan(
    {
      metadataBytes,
      view,
      readRange: async () => ({
        cells: {
          ...base,
          B2: {
            value: null,
            rawValue: '9007199254740993',
            formula: '1+6',
            calculation: { status: 'calculated', value: 7 },
          },
          B3: { value: null, dataType: 'blank' },
        },
      }),
    },
    chart
  );
  assert.match(prepared.plan.replacements[prepared.plan.objectPart], /<c:pt idx="0"><c:v>7<\/c:v>/);
  assert.ok(!/<c:pt idx="1"><c:v>null/.test(prepared.plan.replacements[prepared.plan.objectPart]));
  const typed = await fixture(t, path.join(references, 'typed-refusals.xlsx'));
  await assert.rejects(typed.prepare(chart), /error|unresolved/);
  const special = await fixture(t, path.join(references, 'special-formula.xlsx'));
  await assert.rejects(special.prepare(chart), /special|error/);
  assert.deepEqual(await fs.readFile(f.source), f.original);
});
test('preview and export cache consume the requested accepted version and refuse later source errors without trapping deletion', async (t) => {
  const f = await fixture(t),
    created = await f.write((await f.prepare(chart)).plan),
    g = await fixture(t, created),
    entry = (await g.list()).objects[0],
    changed = await g.book.apply('0', [{ sheet: 'Sales', cell: 'B2', value: 21 }]),
    state = changed.view.storage.state;
  assert.equal((await g.preview(chartId, state)).data.series[0].values[0], 21);
  assert.equal((await g.preview(chartId)).data.series[0].values[0], 2);
  const metadataBytes = await fs.readFile(path.join(g.root(), 'metadata.xlsx')),
    cache = await indexedObjectCachePlan({
      metadataBytes,
      view: g.book.view(state),
      readRange: (sheet, range) => g.book.read(state, sheet, range),
    });
  assert.match(Object.values(cache.replacements)[0], /<c:pt idx="0"><c:v>21<\/c:v>/);
  const failed = await g.book.apply(state, [{ sheet: 'Sales', cell: 'B2', formula: '1/0' }]),
    failedState = failed.view.storage.state;
  await assert.rejects(g.preview(chartId, failedState), /error|unresolved/);
  await assert.rejects(
    indexedObjectCachePlan({
      metadataBytes,
      view: g.book.view(failedState),
      readRange: (sheet, range) => g.book.read(failedState, sheet, range),
    }),
    /error|unresolved/
  );
  assert.equal(
    (
      await g.prepare(
        { operation: 'deleteObject', sheet: 'Sales', objectId: chartId, expectedObjectRevision: entry.revision },
        failedState
      )
    ).plan.kind,
    'deleteObject'
  );
  assert.match(
    indexedObjectReferenceGuard(metadataBytes, g.book.view(state), { operation: 'renameSheet', sheet: 'Sales' }),
    /not qualified/
  );
  assert.match(
    indexedObjectReferenceGuard(metadataBytes, g.book.view(state), { operation: 'duplicateSheet', sheet: 'Sales' }),
    /independent ownership/
  );
});
test('command scope, source bounds, protected graph and binary metadata are checked before touching source', async (t) => {
  const f = await fixture(t);
  for (const command of [
    { ...chart, expectedObjectRevision: undefined },
    { ...chart, sheet: 'Other', objectId: 'arbitrary' },
    { ...chart, path: '/tmp/file' },
    { ...chart, chart: { ...chart.chart, range: 'A1:C1000' } },
    { ...chart, chart: { ...chart.chart, kind: 'pie' } },
  ])
    await assert.rejects(f.prepare(command), /guard|identity|fields|command|1000|2D/);
  const png = await insertedPng();
  for (const asset of [
    { ...png.asset, mime: 'image/jpeg' },
    { ...png.asset, widthPx: 200 },
    { ...png.asset, sha256: '0'.repeat(64) },
    { ...png.asset, base64: png.asset.base64 + ' ' },
    { ...png.asset, base64: 'a'.repeat(200000) },
  ])
    assert.throws(() => indexedPngPayload(asset), /PNG/);
  for (const name of ['external-chart', 'extended-chart', 'pivot-chart', 'unknown-dependent', 'shared-mutable-chart']) {
    const g = await fixture(t, path.join(references, name + '.xlsx')),
      entry = (await g.list()).objects.find((item) => item.id === chartId);
    assert.ok(entry.protectedReason);
    await assert.rejects(
      g.prepare({
        operation: 'deleteObject',
        sheet: 'Sales',
        objectId: chartId,
        expectedObjectRevision: entry.revision,
      }),
      /protected/
    );
    assert.deepEqual(await fs.readFile(g.source), g.original);
  }
});
test('native rejects malformed additions, unrelated XML, graph orphans and duplicate claims atomically', async (t) => {
  const f = await fixture(t),
    prepared = await f.prepare(await insertedPng()),
    source = await f.book.export('0'),
    output = path.join(f.root(), 'package.xlsx'),
    sentinel = Buffer.from('previous complete output');
  await fs.writeFile(output, sentinel);
  const badPlans = [];
  const clone = () => structuredClone(prepared.plan),
    part = prepared.plan.objectPart;
  let p = clone();
  p.binaryAdditions[part].sha256 = '0'.repeat(64);
  badPlans.push(p);
  p = clone();
  p.binaryAdditions[part].widthPx = 3;
  badPlans.push(p);
  p = clone();
  p.binaryAdditions[part].base64 += ' ';
  badPlans.push(p);
  p = clone();
  p.replacements['xl/workbook.xml'] = '<oops/>';
  badPlans.push(p);
  p = clone();
  p.removed = ['custom/opaque.bin'];
  badPlans.push(p);
  p = clone();
  p.replacements[part] = 'not XML';
  badPlans.push(p);
  p = clone();
  p.replacements[p.drawingPart] = p.replacements[p.drawingPart].replace(
    'r:embed="workbookObject_',
    'r:embed="missing_'
  );
  badPlans.push(p);
  p = clone();
  p.replacements['[Content_Types].xml'] = p.replacements['[Content_Types].xml'].replace('image/png', 'image/jpeg');
  badPlans.push(p);
  p = clone();
  p.replacements[p.drawingPart] += '<extra/>';
  badPlans.push(p);
  p = clone();
  p.priorObjectRevision = '0'.repeat(64);
  badPlans.push(p);
  for (const plan of badPlans) {
    await fs.writeFile(path.join(f.root(), 'package-metadata.json'), JSON.stringify(plan));
    await assert.rejects(
      native({
        action: 'workbook',
        operation: 'package-edit',
        root: f.root(),
        version: 0,
        source: path.basename(source.path),
      }),
      /PNG|Object|object|XML|part|graph|ownership|claim|content type|revision|relationship|root/
    );
    assert.deepEqual(await fs.readFile(output), sentinel);
    await assert.rejects(fs.stat(path.join(f.root(), 'package.partial')), (error) => error.code === 'ENOENT');
  }
  assert.deepEqual(await fs.readFile(f.source), f.original);
});
test('native refuses unscoped opaque anchor edits and duplicate object IDs while preserving its previous complete output', async (t) => {
  const f = await fixture(t, path.join(references, 'opaque-shared.xlsx')),
    prepared = await f.prepare({ ...chart, objectId: extraId }),
    source = await f.book.export('0'),
    output = path.join(f.root(), 'package.xlsx'),
    sentinel = Buffer.from('previous complete output');
  await fs.writeFile(output, sentinel);
  for (const change of [
    (value) => value.replace('Opaque protected shape', 'Tampered protected shape'),
    (value) => value.replace('name="Protected shape"', 'name="workbook:' + extraId + '"'),
  ]) {
    const plan = structuredClone(prepared.plan);
    plan.replacements[plan.drawingPart] = change(plan.replacements[plan.drawingPart]);
    await fs.writeFile(path.join(f.root(), 'package-metadata.json'), JSON.stringify(plan));
    await assert.rejects(
      native({
        action: 'workbook',
        operation: 'package-edit',
        root: f.root(),
        version: 0,
        source: path.basename(source.path),
      }),
      /drawing bytes|Duplicate|ambiguous/
    );
    assert.deepEqual(await fs.readFile(output), sentinel);
  }
});
test('cancel marker and a real bounded-writer quota failure leave prior complete output and source intact', async (t) => {
  const base = unzipSync(await fs.readFile(path.join(references, 'source.xlsx')));
  base['custom/opaque.bin'] = randomBytes(220000);
  const input = path.join(os.tmpdir(), 'w7-quota-' + Date.now() + '.xlsx');
  await fs.writeFile(input, zipSync(base));
  t.after(() => fs.rm(input, { force: true }));
  const f = await fixture(t, input),
    prepared = await f.prepare(chart),
    source = await f.book.export('0'),
    output = path.join(f.root(), 'package.xlsx'),
    sentinel = Buffer.from('previous complete output');
  await fs.writeFile(output, sentinel);
  await fs.writeFile(path.join(f.root(), 'package-metadata.json'), JSON.stringify(prepared.plan));
  const request = {
    action: 'workbook',
    operation: 'package-edit',
    root: f.root(),
    version: 0,
    source: path.basename(source.path),
  };
  await fs.writeFile(path.join(f.root(), 'cancelled'), '1');
  await assert.rejects(native(request), { code: 'operation-cancelled' });
  await fs.rm(path.join(f.root(), 'cancelled'));
  assert.deepEqual(await fs.readFile(output), sentinel);
  async function used(directory) {
    let total = 0;
    for (const entry of await fs.readdir(directory, { withFileTypes: true })) {
      const file = path.join(directory, entry.name);
      total += entry.isDirectory() ? await used(file) : (await fs.stat(file)).size;
    }
    return total;
  }
  const session = path.dirname(f.root()),
    ownerPath = path.join(session, 'owner.json'),
    owner = JSON.parse(await fs.readFile(ownerPath, 'utf8')),
    dbBytes = (await fs.stat(path.join(f.root(), 'cells.sqlite'))).size;
  owner.maxSessionBytes = (await used(session)) + dbBytes * 2 + 6 * 17000;
  await fs.writeFile(ownerPath, JSON.stringify(owner));
  await assert.rejects(native(request), /budget|quota/);
  assert.deepEqual(await fs.readFile(output), sentinel);
  await assert.rejects(fs.stat(path.join(f.root(), 'package.partial')), (error) => error.code === 'ENOENT');
  assert.deepEqual(await fs.readFile(f.source), f.original);
});

test('direct line creation and chart deletion preserve the independently authored PNG sibling', async (t) => {
  const f = await fixture(t),
    line = await f.write((await f.prepare({ ...chart, chart: { ...chart.chart, kind: 'line' } })).plan);
  assert.equal((await oracle(f, line, [action('create', chartId)], [await expectedChart('line')])).passed, true);
  const g = await fixture(t, line),
    both = await g.write((await g.prepare(await insertedPng())).plan),
    h = await fixture(t, both),
    entry = (await h.list()).objects.find((item) => item.id === chartId);
  const removed = await h.write(
    (
      await h.prepare({
        operation: 'deleteObject',
        sheet: 'Sales',
        objectId: chartId,
        expectedObjectRevision: entry.revision,
      })
    ).plan
  );
  assert.equal((await oracle(h, removed, [action('delete', chartId)])).passed, true);
  const reopened = await fixture(t, removed);
  assert.deepEqual(
    (await reopened.list()).objects.map((item) => item.kind),
    ['image']
  );
  assert.equal((await reopened.preview(pngId)).image.bytes, 73);
});
test('deleting either of two managed PNG anchors keeps shared bytes until the final graph reference is deleted', async (t) => {
  const f = await fixture(t),
    one = await f.write((await f.prepare(await insertedPng())).plan),
    parts = unzipSync(await fs.readFile(one)),
    drawing = Object.keys(parts).find((name) => name.startsWith('xl/drawings/') && !name.endsWith('.rels')),
    rels = drawing.replace('/drawings/', '/drawings/_rels/') + '.rels',
    raw = strFromU8(parts[drawing]),
    anchor = /<xdr:oneCellAnchor[\s\S]*?<\/xdr:oneCellAnchor>/.exec(raw)[0],
    oldRel = 'workbookObject_' + pngId.replaceAll('-', ''),
    newRel = 'workbookObject_' + extraId.replaceAll('-', '');
  const duplicate = anchor
    .replace('id="1"', 'id="2"')
    .replace('workbook:' + pngId, 'workbook:' + extraId)
    .replace(oldRel, newRel);
  parts[drawing] = strToU8(raw.replace('</xdr:wsDr>', duplicate + '</xdr:wsDr>'));
  const edge = /<Relationship[^>]+\/>/.exec(strFromU8(parts[rels]))[0];
  parts[rels] = strToU8(
    strFromU8(parts[rels]).replace('</Relationships>', edge.replace(oldRel, newRel) + '</Relationships>')
  );
  const shared = path.join(f.directory, 'shared.xlsx');
  await fs.writeFile(shared, zipSync(parts));
  const g = await fixture(t, shared),
    first = (await g.list()).objects.find((item) => item.id === pngId);
  assert.equal(first.canDelete, true);
  const removed = await g.write(
    (
      await g.prepare({
        operation: 'deleteObject',
        sheet: 'Sales',
        objectId: pngId,
        expectedObjectRevision: first.revision,
      })
    ).plan
  );
  assert.equal((await oracle(g, removed, [action('delete', pngId)])).passed, true);
  assert.ok(Object.keys(unzipSync(await fs.readFile(removed))).some((name) => name.startsWith('xl/media/')));
  const h = await fixture(t, removed),
    last = (await h.list()).objects[0];
  assert.equal((await h.preview(extraId)).image.bytes, 73);
  const empty = await h.write(
    (
      await h.prepare({
        operation: 'deleteObject',
        sheet: 'Sales',
        objectId: extraId,
        expectedObjectRevision: last.revision,
      })
    ).plan
  );
  assert.equal((await oracle(h, empty, [action('delete', extraId)])).passed, true);
  assert.ok(!Object.keys(unzipSync(await fs.readFile(empty))).some((name) => name.startsWith('xl/media/')));
});

test('cancel after package.partial exists interrupts the streaming writer and removes its partial package', async (t) => {
  const parts = unzipSync(await fs.readFile(path.join(references, 'source.xlsx')));
  parts['custom/opaque.bin'] = randomBytes(20 * 1024 * 1024);
  const file = path.join(os.tmpdir(), 'w7-cancel-stream-' + Date.now() + '.xlsx');
  await fs.writeFile(file, zipSync(parts, { level: 0 }));
  t.after(() => fs.rm(file, { force: true }));
  const f = await fixture(t, file),
    prepared = await f.prepare(chart),
    source = await f.book.export('0'),
    partial = path.join(f.root(), 'package.partial'),
    output = path.join(f.root(), 'package.xlsx'),
    sentinel = Buffer.from('previous complete output');
  await fs.writeFile(output, sentinel);
  await fs.writeFile(path.join(f.root(), 'package-metadata.json'), JSON.stringify(prepared.plan));
  let finished = false,
    wroteMarker = false;
  const writing = native({
    action: 'workbook',
    operation: 'package-edit',
    root: f.root(),
    version: 0,
    source: path.basename(source.path),
  }).finally(() => {
    finished = true;
  });
  const interrupted = (async () => {
    const deadline = Date.now() + 5000;
    while (!finished && Date.now() < deadline) {
      try {
        await fs.stat(partial);
        await fs.writeFile(path.join(f.root(), 'cancelled'), '1');
        wroteMarker = true;
        return;
      } catch (error) {
        if (error.code !== 'ENOENT') throw error;
      }
      await new Promise((resolve) => setTimeout(resolve, 1));
    }
  })();
  await assert.rejects(writing, /cancelled/);
  await interrupted;
  assert.equal(wroteMarker, true);
  assert.deepEqual(await fs.readFile(output), sentinel);
  await assert.rejects(fs.stat(partial), (error) => error.code === 'ENOENT');
  assert.deepEqual(await fs.readFile(f.source), f.original);
});

test('extended managed anchors and broken primary axes stay protected and retain chart referential refusals', async (t) => {
  const f = await fixture(t),
    created = await f.write((await f.prepare(chart)).plan),
    parts = unzipSync(await fs.readFile(created)),
    drawing = Object.keys(parts).find((name) => name.startsWith('xl/drawings/') && !name.endsWith('.rels')),
    chartPart = Object.keys(parts).find((name) => name.startsWith('xl/charts/'));
  for (const kind of ['anchor-attribute', 'broken-axis']) {
    const next = { ...parts };
    if (kind === 'anchor-attribute')
      next[drawing] = strToU8(
        strFromU8(parts[drawing]).replace('<xdr:graphicFrame>', '<xdr:graphicFrame macro="unqualified">')
      );
    else
      next[chartPart] = strToU8(strFromU8(parts[chartPart]).replace('<c:crossAx val="20"/>', '<c:crossAx val="999"/>'));
    const file = path.join(f.directory, kind + '.xlsx');
    await fs.writeFile(file, zipSync(next));
    const g = await fixture(t, file),
      entry = (await g.list()).objects[0];
    assert.equal(entry.kind, 'chart');
    assert.equal(entry.canEdit, false);
    assert.ok(entry.protectedReason);
    await assert.rejects(
      g.prepare({
        operation: 'deleteObject',
        sheet: 'Sales',
        objectId: chartId,
        expectedObjectRevision: entry.revision,
      }),
      /protected/
    );
    assert.match(
      indexedObjectReferenceGuard(await fs.readFile(path.join(g.root(), 'metadata.xlsx')), g.book.view('0'), {
        operation: 'renameSheet',
        sheet: 'Sales',
      }),
      /not qualified/
    );
    assert.deepEqual(await fs.readFile(g.source), g.original);
  }
});
