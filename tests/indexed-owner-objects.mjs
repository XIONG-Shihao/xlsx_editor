import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { unzipSync, zipSync, strToU8 } from 'fflate';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { inventoryObjectPackage, checkObjectPackage } from '../scripts/test/w7-object-oracle.mjs';
after(() => closeNativeWorkbookWorker());
const references = path.resolve('tests/support/w7-reference-data'),
  id = '11111111-1111-4111-8111-111111111111',
  imageId = '22222222-2222-4222-8222-222222222222';
const create = {
  operation: 'createChart',
  sheet: 'Sales',
  objectId: id,
  expectedObjectRevision: null,
  chart: { kind: 'column', range: 'A1:C4', title: 'Monthly sales' },
};
async function fixture(t, file = path.join(references, 'source.xlsx'), transport = native) {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-owner-object-'));
  t.after(() => fs.rm(dir, { recursive: true, force: true }));
  const source = path.join(dir, 'source.xlsx');
  await fs.copyFile(file, source);
  const original = await fs.readFile(source),
    book = new IndexedWorkbook(
      transport,
      () => assert.fail('No compatibility writer'),
      (input) => calculateIndexedWorkbook(input, transport)
    );
  await book.open(source);
  t.after(() => book.close());
  return { book, dir, source, original };
}
test('owner publishes one chart state, normalizes durable commands and keeps undo/redo versions isolated', async (t) => {
  const f = await fixture(t),
    created = await f.book.apply('0', [create]),
    state = created.view.storage.state;
  assert.equal(f.book.diagnosticStats().retainedStates, 2);
  assert.equal(created.commands[0].chart.blanks, 'gap');
  assert.equal(created.commands[0].anchor.cell, 'E2');
  assert.deepEqual(created.commands, created.edits);
  assert.equal((await f.book.objects('0', 'Sales')).total, 0);
  const entry = (await f.book.objects(state, 'Sales')).objects[0],
    changed = await f.book.apply(state, [{ sheet: 'Sales', cell: 'B2', value: 21 }]),
    later = changed.view.storage.state;
  assert.equal((await f.book.object(later, 'Sales', id)).data.series[0].values[0], 21);
  assert.equal((await f.book.object(state, 'Sales', id)).data.series[0].values[0], 2);
  const saved = await f.book.export(later);
  assert.equal(inventoryObjectPackage(saved.path).objects[0].chart.series[0].values.cache[0], 21);
  const redo = await f.book.apply(later, [
    {
      operation: 'updateChart',
      sheet: 'Sales',
      objectId: id,
      expectedObjectRevision: entry.revision,
      patch: { kind: 'line' },
    },
  ]);
  assert.equal((await f.book.object(redo.view.storage.state, 'Sales', id)).chart.kind, 'line');
  assert.equal((await f.book.object(state, 'Sales', id)).chart.kind, 'column');
  await assert.rejects(
    f.book.apply(state, [{ operation: 'renameSheet', sheet: 'Sales', name: 'Renamed' }]),
    /not qualified/
  );
  await assert.rejects(
    f.book.apply(state, [{ operation: 'duplicateSheet', sheet: 'Sales', name: 'Copy' }]),
    /independent ownership/
  );
  assert.deepEqual(await fs.readFile(f.source), f.original);
});
test('self-contained PNG normalization replays identically and scoped owner media never returns package paths', async (t) => {
  const f = await fixture(t),
    png = (await fs.readFile(path.join(references, 'marker.png'))).toString('base64'),
    accepted = await f.book.apply('0', [
      {
        operation: 'insertImage',
        sheet: 'Sales',
        objectId: imageId,
        expectedObjectRevision: null,
        asset: { mime: 'image/png', base64: png },
        altText: 'Logo',
      },
    ]),
    state = accepted.view.storage.state;
  assert.equal(accepted.commands[0].asset.bytes, 73);
  const preview = await f.book.object(state, 'Sales', imageId);
  assert.equal(preview.image.base64, png);
  assert.ok(!JSON.stringify(preview).includes('xl/'));
  assert.ok(!JSON.stringify(await f.book.objects(state, 'Sales')).includes(f.dir));
  await assert.rejects(f.book.object(state, 'Other', imageId), /not granted/);
  const replay = await fixture(t),
    restored = await replay.book.apply('0', accepted.commands);
  assert.deepEqual(
    await replay.book.objects(restored.view.storage.state, 'Sales'),
    await f.book.objects(state, 'Sales')
  );
  const old = (await f.book.objects(state, 'Sales')).objects[0],
    removed = await f.book.apply(state, [
      { operation: 'deleteObject', sheet: 'Sales', objectId: imageId, expectedObjectRevision: old.revision },
    ]);
  assert.equal((await f.book.objects(removed.view.storage.state, 'Sales')).total, 0);
  assert.equal((await f.book.objects(state, 'Sales')).total, 1);
});
test('loaded formula chart refreshes accepted caches before export fast paths and can be deleted after source errors', async (t) => {
  const f = await fixture(t),
    created = await f.book.apply('0', [create]),
    first = await f.book.export(created.view.storage.state),
    parts = unzipSync(await fs.readFile(first.path));
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    new TextDecoder()
      .decode(parts['xl/worksheets/sheet1.xml'])
      .replace('<c r="B2"><v>2</v></c>', '<c r="B2"><f>7+4</f><v>9007199254740993</v></c>')
  );
  const input = path.join(f.dir, 'formula.xlsx');
  await fs.writeFile(input, zipSync(parts));
  const g = await fixture(t, input),
    fresh = await g.book.export('0');
  assert.equal(inventoryObjectPackage(fresh.path).objects[0].chart.series[0].values.cache[0], 11);
  const source = await g.book.read('0', 'Sales', 'B2');
  assert.equal(source.cells.B2.calculation.value, 11);
  const entry = (await g.book.objects('0', 'Sales')).objects[0],
    bad = await g.book.apply('0', [{ sheet: 'Sales', cell: 'B2', formula: '1/0' }]),
    state = bad.view.storage.state;
  await assert.rejects(g.book.export(state), /error|unresolved/);
  const deleted = await g.book.apply(state, [
    { operation: 'deleteObject', sheet: 'Sales', objectId: id, expectedObjectRevision: entry.revision },
  ]);
  assert.equal((await g.book.objects(deleted.view.storage.state, 'Sales')).total, 0);
  await g.book.export(deleted.view.storage.state);
  assert.equal((await g.book.objects('0', 'Sales')).total, 1);
});
test('mixed object/cell batches publish only one state and rollback all candidate roots on a later failure', async (t) => {
  const f = await fixture(t),
    before = f.book.diagnosticStats(),
    failed = [create, { sheet: 'Sales', cell: 'XFE1', value: 5 }];
  await assert.rejects(f.book.apply('0', failed), /limits|address/);
  assert.equal((await f.book.objects('0', 'Sales')).total, 0);
  assert.equal(f.book.diagnosticStats().storageRoots, before.storageRoots);
  assert.equal(f.book.diagnosticStats().retainedStates, before.retainedStates);
  const ok = await f.book.apply('0', [create, { sheet: 'Sales', cell: 'B2', value: 10 }]);
  assert.equal(f.book.diagnosticStats().retainedStates, 2);
  assert.equal(ok.commands.length, 2);
  assert.equal(ok.commands[0].chart.blanks, 'gap');
  assert.equal((await f.book.object(ok.view.storage.state, 'Sales', id)).data.series[0].values[0], 10);
});
test('late mixed guard failure prunes private same-root SQL versions while preserving old and redo states', async (t) => {
  let root;
  const patches = [];
  const bridge = async (request) => {
    const result = await native(request);
    if (request.operation === 'index') root = request.root;
    if (request.operation === 'patch') patches.push(result.version);
    return result;
  };
  const f = await fixture(t, undefined, bridge),
    first = await f.book.apply('0', [{ sheet: 'Sales', cell: 'B2', value: 9 }]),
    before = first.view.storage.state,
    redo = await f.book.apply('0', [{ sheet: 'Sales', cell: 'C2', value: 12 }]),
    later = redo.view.storage.state;
  const previous = patches.length;
  await assert.rejects(
    f.book.apply(before, [
      { sheet: 'Sales', cell: 'B2', value: 17 },
      { operation: 'deleteObject', sheet: 'Sales', objectId: id, expectedObjectRevision: 'a'.repeat(64) },
    ]),
    /changed|protected/
  );
  assert.equal(patches.length, previous + 1);
  await assert.rejects(
    native({
      action: 'workbook',
      operation: 'range',
      root,
      version: patches.at(-1),
      sheet: 'xl/worksheets/sheet1.xml',
      range: 'B2',
      offset: 0,
      limit: 1,
    }),
    { code: 'internal-unknown-reference' }
  );
  assert.equal((await f.book.read('0', 'Sales', 'B2')).cells.B2.value, 2);
  assert.equal((await f.book.read(before, 'Sales', 'B2')).cells.B2.value, 9);
  assert.equal((await f.book.read(later, 'Sales', 'C2')).cells.C2.value, 12);
  assert.equal(f.book.diagnosticStats().retainedStates, 3);
});
test('cache refresh preserves raw non-cache comments and namespace prefixes in a loaded managed chart', async (t) => {
  const f = await fixture(t),
    created = await f.book.apply('0', [create]),
    saved = await f.book.export(created.view.storage.state),
    parts = unzipSync(await fs.readFile(saved.path)),
    part = Object.keys(parts).find((name) => name.startsWith('xl/charts/workbook-chart-'));
  const original = new TextDecoder()
    .decode(parts[part])
    .replaceAll('c:', 'chart:')
    .replace('xmlns:c=', 'xmlns:chart=')
    .replace('<chart:plotArea>', '<!-- retained chart comment -->\n<chart:plotArea>');
  parts[part] = strToU8(original);
  const input = path.join(f.dir, 'commented-chart.xlsx');
  await fs.writeFile(input, zipSync(parts));
  const g = await fixture(t, input),
    changed = await g.book.apply('0', [{ sheet: 'Sales', cell: 'B2', value: 23 }]),
    output = await g.book.export(changed.view.storage.state),
    result = new TextDecoder().decode(unzipSync(await fs.readFile(output.path))[part]);
  const envelope = (xml) => xml.replace(/<chart:(strCache|numCache)>[\s\S]*?<\/chart:\1>/g, '<cache/>');
  assert.equal(envelope(result), envelope(original));
  assert.match(result, /<chart:v>23<\/chart:v>/);
  assert.equal((await g.book.object(changed.view.storage.state, 'Sales', id)).data.series[0].values[0], 23);
});
test('literal signed zero survives accepted chart preview and native saved cache; rounded raw sources refuse atomically', async (t) => {
  const base = unzipSync(await fs.readFile(path.join(references, 'source.xlsx'))),
    directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-chart-numeric-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  for (const raw of ['-0', '0.10000000000000001', '9007199254740992']) {
    const parts = { ...base };
    parts['xl/worksheets/sheet1.xml'] = strToU8(
      new TextDecoder()
        .decode(base['xl/worksheets/sheet1.xml'])
        .replace('<c r="B2"><v>2</v></c>', `<c r="B2"><v>${raw}</v></c>`)
    );
    const input = path.join(directory, raw + '.xlsx');
    await fs.writeFile(input, zipSync(parts));
    const f = await fixture(t, input);
    if (raw !== '-0') {
      await assert.rejects(f.book.apply('0', [create]), /unsafe|precision/);
      assert.equal(f.book.diagnosticStats().retainedStates, 1);
      assert.equal((await f.book.objects('0', 'Sales')).total, 0);
      continue;
    }
    const made = await f.book.apply('0', [create]),
      state = made.view.storage.state;
    assert.ok(Object.is((await f.book.object(state, 'Sales', id)).data.series[0].values[0], -0));
    const saved = await f.book.export(state),
      chart = inventoryObjectPackage(saved.path).objects[0];
    assert.ok(Object.is(chart.chart.series[0].values.cache[0], -0));
  }
});
