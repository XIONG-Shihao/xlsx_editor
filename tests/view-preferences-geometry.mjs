import './support/scratch-dir.mjs';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createHash } from 'node:crypto';
import { build } from 'esbuild';

const component = fileURLToPath(new URL('../', import.meta.url));
const output = path.resolve(
  component,
  '.build/tests/view-model'
);
await fs.mkdir(output, { recursive: true });
const modules = {};
for (const name of ['viewPreferences', 'viewGeometry']) {
  const file = path.join(output, name + '.mjs');
  await build({
    entryPoints: [path.join(component, 'ui', name + '.ts')],
    bundle: true,
    format: 'esm',
    target: 'es2022',
    outfile: file,
    logLevel: 'silent',
  });
  modules[name] = await import(pathToFileURL(file));
}
const { createViewPreferenceStore, DEFAULT_VIEW_PREFERENCES, VIEW_PREFERENCES_KEY, validateViewPreferencesPatch } =
  modules.viewPreferences;
// Retain the old roomy delimited profile's geometry assertions explicitly.
// Source-backed XLSX/default geometry is tested in source-geometry-ui.mjs.
const literalGeometry = {
  ...modules.viewGeometry,
  createViewGeometry: (sheet, options) => modules.viewGeometry.createViewGeometry(sheet, { ...options, literal: true }),
};
for (const name of ['rowHeight', 'columnWidth', 'rowOffsets', 'columnOffsets', 'offsetForRow', 'offsetForColumn'])
  literalGeometry[name] = (sheet, index, zoom) => modules.viewGeometry[name](sheet, index, zoom, { literal: true });
const {
  createViewGeometry,
  rowHeight,
  columnWidth,
  rowOffsets,
  columnOffsets,
  offsetForRow,
  offsetForColumn,
  virtualWindow,
  hitTestCell,
  mergedCellBounds,
  cellRectangle,
  revealCellScroll,
  zoomToSelection,
} = literalGeometry;
function storage(initial = null) {
  let value = initial,
    writes = 0,
    blocked = false;
  return {
    getItem: () => value,
    setItem: (_key, next) => {
      if (blocked) throw Error('Quota');
      value = next;
      writes++;
    },
    get value() {
      return value;
    },
    get writes() {
      return writes;
    },
    set blocked(next) {
      blocked = next;
    },
  };
}
const geometrySheet = {
  rowHeights: [
    { row: 0, height: 18 },
    { row: 1, height: 30 },
    { row: 2, height: 60 },
    { row: 4, height: 0 },
  ],
  hiddenRows: [2],
  widths: [
    { min: 1, max: 1, width: 5, hidden: false },
    { min: 2, max: 2, width: 12, hidden: true },
    { min: 3, max: 3, width: 20, hidden: false },
    { min: 5, max: 5, width: 8, hidden: false },
    { min: 5, max: 6, width: 30, hidden: false },
  ],
  freeze: { rows: 2, columns: 2 },
  merges: ['C4:D6'],
};
const close = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-8, actual + ' != ' + expected);

test('preferences have closed scalar validation, isolated returned values and no workbook/history mutation', () => {
  const profile = storage(),
    store = createViewPreferenceStore(profile),
    workbook = {
      state: 'accepted-17',
      sourceSha256: 'untouched',
      canUndo: true,
      canRedo: true,
      cells: { A1: { formula: '1+2', value: 3, calculation: { status: 'calculated' } } },
    },
    before = createHash('sha256').update(JSON.stringify(workbook)).digest('hex');
  assert.deepEqual(store.get('file:A', 'Sheet1'), DEFAULT_VIEW_PREFERENCES);
  const result = store.patch('file:A', 'Sheet1', { zoom: 125, headings: false, showFormulas: true });
  result.zoom = 400;
  assert.equal(store.get('file:A', 'Sheet1').zoom, 125);
  const saved = profile.value,
    writes = profile.writes;
  for (const patch of [
    { zoom: 9 },
    { zoom: 401 },
    { zoom: 125.5 },
    { zoom: NaN },
    { zoom: '125' },
    { gridlines: 1 },
    { formulaBar: null },
    { extra: true },
    [],
    null,
    JSON.parse('{"__proto__":true}'),
  ])
    assert.throws(() => store.patch('file:A', 'Sheet1', patch));
  assert.equal(profile.value, saved);
  assert.equal(profile.writes, writes);
  assert.equal(store.get('file:A', 'Sheet1').showFormulas, true);
  assert.equal(createHash('sha256').update(JSON.stringify(workbook)).digest('hex'), before);
  assert.deepEqual(validateViewPreferencesPatch({ zoom: 10, gridlines: false }), { zoom: 10, gridlines: false });
  assert.deepEqual(validateViewPreferencesPatch({ zoom: 400 }), { zoom: 400 });
});
test('two live store instances merge documents, sheets and individual fields before every write', () => {
  const profile = storage(),
    a = createViewPreferenceStore(profile),
    b = createViewPreferenceStore(profile);
  a.patch('file:A', 'Sheet1', { zoom: 125 });
  b.patch('file:B', 'Sheet1', { headings: false });
  a.patch('file:A', 'Sheet2', { zoom: 200 });
  b.patch('file:A', 'Sheet1', { showFormulas: true });
  const reopened = createViewPreferenceStore(profile);
  assert.deepEqual(reopened.get('file:A', 'Sheet1'), { ...DEFAULT_VIEW_PREFERENCES, zoom: 125, showFormulas: true });
  assert.equal(reopened.get('file:A', 'Sheet2').zoom, 200);
  assert.equal(reopened.get('file:B', 'Sheet1').headings, false);
  assert.deepEqual(
    reopened.get('file:A:different-revision', 'Sheet1'),
    DEFAULT_VIEW_PREFERENCES,
    'Caller must supply stable document ID, not a revision-keyed ID'
  );
  assert.equal(a.removeDocument('file:A'), true);
  assert.equal(a.removeDocument('file:A'), false);
  b.patch('file:B', 'Sheet2', { zoom: 40 });
  const final = createViewPreferenceStore(profile);
  assert.deepEqual(final.get('file:A', 'Sheet1'), DEFAULT_VIEW_PREFERENCES);
  assert.equal(final.get('file:B', 'Sheet1').headings, false);
  assert.equal(final.get('file:B', 'Sheet2').zoom, 40);
});
test('blocked profile writes preserve local choice then merge failed fields with another store on retry', () => {
  const profile = storage(),
    a = createViewPreferenceStore(profile),
    b = createViewPreferenceStore(profile);
  profile.blocked = true;
  a.patch('file:A', 'Sheet1', { zoom: 200 });
  assert.equal(a.get('file:A', 'Sheet1').zoom, 200);
  assert.match(a.warning, /could not be saved/);
  assert.equal(profile.writes, 0);
  profile.blocked = false;
  b.patch('file:A', 'Sheet1', { headings: false });
  b.patch('file:B', 'Sheet1', { formulaBar: false });
  a.patch('file:A', 'Sheet1', { showFormulas: true });
  const reopened = createViewPreferenceStore(profile);
  assert.deepEqual(reopened.get('file:A', 'Sheet1'), {
    ...DEFAULT_VIEW_PREFERENCES,
    zoom: 200,
    headings: false,
    showFormulas: true,
  });
  assert.equal(reopened.get('file:B', 'Sheet1').formulaBar, false);
  assert.equal(a.warning, null);
});
test('first-Save preference migration refreshes, moves all sheets and resolves collisions in one profile write', () => {
  const profile = storage(),
    owner = createViewPreferenceStore(profile),
    other = createViewPreferenceStore(profile),
    workbook = { state: 'accepted-new', sourceSha256: 'unchanged', canUndo: true, cells: { A1: { value: 25 } } },
    source = JSON.stringify(workbook);
  owner.patch('new:uuid', 'One', { zoom: 125, headings: false });
  owner.patch('new:uuid', 'Two', { zoom: 40, showFormulas: true });
  owner.patch('file:canonical', 'One', { zoom: 400, formulaBar: false });
  owner.patch('file:canonical', 'Only target', { gridlines: false });
  other.patch('new:uuid', 'One', { gridlines: false });
  other.patch('foreign', 'One', { zoom: 200 });
  const writes = profile.writes;
  assert.equal(owner.moveDocument('new:uuid', 'file:canonical'), true);
  assert.equal(profile.writes, writes + 1, 'Migration commits the whole profile exactly once');
  const cold = createViewPreferenceStore(profile);
  assert.deepEqual(cold.get('file:canonical', 'One'), {
    ...DEFAULT_VIEW_PREFERENCES,
    zoom: 125,
    headings: false,
    gridlines: false,
  });
  assert.deepEqual(cold.get('file:canonical', 'Two'), { ...DEFAULT_VIEW_PREFERENCES, zoom: 40, showFormulas: true });
  assert.equal(cold.get('file:canonical', 'Only target').gridlines, false);
  assert.equal(cold.get('foreign', 'One').zoom, 200);
  assert.deepEqual(cold.get('new:uuid', 'One'), DEFAULT_VIEW_PREFERENCES);
  assert.ok(JSON.parse(profile.value).entries.every((entry) => entry.documentId !== 'new:uuid'));
  const saved = profile.value;
  assert.equal(owner.moveDocument('file:canonical', 'file:canonical'), false);
  assert.equal(owner.moveDocument('absent', 'file:canonical'), false);
  for (const pair of [
    ['', 'target'],
    ['from', ' '],
    ['bad\0id', 'target'],
    ['from', 'x'.repeat(513)],
    [undefined, 'target'],
  ])
    assert.throws(() => owner.moveDocument(...pair));
  assert.equal(profile.value, saved);
  assert.equal(profile.writes, writes + 1);
  assert.equal(JSON.stringify(workbook), source);
});
test('failed migration retains local sheets and source tombstone then recovers atomically alongside foreign stores', () => {
  const profile = storage(),
    owner = createViewPreferenceStore(profile),
    other = createViewPreferenceStore(profile);
  owner.patch('new:uuid', 'One', { zoom: 125, formulaBar: false });
  owner.patch('new:uuid', 'Two', { zoom: 200, showFormulas: true });
  owner.patch('file:canonical', 'One', { zoom: 400, gridlines: false });
  const saved = profile.value,
    writes = profile.writes;
  profile.blocked = true;
  assert.equal(owner.moveDocument('new:uuid', 'file:canonical'), true);
  assert.equal(profile.value, saved);
  assert.equal(profile.writes, writes);
  assert.match(owner.warning, /could not be saved/);
  assert.deepEqual(owner.get('new:uuid', 'One'), DEFAULT_VIEW_PREFERENCES);
  assert.deepEqual(owner.get('file:canonical', 'One'), { ...DEFAULT_VIEW_PREFERENCES, zoom: 125, formulaBar: false });
  assert.equal(owner.get('file:canonical', 'Two').showFormulas, true);
  profile.blocked = false;
  other.patch('foreign', 'One', { zoom: 40 });
  other.patch('file:canonical', 'Untouched', { headings: false });
  const retry = profile.writes;
  owner.patch('file:canonical', 'Two', { gridlines: false });
  assert.equal(profile.writes, retry + 1);
  assert.equal(owner.warning, null);
  const cold = createViewPreferenceStore(profile);
  assert.deepEqual(cold.get('new:uuid', 'One'), DEFAULT_VIEW_PREFERENCES);
  assert.deepEqual(cold.get('file:canonical', 'One'), { ...DEFAULT_VIEW_PREFERENCES, zoom: 125, formulaBar: false });
  assert.deepEqual(cold.get('file:canonical', 'Two'), {
    ...DEFAULT_VIEW_PREFERENCES,
    zoom: 200,
    showFormulas: true,
    gridlines: false,
  });
  assert.equal(cold.get('file:canonical', 'Untouched').headings, false);
  assert.equal(cold.get('foreign', 'One').zoom, 40);
  const limitedProfile = storage(),
    limited = createViewPreferenceStore(limitedProfile, { maxBytes: 400 });
  limited.patch('new', 'One', { zoom: 75 });
  const smallSaved = limitedProfile.value;
  assert.equal(limited.moveDocument('new', 'target'.repeat(80)), true);
  assert.equal(limitedProfile.value, smallSaved, 'An oversized migrated profile is never written');
  assert.equal(limited.get('target'.repeat(80), 'One').zoom, 75);
  assert.deepEqual(limited.get('new', 'One'), DEFAULT_VIEW_PREFERENCES);
  assert.match(limited.warning, /storage limit/);
});
test('corrupt, duplicate, unknown and oversized profile data fail gracefully without write; missing IDs stay private', () => {
  const entry = { documentId: 'file:A', sheet: 'Sheet1', preferences: { ...DEFAULT_VIEW_PREFERENCES, zoom: 125 } };
  for (const raw of [
    'not json',
    JSON.stringify({ schema: 3, entries: [entry] }),
    JSON.stringify({ schema: 1, entries: [entry], extra: true }),
    JSON.stringify({ schema: 1, entries: [entry, entry] }),
    JSON.stringify({ schema: 1, entries: [{ ...entry, preferences: { zoom: 125 } }] }),
    JSON.stringify({ schema: 1, entries: [{ ...entry, preferences: { ...entry.preferences, zoom: Infinity } }] }),
    ' '.repeat(65537),
  ]) {
    const profile = storage(raw),
      store = createViewPreferenceStore(profile);
    assert.deepEqual(store.get('file:A', 'Sheet1'), DEFAULT_VIEW_PREFERENCES);
    assert.ok(store.warning);
    assert.equal(profile.writes, 0);
    assert.equal(profile.value, raw);
  }
  const profile = storage(),
    anonymous = createViewPreferenceStore(profile),
    other = createViewPreferenceStore(profile);
  anonymous.patch(undefined, 'Sheet1', { zoom: 75 });
  assert.equal(anonymous.get(undefined, 'Sheet1').zoom, 75);
  assert.equal(other.get(undefined, 'Sheet1').zoom, 100);
  assert.equal(profile.writes, 0);
  assert.match(anonymous.warning, /Stable document identity/);
  assert.throws(() => anonymous.patch('bad\0id', 'Sheet1', { zoom: 100 }));
  assert.throws(() => anonymous.get(0, 'Sheet1'));
  const denied = createViewPreferenceStore({
    getItem() {
      throw Error('Denied');
    },
    setItem() {
      throw Error('Denied');
    },
  });
  assert.doesNotThrow(() => denied.patch('file:A', 'Sheet1', { zoom: 400 }));
  assert.equal(denied.get('file:A', 'Sheet1').zoom, 400);
});
test('profile entry and UTF8 byte quotas evict old private preferences, never write oversized data', () => {
  const profile = storage(),
    store = createViewPreferenceStore(profile, { maxEntries: 2, maxBytes: 1024 });
  store.patch('A', 'One', { zoom: 125 });
  store.patch('B', 'Two', { zoom: 200 });
  store.patch('C', 'Three', { zoom: 40 });
  const loaded = createViewPreferenceStore(profile, { maxEntries: 2, maxBytes: 1024 });
  assert.deepEqual(loaded.get('A', 'One'), DEFAULT_VIEW_PREFERENCES);
  assert.equal(loaded.get('B', 'Two').zoom, 200);
  assert.equal(loaded.get('C', 'Three').zoom, 40);
  assert.equal(JSON.parse(profile.value).entries.length, 2);
  assert.ok(Buffer.byteLength(profile.value) <= 1024);
  const unicode = storage(),
    limited = createViewPreferenceStore(unicode, { maxBytes: 900 });
  limited.patch('A'.repeat(200), 'One', { zoom: 125 });
  limited.patch('🙂'.repeat(200), 'Two', { zoom: 200 });
  assert.ok(unicode.value === null || Buffer.byteLength(unicode.value) <= 900);
  assert.equal(limited.get('🙂'.repeat(200), 'Two').zoom, 200);
  const tiny = storage(),
    small = createViewPreferenceStore(tiny, { maxBytes: 1 });
  small.patch('A', 'One', { zoom: 125 });
  assert.equal(tiny.writes, 0);
  assert.equal(small.get('A', 'One').zoom, 125);
  assert.ok(small.warning);
  assert.equal(small.removeDocument('A'), true);
  assert.equal(tiny.writes, 0);
  assert.throws(() => createViewPreferenceStore(profile, { maxEntries: 65 }));
  assert.throws(() => createViewPreferenceStore(profile, { maxBytes: 65537 }));
  assert.equal(VIEW_PREFERENCES_KEY, 'workbook.view-preferences.v1');
});
test('125% roomy delimited prefixes honor custom units, hidden/zero rows, hidden columns and first overlapping interval', () => {
  const before = JSON.stringify(geometrySheet),
    g = createViewGeometry(geometrySheet, { rows: 6, columns: 6, zoom: 125, headings: true });
  assert.deepEqual([...g.rowOffsets], [0, 30, 80, 80, 152.5, 152.5, 225]);
  assert.deepEqual([...g.columnOffsets], [0, 75, 75, 275, 400, 480, 780]);
  assert.equal(g.rowHeaderWidth, 60);
  assert.equal(g.columnHeaderHeight, 40);
  assert.equal(g.frozenHeight, 80);
  assert.equal(g.frozenWidth, 75);
  assert.equal(rowHeight(geometrySheet, 3, 125), 72.5);
  assert.equal(rowHeight(geometrySheet, 2, 125), 0);
  assert.equal(columnWidth(geometrySheet, 4, 125), 80);
  assert.equal(columnWidth(geometrySheet, 1, 125), 0);
  assert.equal(offsetForRow(geometrySheet, 3, 125), 80);
  assert.equal(offsetForColumn(geometrySheet, 3, 125), 275);
  assert.equal(JSON.stringify(geometrySheet), before);
  for (const zoom of [10, 40, 100, 200, 400]) {
    const scale = zoom / 100,
      model = createViewGeometry(geometrySheet, { rows: 6, columns: 6, zoom, headings: false });
    close(model.rowOffsets[6], 180 * scale);
    close(model.columnOffsets[6], 624 * scale);
    assert.equal(model.rowHeaderWidth, 0);
    assert.equal(model.columnHeaderHeight, 0);
  }
  assert.throws(() => rowOffsets({}, 1_048_577));
  assert.throws(() => columnOffsets({}, 16_385));
  assert.throws(() => createViewGeometry({}, { rows: -1, columns: 12 }));
  assert.throws(() => rowOffsets({ rowHeights: [{ row: 0, height: NaN }] }, 2));
  assert.throws(() => columnOffsets({ widths: [{ min: 3, max: 2, width: 5, hidden: false }] }, 6));
});
test('binary windows skip zero-size runs, include physical viewport and frozen cells without silent visible cap', () => {
  const g = createViewGeometry(geometrySheet, { rows: 6, columns: 6, zoom: 125 });
  assert.deepEqual(virtualWindow(g.rowOffsets, 75, 200, { headerExtent: 40, frozenCount: 2, overscan: 0 }), {
    first: 5,
    last: 6,
    indices: [5],
    frozenIndices: [0, 1],
  });
  assert.deepEqual(virtualWindow(g.columnOffsets, 180, 300, { headerExtent: 60, frozenCount: 2, overscan: 0 }), {
    first: 2,
    last: 5,
    indices: [2, 3, 4],
    frozenIndices: [0],
  });
  const small = createViewGeometry({}, { rows: 1000, columns: 1000, zoom: 10, headings: false }),
    window = virtualWindow(small.rowOffsets, 0, 1160, { overscan: 0 });
  assert.equal(window.first, 0);
  assert.equal(window.last, 200);
  assert.equal(window.indices.length, 200, 'Do not silently retain a legacy 30-row viewport cap');
  const hidden = createViewGeometry({ hiddenRows: [0, 1, 2, 3, 4, 5, 6, 7, 8] }, { rows: 10, columns: 1 });
  assert.deepEqual(virtualWindow(hidden.rowOffsets, 0, 58, { overscan: 0 }).indices, [9]);
  assert.deepEqual(virtualWindow(hidden.rowOffsets, 0, 0, { overscan: 0 }).indices, []);
  assert.deepEqual(virtualWindow(new Float64Array([0]), 0, 100), { first: 0, last: 0, indices: [], frozenIndices: [] });
});
test('pointer coordinates map actual zoom/hidden/frozen body and merged anchors without changing input model', () => {
  const g = createViewGeometry(geometrySheet, { rows: 6, columns: 6, zoom: 125 });
  assert.deepEqual(hitTestCell(g, 140, 145, { scrollLeft: 180, scrollTop: 75 }), { row: 5, col: 2 });
  assert.deepEqual(hitTestCell(g, 80, 50, { scrollLeft: 180, scrollTop: 75 }), { row: 0, col: 0 });
  assert.equal(hitTestCell(g, 5, 50, { scrollLeft: 0, scrollTop: 0 }), null);
  assert.deepEqual(hitTestCell(g, 140, 145, { scrollLeft: 180, scrollTop: 75 }, geometrySheet), { row: 3, col: 2 });
  const bounds = mergedCellBounds(geometrySheet, 5, 3);
  assert.deepEqual(bounds, { first: { row: 3, col: 2 }, last: { row: 5, col: 3 } });
  assert.deepEqual(cellRectangle(g, bounds), { left: 75, top: 80, width: 325, height: 145 });
  assert.deepEqual(mergedCellBounds({ merges: ['$A$1:$C$4'] }, 3, 2), {
    first: { row: 0, col: 0 },
    last: { row: 3, col: 2 },
  });
  assert.equal(hitTestCell(g, 5000, 5000, { scrollLeft: 0, scrollTop: 0 }), null);
  assert.throws(() => mergedCellBounds({ merges: ['XFE1:XFE2'] }, 0, 0));
  assert.throws(() => mergedCellBounds({ merges: ['B2:A1'] }, 0, 0));
});
test('reveal uses sticky headers/frozen strips; selection fitting uses viewport geometry and bounded integer zoom', () => {
  const g = createViewGeometry(geometrySheet, { rows: 10, columns: 10, zoom: 125 });
  assert.deepEqual(revealCellScroll(g, { width: 360, height: 240, scrollLeft: 0, scrollTop: 0 }, 5, 3), {
    scrollLeft: 108,
    scrollTop: 33,
  });
  assert.deepEqual(revealCellScroll(g, { width: 360, height: 240, scrollLeft: 108, scrollTop: 33 }, 0, 0), {
    scrollLeft: 108,
    scrollTop: 33,
  });
  assert.deepEqual(
    revealCellScroll(g, { width: 360, height: 240, scrollLeft: 0, scrollTop: 0 }, 5, 3, geometrySheet),
    { scrollLeft: 0, scrollTop: 0 },
    'Oversized merged rectangle aligns its anchor instead of oscillating between edges'
  );
  const cropped = createViewGeometry(geometrySheet, { rows: 6, columns: 6, zoom: 125 });
  assert.deepEqual(
    revealCellScroll(cropped, { width: 360, height: 240, scrollLeft: 0, scrollTop: 0 }, 5, 3),
    { scrollLeft: 108, scrollTop: 25 },
    'Bottom reveal padding is clamped to the actual content end'
  );
  const target = { first: { row: 5, col: 3 }, last: { row: 5, col: 3 } };
  assert.equal(zoomToSelection(g, target, { width: 416, height: 308 }), 200);
  assert.equal(zoomToSelection(g, target, { width: 208, height: 154 }), 100);
  assert.equal(zoomToSelection(g, target, { width: 10, height: 10 }), 10);
  assert.equal(zoomToSelection(g, target, { width: 10000, height: 10000 }), 400);
  assert.equal(zoomToSelection(g, target, { width: 0, height: 0 }), 125);
  const noHeadings = createViewGeometry(geometrySheet, { rows: 6, columns: 6, zoom: 125, headings: false });
  assert.equal(zoomToSelection(noHeadings, target, { width: 320, height: 244 }), 200);
  assert.equal(
    zoomToSelection(g, { first: { row: 0, col: 0 }, last: { row: 5, col: 5 } }, { width: 672, height: 212 }),
    100
  );
  assert.equal(
    zoomToSelection(g, { first: { row: 2, col: 1 }, last: { row: 2, col: 1 } }, { width: 600, height: 400 }),
    125
  );
});
