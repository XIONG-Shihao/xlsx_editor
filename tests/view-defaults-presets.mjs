import test from 'node:test';
import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8 } from 'fflate';
import { createWorkbook, readWorkbook, readWorkbookMetadata } from '../core/workbook-package.mjs';
const bundled = await build({
  entryPoints: [
    fileURLToPath(new URL('../ui/viewPresets.ts', import.meta.url)),
    fileURLToPath(new URL('../ui/viewPreferences.ts', import.meta.url)),
  ],
  bundle: true,
  write: false,
  outdir: 'out',
  platform: 'browser',
  format: 'esm',
  logLevel: 'silent',
});
const modules = Object.fromEntries(
  await Promise.all(
    bundled.outputFiles.map(async (file) => [
      file.path.split('/').at(-1).replace('.js', ''),
      await import('data:text/javascript;base64,' + Buffer.from(file.contents).toString('base64')),
    ])
  )
);
const { ViewPresetStore, BUILTIN_VIEW_PRESETS, VIEW_PRESETS_KEY } = modules.viewPresets,
  { createViewPreferenceStore, DEFAULT_VIEW_PREFERENCES, VIEW_PREFERENCES_KEY } = modules.viewPreferences;
function storage() {
  const values = new Map();
  let denied = false;
  return {
    getItem: (key) => values.get(key) ?? null,
    setItem: (key, value) => {
      if (denied) throw Error('denied');
      values.set(key, value);
    },
    values,
    set denied(value) {
      denied = value;
    },
  };
}
function fixture(views) {
  const parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${views}<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>`
  );
  return zipSync(parts);
}
test('primary source zoom/gridline/heading flags are read from full and metadata views without rewriting bytes', () => {
  const bytes = fixture(
      '<sheetViews><sheetView workbookViewId="0" zoomScale="175" showGridLines="false" showRowColHeaders="0"/></sheetViews>'
    ),
    before = new Uint8Array(bytes);
  assert.deepEqual(readWorkbook(bytes).sheets[0].sourceView, { zoom: 175, gridlines: false, headings: false });
  assert.deepEqual(readWorkbookMetadata(bytes).view.sheets[0].sourceView, {
    zoom: 175,
    gridlines: false,
    headings: false,
  });
  assert.deepEqual(bytes, before);
});
test('invalid/ambiguous/foreign source views do not grant guessed defaults; valid flags survive an invalid zoom', () => {
  for (const xml of [
    '',
    '<sheetViews><sheetView workbookViewId="9" zoomScale="200"/></sheetViews>',
    '<sheetViews><sheetView workbookViewId="0"/><sheetView workbookViewId="0" zoomScale="200"/></sheetViews>',
    '<sheetViews><sheetView xmlns="urn:foreign" workbookViewId="0" zoomScale="200"/></sheetViews>',
  ])
    assert.deepEqual(readWorkbook(fixture(xml)).sheets[0].sourceView, {});
  assert.deepEqual(
    readWorkbook(
      fixture(
        '<sheetViews><sheetView workbookViewId="0" zoomScale="401" showGridLines="1" showRowColHeaders="bad"/></sheetViews>'
      )
    ).sheets[0].sourceView,
    { gridlines: true }
  );
});
test('only explicit local fields override source, and a persisted reset reveals changing source defaults', () => {
  const profile = storage(),
    store = createViewPreferenceStore(profile),
    source = { zoom: 175, gridlines: false, headings: false };
  assert.deepEqual(store.get('file', 'One', source), { ...DEFAULT_VIEW_PREFERENCES, ...source });
  store.patch('file', 'One', { zoom: 200 }, source);
  assert.deepEqual(JSON.parse(profile.values.get(VIEW_PREFERENCES_KEY)).entries[0].preferences, { zoom: 200 });
  const cold = createViewPreferenceStore(profile);
  assert.deepEqual(cold.get('file', 'One', { zoom: 80, gridlines: true, headings: false }), {
    ...DEFAULT_VIEW_PREFERENCES,
    zoom: 200,
    headings: false,
  });
  cold.reset('file', 'One');
  assert.deepEqual(createViewPreferenceStore(profile).get('file', 'One', source), {
    ...DEFAULT_VIEW_PREFERENCES,
    ...source,
  });
});
test('schema1 choices stay explicit after migration while schema2 sparse fields continue following source', () => {
  const profile = storage();
  profile.values.set(
    VIEW_PREFERENCES_KEY,
    JSON.stringify({
      schema: 1,
      entries: [{ documentId: 'old', sheet: 'One', preferences: { ...DEFAULT_VIEW_PREFERENCES, zoom: 125 } }],
    })
  );
  const store = createViewPreferenceStore(profile);
  assert.equal(store.get('old', 'One', { gridlines: false }).gridlines, true);
  store.patch('new', 'One', { headings: false });
  const parsed = JSON.parse(profile.values.get(VIEW_PREFERENCES_KEY));
  assert.equal(parsed.schema, 2);
  assert.deepEqual(parsed.entries.find((entry) => entry.documentId === 'new').preferences, { headings: false });
});
test('failed reset does not resurrect overrides and survives another store write on retry', () => {
  const profile = storage(),
    store = createViewPreferenceStore(profile);
  store.patch('file', 'One', { zoom: 200 });
  profile.denied = true;
  store.reset('file', 'One');
  assert.equal(store.get('file', 'One', { zoom: 175 }).zoom, 175);
  profile.denied = false;
  createViewPreferenceStore(profile).patch('other', 'One', { zoom: 50 });
  store.patch('file', 'One', { headings: false });
  assert.equal(createViewPreferenceStore(profile).get('file', 'One', { zoom: 175 }).zoom, 175);
});
test('three built-ins and named create/update/delete remain isolated by file/sheet with cold persistence', () => {
  assert.equal(BUILTIN_VIEW_PRESETS.length, 3);
  const profile = storage(),
    store = new ViewPresetStore(profile);
  store.create('a', 'One', 'Entry', BUILTIN_VIEW_PRESETS[1].preferences);
  store.create('a', 'Two', 'Entry', BUILTIN_VIEW_PRESETS[2].preferences);
  store.create('b', 'One', 'Entry', DEFAULT_VIEW_PREFERENCES);
  assert.throws(() => store.create('a', 'One', 'entry', DEFAULT_VIEW_PREFERENCES), /already exists/);
  assert.throws(() => store.create('a', 'One', 'Default', DEFAULT_VIEW_PREFERENCES), /built-in/);
  store.update('a', 'One', 'Entry', { ...DEFAULT_VIEW_PREFERENCES, zoom: 200 });
  const cold = new ViewPresetStore(profile);
  assert.equal(cold.list('a', 'One')[0].preferences.zoom, 200);
  assert.equal(cold.list('a', 'Two')[0].preferences.showFormulas, true);
  assert.equal(cold.list('b', 'One')[0].preferences.zoom, 100);
  cold.delete('a', 'One', 'Entry');
  assert.deepEqual(new ViewPresetStore(profile).list('a', 'One'), []);
});
test('denied named writes retain this session, merge unrelated updates on retry and migrate first-Save identities', () => {
  const profile = storage(),
    a = new ViewPresetStore(profile),
    b = new ViewPresetStore(profile);
  profile.denied = true;
  a.create('new', 'One', 'Draft', { ...DEFAULT_VIEW_PREFERENCES, zoom: 125 });
  assert(a.warning);
  assert.equal(a.list('new', 'One')[0].preferences.zoom, 125);
  a.moveDocument('new', 'saved');
  assert.deepEqual(a.list('new', 'One'), []);
  assert.equal(a.list('saved', 'One')[0].preferences.zoom, 125);
  profile.denied = false;
  b.create('other', 'Two', 'Keep', DEFAULT_VIEW_PREFERENCES);
  a.update('saved', 'One', 'Draft', { ...DEFAULT_VIEW_PREFERENCES, zoom: 150 });
  const cold = new ViewPresetStore(profile);
  assert.equal(cold.list('saved', 'One')[0].preferences.zoom, 150);
  assert.equal(cold.list('other', 'Two').length, 1);
  assert.deepEqual(cold.list('new', 'One'), []);
  const transient = new ViewPresetStore(null);
  transient.create('new', 'One', 'Session', DEFAULT_VIEW_PREFERENCES);
  transient.moveDocument('new', 'saved');
  assert.equal(transient.list('saved', 'One').length, 1);
});
test('bounded named catalogs, detached snapshots and corrupt input refuse without silently writing during read', () => {
  const profile = storage(),
    store = new ViewPresetStore(profile);
  for (let i = 0; i < 32; i++) store.create('file', 'One', 'View ' + i, DEFAULT_VIEW_PREFERENCES);
  assert.throws(() => store.create('file', 'One', 'Extra', DEFAULT_VIEW_PREFERENCES), /32 saved/);
  assert.equal(store.list('file', 'One').length, 32);
  assert(Buffer.byteLength(profile.values.get(VIEW_PRESETS_KEY)) <= 65536);
  const returned = store.list('file', 'One');
  returned[0].preferences.zoom = 400;
  assert.equal(store.list('file', 'One')[0].preferences.zoom, 100);
  for (const name of ['', 'bad\0name', 'x'.repeat(65)])
    assert.throws(() => store.create('file', 'One', name, DEFAULT_VIEW_PREFERENCES));
  const corrupt = storage();
  corrupt.values.set(VIEW_PRESETS_KEY, 'invalid');
  const cold = new ViewPresetStore(corrupt);
  assert.deepEqual(cold.list('file', 'One'), []);
  assert(cold.warning);
  assert.equal(corrupt.values.get(VIEW_PRESETS_KEY), 'invalid');
});
