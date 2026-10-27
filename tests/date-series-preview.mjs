import test from 'node:test';
import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { fileURLToPath } from 'node:url';
const bundle = await build({
  entryPoints: [fileURLToPath(new URL('../ui/fillSeriesPreview.ts', import.meta.url))],
  bundle: true,
  write: false,
  platform: 'browser',
  format: 'esm',
  logLevel: 'silent',
});
const { createFillSeriesPreview, fillSeriesPreviewValue, fillSeriesCommands, fillSeriesPreviewPage } = await import(
  'data:text/javascript;base64,' + Buffer.from(bundle.outputFiles[0].contents).toString('base64')
);
import { projectWorkbookCommands } from '../core/workbook-projection.mjs';
const binding = (range) => ({
  documentId: 'book',
  ownerIdentity: 'native-owner',
  state: 'head',
  sheet: 'Data',
  selection: range,
  selectedCell: range.split(':')[0],
  activationEpoch: 0,
  navigationGeneration: 0,
  modeGeneration: 0,
  specialSelectionEpoch: 0,
  mode: 'normal',
  specialSelection: null,
});
function input(a = 45350, b = 45351, range = 'A1:A5', date1904 = false) {
  return {
    binding: binding(range),
    kind: 'date',
    sheet: {
      name: 'Data',
      path: 'xl/worksheets/sheet1.xml',
      rows: 10000,
      columns: 26,
      cells: {},
      merges: [],
      widths: [],
      date1904,
    },
    cells: {
      A1: { value: a, rawValue: String(a), format: 14, fontColor: '#123456' },
      [range.endsWith('1') ? 'B1' : 'A2']: { value: b, rawValue: String(b), format: 14, bold: true },
    },
    readRange: range,
  };
}
function ready(data) {
  const result = createFillSeriesPreview(data);
  assert.equal(result.status, 'ready', result.reason);
  return result.preview;
}
test('whole-day dates preview leap/day/month rollover and generate contents plus date formatting as one batch', () => {
  const data = input(),
    before = structuredClone(data),
    p = ready(data);
  assert.deepEqual(p.date, { date1904: false });
  assert.deepEqual(
    p.destinations.map((row) => fillSeriesPreviewValue(p, row.value)),
    ['2024-03-01', '2024-03-02', '2024-03-03']
  );
  assert.equal(fillSeriesPreviewValue(p, p.seeds[0].value), '2024-02-28');
  assert.equal(fillSeriesPreviewValue(p, p.seeds[1].value), '2024-02-29');
  assert.deepEqual(fillSeriesCommands(p), [
    { type: 'fillSeries', sheet: 'Data', range: 'A1:A5' },
    { type: 'formatRange', sheet: 'Data', range: 'A3:A5', format: { numberFormat: 'date' } },
  ]);
  assert.deepEqual(data, before);
  assert(Object.isFrozen(p.date));
});
test('1904 source calendar and horizontal descending/constant steps use UTC dates', () => {
  const p = ready(input(0, 1, 'A1:D1', true));
  assert.deepEqual(
    p.destinations.map((row) => fillSeriesPreviewValue(p, row.value)),
    ['1904-01-03', '1904-01-04']
  );
  const descending = ready(input(44929, 44928, 'A1:D1'));
  assert.deepEqual(
    descending.destinations.map((row) => fillSeriesPreviewValue(descending, row.value)),
    ['2023-01-01', '2022-12-31']
  );
  const same = ready(input(45292, 45292));
  assert(same.destinations.every((row) => fillSeriesPreviewValue(same, row.value) === '2024-01-01'));
});
test('1900 compatibility serial 60 and upper date limits are explicit, without time fractions or overflowing dates', () => {
  const p = ready(input(58, 59));
  assert.deepEqual(
    p.destinations.map((row) => fillSeriesPreviewValue(p, row.value)),
    ['1900-02-29', '1900-03-01', '1900-03-02']
  );
  for (const [a, b, calendar] of [
    [1, 0, false],
    [0, 1, false],
    [1, 1.5, false],
    [2958464, 2958465, false],
    [2957002, 2957003, true],
  ])
    assert.equal(createFillSeriesPreview(input(a, b, 'A1:A5', calendar)).status, 'refused');
  const end = ready(input(2958463, 2958464, 'A1:A3'));
  assert.equal(fillSeriesPreviewValue(end, end.destinations[0].value), '9999-12-31');
});
test('date text, calculated seeds, occupied destinations and unsupported mode cannot become a partial plan', () => {
  for (const patch of [
    { A1: { value: '2024-02-28' } },
    { A1: { value: 45350, formula: 'DATE(2024,2,28)' } },
    { A4: { value: 'keep' } },
    { A4: { value: null, formulaType: 'array' } },
  ]) {
    const data = input();
    Object.assign(data.cells, patch);
    const before = structuredClone(data),
      result = createFillSeriesPreview(data);
    assert.equal(result.status, 'refused');
    assert(!Object.hasOwn(result, 'preview'));
    assert.deepEqual(data, before);
  }
  assert.equal(createFillSeriesPreview({ ...input(), kind: 'months' }).status, 'refused');
  assert.equal(
    createFillSeriesPreview({ ...input(), sheet: { ...input().sheet, date1904: 'unknown' } }).status,
    'refused'
  );
});
test('date format changes only generated number format while seeds and unrelated destination formatting survive', async () => {
  const data = input();
  data.cells.A3 = { value: null, format: 2, bold: true, fontColor: '#FF0000' };
  const p = ready(data),
    view = { revision: 'head', readOnly: false, limitations: [], sheets: [{ ...data.sheet, cells: data.cells }] };
  const result = await projectWorkbookCommands(view, fillSeriesCommands(p), {}),
    cells = result.view.sheets[0].cells;
  assert.deepEqual(cells.A1, data.cells.A1);
  assert.deepEqual(cells.A2, data.cells.A2);
  assert.equal(cells.A3.value, 45352);
  assert.equal(cells.A3.bold, true);
  assert.equal(cells.A3.fontColor, '#FF0000');
  assert.equal(cells.A3.format, 14);
});
test('date pages retain their calendar and reject corrupted date metadata', () => {
  const p = ready(input());
  assert.equal(fillSeriesPreviewPage(p).total, 3);
  assert.throws(() => fillSeriesPreviewPage({ ...p, date: { date1904: 'false' } }), /Invalid Date/);
  assert.throws(() => fillSeriesPreviewPage({ ...p, date: { date1904: false, foreign: true } }), /Invalid Date/);
  assert.throws(
    () => fillSeriesPreviewPage({ ...p, destinations: [{ cell: 'A3', value: -1 }, ...p.destinations.slice(1)] }),
    /Invalid Date/
  );
});
