import { test } from 'node:test';
import assert from 'node:assert/strict';
import { composePrint } from '../core/print-layout.mjs';
import { defaultPrintSettings } from '../core/print-model.mjs';
import { pageBoundaryModel, projectPageBoundary } from '../ui/pageBoundaries.ts';
import { createViewGeometry } from '../ui/viewGeometry.ts';
const col = (n) => String.fromCharCode(65 + n);
function snapshot(patch = {}) {
  const settings = {
    ...defaultPrintSettings(),
    area: 'A1:H80',
    marginsPt: { top: 54, right: 54, bottom: 54, left: 54, header: 18, footer: 18 },
    repeatRows: { first: 1, last: 1 },
    repeatColumns: { first: 'A', last: 'A' },
    ...patch,
  };
  return {
    schema: 1,
    sessionId: 'owner',
    state: 'head',
    revision: 'source',
    sheet: 'Main',
    date1904: false,
    fingerprint: 'a'.repeat(64),
    settings,
    rows: Array.from({ length: 80 }, (_, index) => ({ index, heightPt: 24, hidden: index === 4 })),
    columns: Array.from({ length: 8 }, (_, index) => ({ index, widthPt: 120, hidden: index === 7 })),
    cells: Array.from({ length: 80 }, (_, r) =>
      Array.from({ length: 8 }, (_, c) => ({ ref: col(c) + (r + 1), text: 'r' + r + 'c' + c, style: { fontSize: 8 } }))
    ).flat(),
    merges: [],
    objects: [],
  };
}
function job(patch = {}) {
  const layout = composePrint(snapshot(patch));
  return { ...layout, id: 'lease', state: 'head', pageCount: layout.pages.length, pages: [layout.pages[0]] };
}
test('authoritative page groups summarize repeats/hidden axes without transporting all physical cells', () => {
  const j = job();
  assert.equal(j.pageCount, 6);
  assert.equal(j.pages.length, 1);
  assert.deepEqual(j.boundaries[0], {
    number: 1,
    rows: [0, 30],
    columns: [0, 3],
    bodyRows: [1, 30],
    bodyColumns: [1, 3],
  });
  assert.deepEqual(j.boundaries[1], {
    number: 2,
    rows: [0, 59],
    columns: [0, 3],
    bodyRows: [31, 59],
    bodyColumns: [1, 3],
  });
  assert.deepEqual(j.boundaries[3], {
    number: 4,
    rows: [0, 30],
    columns: [0, 6],
    bodyRows: [1, 30],
    bodyColumns: [4, 6],
  });
  assert(JSON.stringify(j.boundaries).length < 2000);
});
test('source map repeats prefixes once and keeps actual page order and range identity', () => {
  const model = pageBoundaryModel(job());
  assert.deepEqual(
    model.pages.map((p) => p.range),
    ['A1:D31', 'A32:D60', 'A61:D80', 'E1:G31', 'E32:G60', 'E61:G80']
  );
  assert.equal(model.repeatRows, '1:1');
  assert.equal(model.repeatColumns, 'A:A');
  assert.equal(model.rows, 80);
  assert.equal(model.columns, 8);
});
test('projection uses the shared source axes at zoom/scroll including hidden zero extents', () => {
  const model = pageBoundaryModel(job()),
    sheet = {
      sheetFormat: { defaultRowHeight: 24 },
      rowHeights: [],
      hiddenRows: [4],
      widths: [
        { min: 1, max: 7, width: 12, hidden: false },
        { min: 8, max: 8, width: 12, hidden: true },
      ],
      freeze: { rows: 0, columns: 0 },
    },
    geometry = createViewGeometry(sheet, { rows: 80, columns: 8, zoom: 150, headings: true }),
    viewport = { width: 1000, height: 600, scrollLeft: 37, scrollTop: 111 },
    first = projectPageBoundary(model.pages[0], geometry, viewport),
    second = projectPageBoundary(model.pages[1], geometry, viewport);
  assert.equal(first.top, 48 - 111);
  assert.equal(first.height, ((30 * 24 * 96) / 72) * 1.5);
  assert.equal(second.top, first.top + first.height);
  assert.equal(second.height, ((29 * 24 * 96) / 72) * 1.5);
  assert.equal(first.left, 72 - 37);
  assert.equal(first.width, geometry.columnOffsets[4]);
  assert.throws(
    () => projectPageBoundary(model.pages[5], createViewGeometry(sheet, { rows: 10, columns: 2 }), viewport),
    /exceed/
  );
});
test('accepted scale/fit/orientation settings change actual splits, not the source projection algorithm', () => {
  const base = job(),
    scaled = job({ scaling: { mode: 'scale', percent: 150 } }),
    fit = job({ scaling: { mode: 'fit', width: 1, height: 1 } }),
    landscape = job({ orientation: 'landscape' });
  assert(scaled.pageCount > base.pageCount);
  assert.equal(fit.pageCount, 1);
  assert.notDeepEqual(landscape.boundaries, base.boundaries);
  for (const j of [base, scaled, fit, landscape]) assert.equal(pageBoundaryModel(j).pages.length, j.pageCount);
});
test('malformed/foreign summaries and a first-page mismatch cannot become visible boundaries', () => {
  const j = job();
  for (const patch of [
    { boundaries: undefined },
    { pageCount: 7 },
    { boundaries: [...j.boundaries].reverse() },
    { boundaries: j.boundaries.map((p, i) => (i ? p : { ...p, rows: [-1, 30] })) },
    { boundaries: j.boundaries.map((p, i) => (i ? p : { ...p, bodyColumns: [0, 3] })) },
    { boundaries: j.boundaries.map((p, i) => (i ? p : { ...p, rows: [0, 29] })) },
    { boundaries: j.boundaries.map((p, i) => (i ? p : { ...p, extra: 'grant' })) },
  ])
    assert.throws(() => pageBoundaryModel({ ...j, ...patch }));
});
