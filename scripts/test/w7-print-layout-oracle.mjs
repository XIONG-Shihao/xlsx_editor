/** Normative layout checks only. Does not compose pages or produce PDF output. */
import fs from 'node:fs';
import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';
const defaults = fileURLToPath(new URL('../../tests/support/w7-print-layout-expected.json', import.meta.url));
export function printLayoutCase(id, filename = defaults) {
  const fixture = JSON.parse(fs.readFileSync(filename, 'utf8')),
    entry = fixture.cases.find((item) => item.id === id);
  assert(entry, 'Unknown expected print case: ' + id);
  return { ...entry, tolerancePt: fixture.tolerancePt };
}
function numbers(actual, expected, tolerance, label) {
  assert.equal(actual?.length, expected.length, label + ' geometry length');
  expected.forEach((value, index) =>
    assert(
      Number.isFinite(actual[index]) && Math.abs(actual[index] - value) <= tolerance,
      `${label}[${index}] expected ${value}, got ${actual[index]}`
    )
  );
}
function expectedCells(test, page) {
  const { columnWidthsPt: widths, rowHeightPt: height, scale, marginsPt, merges = [] } = test.input,
    rows =
      page.rows.length === 2
        ? Array.from({ length: page.rows[1] - page.rows[0] + 1 }, (_, i) => page.rows[0] + i)
        : page.rows;
  const result = new Map(),
    continuations = new Set(),
    merged = new Map();
  for (const range of merges) {
    const [a, b] = range.split(':'),
      column = a.charCodeAt(0) - 65,
      last = b.charCodeAt(0) - 65;
    merged.set(a, { width: widths.slice(column, last + 1).reduce((sum, value) => sum + value, 0), range });
    for (let c = column + 1; c <= last; c++) continuations.add(String.fromCharCode(65 + c) + a.slice(1));
  }
  rows.forEach((row, r) => {
    let x = marginsPt[3];
    widths.forEach((width, c) => {
      const ref = String.fromCharCode(65 + c) + row,
        merge = merged.get(ref);
      if (!continuations.has(ref))
        result.set(ref, {
          bbox: [x, marginsPt[0] + r * height * scale, (merge?.width ?? width) * scale, height * scale],
          ...(merge ? { merge: merge.range } : {}),
        });
      x += width * scale;
    });
  });
  return result;
}
/** layout: {units:'pt',pageSizePt:[w,h],pages:[{number,cells:[{ref,bbox,text,merge?}],objects:[{id,kind,bbox,clipBox,clipped}]}]}. */
export function checkPrintLayout(layout, id, filename = defaults) {
  const test = printLayoutCase(id, filename),
    expected = test.expected;
  if (test.expectedRefusal) {
    for (const [key, value] of Object.entries(test.expectedRefusal))
      assert.equal(layout?.refusal?.[key], value, 'Expected print refusal ' + key);
    assert(!layout.pages?.length, 'Refused job cannot publish composed pages');
    return { passed: true, id, method: 'normative-refusal-contract' };
  }
  assert.equal(layout.units, 'pt');
  numbers(layout.pageSizePt, expected.pageSizePt, test.tolerancePt, 'page size');
  assert.equal(layout.pages.length, expected.pageCount, 'page count');
  for (const wanted of expected.pages) {
    const page = layout.pages.find((value) => value.number === wanted.number);
    assert(page, 'Expected page ' + wanted.number);
    const cells = new Map(page.cells.map((cell) => [cell.ref, cell]));
    assert.equal(cells.size, page.cells.length, 'duplicate page cell refs');
    assert.equal(cells.size, wanted.cellCount, 'page cell count');
    const all = expectedCells(test, wanted);
    assert.deepEqual([...cells.keys()].sort(), [...all.keys()].sort(), 'Complete printed range coverage');
    for (const [ref, geometry] of all) {
      numbers(cells.get(ref).bbox, geometry.bbox, test.tolerancePt, ref + ' complete geometry');
      assert.equal(cells.get(ref).merge, geometry.merge, ref + ' merge geometry');
    }
    for (const [ref, value] of Object.entries(wanted.representativeCells)) {
      const cell = cells.get(ref);
      assert(cell, 'Expected page cell ' + ref);
      numbers(cell.bbox, value.bbox, test.tolerancePt, ref);
      assert.equal(cell.text, value.text, ref + ' text');
      if (value.merge) assert.equal(cell.merge, value.merge, ref + ' merge');
    }
    for (const ref of wanted.absentCells ?? [])
      assert(!cells.has(ref), 'Hidden/merged continuation cell unexpectedly rendered: ' + ref);
    const objects = new Map((page.objects ?? []).map((object) => [object.id, object]));
    assert.equal(objects.size, (page.objects ?? []).length, 'duplicate page object IDs');
    assert.equal(objects.size, wanted.objects?.length ?? 0, 'page object count');
    for (const wantedObject of wanted.objects ?? []) {
      const object = objects.get(wantedObject.id);
      assert(object, 'Expected object ' + wantedObject.id);
      assert.equal(object.kind, wantedObject.kind);
      numbers(object.bbox, wantedObject.bbox, test.tolerancePt, object.id + ' bounds');
      numbers(object.clipBox, wantedObject.clipBox, test.tolerancePt, object.id + ' clip');
      assert.equal(object.clipped, wantedObject.clipped);
    }
  }
  return { passed: true, id, pages: layout.pages.length, method: 'normative-page-composition-contract' };
}
