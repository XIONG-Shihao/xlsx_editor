import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  INK_EMU_PER_PIXEL as E,
  INK_LIMITS,
  INK_PENS,
  inkPen,
  normalizeInkStroke,
  inkStrokeFromLogicalPoints,
  inkStrokePath,
  inkStrokeHit,
} from '../core/ink-stroke.mjs';
import { inkStrokeAnchorXml, readInkStrokeAnchor } from '../core/ink-stroke-xml.mjs';
import { inkWorksheetPoint, anchoredInkStroke } from '../ui/inkGeometry.ts';
import { createObjectAxes, objectScreenBox } from '../ui/worksheetObjectGeometry.ts';
const id = '11111111-1111-4111-8111-111111111111',
  make = (
    points = [
      [20, 40],
      [120, 40],
    ],
    pen = 'black'
  ) => inkStrokeFromLogicalPoints(points, pen),
  anchor = (stroke) => ({
    kind: 'oneCell',
    cell: 'B3',
    dxEmu: 10 * E,
    dyEmu: 5 * E,
    widthEmu: stroke.widthEmu,
    heightEmu: stroke.heightEmu,
  }),
  xml = (value = make().stroke, alt = 'Plain <&é\r\n😀 stroke') => inkStrokeAnchorXml(id, 1, value, anchor(value), alt);
test('logical capture pads horizontal/vertical paths, keeps ordered loops and detaches caller points', () => {
  const input = [
      [20, 40],
      [20, 40],
      [120, 40],
    ],
    v = inkStrokeFromLogicalPoints(input);
  assert.deepEqual(v, {
    stroke: {
      schema: 1,
      pen: 'black',
      widthEmu: 102 * E,
      heightEmu: 2 * E,
      points: [
        [E, E],
        [101 * E, E],
      ],
    },
    originEmu: { x: 19 * E, y: 39 * E },
  });
  input[0][0] = 999;
  assert.equal(v.stroke.points[0][0], E);
  const vertical = make([
    [20, 30],
    [20, 120],
  ]).stroke;
  assert.equal(vertical.widthEmu, 2 * E);
  assert.equal(vertical.heightEmu, 92 * E);
  const loop = make([
    [20, 20],
    [60, 50],
    [20, 20],
  ]).stroke;
  assert.equal(loop.points.length, 3);
  assert.equal(
    inkStrokePath(loop)
      .split(' ')
      .filter((s) => s.startsWith('M')).length,
    1
  );
  assert(!inkStrokePath(loop).includes('Z'));
});
test('three fixed pen styles give actual distinct widths/caps/opacity and retained model identities', () => {
  assert.equal(INK_PENS.length, 3);
  const bodies = INK_PENS.map((p) =>
    xml(
      make(
        [
          [20, 40],
          [120, 40],
        ],
        p.id
      ).stroke
    )
  );
  assert.equal(new Set(bodies).size, 3);
  for (const p of INK_PENS) {
    assert(Object.isFrozen(p));
    const read = readInkStrokeAnchor(
      xml(
        make(
          [
            [20, 40],
            [120, 40],
          ],
          p.id
        ).stroke
      )
    );
    assert.equal(read.stroke.pen, p.id);
    assert.equal(read.stroke.heightEmu, p.widthEmu);
  }
  assert.throws(() => inkPen('imported'));
  assert.throws(() => inkPen({ id: 'black' }));
});
test('finite point/coordinate bounds reject invalid aliases, degenerate taps, oversized drafts and arbitrary styles', () => {
  const valid = make().stroke;
  normalizeInkStroke(valid);
  for (const v of [
    { ...valid, pressure: 0.5 },
    { ...valid, schema: 2 },
    { ...valid, pen: 'blue' },
    { ...valid, widthEmu: Infinity },
    { ...valid, heightEmu: valid.heightEmu + 1 },
    {
      ...valid,
      points: [
        [E, E],
        [E, E],
      ],
    },
    {
      ...valid,
      points: [
        [0, E],
        [101 * E, E],
      ],
    },
    {
      ...valid,
      points: [
        [E, E],
        [101 * E, NaN],
      ],
    },
  ])
    assert.throws(() => normalizeInkStroke(v));
  for (const points of [
    [[1, 1]],
    [
      [1, 1],
      [1, 1],
    ],
    [
      [0, 0],
      [10, 10],
    ],
    [
      [NaN, 1],
      [2, 2],
    ],
    [
      [1, Infinity],
      [2, 2],
    ],
    [
      [1, 1],
      [1e9 + 1, 1],
    ],
    [
      [2, 2],
      [6000, 2],
    ],
    Array.from({ length: 513 }, (_, i) => [20 + i, 20]),
    Array.from({ length: INK_LIMITS.inputPoints + 1 }, () => [20, 20]),
  ])
    assert.throws(() => make(points));
  const max = make(Array.from({ length: INK_LIMITS.points }, (_, i) => [20 + i, 20 + (i % 3)])).stroke;
  assert.equal(max.points.length, 512);
});
test('ordinary zoom/scroll conversion maps back to one logical worksheet stroke and original one-cell axes', () => {
  const points = [
      [90, 50],
      [130, 80],
      [190, 55],
    ],
    sheet = {
      sheetFormat: { defaultColumnWidth: 10, defaultRowHeight: 15 },
      normalFont: { fontFamily: 'Calibri', fontSize: 11, bold: false, italic: false },
    },
    axes = createObjectAxes(sheet),
    expected = anchoredInkStroke(points, 'black', axes);
  assert.equal(expected.anchor.cell, 'B3');
  assert.equal(expected.anchor.dxEmu, 19 * E);
  assert.equal(expected.anchor.dyEmu, 9 * E);
  for (const scale of [0.5, 1, 1.25, 2, 4]) {
    const view = {
        scale,
        rowHeaderWidth: 48 * scale,
        columnHeaderHeight: 32 * scale,
        scrollLeft: 25 * scale,
        scrollTop: 30 * scale,
      },
      converted = points.map(([x, y]) =>
        inkWorksheetPoint(
          view,
          view.rowHeaderWidth + x * scale - view.scrollLeft,
          view.columnHeaderHeight + y * scale - view.scrollTop
        )
      );
    assert.deepEqual(anchoredInkStroke(converted, 'black', axes), expected);
    const shown = objectScreenBox(expected.box, view);
    assert.equal(shown.width, expected.box.width * scale);
    assert.equal(shown.height, expected.box.height * scale);
  }
  assert.equal(
    inkWorksheetPoint({ scale: 1, rowHeaderWidth: 48, columnHeaderHeight: 32, scrollLeft: 0, scrollTop: 0 }, 20, 20),
    null
  );
  assert.throws(
    () =>
      inkWorksheetPoint(
        { scale: 1, rowHeaderWidth: 48, columnHeaderHeight: 32, scrollLeft: 0, scrollTop: 0, frozenRows: 1 },
        100,
        100
      ),
    /frozen/
  );
  assert.throws(() =>
    inkWorksheetPoint({ scale: 0, rowHeaderWidth: 0, columnHeaderHeight: 0, scrollLeft: 0, scrollTop: 0 }, 100, 100)
  );
});
test('one-cell DrawingML roundtrip preserves exact paths, semantic identity, style and literal alternative text', () => {
  for (const pen of INK_PENS) {
    const model = make(
        [
          [20, 20],
          [90, 45],
          [70, 80],
        ],
        pen.id
      ).stroke,
      source = xml(model),
      read = readInkStrokeAnchor(source);
    assert.deepEqual(read, { id, number: 1, stroke: model, anchor: anchor(model), altText: 'Plain <&é\r\n😀 stroke' });
    assert(!source.includes('<a:close'));
    assert(!source.includes('<xdr:pic'));
    assert(!source.includes('<xdr:txBody'));
  }
  for (const a of [
    { ...anchor(make().stroke), cell: 'a1' },
    { ...anchor(make().stroke), cell: 'XFE1' },
    { ...anchor(make().stroke), dxEmu: -1 },
    { ...anchor(make().stroke), widthEmu: 1 },
    { ...anchor(make().stroke), kind: 'twoCell' },
  ])
    assert.throws(() => inkStrokeAnchorXml(id, 1, make().stroke, a));
  assert.throws(() => inkStrokeAnchorXml('imported', 1, make().stroke, anchor(make().stroke)));
  assert.throws(() => inkStrokeAnchorXml(id, 0, make().stroke, anchor(make().stroke)));
  for (const alt of ['\u0000', '\ud800', '_x0000_', 'x'.repeat(257)]) assert.throws(() => xml(make().stroke, alt));
});
test('semantic codec accepts namespace-prefix/attribute-order/spacing differences without claiming general imported ink', () => {
  const original = xml(),
    alias = original
      .replaceAll('xdr:', 'sheetDraw:')
      .replaceAll('a:', 'drawing:')
      .replace('xmlns:xdr=', 'xmlns:sheetDraw=')
      .replace('xmlns:a=', 'xmlns:drawing=')
      .replace('id="1" name="workbook:' + id + '"', 'name="workbook:' + id + '" id="1"')
      .replaceAll('><', '>\n  <');
  assert.deepEqual(readInkStrokeAnchor(alias), readInkStrokeAnchor(original));
});
test('extra paths, closed fills, text/effects, hyperlinks, altered dimensions/styles and XML entities remain outside the closed codec', () => {
  const source = xml();
  for (const changed of [
    source.replace('fill="none"', 'fill="norm"'),
    source.replace('</a:path>', '<a:close/></a:path>'),
    source.replace('w="19050"', 'w="19051"'),
    source.replace('<a:alpha val="100000"/>', '<a:alpha val="50000"/>'),
    source.replace('<a:xfrm>', '<a:xfrm rot="60000">'),
    source.replace('<xdr:cNvSpPr/>', '<xdr:cNvSpPr/><a:hlinkClick r:id="foreign" xmlns:r="urn:foreign"/>'),
    source.replace('</xdr:sp>', '<xdr:txBody/></xdr:sp>'),
    source.replace('</a:pathLst>', '<a:path/></a:pathLst>'),
    source.replace('<a:round/>', '<a:miter/>'),
    source.replace('name="workbook:', 'name="foreign:'),
    '<!DOCTYPE root [<!ENTITY outside SYSTEM "file:///outside">]>' + source,
  ])
    assert.throws(() => readInkStrokeAnchor(changed));
});
test('whole-path hit groundwork follows the actual segments instead of treating the entire bounding rectangle as ink', () => {
  const model = make([
    [20, 20],
    [80, 80],
  ]).stroke;
  assert(inkStrokeHit(model, 31 * E, 31 * E));
  assert(!inkStrokeHit(model, 31 * E, 3 * E));
  assert(!inkStrokeHit(model, -1e20, 0));
  assert.throws(() => inkStrokeHit(model, NaN, 0));
  assert.throws(() => inkStrokeHit(model, 0, 0, -1));
  const x = 31 * E,
    y = 34 * E;
  assert(!inkStrokeHit(model, x, y));
  assert(inkStrokeHit(model, x, y, 3 * E));
});
