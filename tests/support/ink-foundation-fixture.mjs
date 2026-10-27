/** Constructor-only format/preview samples. This is not the native workbook writer. */
import fs from 'node:fs/promises';
import path from 'node:path';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook } from '../../core/workbook-package.mjs';
import { inkStrokeAnchorXml } from '../../core/ink-stroke-xml.mjs';
import { anchoredInkStroke } from '../../ui/inkGeometry.ts';
import { createObjectAxes } from '../../ui/worksheetObjectGeometry.ts';
const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  D = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
export const INK_FOUNDATION_SAMPLES = Object.freeze([
  {
    id: '11111111-1111-4111-8111-111111111111',
    number: 1,
    pen: 'black',
    label: 'Black zigzag',
    points: [
      [90, 50],
      [130, 80],
      [190, 55],
    ],
  },
  {
    id: '22222222-2222-4222-8222-222222222222',
    number: 2,
    pen: 'red',
    label: 'Red check',
    points: [
      [260, 80],
      [280, 100],
      [330, 35],
    ],
  },
  {
    id: '33333333-3333-4333-8333-333333333333',
    number: 3,
    pen: 'yellow-highlighter',
    label: 'Yellow line',
    points: [
      [20, 120],
      [200, 120],
    ],
  },
]);
export function inkFoundationCases() {
  const axes = createObjectAxes({ sheetFormat: { defaultColumnWidth: 10, defaultRowHeight: 15 } });
  return INK_FOUNDATION_SAMPLES.map((s) => ({ ...structuredClone(s), ...anchoredInkStroke(s.points, s.pen, axes) }));
}
export async function inkFoundationFixture(directory) {
  await fs.mkdir(directory, { recursive: true });
  const source = path.join(directory, 'source.xlsx'),
    saved = path.join(directory, 'vector-fixture.xlsx'),
    manifest = path.join(directory, 'expected-input.json'),
    parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${S}"><sheetFormatPr defaultColWidth="10" defaultRowHeight="15"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Preserved cell &amp; style</t></is></c><c r="B1"><v>42</v></c></row></sheetData></worksheet>`
  );
  parts['custom/ink-opaque.bin'] = new Uint8Array([1, 9, 17, 255]);
  await fs.writeFile(source, zipSync(parts));
  const cases = inkFoundationCases(),
    anchors = cases.map((c) => inkStrokeAnchorXml(c.id, c.number, c.stroke, c.anchor, c.label)).join('');
  parts['xl/drawings/workbook-ink-fixture.xml'] = strToU8(`<xdr:wsDr xmlns:xdr="${D}">${anchors}</xdr:wsDr>`);
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    strFromU8(parts['xl/worksheets/sheet1.xml']).replace(
      '</worksheet>',
      `<drawing xmlns:r="${R}" r:id="inkDrawing"/></worksheet>`
    )
  );
  parts['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="inkDrawing" Type="${R}/drawing" Target="../drawings/workbook-ink-fixture.xml"/></Relationships>`
  );
  parts['[Content_Types].xml'] = strToU8(
    strFromU8(parts['[Content_Types].xml']).replace(
      '</Types>',
      '<Override PartName="/xl/drawings/workbook-ink-fixture.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>'
    )
  );
  await fs.writeFile(saved, zipSync(parts));
  await fs.writeFile(
    manifest,
    JSON.stringify(
      {
        schema: 'workbook.xlsx.ink-foundation-input.v1',
        scope: 'Constructed format fixture; no native/editor write qualification',
        geometry: { columnPixels: 70, rowPixels: 20 },
        samples: INK_FOUNDATION_SAMPLES,
      },
      null,
      2
    ) + '\n'
  );
  return { source, saved, manifest, cases };
}
