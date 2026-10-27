import assert from 'node:assert/strict';
import { test } from 'node:test';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { createWorkbook, readWorkbook } from '../core/workbook-package.mjs';
import { indexedPrintGeometry } from '../core/indexed-print.mjs';
import { defaultPrintSettings } from '../core/print-model.mjs';
import { referenceGeometryProfile } from '../core/worksheet-geometry.mjs';
import { printPageHtml } from '../core/print-layout.mjs';

function source(format = '<sheetFormatPr defaultRowHeight="22.5" defaultColWidth="9.140625" zeroHeight="true"/>') {
  const parts = unzipSync(createWorkbook());
  const body =
    format +
    '<cols><col min="1" max="1" width="8.7109375"/><col min="2" max="2" width="0.4"/><col min="3" max="3" width="12" hidden="1"/></cols><sheetData><row r="1" hidden="0"/><row r="2" hidden="1"/><row r="3" hidden="false" ht="30"/></sheetData>';
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    strFromU8(parts['xl/worksheets/sheet1.xml']).replace('<sheetData/>', body)
  );
  return zipSync(parts);
}
const settings = { ...defaultPrintSettings(), area: 'A1:D4' };
test('narrow print cells cannot grow beyond their accepted width because of padding or borders', () => {
  const layout = {
    sheet: 'Sheet1',
    scale: 1,
    settings: { gridlines: true },
    pageSizePt: [595, 842],
    pages: [{ number: 1, objects: [], cells: [{ ref: 'B1', text: '', bbox: [54, 54, 0.2, 15], style: {} }] }],
  };
  const html = printPageHtml(layout, 1);
  assert(html.includes('width:0.2pt;'));
  assert(html.includes('padding:0 0pt;'));
  assert(html.includes('border:0;box-shadow:inset 0 0 0 0.1pt'));
  layout.settings.gridlines = false;
  layout.pages[0].cells[0].bbox[2] = 2.25;
  assert(printPageHtml(layout, 1).includes('padding:0 1.125pt;'));
});
test('physical axes use source defaults, narrow widths and hidden-row overrides with the same digit profile', () => {
  const bytes = source(),
    before = bytes.slice(),
    view = readWorkbook(bytes),
    profile = referenceGeometryProfile(view.sheets[0].normalFont);
  const geometry = indexedPrintGeometry(bytes, view, 'Sheet1', settings, profile);
  assert.deepEqual(geometry.rows, [
    { index: 0, heightPt: 22.5, hidden: false },
    { index: 1, heightPt: 0, hidden: true },
    { index: 2, heightPt: 30, hidden: false },
    { index: 3, heightPt: 0, hidden: true },
  ]);
  // Microsoft's 8-character/7-pixel-digit example encodes 8.7109375 and yields 61px.
  assert.deepEqual(geometry.columns, [
    { index: 0, widthPt: 45.75, hidden: false },
    { index: 1, widthPt: 2.25, hidden: false },
    { index: 2, widthPt: 0, hidden: true },
    { index: 3, widthPt: 48, hidden: false },
  ]);
  assert.deepEqual(
    geometry.objectBox({ cell: 'D3', dxEmu: 12700, dyEmu: 0, widthEmu: 127000, heightEmu: 254000 }),
    [49, 22.5, 10, 20]
  );
  assert.equal(geometry.objectBox({ cell: 'C3', dxEmu: 0, dyEmu: 0, widthEmu: 127000, heightEmu: 127000 }), null);
  assert.deepEqual(geometry.geometryProfile, profile);
  assert.deepEqual(bytes, before);
});
test('print profiles match the accepted Normal font', () => {
  const bytes = source(),
    view = readWorkbook(bytes),
    profile = referenceGeometryProfile(view.sheets[0].normalFont);
  assert.throws(
    () => indexedPrintGeometry(bytes, view, 'Sheet1', settings, { ...profile, font: { ...profile.font, bold: true } }),
    { code: 'layout-bad-input' }
  );
});
test('unresolved source dimensions refuse print rather than silently publishing fallback axes', () => {
  const bytes = source('<sheetFormatPr defaultRowHeight="NaN"/>'),
    view = readWorkbook(bytes);
  assert.throws(
    () => indexedPrintGeometry(bytes, view, 'Sheet1', settings),
    (error) => error.code === 'PRINT_UNRESOLVED_GEOMETRY'
  );
});
