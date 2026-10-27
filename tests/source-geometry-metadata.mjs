import assert from 'node:assert/strict';
import { test } from 'node:test';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { createWorkbook, readWorkbook, readWorkbookMetadata } from '../core/workbook-package.mjs';

function fixture(format, rows = '', normal = false) {
  const parts = unzipSync(createWorkbook());
  let sheet = strFromU8(parts['xl/worksheets/sheet1.xml']);
  sheet = sheet.replace('<sheetData/>', `${format}<sheetData>${rows}</sheetData>`);
  parts['xl/worksheets/sheet1.xml'] = strToU8(sheet);
  if (normal) {
    let styles = strFromU8(parts['xl/styles.xml']);
    styles = styles
      .replace('<fonts count="1">', '<fonts count="2">')
      .replace('</fonts>', '<font><sz val="12"/><name val="Arial"/><b/><i val="false"/></font></fonts>');
    styles = styles
      .replace('<cellStyleXfs count="1">', '<cellStyleXfs count="2">')
      .replace(
        '</cellStyleXfs>',
        '<xf fontId="1"/></cellStyleXfs><cellStyles count="1"><cellStyle name="Normal" builtinId="0" xfId="1"/></cellStyles>'
      );
    parts['xl/styles.xml'] = strToU8(styles);
  }
  return zipSync(parts);
}

test('source defaults and Normal style remain identical in full and cell-free metadata views', () => {
  const bytes = fixture(
    '<sheetFormatPr defaultRowHeight="22.5" defaultColWidth="9.140625" baseColWidth="11" zeroHeight="true"/>',
    '<row r="1" hidden="false"/><row r="2" hidden="1"/><row r="3" hidden="0" ht="30"/>',
    true
  );
  const before = bytes.slice(),
    sheet = readWorkbook(bytes).sheets[0];
  assert.deepEqual(sheet.sheetFormat, {
    defaultRowHeight: 22.5,
    defaultColumnWidth: 9.140625,
    baseColumnWidth: 11,
    rowsHiddenByDefault: true,
  });
  assert.deepEqual(sheet.normalFont, { fontFamily: 'Arial', fontSize: 12, bold: true, italic: false });
  assert.deepEqual(sheet.visibleRows, [0, 2]);
  assert.deepEqual(sheet.hiddenRows, [1]);
  assert.deepEqual(sheet.rowHeights, [{ row: 2, height: 30 }]);
  const parts = unzipSync(bytes);
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    strFromU8(parts['xl/worksheets/sheet1.xml']).replace(/<sheetData>[\s\S]*?<\/sheetData>/, '<sheetData/>')
  );
  const indexed = readWorkbookMetadata(zipSync(parts)).view.sheets[0];
  assert.deepEqual(indexed.sheetFormat, sheet.sheetFormat);
  assert.deepEqual(indexed.normalFont, sheet.normalFont);
  assert.deepEqual(bytes, before);
});
test('absent defaults stay absent; source zero sizes and explicit false visibility stay meaningful', () => {
  const blank = readWorkbook(createWorkbook()).sheets[0];
  assert.deepEqual(blank.sheetFormat, { rowsHiddenByDefault: false });
  assert.deepEqual(blank.normalFont, { fontFamily: 'Calibri', fontSize: 11, bold: false, italic: false });
  const sheet = readWorkbook(
    fixture('<sheetFormatPr defaultRowHeight="0" defaultColWidth="0" baseColWidth="0" zeroHeight="0"/>')
  ).sheets[0];
  assert.deepEqual(sheet.sheetFormat, {
    defaultRowHeight: 0,
    defaultColumnWidth: 0,
    baseColumnWidth: 0,
    rowsHiddenByDefault: false,
  });
});
test('invalid numeric or visibility defaults remain explicitly unresolved instead of becoming accepted sizes', () => {
  const sheet = readWorkbook(
    fixture('<sheetFormatPr defaultRowHeight="NaN" defaultColWidth="0x10" baseColWidth="1.5" zeroHeight="maybe"/>')
  ).sheets[0];
  assert.deepEqual(sheet.sheetFormat, { rowsHiddenByDefault: false });
  assert.equal(sheet.geometryIssues.length, 4);
});
test('false bold and italic source flags do not change the Normal font metrics', () => {
  let parts = unzipSync(createWorkbook());
  parts['xl/styles.xml'] = strToU8(
    strFromU8(parts['xl/styles.xml']).replace('</font>', '<b val="0"/><i val="false"/></font>')
  );
  assert.deepEqual(readWorkbook(zipSync(parts)).sheets[0].normalFont, {
    fontFamily: 'Calibri',
    fontSize: 11,
    bold: false,
    italic: false,
  });
});
test('an unresolved Normal font stays readable with an explicit geometry issue and unchanged source', () => {
  const parts = unzipSync(createWorkbook());
  parts['xl/styles.xml'] = strToU8(
    strFromU8(parts['xl/styles.xml']).replace('val="Calibri"', 'val=""').replace('val="11"', 'val="1e999"')
  );
  const bytes = zipSync(parts),
    before = bytes.slice(),
    sheet = readWorkbook(bytes).sheets[0];
  assert.equal(sheet.normalFont.fontFamily, 'Calibri');
  assert.equal(sheet.normalFont.fontSize, 11);
  assert.equal(sheet.geometryIssues.length, 2);
  assert.deepEqual(bytes, before);
});
