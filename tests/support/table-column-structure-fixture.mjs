import { createWorkbook, editWorkbook } from '../../core/workbook-package.mjs';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
export function tableColumnStructureFixture({
  occupied = false,
  crossing = false,
  sparse = false,
  readonly = false,
} = {}) {
  let bytes = editWorkbook(
    createWorkbook(),
    [
      ['A1', 'Value'],
      ['B1', 'Cost'],
      ['C1', 'Total'],
      ['A2', 2],
      ['B2', 10],
      ['A3', 3],
      ['B3', 20],
      ['A4', 4],
      ['B4', 30],
      ['A8', 77],
    ].map(([cell, value]) => ({ sheet: 'Sheet1', cell, value }))
  );
  bytes = editWorkbook(bytes, [
    { operation: 'createTable', sheet: 'Sheet1', range: 'A1:C4', tableName: 'Prices' },
    { operation: 'addSheet', sheet: 'Other' },
  ]);
  bytes = editWorkbook(bytes, [
    { sheet: 'Sheet1', cell: 'A1', bold: true, fontColor: '11AA44' },
    { sheet: 'Sheet1', cell: 'B2', numberFormat: 'currency' },
    ...['2', '3', '4'].map((r) => ({ sheet: 'Sheet1', cell: 'C' + r, formula: 'prices[[#This Row],[Value]]*2' })),
    { sheet: 'Sheet1', cell: 'H1', formula: 'SUM(Prices[Value])' },
    { sheet: 'Sheet1', cell: 'H2', formula: 'SUM(Prices[Cost])' },
    { sheet: 'Sheet1', cell: 'H3', formula: 'SUM(Prices[Total])' },
    { sheet: 'Sheet1', cell: 'H4', formula: 'H2+1' },
    { sheet: 'Sheet1', cell: 'H5', formula: 'SUM(A2:B4)' },
    { sheet: 'Sheet1', cell: 'H6', formula: '"Prices[Cost]"' },
    { sheet: 'Other', cell: 'A1', formula: 'SUM(PRICES[Cost])' },
    ...(occupied ? [{ sheet: 'Sheet1', cell: 'D2', value: 'Do not absorb' }] : []),
    ...(crossing ? [{ sheet: 'Sheet1', cell: 'H7', formula: 'SUM(A2:D4)' }] : []),
  ]);
  const parts = unzipSync(bytes);
  parts['customXml/table-opaque.xml'] = strToU8('<opaque> untouched </opaque>');
  if (sparse)
    parts['xl/worksheets/sheet1.xml'] = strToU8(
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Value</t></is></c><c r="B1" t="inlineStr"><is><t>Cost</t></is></c><c r="C1" t="inlineStr"><is><t>Total</t></is></c></row><row r="2" ht="20" customHeight="1"/><row r="4"><c r="B4" s="0"/></row></sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>'
    );
  if (readonly)
    parts['xl/worksheets/sheet1.xml'] = strToU8(
      strFromU8(parts['xl/worksheets/sheet1.xml']).replace('</worksheet>', '<sheetProtection sheet="1"/></worksheet>')
    );
  return Buffer.from(zipSync(parts));
}
