/** Authored independent inputs for the bounded scalar picker profile. */
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook } from '../../core/workbook-package.mjs';

const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const functionLibraryCases = Object.freeze([
  { cell: 'C1', id: 'SUM', args: [['reference', 'A1:A3']], formula: 'SUM(A1:A3)', type: 'number', value: 9 },
  {
    cell: 'C2',
    id: 'IF',
    args: [
      ['boolean', true],
      ['text', '00123 "quoted" <&>'],
      ['boolean', false],
    ],
    formula: 'IF(TRUE,"00123 ""quoted"" <&>",FALSE)',
    type: 'string',
    value: '00123 "quoted" <&>',
  },
  {
    cell: 'C3',
    id: 'IF',
    args: [
      ['boolean', false],
      ['number', 17],
      ['boolean', true],
    ],
    formula: 'IF(FALSE,17,TRUE)',
    type: 'boolean',
    value: true,
  },
  {
    cell: 'C4',
    id: 'IF',
    args: [
      ['boolean', true],
      ['reference', 'A4'],
      ['boolean', false],
    ],
    formula: 'IF(TRUE,A4,FALSE)',
    type: 'error',
    value: '#DIV/0!',
  },
  {
    cell: 'C5',
    id: 'IFERROR',
    args: [
      ['reference', 'A4'],
      ['boolean', true],
    ],
    formula: 'IFERROR(A4,TRUE)',
    type: 'boolean',
    value: true,
  },
  {
    cell: 'C6',
    id: 'IFERROR',
    args: [
      ['reference', 'A5'],
      ['text', 'fallback'],
    ],
    formula: 'IFERROR(A5,"fallback")',
    type: 'string',
    value: '#DIV/0!',
  },
  { cell: 'C7', id: 'DEC2HEX', args: [['number', 255]], formula: 'DEC2HEX(255)', type: 'string', value: 'FF' },
  {
    cell: 'C8',
    id: 'DATE',
    args: [
      ['number', 2000],
      ['number', 2],
      ['number', 29],
    ],
    formula: 'DATE(2000,2,29)',
    type: 'date',
    date: [2000, 2, 29],
  },
]);
export const functionLibraryOwnerCases = Object.freeze([
  ...functionLibraryCases,
  { cell: 'D2', id: 'AVERAGE', formula: 'AVERAGE(2,4)', type: 'number', value: 3 },
  { cell: 'D3', id: 'COUNT', formula: 'COUNT(H1:H5)', type: 'number', value: 2 },
  { cell: 'D4', id: 'MIN', formula: 'MIN(2,3,4)', type: 'number', value: 2 },
  { cell: 'D5', id: 'MAX', formula: 'MAX(2,3,4)', type: 'number', value: 4 },
  { cell: 'D6', id: 'LEN', formula: 'LEN("東京<&>")', type: 'string-length', value: 5 },
  { cell: 'D7', id: 'LEFT', formula: 'LEFT("東京abc")', type: 'string', value: '東' },
  { cell: 'D8', id: 'TRIM', formula: 'TRIM("  a   b  ")', type: 'string', value: 'a b' },
  { cell: 'D9', id: 'YEAR', formula: 'YEAR(C8)', type: 'number', value: 2000 },
  { cell: 'D10', id: 'MONTH', formula: 'MONTH(C8)', type: 'number', value: 2 },
  { cell: 'D11', id: 'DAY', formula: 'DAY(C8)', type: 'number', value: 29 },
  { cell: 'D12', id: 'ABS', formula: 'ABS(-5)', type: 'number', value: 5 },
  { cell: 'D13', id: 'ROUND', formula: 'ROUND(4.5,1)', type: 'number', value: 4.5 },
  { cell: 'D14', id: 'ROWS', formula: 'ROWS(A1:B3)', type: 'number', value: 3 },
  { cell: 'D15', id: 'COLUMNS', formula: 'COLUMNS(A1:B3)', type: 'number', value: 2 },
]);
export function functionLibraryFixture({ calendar = 1900, protectedArray = false } = {}) {
  if (![1900, 1904].includes(calendar)) throw Error('Unknown fixture calendar');
  const parts = unzipSync(createWorkbook());
  parts['xl/workbook.xml'] = strToU8(
    strFromU8(parts['xl/workbook.xml'])
      .replace('<sheets>', `<workbookPr date1904="${calendar === 1904 ? '1' : '0'}"/><sheets>`)
      .replace('</workbook>', '<calcPr fullCalcOnLoad="0" forceFullCalc="0" calcMode="auto"/></workbook>')
  );
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${ns}"><dimension ref="A1:Z900"/><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>SUM(A1:A3)</f><v>9</v></c><c r="D1" t="inlineStr"><is><t>merged &lt;&amp;&gt;</t></is></c>${protectedArray ? '<c r="F1"><f t="array" ref="F1:F2">1+1</f><v>2</v></c>' : ''}<c r="H1"><v>2</v></c></row><row r="2"><c r="A2"><v>3</v></c>${protectedArray ? '<c r="F2"><v>2</v></c>' : ''}<c r="H2"><v>4</v></c></row><row r="3"><c r="A3"><v>4</v></c><c r="H3" t="inlineStr"><is><t>4</t></is></c></row><row r="4"><c r="A4" t="e"><v>#DIV/0!</v></c><c r="H4" t="b"><v>1</v></c></row><row r="5"><c r="A5" t="inlineStr"><is><t>#DIV/0!</t></is></c></row><row r="33"><c r="C33" s="0"><v>933</v></c></row><row r="900" ht="27" customHeight="1"><c r="Z900"><v>4444</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="D1:E1"/></mergeCells></worksheet>`
  );
  parts['custom/function-library-opaque.bin'] = new Uint8Array([0, 255, 19, 3, 17]);
  return zipSync(parts, { mtime: new Date(2000, 0, 1) });
}
export function functionLibraryExpected(
  calendar,
  { replacement = false, onlyDate = false, ownerCatalog = false } = {}
) {
  const cases = (
    onlyDate
      ? functionLibraryCases.filter((c) => c.id === 'DATE')
      : ownerCatalog
        ? functionLibraryOwnerCases
        : functionLibraryCases
  ).map((c) => ({
    cell: c.cell,
    formula: c.formula,
    type: c.type === 'string-length' ? 'number' : c.type,
    ...(c.type === 'date' ? { date: c.date } : { value: c.value }),
  }));
  if (replacement) cases.push({ cell: 'C33', formula: 'SUM(A1:A3)', type: 'number', value: 9 });
  return {
    schema: 'workbook.xlsx.function-library-expected.v1',
    calendar,
    sheet: 'xl/worksheets/sheet1.xml',
    cases,
    scope:
      'Authored scalar inputs/expected types; DATE serial computed independently by Python Gregorian arithmetic; no output-derived expectations',
  };
}
