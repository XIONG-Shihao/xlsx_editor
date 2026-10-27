/** Authored expectations; never derived from the workbook evaluator. */
import { createWorkbook } from '../../core/workbook-package.mjs';
import { unzipSync, zipSync, strToU8 } from 'fflate';
const n = (value) => ({ kind: 'number', value }),
  s = (value) => ({ kind: 'text', value }),
  b = (value) => ({ kind: 'boolean', value }),
  r = (value) => ({ kind: 'reference', value });
export const commonFunctionCases = [
  ['SUMIF', [r('A1:A4'), s('sales'), r('B1:B4')], 'SUMIF(A1:A4,"sales",B1:B4)', 5],
  ['COUNTIF', [r('A1:A4'), s('S*')], 'COUNTIF(A1:A4,"S*")', 3],
  ['SUMIF', [r('B1:B4'), s('>=2.5')], 'SUMIF(B1:B4,">=2.5")', 10.75],
  ['COUNTIF', [r('B1:B4'), r('D5')], 'COUNTIF(B1:B4,D5)', 1],
  ['VLOOKUP', [n(20), r('C1:D4'), n(2)], 'VLOOKUP(20,C1:D4,2,FALSE)', 9.5],
  ['VLOOKUP', [s('sAlEs'), r('A1:B4'), n(2)], 'VLOOKUP("sAlEs",A1:B4,2,FALSE)', 1.25],
  ['VLOOKUP', [n(99), r('C1:D4'), n(2)], 'VLOOKUP(99,C1:D4,2,FALSE)', { error: '#N/A' }],
  ['XLOOKUP', [s('sales'), r('A1:A4'), r('B1:B4')], 'XLOOKUP("sales",A1:A4,B1:B4)', 1.25],
  ['XLOOKUP', [n(10), r('C1:C4'), r('D1:D4')], 'XLOOKUP(10,C1:C4,D1:D4)', 7.25],
  ['XLOOKUP', [n(999), r('C1:C4'), r('D1:D4'), s('Missing')], 'XLOOKUP(999,C1:C4,D1:D4,"Missing")', 'Missing'],
  ['XLOOKUP', [s('ErrorStringKey'), r('G1:G2'), r('F1:F2')], 'XLOOKUP("ErrorStringKey",G1:G2,F1:F2)', '#N/A'],
  ['XLOOKUP', [s('Flag'), r('G1:G2'), r('F1:F2')], 'XLOOKUP("Flag",G1:G2,F1:F2)', true],
  ['XLOOKUP', [n(2.25), r('K1:K3'), r('L1:L3')], 'XLOOKUP(2.25,K1:K3,L1:L3)', 8.25],
  ['XLOOKUP', [n(2.25), r('A6:C6'), r('A7:C7')], 'XLOOKUP(2.25,A6:C6,A7:C7)', 'middle'],
  ['XLOOKUP', [n(999), r('C1:C4'), r('D1:D4'), s('')], 'XLOOKUP(999,C1:C4,D1:D4,"")', ''],
  ['COUNTIF', [r('A1:A4'), s('Sales~*')], 'COUNTIF(A1:A4,"Sales~*")', 1],
  ['COUNTIF', [r('F1:F4'), b(true)], 'COUNTIF(F1:F4,TRUE)', 2],
  ['SUMIF', [r('F1:F4'), b(true), r('D1:D4')], 'SUMIF(F1:F4,TRUE,D1:D4)', 20.5],
  ['SUMIF', [r('A1:A4'), r('E1'), r('B1:B4')], 'SUMIF(A1:A4,E1,B1:B4)', 5],
  ['XLOOKUP', [n(999), r('C1:C4'), r('D1:D4')], 'XLOOKUP(999,C1:C4,D1:D4)', { error: '#N/A' }],
].map(([id, args, formula, value], i) => ({ cell: 'H' + (i + 1), id, args, formula, value }));
export function commonFunctionFixture() {
  const parts = unzipSync(createWorkbook()),
    rows = new Map(),
    values = {
      A1: 'Sales',
      A2: 'Ops',
      A3: 'sales',
      A4: 'Sales*',
      B1: 1.25,
      B2: 2.5,
      B3: 3.75,
      B4: 4.5,
      C1: 10,
      C2: 20,
      C3: 10,
      C4: 30,
      D1: 7.25,
      D2: 9.5,
      D3: 11,
      D4: 15,
      D5: 2.5,
      E1: 'Sales',
      F1: '#N/A',
      F2: true,
      F3: true,
      F4: false,
      G1: 'ErrorStringKey',
      G2: 'Flag',
      K1: 1.5,
      K2: 2.25,
      K3: 3.25,
      L1: 7.75,
      L2: 8.25,
      L3: 9.25,
      A6: 1.5,
      B6: 2.25,
      C6: 3.25,
      A7: 'first',
      B7: 'middle',
      C7: 'last',
    };
  const escape = (text) => text.replace(/[&<>]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[ch]);
  for (const [ref, value] of Object.entries(values)) {
    const row = Number(ref.match(/\d+$/)[0]),
      cells = rows.get(row) ?? [];
    rows.set(row, cells);
    cells.push(
      typeof value === 'string'
        ? '<c r="' + ref + '" t="inlineStr"><is><t>' + escape(value) + '</t></is></c>'
        : typeof value === 'boolean'
          ? '<c r="' + ref + '" t="b"><v>' + (value ? 1 : 0) + '</v></c>'
          : '<c r="' + ref + '"><v>' + value + '</v></c>'
    );
  }
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:L20"/><sheetData>' +
      [...rows]
        .sort((a, b) => a[0] - b[0])
        .map(([row, cells]) => '<row r="' + row + '">' + cells.join('') + '</row>')
        .join('') +
      '</sheetData></worksheet>'
  );
  parts['custom/common-functions.bin'] = Uint8Array.from([0, 18, 255, 9]);
  return zipSync(parts);
}
