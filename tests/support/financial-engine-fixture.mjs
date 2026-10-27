import { createWorkbook } from '../../core/workbook-package.mjs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
export const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export function fixture({
  calendar = 1900,
  formulas = [],
  mixed = false,
  rawRate = '0',
  rawCash = '1200',
  extraCells = '',
} = {}) {
  const parts = unzipSync(createWorkbook());
  parts['xl/workbook.xml'] = strToU8(
    strFromU8(parts['xl/workbook.xml'])
      .replace('<sheets>', `<workbookPr date1904="${calendar === 1904 ? '1' : '0'}"/><sheets>`)
      .replace('</workbook>', '<calcPr fullCalcOnLoad="0" forceFullCalc="0" calcMode="auto"/></workbook>')
  );
  const rows = new Map();
  const add = (row, xml) => rows.set(row, (rows.get(row) ?? '') + xml);
  add(
    1,
    `<c r="A1"><v>${rawRate}</v></c><c r="C1"><f>SUM(B1:B2)</f><v>0</v></c><c r="D1" t="inlineStr"><is><t xml:space="preserve">  opaque &lt;&amp;&gt;  </t></is></c>`
  );
  add(2, '<c r="A2"><v>12</v></c>');
  add(3, `<c r="A3"><v>${rawCash}</v></c>`);
  if (mixed) add(2, '<c r="C2"><f>DEC2HEX(255,4)</f><v>999</v></c>');
  for (const [index, formula] of formulas.entries())
    add(index + 1, `<c r="B${index + 1}"><f>${formula.replaceAll('&', '&amp;').replaceAll('<', '&lt;')}</f></c>`);
  add(40, '<c r="J40" s="0"><v>4444</v></c>' + extraCells);
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${ns}"><dimension ref="A1:J40"/><sheetFormatPr defaultRowHeight="15"/><sheetData>${[...rows]
      .sort((a, b) => a[0] - b[0])
      .map(([r, c]) => `<row r="${r}"${r === 40 ? ' ht="27" customHeight="1"' : ''}>${c}</row>`)
      .join('')}</sheetData></worksheet>`
  );
  parts['custom/financial-opaque.bin'] = new Uint8Array([0, 255, 19, 3, 17]);
  return Buffer.from(zipSync(parts, { mtime: new Date(2000, 0, 1) }));
}
