/** Authored records and expected retained rows; never invokes a dedup engine. */
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook, editWorkbook } from '../../core/workbook-package.mjs';

export const dedupSheet = 'Records';
export const dedupRange = 'A1:E14';
export const dedupColumns = ['A', 'B'];
export const dedupKeptRows = [2, 4, 5, 6, 7, 8, 9, 10, 12, 13];
export const dedupRemovedRows = [
  { row: 3, keptRow: 2 },
  { row: 11, keptRow: 2 },
  { row: 14, keptRow: 12 },
];
export const dedupCalcPr =
  '<calcPr calcId="191029" calcMode="auto" calcOnSave="0" fullCalcOnLoad="0" forceFullCalc="0"/>';
export const dedupRecords = [
  { row: 2, type: 'number', raw: '1', region: 'east' },
  { row: 3, type: 'number', raw: '1.00', region: 'east' },
  { row: 4, type: 'text', value: '1', region: 'east' },
  { row: 5, type: 'boolean', value: true, region: 'east' },
  { row: 6, type: 'blank', region: 'east' },
  { row: 7, type: 'text', value: '', region: 'east' },
  { row: 8, type: 'error', value: '#N/A', region: 'east' },
  { row: 9, type: 'text', value: '#N/A', region: 'east' },
  { row: 10, type: 'number', raw: '1', region: 'west' },
  { row: 11, type: 'number', raw: '1', region: 'east' },
  { row: 12, type: 'number', raw: '9007199254740993', region: 'east' },
  { row: 13, type: 'number', raw: '9007199254740994', region: 'east' },
  { row: 14, type: 'number', raw: '9007199254740993', region: 'east' },
];
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const text = (ref, value) => `<c r="${ref}" t="inlineStr"><is><t>${value}</t></is></c>`;
const number = (ref, raw) => `<c r="${ref}"><v>${raw}</v></c>`;
function key(record) {
  const ref = 'A' + record.row;
  switch (record.type) {
    case 'number':
      return number(ref, record.raw);
    case 'text':
      return text(ref, record.value);
    case 'boolean':
      return `<c r="${ref}" t="b"><v>${Number(record.value)}</v></c>`;
    case 'error':
      return `<c r="${ref}" t="e"><v>${record.value}</v></c>`;
    case 'blank':
      return '';
    default:
      throw Error('Unknown authored key type');
  }
}
export function dedupDesktopFixture() {
  const parts = unzipSync(createWorkbook());
  const body = dedupRecords
    .map((record) => {
      const r = record.row,
        id = 100 + r;
      return `<row r="${r}" ht="${18 + r}" customHeight="1"${r === 6 ? ' hidden="true"' : ''}>${key(record)}${text('B' + r, record.region)}${number('C' + r, id)}${text('D' + r, 'record-' + r)}<c r="E${r}"><f>C${r}*10+$F$20</f><v>${id * 10 + 10}</v></c>${text('F' + r, 'outside-row-' + r)}</row>`;
    })
    .join('');
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${ns}"><dimension ref="A1:F20"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="6" width="16" customWidth="1"/></cols><sheetData><row r="1">${['Key', 'Region', 'Value', 'Record ID', 'Amount', 'Outside selection'].map((value, i) => text(String.fromCharCode(65 + i) + '1', value)).join('')}</row>${body}<row r="20">${text('A20', 'outside-unchanged')}${number('F20', '10')}</row></sheetData></worksheet>`
  );
  let workbook = strFromU8(parts['xl/workbook.xml']).replace('name="Sheet1"', 'name="Records"');
  workbook = /<calcPr\b[^>]*\/>/.test(workbook)
    ? workbook.replace(/<calcPr\b[^>]*\/>/, dedupCalcPr)
    : workbook.replace('</workbook>', dedupCalcPr + '</workbook>');
  parts['xl/workbook.xml'] = strToU8(workbook);
  parts['custom/dedup-desktop-opaque.bin'] = new Uint8Array([7, 255, 0, 91]);
  const styled = unzipSync(
    editWorkbook(
      zipSync(parts, { mtime: new Date(2000, 0, 1) }),
      dedupRecords.map(({ row }) => ({
        sheet: dedupSheet,
        cell: 'D' + row,
        fontColor: '0000' + row.toString(16).padStart(2, '0'),
        bold: row % 2 === 1,
      }))
    )
  );
  styled['xl/workbook.xml'] = strToU8(strFromU8(styled['xl/workbook.xml']).replace(/<calcPr\b[^>]*\/>/, dedupCalcPr));
  let xml = strFromU8(styled['xl/worksheets/sheet1.xml']);
  for (const { row } of dedupRecords)
    xml = xml.replace(
      new RegExp(`(<c\\b[^>]*r="E${row}"[^>]*><f[^>]*>[^<]*<\\/f>)(?:<v[^>]*>[^<]*<\\/v>)?(<\\/c>)`),
      (_all, start, end) => start + `<v>${(100 + row) * 10 + 10}</v>` + end
    );
  styled['xl/worksheets/sheet1.xml'] = strToU8(xml);
  return Buffer.from(zipSync(styled, { mtime: new Date(2000, 0, 1) }));
}
export function dedupSavedExpectation() {
  return {
    sheet: dedupSheet,
    range: dedupRange,
    columns: dedupColumns,
    hasHeaders: true,
    keptRows: dedupKeptRows,
    removedRows: dedupRemovedRows,
    records: dedupRecords,
    formula: 'C{row}*10+$F$20',
    formulaRawValues: dedupKeptRows.map((row) => ((100 + row) * 10 + 10).toFixed(1)),
    rowMetadata: 'unchanged-at-original-worksheet-row',
    trailing: 'selected-cells-removed',
    dimension: { before: 'A1:F20', after: 'A1:F20' },
  };
}
