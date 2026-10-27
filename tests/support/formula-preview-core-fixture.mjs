/** Independently authored preview inputs/results; source caches deliberately differ. */
import { createWorkbook } from '../../core/workbook-package.mjs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
export function formulaPreviewCoreFixture({ manual = true, special = false } = {}) {
  const parts = unzipSync(createWorkbook()),
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  parts['xl/workbook.xml'] = strToU8(
    `<workbook xmlns="${ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId3"/></sheets><definedNames><definedName name="Amount">'Sheet1'!$A$1:$A$3</definedName><definedName name="Tax">'Other'!$A$1</definedName><definedName name="Tax" localSheetId="0">'Sheet1'!$A$2</definedName><definedName name="BadScope" localSheetId="9">'Sheet1'!$A$1</definedName><definedName name="Dup">'Sheet1'!$A$1</definedName><definedName name="dup">'Sheet1'!$A$2</definedName></definedNames><calcPr calcMode="${manual ? 'manual' : 'auto'}" calcOnSave="0"/></workbook>`
  );
  parts['xl/_rels/workbook.xml.rels'] = strToU8(
    strFromU8(parts['xl/_rels/workbook.xml.rels']).replace(
      '</Relationships>',
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>'
    )
  );
  parts['[Content_Types].xml'] = strToU8(
    strFromU8(parts['[Content_Types].xml']).replace(
      '</Types>',
      '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
    )
  );
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${ns}"><sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1+3</f><v>103</v></c><c r="C1" t="inlineStr"><is><t>#N/A</t></is></c><c r="D1" t="e"><v>#N/A</v></c><c r="F1"><f>F2+1</f><v>91</v></c><c r="G1"><v>9007199254740993</v></c>${special ? '<c r="K1"><f t="array" ref="K1:L2">SEQUENCE(2,2)</f><v>1</v></c><c r="L1"><v>2</v></c>' : ''}</row><row r="2"><c r="A2"><v>4</v></c><c r="B2"><f>B1*2</f><v>105</v></c><c r="C2" t="b"><v>0</v></c><c r="F2"><f>F1+1</f><v>92</v></c>${special ? '<c r="K2"><v>3</v></c><c r="L2"><v>4</v></c>' : ''}</row><row r="3"><c r="A3"><v>6</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="M1:M2"/></mergeCells></worksheet>`
  );
  parts['xl/worksheets/sheet2.xml'] = strToU8(
    `<worksheet xmlns="${ns}"><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>`
  );
  parts['custom/unchanged-preview.bin'] = new Uint8Array([0, 255, 42]);
  return zipSync(parts, { mtime: new Date(2000, 0, 1) });
}
export const formulaPreviewCoreCases = Object.freeze([
  ['SUM(A1:A3)', { status: 'value', type: 'number', value: 12 }],
  ['B2+1', { status: 'value', type: 'number', value: 11 }],
  ['SUM(Amount)+Tax', { status: 'value', type: 'number', value: 16 }],
  ['Other!A1+1', { status: 'value', type: 'number', value: 8 }],
  ['C1', { status: 'value', type: 'text', value: '#N/A' }],
  ['C2', { status: 'value', type: 'boolean', value: false }],
  ['D1', { status: 'value', type: 'error', value: '#N/A' }],
  ['1/0', { status: 'value', type: 'error', value: '#DIV/0!' }],
  ['IFERROR(D1,99)', { status: 'value', type: 'number', value: 99 }],
  ['"a"&"b"', { status: 'value', type: 'text', value: 'ab' }],
  ['-0', { status: 'value', type: 'number', value: -0 }],
  ['PMT(0,12,1200)', { status: 'value', type: 'number', value: -100 }],
]);
