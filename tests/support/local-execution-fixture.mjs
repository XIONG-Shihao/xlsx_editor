import { createWorkbook } from '../../core/workbook-package.mjs';
import { unzipSync, zipSync, strToU8 } from 'fflate';
const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export function localExecutionFixture({
  defaultHidden = false,
  filter = false,
  table = false,
  protectedSheet = false,
  manual = false,
} = {}) {
  const parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<s:worksheet xmlns:s="${S}" xmlns:x="urn:test"><s:sheetFormatPr defaultRowHeight="22" defaultColWidth="12"${defaultHidden ? ' zeroHeight="1"' : ''}/><s:cols><s:col min="2" max="2" width="18.5" hidden="true" customWidth="1"/><s:col min="4" max="4" width="0" customWidth="1"/></s:cols><s:sheetData><s:row r="1" hidden="0"><s:c r="A1"><s:v>42</s:v></s:c><s:c r="B1"><s:f>A1+2</s:f><s:v>7</s:v></s:c></s:row><s:row r="2" hidden="1" ht="37.5" customHeight="1" x:note="A &amp; B"><s:c r="A2" t="inlineStr"><s:is><s:t>Hidden &amp; literal</s:t></s:is></s:c></s:row><s:row r="3" ht="0" customHeight="1"/><s:row r="700000" hidden="true" s="0"/></s:sheetData>${filter ? '<s:autoFilter ref="A1:B3"/>' : ''}${table ? '<s:tableParts count="0"/>' : ''}${protectedSheet ? '<s:sheetProtection sheet="1"/>' : ''}</s:worksheet>`
  );
  parts['custom/visibility-opaque.bin'] = new Uint8Array([0, 7, 42, 255]);
  if (manual)
    parts['xl/workbook.xml'] = strToU8(
      new TextDecoder()
        .decode(parts['xl/workbook.xml'])
        .replace('</workbook>', '<calcPr calcMode="manual" calcOnSave="0"/></workbook>')
    );
  return zipSync(parts);
}
