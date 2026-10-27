import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook } from '../../core/workbook-package.mjs';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export function preservedFixture(kind = 'shared') {
  const parts = unzipSync(createWorkbook());
  const formula =
    kind === 'shared'
      // An attribute the editor does not understand keeps the workbook in preserve mode.
      ? '<f t="shared" si="0" ref="A1:A2" ca="1">B1+1</f><v>2</v>'
      : kind === 'cache'
        ? '<f>B1+1</f><v xmlns:u="urn:keep" u:tag="cached">2</v>'
        : '<f>B1+1</f><v>2</v>';
  const second = kind === 'shared' ? '<c r="A2"><f t="shared" si="1" ref="A2:A3">B2+2</f><v>3</v></c>' : '';
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${NS}"><dimension ref="A1:D2"/><sheetData><row r="1"><c r="A1">${formula}</c><c r="B1"><v>1</v></c><c r="C1" t="inlineStr"><is><r><rPr><b/></rPr><t>Keep rich</t></r></is></c><c r="D1"><v>9</v></c></row><row r="2">${second}<c r="B2"><v>1</v></c></row></sheetData></worksheet>`
  );
  if (kind === 'iterative' || kind === 'opaque')
    parts['xl/workbook.xml'] = strToU8(
      strFromU8(parts['xl/workbook.xml']).replace(
        '</workbook>',
        kind === 'iterative'
          ? '<calcPr iterate="1" iterateCount="23" iterateDelta="0.004" fullCalcOnLoad="0"/></workbook>'
          : '<calcPr xmlns:u="urn:keep" u:flag="preserve"><u:child value="&gt;"/></calcPr></workbook>'
      )
    );
  parts['customXml/opaque.bin'] = new Uint8Array([0, 255, 7, 8]);
  return Buffer.from(zipSync(parts));
}
