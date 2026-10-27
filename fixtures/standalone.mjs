import {zipSync,unzipSync,strToU8} from 'fflate';
import {createWorkbook} from '../core/workbook-package.mjs';
export function workbook() {
  const parts=unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml']=strToU8('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>10</v></c><c r="B1"><v>20</v></c><c r="C1"><f>A1+B1</f><v>30</v></c></row></sheetData></worksheet>');
  parts['custom/preserved.bin']=new Uint8Array([0,1,2,255]);return zipSync(parts);
}
