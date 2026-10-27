import { createWorkbook, readWorkbookMetadata } from '../../core/workbook-package.mjs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
export const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export function accessibilityFixture({
  protectedSheet = false,
  extended = false,
  formula = false,
  headerless = false,
  orphan = false,
  twoSheets = true,
} = {}) {
  const parts = unzipSync(createWorkbook()),
    count = twoSheets ? 2 : 1;
  if (twoSheets) {
    parts['xl/workbook.xml'] = strToU8(
      strFromU8(parts['xl/workbook.xml']).replace(
        '</sheets>',
        '<sheet name="Details" sheetId="2" r:id="accessibilitySheet2"/></sheets>'
      )
    );
    parts['xl/_rels/workbook.xml.rels'] = strToU8(
      strFromU8(parts['xl/_rels/workbook.xml.rels']).replace(
        '</Relationships>',
        `<Relationship Id="accessibilitySheet2" Type="${R}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`
      )
    );
    parts['[Content_Types].xml'] = strToU8(
      strFromU8(parts['[Content_Types].xml']).replace(
        '</Types>',
        '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
      )
    );
  }
  for (let i = 1; i <= count; i++) {
    parts[`xl/worksheets/sheet${i}.xml`] = strToU8(
      `<worksheet xmlns="${S}" xmlns:r="${R}"><dimension ref="A1:C3"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${i === 1 ? 'Label' : 'Good'}</t></is></c><c r="B1"${formula && i === 1 ? ' t="str"' : ' t="inlineStr"'}>${formula && i === 1 ? '<f>"LABEL"</f><v>LABEL</v>' : `<is><t>${i === 1 ? ' label ' : 'Distinct'}</t></is>`}</c><c r="C1" t="inlineStr"><is><t>${i === 1 ? '   ' : 'Third'}</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>One</t></is></c><c r="B2"><v>7</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Two</t></is></c><c r="B3"><v>8</v></c></row></sheetData>${protectedSheet && i === 1 ? '<sheetProtection sheet="1" objects="1"/>' : ''}${orphan && i === 1 ? '' : `<tableParts count="1"><tablePart r:id="table${i}"/></tableParts>`}</worksheet>`
    );
    parts[`xl/worksheets/_rels/sheet${i}.xml.rels`] = strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="table${i}" Type="${R}/table" Target="../tables/table${i}.xml"/></Relationships>`
    );
    parts[`xl/tables/table${i}.xml`] = strToU8(
      `<table xmlns="${S}" id="${i}" name="Data${i}" displayName="Data${i}" ref="A1:C3"${headerless && i === 1 ? ' headerRowCount="0"' : ''}><autoFilter ref="A1:C3"/><tableColumns count="3"><tableColumn id="1" name="First"/><tableColumn id="2" name="Second"/><tableColumn id="3" name="Third"/></tableColumns>${extended && i === 1 ? '<extLst><ext uri="foreign"/></extLst>' : ''}</table>`
    );
    parts['[Content_Types].xml'] = strToU8(
      strFromU8(parts['[Content_Types].xml']).replace(
        '</Types>',
        `<Override PartName="/xl/tables/table${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>`
      )
    );
  }
  parts['custom/accessibility-opaque.bin'] = new Uint8Array([3, 5, 7, 255]);
  const bytes = zipSync(parts);
  return { bytes, view: readWorkbookMetadata(bytes).view };
}
