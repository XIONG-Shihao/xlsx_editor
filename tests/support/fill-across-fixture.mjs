/** Authored ordinary cell-copy facts; this is not an observed Excel reference. */
import { createWorkbook } from '../../core/workbook-package.mjs';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const esc = (v) =>
  String(v).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[c]);
export const acrossNames = ['Template', 'January', 'February', 'Untouched'];
export const acrossRange = 'A1:C4';
export const acrossSource = {
  A1: { value: 7, kind: 'number', style: 2 },
  A2: { value: '雪 😀', kind: 'text', style: 0 },
  A3: { value: false, kind: 'boolean', style: 1 },
  A4: { value: '#N/A', kind: 'error', style: 2 },
  B1: { value: '#N/A', kind: 'text', style: 0 },
  B2: { value: '', kind: 'text', style: 1 },
  B3: { value: null, kind: 'blank', style: 2 },
  C1: { formula: 'F1+1', cache: 11, style: 1 },
  C2: { formula: 'Template!F1+2', cache: 12, style: 2 },
  C3: { formula: 'IFERROR(#REF!,99)', cache: 99, style: 0 },
};
function cell(ref, c) {
  return (
    '<c r="' +
    ref +
    '"' +
    (c.style ? ' s="' + c.style + '"' : '') +
    (c.formula ? '' : ({ text: ' t="inlineStr"', boolean: ' t="b"', error: ' t="e"' }[c.kind] ?? '')) +
    '>' +
    (c.formula
      ? '<f>' + esc(c.formula) + '</f><v>' + c.cache + '</v>'
      : c.kind === 'blank'
        ? ''
        : c.kind === 'text'
          ? '<is><t xml:space="preserve">' + esc(c.value) + '</t></is>'
          : '<v>' + esc(c.kind === 'boolean' ? (c.value ? '1' : '0') : c.value) + '</v>') +
    '</c>'
  );
}
function sheet(cells, extra = '') {
  const rows = new Map();
  for (const [ref, c] of Object.entries(cells)) {
    const row = Number(ref.match(/\d+$/)[0]);
    if (!rows.has(row)) rows.set(row, []);
    rows.get(row).push(cell(ref, c));
  }
  return (
    '<worksheet xmlns="' +
    ns +
    '" xmlns:r="' +
    rel +
    '"><dimension ref="A1:F4"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData>' +
    [...rows]
      .sort(([a], [b]) => a - b)
      .map(([row, cells]) => '<row r="' + row + '">' + cells.join('') + '</row>')
      .join('') +
    '</sheetData>' +
    extra +
    '</worksheet>'
  );
}
export function fillAcrossFixture({ calculation = 'auto', boundary = null } = {}) {
  if (
    !['auto', 'manual'].includes(calculation) ||
    ![null, 'merge', 'hyperlink', 'validation', 'axis'].includes(boundary)
  )
    throw Error('Closed Across fixture options');
  const parts = unzipSync(createWorkbook()),
    all = [{ ...acrossSource, F1: { value: 10, kind: 'number', style: 0 } }];
  for (let i = 1; i < 4; i++) {
    const cells = {};
    for (const column of ['A', 'B', 'C'])
      for (let row = 1; row <= 4; row++) cells[column + row] = { value: 900 + i, kind: 'number', style: 1 };
    cells.C1 = { formula: 'F1*2', cache: calculation === 'manual' ? 710 + i : 20 * (i + 1), style: 2 };
    cells.F1 = { value: 10 * (i + 1), kind: 'number', style: 0 };
    all.push(cells);
  }
  all.forEach((cells, i) => {
    let extra =
      i === 2
        ? boundary === 'merge'
          ? '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>'
          : boundary === 'hyperlink'
            ? '<hyperlinks><hyperlink ref="A1" location="Template!A1"/></hyperlinks>'
            : boundary === 'validation'
              ? '<dataValidations count="1"><dataValidation type="whole" sqref="A1"><formula1>0</formula1><formula2>20</formula2></dataValidation></dataValidations>'
              : ''
        : '';
    let xml = sheet(cells, extra);
    if (i === 2 && boundary === 'axis')
      xml = xml.replace('<sheetData>', '<cols><col min="1" max="1" style="2"/></cols><sheetData>');
    parts['xl/worksheets/sheet' + (i + 1) + '.xml'] = strToU8(xml);
  });
  parts['xl/workbook.xml'] = strToU8(
    strFromU8(parts['xl/workbook.xml'])
      .replace(
        /<sheets>[\s\S]*?<\/sheets>/,
        '<sheets>' +
          acrossNames
            .map((name, i) => '<sheet name="' + name + '" sheetId="' + (i + 1) + '" r:id="sheet' + (i + 1) + '"/>')
            .join('') +
          '</sheets>'
      )
      .replace(
        '</workbook>',
        '<calcPr calcId="191029" calcMode="' +
          calculation +
          '" calcOnSave="0" fullCalcOnLoad="0" forceFullCalc="0"/></workbook>'
      )
  );
  parts['xl/_rels/workbook.xml.rels'] = strToU8(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      acrossNames
        .map(
          (name, i) =>
            '<Relationship Id="sheet' +
            (i + 1) +
            '" Type="' +
            rel +
            '/worksheet" Target="worksheets/sheet' +
            (i + 1) +
            '.xml"/>'
        )
        .join('') +
      '<Relationship Id="styles" Type="' +
      rel +
      '/styles" Target="styles.xml"/></Relationships>'
  );
  parts['[Content_Types].xml'] = strToU8(
    strFromU8(parts['[Content_Types].xml']).replace(
      '</Types>',
      [2, 3, 4]
        .map(
          (i) =>
            '<Override PartName="/xl/worksheets/sheet' +
            i +
            '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
        .join('') + '</Types>'
    )
  );
  parts['xl/styles.xml'] = strToU8(
    '<styleSheet xmlns="' +
      ns +
      '"><numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.0000;[Red](&quot;$&quot;#,##0.0000)"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="14"/><color rgb="FF654321"/><name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="2" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>'
  );
  parts['custom/across-opaque.bin'] = new Uint8Array([0, 255, 17, 29]);
  return Buffer.from(zipSync(parts, { level: 0, mtime: new Date('2020-01-01T00:00:00Z') }));
}
export const fillAcrossCommands = (mode) =>
  ['January', 'February'].map((sheet) => ({
    type: 'fillAcrossSheets',
    sheet,
    sourceSheet: 'Template',
    range: acrossRange,
    mode,
  }));
