/** Independent static name/scope cases; expected values do not use the evaluator. */
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook } from '../../core/workbook-package.mjs';
const n = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  p = 'http://schemas.openxmlformats.org/package/2006/relationships';
const escape = (v) => String(v).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('"', '&quot;');
export const staticNameCases = Object.freeze([
  { sheet: 'Inputs', cell: 'C1', formula: 'RaTe', value: 2 },
  { sheet: 'Report', cell: 'C1', formula: 'RATE', value: 5 },
  { sheet: 'Report', cell: 'C2', formula: 'SUM(Total)', value: 7 },
  { sheet: 'Report', cell: 'C3', formula: 'QuotedValue', value: 7 },
  { sheet: "O'Brien Data", cell: 'C1', formula: 'LocalQuote', value: 11 },
  { sheet: "O'Brien Data", cell: 'C2', formula: "'O''Brien Data'!LocalQuote", value: 11 },
  { sheet: 'Other', cell: 'C1', formula: 'OtherOnly', value: 19 },
  { sheet: 'Report', cell: 'C4', formula: "'Other'!OtherOnly", value: 19 },
  { sheet: 'Inputs', cell: 'C2', formula: 'LeadingZero', value: 13 },
  { sheet: 'Report', cell: 'C5', formula: 'LeadingOne', value: 17 },
]);
export const invalidScopeIds = Object.freeze(['', '99', 'bad', '-1', '1.5', '4294967296']);
const definition = (name, formula, scope) =>
  `<definedName name="${name}"${scope === undefined ? '' : ` localSheetId="${escape(scope)}"`}>${escape(formula)}</definedName>`;
export function nameScopeOwnerFixture({ kind = 'valid' } = {}) {
  if (!['valid', 'invalid', 'duplicate'].includes(kind)) throw Error('Invalid name fixture kind');
  const parts = unzipSync(createWorkbook()),
    sheets = ['Inputs', 'Report', "O'Brien Data", 'Other'];
  const names =
    kind === 'valid'
      ? [
          definition('Rate', 'Inputs!$A$1'),
          definition('rate', 'Inputs!$A$2', '1'),
          definition('Total', 'Inputs!$A$1:$A$2'),
          definition('QuotedValue', "'O''Brien Data'!$A$1"),
          definition('LocalQuote', "'O''Brien Data'!$A$2", '2'),
          definition('OtherOnly', 'Other!$A$1', '3'),
          definition('LeadingZero', 'Inputs!$A$3', '00'),
          definition('LeadingOne', 'Inputs!$A$4', '01'),
        ]
      : kind === 'invalid'
        ? invalidScopeIds.map((id, i) => definition('BadScope' + i, 'Inputs!$A$3', id))
        : [
            definition('Dupe', 'Inputs!$A$1'),
            definition('dupe', 'Inputs!$A$2'),
            definition('LocalDupe', 'Inputs!$A$1', '1'),
            definition('localdupe', 'Inputs!$A$2', '1'),
          ];
  parts['xl/workbook.xml'] = strToU8(
    `<workbook xmlns="${n}" xmlns:r="${r}"><sheets>${sheets.map((name, i) => `<sheet name="${escape(name)}" sheetId="${i + 1}" r:id="sheet${i + 1}"/>`).join('')}</sheets><definedNames>${names.join('')}</definedNames><calcPr calcId="191029" calcMode="auto" calcOnSave="1" fullCalcOnLoad="0" forceFullCalc="0"/></workbook>`
  );
  parts['xl/_rels/workbook.xml.rels'] = strToU8(
    `<Relationships xmlns="${p}">${sheets.map((_, i) => `<Relationship Id="sheet${i + 1}" Type="${r}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="styles" Type="${r}/styles" Target="styles.xml"/></Relationships>`
  );
  parts['[Content_Types].xml'] = strToU8(
    strFromU8(parts['[Content_Types].xml']).replace(
      '</Types>',
      [2, 3, 4]
        .map(
          (i) =>
            `<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
        )
        .join('') + '</Types>'
    )
  );
  const literals = [[2, 5, 13, 17], [23], [7, 11], [19]];
  sheets.forEach((_, i) => {
    const rows = Array.from({ length: 6 }, (_, j) => {
      const row = j + 1,
        values = [];
      if (literals[i][j] !== undefined) values.push(`<c r="A${row}"><v>${literals[i][j]}</v></c>`);
      values.push(`<c r="C${row}" s="0"/>`);
      if (kind === 'invalid' && i === 1) {
        values.push(`<c r="B${row}"><f>BadScope${j}*2</f><v>999</v></c>`);
        values.push(`<c r="D${row}"><f>IFERROR(B${row},99)</f><v>99</v></c>`);
      }
      if (kind === 'duplicate' && i === 1 && row <= 2)
        values.push(`<c r="B${row}"><f>${row === 1 ? 'Dupe' : 'LocalDupe'}</f><v>999</v></c>`);
      if (row === 6)
        values.push(
          '<c r="F6" t="inlineStr"><is><t xml:space="preserve">  00123 =SUM(A1) &lt;&amp;&gt;  </t></is></c>'
        );
      return `<row r="${row}"${row === 6 ? ' ht="24" customHeight="1"' : ''}>${values.join('')}</row>`;
    });
    parts[`xl/worksheets/sheet${i + 1}.xml`] = strToU8(
      `<worksheet xmlns="${n}"><dimension ref="A1:F6"/><sheetFormatPr defaultRowHeight="15"/><sheetData>${rows.join('')}</sheetData></worksheet>`
    );
  });
  parts['custom/name-scope-preserved.bin'] = new Uint8Array([0, 255, 81, 9, 21, 34]);
  return Buffer.from(zipSync(parts, { mtime: new Date(2000, 0, 1) }));
}
