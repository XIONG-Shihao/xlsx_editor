/** Authored draft-preview source and expectations; no evaluator supplies results. */
import { zipSync, unzipSync, strFromU8, strToU8 } from 'fflate';
import { createWorkbook } from '../../core/workbook-package.mjs';
const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const P = 'http://schemas.openxmlformats.org/package/2006/relationships';
const esc = (value) =>
  String(value).replace(
    /[&<>"']/g,
    (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[c]
  );
const number = (cell, value) => `<c r="${cell}"><v>${value}</v></c>`;
const text = (cell, value) => `<c r="${cell}" t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`;
const formula = (cell, value, cache = 99, extra = '') =>
  `<c r="${cell}"><f${extra}>${esc(value)}</f><v>${cache}</v></c>`;
function sheet(rows, dimension = 'A1:H12', tail = '') {
  return `<worksheet xmlns="${S}" xmlns:r="${R}"><dimension ref="${dimension}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData>${rows.map(([row, cells]) => `<row r="${row}">${cells.join('')}</row>`).join('')}</sheetData>${tail}</worksheet>`;
}
export const formulaPreviewSheets = Object.freeze(['Inputs', 'Preview', "O'Brien Data", 'Quota']);
export const formulaPreviewCases = Object.freeze([
  { id: 'occupied-formula', sheet: 'Preview', cell: 'C1', formula: 'SUM(Inputs!A1:A3)+10', type: 'number', value: 22 },
  {
    id: 'occupied-literal',
    sheet: 'Preview',
    cell: 'D2',
    formula: 'SUM(Inputs!A1,Inputs!A2)',
    type: 'number',
    value: 6,
  },
  { id: 'typed-text', sheet: 'Preview', cell: 'E1', formula: 'IF(TRUE,"#N/A",FALSE)', type: 'text', value: '#N/A' },
  { id: 'literal-error-reference', sheet: 'Preview', cell: 'E1', formula: 'Inputs!A4', type: 'error', value: '#N/A' },
  {
    id: 'error-looking-text',
    sheet: 'Preview',
    cell: 'E1',
    formula: 'IFERROR(Inputs!A5,"fallback")',
    type: 'text',
    value: '#N/A',
  },
  { id: 'boolean', sheet: 'Preview', cell: 'E1', formula: 'IF(FALSE,17,TRUE)', type: 'boolean', value: true },
  { id: 'division-error', sheet: 'Preview', cell: 'E1', formula: '1/0', type: 'error', value: '#DIV/0!' },
  {
    id: 'negative-zero',
    sheet: 'Preview',
    cell: 'E1',
    formula: '-0',
    type: 'number',
    value: -0,
    binary64: '8000000000000000',
  },
  {
    id: 'financial-positive-zero',
    sheet: 'Preview',
    cell: 'E1',
    formula: 'PMT(0,12,0)',
    type: 'number',
    value: 0,
    binary64: '0000000000000000',
  },
  { id: 'global-name', sheet: 'Preview', cell: 'E1', formula: 'SUM(Revenue)+10', type: 'number', value: 22 },
  { id: 'local-shadow', sheet: 'Preview', cell: 'E1', formula: 'SUM(Tax)', type: 'number', value: 4 },
  { id: 'quoted-sheet-name', sheet: 'Preview', cell: 'E1', formula: 'SUM(QuotedRevenue)+1', type: 'number', value: 16 },
]);
export const formulaPreviewRefusals = Object.freeze([
  { id: 'self-cycle', sheet: 'Preview', cell: 'C1', formula: 'C1+1', statuses: ['invalid'] },
  { id: 'indirect-cycle', sheet: 'Preview', cell: 'C1', formula: 'A1+1', statuses: ['invalid'] },
  { id: 'invalid-scope', sheet: 'Preview', cell: 'E1', formula: 'BadScope', statuses: ['unavailable'] },
  { id: 'duplicate-name', sheet: 'Preview', cell: 'E1', formula: 'Duplicate', statuses: ['unavailable'] },
  { id: 'foreign-local', sheet: 'Preview', cell: 'E1', formula: 'OtherLocal', statuses: ['unavailable'] },
  {
    id: 'dynamic-reference',
    sheet: 'Preview',
    cell: 'E1',
    formula: 'INDIRECT("Inputs!A1")',
    statuses: ['unavailable'],
  },
  {
    id: 'external-reference',
    sheet: 'Preview',
    cell: 'E1',
    formula: "'[other.xlsx]Inputs'!A1",
    statuses: ['invalid', 'unavailable'],
  },
  {
    id: 'unqualified-call-hidden-branch',
    sheet: 'Preview',
    cell: 'E1',
    formula: 'IF(TRUE,7,UNKNOWN(Inputs!A1))',
    statuses: ['unavailable'],
  },
  { id: 'unsafe-numeric-source', sheet: 'Preview', cell: 'E1', formula: 'SUM(Inputs!A7)', statuses: ['unavailable'] },
  { id: 'unsafe-long-decimal', sheet: 'Preview', cell: 'E1', formula: 'SUM(Inputs!A8)', statuses: ['unavailable'] },
]);
export const formulaPreviewQuotas = Object.freeze([
  { id: 'positions-at-limit', kind: 'positions', size: 9999, formula: 'SUM(A1:A9999)', value: 9999, positions: 10000 },
  {
    id: 'positions-over-limit',
    kind: 'positions',
    size: 10000,
    formula: 'SUM(A1:A10000)',
    status: 'unavailable',
    positions: 10001,
  },
  { id: 'formulas-at-limit', kind: 'formulas', size: 255, formula: 'SUM(A1:A255)', value: 255, formulaNodes: 256 },
  {
    id: 'formulas-over-limit',
    kind: 'formulas',
    size: 256,
    formula: 'SUM(A1:A256)',
    status: 'unavailable',
    formulaNodes: 257,
  },
  { id: 'depth-at-limit', kind: 'depth', size: 31, formula: 'A1', value: 31, depth: 32 },
  { id: 'depth-over-limit', kind: 'depth', size: 32, formula: 'A1', status: 'unavailable', depth: 33 },
]);
export function formulaPreviewFixture({ mode = 'auto', readonly = false, context = 'ordinary', quota } = {}) {
  if (!['auto', 'manual'].includes(mode) || !['ordinary', 'shared', 'spill', 'pivot-seed'].includes(context))
    throw Error('Invalid authored preview fixture');
  const parts = unzipSync(createWorkbook());
  const names =
    context === 'pivot-seed'
      ? ''
      : `<definedNames><definedName name="Revenue">'Inputs'!$A$1:$A$3</definedName><definedName name="Tax">'Inputs'!$A$1</definedName><definedName name="Tax" localSheetId="1">'Inputs'!$A$2</definedName><definedName name="QuotedRevenue">'O''Brien Data'!$B$1:$B$2</definedName><definedName name="OtherLocal" localSheetId="2">'O''Brien Data'!$B$1</definedName><definedName name="BadScope" localSheetId="9">'Inputs'!$A$1</definedName><definedName name="Duplicate">'Inputs'!$A$1</definedName><definedName name="duplicate">'Inputs'!$A$2</definedName></definedNames>`;
  parts['xl/workbook.xml'] = strToU8(
    `<workbook xmlns="${S}" xmlns:r="${R}"><workbookPr/><bookViews><workbookView/></bookViews><sheets>${formulaPreviewSheets.map((name, i) => `<sheet name="${esc(name)}" sheetId="${i + 1}" r:id="sheet${i + 1}"/>`).join('')}</sheets>${names}<calcPr calcId="191029" calcMode="${mode}" calcOnSave="0" fullCalcOnLoad="0" forceFullCalc="0"/></workbook>`
  );
  parts['xl/_rels/workbook.xml.rels'] = strToU8(
    `<Relationships xmlns="${P}">${formulaPreviewSheets.map((_, i) => `<Relationship Id="sheet${i + 1}" Type="${R}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/></Relationships>`
  );
  const extra = [2, 3, 4]
    .map(
      (i) =>
        `<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
    )
    .join('');
  parts['[Content_Types].xml'] = strToU8(
    strFromU8(parts['[Content_Types].xml']).replace(
      '</Types>',
      extra +
        (readonly
          ? '<Override PartName="/custom/preview-opaque.bin" ContentType="application/vnd.ms-office.vbaProject"/>'
          : '') +
        '</Types>'
    )
  );
  const inputs =
    context === 'pivot-seed'
      ? [
          [1, [text('A1', 'Group'), text('B1', 'Amount')]],
          [2, [text('A2', 'A'), number('B2', 3)]],
          [3, [text('A3', 'B'), number('B3', 5)]],
        ]
      : [
          [1, [number('A1', mode === 'manual' ? 5 : 2), formula('B1', 'A1+1', mode === 'manual' ? 103 : 3)]],
          [2, [number('A2', 4)]],
          [3, [number('A3', 6)]],
          [4, ['<c r="A4" t="e"><v>#N/A</v></c>']],
          [5, [text('A5', '#N/A')]],
          [6, ['<c r="A6" t="b"><v>0</v></c>']],
          [7, [number('A7', '9007199254740993')]],
          [8, [number('A8', '0.1234567890123456789012345')]],
        ];
  if (context === 'shared') {
    inputs[0][1].push(formula('F1', 'A1+1', 3, ' t="shared" si="0" ref="F1:F2"'));
    inputs[1][1].push('<c r="F2"><f t="shared" si="0"/><v>5</v></c>');
  }
  if (context === 'spill') inputs[0][1].push('<c r="F1"><f>SEQUENCE(3,2)</f></c>');
  parts['xl/worksheets/sheet1.xml'] = strToU8(sheet(inputs));
  parts['xl/worksheets/sheet2.xml'] = strToU8(
    sheet(
      context === 'pivot-seed'
        ? []
        : [
            [
              1,
              [formula('A1', 'C1+1', 913), formula('C1', 'SUM(Inputs!A1:A3)', 912), text('H1', ' 00123 =SUM(A1) <&> ')],
            ],
            [2, [number('D2', 777)]],
          ],
      'A1:H12'
    )
  );
  parts['xl/worksheets/sheet3.xml'] = strToU8(
    sheet(
      [
        [1, [number('B1', 7)]],
        [2, [number('B2', 8)]],
      ],
      'B1:B2'
    )
  );
  let quotaRows = [];
  if (quota) {
    if (
      !['positions', 'formulas', 'depth'].includes(quota.kind) ||
      !Number.isInteger(quota.size) ||
      quota.size < 1 ||
      quota.size > 10000
    )
      throw Error('Invalid preview quota fixture');
    quotaRows = Array.from({ length: quota.size }, (_, index) => [
      index + 1,
      [
        quota.kind === 'positions'
          ? number('A' + (index + 1), 1)
          : formula(
              'A' + (index + 1),
              quota.kind === 'depth' && index + 1 < quota.size ? 'A' + (index + 2) + '+1' : '1',
              900 + index
            ),
      ],
    ]);
  }
  parts['xl/worksheets/sheet4.xml'] = strToU8(sheet(quotaRows, 'A1:D' + Math.max(1, quota?.size ?? 1)));
  parts['custom/preview-opaque.bin'] = Uint8Array.from([0, 255, 31, 3, 17, 42]);
  return Buffer.from(zipSync(parts, { mtime: new Date(2000, 0, 1) }));
}
