// Independent authored formulas, caches and expected edges for calculation controls/tracing.
// Expectations are deliberately literal; no production parser/evaluator generates them.
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook } from '../../core/workbook-package.mjs';
const n = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  p = 'http://schemas.openxmlformats.org/package/2006/relationships';
const esc = (v) => String(v).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('"', '&quot;');
function sheet(rows, { table = false } = {}) {
  return `<worksheet xmlns="${n}" xmlns:r="${r}"><dimension ref="A1:Z4"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData>${rows.map(([row, cells]) => `<row r="${row}">${cells.join('')}</row>`).join('')}</sheetData>${table ? '<tableParts count="1"><tablePart r:id="table1"/></tableParts>' : ''}</worksheet>`;
}
const literal = (cell, value) =>
  typeof value === 'string'
    ? `<c r="${cell}" t="inlineStr"><is><t>${esc(value)}</t></is></c>`
    : `<c r="${cell}"><v>${value}</v></c>`;
const formula = (cell, text, value) =>
  `<c r="${cell}"><f>${esc(text)}</f>${value === undefined ? '' : `<v>${value}</v>`}</c>`;
export function calculationTraceFixture({
  mode = 'auto',
  recalcBeforeSave = true,
  markers = false,
  readonly = false,
} = {}) {
  if (!['auto', 'manual'].includes(mode)) throw Error('Invalid fixture mode');
  const parts = unzipSync(createWorkbook()),
    names = ['Inputs', 'Data', 'Report', 'Downstream'];
  parts['xl/workbook.xml'] = strToU8(
    `<workbook xmlns="${n}" xmlns:r="${r}"><workbookPr/><bookViews><workbookView/></bookViews><sheets>${names.map((name, i) => `<sheet name="${name}" sheetId="${i + 1}" r:id="sheet${i + 1}"/>`).join('')}</sheets><definedNames><definedName name="RateInput">Inputs!$A$1</definedName></definedNames><calcPr calcId="191029" calcMode="${mode}" calcOnSave="${recalcBeforeSave ? '1' : '0'}" fullCalcOnLoad="0" forceFullCalc="0" iterate="0"/></workbook>`
  );
  parts['xl/_rels/workbook.xml.rels'] = strToU8(
    `<Relationships xmlns="${p}">${names.map((_, i) => `<Relationship Id="sheet${i + 1}" Type="${r}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="styles" Type="${r}/styles" Target="styles.xml"/></Relationships>`
  );
  parts['[Content_Types].xml'] = strToU8(
    strFromU8(parts['[Content_Types].xml']).replace(
      '</Types>',
      `${[2, 3, 4].map((i) => `<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>${readonly ? '<Override PartName="/custom/control-opaque.bin" ContentType="application/vnd.ms-office.vbaProject"/>' : ''}</Types>`
    )
  );
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    sheet([
      [1, [literal('A1', 2), formula('B1', 'A1+1', 3), formula('D1', 'Data!B2*2', 6)]],
      [2, [literal('A2', 3), formula('B2', 'SUM(A1:A2)', 5)]],
    ])
  );
  parts['xl/worksheets/sheet2.xml'] = strToU8(
    sheet(
      [
        [1, [literal('A1', 'Key'), literal('B1', 'Amount')]],
        [2, [literal('A2', 'alpha'), literal('B2', 3)]],
        [3, [literal('A3', 'beta'), literal('B3', 5)]],
        [4, [literal('A4', 'gamma'), literal('B4', 7)]],
      ],
      { table: true }
    )
  );
  const reportRows = [
    [1, [formula('A1', 'Inputs!B1*2', 6), formula('B1', 'SUM(Rates[Amount])', 15)]],
    [2, [formula('A2', 'Inputs!B2+Data!B2', 8), formula('B2', 'RateInput*3', 6)]],
  ];
  if (markers) {
    reportRows[0][1].push(formula('C1', 'INDIRECT("Inputs!A1")+Inputs!A2'));
    reportRows[1][1].push(formula('C2', 'UNKNOWN(Inputs!A1)'));
    reportRows.push([3, [formula('C3', 'C4+1')]]);
    reportRows.push([4, [formula('C4', 'C3+1')]]);
  }
  parts['xl/worksheets/sheet3.xml'] = strToU8(sheet(reportRows));
  parts['xl/worksheets/sheet4.xml'] = strToU8(sheet([[1, [formula('A1', 'Report!A1+1', 7)]]]));
  parts['xl/worksheets/_rels/sheet2.xml.rels'] = strToU8(
    `<Relationships xmlns="${p}"><Relationship Id="table1" Type="${r}/table" Target="../tables/table1.xml"/></Relationships>`
  );
  parts['xl/tables/table1.xml'] = strToU8(
    `<table xmlns="${n}" id="1" name="Rates" displayName="Rates" ref="A1:B4" totalsRowShown="0"><autoFilter ref="A1:B4"/><tableColumns count="2"><tableColumn id="1" name="Key"/><tableColumn id="2" name="Amount"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`
  );
  parts['custom/control-opaque.bin'] = new Uint8Array([0, 255, 21, 13, 9, 14]);
  return Buffer.from(zipSync(parts, { mtime: new Date(2000, 0, 1) }));
}
export const calculationControlExpected = Object.freeze({
  source: {
    'Inputs!B1': 3,
    'Inputs!B2': 5,
    'Inputs!D1': 6,
    'Report!A1': 6,
    'Report!A2': 8,
    'Report!B1': 15,
    'Report!B2': 6,
    'Downstream!A1': 7,
  },
  edits: [
    { sheet: 'Inputs', cell: 'A1', value: 5 },
    { sheet: 'Data', cell: 'B2', value: 9 },
  ],
  afterAll: {
    'Inputs!B1': 6,
    'Inputs!B2': 8,
    'Inputs!D1': 18,
    'Report!A1': 12,
    'Report!A2': 17,
    'Report!B1': 21,
    'Report!B2': 15,
    'Downstream!A1': 13,
  },
  reportClosure: ['Inputs!B1', 'Inputs!B2', 'Report!A1', 'Report!A2', 'Report!B1', 'Report!B2'],
  pendingAfterReport: ['Inputs!D1', 'Downstream!A1'],
  manualBeforeCalculation:
    'Changed input literals are current; retained formula results must be explicitly stored/pending, never newly certified calculated.',
  sheetPolicy:
    'Local qualified selected-sheet plus required formula-ancestor closure; not an independently captured Excel worksheet-only result.',
});
export const traceExpected = Object.freeze([
  {
    sheet: 'Inputs',
    cell: 'B1',
    direction: 'precedents',
    edges: [{ from: { sheet: 'Inputs', range: 'A1' }, to: { sheet: 'Inputs', range: 'B1' }, via: 'formula' }],
  },
  {
    sheet: 'Inputs',
    cell: 'B2',
    direction: 'precedents',
    edges: [{ from: { sheet: 'Inputs', range: 'A1:A2' }, to: { sheet: 'Inputs', range: 'B2' }, via: 'formula' }],
  },
  {
    sheet: 'Report',
    cell: 'A1',
    direction: 'precedents',
    edges: [{ from: { sheet: 'Inputs', range: 'B1' }, to: { sheet: 'Report', range: 'A1' }, via: 'formula' }],
  },
  {
    sheet: 'Report',
    cell: 'B1',
    direction: 'precedents',
    edges: [{ from: { sheet: 'Data', range: 'B2:B4' }, to: { sheet: 'Report', range: 'B1' }, via: 'table' }],
  },
  {
    sheet: 'Report',
    cell: 'B2',
    direction: 'precedents',
    edges: [{ from: { sheet: 'Inputs', range: 'A1' }, to: { sheet: 'Report', range: 'B2' }, via: 'name' }],
  },
  {
    sheet: 'Inputs',
    cell: 'A1',
    direction: 'dependents',
    edges: [
      { from: { sheet: 'Inputs', range: 'A1' }, to: { sheet: 'Inputs', range: 'B1' }, via: 'formula' },
      { from: { sheet: 'Inputs', range: 'A1:A2' }, to: { sheet: 'Inputs', range: 'B2' }, via: 'formula' },
      { from: { sheet: 'Inputs', range: 'A1' }, to: { sheet: 'Report', range: 'B2' }, via: 'name' },
    ],
  },
  {
    sheet: 'Report',
    cell: 'C1',
    direction: 'precedents',
    edges: [{ from: { sheet: 'Inputs', range: 'A2' }, to: { sheet: 'Report', range: 'C1' }, via: 'formula' }],
    requiredMarker: 'dynamic',
    complete: false,
  },
  {
    sheet: 'Report',
    cell: 'C2',
    direction: 'precedents',
    edges: [{ from: { sheet: 'Inputs', range: 'A1' }, to: { sheet: 'Report', range: 'C2' }, via: 'formula' }],
    requiredMarker: 'unsupported',
    complete: false,
  },
  {
    sheet: 'Report',
    cell: 'C3',
    direction: 'precedents',
    edges: [{ from: { sheet: 'Report', range: 'C4' }, to: { sheet: 'Report', range: 'C3' }, via: 'formula' }],
  },
  {
    sheet: 'Report',
    cell: 'C4',
    direction: 'precedents',
    edges: [{ from: { sheet: 'Report', range: 'C3' }, to: { sheet: 'Report', range: 'C4' }, via: 'formula' }],
  },
]);
