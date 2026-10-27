/** Closed worksheet PivotTable model. A cache and its displayed cells share one model. */
import { createHash } from 'node:crypto';
import { pivotLayoutSample, pivotOutputData } from './pivot-layout-samples.mjs';
import { hasExactNumericValue } from './calculation.mjs';
import { rangeBounds, columnName, assertXmlText } from './workbook-commands.mjs';
import { refusal } from './refusal-codes.mjs';
export const PIVOT_LIMITS = Object.freeze({ positions: 8000, rows: 1001, columns: 8, groups: 512, text: 128 });
export const PIVOT_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const PIVOT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const PIVOT_TYPES = Object.freeze({
  table: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
  definition: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
  records: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
});
const hash = (value) => createHash('sha256').update(value).digest('hex');
export function pivotEscape(value) {
  assertXmlText(String(value));
  return String(value).replace(
    /[&<>"\r\n\t]/g,
    (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', '\r': '&#13;', '\n': '&#10;', '\t': '&#9;' })[ch]
  );
}
export function pivotIdentifier(value) {
  return (
    typeof value === 'string' &&
    /^[A-Za-z_][A-Za-z0-9_]{0,127}$/.test(value) &&
    !/^[A-Z]{1,3}[1-9]\d*$/i.test(value) &&
    !/^(?:r|c|r\d+c\d+|_xlnm\..*)$/i.test(value)
  );
}
function scalar(cell) {
  if (!cell) return null;
  if (
    cell.formulaOutput ||
    cell.formulaType ||
    cell.spill ||
    cell.calculation?.pending ||
    cell.richText ||
    (!cell.formula && cell.dataType === 'error') ||
    cell.vm !== undefined ||
    cell.cm !== undefined
  )
    throw refusal('pivot-bad-input', 'Pivot source requires ordinary current cells');
  if (cell.formula && (cell.calculation?.status !== 'calculated' || typeof cell.calculation.value !== 'number'))
    throw refusal('pivot-bad-input', 'Pivot formula source requires a current numeric calculation');
  const value = cell.formula ? cell.calculation.value : cell.value;
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number') {
    if (!cell.formula && !hasExactNumericValue(value, cell.rawValue))
      throw refusal('pivot-unsupported', 'Pivot literal numeric precision requires qualification');
    if (!Number.isSafeInteger(value)) throw refusal('pivot-bad-input', 'Pivot numeric source requires safe integers');
    return Object.is(value, -0) ? 0 : value;
  }
  if (typeof value === 'boolean') return value;
  if (typeof value === 'string') {
    if (value.length > PIVOT_LIMITS.text) throw refusal('pivot-limit', 'Pivot source text exceeds its bound');
    assertXmlText(value);
    if (/_x[0-9a-f]{4}_/i.test(value))
      throw refusal('pivot-unsupported', 'Literal SpreadsheetML escape patterns require separate pivot qualification');
    return value;
  }
  throw refusal('pivot-unsupported', 'Unsupported pivot source value');
}
const scalarKey = (value) => JSON.stringify([value === null ? 'blank' : typeof value, value]);
export function createPivotModel({ sheet, range, cells, rowField, valueField, aggregate = 'sum' }) {
  const b = rangeBounds(range);
  if (
    b.rows < 2 ||
    b.rows > PIVOT_LIMITS.rows ||
    b.columns < 2 ||
    b.columns > PIVOT_LIMITS.columns ||
    b.size > PIVOT_LIMITS.positions
  )
    throw refusal('pivot-limit', 'Pivot source exceeds the bounded rectangle');
  if (!['sum', 'count'].includes(aggregate))
    throw refusal('pivot-unsupported', 'Only SUM and numeric COUNT pivot aggregates are qualified');
  const headers = [];
  for (let c = b.first.col; c <= b.last.col; c++) {
    const v = scalar(cells[columnName(c) + (b.first.row + 1)]);
    if (typeof v !== 'string' || !v.trim() || headers.some((h) => h.toLowerCase() === v.toLowerCase()))
      throw refusal('pivot-bad-input', 'Pivot headers must be distinct nonempty text');
    headers.push(v);
  }
  const row = typeof rowField === 'number' ? rowField : headers.indexOf(rowField),
    value = typeof valueField === 'number' ? valueField : headers.indexOf(valueField);
  if (
    !Number.isInteger(row) ||
    !Number.isInteger(value) ||
    row < 0 ||
    value < 0 ||
    row >= headers.length ||
    value >= headers.length ||
    row === value
  )
    throw refusal('pivot-unsupported', 'Choose different existing row and value fields');
  const fields = headers.map((name) => ({ name, items: [], keys: new Map() })),
    records = [],
    groups = [],
    groupMap = new Map(),
    folded = new Map();
  let total = 0;
  for (let r = b.first.row + 1; r <= b.last.row; r++) {
    const values = headers.map((_, i) => scalar(cells[columnName(b.first.col + i) + (r + 1)]));
    if (typeof values[row] !== 'string')
      throw refusal('pivot-bad-input', 'Pivot row labels require nonempty literal text');
    const label = values[row],
      fold = label.toLowerCase();
    if (folded.has(fold) && folded.get(fold) !== label)
      throw refusal('pivot-conflict', 'Case-colliding pivot labels require separate qualification');
    folded.set(fold, label);
    if (values[value] !== null && typeof values[value] !== 'number')
      throw refusal('pivot-bad-input', 'Pivot value field requires numeric or blank literals');
    const record = values.map((v, i) => {
      const f = fields[i],
        key = scalarKey(v);
      if (!f.keys.has(key)) {
        f.keys.set(key, f.items.length);
        f.items.push(v);
      }
      return f.keys.get(key);
    });
    records.push(record);
    if (!groupMap.has(label)) {
      if (groups.length >= PIVOT_LIMITS.groups) throw refusal('pivot-limit', 'Pivot row groups exceed their bound');
      groupMap.set(label, groups.length);
      groups.push({ label, item: record[row], value: 0 });
    }
    const amount = values[value] === null ? 0 : aggregate === 'count' ? 1 : values[value],
      g = groups[groupMap.get(label)];
    g.value += amount;
    total += amount;
    if (!Number.isSafeInteger(g.value) || !Number.isSafeInteger(total))
      throw refusal('pivot-limit', 'Pivot aggregate exceeds exact integer qualification');
  }
  const model = {
    sheet,
    range: columnName(b.first.col) + (b.first.row + 1) + ':' + columnName(b.last.col) + (b.last.row + 1),
    rowField: row,
    valueField: value,
    aggregate,
    fields: fields.map(({ name, items }) => ({ name, items })),
    records,
    groups,
    total,
  };
  return { ...model, fingerprint: hash(JSON.stringify(model)) };
}
function itemXml(v) {
  if (v === null) return '<m/>';
  if (typeof v === 'string') return '<s v="' + pivotEscape(v) + '"/>';
  if (typeof v === 'boolean') return '<b v="' + (v ? '1' : '0') + '"/>';
  return '<n v="' + v + '"/>';
}
function sharedXml(items) {
  // Office groups Boolean values with text for cache type flags. Blank values
  // also require semi-mixed typing, even in an otherwise numeric field.
  const types = new Set(items.filter((v) => v !== null).map((v) => (typeof v === 'number' ? 'number' : 'string')));
  const numbers = types.has('number'),
    strings = types.has('string'),
    blank = items.includes(null);
  return (
    '<sharedItems count="' +
    items.length +
    '" containsBlank="' +
    Number(blank) +
    '" containsString="' +
    Number(strings) +
    '" containsNumber="' +
    Number(numbers) +
    '" containsInteger="' +
    Number(numbers) +
    '" containsDate="0" containsNonDate="1" containsSemiMixedTypes="' +
    Number(strings || blank) +
    '" containsMixedTypes="' +
    Number(types.size > 1) +
    '">' +
    items.map(itemXml).join('') +
    '</sharedItems>'
  );
}
export function pivotParts(model, { name, cacheId = 1, sample = 'tabular-light', grandTotals = true }) {
  const layout = pivotLayoutSample(sample);
  if (typeof grandTotals !== 'boolean')
    throw refusal('pivot-bad-input', 'Pivot grand totals require an explicit Boolean');
  const count = model.groups.length + Number(grandTotals);
  if (!pivotIdentifier(name) || !Number.isSafeInteger(cacheId) || cacheId < 1)
    throw refusal('pivot-malformed', 'Invalid pivot identity');
  const caption = (model.aggregate === 'sum' ? 'Sum of ' : 'Count of ') + model.fields[model.valueField].name,
    rows = model.groups.length + 1 + Number(grandTotals),
    output = 'A1:B' + rows;
  const definition =
    '<pivotCacheDefinition xmlns="' +
    PIVOT_NS +
    '" xmlns:r="' +
    PIVOT_REL +
    '" r:id="rIdRecords" saveData="1" refreshOnLoad="0" enableRefresh="1" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="' +
    model.records.length +
    '"><cacheSource type="worksheet"><worksheetSource ref="' +
    model.range +
    '" sheet="' +
    pivotEscape(model.sheet) +
    '"/></cacheSource><cacheFields count="' +
    model.fields.length +
    '">' +
    model.fields
      .map((f) => '<cacheField name="' + pivotEscape(f.name) + '" numFmtId="0">' + sharedXml(f.items) + '</cacheField>')
      .join('') +
    '</cacheFields></pivotCacheDefinition>';
  const records =
    '<pivotCacheRecords xmlns="' +
    PIVOT_NS +
    '" count="' +
    model.records.length +
    '">' +
    model.records.map((r) => '<r>' + r.map((i) => '<x v="' + i + '"/>').join('') + '</r>').join('') +
    '</pivotCacheRecords>';
  const fields = model.fields
    .map((f, i) =>
      i === model.rowField
        ? '<pivotField axis="axisRow" showAll="0" defaultSubtotal="0" compact="' +
          layout.compact +
          '" outline="' +
          layout.outline +
          '"><items count="' +
          model.groups.length +
          '">' +
          model.groups.map((g) => '<item x="' + g.item + '"/>').join('') +
          '</items></pivotField>'
        : '<pivotField' + (i === model.valueField ? ' dataField="1"' : '') + ' showAll="0"/>'
    )
    .join('');
  const table =
    '<pivotTableDefinition xmlns="' +
    PIVOT_NS +
    '" name="' +
    name +
    '" cacheId="' +
    cacheId +
    '" dataCaption="Values" grandTotalCaption="Grand Total" rowGrandTotals="' +
    Number(grandTotals) +
    '" colGrandTotals="' +
    Number(grandTotals) +
    '" compact="' +
    layout.compact +
    '" compactData="' +
    layout.compact +
    '" outline="' +
    layout.outline +
    '" outlineData="' +
    layout.outline +
    '" gridDropZones="0" updatedVersion="6" minRefreshableVersion="3" createdVersion="6" useAutoFormatting="0"><location ref="' +
    output +
    '" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="' +
    model.fields.length +
    '">' +
    fields +
    '</pivotFields><rowFields count="1"><field x="' +
    model.rowField +
    '"/></rowFields><rowItems count="' +
    count +
    '">' +
    model.groups.map((g, i) => '<i><x v="' + i + '"/></i>').join('') +
    (grandTotals ? '<i t="grand"><x/></i>' : '') +
    '</rowItems><colItems count="1"><i/></colItems><dataFields count="1"><dataField name="' +
    pivotEscape(caption) +
    '" fld="' +
    model.valueField +
    '" subtotal="' +
    model.aggregate +
    '" baseField="0" baseItem="0" numFmtId="0"/></dataFields><pivotTableStyleInfo name="' +
    layout.style +
    '" showRowHeaders="1" showColHeaders="1" showRowStripes="' +
    layout.stripes +
    '" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>';
  const data = pivotOutputData(model, { sample, grandTotals });
  const cells = data
    .map(
      (r, i) =>
        '<row r="' +
        (i + 1) +
        '">' +
        r
          .map((v, c) => {
            const ref = columnName(c) + (i + 1);
            return typeof v === 'number'
              ? '<c r="' + ref + '"><v>' + v + '</v></c>'
              : '<c r="' + ref + '" t="inlineStr"><is><t xml:space="preserve">' + pivotEscape(v) + '</t></is></c>';
          })
          .join('') +
        '</row>'
    )
    .join('');
  const worksheet =
    '<worksheet xmlns="' +
    PIVOT_NS +
    '" xmlns:r="' +
    PIVOT_REL +
    '"><dimension ref="' +
    output +
    '"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="24" customWidth="1"/><col min="2" max="2" width="20" customWidth="1"/></cols><sheetData>' +
    cells +
    '</sheetData></worksheet>';
  return { definition, records, table, worksheet, output };
}
