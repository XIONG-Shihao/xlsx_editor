import { rangeBounds, columnName, assertXmlText } from './workbook-commands.mjs';
import { chartStyle } from './chart-styles.mjs';
import { refusal } from './refusal-codes.mjs';
export const CHART_SOURCE_LIMIT = 1000;
const object = (value, keys, label) => {
  if (!value || typeof value !== 'object' || Array.isArray(value) || Object.keys(value).some((k) => !keys.includes(k)))
    throw refusal('internal-invalid-field', 'Invalid ' + label);
};
const text = (value, limit, label) => {
  if (typeof value !== 'string' || value.length > limit) throw refusal('internal-invalid-field', 'Invalid ' + label);
  assertXmlText(value);
  return value;
};
export function normalizeChartModel(value) {
  object(value, ['kind', 'range', 'title', 'legend', 'blanks', 'includeHidden', 'style'], 'chart model');
  if (!['column', 'bar', 'line'].includes(value.kind))
    throw refusal('chart-unsupported', 'Only 2D column, bar and line charts are qualified');
  const bounds = rangeBounds(value.range);
  if (bounds.rows < 2 || bounds.columns < 2 || bounds.columns > 9 || bounds.size > CHART_SOURCE_LIMIT)
    throw refusal('chart-limit', 'Chart source requires headers/categories and at most 1000 positions/eight series');
  if (
    (value.blanks !== undefined && value.blanks !== 'gap') ||
    (value.includeHidden !== undefined && value.includeHidden !== false)
  )
    throw refusal('chart-unsupported', 'Chart blank/hidden policy is not qualified');
  const legend = value.legend ?? 'bottom';
  if (!['bottom', 'none'].includes(legend)) throw refusal('chart-unsupported', 'Unsupported chart legend');
  if (value.style !== undefined) chartStyle(value.style);
  return {
    ...(value.style !== undefined ? { style: value.style } : {}),
    kind: value.kind,
    range: `${columnName(bounds.first.col)}${bounds.first.row + 1}:${columnName(bounds.last.col)}${bounds.last.row + 1}`,
    title: text(value.title ?? '', 128, 'chart title'),
    legend,
    blanks: 'gap',
    includeHidden: false,
  };
}
function numericMeaning(text) {
  const match = /^([+-]?)(\d*(?:\.\d*)?)(?:[eE]([+-]?\d+))?$/.exec(text);
  if (!match || !/[0-9]/.test(match[2])) return null;
  const [whole, fraction = ''] = match[2].split('.');
  let exponent = Number(match[3] ?? 0) - fraction.length,
    digits = (whole + fraction).replace(/^0+/, '');
  if (!Number.isSafeInteger(exponent)) return null;
  if (!digits) return (match[1] === '-' ? '-' : '') + '0:0';
  while (digits.endsWith('0')) {
    digits = digits.slice(0, -1);
    exponent++;
  }
  return (match[1] === '-' ? '-' : '') + digits + ':' + exponent;
}
function acceptedValue(cell, ref) {
  if (!cell) return null;
  if (
    cell.dataType === 'error' ||
    cell.error ||
    cell.calculation?.status === 'error' ||
    cell.calculation?.error ||
    cell.formulaType ||
    cell.scenarioResult
  )
    throw refusal('chart-unsupported', `Chart source ${ref} is an error/special result`);
  let value = cell.value;
  if (cell.formula !== undefined) {
    if (cell.calculation?.status !== 'calculated' || cell.calculation.pending)
      throw refusal('chart-unsupported', `Chart source ${ref} is unresolved`);
    value = cell.calculation.value;
  } else if (cell.dataType === 'blank' || value === null) return null;
  if (typeof value === 'number') {
    if (
      !Number.isFinite(value) ||
      (cell.formula === undefined &&
        typeof cell.rawValue === 'string' &&
        /^[+-]?\d+$/.test(cell.rawValue) &&
        Math.abs(Number(cell.rawValue)) > Number.MAX_SAFE_INTEGER)
    )
      throw refusal('chart-unsupported', `Chart source ${ref} is an unsafe number`);
    if (
      cell.formula === undefined &&
      typeof cell.rawValue === 'string' &&
      numericMeaning(cell.rawValue) !== numericMeaning(Object.is(value, -0) ? '-0' : String(value))
    )
      throw refusal('chart-unsupported', `Chart source ${ref} is an unsafe raw number`);
  }
  return value;
}
export async function readChartData(readRange, sheet, model, view) {
  model = normalizeChartModel(model);
  if (typeof readRange !== 'function')
    throw refusal('chart-unsupported', 'Accepted chart source reader is unavailable');
  const bounds = rangeBounds(model.range),
    reply = await readRange(sheet, model.range),
    cells = reply.cells ?? reply;
  const headers = [];
  for (let col = bounds.first.col + 1; col <= bounds.last.col; col++) {
    const ref = columnName(col) + (bounds.first.row + 1),
      value = acceptedValue(cells[ref], ref);
    if (typeof value !== 'string' || !value.trim() || value.length > 128 || cells[ref]?.formula !== undefined)
      throw refusal('chart-bad-input', `Chart header ${ref} needs literal text`);
    assertXmlText(value);
    headers.push(value);
  }
  const categories = [],
    series = headers.map((name) => ({
      name,
      values: [],
    })),
    hidden = [];
  const hiddenRows = new Set(view.sheets.find((item) => item.name === sheet)?.hiddenRows ?? []);
  for (let row = bounds.first.row + 1; row <= bounds.last.row; row++) {
    const ref = columnName(bounds.first.col) + (row + 1),
      value = acceptedValue(cells[ref], ref);
    if (
      (value !== null && typeof value !== 'string' && typeof value !== 'number') ||
      (typeof value === 'string' && value.length > 256) ||
      cells[ref]?.formula !== undefined
    )
      throw refusal('chart-bad-input', `Chart category ${ref} needs literal text/number`);
    categories.push(value === null ? '' : Object.is(value, -0) ? '-0' : String(value));
    hidden.push(hiddenRows.has(row));
    for (let col = bounds.first.col + 1; col <= bounds.last.col; col++) {
      const ref = columnName(col) + (row + 1),
        number = acceptedValue(cells[ref], ref);
      if (number !== null && typeof number !== 'number')
        throw refusal('chart-bad-input', `Chart value ${ref} needs a numeric result or blank`);
      series[col - bounds.first.col - 1].values.push(number);
    }
  }
  return {
    categories,
    series,
    hidden,
  };
}

/** Finite, data-dependent proposals. This primitive does not grant an editing frame. */
export async function createChartRecommendations({ sheet, range, view, readRange }) {
  const base = normalizeChartModel({ kind: 'column', range, title: '', legend: 'bottom' }),
    data = await readChartData(readRange, sheet, base, view),
    count = data.categories.length;
  if (!data.series.some((s) => s.values.some((v) => v !== null)))
    throw refusal('chart-bad-input', 'Chart recommendations need at least one numeric value');
  const longLabels = data.categories.some((s) => s.length > 12),
    ordered =
      data.categories.length > 1 &&
      data.categories.every((s) => /^-?\d+(?:\.\d+)?$/.test(s)) &&
      data.categories.every((s, i, a) => !i || Number(s) > Number(a[i - 1]));
  const title = data.series
      .map((s) => s.name)
      .join(' / ')
      .slice(0, 128),
    proposals = [
      {
        id: 'column-comparison',
        label: 'Compare categories',
        reason: `${count} categories and ${data.series.length} numeric series.`,
        chart: normalizeChartModel({ ...base, title, style: 'colorful' }),
        data,
      },
      {
        id: 'bar-comparison',
        label: 'Compare labels',
        reason: longLabels
          ? 'Horizontal bars give longer category labels more room.'
          : `${count} categories in a horizontal comparison.`,
        chart: normalizeChartModel({ ...base, kind: 'bar', title, style: 'monochrome' }),
        data,
      },
      {
        id: 'line-sequence',
        label: 'Follow row order',
        reason: ordered
          ? 'Numeric categories increase in worksheet order.'
          : 'Connect values in the selected worksheet row order.',
        chart: normalizeChartModel({ ...base, kind: 'line', title, style: 'outline' }),
        data,
      },
    ];
  if (longLabels || count > 8) proposals.unshift(proposals.splice(1, 1)[0]);
  else if (ordered) proposals.unshift(proposals.pop());
  return { sheet, range: base.range, rows: count, series: data.series.length, proposals };
}
