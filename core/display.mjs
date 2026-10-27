/** One typed display contract for Grid, native-search parity, Replace and print.
 * Formatting never creates calculation authority or changes authored values. */
import { hasExactNumericValue } from './calculation.mjs';
import { formatCustomNumber } from './custom-number-format.mjs';
const plain = (text) => ({ text: String(text), formatStatus: 'not-applicable' });
const preserved = (value) => ({ text: String(value), formatStatus: 'not-applicable', reason: 'precision-preserved' });
function date(value, date1904) {
  if (!date1904 && Math.floor(value) === 60) return '1900-02-29';
  const days = date1904 ? value : value < 60 ? value + 1 : value;
  const result = new Date((date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30)) + days * 86400000);
  return Number.isFinite(result.getTime()) ? result.toISOString().slice(0, 10) : String(value);
}
/** Numeric IDs are built-ins; strings are authored custom codes, never IDs. */
export function formatDisplayDetails(value, format, date1904 = false) {
  if (typeof format === 'string') {
    if (typeof value === 'number' && !Number.isFinite(value)) return plain(value);
    if (typeof value === 'number' || typeof value === 'string') {
      const result = formatCustomNumber(value, format, date1904);
      if (result)
        return { text: result.text, formatStatus: 'supported', ...(result.color ? { color: result.color } : {}) };
      return { text: String(value), formatStatus: 'unsupported', reason: 'unsupported-custom-format' };
    }
    return plain(value);
  }
  if (typeof value !== 'number' || !Number.isFinite(value)) return plain(value);
  let text;
  if (format === 9 || format === 10) text = (value * 100).toFixed(format === 9 ? 0 : 2) + '%';
  else if ([14, 15, 16, 17, 22].includes(format)) text = date(value, date1904);
  else if (format === 2) text = value.toFixed(2);
  else if (format === 1) text = value.toFixed(0);
  else if (format === 3) text = value.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
  else if (format === 4 || format === 44)
    text = value.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  else if (format === undefined || format === 0 || format === 49) text = String(value);
  else return { text: String(value), formatStatus: 'unsupported', reason: 'unsupported-custom-format' };
  return { text, formatStatus: 'supported' };
}
export function formatDisplayValue(value, format, date1904 = false) {
  return formatDisplayDetails(value, format, date1904).text;
}
/** Guard raw values and typed errors before a text section can decorate them. */
export function displayCellDetails(cell, date1904 = false) {
  if (cell?.calculation?.pending) {
    const cache = cell.cachedResult;
    if (cache?.status === 'error') return plain('Stored: ' + (cache.error ?? '#ERROR!'));
    const value = cache?.value;
    if (value === undefined) return plain('Pending calculation');
    const details =
      typeof value === 'number' && !hasExactNumericValue(value, cache.rawValue)
        ? preserved(cache.rawValue)
        : formatDisplayDetails(value, cell?.format, date1904);
    return { ...details, text: 'Stored: ' + details.text };
  }
  if (cell?.formula !== undefined || cell?.formulaType !== undefined || cell?.scenarioResult) {
    const result = cell.calculation;
    if (!result) return plain('Calculating…');
    if (result.status === 'circular') return plain('#CYCLE!');
    if (result.status === 'error') return plain(result.error ?? '#ERROR!');
    if (!['calculated', 'stored'].includes(result.status)) return plain('Unresolved');
    if (result.value == null) return plain(result.error ?? '');
    if (result.status === 'stored') {
      if (result.error !== undefined || cell.dataType === 'error' || cell.error !== undefined)
        return plain(result.error ?? cell.error ?? result.value);
      const raw = result.rawValue ?? cell.rawValue;
      if (typeof result.value === 'number' && !hasExactNumericValue(result.value, raw)) return preserved(raw);
    }
    return formatDisplayDetails(result.value, cell?.format, date1904);
  }
  if (cell?.dataType === 'error' || cell?.error !== undefined) return plain(cell.error ?? cell.value ?? '');
  if (typeof cell?.value === 'number' && !hasExactNumericValue(cell.value, cell.rawValue))
    return preserved(cell.rawValue);
  if (cell?.value == null) return plain('');
  return formatDisplayDetails(cell.value, cell?.format, date1904);
}
export function displayCell(cell, date1904 = false) {
  return displayCellDetails(cell, date1904).text;
}
