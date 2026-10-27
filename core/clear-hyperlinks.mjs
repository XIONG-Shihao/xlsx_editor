/** Native-owned hyperlink metadata removal. Target geometry never reads cells. */
import { rangeBounds, columnName } from './workbook-commands.mjs';
import { refusal } from './refusal-codes.mjs';
export const CLEAR_HYPERLINKS_LIMITS = Object.freeze({ ranges: 10000, actionBytes: 2 * 1024 * 1024 });
const fields = ['type', 'sheet', 'ranges'];
const plain = (value) =>
  !!value &&
  typeof value === 'object' &&
  !Array.isArray(value) &&
  (Object.getPrototypeOf(value) === Object.prototype || Object.getPrototypeOf(value) === null);
const sheetName = (value) =>
  typeof value === 'string' &&
  value.length > 0 &&
  value.length <= 31 &&
  !!value.trim() &&
  !/[\\/*?:\[\]\u0000-\u001f\uD800-\uDFFF\uFFFE\uFFFF]/u.test(value) &&
  !value.startsWith("'") &&
  !value.endsWith("'");
function area(value) {
  if (typeof value !== 'string' || !/^([A-Z]{1,3}[1-9]\d{0,6})(?::([A-Z]{1,3}[1-9]\d{0,6}))?$/.test(value))
    throw refusal('hyperlink-bad-input', 'Clear Hyperlinks needs canonical worksheet rectangles.');
  const bounds = rangeBounds(value),
    first = columnName(bounds.first.col) + (bounds.first.row + 1),
    last = columnName(bounds.last.col) + (bounds.last.row + 1);
  if (value !== (first === last ? first : first + ':' + last))
    throw refusal('hyperlink-bad-input', 'Clear Hyperlinks needs canonical worksheet rectangles.');
  return { ...bounds, range: value };
}
/** Sweep row events, querying occupied column intervals. Fixed worksheet
 * geometry bounds memory; full-sheet selections do not enumerate any cells. */
function disjoint(areas) {
  if (areas.length < 2) return;
  const events = areas
    .flatMap((a) => [
      { row: a.first.row, first: a.first.col, last: a.last.col, delta: 1 },
      { row: a.last.row + 1, first: a.first.col, last: a.last.col, delta: -1 },
    ])
    .sort((a, b) => a.row - b.row || a.delta - b.delta);
  const maximum = new Int32Array(65536),
    lazy = new Int32Array(65536);
  function update(node, first, last, start, end, delta) {
    if (start <= first && last <= end) {
      maximum[node] += delta;
      lazy[node] += delta;
      return;
    }
    const middle = (first + last) >> 1;
    if (start <= middle) update(node * 2, first, middle, start, end, delta);
    if (end > middle) update(node * 2 + 1, middle + 1, last, start, end, delta);
    maximum[node] = lazy[node] + Math.max(maximum[node * 2], maximum[node * 2 + 1]);
  }
  function query(node, first, last, start, end, carried = 0) {
    if (start <= first && last <= end) return carried + maximum[node];
    const middle = (first + last) >> 1;
    let value = 0;
    carried += lazy[node];
    if (start <= middle) value = query(node * 2, first, middle, start, end, carried);
    if (end > middle) value = Math.max(value, query(node * 2 + 1, middle + 1, last, start, end, carried));
    return value;
  }
  for (const event of events) {
    if (event.delta > 0 && query(1, 0, 16383, event.first, event.last) > 0)
      throw refusal('hyperlink-bad-input', 'Clear Hyperlinks target rectangles must be disjoint.');
    update(1, 0, 16383, event.first, event.last, event.delta);
  }
}
export function clearHyperlinksCommand(value) {
  if (
    !plain(value) ||
    Object.keys(value).length !== fields.length ||
    fields.some((key) => !Object.hasOwn(value, key)) ||
    value.type !== 'clearHyperlinks' ||
    !sheetName(value.sheet) ||
    !Array.isArray(value.ranges) ||
    !value.ranges.length ||
    value.ranges.length > CLEAR_HYPERLINKS_LIMITS.ranges
  )
    throw refusal(
      'internal-invalid-field',
      'Invalid Clear Hyperlinks command: use one worksheet and 1–10,000 disjoint rectangles.'
    );
  const areas = value.ranges.map(area);
  disjoint(areas);
  areas.sort(
    (a, b) =>
      a.first.row - b.first.row || a.first.col - b.first.col || a.last.row - b.last.row || a.last.col - b.last.col
  );
  const command = { type: 'clearHyperlinks', sheet: value.sheet, ranges: areas.map((a) => a.range) };
  if (
    new TextEncoder().encode(JSON.stringify({ operation: 'apply', commands: [command] })).length >
    CLEAR_HYPERLINKS_LIMITS.actionBytes
  )
    throw refusal('hyperlink-limit', 'Clear Hyperlinks action exceeds its 2 MiB bound.');
  return command;
}
/** Capture the complete batch before queueing. Canonicalized commands remain
 * compact durable metadata commands, never cell/style projections. */
export function normalizeClearHyperlinksCommands(commands, view) {
  if (!Array.isArray(commands) || !commands.length)
    throw refusal('internal-invalid-field', 'Invalid workbook command batch');
  const copied = structuredClone(commands),
    normalized = copied.map((command) => {
      if (command?.type !== 'clearHyperlinks') return command;
      const accepted = clearHyperlinksCommand(command);
      if (
        view &&
        (!Array.isArray(view.sheets) ||
          view.sheets.filter((sheet) => sheet.name === accepted.sheet).length !== 1 ||
          !view.sheets.find((sheet) => sheet.name === accepted.sheet)?.path)
      )
        throw refusal('hyperlink-bad-input', 'Clear Hyperlinks requires an available original worksheet.');
      return accepted;
    });
  if (
    normalized.some((command) => command?.type === 'clearHyperlinks') &&
    new TextEncoder().encode(JSON.stringify({ operation: 'apply', commands: normalized })).length >
      CLEAR_HYPERLINKS_LIMITS.actionBytes
  )
    throw refusal('hyperlink-limit', 'Clear Hyperlinks action exceeds its 2 MiB bound.');
  return normalized;
}
