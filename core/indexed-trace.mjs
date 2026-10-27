/** Owner-private signed continuation. No retained cells, native roots or mutations. */
import { randomBytes, createHmac, createHash, timingSafeEqual } from 'node:crypto';
import { cellPosition, rangeBounds } from './workbook-commands.mjs';
import { refusal } from './refusal-codes.mjs';
export const traceCursorKey = () => randomBytes(32);
const keys = (o, wanted) =>
  o && typeof o === 'object' && !Array.isArray(o) && Object.keys(o).every((k) => wanted.includes(k));
const bound = (n, max) => Number.isSafeInteger(n) && n >= 0 && n <= max;
const text = (s) => typeof s === 'string' && s.length > 0 && s.length <= 1024 && !s.includes('\0');
export function traceRequest(sheet, cell, options = {}) {
  if (
    !keys(options, ['direction', 'after', 'limit']) ||
    !['precedents', 'dependents'].includes(options.direction) ||
    !text(sheet) ||
    !text(cell) ||
    !/^[A-Z]{1,3}[1-9][0-9]{0,6}$/.test(cell)
  )
    throw refusal('internal-invalid-field', 'Invalid formula trace request');
  cellPosition(cell);
  const limit = options.limit ?? 100;
  if (
    !Number.isInteger(limit) ||
    limit < 1 ||
    limit > 100 ||
    (options.after !== undefined && (typeof options.after !== 'string' || options.after.length > 4096))
  )
    throw refusal('internal-invalid-field', 'Invalid formula trace bounds');
  return { direction: options.direction, limit };
}
function signed(key, data) {
  return createHmac('sha256', key).update(data).digest();
}
const bindingToken = (binding) => createHash('sha256').update(JSON.stringify(binding)).digest('hex');
export function readTraceCursor(key, after, binding) {
  if (after === undefined) return { cursor: null, pages: 0, scanned: 0, unknown: false };
  const [body, signature, ...extra] = after.split('.');
  if (extra.length || !body || !signature || !/^[A-Za-z0-9_-]+$/.test(body) || !/^[A-Za-z0-9_-]{43}$/.test(signature))
    throw refusal('internal-invalid-field', 'Invalid formula trace continuation');
  const expected = signed(key, body),
    actual = Buffer.from(signature, 'base64url'),
    payload = Buffer.from(body, 'base64url');
  if (actual.toString('base64url') !== signature || payload.toString('base64url') !== body)
    throw refusal('internal-invalid-field', 'Invalid noncanonical formula trace continuation');
  if (expected.length !== actual.length || !timingSafeEqual(expected, actual))
    throw refusal('formula-stale', 'Formula trace continuation belongs to another owner');
  let data;
  try {
    data = JSON.parse(payload.toString('utf8'));
  } catch {
    throw refusal('internal-invalid-field', 'Invalid formula trace continuation');
  }
  if (
    !keys(data, ['binding', 'cursor', 'pages', 'scanned', 'unknown']) ||
    data.binding !== bindingToken(binding) ||
    !bound(data.pages, 63) ||
    data.pages < 1 ||
    !bound(data.scanned, 131072) ||
    typeof data.unknown !== 'boolean' ||
    !keys(data.cursor, ['sheetIndex', 'after', 'offset']) ||
    !bound(data.cursor.sheetIndex, 1023) ||
    !bound(data.cursor.offset, 1024) ||
    (data.cursor.after !== null && !text(data.cursor.after))
  )
    throw refusal('formula-stale', 'Formula trace continuation is stale or changed its target');
  return data;
}
export function traceReceipt(key, binding, previous, result, limit) {
  if (
    !keys(result, [
      'version',
      'sheet',
      'cell',
      'direction',
      'edges',
      'markers',
      'next',
      'scannedFormulas',
      'unknown',
      'limited',
    ]) ||
    result.version !== binding.version ||
    result.sheet !== binding.sheet ||
    result.cell !== binding.cell ||
    result.direction !== binding.direction ||
    !Array.isArray(result.edges) ||
    !Array.isArray(result.markers) ||
    result.edges.length + result.markers.length > limit ||
    !bound(result.scannedFormulas, 2048) ||
    typeof result.unknown !== 'boolean' ||
    typeof result.limited !== 'boolean'
  )
    throw refusal('internal-invalid-field', 'Invalid version-bound formula trace receipt');
  const position = (p) => {
    if (!keys(p, ['sheet', 'range']) || !binding.sheets.includes(p.sheet) || !text(p.range))
      throw refusal('internal-invalid-field', 'Invalid trace endpoint');
    rangeBounds(p.range);
  };
  for (const edge of result.edges) {
    if (
      !keys(edge, ['from', 'to', 'kind', 'via']) ||
      !['reference', 'spill-origin', 'spill-range'].includes(edge.kind) ||
      !['formula', 'name', 'table'].includes(edge.via)
    )
      throw refusal('internal-invalid-field', 'Invalid trace edge');
    position(edge.from);
    position(edge.to);
    if (binding.direction === 'precedents') {
      if (edge.to.sheet !== binding.sheet || edge.to.range !== binding.cell)
        throw refusal('formula-stale', 'Precedent edge changed its selected formula');
    } else {
      const box = rangeBounds(edge.from.range),
        at = cellPosition(binding.cell);
      if (
        edge.from.sheet !== binding.sheet ||
        at.row < box.first.row ||
        at.row > box.last.row ||
        at.col < box.first.col ||
        at.col > box.last.col
      )
        throw refusal('formula-bad-input', 'Dependent edge does not refer to its selected cell');
    }
  }
  for (const marker of result.markers) {
    if (
      !keys(marker, ['sheet', 'cell', 'kind', 'reason']) ||
      !binding.sheets.includes(marker.sheet) ||
      (marker.cell !== undefined && (!text(marker.cell) || !cellPosition(marker.cell))) ||
      !['dynamic', 'external', 'unsupported', 'special', 'limit'].includes(marker.kind) ||
      !text(marker.reason)
    )
      throw refusal('internal-invalid-field', 'Invalid trace marker');
  }
  if (
    result.next !== null &&
    (!keys(result.next, ['sheetIndex', 'after', 'offset']) ||
      !bound(result.next.sheetIndex, binding.sheets.length - 1) ||
      !bound(result.next.offset, 1024) ||
      (result.next.after !== null && !text(result.next.after)))
  )
    throw refusal('internal-invalid-field', 'Invalid native trace cursor');
  const pages = previous.pages + 1,
    scanned = previous.scanned + result.scannedFormulas,
    unknown = previous.unknown || result.unknown,
    limited = result.limited || (pages >= 64 && result.next !== null);
  let nextCursor = null;
  if (result.next !== null && !limited) {
    const body = Buffer.from(
      JSON.stringify({ binding: bindingToken(binding), cursor: result.next, pages, scanned, unknown })
    ).toString('base64url');
    nextCursor = body + '.' + signed(key, body).toString('base64url');
  }
  const markers = [...result.markers];
  if (limited && markers.length + result.edges.length < limit)
    markers.push({
      sheet: binding.sheet,
      cell: binding.cell,
      kind: 'limit',
      reason: 'Trace stopped at its64-page bounded coverage limit',
    });
  return {
    schema: 1,
    ownerId: binding.ownerId,
    state: binding.state,
    revision: binding.revision,
    sheet: binding.sheet,
    cell: binding.cell,
    direction: binding.direction,
    edges: result.edges,
    markers,
    nextCursor,
    coverage: {
      mode: 'direct-static-references',
      complete: result.next === null && !unknown && !limited,
      scannedFormulas: scanned,
      pages,
      limited,
    },
  };
}
