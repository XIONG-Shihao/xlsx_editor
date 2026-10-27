import { convertTextField } from './text-conversion.mjs';
import { parseCellInput } from './cell-input.mjs';
import { cellPosition, columnName, rangeBounds } from './workbook-commands.mjs';
import { projectCellCommands } from './workbook-projection.mjs';
import {
  PASTE_PAGE_BYTES,
  PASTE_PAGE_CELLS,
  pasteBytes,
  validatePasteDescriptor,
  validatePasteCells,
} from './paste-protocol.mjs';
import { refusal } from './refusal-codes.mjs';

const own = (object, key) => Object.hasOwn(object, key);
const nativeValue = (value) =>
  JSON.parse(
    JSON.stringify(value, (_key, item) =>
      typeof item === 'number' && (!Number.isFinite(item) || Object.is(item, -0))
        ? { $workbookNumber: Object.is(item, -0) ? '-0' : String(item) }
        : item
    )
  );
const ref = (row, col) => columnName(col) + (row + 1);
function decodeCell(cell, styles) {
  const result = { ...(styles[cell.styleId ?? 0] ?? { wrapText: false }), ...cell };
  for (const key of ['value', 'cachedValue'])
    if (result[key] && typeof result[key] === 'object' && own(result[key], '$workbookNumber'))
      result[key] = Number(result[key].$workbookNumber);
  return result;
}
/** Greedy bounded bounding boxes reduce native calls for dense or narrow pages.
 * Even a high-column ragged paste never reads more than 1,000 positions at once. */
function targetRanges(entries, origin, columns) {
  const ranges = [];
  let bounds;
  for (const [offset] of entries) {
    const row = origin.row + Math.floor(offset / columns),
      col = origin.col + (offset % columns);
    const next = bounds
      ? {
          r0: Math.min(bounds.r0, row),
          r1: Math.max(bounds.r1, row),
          c0: Math.min(bounds.c0, col),
          c1: Math.max(bounds.c1, col),
        }
      : { r0: row, r1: row, c0: col, c1: col };
    if (bounds && (next.r1 - next.r0 + 1) * (next.c1 - next.c0 + 1) > PASTE_PAGE_CELLS) {
      ranges.push(ref(bounds.r0, bounds.c0) + ':' + ref(bounds.r1, bounds.c1));
      bounds = { r0: row, r1: row, c0: col, c1: col };
    } else bounds = next;
  }
  if (bounds) ranges.push(ref(bounds.r0, bounds.c0) + ':' + ref(bounds.r1, bounds.c1));
  return ranges;
}
async function readTargets(working, sheet, ranges, wanted, native, check) {
  const cells = {};
  for (const range of ranges) {
    check();
    const result = await native({
      action: 'workbook',
      operation: 'range',
      root: working.root,
      version: working.version,
      sheet: sheet.path,
      range,
      offset: 0,
      limit: PASTE_PAGE_CELLS,
    });
    check();
    const bounds = rangeBounds(range),
      seen = new Set();
    if (!Array.isArray(result.cells) || result.cells.length > PASTE_PAGE_CELLS || bounds.size > PASTE_PAGE_CELLS)
      throw refusal('package-limit', 'Native paste target page exceeded its declared bound');
    for (const entry of result.cells) {
      const position = Array.isArray(entry) && typeof entry[0] === 'string' ? cellPosition(entry[0]) : null;
      if (
        !position ||
        position.row < bounds.first.row ||
        position.row > bounds.last.row ||
        position.col < bounds.first.col ||
        position.col > bounds.last.col ||
        seen.has(entry[0])
      )
        throw refusal('package-refused', 'Native paste target page has duplicate or out-of-range positions');
      seen.add(entry[0]);
    }
    // Native range advertises another offset at exactly its transport limit.
    // A bounded geometry with every unique position present is already complete.
    if (
      result.nextOffset !== null &&
      !(result.nextOffset === result.cells.length && result.cells.length === bounds.size)
    )
      throw refusal('package-limit', 'Native paste target page exceeded its declared bound');
    for (const [address, cell] of result.cells)
      if (wanted.has(address)) cells[address] = decodeCell(cell, working.styles ?? []);
  }
  return cells;
}
/** Apply only private native versions. The session owns final calculation,
 * accepted-state publication and its one durable descriptor/undo frame. */
export async function applyIndexedImport({
  descriptor: input,
  pages,
  current,
  native,
  checkCancelled = () => {},
  progress = () => {},
}) {
  const descriptor = validatePasteDescriptor(input),
    { shape } = descriptor,
    origin = cellPosition(shape.start);
  if (
    !current ||
    current.view?.readOnly ||
    !current.view?.sheets?.some((sheet) => sheet.name === shape.sheet) ||
    typeof native !== 'function' ||
    !pages?.[Symbol.asyncIterator]
  )
    throw refusal('package-malformed', 'Invalid private indexed paste context');
  const working = {
    ...current,
    view: { ...current.view, sheets: current.view.sheets.map((sheet) => ({ ...sheet, cells: {} })) },
  };
  const target = working.view.sheets.find((sheet) => sheet.name === shape.sheet);
  const typed = descriptor.policy === 'typed-text-v1';
  if (typed && target.date1904 !== descriptor.conversion.date1904)
    throw refusal('package-stale', 'Typed query target calendar changed.');
  let sequence = 0,
    previous = -1,
    count = 0,
    maxColumn = -1,
    formulaDelta = 0,
    formulaChanged = false,
    peakPageCells = 0,
    peakReadPositions = 0,
    patchCalls = 0;
  const check = () => checkCancelled();
  for await (const page of pages) {
    check();
    if (!page || Object.keys(page).some((key) => !['sequence', 'cells'].includes(key)) || page.sequence !== sequence++)
      throw refusal('package-bad-input', 'Staged paste replay sequence is not advancing');
    const validation = validatePasteCells(page.cells, shape, previous);
    previous = validation.lastOffset;
    const entries = page.cells.map(([offset, text]) => [offset, text]);
    for (const [offset] of entries) maxColumn = Math.max(maxColumn, offset % shape.columns);
    if (count + entries.length > descriptor.count)
      throw refusal('package-limit', 'Staged paste replay exceeds declared count');
    peakPageCells = Math.max(peakPageCells, entries.length);
    const ranges = targetRanges(entries, origin, shape.columns);
    for (const range of ranges) {
      const [a, b] = range.split(':').map(cellPosition);
      peakReadPositions = Math.max(peakReadPositions, (b.row - a.row + 1) * (b.col - a.col + 1));
    }
    const wanted = new Set(
      entries.map(([offset]) =>
        ref(origin.row + Math.floor(offset / shape.columns), origin.col + (offset % shape.columns))
      )
    );
    const cells = await readTargets(working, target, ranges, wanted, native, check);
    const commands = entries.map(([offset, text]) => {
      const address = ref(origin.row + Math.floor(offset / shape.columns), origin.col + (offset % shape.columns)),
        prior = cells[address];
      if (typed) {
        const row = Math.floor(offset / shape.columns),
          column = offset % shape.columns,
          conversion = descriptor.conversion,
          padded = row >= conversion.outputRows || column >= conversion.outputColumns;
        if (padded && text !== '')
          throw refusal('package-bad-input', 'Typed query clearing padding must be empty text.');
        const policy = padded || (conversion.keepHeader && row === 0) ? 'text' : conversion.policies[column],
          result = convertTextField(text, policy, { date1904: conversion.date1904, padded });
        if (result.status !== 'converted')
          throw refusal('package-refused', `Typed query field row ${row + 1}, column ${column + 1}: ${result.reason}`);
        return {
          sheet: shape.sheet,
          cell: address,
          value: result.value,
          ...(result.numberFormat ? { numberFormat: result.numberFormat } : {}),
        };
      }
      const literal = descriptor.policy === 'text' || prior?.format === 49 || prior?.format === '@';
      return { sheet: shape.sheet, cell: address, ...parseCellInput(text, literal) };
    });
    const view = {
      ...working.view,
      sheets: working.view.sheets.map((sheet) => ({ ...sheet, cells: sheet.name === shape.sheet ? cells : {} })),
    };
    // Genuine normalized edits reach native patch guards, including special
    // formula follower ranges that are not marked on individual read cells.
    const projected = projectCellCommands(view, commands);
    if (!projected) throw refusal('package-bad-input', 'Staged paste requires ordinary cell projection');
    const byCell = new Map(projected.edits.map((edit) => [edit.cell, [edit]]));
    let batch = [],
      bytes = 2;
    const patch = async () => {
      if (!batch.length) return;
      check();
      const result = await native({
        action: 'workbook',
        operation: 'patch',
        root: working.root,
        version: working.version,
        changes: batch,
      });
      check();
      if (!Number.isSafeInteger(result.version) || result.version < 0)
        throw refusal('internal-invalid-field', 'Invalid native paste candidate version');
      working.version = result.version;
      patchCalls++;
      batch = [];
      bytes = 2;
    };
    for (const change of projected.changes) {
      const old = cells[change.cell],
        delta =
          Number(!!(change.value.formula || change.value.formulaType)) - Number(!!(old?.formula || old?.formulaType));
      formulaDelta += delta;
      formulaChanged ||= !!(change.value.formula || old?.formula || old?.formulaType);
      const value = nativeValue({ ...change, sheet: target.path, edits: byCell.get(change.cell) }),
        size = pasteBytes(JSON.stringify(value)) + 1;
      if (size > PASTE_PAGE_BYTES)
        throw refusal('package-limit', 'Projected paste cell exceeds native page byte limit');
      if (batch.length && (batch.length >= 256 || bytes + size > PASTE_PAGE_BYTES)) await patch();
      batch.push(value);
      bytes += size;
    }
    await patch();
    count += entries.length;
    check();
    progress({ phase: 'applying-paste', completed: count, total: descriptor.count });
  }
  check();
  if (
    count !== descriptor.count ||
    Math.floor(previous / shape.columns) + 1 !== shape.rows ||
    maxColumn + 1 !== shape.columns
  )
    throw refusal('package-unsupported', 'Staged paste replay is incomplete or has a different shape');
  target.storedFormulaCount = (target.storedFormulaCount ?? 0) + formulaDelta;
  target.rows = Math.max(target.rows ?? 1, origin.row + shape.rows);
  target.columns = Math.max(target.columns ?? 1, origin.col + shape.columns);
  const range = shape.start + ':' + ref(origin.row + shape.rows - 1, origin.col + shape.columns - 1);
  return {
    working,
    version: working.version,
    view: working.view,
    changedRanges: [{ sheet: shape.sheet, range }],
    contentChanged: true,
    formulaChanged,
    formulaDelta,
    editCount: count,
    journal: descriptor,
    diagnostics: { peakPageCells, peakReadPositions, patchCalls },
  };
}
