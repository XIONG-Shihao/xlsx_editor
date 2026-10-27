import { createHash } from 'node:crypto';
import {
  readWorkbook,
  editWorkbookResult,
  persistCalculation,
  assertStructuralEdit,
  applyProjectedOperation,
  encodeWorkbookString,
  readWorkbookFilterStates,
} from './workbook-package.mjs';
import * as engine from './calculation-node.mjs';
import { projectWorkbookCommands, projectCellCommands } from './workbook-projection.mjs';
import { rangeBounds, columnName, cellPosition, keyCellValue, resolveWorkbookEdits } from './workbook-commands.mjs';
import { refusal } from './refusal-codes.mjs';

export const WORKBOOK_PROFILE = 'xlsx.oss.v1';
const digest = (bytes) => createHash('sha256').update(bytes).digest('hex');
const { calculateWorkbook } = engine;
function bindCalculation(view, revision) {
  // This is the unpublished result of calculating a newly serialized batch.
  // No accepted view or prior receipt owns it; rebinding it in place avoids
  // another full cell-index allocation while the preserving writer is live.
  view.revision = revision;
  for (const sheet of view.sheets)
    for (const ref in sheet.cells) {
      const cell = sheet.cells[ref];
      if (!cell.calculation) continue;
      Object.assign(cell, persistedCache(cell.calculation));
      cell.calculation.revision = revision;
    }
  return view;
}
function persistedCache(result) {
  if (result.status === 'calculated') {
    const value = result.value ?? '';
    return {
      value,
      cachedValue: value,
      rawValue:
        typeof value === 'boolean'
          ? String(Number(value))
          : typeof value === 'string'
            ? encodeWorkbookString(value)
            : String(value),
      dataType: typeof value === 'boolean' ? 'boolean' : typeof value === 'string' ? 'text' : 'number',
    };
  }
  if (result.status === 'error')
    return {
      value: result.error,
      cachedValue: result.error,
      rawValue: result.error,
      dataType: 'error',
      error: result.error,
    };
  return { value: null, cachedValue: null, rawValue: null, dataType: 'blank' };
}

export async function readCalculatedWorkbook(bytes, options) {
  return calculateWorkbook(readWorkbook(bytes), { ...options, takeOwnership: true });
}

/** Accepted batches only replace state after validation, calculation and serialization.
 * The desktop recovery and MCP durable-operation owners persist their own receipts.
 */
export class WorkbookSession {
  #bytes;
  #view;
  #receipts = new Map();
  #busy = false;
  constructor(bytes) {
    this.#bytes = bytes;
  }
  async read() {
    return (this.#view ??= await readCalculatedWorkbook(this.#bytes));
  }
  get bytes() {
    return this.#bytes;
  }
  async apply({ operationId, expectedRevision, commands }, { signal } = {}) {
    if (typeof operationId !== 'string' || !operationId || operationId.length > 150)
      throw refusal('internal-invalid-field', 'Invalid operation ID');
    const fingerprint = digest(
      JSON.stringify({ expectedRevision, commands }, (_key, value) =>
        Object.is(value, -0) ? { $workbookNumber: '-0' } : value
      )
    );
    const receipt = this.#receipts.get(operationId);
    if (receipt) {
      if (receipt.fingerprint !== fingerprint)
        throw refusal('workbook-duplicate', 'Conflict: operation ID was already used with other arguments');
      return receipt.result;
    }
    if (this.#busy) throw refusal('session-busy', 'Workbook is busy; retry after the accepted batch');
    if (this.#receipts.size >= 10000) throw refusal('workbook-limit', 'Workbook operation receipt limit reached');
    this.#busy = true;
    try {
      if (signal?.aborted) throw refusal('operation-cancelled', 'Workbook edit cancelled');
      const before = await this.read();
      if (before.revision !== expectedRevision.replace(/^sha256:/, ''))
        throw refusal('workbook-stale', 'Conflict: workbook revision is stale');
      let edited = await editCommandBatch(this.#bytes, commands, before);
      const calculated = await calculateWorkbook(edited.view, { takeOwnership: true });
      edited.view = undefined;
      const bytes = persistCalculation(edited.bytes, calculated, edited.context);
      const changedCells = edited.changedCells;
      const acceptedEdits = edited.acceptedEdits;
      edited = null;
      const revision = digest(bytes);
      const view = bindCalculation(calculated, revision);
      if (signal?.aborted) throw refusal('operation-cancelled', 'Workbook edit cancelled');
      const affected = [];
      for (const sheet of view.sheets) {
        const previous = before.sheets.find((old) => old.name === sheet.name);
        const refs = changedCells ? (changedCells[sheet.name] ?? []) : Object.keys(sheet.cells);
        for (const cell of refs) {
          const after = sheet.cells[cell];
          const old = previous?.cells[cell];
          const state = (item) =>
            JSON.stringify({
              value: item?.value,
              formula: item?.formula,
              format: item?.format,
              bold: item?.bold,
              italic: item?.italic,
              calculation: item?.calculation
                ? { status: item.calculation.status, value: item.calculation.value, error: item.calculation.error }
                : undefined,
            });
          if (state(old) !== state(after))
            affected.push({
              sheet: sheet.name,
              cell,
              before: { value: old?.value ?? null, formula: old?.formula ?? null },
              after: { value: after.value, formula: after.formula ?? null, calculation: after.calculation },
            });
        }
      }
      const result = {
        bytes,
        view,
        operationId,
        affected,
        edits: acceptedEdits,
        profile: WORKBOOK_PROFILE,
        calculation: calculationSummary(view),
      };
      this.#bytes = bytes;
      this.#view = view;
      this.#receipts.set(operationId, { fingerprint, result });
      return result;
    } finally {
      this.#busy = false;
    }
  }
}

async function editCommandBatch(input, commands, acceptedView) {
  if (!Array.isArray(commands) || !commands.length)
    throw refusal('internal-invalid-field', 'Invalid workbook command batch');
  // Projection keeps its input immutable. Reuse the accepted view rather than
  // parsing the full package again; each writer still validates source bytes,
  // expected revision and guards before returning a new, private view.
  let bytes = input,
    view = acceptedView,
    pending = [],
    context,
    changedCells = {};
  const acceptedEdits = [];
  const recordChanges = (result) => {
    if (!result.changedCells) {
      changedCells = undefined;
      return;
    }
    if (changedCells)
      for (const [sheet, refs] of Object.entries(result.changedCells))
        changedCells[sheet] = [...new Set([...(changedCells[sheet] ?? []), ...refs])];
  };
  const flush = async () => {
    if (!pending.length) return;
    context = undefined;
    // Ordinary cell edits need only addressed coordinates. The preserving writer
    // still checks actual package values; an unused full projection otherwise
    // duplicates every accepted cell while the worksheet DOM is being built.
    let projected = pending.every(
      (command) =>
        !command.operation && [undefined, 'setCells', 'formatRange', 'updateRows', 'recalculate'].includes(command.type)
    )
      ? { edits: resolveWorkbookEdits(view, pending) }
      : (projectCellCommands(view, pending) ?? (await projectWorkbookCommands(view, pending, engine)));
    const edits = projected.edits;
    projected = null;
    const result = editWorkbookResult(bytes, edits, view.revision, true, undefined, true, { sourceView: view });
    recordChanges(result);
    acceptedEdits.push(...edits);
    bytes = result.bytes;
    view = result.view;
    context = result.context;
    pending = [];
  };
  for (const command of commands) {
    if (
      [
        'addSheet',
        'renameSheet',
        'duplicateSheet',
        'mergeCells',
        'unmergeCells',
        'addConditionalRule',
        'createTable',
      ].includes(command.operation)
    ) {
      await flush();
      context = undefined;
      await projectWorkbookCommands(view, [command], engine);
      const result = editWorkbookResult(bytes, [command], view.revision, true, undefined, true);
      recordChanges(result);
      acceptedEdits.push(command);
      bytes = result.bytes;
      view = result.view;
      context = result.context;
    } else if (command.operation === 'insertRow') {
      await flush();
      context = undefined;
      assertStructuralEdit(bytes, command);
      const transformed = await projectWorkbookCommands(view, [command], engine);
      const result = editWorkbookResult(bytes, [command], view.revision, true, transformed.view, true);
      recordChanges(result);
      acceptedEdits.push(command);
      bytes = result.bytes;
      view = result.view;
      context = result.context;
    } else if (
      [
        'sortRows',
        'filterRows',
        'reapplyFilter',
        'clearFilter',
        'removeFilter',
        'resizeColumn',
        'resizeRow',
        'freeze',
        'unhideAll',
      ].includes(command.operation)
    ) {
      await flush();
      context = undefined;
      const projected = await projectWorkbookCommands(view, [command], engine, {
        filterStates: readWorkbookFilterStates(bytes),
      });
      const result = applyProjectedOperation(bytes, command, projected.view, true);
      recordChanges(result);
      acceptedEdits.push(command);
      bytes = result.bytes;
      view = result.view;
      context = result.context;
    } else pending.push(command);
  }
  await flush();
  return { bytes, view, context, changedCells, acceptedEdits };
}

export function calculationSummary(view) {
  const counts = { calculated: 0, unresolved: 0, error: 0, circular: 0 };
  for (const sheet of view.sheets)
    for (const ref in sheet.cells) {
      const cell = sheet.cells[ref];
      if (cell.formula || cell.formulaType) counts[cell.calculation?.status ?? 'unresolved']++;
    }
  return { revision: view.revision, engine: view.calculationEngine, ...counts };
}

export function readWorkbookRange(view, { sheet: name, range = 'A1:J20', offset = 0, limit = 100, keys }) {
  const sheet = view.sheets.find((item) => item.name === name);
  if (!sheet) throw refusal('workbook-not-found', `Sheet not found: ${name}`);
  if (!Number.isInteger(offset) || offset < 0 || !Number.isInteger(limit) || limit < 1 || limit > 500)
    throw refusal('internal-invalid-field', 'Invalid read window');
  const bounds = rangeBounds(range);
  let total = bounds.size,
    selectedRows;
  const missingKeys = [],
    duplicateKeys = [];
  if (keys) {
    const column = keys.column;
    const keyColumn = cellPosition(`${column}1`).col;
    if (keyColumn < bounds.first.col || keyColumn > bounds.last.col)
      throw refusal('workbook-unsupported', 'Key column is outside the requested range');
    const negativeZero = Symbol('negative-zero-key'),
      identity = (value) => (Object.is(value, -0) ? negativeZero : value);
    const matches = new Map(keys.values.map((value) => [identity(value), []])),
      rows = new Set();
    for (let row = bounds.first.row; row <= bounds.last.row; row++) {
      const value = keyCellValue(sheet.cells[`${columnName(keyColumn)}${row + 1}`]);
      matches.get(identity(value))?.push(row);
    }
    for (const value of keys.values) {
      const found = matches.get(identity(value));
      if (!found.length) missingKeys.push(value);
      if (found.length > 1) duplicateKeys.push(value);
      for (const row of found) rows.add(row);
    }
    selectedRows = [...rows].sort((a, b) => a - b);
    total = selectedRows.length * bounds.columns;
  }
  const items = Array.from({ length: Math.max(0, Math.min(limit, total - offset)) }, (_, index) => {
    const position = offset + index,
      row = selectedRows
        ? selectedRows[Math.floor(position / bounds.columns)]
        : bounds.first.row + Math.floor(position / bounds.columns);
    const ref = `${columnName(bounds.first.col + (position % bounds.columns))}${row + 1}`;
    const cell = sheet.cells[ref];
    return {
      sheet: name,
      cell: ref,
      target: { sheet: name, cell: ref, revision: view.revision },
      value: cell?.value ?? null,
      formula: cell?.formula ?? null,
      formulaType: cell?.formulaType ?? null,
      sharedFormulaId: cell?.sharedFormulaId ?? null,
      formulaReference: cell?.formulaReference ?? null,
      dataType: cell?.dataType ?? 'blank',
      rawValue: cell?.rawValue ?? null,
      cachedValue: cell?.cachedValue ?? null,
      calculation: cell?.calculation ?? null,
      format: cell?.format ?? 'general',
      bold: cell?.bold ?? false,
      italic: cell?.italic ?? false,
      expectedValue: cell?.value ?? null,
      expectedFormula: cell?.formula ?? null,
    };
  });
  return {
    items,
    total,
    offset,
    ...(keys ? { missingKeys, duplicateKeys } : {}),
    nextOffset: offset + items.length < total ? offset + items.length : null,
    calculation: calculationSummary(view),
    coverage: {
      sheet: name,
      range,
      ...(keys ? { keys } : {}),
      completeDocument: false,
      omissions: ['unrequested-ranges', 'external-data-freshness', 'visual-layout'],
    },
  };
}
