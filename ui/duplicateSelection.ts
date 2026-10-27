import { addressPosition, colName } from './sheetModel';
import type { HomeBinding } from './homeCommands';
import type { WorkbookCell, WorkbookDuplicatePreview } from '../core/contracts';
export interface DuplicateConfig {
  sheet: string;
  range: string;
  hasHeaders: boolean;
  columns: string[];
}
export type DuplicatePreview = WorkbookDuplicatePreview;
export type DuplicateReceipt = Pick<WorkbookDuplicatePreview, 'documentId' | 'ownerId' | 'state' | 'revision'>;
export function duplicateReceiptKey(receipt: DuplicateReceipt | undefined): string {
  return JSON.stringify(receipt ? [receipt.documentId, receipt.ownerId, receipt.state, receipt.revision] : null);
}
export interface DuplicateCommand {
  operation: 'removeDuplicates';
  sheet: string;
  range: string;
  hasHeaders: boolean;
  columns: string[];
  expectedFingerprint: string;
}
export interface DuplicateContext {
  binding: HomeBinding;
  receipt: DuplicateReceipt;
  serial: number;
  caption: string;
  initialHeader: boolean;
  isCurrent: () => boolean;
  isRuntimeCurrent: () => boolean;
  disabledReason: (action: 'preview' | 'remove') => string | undefined;
  read: (config: DuplicateConfig, offset: number) => Promise<unknown>;
  readHeaders: (range: string) => Promise<Record<string, WorkbookCell>>;
  accept: (commands: DuplicateCommand[]) => Promise<void>;
}
export class DuplicateDraftRetired extends Error {
  constructor() {
    super('The workbook or selection changed. Reopen Remove duplicates.');
  }
}
const record = (value: unknown): value is Record<string, unknown> =>
  !!value && typeof value === 'object' && !Array.isArray(value);
const exact = (value: unknown, keys: string[]): value is Record<string, unknown> =>
  record(value) && Object.keys(value).length === keys.length && keys.every((key) => Object.hasOwn(value, key));
function bounds(range: string): {
  first: { row: number; col: number };
  last: { row: number; col: number };
  height: number;
  width: number;
} {
  if (!/^[A-Z]{1,3}[1-9][0-9]{0,6}(?::[A-Z]{1,3}[1-9][0-9]{0,6})?$/.test(range))
    throw Error('Select a rectangular range to compare.');
  const pieces = range.split(':'),
    first = addressPosition(pieces[0]),
    last = addressPosition(pieces[1] ?? pieces[0]);
  if (last.row < first.row || last.col < first.col || last.row >= 1048576 || last.col >= 16384)
    throw Error('Select a rectangular range to compare.');
  const height = last.row - first.row + 1,
    width = last.col - first.col + 1;
  if (height * width > 1000000) throw Error('Select up to 1,000,000 cells to compare.');
  return { first, last, height, width };
}
export function duplicateColumns(range: string): string[] {
  const b = bounds(range);
  return Array.from({ length: b.width }, (_, index) => colName(b.first.col + index));
}
export function duplicateConfig(
  context: DuplicateContext,
  hasHeaders: boolean,
  columns: readonly string[]
): DuplicateConfig {
  const b = bounds(context.binding.selection),
    allowed = new Set(duplicateColumns(context.binding.selection));
  if (typeof hasHeaders !== 'boolean' || b.height - (hasHeaders ? 1 : 0) < 1)
    throw Error('Select at least one data row to compare.');
  if (
    !Array.isArray(columns) ||
    columns.length < 1 ||
    columns.length > 16 ||
    new Set(columns).size !== columns.length ||
    columns.some((column) => !allowed.has(column))
  )
    throw Error('Choose 1–16 different columns from the selected range.');
  return { sheet: context.binding.sheet, range: context.binding.selection, hasHeaders, columns: [...columns] };
}
export function duplicateConfigKey(config: DuplicateConfig): string {
  return JSON.stringify([config.sheet, config.range, config.hasHeaders, config.columns]);
}
export function duplicateContextKey(context: DuplicateContext): string {
  return JSON.stringify([context.serial, context.binding, context.receipt]);
}
function validateConfig(config: DuplicateConfig): void {
  if (
    !exact(config, ['sheet', 'range', 'hasHeaders', 'columns']) ||
    typeof config.sheet !== 'string' ||
    !config.sheet ||
    typeof config.hasHeaders !== 'boolean'
  )
    throw Error('Invalid duplicate-removal options.');
  const b = bounds(config.range),
    allowed = new Set(duplicateColumns(config.range));
  if (
    b.height - (config.hasHeaders ? 1 : 0) < 1 ||
    !Array.isArray(config.columns) ||
    config.columns.length < 1 ||
    config.columns.length > 16 ||
    new Set(config.columns).size !== config.columns.length ||
    config.columns.some((column) => !allowed.has(column))
  )
    throw Error('Choose 1–16 different columns from the selected range.');
}
function assertConfig(context: DuplicateContext, config: DuplicateConfig): void {
  validateConfig(config);
  if (config.sheet !== context.binding.sheet || config.range !== context.binding.selection)
    throw new DuplicateDraftRetired();
}
function assertCurrent(context: DuplicateContext, action: 'preview' | 'remove'): void {
  if (
    !context.isCurrent() ||
    !context.binding.ownerIdentity ||
    context.binding.mode !== 'normal' ||
    context.binding.specialSelection
  )
    throw new DuplicateDraftRetired();
  const reason = context.disabledReason(action);
  if (reason) throw Error(reason);
}
export function validateDuplicatePreview(
  value: unknown,
  config: DuplicateConfig,
  offset: number,
  receipt: DuplicateReceipt,
  previous?: DuplicatePreview
): DuplicatePreview {
  validateConfig(config);
  const fail = () => {
    throw Error('The duplicate preview did not match the selected rows.');
  };
  if (
    !Number.isSafeInteger(offset) ||
    offset < 0 ||
    !exact(value, [
      'schema',
      'documentId',
      'ownerId',
      'state',
      'revision',
      'sheet',
      'range',
      'hasHeaders',
      'columns',
      'fingerprint',
      'totalRows',
      'keptRows',
      'removedRows',
      'comparison',
      'samples',
      'offset',
      'nextOffset',
    ])
  )
    return fail();
  if (
    !exact(receipt, ['documentId', 'ownerId', 'state', 'revision']) ||
    (['documentId', 'ownerId', 'state', 'revision'] as const).some(
      (key) => typeof receipt[key] !== 'string' || !receipt[key] || value[key] !== receipt[key]
    ) ||
    value.schema !== 1 ||
    value.sheet !== config.sheet ||
    value.range !== config.range ||
    value.hasHeaders !== config.hasHeaders ||
    !Array.isArray(value.columns) ||
    JSON.stringify(value.columns) !== JSON.stringify(config.columns)
  )
    return fail();
  const b = bounds(config.range),
    first = b.first.row + 1 + (config.hasHeaders ? 1 : 0),
    last = b.last.row + 1;
  if (
    typeof value.fingerprint !== 'string' ||
    !/^[a-f0-9]{64}$/.test(value.fingerprint) ||
    value.comparison !== 'exact-typed-values-v1' ||
    value.offset !== offset
  )
    return fail();
  if (
    !['totalRows', 'keptRows', 'removedRows'].every(
      (key) => Number.isSafeInteger(value[key]) && Number(value[key]) >= 0
    ) ||
    value.totalRows !== b.height - (config.hasHeaders ? 1 : 0) ||
    Number(value.keptRows) + Number(value.removedRows) !== value.totalRows ||
    Number(value.keptRows) < 1 ||
    offset > Number(value.removedRows)
  )
    return fail();
  if (!Array.isArray(value.samples) || value.samples.length !== Math.min(100, Number(value.removedRows) - offset))
    return fail();
  let prior = 0;
  for (const item of value.samples) {
    if (
      !exact(item, ['row', 'keptRow']) ||
      !Number.isSafeInteger(item.row) ||
      !Number.isSafeInteger(item.keptRow) ||
      Number(item.row) <= prior ||
      Number(item.row) < first ||
      Number(item.row) > last ||
      Number(item.keptRow) < first ||
      Number(item.keptRow) >= Number(item.row)
    )
      return fail();
    prior = Number(item.row);
  }
  const end = offset + value.samples.length,
    expectedNext = end < Number(value.removedRows) ? end : null;
  if (value.nextOffset !== expectedNext) return fail();
  const result = structuredClone(value) as unknown as DuplicatePreview;
  if (
    previous &&
    (result.fingerprint !== previous.fingerprint ||
      result.totalRows !== previous.totalRows ||
      result.keptRows !== previous.keptRows ||
      result.removedRows !== previous.removedRows)
  )
    return fail();
  if (
    previous &&
    offset === previous.offset + previous.samples.length &&
    previous.samples.length &&
    result.samples.length &&
    result.samples[0].row <= previous.samples.at(-1)!.row
  )
    return fail();
  return result;
}
export async function readDuplicateHeaders(
  context: DuplicateContext,
  page: number,
  signal?: AbortSignal
): Promise<Record<string, string>> {
  assertCurrent(context, 'preview');
  if (!Number.isSafeInteger(page) || page < 0) throw Error('Choose a valid column page.');
  const b = bounds(context.binding.selection),
    columns = duplicateColumns(context.binding.selection).slice(page * 100, (page + 1) * 100);
  if (!columns.length) return {};
  const range = `${columns[0]}${b.first.row + 1}:${columns.at(-1)}${b.first.row + 1}`;
  const cells = await context.readHeaders(range);
  assertCurrent(context, 'preview');
  if (signal?.aborted) throw new DuplicateDraftRetired();
  const allowed = new Set(columns.map((column) => column + (b.first.row + 1)));
  if (!record(cells) || Object.entries(cells).some(([ref, cell]) => !allowed.has(ref) || !record(cell)))
    throw Error('The column-heading reply did not match the selection.');
  const labels: Record<string, string> = {};
  for (const column of columns) {
    const cell = cells[column + (b.first.row + 1)];
    if (cell && !cell.formula && !cell.calculation?.pending && typeof cell.value === 'string' && cell.value.trim())
      labels[column] = cell.value.trim().slice(0, 100);
  }
  return labels;
}
export async function previewDuplicates(
  context: DuplicateContext,
  config: DuplicateConfig,
  offset = 0,
  previous?: DuplicatePreview
): Promise<DuplicatePreview> {
  assertCurrent(context, 'preview');
  assertConfig(context, config);
  const result = await context.read(structuredClone(config), offset);
  assertCurrent(context, 'preview');
  return validateDuplicatePreview(result, config, offset, context.receipt, previous);
}
export async function removeDuplicates(
  context: DuplicateContext,
  config: DuplicateConfig,
  preview: DuplicatePreview
): Promise<{ accepted: true; removedRows: number; runtimeCurrent: boolean }> {
  assertCurrent(context, 'remove');
  assertConfig(context, config);
  const receipt = validateDuplicatePreview(preview, config, preview.offset, context.receipt);
  if (!receipt.removedRows) throw Error('No duplicate rows were found.');
  const command: DuplicateCommand = {
    operation: 'removeDuplicates',
    sheet: config.sheet,
    range: config.range,
    hasHeaders: config.hasHeaders,
    columns: [...config.columns],
    expectedFingerprint: receipt.fingerprint,
  };
  const pending = context.accept([command]);
  if (!pending || typeof pending.then !== 'function')
    throw Error('Removing duplicates requires an acknowledged workbook command.');
  await pending;
  return { accepted: true, removedRows: receipt.removedRows, runtimeCurrent: context.isRuntimeCurrent() };
}
