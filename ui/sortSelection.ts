/** Sort drafts stay bound to their original selection; the owner admits rows. */
import type { WorkbookCell, WorkbookCommand, WorkbookSortKey } from '../core/contracts';
import { normalizeSortKeys } from '../core/sort-keys.mjs';
import { addressPosition, colName } from './sheetModel';
import type { HomeBinding } from './homeCommands';
export interface SortContext {
  binding: HomeBinding;
  serial: number;
  caption: string;
  initialColumn: string;
  initialHeader: boolean;
  isCurrent: (binding: HomeBinding) => boolean;
  isRuntimeCurrent: (binding: HomeBinding) => boolean;
  disabledReason: () => string | undefined;
  read: (range: string) => Promise<Record<string, WorkbookCell>>;
  accept: (commands: WorkbookCommand[]) => Promise<void>;
}
export class SortSelectionRetired extends Error {
  constructor() {
    super('The workbook or selection changed. Reopen Sort.');
  }
}
export function sortContextKey(context: SortContext): string {
  return JSON.stringify([context.serial, context.binding]);
}
function rangeBounds(selection: string): {
  first: { row: number; col: number };
  last: { row: number; col: number };
  size: number;
} {
  if (!/^[A-Z]{1,3}[1-9][0-9]{0,6}(?::[A-Z]{1,3}[1-9][0-9]{0,6})?$/.test(selection))
    throw Error('Select a rectangular range to sort.');
  const parts = selection.split(':'),
    first = addressPosition(parts[0]),
    last = addressPosition(parts[1] ?? parts[0]);
  if (last.row < first.row || last.col < first.col || last.row >= 1048576 || last.col >= 16384)
    throw Error('Select a rectangular range to sort.');
  return { first, last, size: (last.row - first.row + 1) * (last.col - first.col + 1) };
}
export function sortSelectionColumns(selection: string): string[] {
  const b = rangeBounds(selection);
  if (!b.size || b.size > 10000) throw Error('Select a rectangular range of up to 10,000 cells to sort.');
  return Array.from({ length: b.last.col - b.first.col + 1 }, (_, index) => colName(b.first.col + index));
}
export function buildSelectionSort(
  sheet: string,
  selection: string,
  header: boolean,
  keys: readonly WorkbookSortKey[]
): WorkbookCommand {
  if (typeof header !== 'boolean') throw Error('Choose whether the selection has a header row.');
  sortSelectionColumns(selection);
  const b = rangeBounds(selection),
    first = b.first.row + (header ? 1 : 0);
  if (b.last.row - first < 1) throw Error('Select at least two data rows to sort.');
  const range = `${colName(b.first.col)}${first + 1}:${colName(b.last.col)}${b.last.row + 1}`;
  const edit = { operation: 'sortRows' as const, sheet, range, keys: keys.map((key) => ({ ...key })) };
  return { ...edit, keys: [...normalizeSortKeys(edit)] };
}
function assertCurrent(context: SortContext): void {
  if (!context.isCurrent(context.binding)) throw new SortSelectionRetired();
  const reason = context.disabledReason();
  if (reason) throw Error(reason);
}
export async function readSortHeaders(context: SortContext, signal?: AbortSignal): Promise<Record<string, string>> {
  assertCurrent(context);
  const b = rangeBounds(context.binding.selection),
    columns = sortSelectionColumns(context.binding.selection),
    labels: Record<string, string> = {};
  for (let index = 0; index < columns.length; index += 1000) {
    if (signal?.aborted) throw new SortSelectionRetired();
    assertCurrent(context);
    const chunk = columns.slice(index, index + 1000),
      range = `${chunk[0]}${b.first.row + 1}:${chunk.at(-1)}${b.first.row + 1}`;
    const cells = await context.read(range);
    assertCurrent(context);
    if (signal?.aborted) throw new SortSelectionRetired();
    const allowed = new Set(chunk.map((column) => column + (b.first.row + 1)));
    if (
      !cells ||
      typeof cells !== 'object' ||
      Array.isArray(cells) ||
      Object.entries(cells).some(
        ([ref, cell]) => !allowed.has(ref) || !cell || typeof cell !== 'object' || Array.isArray(cell)
      )
    )
      throw Error('The header reply did not match the selected row.');
    for (const column of chunk) {
      const cell = cells[column + (b.first.row + 1)];
      if (cell && !cell.formula && !cell.calculation?.pending && typeof cell.value === 'string' && cell.value.trim())
        labels[column] = cell.value.trim().slice(0, 100);
    }
  }
  return labels;
}
export async function applySelectionSort(
  context: SortContext,
  header: boolean,
  keys: readonly WorkbookSortKey[]
): Promise<{ accepted: true; runtimeCurrent: boolean }> {
  assertCurrent(context);
  const command = buildSelectionSort(context.binding.sheet, context.binding.selection, header, keys);
  assertCurrent(context);
  const pending = context.accept([command]);
  if (!pending || typeof pending.then !== 'function') throw Error('Sorting requires an acknowledged workbook command.');
  await pending;
  return { accepted: true, runtimeCurrent: context.isRuntimeCurrent(context.binding) };
}
