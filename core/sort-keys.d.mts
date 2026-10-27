import type { WorkbookEdit, WorkbookSortKey } from './contracts.js';
/** Validates shape/bounds only; owner source and reference guards still apply. */
export function normalizeSortKeys(
  edit: Pick<WorkbookEdit, 'range' | 'column' | 'direction' | 'keys'>
): readonly Readonly<WorkbookSortKey>[];
