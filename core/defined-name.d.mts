import type { WorkbookView } from './contracts.js';
export function validateDefinedName(name: string): string;
export function definedNameConflict(
  name: string,
  inventory: Pick<WorkbookView, 'namedRanges' | 'tables'>,
  options?: { scopeSheet?: string; visibleSheet?: string }
): string | null;
