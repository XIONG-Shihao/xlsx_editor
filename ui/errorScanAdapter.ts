/** Scoped readonly scan: opaque UI token binds acceptance; the captured host
 * callback supplies the actual original backend grant/session/revision. */
import type { WorkbookCell } from '../core/contracts';
import { errorScanScopeKey, type ErrorScanner, type ErrorScanScope } from './errorChecking';
export interface ErrorScanPage {
  cells: Record<string, WorkbookCell>;
  nextCursor: string | null;
}
export interface ErrorScanFrame {
  scope: ErrorScanScope;
  scan: (sheet: string, after?: string) => Promise<ErrorScanPage>;
}
function canonical(reference: unknown): reference is string {
  const match = typeof reference === 'string' && /^([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(reference);
  return (
    !!match && [...match[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) <= 16384 && Number(match[2]) <= 1048576
  );
}
/** Missing receipt fields are a failure, never an empty complete worksheet. */
export function validateErrorScanPage(value: unknown): ErrorScanPage {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw Error('The worksheet scan returned no valid page.');
  const page = value as Partial<ErrorScanPage>;
  if (
    !page.cells ||
    typeof page.cells !== 'object' ||
    Array.isArray(page.cells) ||
    !Object.hasOwn(value, 'nextCursor') ||
    (page.nextCursor !== null && !canonical(page.nextCursor)) ||
    Object.keys(page.cells).length > 2048 ||
    Object.keys(page.cells).some((ref) => !canonical(ref)) ||
    Object.values(page.cells).some((cell) => !cell || typeof cell !== 'object' || Array.isArray(cell))
  )
    throw Error('The worksheet scan returned an invalid cell page or cursor.');
  return { cells: page.cells, nextCursor: page.nextCursor as string | null };
}
export function createScopedErrorScanner(current: () => ErrorScanFrame | null): ErrorScanner {
  return async (scope, after, signal) => {
    const frame = current(),
      key = errorScanScopeKey(scope);
    if (signal.aborted || !scope.active || !frame || errorScanScopeKey(frame.scope) !== key)
      throw Error('The worksheet or workbook revision changed.');
    const result = await frame.scan(scope.sheet, after ?? undefined),
      latest = current();
    if (signal.aborted || !latest || errorScanScopeKey(latest.scope) !== key)
      throw Error('The worksheet or workbook revision changed.');
    const page = validateErrorScanPage(result);
    return { scope, after, ...page };
  };
}
