/** Bind readonly cell requests to the acknowledged owner and opaque UI state token.
 * The captured host reader owns the actual backend state; this module never parses it. */
import type { WorkbookCell } from '../core/contracts';
import type { WatchReader, WatchScope } from './watchWindow';
export interface WatchReadFrame {
  scope: WatchScope;
  read: (sheet: string, range: string) => Promise<Record<string, WorkbookCell>>;
}
export const watchScopeKey = (scope: WatchScope): string =>
  JSON.stringify([
    scope.documentId,
    scope.ownerIdentity,
    scope.state,
    scope.activationEpoch,
    scope.active,
    scope.sheets,
  ]);
export function createScopedWatchReader(current: () => WatchReadFrame | null): WatchReader {
  return async (scope, target, signal) => {
    const frame = current(),
      key = watchScopeKey(scope);
    if (
      signal.aborted ||
      !scope.active ||
      !frame ||
      watchScopeKey(frame.scope) !== key ||
      !scope.sheets.includes(target.sheet)
    )
      throw Error('The watch owner or accepted state changed.');
    const cells = await frame.read(target.sheet, target.reference + ':' + target.reference);
    const latest = current();
    if (signal.aborted || !latest || watchScopeKey(latest.scope) !== key)
      throw Error('The watch owner or accepted state changed.');
    if (
      !cells ||
      typeof cells !== 'object' ||
      Array.isArray(cells) ||
      Object.keys(cells).some((reference) => reference !== target.reference)
    )
      throw Error('The cell reply did not match the watched target.');
    return { scope, target, cell: cells[target.reference] };
  };
}
