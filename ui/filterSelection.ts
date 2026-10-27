/** UI receipts describe accepted filters, never a projected hidden-row baseline. */
import type { WorkbookCommand, WorkbookFilterValuePage, WorkbookSheet } from '../core/contracts';
import type { HomeBinding } from './homeCommands';
export type AcceptedFilter = NonNullable<WorkbookSheet['filter']>;
export type FilterAction = 'reapply' | 'clear' | 'remove';
export interface FilterContext {
  binding: HomeBinding;
  definition: AcceptedFilter | undefined;
  isCurrent: () => boolean;
  isRuntimeCurrent: () => boolean;
  disabledReason: () => string | undefined;
  accept: (commands: WorkbookCommand[]) => Promise<void>;
}
export class FilterSelectionRetired extends Error {
  constructor() {
    super('The workbook or filter selection changed. Load the values again.');
  }
}
const own = (value: object, key: string) => Object.prototype.hasOwnProperty.call(value, key);
const record = (value: unknown): value is Record<string, unknown> =>
  !!value && typeof value === 'object' && !Array.isArray(value);
const scalar = (value: unknown): value is string | number | boolean | null =>
  value === null ||
  typeof value === 'string' ||
  typeof value === 'boolean' ||
  (typeof value === 'number' && Number.isFinite(value));
function scalarKey(value: unknown): unknown {
  return !scalar(value)
    ? ['invalid']
    : typeof value === 'number'
      ? [Object.is(value, -0) ? '-0' : String(value), 'number']
      : [value, typeof value];
}
export function filterDefinitionKey(value: AcceptedFilter | undefined): string {
  return JSON.stringify(
    value
      ? [
          value.range,
          value.column,
          Array.isArray(value.values) ? value.values.map(scalarKey) : value.values === undefined ? null : 'invalid',
          value.owned,
          value.revision,
          value.canReapply,
          value.canClear,
          value.reason,
        ]
      : null
  );
}
export function acceptedFilterReason(value: AcceptedFilter | undefined, action: FilterAction): string | undefined {
  if (!value) return 'This worksheet has no current filter.';
  const allowed = ['range', 'column', 'values', 'owned', 'revision', 'canReapply', 'canClear', 'reason'];
  if (
    !record(value) ||
    Object.keys(value).some((key) => !allowed.includes(key)) ||
    typeof value.range !== 'string' ||
    !value.range ||
    typeof value.owned !== 'boolean' ||
    typeof value.canReapply !== 'boolean' ||
    typeof value.canClear !== 'boolean' ||
    (value.reason !== undefined && typeof value.reason !== 'string')
  )
    return 'The current filter definition is unavailable.';
  if (!value.owned || typeof value.revision !== 'string' || !/^[a-f0-9]{64}$/.test(value.revision))
    return value.reason || 'This filter cannot be changed here.';
  if (action === 'remove') {
    if (
      (value.column !== undefined && typeof value.column !== 'string') ||
      (value.values !== undefined &&
        (!Array.isArray(value.values) || value.values.length > 1000 || value.values.some((item) => !scalar(item))))
    )
      return 'The current filter definition is unavailable.';
    return;
  }
  if (
    !value.column ||
    !Array.isArray(value.values) ||
    value.values.length > 1000 ||
    value.values.some((item) => !scalar(item))
  )
    return 'The current filter definition is unavailable.';
  if (action === 'reapply' ? !value.canReapply : !value.canClear)
    return (
      value.reason ||
      (action === 'reapply' ? 'This filter cannot be reapplied here.' : 'This filter cannot be cleared here.')
    );
}
export function acceptedFilterCommand(
  sheet: string,
  value: AcceptedFilter | undefined,
  action: FilterAction
): WorkbookCommand {
  const reason = acceptedFilterReason(value, action);
  if (reason) throw Error(reason);
  return {
    operation: action === 'reapply' ? 'reapplyFilter' : action === 'clear' ? 'clearFilter' : 'removeFilter',
    sheet,
    expectedFilterRevision: value!.revision!,
  };
}
export function validateFilterValuePage(value: unknown, offset: number): WorkbookFilterValuePage {
  const invalid = () => {
    throw Error('Invalid filter value page.');
  };
  if (
    !record(value) ||
    Object.keys(value).length !== 4 ||
    !['values', 'total', 'nextOffset', 'protected'].every((key) => own(value, key)) ||
    !Array.isArray(value.values) ||
    value.values.length > 100 ||
    !Number.isSafeInteger(value.total) ||
    Number(value.total) < value.values.length ||
    (value.nextOffset !== null && (!Number.isSafeInteger(value.nextOffset) || Number(value.nextOffset) <= offset)) ||
    !record(value.protected)
  )
    return invalid();
  const protectedKeys = ['errors', 'unresolved', 'special', 'unsafeNumbers'];
  if (
    Object.keys(value.protected).length !== 4 ||
    protectedKeys.some(
      (key) =>
        !own(value.protected as object, key) ||
        !Number.isSafeInteger((value.protected as Record<string, unknown>)[key]) ||
        Number((value.protected as Record<string, unknown>)[key]) < 0
    )
  )
    return invalid();
  const keys = new Set<string>();
  for (const item of value.values) {
    if (
      !record(item) ||
      Object.keys(item).length !== 3 ||
      !['key', 'value', 'count'].every((key) => own(item, key)) ||
      typeof item.key !== 'string' ||
      !item.key ||
      keys.has(item.key) ||
      !scalar(item.value) ||
      !Number.isSafeInteger(item.count) ||
      Number(item.count) < 1
    )
      return invalid();
    keys.add(item.key);
  }
  return structuredClone(value) as unknown as WorkbookFilterValuePage;
}
export async function acceptFilterCommand(
  context: FilterContext,
  command: WorkbookCommand
): Promise<{ accepted: true; runtimeCurrent: boolean }> {
  if (!context.isCurrent()) throw new FilterSelectionRetired();
  const reason = context.disabledReason();
  if (reason) throw Error(reason);
  const pending = context.accept([command]);
  if (!pending || typeof pending.then !== 'function')
    throw Error('Filtering requires an acknowledged workbook command.');
  await pending;
  return { accepted: true, runtimeCurrent: context.isRuntimeCurrent() };
}
