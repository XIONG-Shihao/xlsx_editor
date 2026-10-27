/** Selection is a view; only a captured, eligible definition can authorize refresh. */
import type { WorkbookPivotDefinition, WorkbookCommand } from '../core/contracts';
function bounds(range: string): [number, number, number, number] | null {
  const m = /^([A-Z]{1,3})([1-9]\d{0,6})(?::([A-Z]{1,3})([1-9]\d{0,6}))?$/i.exec(range);
  if (!m) return null;
  const column = (s: string) => [...s.toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  const a = column(m[1]),
    r = Number(m[2]),
    b = column(m[3] ?? m[1]),
    z = Number(m[4] ?? m[2]);
  return a <= b && r <= z && b <= 16384 && z <= 1048576 ? [a, r, b, z] : null;
}
export function selectedPivot(
  pivots: readonly WorkbookPivotDefinition[],
  sheet: string,
  selection: string,
  regular = true
): WorkbookPivotDefinition | null {
  if (!regular || pivots.length > 32) return null;
  const target = bounds(selection);
  if (!target) return null;
  const matches = pivots.filter((p) => {
    const box = p.sheet === sheet && p.outputRange ? bounds(p.outputRange) : null;
    return box && box[0] <= target[0] && box[1] <= target[1] && box[2] >= target[2] && box[3] >= target[3];
  });
  return matches.length === 1 ? matches[0] : null;
}
export interface PivotContextFrame {
  key: string;
  pivot: WorkbookPivotDefinition;
  active: boolean;
  normal: boolean;
  readOnly: boolean;
  busy: boolean;
  accept?: (commands: WorkbookCommand[]) => Promise<unknown>;
}
export function pivotRefreshReason(frame: PivotContextFrame | null): string | undefined {
  if (!frame) return 'Select a known pivot output first.';
  if (!frame.active) return 'Activate this workbook to refresh the pivot.';
  if (!frame.normal) return 'Return to Normal view to refresh the pivot.';
  if (frame.readOnly) return 'This workbook is read-only.';
  if (frame.busy) return 'Wait for the workbook operation to finish.';
  if (!frame.accept) return 'Pivot refresh is unavailable in this editor.';
  const p = frame.pivot;
  if (!p.canEdit || !p.revision || !p.sheet || !p.sourceSheet || !p.sourceRange)
    return p.reason ?? 'This pivot is preserved and protected.';
  return undefined;
}
export function pivotContextKey(owner: readonly unknown[], pivot: WorkbookPivotDefinition): string {
  return JSON.stringify([
    ...owner,
    pivot.name,
    pivot.sheet,
    pivot.revision,
    pivot.outputRange,
    pivot.sourceSheet,
    pivot.sourceRange,
    pivot.rowField,
    pivot.valueField,
    pivot.aggregate,
    pivot.fieldNames ?? [],
    pivot.sample ?? null,
    pivot.grandTotals ?? null,
    pivot.canEdit,
    pivot.reason ?? null,
  ]);
}
export async function refreshPivotContext(
  captured: PivotContextFrame,
  current: () => PivotContextFrame | null
): Promise<{ accepted: true; key: string; current: boolean }> {
  const live = current();
  if (live?.key !== captured.key) throw Error('The selected pivot or workbook changed. Select the pivot again.');
  const reason = pivotRefreshReason(live) ?? pivotRefreshReason(captured);
  if (reason) throw Error(reason);
  const p = captured.pivot;
  await captured.accept!([
    { operation: 'refreshPivot', sheet: p.sheet!, name: p.name, expectedPivotRevision: p.revision! },
  ]);
  return { accepted: true, key: captured.key, current: current()?.key === captured.key };
}

export interface PivotFieldDraft {
  sourceSheet: string;
  range: string;
  rowField: number;
  valueField: number;
  aggregate: 'sum' | 'count';
}
export function pivotFieldDraftReason(draft: PivotFieldDraft): string | undefined {
  const box = bounds(draft.range);
  if (!draft.sourceSheet || !box) return 'Choose a source worksheet and a valid rectangle.';
  const columns = box[2] - box[0] + 1,
    rows = box[3] - box[1] + 1;
  if (columns < 2 || columns > 8 || rows < 2 || rows > 1001 || columns * rows > 8000)
    return 'Use 2–8 source fields, one header and at most 8,000 positions.';
  if (
    !Number.isInteger(draft.rowField) ||
    !Number.isInteger(draft.valueField) ||
    draft.rowField < 0 ||
    draft.valueField < 0 ||
    draft.rowField >= columns ||
    draft.valueField >= columns ||
    draft.rowField === draft.valueField
  )
    return 'Choose distinct row and value fields within the source rectangle.';
  if (!['sum', 'count'].includes(draft.aggregate)) return 'Choose SUM or numeric COUNT.';
  return undefined;
}
export async function updatePivotContext(
  captured: PivotContextFrame,
  current: () => PivotContextFrame | null,
  draft: PivotFieldDraft
): Promise<{ accepted: true; key: string; current: boolean }> {
  const fields = structuredClone(draft),
    live = current();
  if (live?.key !== captured.key) throw Error('The selected pivot or workbook changed. Select the pivot again.');
  const reason = pivotRefreshReason(live) ?? pivotRefreshReason(captured) ?? pivotFieldDraftReason(fields);
  if (reason) throw Error(reason);
  if (Object.keys(fields).sort().join(',') !== 'aggregate,range,rowField,sourceSheet,valueField')
    throw Error('Invalid pivot field draft');
  const p = captured.pivot;
  await captured.accept!([
    { operation: 'updatePivot', sheet: p.sheet!, name: p.name, expectedPivotRevision: p.revision!, ...fields },
  ]);
  return { accepted: true, key: captured.key, current: current()?.key === captured.key };
}
