/** Captured presentation edits share the selected pivot authority. */
import { PIVOT_LAYOUT_SAMPLES } from '../core/pivot-layout-samples.mjs';
import type { WorkbookPivotLayoutSample } from '../core/contracts';
import { pivotRefreshReason, type PivotContextFrame } from './pivotSelection';
export interface PivotLayoutDraft {
  sample: WorkbookPivotLayoutSample;
  grandTotals: boolean;
}
export function pivotLayoutDraftReason(draft: unknown): string | undefined {
  if (
    !draft ||
    typeof draft !== 'object' ||
    Array.isArray(draft) ||
    Object.keys(draft).sort().join(',') !== 'grandTotals,sample'
  )
    return 'Choose a supported pivot layout and grand-total setting.';
  const value = draft as PivotLayoutDraft;
  if (!PIVOT_LAYOUT_SAMPLES.some((s) => s.id === value.sample) || typeof value.grandTotals !== 'boolean')
    return 'Choose a supported pivot layout and grand-total setting.';
  return undefined;
}
export function pivotLayoutReason(frame: PivotContextFrame | null): string | undefined {
  const reason = pivotRefreshReason(frame);
  if (reason)
    return reason
      .replace('refresh the pivot', 'edit the pivot layout')
      .replace('Pivot refresh', 'Pivot layout editing');
  if (!frame?.pivot.sample || typeof frame.pivot.grandTotals !== 'boolean')
    return 'This pivot has no supported layout profile.';
  return pivotLayoutDraftReason({ sample: frame.pivot.sample, grandTotals: frame.pivot.grandTotals });
}
export async function setPivotLayoutContext(
  captured: PivotContextFrame,
  current: () => PivotContextFrame | null,
  draft: PivotLayoutDraft
): Promise<{ accepted: true; key: string; current: boolean }> {
  const fields = structuredClone(draft),
    live = current();
  if (live?.key !== captured.key) throw Error('The selected pivot or workbook changed. Select the pivot again.');
  const reason = pivotLayoutReason(live) ?? pivotLayoutReason(captured) ?? pivotLayoutDraftReason(fields);
  if (reason) throw Error(reason);
  const p = captured.pivot;
  await captured.accept!([
    { operation: 'setPivotLayout', sheet: p.sheet!, name: p.name, expectedPivotRevision: p.revision!, ...fields },
  ]);
  return { accepted: true, key: captured.key, current: current()?.key === captured.key };
}
