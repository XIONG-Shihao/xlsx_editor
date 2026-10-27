import { useEffect, useRef, useState } from 'react';
import type { WorkbookPivotDefinition, WorkbookPivotLayoutSample } from '../core/contracts';
import { PIVOT_LAYOUT_SAMPLES } from '../core/pivot-layout-samples.mjs';
import { pivotLayoutDraftReason, type PivotLayoutDraft } from './pivotLayout';
import { pivotSamplePalette } from './pivotStyles';
export function WorkbookPivotDesign({
  pivot,
  bindingKey,
  reason,
  running,
  visible,
  onApply,
  words = (s: string) => s,
}: {
  pivot: WorkbookPivotDefinition;
  bindingKey: string;
  reason?: string;
  running: boolean;
  visible: boolean;
  onApply: (draft: PivotLayoutDraft) => Promise<unknown>;
  words?: (s: string) => string;
}): JSX.Element {
  const initial = () => ({ sample: pivot.sample ?? '', grandTotals: pivot.grandTotals ?? true });
  const [draft, setDraft] = useState(initial),
    live = useRef({ key: bindingKey, draft, enabled: visible && !reason && !running });
  live.current = { key: bindingKey, draft, enabled: visible && !reason && !running };
  useEffect(() => setDraft(initial()), [bindingKey]);
  const invalid = pivotLayoutDraftReason(draft),
    unchanged = draft.sample === pivot.sample && draft.grandTotals === pivot.grandTotals,
    palette = pivotSamplePalette(draft.sample);
  const apply = () => {
    if (
      !visible ||
      !live.current.enabled ||
      live.current.key !== bindingKey ||
      live.current.draft !== draft ||
      invalid ||
      unchanged
    )
      return;
    void onApply({ sample: draft.sample as WorkbookPivotLayoutSample, grandTotals: draft.grandTotals });
  };
  return (
    <section
      aria-label={words('PivotTable Design')}
      data-pivot-name={pivot.name}
      data-pivot-revision={pivot.revision ?? ''}
    >
      <h3>{pivot.name}</h3>
      <p>
        {words('Output')}: {pivot.sheet}!{pivot.outputRange}
      </p>
      <p>
        {words('Applied layout')}:{' '}
        {words(PIVOT_LAYOUT_SAMPLES.find((s) => s.id === pivot.sample)?.label ?? 'Unavailable')}
      </p>
      {reason && <p role="status">{words(reason)}</p>}
      <fieldset disabled={!visible || !!reason || running}>
        <legend>{words('PivotTable Styles')}</legend>
        <label>
          {words('Layout and style')}
          <select
            aria-label={words('Pivot layout sample')}
            value={draft.sample}
            onChange={(e) => setDraft((previous) => ({ ...previous, sample: e.target.value }))}
          >
            {!pivot.sample && <option value="">{words('Unavailable')}</option>}
            {PIVOT_LAYOUT_SAMPLES.map((sample) => (
              <option key={sample.id} value={sample.id}>
                {words(sample.label)}
              </option>
            ))}
          </select>
        </label>
        <label>
          <input
            type="checkbox"
            aria-label={words('Show pivot grand total')}
            checked={draft.grandTotals}
            onChange={(e) => setDraft((previous) => ({ ...previous, grandTotals: e.target.checked }))}
          />
          {words('Grand total')}
        </label>
        {palette && (
          <table
            aria-label={words('Pivot style sample')}
            style={{ width: '100%', borderCollapse: 'collapse', margin: '12px 0' }}
          >
            <caption>{words('Style sample')}</caption>
            <thead>
              <tr style={{ background: palette.header, color: palette.headerText }}>
                <th style={{ padding: 6, textAlign: 'left' }}>{words(palette.compact ? 'Row Labels' : 'Category')}</th>
                <th style={{ padding: 6, textAlign: 'right' }}>{words('Sum of Amount')}</th>
              </tr>
            </thead>
            <tbody>
              {[
                ['A', 4],
                ['B', 6],
              ].map(([label, value], index) => (
                <tr
                  key={label}
                  style={{ background: palette.striped && index === 0 ? palette.band : '#FFFFFF', color: palette.text }}
                >
                  <td style={{ padding: 6, paddingLeft: palette.compact ? 18 : 6 }}>{label}</td>
                  <td style={{ padding: 6, textAlign: 'right' }}>{value}</td>
                </tr>
              ))}
              {draft.grandTotals && (
                <tr style={{ background: palette.header, color: palette.headerText, fontWeight: 700 }}>
                  <td style={{ padding: 6 }}>{words('Grand Total')}</td>
                  <td style={{ padding: 6, textAlign: 'right' }}>10</td>
                </tr>
              )}
            </tbody>
          </table>
        )}
        {invalid && <p role="status">{words(invalid)}</p>}
        <button type="button" disabled={!!invalid || unchanged} onClick={apply}>
          {words(running ? 'Applying pivot layout…' : 'Apply pivot layout')}
        </button>
      </fieldset>
      <p>{words('Applying a layout keeps the pivot data. Use Refresh to update values.')}</p>
    </section>
  );
}
