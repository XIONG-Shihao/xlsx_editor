import { useEffect, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react';
import type { WorkbookChartStyle, WorkbookObjectCommand, WorkbookObjectPreview } from '../core/contracts';
import { CHART_STYLES } from '../core/chart-styles.mjs';
import { createChartRecommendationsController, type ChartRecommendationFrame } from './chartRecommendationsController';
import { WorkbookObjectPreviewView } from './WorkbookObjects';
export function WorkbookChartRecommendations({
  frame,
  pending,
  onAccept,
  words = (s: string) => s,
}: {
  frame: ChartRecommendationFrame;
  pending: boolean;
  onAccept: (frame: ChartRecommendationFrame, command: WorkbookObjectCommand) => Promise<void>;
  words?: (s: string) => string;
}): JSX.Element {
  const controller = useMemo(() => createChartRecommendationsController(), []),
    snapshot = useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot),
    live = useRef(frame);
  live.current = frame;
  useLayoutEffect(() => controller.setFrame(frame), [controller, frame.key, frame.active]);
  const disposal = useRef<ReturnType<typeof setTimeout> | null>(null);
  useEffect(() => {
    if (disposal.current) clearTimeout(disposal.current);
    return () => {
      controller.setFrame(null);
      disposal.current = setTimeout(() => controller.dispose(), 0);
    };
  }, [controller]);
  const [draft, setDraft] = useState({
      key: frame.key,
      proposal: '',
      style: null as WorkbookChartStyle | null,
      cell: 'E2',
    }),
    [notice, setNotice] = useState({ key: '', text: '' }),
    current = draft.key === frame.key ? draft : { key: frame.key, proposal: '', style: null, cell: 'E2' },
    fresh = snapshot.key === frame.key && frame.active,
    result = fresh ? snapshot.result : null,
    proposal = result?.proposals.find((p) => p.id === current.proposal) ?? result?.proposals[0],
    style = current.style ?? proposal?.chart.style ?? 'colorful',
    disabled =
      pending || !frame.active || !!frame.reason || !fresh || snapshot.loading || !!snapshot.error || !proposal;
  const update = (patch: Partial<typeof draft>) => {
    if (live.current.key === frame.key && frame.active) setDraft({ ...current, ...patch });
  };
  async function apply() {
    if (disabled || live.current.key !== frame.key) return;
    try {
      const chart = controller.capture(frame, proposal!.id, style);
      await onAccept(
        { ...frame, view: structuredClone(frame.view) },
        {
          operation: 'createChart',
          sheet: frame.sheet,
          objectId: crypto.randomUUID(),
          expectedObjectRevision: null,
          chart,
          anchor: { cell: current.cell },
        }
      );
    } catch (error) {
      if (live.current.key === frame.key)
        setNotice({ key: frame.key, text: String(error instanceof Error ? error.message : error) });
    }
  }
  return (
    <section aria-label={words('Recommended Charts')} style={{ display: 'grid', gap: 12, padding: 12 }}>
      <p>
        {words(
          'Recommendations use the selected header, category and numeric cells. Choose a chart and style, then insert it.'
        )}
      </p>
      <p>{frame.sheet + '!' + frame.range}</p>
      {result && (
        <div role="group" aria-label={words('Chart recommendations')} style={{ display: 'grid', gap: 10 }}>
          {result.proposals.map((p) => {
            const selected = proposal?.id === p.id,
              chart = { ...p.chart, ...(selected ? { style } : {}) },
              object = {
                id: p.id,
                name: p.label,
                kind: 'chart',
                revision: 'preview',
                anchor: null,
                managed: false,
                canEdit: false,
                canDelete: false,
                protectedReason: null,
                chart,
                data: p.data,
              } as WorkbookObjectPreview;
            return (
              <div key={p.id} data-chart-proposal={p.id}>
                <button
                  type="button"
                  aria-pressed={selected}
                  disabled={pending || !frame.active}
                  onClick={() => update({ proposal: p.id, style: null })}
                >
                  {words(p.label)}
                </button>
                <p>{words(p.reason)}</p>
                <div style={{ height: 160, pointerEvents: 'none' }}>
                  <WorkbookObjectPreviewView object={object} embedded words={words} />
                </div>
              </div>
            );
          })}
        </div>
      )}
      <form
        onSubmit={(e) => {
          e.preventDefault();
          void apply();
        }}
      >
        <fieldset disabled={disabled} style={{ display: 'grid', gap: 8 }}>
          <legend>{words('Insert recommended chart')}</legend>
          <label>
            {words('Chart style')}
            <select
              aria-label={words('Chart style')}
              value={style}
              onChange={(e) => update({ style: e.target.value as WorkbookChartStyle })}
            >
              {CHART_STYLES.map((s) => (
                <option key={s.id} value={s.id}>
                  {words(s.label)}
                </option>
              ))}
            </select>
          </label>
          <label>
            {words('Anchor cell')}
            <input
              aria-label={words('Recommended chart anchor')}
              required
              value={current.cell}
              onChange={(e) => update({ cell: e.target.value })}
            />
          </label>
          <button type="submit">{words('Insert recommended chart')}</button>
        </fieldset>
      </form>
      {snapshot.loading && fresh && <p role="status">{words('Reading selected chart data…')}</p>}
      {frame.reason && <p role="status">{words(frame.reason)}</p>}
      {((fresh && snapshot.error) || (notice.key === frame.key && notice.text)) && (
        <p role="alert">{(fresh && snapshot.error) || notice.text}</p>
      )}
      <button
        type="button"
        disabled={pending || !frame.active || snapshot.loading}
        onClick={() => {
          setNotice({ key: frame.key, text: '' });
          controller.refresh();
        }}
      >
        {words('Refresh recommendations')}
      </button>
    </section>
  );
}
