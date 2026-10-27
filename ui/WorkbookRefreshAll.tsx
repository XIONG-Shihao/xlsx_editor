import { useEffect, useLayoutEffect, useMemo, useRef, useSyncExternalStore } from 'react';
import { createQueryRefreshAllController, type QueryRefreshAllFrame } from './queryRefreshAllController';
const labels: Record<string, string> = {
  ready: 'Ready',
  refused: 'Refused',
  refreshed: 'Refreshed',
  cancelled: 'Cancelled',
  'not-run': 'Not run',
  uncertain: 'Acceptance uncertain',
};
export function WorkbookRefreshAll({
  frame,
  intent,
  words = (s) => s,
}: {
  frame: QueryRefreshAllFrame | null;
  intent: number;
  words?: (s: string) => string;
}): JSX.Element {
  const controller = useMemo(() => createQueryRefreshAllController(), []),
    seen = useRef(-1),
    retirement = useRef<ReturnType<typeof setTimeout> | null>(null);
  const state = useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot);
  useLayoutEffect(() => controller.setFrame(frame));
  useEffect(() => {
    if (retirement.current !== null) clearTimeout(retirement.current);
    return () => {
      controller.setFrame(null);
      retirement.current = setTimeout(() => controller.dispose(), 0);
    };
  }, [controller]);
  useEffect(() => {
    if (!frame?.active || intent <= seen.current) return;
    seen.current = intent;
    void controller.preview();
  }, [controller, intent, frame?.active]);
  const current = !!frame?.active && state.ownerKey === frame.ownerKey,
    preview = current && (state.pending || state.key === frame?.key) ? state.preview : null,
    result = current ? state.result : null;
  return (
    <section
      aria-label={words('Refresh All local queries')}
      hidden={!frame?.active}
      className="workbook-refresh-all"
      data-formula-draft-zone="true"
    >
      <p>
        {words(
          'Preview every saved local CSV/TSV connection in this workbook, then refresh the ready connections. Each accepted connection has one Undo step. Cancel lets a current acceptance finish and stops remaining connections.'
        )}
      </p>
      <p>
        {words('Captured workbook')}: {frame?.name}
      </p>
      <div>
        <button
          type="button"
          disabled={!frame?.active || !!frame.reason || frame.busy || state.pending}
          onClick={() => void controller.preview()}
        >
          {words('Preview Refresh All')}
        </button>{' '}
        <button
          type="button"
          disabled={
            !preview ||
            state.phase !== 'preview' ||
            state.pending ||
            !!frame?.reason ||
            frame?.busy ||
            !preview.entries.some((e) => e.status === 'ready')
          }
          onClick={() => void controller.run()}
        >
          {words('Refresh ready connections')}
        </button>{' '}
        <button
          type="button"
          disabled={state.cancelPending || (!state.pending && !preview)}
          onClick={() => void controller.cancel()}
        >
          {words(state.pending ? 'Cancel remaining refreshes' : 'Cancel Refresh All preview')}
        </button>
      </div>
      {frame?.reason && <p role="status">{words(frame.reason)}</p>}
      {current && state.message && (
        <p role="status" aria-live="polite">
          {words(state.message)}
        </p>
      )}
      {preview && (
        <div aria-label={words('All connection previews')}>
          <p>
            {preview.entries.length} {words('saved local connections')} · {words(preview.undoPolicy)}
          </p>
          {preview.warning && <p role="alert">{words(preview.warning)}</p>}
          {preview.entries.length === 0 && (
            <p>{words('No saved local connections. Use Queries and refresh to add a CSV/TSV source.')}</p>
          )}
          {preview.entries.map((e) => (
            <article key={e.connectionId} data-refresh-connection={e.connectionId}>
              <h3>
                {e.name} · {words(labels[e.status])}
              </h3>
              <p>
                {e.source?.name ?? words('Source unavailable')} → {e.target.sheet}!{e.range ?? e.target.start}
                {e.outputRows !== undefined ? ' · ' + e.outputRows + ' × ' + e.outputColumns : ''}
              </p>
              {e.reason && <p role="status">{words(e.reason)}</p>}
              {e.conversions && (
                <p>
                  {words('Saved output types')}: {[...new Set(e.conversions.policies)].join(', ')} ·{' '}
                  {e.conversions.rejected} {words('rejected fields')} · {e.conversions.changed}{' '}
                  {words('conversion notices')}
                </p>
              )}
              {e.sample && (
                <details>
                  <summary>{words('Source sample (up to 3 rows and 4 fields)')}</summary>
                  <p>{words('The output uses the saved field types shown above.')}</p>
                  <table>
                    <tbody>
                      {e.sample.map((row, i) => (
                        <tr key={i}>
                          {row.map((value, j) => (
                            <td key={j}>
                              <code>{JSON.stringify(value)}</code>
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {e.sampleTruncated && <p>{words('Bounded sample; all source fields are inspected.')}</p>}
                </details>
              )}
            </article>
          ))}
        </div>
      )}
      {result && (
        <div aria-label={words('Refresh All outcomes')}>
          <p>
            {result.accepted} {words('accepted connections')} · {result.uncertain} {words('uncertain connections')} ·{' '}
            {result.undoUnits === null
              ? words('Undo count is uncertain')
              : result.undoUnits + ' ' + words('Undo steps')}
          </p>
          {result.outcomes.map((e) => (
            <article key={e.connectionId}>
              <h3>
                {e.name} · {words(labels[e.status])}
              </h3>
              <p>
                {e.target.sheet}!{e.target.start} ·{' '}
                {e.undoUnits === null ? words('Acceptance may be retained') : e.undoUnits + ' ' + words('Undo steps')}
              </p>
              {e.outputRows !== undefined && (
                <p>
                  {e.outputRows} × {e.outputColumns}
                </p>
              )}
              {e.reason && <p>{words(e.reason)}</p>}
              {e.warning && <p role="alert">{words(e.warning)}</p>}
            </article>
          ))}
        </div>
      )}
    </section>
  );
}
