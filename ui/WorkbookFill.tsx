import { useLayoutEffect, useRef, useState } from 'react';
import {
  fillSeriesPreviewPage,
  fillSeriesPreviewValue,
  type FillSeriesKind,
  type FillSeriesPreview,
} from './fillSeriesPreview';
import { captureHomeBinding, homeBindingMatches, type HomeBinding } from './homeCommands';

export interface FillSeriesContext {
  readonly key: string;
  readonly binding: HomeBinding;
  readonly caption: string;
  isCurrent(): boolean;
  isRuntimeCurrent(): boolean;
  disabledReason(action: 'preview' | 'apply'): string | undefined;
  preview(signal: AbortSignal, kind?: FillSeriesKind): Promise<FillSeriesPreview>;
  apply(preview: FillSeriesPreview): Promise<void>;
}
export interface WorkbookFillProps {
  context: FillSeriesContext | null;
  visible: boolean;
  words?: (text: string) => string;
}
const identity = (text: string) => text;

/** Preview only describes the captured range. The original owner callback owns
 * the before-acceptance re-read and native fillSeries command admission. */
export function WorkbookFill({ context, visible, words = identity }: WorkbookFillProps): JSX.Element {
  const key = context?.key ?? '',
    live = useRef({ key, visible });
  live.current = { key, visible };
  const mounted = useRef(true),
    generation = useRef(0),
    physical = useRef(false),
    abort = useRef<AbortController | null>(null);
  const [phase, setPhase] = useState<'idle' | 'reading' | 'accepting' | 'stopping'>('idle'),
    phaseRef = useRef(phase);
  phaseRef.current = phase;
  const [kind, setKind] = useState<FillSeriesKind>('number'),
    kindLive = useRef(kind);
  kindLive.current = kind;
  const [preview, setPreview] = useState<{ context: FillSeriesContext; value: FillSeriesPreview } | null>(null),
    [offset, setOffset] = useState(0),
    [error, setError] = useState(''),
    [accepted, setAccepted] = useState(false);
  const previewLive = useRef(preview);
  previewLive.current = preview;
  const current = () => mounted.current && live.current.key === key && live.current.visible;
  const fresh = !!context && context.isCurrent(),
    runtime = !!context && context.isRuntimeCurrent();
  const busy = phase !== 'idle',
    previewReason = context?.disabledReason('preview'),
    applyReason = context?.disabledReason('apply');
  const shown =
    visible && fresh && preview && preview.context.key === key && preview.context.isCurrent() && !accepted
      ? preview
      : null;
  const page = shown ? fillSeriesPreviewPage(shown.value, offset, 100) : null;

  useLayoutEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
      generation.current++;
      abort.current?.abort();
    };
  }, []);
  useLayoutEffect(() => {
    generation.current++;
    abort.current?.abort();
    setKind('number');
    setPreview(null);
    setOffset(0);
    setError('');
    setAccepted(false);
    setPhase(physical.current ? 'stopping' : 'idle');
  }, [key]);
  useLayoutEffect(() => {
    if (visible && fresh) return;
    // A normal acknowledged command advances the head before its Promise may
    // settle. Keep that original acceptance truthful under the same runtime.
    if (phaseRef.current === 'accepting' && visible && runtime) return;
    generation.current++;
    abort.current?.abort();
    setPreview(null);
    setOffset(0);
    setError('');
    setPhase(physical.current ? 'stopping' : 'idle');
  }, [visible, fresh, runtime]);

  const inspect = async (): Promise<void> => {
    if (
      !context ||
      !current() ||
      !context.isCurrent() ||
      physical.current ||
      accepted ||
      context.disabledReason('preview')
    )
      return;
    const captured = Object.freeze({ ...context, binding: captureHomeBinding(context.binding) }),
      serial = ++generation.current,
      controller = new AbortController();
    abort.current = controller;
    physical.current = true;
    setPhase('reading');
    setError('');
    setPreview(null);
    setOffset(0);
    try {
      const value = structuredClone(await captured.preview(controller.signal, kind));
      if (!current() || generation.current !== serial || controller.signal.aborted || !captured.isCurrent()) return;
      if (
        (value.date ? 'date' : 'number') !== kindLive.current ||
        !homeBindingMatches(value.binding, captured.binding) ||
        value.sheet !== captured.binding.sheet ||
        value.range !== captured.binding.selection
      )
        throw Error('The series preview does not match the captured workbook and range.');
      // The browser-safe model validates the complete numeric result before
      // returning it; paging never asks a new owner or calculates a new series.
      fillSeriesPreviewPage(value, 0, 100);
      setPreview({ context: captured, value });
    } catch (failure) {
      if (current() && generation.current === serial && !controller.signal.aborted && captured.isCurrent())
        setError(failure instanceof Error ? failure.message : String(failure));
    } finally {
      physical.current = false;
      if (mounted.current) setPhase('idle');
    }
  };
  const apply = async (): Promise<void> => {
    if (
      !shown ||
      previewLive.current !== shown ||
      !context ||
      !current() ||
      !context.isCurrent() ||
      !shown.context.isCurrent() ||
      physical.current ||
      accepted ||
      context.disabledReason('apply') ||
      shown.context.disabledReason('apply')
    )
      return;
    const captured = shown.context,
      value = structuredClone(shown.value),
      serial = ++generation.current;
    physical.current = true;
    setPhase('accepting');
    setError('');
    try {
      await captured.apply(value);
      if (current() && generation.current === serial && captured.isRuntimeCurrent()) setAccepted(true);
    } catch (failure) {
      if (current() && generation.current === serial && captured.isCurrent())
        setError(failure instanceof Error ? failure.message : String(failure));
    } finally {
      physical.current = false;
      if (mounted.current) setPhase('idle');
    }
  };
  const cancel = (): void => {
    if (!current() || !context?.isCurrent() || phaseRef.current === 'accepting' || accepted) return;
    generation.current++;
    abort.current?.abort();
    setPreview(null);
    setOffset(0);
    setError('');
    setPhase(physical.current ? 'stopping' : 'idle');
  };
  const move = (next: number): void => {
    if (
      previewLive.current !== shown ||
      !current() ||
      !context?.isCurrent() ||
      !shown?.context.isCurrent() ||
      busy ||
      physical.current ||
      accepted
    )
      return;
    setOffset(next);
  };

  const changeKind = (next: string): void => {
    if (
      !current() ||
      !context?.isCurrent() ||
      busy ||
      physical.current ||
      accepted ||
      (next !== 'number' && next !== 'date')
    )
      return;
    generation.current++;
    abort.current?.abort();
    setKind(next);
    setPreview(null);
    setOffset(0);
    setError('');
  };

  return (
    <section
      className="sheet-fill"
      aria-label={words('Fill series')}
      data-formula-draft-zone
      data-fill-state={context?.binding.state}
      data-fill-selection={context?.binding.selection}
    >
      {visible && context && runtime && (
        <>
          <p>
            <strong>{context.caption}</strong>
          </p>
          <p>
            {words(
              'Continue Numbers or Dates using the first two cells. The seed values and formatting stay unchanged.'
            )}
          </p>
        </>
      )}
      {visible && !context && <p role="status">{words('Fill series is unavailable here.')}</p>}
      {visible && context && !fresh && !accepted && (
        <p role="status">{words('The workbook or selection changed. Reopen Fill series.')}</p>
      )}
      {visible && context && runtime && (fresh || accepted) && (
        <>
          <label>
            {words('Series type')}
            <select
              aria-label={words('Series type')}
              value={kind}
              disabled={busy || !fresh || accepted}
              onChange={(event) => changeKind(event.target.value)}
            >
              <option value="number">{words('Numbers')}</option>
              <option value="date">{words('Dates (day step)')}</option>
            </select>
          </label>
          {kind === 'date' && (
            <p>
              {words(
                'Use two whole-day Excel date cells stored as numbers. Dates use the workbook date system and a constant day step. Generated cells get date formatting; other formatting stays. Date text, times and month/year/business-calendar series are not supported.'
              )}
            </p>
          )}
          <div className="sheet-fill-actions">
            <button
              type="button"
              disabled={busy || !fresh || accepted || !!previewReason}
              title={previewReason ? words(previewReason) : undefined}
              onClick={() => void inspect()}
            >
              {words('Preview series')}
            </button>
            {shown && page && (
              <button
                type="button"
                disabled={busy || !fresh || accepted || !!applyReason}
                title={applyReason ? words(applyReason) : undefined}
                onClick={() => void apply()}
              >
                {words('Apply series')}
              </button>
            )}
            <button type="button" disabled={!fresh || accepted || phase === 'accepting'} onClick={cancel}>
              {words('Cancel series')}
            </button>
          </div>
          {previewReason && fresh && !accepted && <p role="status">{words(previewReason)}</p>}
          {applyReason && applyReason !== previewReason && fresh && !accepted && (
            <p role="status">{words(applyReason)}</p>
          )}
          {phase === 'reading' && fresh && <p role="status">{words('Reading the selected cells…')}</p>}
          {phase === 'stopping' && <p role="status">{words('Finishing the preview…')}</p>}
          {phase === 'accepting' && <p role="status">{words('Finishing the series…')}</p>}
          {shown && page && (
            <>
              <dl aria-label={words('Series summary')}>
                <dt>{words('First value')}</dt>
                <dd>
                  {shown.value.seeds[0].cell}: {fillSeriesPreviewValue(shown.value, shown.value.seeds[0].value)}
                </dd>
                <dt>{words('Second value')}</dt>
                <dd>
                  {shown.value.seeds[1].cell}: {fillSeriesPreviewValue(shown.value, shown.value.seeds[1].value)}
                </dd>
                <dt>{words(shown.value.date ? 'Step (days)' : 'Step')}</dt>
                <dd>{String(shown.value.step)}</dd>
                <dt>{words('Destination cells')}</dt>
                <dd>{shown.value.destinationCount}</dd>
                <dt>{words('Range')}</dt>
                <dd>
                  {shown.value.sheet}!{shown.value.range}
                </dd>
              </dl>
              <p>
                {words(
                  shown.value.date
                    ? 'Dates are displayed as YYYY-MM-DD. Protected or unsupported cells may prevent Apply.'
                    : 'Generated numbers use up to 15 significant digits. Protected or unsupported cells may prevent Apply.'
                )}
              </p>
              <div className="sheet-fill-destinations">
                <table aria-label={words('Series destinations')}>
                  <thead>
                    <tr>
                      <th>{words('Cell')}</th>
                      <th>{words('New value')}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {page.rows.map((row) => (
                      <tr key={row.cell}>
                        <td>{row.cell}</td>
                        <td>{fillSeriesPreviewValue(shown.value, row.value)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <p role="status" aria-label={words('Series cells shown')}>
                {page.offset + 1}–{page.offset + page.rows.length} {words('of')} {page.total}
              </p>
              <button
                type="button"
                disabled={busy || page.offset === 0}
                onClick={() => move(Math.max(0, page.offset - 100))}
              >
                {words('Previous series cells')}
              </button>
              <button
                type="button"
                disabled={busy || page.nextOffset === null}
                onClick={() => {
                  if (page.nextOffset !== null) move(page.nextOffset);
                }}
              >
                {words('More series cells')}
              </button>
            </>
          )}
          {accepted && runtime && (
            <p role="status">{words('Series filled. Use Undo to restore the previous cells.')}</p>
          )}
        </>
      )}
      {error && fresh && visible && <p role="alert">{words(error)}</p>}
    </section>
  );
}
