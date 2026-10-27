import { useLayoutEffect, useRef, useState } from 'react';
import { captureHomeBinding, homeBindingMatches, type HomeBinding, type HomeFillAcrossOptions } from './homeCommands';
import { FILL_ACROSS_LIMITS, type FillAcrossPreview, type FillAcrossSheet } from './fillAcrossPreview';

export interface FillAcrossContext {
  readonly key: string;
  readonly binding: HomeBinding;
  readonly caption: string;
  readonly sheets: readonly FillAcrossSheet[];
  isCurrent(): boolean;
  isRuntimeCurrent(): boolean;
  disabledReason(): string | undefined;
  preview(options: HomeFillAcrossOptions, signal: AbortSignal): Promise<FillAcrossPreview>;
  apply(preview: FillAcrossPreview): Promise<void>;
}
export interface WorkbookFillAcrossProps {
  context: FillAcrossContext | null;
  visible: boolean;
  words?: (text: string) => string;
}
const identity = (value: string) => value;
const modes = [
  ['all', 'All'],
  ['contents', 'Contents'],
  ['formats', 'Formats'],
] as const;

export function WorkbookFillAcross({ context, visible, words = identity }: WorkbookFillAcrossProps): JSX.Element {
  const key = context?.key ?? '',
    live = useRef({ key, visible });
  live.current = { key, visible };
  const mounted = useRef(true),
    generation = useRef(0),
    physical = useRef(false),
    abort = useRef<AbortController | null>(null);
  const [options, setOptions] = useState<HomeFillAcrossOptions>({ targets: [], mode: 'all' }),
    optionsLive = useRef(options);
  optionsLive.current = options;
  const [phase, setPhase] = useState<'idle' | 'previewing' | 'accepting' | 'stopping'>('idle'),
    phaseRef = useRef(phase);
  phaseRef.current = phase;
  const [preview, setPreview] = useState<{ context: FillAcrossContext; value: FillAcrossPreview } | null>(null),
    [offset, setOffset] = useState(0),
    [error, setError] = useState(''),
    [accepted, setAccepted] = useState(false);
  const current = () => mounted.current && live.current.key === key && live.current.visible;
  const fresh = !!context && context.isCurrent(),
    runtime = !!context && context.isRuntimeCurrent(),
    busy = phase !== 'idle',
    reason = context?.disabledReason();
  const shown =
    visible && fresh && preview && preview.context.key === key && preview.context.isCurrent() && !accepted
      ? preview
      : null;
  const shownLive = useRef(shown);
  shownLive.current = shown;
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
    setOptions({ targets: [], mode: 'all' });
    setPreview(null);
    setOffset(0);
    setError('');
    setAccepted(false);
    setPhase(physical.current ? 'stopping' : 'idle');
  }, [key]);
  useLayoutEffect(() => {
    if ((visible && fresh) || (phaseRef.current === 'accepting' && visible && runtime)) return;
    generation.current++;
    abort.current?.abort();
    setPreview(null);
    setOffset(0);
    setError('');
    setPhase(physical.current ? 'stopping' : 'idle');
  }, [visible, fresh, runtime]);
  const change = (next: HomeFillAcrossOptions): void => {
    if (!current() || !context?.isCurrent() || physical.current || accepted) return;
    generation.current++;
    abort.current?.abort();
    shownLive.current = null;
    setOptions(next);
    setPreview(null);
    setOffset(0);
    setError('');
  };
  const inspect = async (): Promise<void> => {
    if (
      !context ||
      !current() ||
      !context.isCurrent() ||
      physical.current ||
      accepted ||
      reason ||
      !optionsLive.current.targets.length
    )
      return;
    const captured = Object.freeze({ ...context, binding: captureHomeBinding(context.binding) }),
      options = structuredClone(optionsLive.current),
      serial = ++generation.current,
      controller = new AbortController();
    abort.current = controller;
    physical.current = true;
    setPhase('previewing');
    setPreview(null);
    setOffset(0);
    setError('');
    try {
      const value = structuredClone(await captured.preview(options, controller.signal));
      if (!current() || generation.current !== serial || controller.signal.aborted || !captured.isCurrent()) return;
      if (
        !homeBindingMatches(value.binding, captured.binding) ||
        value.sourceSheet !== captured.binding.sheet ||
        value.range !== captured.binding.selection ||
        value.mode !== options.mode ||
        value.targets.length !== options.targets.length ||
        value.targets.some((target) => !options.targets.includes(target))
      )
        throw Error('The Across Worksheets preview belongs to another workbook, selection or destination.');
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
      shownLive.current !== shown ||
      !current() ||
      !context?.isCurrent() ||
      !shown.context.isCurrent() ||
      physical.current ||
      accepted ||
      reason ||
      shown.context.disabledReason()
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
    shownLive.current = null;
    setPreview(null);
    setOffset(0);
    setError('');
    setPhase(physical.current ? 'stopping' : 'idle');
  };
  const page = shown?.value.targets.slice(offset, offset + FILL_ACROSS_LIMITS.page) ?? [];
  return (
    <section className="sheet-fill-across" aria-label={words('Across Worksheets')} data-formula-draft-zone>
      {visible && context && runtime && (
        <>
          <p>
            <strong>{context.caption}</strong>
          </p>
          <p>
            {words(
              'Copy this range to the same cell addresses on the selected worksheets. Each Apply has one Undo step.'
            )}
          </p>
        </>
      )}
      {visible && !context && <p role="status">{words('Across Worksheets is unavailable here.')}</p>}
      {visible && context && !fresh && !accepted && (
        <p role="status">{words('The workbook or selection changed. Reopen Across Worksheets.')}</p>
      )}
      {visible && context && runtime && (fresh || accepted) && (
        <>
          <fieldset disabled={busy || !fresh || accepted}>
            <legend>{words('Copy mode')}</legend>
            {modes.map(([mode, label]) => (
              <label key={mode}>
                <input
                  type="radio"
                  name={'fill-across-mode-' + key}
                  aria-label={words(label)}
                  checked={options.mode === mode}
                  onChange={() => change({ ...optionsLive.current, mode })}
                />
                {words(label)}
              </label>
            ))}
          </fieldset>
          <fieldset disabled={busy || !fresh || accepted}>
            <legend>{words('Destination worksheets')}</legend>
            <div className="sheet-fill-across-picker">
              {context.sheets
                .filter((sheet) => sheet.name !== context.binding.sheet && sheet.state === 'visible')
                .map((sheet) => (
                  <label key={sheet.name}>
                    <input
                      type="checkbox"
                      aria-label={words('Destination worksheet') + ' ' + sheet.name}
                      checked={options.targets.includes(sheet.name)}
                      onChange={(event) =>
                        change({
                          ...optionsLive.current,
                          targets: event.target.checked
                            ? [...optionsLive.current.targets, sheet.name]
                            : optionsLive.current.targets.filter((target) => target !== sheet.name),
                        })
                      }
                    />
                    {sheet.name}
                  </label>
                ))}
            </div>
          </fieldset>
          <p>
            {words(
              options.mode === 'all'
                ? 'All copies cell contents and formatting.'
                : options.mode === 'contents'
                  ? 'Contents copies values and formulas while preserving destination formatting.'
                  : 'Formats copies cell formatting while preserving destination contents.'
            )}
          </p>
          <div className="sheet-fill-across-actions">
            <button
              type="button"
              disabled={busy || !fresh || accepted || !!reason || !options.targets.length}
              title={reason ? words(reason) : undefined}
              onClick={() => void inspect()}
            >
              {words('Preview destinations')}
            </button>
            {shown && (
              <button
                type="button"
                disabled={busy || !fresh || accepted || !!reason}
                title={reason ? words(reason) : undefined}
                onClick={() => void apply()}
              >
                {words('Apply Across Worksheets')}
              </button>
            )}
            <button type="button" disabled={!fresh || accepted || phase === 'accepting'} onClick={cancel}>
              {words('Cancel Across Worksheets')}
            </button>
          </div>
          {reason && fresh && !accepted && <p role="status">{words(reason)}</p>}
          {phase === 'previewing' && fresh && <p role="status">{words('Preparing the destination preview…')}</p>}
          {phase === 'accepting' && <p role="status">{words('Filling the destination worksheets…')}</p>}
          {phase === 'stopping' && <p role="status">{words('Finishing the destination preview…')}</p>}
          {shown && (
            <>
              <dl aria-label={words('Across Worksheets summary')}>
                <dt>{words('Source range')}</dt>
                <dd>
                  {shown.value.sourceSheet}!{shown.value.range}
                </dd>
                <dt>{words('Copy mode')}</dt>
                <dd>{words(modes.find(([mode]) => mode === shown.value.mode)![1])}</dd>
                <dt>{words('Destination worksheets')}</dt>
                <dd>{shown.value.targets.length}</dd>
                <dt>{words('Destination cells')}</dt>
                <dd>{shown.value.destinationPositions}</dd>
              </dl>
              <div className="sheet-fill-across-destinations">
                <table aria-label={words('Across Worksheets destinations')}>
                  <thead>
                    <tr>
                      <th>{words('Worksheet')}</th>
                      <th>{words('Range')}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {page.map((target) => (
                      <tr key={target}>
                        <td>{target}</td>
                        <td>{shown.value.range}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <p role="status">
                {offset + 1}–{offset + page.length} {words('of')} {shown.value.targets.length}
              </p>
              <button
                type="button"
                disabled={busy || offset === 0}
                onClick={() => {
                  if (current() && context.isCurrent()) setOffset(Math.max(0, offset - FILL_ACROSS_LIMITS.page));
                }}
              >
                {words('Previous destinations')}
              </button>
              <button
                type="button"
                disabled={busy || offset + page.length === shown.value.targets.length}
                onClick={() => {
                  if (current() && context.isCurrent()) setOffset(offset + FILL_ACROSS_LIMITS.page);
                }}
              >
                {words('More destinations')}
              </button>
            </>
          )}
          {accepted && runtime && (
            <p role="status">{words('Destination worksheets filled. Use Undo to restore their previous cells.')}</p>
          )}
        </>
      )}
      {error && visible && fresh && <p role="alert">{words(error)}</p>}
    </section>
  );
}
