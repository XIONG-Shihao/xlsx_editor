import { useLayoutEffect, useRef, useState } from 'react';
import { captureHomeBinding, homeBindingMatches, type HomeBinding } from './homeCommands';
import {
  DEFAULT_SPECIAL_SELECTION_OPTIONS,
  SPECIAL_SELECTION_KINDS,
  normalizeSpecialSelectionOptions,
  specialSelectionPreviewPage,
  type SpecialSelectionOptions,
  type SpecialSelectionPreview,
  type SpecialSelectionTypes,
} from './specialSelection';

export interface SpecialSelectionContext {
  readonly key: string;
  readonly binding: HomeBinding;
  readonly caption: string;
  readonly initialOptions?: SpecialSelectionOptions;
  isCurrent(): boolean;
  /** A view-selection acknowledgement can retire the old selection binding.
   * This callback identifies only the same live workbook runtime. */
  isRuntimeCurrent(): boolean;
  disabledReason(action: 'preview' | 'apply'): string | undefined;
  preview(options: SpecialSelectionOptions, signal: AbortSignal): Promise<SpecialSelectionPreview>;
  apply(preview: SpecialSelectionPreview): Promise<void>;
}
export interface WorkbookSpecialSelectionProps {
  context: SpecialSelectionContext | null;
  visible: boolean;
  words?: (text: string) => string;
}
const identity = (text: string) => text;
const typeLabels: ReadonlyArray<[keyof SpecialSelectionTypes, string]> = [
  ['number', 'Numbers'],
  ['text', 'Text'],
  ['logical', 'Logicals'],
  ['error', 'Errors'],
];

/** This pane selects captured source references only. The host re-reads the
 * original rectangle before changing its local selection; no workbook Apply. */
export function WorkbookSpecialSelection({
  context,
  visible,
  words = identity,
}: WorkbookSpecialSelectionProps): JSX.Element {
  const key = context?.key ?? '',
    live = useRef({ key, visible });
  live.current = { key, visible };
  const mounted = useRef(true),
    generation = useRef(0),
    physical = useRef(false),
    abort = useRef<AbortController | null>(null);
  const [options, setOptions] = useState<SpecialSelectionOptions>(() =>
      normalizeSpecialSelectionOptions(context?.initialOptions ?? DEFAULT_SPECIAL_SELECTION_OPTIONS)
    ),
    optionsLive = useRef(options);
  optionsLive.current = options;
  const [phase, setPhase] = useState<'idle' | 'reading' | 'selecting' | 'stopping'>('idle'),
    phaseRef = useRef(phase);
  phaseRef.current = phase;
  const [preview, setPreview] = useState<{ context: SpecialSelectionContext; value: SpecialSelectionPreview } | null>(
      null
    ),
    [offset, setOffset] = useState(0),
    [error, setError] = useState(''),
    [accepted, setAccepted] = useState(false);
  const current = () => mounted.current && live.current.key === key && live.current.visible;
  const fresh = !!context && context.isCurrent(),
    runtime = !!context && context.isRuntimeCurrent(),
    busy = phase !== 'idle';
  const shown =
    visible && fresh && preview && preview.context.key === key && preview.context.isCurrent() && !accepted
      ? preview
      : null;
  const page = shown ? specialSelectionPreviewPage(shown.value, offset, 100) : null;
  const previewReason = context?.disabledReason('preview'),
    applyReason = context?.disabledReason('apply');
  const typed = options.kind === 'constants' || options.kind === 'formulas',
    typesSelected = Object.values(options.types).some(Boolean);
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
    setPreview(null);
    setOffset(0);
    setError('');
    setAccepted(false);
    setOptions(normalizeSpecialSelectionOptions(context?.initialOptions ?? DEFAULT_SPECIAL_SELECTION_OPTIONS));
    setPhase(physical.current ? 'stopping' : 'idle');
  }, [key]);
  useLayoutEffect(() => {
    if ((visible && fresh) || (phaseRef.current === 'selecting' && visible && runtime)) return;
    generation.current++;
    abort.current?.abort();
    setPreview(null);
    setOffset(0);
    setError('');
    setPhase(physical.current ? 'stopping' : 'idle');
  }, [visible, fresh, runtime]);
  const change = (next: SpecialSelectionOptions): void => {
    if (!current() || !context?.isCurrent() || physical.current || accepted) return;
    generation.current++;
    abort.current?.abort();
    setPreview(null);
    setOffset(0);
    setError('');
    setOptions(next);
  };
  const inspect = async (): Promise<void> => {
    if (
      !context ||
      !current() ||
      !context.isCurrent() ||
      physical.current ||
      accepted ||
      context.disabledReason('preview') ||
      (typed && !typesSelected)
    )
      return;
    const captured = Object.freeze({ ...context, binding: captureHomeBinding(context.binding) }),
      capturedOptions = normalizeSpecialSelectionOptions(optionsLive.current),
      serial = ++generation.current,
      controller = new AbortController();
    abort.current = controller;
    physical.current = true;
    setPhase('reading');
    setPreview(null);
    setOffset(0);
    setError('');
    try {
      const value = structuredClone(await captured.preview(capturedOptions, controller.signal));
      if (!current() || generation.current !== serial || controller.signal.aborted || !captured.isCurrent()) return;
      if (
        !homeBindingMatches(value.binding, captured.binding) ||
        JSON.stringify(value.options) !== JSON.stringify(capturedOptions)
      )
        throw Error('The Go To Special preview belongs to another workbook, range or cell type.');
      specialSelectionPreviewPage(value, 0, 100);
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
      !context ||
      !current() ||
      !context.isCurrent() ||
      !shown.context.isCurrent() ||
      physical.current ||
      accepted ||
      !shown.value.references.length ||
      context.disabledReason('apply') ||
      shown.context.disabledReason('apply')
    )
      return;
    const captured = shown.context,
      value = structuredClone(shown.value),
      serial = ++generation.current;
    physical.current = true;
    setPhase('selecting');
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
    if (!current() || !context?.isCurrent() || phaseRef.current === 'selecting' || accepted) return;
    generation.current++;
    abort.current?.abort();
    setPreview(null);
    setOffset(0);
    setError('');
    setPhase(physical.current ? 'stopping' : 'idle');
  };
  const move = (next: number): void => {
    if (current() && context?.isCurrent() && shown?.context.isCurrent() && !busy && !physical.current && !accepted)
      setOffset(next);
  };
  return (
    <section className="sheet-special-selection" aria-label={words('Go To Special')} data-formula-draft-zone>
      {visible && context && runtime && (
        <>
          <p>
            <strong>{context.caption}</strong>
          </p>
          <p>
            {words(
              'Inspect matching cells, then select them together. Constants, Formulas and Blanks use the worksheet used range when one cell or one merged area is selected. Last cell uses the whole worksheet.'
            )}
          </p>
        </>
      )}
      {visible && !context && <p role="status">{words('Go To Special is unavailable here.')}</p>}
      {visible && context && !fresh && !accepted && (
        <p role="status">{words('The workbook or selection changed. Reopen Go To Special.')}</p>
      )}
      {visible && context && runtime && (fresh || accepted) && (
        <>
          <label>
            {words('Selection type')}
            <select
              aria-label={words('Selection type')}
              value={options.kind}
              disabled={busy || !fresh || accepted}
              onChange={(event) => {
                const kind = SPECIAL_SELECTION_KINDS.find((k) => k.id === event.target.value);
                if (kind?.supported) change({ ...optionsLive.current, kind: kind.id });
              }}
            >
              {SPECIAL_SELECTION_KINDS.map((kind) => (
                <option key={kind.id} value={kind.id} disabled={!kind.supported}>
                  {words(kind.label)}
                  {!kind.supported ? ' — ' + words('Not available yet') : ''}
                </option>
              ))}
            </select>
          </label>
          {options.kind === 'current-region' && (
            <p>
              {words(
                'Selects a regular data block around the active populated cell. Up to 10,000 positions; merged, grouped and diagonal-only blocks are outside this profile.'
              )}
            </p>
          )}
          {typed && (
            <fieldset>
              <legend>{words('Cell types')}</legend>
              {typeLabels.map(([type, label]) => (
                <label key={type}>
                  <input
                    type="checkbox"
                    aria-label={words(label)}
                    checked={options.types[type]}
                    disabled={busy || !fresh || accepted}
                    onChange={(event) =>
                      change({
                        ...optionsLive.current,
                        types: { ...optionsLive.current.types, [type]: event.target.checked },
                      })
                    }
                  />
                  {words(label)}
                </label>
              ))}
            </fieldset>
          )}
          {typed && !typesSelected && <p role="status">{words('Choose at least one cell type.')}</p>}
          <div className="sheet-special-selection-actions">
            <button
              type="button"
              disabled={busy || !fresh || accepted || !!previewReason || (typed && !typesSelected)}
              title={previewReason ? words(previewReason) : undefined}
              onClick={() => void inspect()}
            >
              {words('Preview matching cells')}
            </button>
            {shown && page && (
              <button
                type="button"
                disabled={busy || !fresh || accepted || !shown.value.references.length || !!applyReason}
                title={applyReason ? words(applyReason) : undefined}
                onClick={() => void apply()}
              >
                {words('Select matching cells')}
              </button>
            )}
            <button type="button" disabled={!fresh || accepted || phase === 'selecting'} onClick={cancel}>
              {words('Cancel selection')}
            </button>
          </div>
          {previewReason && fresh && !accepted && <p role="status">{words(previewReason)}</p>}
          {applyReason && applyReason !== previewReason && fresh && !accepted && (
            <p role="status">{words(applyReason)}</p>
          )}
          {phase === 'reading' && fresh && <p role="status">{words('Reading the complete selection scope…')}</p>}
          {phase === 'stopping' && <p role="status">{words('Finishing the selection preview…')}</p>}
          {phase === 'selecting' && <p role="status">{words('Selecting the matching cells…')}</p>}
          {shown && page && (
            <>
              <dl aria-label={words('Special selection summary')}>
                <dt>{words('Source positions')}</dt>
                <dd>{shown.value.positions}</dd>
                <dt>{words('Matching cells')}</dt>
                <dd>{shown.value.references.length}</dd>
                <dt>{words('Range')}</dt>
                <dd>
                  {shown.value.sheet}!{shown.value.range}
                </dd>
              </dl>
              {shown.value.scope.usedRange?.status === 'complete' && (
                <p>
                  {words('Worksheet used range')}: {shown.value.sheet}!{shown.value.scope.usedRange.range}
                </p>
              )}
              {!page.total && <p role="status">{words('No cells match. The current selection stays unchanged.')}</p>}
              <div className="sheet-special-selection-rows">
                <table aria-label={words('Special selection targets')}>
                  <thead>
                    <tr>
                      <th>{words('Cell')}</th>
                      <th>{words('Type')}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {page.rows.map((row) => (
                      <tr key={row.cell}>
                        <td>
                          {shown.value.sheet}!{row.cell}
                        </td>
                        <td>{words(row.classification)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <p role="status" aria-label={words('Special selection cells shown')}>
                {page.total ? `${page.offset + 1}–${page.offset + page.rows.length}` : '0'} {words('of')} {page.total}
              </p>
              <button
                type="button"
                disabled={busy || page.offset === 0}
                onClick={() => move(Math.max(0, page.offset - 100))}
              >
                {words('Previous matching cells')}
              </button>
              <button
                type="button"
                disabled={busy || page.nextOffset === null}
                onClick={() => {
                  if (page.nextOffset !== null) move(page.nextOffset);
                }}
              >
                {words('More matching cells')}
              </button>
            </>
          )}
          {accepted && runtime && (
            <p role="status">{words('Matching cells selected. Workbook values and history stay unchanged.')}</p>
          )}
        </>
      )}
      {error && fresh && visible && <p role="alert">{words(error)}</p>}
    </section>
  );
}
