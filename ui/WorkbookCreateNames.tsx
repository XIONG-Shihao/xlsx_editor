/** One retained workbook-name creation workflow; no acceptance during preview. */
import { useLayoutEffect, useRef, useState } from 'react';
import {
  previewCreateNames,
  acceptCreateNames,
  createNamesBindingKey,
  createNameChoiceReason,
  type CreateNamesContext,
  type CreateNamesPreview,
  type NameEdges,
  type NameChoice,
} from './createNamesFromSelection';
export interface WorkbookCreateNamesProps {
  context: CreateNamesContext | null;
  visible: boolean;
  words?: (text: string) => string;
  onOutcome?: (result: Awaited<ReturnType<typeof acceptCreateNames>>) => void;
}
const identity = (text: string) => text;
const locations = [
  ['top', 'Top row'],
  ['left', 'Left column'],
  ['bottom', 'Bottom row'],
  ['right', 'Right column'],
] as const;
export function WorkbookCreateNames({
  context,
  visible,
  words = identity,
  onOutcome,
}: WorkbookCreateNamesProps): JSX.Element {
  const key = context ? createNamesBindingKey(context.binding) : '',
    live = useRef({ key, visible, context });
  live.current = { key, visible, context };
  const [edges, setEdges] = useState<NameEdges>({ top: true, left: false, bottom: false, right: false }),
    [preview, setPreview] = useState<CreateNamesPreview | null>(null),
    [choices, setChoices] = useState<Record<string, NameChoice>>({}),
    [page, setPage] = useState(0);
  const [phase, setPhase] = useState<'idle' | 'previewing' | 'checking' | 'accepting' | 'stopping'>('idle'),
    [pending, setPending] = useState(false),
    [error, setError] = useState(''),
    [outcome, setOutcome] = useState<Awaited<ReturnType<typeof acceptCreateNames>> | null>(null);
  const serial = useRef(0),
    physical = useRef(false),
    abort = useRef<AbortController | null>(null),
    phaseRef = useRef(phase),
    mounted = useRef(true);
  phaseRef.current = phase;
  const fresh = !!context && context.isCurrent(context.binding),
    runtime = !!context && context.isRuntimeCurrent(context.binding);
  const current = () => mounted.current && live.current.key === key && live.current.visible;
  useLayoutEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
      serial.current++;
      if (phaseRef.current !== 'accepting') abort.current?.abort();
    };
  }, []);
  useLayoutEffect(() => {
    serial.current++;
    if (phaseRef.current !== 'accepting') abort.current?.abort();
    setPreview(null);
    setChoices({});
    setOutcome(null);
    setError('');
    setPhase('idle');
    setPage(0);
    return () => {
      serial.current++;
      if (phaseRef.current !== 'accepting') abort.current?.abort();
    };
  }, [key]);
  useLayoutEffect(() => {
    if (!visible && phaseRef.current !== 'accepting') {
      serial.current++;
      abort.current?.abort();
      setPreview(null);
      setPhase('idle');
    }
  }, [visible]);
  const reset = (): void => {
    serial.current++;
    abort.current?.abort();
    setPreview(null);
    setChoices({});
    setOutcome(null);
    setError('');
    setPage(0);
  };
  const cancel = (): void => {
    if (phaseRef.current === 'accepting') return;
    serial.current++;
    abort.current?.abort();
    setPreview(null);
    setChoices({});
    setError('');
    setPhase(physical.current ? 'stopping' : 'idle');
  };
  const inspect = async (): Promise<void> => {
    if (!context || !fresh || !current() || physical.current) return;
    const action = ++serial.current,
      controller = new AbortController();
    physical.current = true;
    abort.current = controller;
    setPending(true);
    setPhase('previewing');
    setError('');
    setOutcome(null);
    setPreview(null);
    try {
      const value = await previewCreateNames(context, edges, { signal: controller.signal });
      if (current() && serial.current === action) {
        setPreview(value);
        setChoices(Object.fromEntries(value.proposals.map((item) => [item.id, { name: item.name, included: true }])));
        setPage(0);
      }
    } catch (failure) {
      if (current() && serial.current === action && context.isCurrent(context.binding) && !controller.signal.aborted)
        setError(failure instanceof Error ? failure.message : String(failure));
    } finally {
      physical.current = false;
      if (mounted.current) setPending(false);
      if (current() && (serial.current === action || phaseRef.current === 'stopping')) setPhase('idle');
    }
  };
  const create = async (): Promise<void> => {
    if (!context || !preview || !fresh || !current() || physical.current) return;
    const action = ++serial.current,
      controller = new AbortController(),
      observer = onOutcome;
    physical.current = true;
    abort.current = controller;
    setPending(true);
    setPhase('checking');
    setError('');
    try {
      const value = await acceptCreateNames(context, preview, structuredClone(choices), {
        signal: controller.signal,
        beforeAcceptance: () => {
          if (!current() || serial.current !== action) controller.abort();
          else {
            phaseRef.current = 'accepting';
            setPhase('accepting');
          }
        },
      });
      try {
        observer?.(value);
      } catch (failure) {
        console.warn('xlsx.create-names-outcome', failure);
      }
      if (current() && serial.current === action && value.runtimeCurrent) setOutcome(value);
    } catch (failure) {
      if (current() && serial.current === action && context.isCurrent(context.binding) && !controller.signal.aborted)
        setError(failure instanceof Error ? failure.message : String(failure));
    } finally {
      physical.current = false;
      if (mounted.current) setPending(false);
      if (current() && (serial.current === action || phaseRef.current === 'stopping')) setPhase('idle');
    }
  };
  const shown = runtime && fresh && visible ? preview : null,
    busy = pending || phase !== 'idle',
    included = shown?.proposals.filter((item) => choices[item.id]?.included ?? true) ?? [];
  const reasons = shown ? included.map((item) => createNameChoiceReason(shown, item, choices)).filter(Boolean) : [],
    access = context?.availability();
  return (
    <section
      className="sheet-create-names"
      aria-label={words('Create from Selection')}
      data-create-names-state={context?.binding.state}
      data-create-names-selection={context?.binding.selection}
    >
      <p>
        {words(
          'Create workbook names from selected text labels. Review each name and its exact range before creating all included names in one Undo step.'
        )}
      </p>
      {(!runtime || (!fresh && !outcome)) && (
        <p role="status">{words('The workbook or selection changed. Reopen Create from Selection.')}</p>
      )}
      {context && runtime && (
        <>
          <p>
            <strong>
              {context.binding.sheet}!{context.binding.selection}
            </strong>
          </p>
          <fieldset disabled={busy || !fresh || !!outcome}>
            <legend>{words('Create names from values in the')}</legend>
            {locations.map(([edge, label]) => (
              <label className="sheet-name-location" key={edge}>
                <input
                  aria-label={words(label)}
                  type="checkbox"
                  checked={edges[edge]}
                  onChange={(event) => {
                    reset();
                    setEdges((value) => ({ ...value, [edge]: event.target.checked }));
                  }}
                />
                {words(label)}
              </label>
            ))}
          </fieldset>
          <button
            type="button"
            disabled={busy || !fresh || !!outcome || access?.readOnly || access?.busy}
            onClick={() => void inspect()}
          >
            {words('Preview names')}
          </button>
          <button type="button" disabled={phase === 'accepting'} onClick={cancel}>
            {words(phase === 'accepting' ? 'Finishing…' : 'Cancel name creation')}
          </button>
          {phase === 'previewing' && <p role="status">{words('Reading the selected labels…')}</p>}
          {phase === 'checking' && <p role="status">{words('Checking labels and existing names…')}</p>}
          {phase === 'accepting' && runtime && <p role="status">{words('Finishing name creation…')}</p>}
          {phase === 'stopping' && <p role="status">{words('Stopping the preview…')}</p>}
          {shown && !outcome && (
            <>
              <p role="status">
                {included.length} {words('names included')} · {words('Data interior')}: {shown.dataRange}
              </p>
              {!!shown.corners.length && (
                <p>
                  {words('Corner cells shared by chosen label edges are excluded')}: {shown.corners.join(', ')}.
                </p>
              )}
              <p>
                {words(
                  'Leading and trailing spaces are trimmed; internal spaces become underscores. Other unsupported characters need an edited ASCII name or explicit exclusion. Existing names are never replaced.'
                )}
              </p>
              <table aria-label={words('Names preview')}>
                <thead>
                  <tr>
                    <th>{words('Include')}</th>
                    <th>{words('Label / proposed name')}</th>
                    <th>{words('Named range')}</th>
                  </tr>
                </thead>
                <tbody>
                  {shown.proposals.slice(page * 20, (page + 1) * 20).map((item) => {
                    const choice = choices[item.id] ?? { name: item.name, included: true },
                      reason = createNameChoiceReason(shown, item, choices);
                    return (
                      <tr key={item.id} data-name-label={item.labelCell}>
                        <td>
                          <input
                            type="checkbox"
                            aria-label={words('Include name from') + ' ' + item.labelCell}
                            checked={choice.included}
                            disabled={busy}
                            onChange={(event) =>
                              setChoices((values) => ({
                                ...values,
                                [item.id]: { ...choice, included: event.target.checked },
                              }))
                            }
                          />
                        </td>
                        <td>
                          <strong>{item.labelCell}</strong>{' '}
                          <output aria-label={words('Source label') + ' ' + item.labelCell}>
                            {JSON.stringify(item.label)}
                          </output>
                          <input
                            aria-label={words('Proposed name') + ' ' + item.labelCell}
                            value={choice.name}
                            maxLength={255}
                            disabled={busy || !choice.included}
                            onChange={(event) =>
                              setChoices((values) => ({
                                ...values,
                                [item.id]: { ...choice, name: event.target.value },
                              }))
                            }
                          />
                          {item.converted && <small>{words('Converted from the label; review this name.')}</small>}
                          {!choice.included ? (
                            <small>{words('Excluded')}</small>
                          ) : (
                            reason && <small role="status">{words(reason)}</small>
                          )}
                        </td>
                        <td>
                          <output aria-label={words('Named range') + ' ' + item.labelCell}>{item.formula}</output>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
              <button disabled={busy || page === 0} onClick={() => setPage((value) => value - 1)}>
                {words('Previous name proposals')}
              </button>
              <span>
                {Math.min(page * 20 + 1, shown.proposals.length)}–{Math.min((page + 1) * 20, shown.proposals.length)} /{' '}
                {shown.proposals.length}
              </span>
              <button
                disabled={busy || (page + 1) * 20 >= shown.proposals.length}
                onClick={() => setPage((value) => value + 1)}
              >
                {words('More name proposals')}
              </button>
              <button
                type="button"
                disabled={busy || !fresh || access?.readOnly || access?.busy || !included.length || !!reasons.length}
                onClick={() => void create()}
              >
                {words('Create names')}
              </button>
            </>
          )}
          {outcome && runtime && (
            <p role="status">
              {words('Created')} {outcome.count} {words('workbook names. Use Undo to remove this batch.')}
            </p>
          )}
        </>
      )}
      {error && runtime && fresh && visible && <p role="alert">{words(error)}</p>}
    </section>
  );
}
