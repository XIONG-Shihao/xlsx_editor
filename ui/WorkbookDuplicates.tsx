import { useLayoutEffect, useRef, useState } from 'react';
import {
  duplicateColumns,
  duplicateConfig,
  duplicateConfigKey,
  duplicateContextKey,
  previewDuplicates,
  readDuplicateHeaders,
  removeDuplicates,
  type DuplicateContext,
  type DuplicatePreview,
} from './duplicateSelection';
export interface WorkbookDuplicatesProps {
  context: DuplicateContext | null;
  visible: boolean;
  words?: (text: string) => string;
}
const identity = (text: string) => text;
export function WorkbookDuplicates({ context, visible, words = identity }: WorkbookDuplicatesProps): JSX.Element {
  const key = context ? duplicateContextKey(context) : '',
    live = useRef({ key, visible });
  live.current = { key, visible };
  const [hasHeaders, setHasHeaders] = useState(true),
    [columns, setColumns] = useState<string[]>([]),
    [columnPage, setColumnPage] = useState(0),
    [labels, setLabels] = useState<Record<string, string>>({}),
    [preview, setPreview] = useState<{ configKey: string; value: DuplicatePreview } | null>(null);
  const [phase, setPhase] = useState<'idle' | 'headings' | 'previewing' | 'accepting' | 'stopping'>('idle'),
    [error, setError] = useState(''),
    [outcome, setOutcome] = useState<{ removedRows: number } | null>(null),
    [settled, setSettled] = useState(0);
  const serial = useRef(0),
    physical = useRef(false),
    mounted = useRef(true),
    abort = useRef<AbortController | null>(null),
    phaseRef = useRef(phase),
    headingRequest = useRef('');
  phaseRef.current = phase;
  const fresh = !!context && context.isCurrent(),
    runtime = !!context && context.isRuntimeCurrent(),
    current = () => mounted.current && live.current.key === key && live.current.visible;
  const available = context ? duplicateColumns(context.binding.selection) : [],
    busy = phase !== 'idle';
  let config = null,
    configReason = '';
  if (context)
    try {
      config = duplicateConfig(context, hasHeaders, columns);
    } catch (failure) {
      configReason = failure instanceof Error ? failure.message : String(failure);
    }
  const configKey = config ? duplicateConfigKey(config) : '',
    liveConfig = useRef(configKey);
  liveConfig.current = configKey;
  const previewReason = context?.disabledReason('preview'),
    removeReason = context?.disabledReason('remove');
  const shown = visible && fresh && preview?.configKey === configKey ? preview.value : null;
  useLayoutEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
      serial.current++;
      abort.current?.abort();
    };
  }, []);
  useLayoutEffect(() => {
    serial.current++;
    abort.current?.abort();
    const all = context ? duplicateColumns(context.binding.selection) : [];
    setHasHeaders(context?.initialHeader ?? true);
    setColumns(all.length <= 16 ? all : []);
    setColumnPage(0);
    setLabels({});
    setPreview(null);
    setError('');
    setOutcome(null);
    setPhase(physical.current ? 'stopping' : 'idle');
  }, [key]);
  useLayoutEffect(() => {
    if (!visible) {
      serial.current++;
      abort.current?.abort();
      setPreview(null);
      if (phaseRef.current !== 'accepting') setPhase(physical.current ? 'stopping' : 'idle');
    }
  }, [visible]);
  useLayoutEffect(() => {
    const request = key + '|' + columnPage;
    if (
      !context ||
      !visible ||
      !fresh ||
      physical.current ||
      headingRequest.current === request ||
      phaseRef.current === 'accepting'
    )
      return;
    const generation = ++serial.current,
      controller = new AbortController();
    abort.current = controller;
    physical.current = true;
    headingRequest.current = request;
    setPhase('headings');
    void readDuplicateHeaders(context, columnPage, controller.signal)
      .then((value) => {
        if (current() && serial.current === generation) setLabels((prior) => ({ ...prior, ...value }));
      })
      .catch((failure) => {
        if (current() && serial.current === generation && !controller.signal.aborted && context.isCurrent())
          setError(failure instanceof Error ? failure.message : String(failure));
      })
      .finally(() => {
        physical.current = false;
        if (mounted.current) {
          setPhase('idle');
          setSettled((value) => value + 1);
        }
      });
    return () => {
      serial.current++;
      controller.abort();
    };
  }, [key, visible, columnPage, settled]);
  const invalidate = (): void => {
    serial.current++;
    abort.current?.abort();
    setPreview(null);
    setOutcome(null);
    setError('');
  };
  const editable = () => current() && fresh && liveConfig.current === configKey && !busy && !outcome;
  const inspect = async (offset = 0): Promise<void> => {
    if (
      !context ||
      !config ||
      !current() ||
      !fresh ||
      liveConfig.current !== configKey ||
      physical.current ||
      context.disabledReason('preview')
    )
      return;
    const generation = ++serial.current,
      captured = structuredClone(config),
      capturedKey = configKey,
      previous = shown;
    physical.current = true;
    setPhase('previewing');
    setError('');
    setOutcome(null);
    try {
      const result = await previewDuplicates(context, captured, offset, previous ?? undefined);
      if (current() && serial.current === generation) setPreview({ configKey: capturedKey, value: result });
    } catch (failure) {
      if (current() && serial.current === generation && context.isCurrent())
        setError(failure instanceof Error ? failure.message : String(failure));
    } finally {
      physical.current = false;
      if (mounted.current) {
        setPhase('idle');
        setSettled((value) => value + 1);
      }
    }
  };
  const remove = async (): Promise<void> => {
    if (
      !context ||
      !config ||
      !shown ||
      !current() ||
      !fresh ||
      liveConfig.current !== configKey ||
      physical.current ||
      context.disabledReason('remove') ||
      !shown.removedRows
    )
      return;
    const generation = ++serial.current,
      captured = structuredClone(config),
      receipt = structuredClone(shown);
    physical.current = true;
    setPhase('accepting');
    setError('');
    try {
      const result = await removeDuplicates(context, captured, receipt);
      if (current() && serial.current === generation && result.runtimeCurrent)
        setOutcome({ removedRows: result.removedRows });
    } catch (failure) {
      if (current() && serial.current === generation && context.isCurrent())
        setError(failure instanceof Error ? failure.message : String(failure));
    } finally {
      physical.current = false;
      if (mounted.current) {
        setPhase('idle');
        setSettled((value) => value + 1);
      }
    }
  };
  const cancel = (): void => {
    if (!current() || !fresh || liveConfig.current !== configKey || phaseRef.current === 'accepting') return;
    invalidate();
    setPhase(physical.current ? 'stopping' : 'idle');
  };
  return (
    <section
      className="sheet-duplicates"
      aria-label={words('Remove duplicates')}
      data-formula-draft-zone
      data-duplicate-state={context?.binding.state}
      data-duplicate-selection={context?.binding.selection}
    >
      {context && runtime && (
        <>
          <p>
            <strong>{context.caption}</strong>
          </p>
          <p>
            {words(
              'Keep the first matching row. Extra selected cells at the end are cleared; cells outside the selected range stay unchanged.'
            )}
          </p>
          <details>
            <summary>{words('How values are compared')}</summary>
            <p>
              {words(
                'Text matches exactly, including case and spaces. Blank cells and empty text differ. Numbers, text, booleans and errors stay separate. Numeric spellings such as 1 and 1.00 count as the same number.'
              )}
            </p>
          </details>
        </>
      )}
      {!fresh && !outcome && (
        <p role="status">{words('The workbook or selection changed. Reopen Remove duplicates.')}</p>
      )}
      {context && runtime && (fresh || outcome) && (
        <>
          <fieldset disabled={busy || !fresh || !!outcome}>
            <legend>{words('Compare these columns')}</legend>
            <label>
              <input
                type="checkbox"
                aria-label={words('My data has headers')}
                checked={hasHeaders}
                onChange={(event) => {
                  if (!editable()) return;
                  invalidate();
                  setHasHeaders(event.target.checked);
                }}
              />
              {words('My data has headers')}
            </label>
            <p>
              {words('Selected columns')}: {columns.join(', ') || words('None')} · {columns.length}/16
            </p>
            {available.slice(columnPage * 100, (columnPage + 1) * 100).map((column) => (
              <label key={column}>
                <input
                  type="checkbox"
                  aria-label={words('Compare column') + ' ' + column}
                  checked={columns.includes(column)}
                  disabled={!columns.includes(column) && columns.length >= 16}
                  onChange={(event) => {
                    if (!editable()) return;
                    invalidate();
                    setColumns((values) =>
                      event.target.checked ? [...values, column] : values.filter((value) => value !== column)
                    );
                  }}
                />
                {column}
                {hasHeaders && labels[column] ? ' — ' + labels[column] : ''}
              </label>
            ))}
            <button
              disabled={available.length > 16}
              onClick={() => {
                if (editable()) {
                  invalidate();
                  setColumns([...available]);
                }
              }}
            >
              {words('Select all columns')}
            </button>
            <button
              onClick={() => {
                if (editable()) {
                  invalidate();
                  setColumns([]);
                }
              }}
            >
              {words('Clear selected columns')}
            </button>
            <button
              disabled={columnPage === 0}
              onClick={() => {
                if (editable()) setColumnPage((value) => value - 1);
              }}
            >
              {words('Previous columns')}
            </button>
            <button
              disabled={(columnPage + 1) * 100 >= available.length}
              onClick={() => {
                if (editable()) setColumnPage((value) => value + 1);
              }}
            >
              {words('More columns')}
            </button>
          </fieldset>
          <button
            disabled={busy || !fresh || !!configReason || !!previewReason || !!outcome}
            title={previewReason ? words(previewReason) : undefined}
            onClick={() => void inspect()}
          >
            {words('Preview duplicates')}
          </button>
          <button disabled={phase === 'accepting'} onClick={cancel}>
            {words('Cancel duplicate removal')}
          </button>
          {phase === 'headings' && <p role="status">{words('Reading column headings…')}</p>}
          {phase === 'previewing' && <p role="status">{words('Comparing the selected rows…')}</p>}
          {phase === 'stopping' && <p role="status">{words('Finishing the preview…')}</p>}
          {phase === 'accepting' && <p role="status">{words('Finishing duplicate removal…')}</p>}
          {configReason && fresh && <p role="status">{words(configReason)}</p>}
          {previewReason && fresh && <p role="status">{words(previewReason)}</p>}
          {shown && !outcome && (
            <>
              <p role="status" aria-label={words('Duplicate preview counts')}>
                {shown.totalRows} {words('data rows')}: {shown.keptRows} {words('kept')}, {shown.removedRows}{' '}
                {words('duplicates')}
              </p>
              <table aria-label={words('Duplicate rows')}>
                <thead>
                  <tr>
                    <th>{words('Duplicate row')}</th>
                    <th>{words('First matching row')}</th>
                  </tr>
                </thead>
                <tbody>
                  {shown.samples.map((item) => (
                    <tr key={item.row}>
                      <td>{item.row}</td>
                      <td>{item.keptRow}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <button
                disabled={busy || shown.offset === 0}
                onClick={() => void inspect(Math.max(0, shown.offset - 100))}
              >
                {words('Previous duplicate rows')}
              </button>
              <button
                disabled={busy || shown.nextOffset === null}
                onClick={() => {
                  if (shown.nextOffset !== null) void inspect(shown.nextOffset);
                }}
              >
                {words('More duplicate rows')}
              </button>
              <button
                disabled={busy || !fresh || !shown.removedRows || !!removeReason}
                title={removeReason ? words(removeReason) : undefined}
                onClick={() => void remove()}
              >
                {words('Remove duplicates')}
              </button>
            </>
          )}
          {outcome && runtime && (
            <p role="status">
              {outcome.removedRows} {words('duplicate rows removed. Use Undo to restore the previous rows.')}
            </p>
          )}
        </>
      )}
      {error && fresh && visible && <p role="alert">{words(error)}</p>}
    </section>
  );
}
