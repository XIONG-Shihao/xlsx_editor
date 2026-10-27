import { useLayoutEffect, useRef, useState } from 'react';
import type { WorkbookSortKey } from '../core/contracts';
import {
  applySelectionSort,
  buildSelectionSort,
  readSortHeaders,
  sortContextKey,
  sortSelectionColumns,
  type SortContext,
} from './sortSelection';
export interface WorkbookSortProps {
  context: SortContext | null;
  visible: boolean;
  words?: (text: string) => string;
  onHeaderChange?: (header: boolean) => void;
}
const identity = (text: string) => text;
export function WorkbookSort({ context, visible, words = identity, onHeaderChange }: WorkbookSortProps): JSX.Element {
  const key = context ? sortContextKey(context) : '',
    live = useRef({ key, visible });
  live.current = { key, visible };
  const [keys, setKeys] = useState<WorkbookSortKey[]>([]),
    [header, setHeader] = useState(true),
    [labels, setLabels] = useState<Record<string, string>>({}),
    [phase, setPhase] = useState<'idle' | 'reading' | 'accepting' | 'stopping'>('idle'),
    [error, setError] = useState(''),
    [accepted, setAccepted] = useState(false),
    [readEpoch, setReadEpoch] = useState(0),
    [headingsError, setHeadingsError] = useState(false);
  const serial = useRef(0),
    physical = useRef(false),
    abort = useRef<AbortController | null>(null),
    mounted = useRef(true);
  const headersRead = useRef('');
  const current = () => mounted.current && live.current.key === key && live.current.visible;
  const fresh = !!context && context.isCurrent(context.binding),
    runtime = !!context && context.isRuntimeCurrent(context.binding);
  const columns = context ? sortSelectionColumns(context.binding.selection) : [];
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
    setKeys(context ? [{ column: context.initialColumn, direction: 'asc' }] : []);
    setHeader(context?.initialHeader ?? true);
    setLabels({});
    setError('');
    setAccepted(false);
    setHeadingsError(false);
    setPhase(physical.current ? 'reading' : 'idle');
  }, [key]);
  useLayoutEffect(() => {
    if (!visible || !context || !context.isCurrent(context.binding) || physical.current || headersRead.current === key)
      return;
    const action = ++serial.current,
      controller = new AbortController();
    headersRead.current = key;
    abort.current = controller;
    physical.current = true;
    setPhase('reading');
    void readSortHeaders(context, controller.signal)
      .then((value) => {
        if (current() && serial.current === action) setLabels(value);
      })
      .catch((failure) => {
        if (
          current() &&
          serial.current === action &&
          context.isCurrent(context.binding) &&
          !controller.signal.aborted
        ) {
          setHeadingsError(true);
          setError(failure instanceof Error ? failure.message : String(failure));
        }
      })
      .finally(() => {
        physical.current = false;
        if (mounted.current) {
          setPhase('idle');
          if (live.current.key !== key) setReadEpoch((value) => value + 1);
        }
      });
    return () => {
      serial.current++;
      controller.abort();
    };
  }, [key, visible, readEpoch]);
  const disabledReason = context?.disabledReason(),
    busy = phase !== 'idle';
  let planReason = '';
  if (context && keys.length)
    try {
      buildSelectionSort(context.binding.sheet, context.binding.selection, header, keys);
    } catch (failure) {
      planReason = failure instanceof Error ? failure.message : String(failure);
    }
  const editCurrent = () => current() && fresh && !busy && !accepted && !disabledReason;
  const change = (index: number, patch: Partial<WorkbookSortKey>): void => {
    if (!editCurrent()) return;
    setError('');
    setKeys((values) => values.map((value, i) => (i === index ? { ...value, ...patch } : value)));
  };
  const move = (index: number, offset: number): void => {
    if (!editCurrent()) return;
    setKeys((values) => {
      const next = [...values];
      [next[index], next[index + offset]] = [next[index + offset], next[index]];
      return next;
    });
  };
  const apply = async (): Promise<void> => {
    if (
      !context ||
      !fresh ||
      !current() ||
      physical.current ||
      busy ||
      planReason ||
      disabledReason ||
      headingsError ||
      accepted
    )
      return;
    const action = ++serial.current;
    physical.current = true;
    setPhase('accepting');
    setError('');
    try {
      const result = await applySelectionSort(context, header, structuredClone(keys));
      if (current() && serial.current === action && result.runtimeCurrent) setAccepted(true);
    } catch (failure) {
      if (current() && serial.current === action && context.isCurrent(context.binding))
        setError(failure instanceof Error ? failure.message : String(failure));
    } finally {
      physical.current = false;
      if (current() && serial.current === action) setPhase('idle');
    }
  };
  return (
    <section
      className="sheet-sort"
      aria-label={words('Sort')}
      data-sort-state={context?.binding.state}
      data-sort-selection={context?.binding.selection}
    >
      {context && runtime && (
        <>
          <p>
            <strong>{context.caption}</strong>
          </p>
          <p>{words('Sort rows together, using the levels in this order.')}</p>
        </>
      )}
      {!fresh && !accepted && <p role="status">{words('The workbook or selection changed. Reopen Sort.')}</p>}
      {context && runtime && (
        <>
          <fieldset disabled={busy || !fresh || accepted || !!disabledReason}>
            <legend>{words('Sort levels')}</legend>
            <label>
              <input
                type="checkbox"
                aria-label={words('My data has headers')}
                checked={header}
                onChange={(event) => {
                  if (!editCurrent()) return;
                  setHeader(event.target.checked);
                  setError('');
                  onHeaderChange?.(event.target.checked);
                }}
              />
              {words('My data has headers')}
            </label>
            <ol>
              {keys.map((level, index) => (
                <li key={index} data-sort-level={index + 1}>
                  <label>
                    {words(index ? 'Then by' : 'Sort by')}
                    <select
                      aria-label={words('Sort column') + ' ' + (index + 1)}
                      value={level.column}
                      onChange={(event) => change(index, { column: event.target.value })}
                    >
                      {columns.map((column) => (
                        <option key={column} value={column}>
                          {column}
                          {header && labels[column] ? ' — ' + labels[column] : ''}
                        </option>
                      ))}
                    </select>
                  </label>
                  <label>
                    {words('Order')}
                    <select
                      aria-label={words('Sort order') + ' ' + (index + 1)}
                      value={level.direction}
                      onChange={(event) => change(index, { direction: event.target.value as 'asc' | 'desc' })}
                    >
                      <option value="asc">{words('Ascending')}</option>
                      <option value="desc">{words('Descending')}</option>
                    </select>
                  </label>
                  <div>
                    <button
                      type="button"
                      aria-label={words('Move sort level') + ' ' + (index + 1) + ' ' + words('up')}
                      disabled={index === 0}
                      onClick={() => move(index, -1)}
                    >
                      ↑
                    </button>
                    <button
                      type="button"
                      aria-label={words('Move sort level') + ' ' + (index + 1) + ' ' + words('down')}
                      disabled={index === keys.length - 1}
                      onClick={() => move(index, 1)}
                    >
                      ↓
                    </button>
                    <button
                      type="button"
                      aria-label={words('Remove sort level') + ' ' + (index + 1)}
                      disabled={keys.length === 1}
                      onClick={() => {
                        if (editCurrent()) setKeys((values) => values.filter((_, i) => i !== index));
                      }}
                    >
                      {words('Remove')}
                    </button>
                  </div>
                </li>
              ))}
            </ol>
            <button
              type="button"
              disabled={keys.length >= Math.min(16, columns.length)}
              onClick={() => {
                if (!editCurrent()) return;
                const column = columns.find((value) => !keys.some((level) => level.column === value));
                if (column) setKeys((values) => [...values, { column, direction: 'asc' }]);
              }}
            >
              {words('Add level')}
            </button>
          </fieldset>
          {phase === 'reading' && fresh && <p role="status">{words('Reading column headings…')}</p>}
          {phase === 'stopping' && <p role="status">{words('Finishing the header read…')}</p>}
          {phase === 'accepting' && <p role="status">{words('Finishing sort…')}</p>}
          {accepted && <p role="status">{words('Rows sorted. Use Undo to restore their previous order.')}</p>}
          {(planReason || disabledReason) && fresh && !accepted && (
            <p role="status">{words(planReason || disabledReason!)}</p>
          )}
          <button
            type="button"
            disabled={busy || !fresh || accepted || !!planReason || !!disabledReason || headingsError || !keys.length}
            onClick={() => void apply()}
          >
            {words('Apply sort')}
          </button>
          <button
            type="button"
            disabled={!fresh || accepted || phase === 'accepting' || phase === 'stopping'}
            onClick={() => {
              if (!current() || !fresh || accepted || phase === 'accepting' || phase === 'stopping') return;
              serial.current++;
              abort.current?.abort();
              setError(headingsError ? 'Reopen Sort to read column headings.' : '');
              setKeys([{ column: context.initialColumn, direction: 'asc' }]);
              setHeader(context.initialHeader);
              onHeaderChange?.(context.initialHeader);
              setPhase(physical.current ? 'stopping' : 'idle');
            }}
          >
            {words('Cancel sort')}
          </button>
        </>
      )}
      {error && fresh && visible && <p role="alert">{words(error)}</p>}
    </section>
  );
}
