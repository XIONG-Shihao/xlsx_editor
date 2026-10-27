/** One retained typed split workflow with explicit conversions/replacements. */
import type { TextConversionPolicy } from '../core/text-conversion.mjs';
import { useLayoutEffect, useRef, useState } from 'react';
import {
  previewTextColumns,
  applyTextColumns,
  textColumnContextKey,
  type TextColumnContext,
  type TextColumnPreview,
  type TextColumnOptions,
} from './textColumns';
export interface WorkbookTextColumnsProps {
  context: TextColumnContext | null;
  visible: boolean;
  words?: (text: string) => string;
  onOutcome?: (result: Awaited<ReturnType<typeof applyTextColumns>>) => void;
}
const identity = (text: string) => text;
function outputText(value: string | number | boolean): string {
  return typeof value === 'number' && Object.is(value, -0) ? '-0' : JSON.stringify(value);
}
function prior(value: TextColumnPreview['outputs'][number]): string {
  const cell = value.previous;
  if (cell?.value === null || cell?.value === undefined) return 'Empty';
  if (typeof cell.value === 'string') return JSON.stringify(cell.value);
  if (typeof cell.value === 'boolean') return cell.value ? 'TRUE' : 'FALSE';
  return Object.is(cell.value, -0) ? '-0' : (cell.rawValue ?? String(cell.value));
}
export function WorkbookTextColumns({
  context,
  visible,
  words = identity,
  onOutcome,
}: WorkbookTextColumnsProps): JSX.Element {
  const key = context ? textColumnContextKey(context) : '',
    live = useRef({ key, visible, context });
  live.current = { key, visible, context };
  const [mode, setMode] = useState<'delimited' | 'fixed-width'>('delimited'),
    [breaks, setBreaks] = useState('');
  const [delimiter, setDelimiter] = useState('tab'),
    [custom, setCustom] = useState('|'),
    [quotes, setQuotes] = useState<'none' | 'double'>('double'),
    [destination, setDestination] = useState(context?.binding.selection.split(':')[0] ?? 'A1');
  const [policies, setPolicies] = useState<TextConversionPolicy[]>([]),
    [fieldCount, setFieldCount] = useState(0);
  const [preview, setPreview] = useState<TextColumnPreview | null>(null),
    [error, setError] = useState(''),
    [phase, setPhase] = useState<'idle' | 'previewing' | 'checking' | 'accepting' | 'stopping'>('idle'),
    [outcome, setOutcome] = useState<Awaited<ReturnType<typeof applyTextColumns>> | null>(null);
  const [rowPage, setRowPage] = useState(0),
    [columnPage, setColumnPage] = useState(0),
    [collisionPage, setCollisionPage] = useState(0);
  const serial = useRef(0),
    abort = useRef<AbortController | null>(null),
    phaseRef = useRef(phase),
    physical = useRef(false),
    mounted = useRef(true);
  phaseRef.current = phase;
  const [pending, setPending] = useState(false);
  useLayoutEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
      serial.current++;
      if (phaseRef.current !== 'accepting') abort.current?.abort();
    };
  }, []);
  const fresh = !!context && context.isCurrent(context.binding),
    runtime = !!context && (context.isRuntimeCurrent?.(context.binding) ?? fresh),
    current = () => live.current.key === key && live.current.visible;
  useLayoutEffect(() => {
    serial.current++;
    if (phaseRef.current !== 'accepting') abort.current?.abort();
    setPreview(null);
    setOutcome(null);
    setError('');
    setDestination(context?.binding.selection.split(':')[0] ?? 'A1');
    setPolicies([]);
    setFieldCount(0);
    setPhase('idle');
    return () => {
      serial.current++;
      if (phaseRef.current !== 'accepting') abort.current?.abort();
    };
  }, [key]);
  useLayoutEffect(() => {
    if (!visible && phaseRef.current !== 'accepting') {
      serial.current++;
      abort.current?.abort();
      setPhase('idle');
    }
  }, [visible]);
  const options = (): TextColumnOptions =>
    mode === 'fixed-width'
      ? {
          mode: 'fixed-width',
          breaks: breaks.split(',').map((value) => Number(value.trim())),
          destination: destination.trim().toUpperCase(),
          policies,
        }
      : {
          delimiter:
            delimiter === 'tab' ? '\t' : delimiter === 'comma' ? ',' : delimiter === 'semicolon' ? ';' : custom,
          quotes,
          destination: destination.trim().toUpperCase(),
          policies,
        };
  const reset = (): void => {
    serial.current++;
    abort.current?.abort();
    setPreview(null);
    setOutcome(null);
    setError('');
    setPhase('idle');
  };
  const inspect = async (): Promise<void> => {
    if (!context || !context.isCurrent(context.binding) || !current() || physical.current) return;
    const action = ++serial.current,
      controller = new AbortController();
    physical.current = true;
    setPending(true);
    abort.current = controller;
    setError('');
    setOutcome(null);
    setPhase('previewing');
    try {
      const result = await previewTextColumns(context, options(), { signal: controller.signal });
      if (serial.current === action && current()) {
        setPreview(result);
        setFieldCount(result.columns);
        setRowPage(0);
        setColumnPage(0);
        setCollisionPage(0);
      }
    } catch (failure) {
      if (serial.current === action && current() && !controller.signal.aborted)
        setError(failure instanceof Error ? failure.message : String(failure));
    } finally {
      physical.current = false;
      if (mounted.current) setPending(false);
      if (mounted.current && current() && (serial.current === action || phaseRef.current === 'stopping')) {
        phaseRef.current = 'idle';
        setPhase('idle');
      }
    }
  };
  const apply = async (): Promise<void> => {
    if (
      !context ||
      !context.isCurrent(context.binding) ||
      !preview ||
      preview.rejections ||
      !current() ||
      physical.current
    )
      return;
    const action = ++serial.current,
      controller = new AbortController();
    physical.current = true;
    setPending(true);
    abort.current = controller;
    setError('');
    setPhase('checking');
    try {
      const result = await applyTextColumns(context, preview, {
        signal: controller.signal,
        replaceFingerprint: preview.collisions ? preview.fingerprint : undefined,
        beforeAcceptance: () => {
          phaseRef.current = 'accepting';
          if (serial.current === action && current()) setPhase('accepting');
        },
      });
      onOutcome?.(result);
      if (serial.current === action && current() && result.runtimeCurrent) setOutcome(result);
    } catch (failure) {
      if (serial.current === action && current() && !controller.signal.aborted)
        setError(failure instanceof Error ? failure.message : String(failure));
    } finally {
      physical.current = false;
      if (mounted.current) setPending(false);
      if (mounted.current && current() && (serial.current === action || phaseRef.current === 'stopping')) {
        phaseRef.current = 'idle';
        setPhase('idle');
      }
    }
  };
  const busy = phase !== 'idle' || pending,
    access = context?.availability(),
    shown = preview && runtime && fresh ? preview : null,
    collisions = shown?.outputs.filter((output) => output.occupied) ?? [];
  const formGeneration = serial.current,
    editCurrent = () =>
      mounted.current &&
      current() &&
      fresh &&
      !!context?.isCurrent(context.binding) &&
      !busy &&
      !outcome &&
      serial.current === formGeneration;
  return (
    <section className="sheet-text-columns" aria-label={words('Text to Columns')}>
      <p>
        {words(
          'Split one selected column. Fields stay Text unless you choose a conversion. Leading zeros and text beginning with = stay text by default; empty and padded fields stay empty text.'
        )}
      </p>
      {(!runtime || (!fresh && !outcome)) && (
        <p role="status">{words('The workbook or selection changed. Reopen Text to Columns.')}</p>
      )}
      {context && runtime && (
        <>
          <p>
            {context.binding.sheet}!{context.binding.selection}
          </p>
          <label>
            {words('Split method')}
            <select
              aria-label={words('Split method')}
              value={mode}
              disabled={busy || !fresh || !!outcome}
              onChange={(event) => {
                if (!editCurrent()) return;
                reset();
                setPolicies([]);
                setFieldCount(0);
                setColumnPage(0);
                setMode(event.target.value as 'delimited' | 'fixed-width');
              }}
            >
              <option value="delimited">{words('Delimited')}</option>
              <option value="fixed-width">{words('Fixed width')}</option>
            </select>
          </label>
          {mode === 'fixed-width' ? (
            <>
              <label>
                {words('Split after characters')}
                <input
                  aria-label={words('Split after characters')}
                  placeholder="5,13,18,28"
                  value={breaks}
                  maxLength={65536}
                  disabled={busy || !fresh || !!outcome}
                  onChange={(event) => {
                    if (!editCurrent()) return;
                    reset();
                    setPolicies([]);
                    setFieldCount(0);
                    setColumnPage(0);
                    setBreaks(event.target.value);
                  }}
                />
              </label>
              <p>
                {words(
                  'Enter increasing positions separated by commas. Positions count Unicode code points: an emoji is one character; a combining accent is a separate character. Spaces and quotes stay unchanged. Short records use empty text beyond their end.'
                )}
              </p>
            </>
          ) : (
            <>
              <label>
                {words('Delimiter')}
                <select
                  aria-label={words('Delimiter')}
                  value={delimiter}
                  disabled={busy || !fresh || !!outcome}
                  onChange={(event) => {
                    if (!editCurrent()) return;
                    reset();
                    setPolicies([]);
                    setFieldCount(0);
                    setDelimiter(event.target.value);
                  }}
                >
                  <option value="tab">{words('Tab')}</option>
                  <option value="comma">{words('Comma')}</option>
                  <option value="semicolon">{words('Semicolon')}</option>
                  <option value="custom">{words('Custom')}</option>
                </select>
              </label>
              {delimiter === 'custom' && (
                <label>
                  {words('Custom delimiter')}
                  <input
                    aria-label={words('Custom delimiter')}
                    value={custom}
                    maxLength={2}
                    disabled={busy || !fresh || !!outcome}
                    onChange={(event) => {
                      if (!editCurrent()) return;
                      reset();
                      setPolicies([]);
                      setFieldCount(0);
                      setCustom(event.target.value);
                    }}
                  />
                </label>
              )}
              <label>
                {words('Quote policy')}
                <select
                  aria-label={words('Quote policy')}
                  value={quotes}
                  disabled={busy || !fresh || !!outcome}
                  onChange={(event) => {
                    if (!editCurrent()) return;
                    reset();
                    setPolicies([]);
                    setFieldCount(0);
                    setQuotes(event.target.value as 'none' | 'double');
                  }}
                >
                  <option value="none">{words('No quotes')}</option>
                  <option value="double">{words('Double quotes')}</option>
                </select>
              </label>
            </>
          )}
          <label>
            {words('Destination start')}
            <input
              aria-label={words('Destination start')}
              value={destination}
              disabled={busy || !fresh || !!outcome}
              onChange={(event) => {
                if (!editCurrent()) return;
                reset();
                setDestination(event.target.value);
              }}
            />
          </label>
          <p aria-label={words('Workbook date calendar')}>
            {words('Date calendar')}:{' '}
            {context.date1904 === true
              ? '1904'
              : context.date1904 === false
                ? '1900'
                : words('Unknown — ISO Date conversion is unavailable.')}
          </p>
          {!!fieldCount && (
            <fieldset disabled={busy || !fresh || !!outcome}>
              <legend>{words('Output field types')}</legend>
              {Array.from({ length: Math.min(10, fieldCount - columnPage * 10) }, (_, index) => {
                const field = columnPage * 10 + index;
                return (
                  <label key={field}>
                    {words('Field')} {field + 1}
                    <select
                      aria-label={words('Field') + ' ' + (field + 1) + ' ' + words('type')}
                      value={policies[field] ?? 'text'}
                      onChange={(event) => {
                        if (!editCurrent()) return;
                        reset();
                        setPolicies((values) =>
                          Array.from({ length: Math.max(values.length, field + 1) }, (_, i) =>
                            i === field ? (event.target.value as TextConversionPolicy) : (values[i] ?? 'text')
                          )
                        );
                      }}
                    >
                      <option value="text">{words('Text')}</option>
                      <option value="number">{words('Number')}</option>
                      <option value="boolean">{words('Boolean')}</option>
                      <option value="iso-date">{words('ISO Date')}</option>
                    </select>
                  </label>
                );
              })}
            </fieldset>
          )}
          {!!fieldCount && !shown && (
            <div>
              <button
                disabled={busy || columnPage === 0}
                onClick={() => {
                  if (editCurrent()) setColumnPage((page) => page - 1);
                }}
              >
                {words('Previous field types')}
              </button>
              <button
                disabled={busy || (columnPage + 1) * 10 >= fieldCount}
                onClick={() => {
                  if (editCurrent()) setColumnPage((page) => page + 1);
                }}
              >
                {words('More field types')}
              </button>
            </div>
          )}
          <button type="button" disabled={busy || !fresh || !!outcome} onClick={() => void inspect()}>
            {words('Preview split')}
          </button>
          {(phase === 'previewing' || phase === 'checking' || (phase === 'idle' && !!fieldCount && !outcome)) && (
            <button
              type="button"
              onClick={() => {
                if (phaseRef.current === 'accepting') return;
                if (phaseRef.current === 'idle') {
                  if (!editCurrent()) return;
                  reset();
                  setPolicies([]);
                  return;
                }
                if (!current() || !context?.isCurrent(context.binding)) return;
                serial.current++;
                abort.current?.abort();
                phaseRef.current = 'stopping';
                setPhase('stopping');
              }}
            >
              {words('Cancel split')}
            </button>
          )}
          {phase === 'stopping' && <p role="status">{words('Stopping the preview…')}</p>}
          {phase === 'accepting' && <p role="status">{words('Finishing workbook changes…')}</p>}
          {access?.readOnly && (
            <p>{words('This workbook is read-only. Preview is available; splitting is disabled.')}</p>
          )}
          {shown && !outcome && (
            <>
              <p role="status">
                {shown.source} → {shown.destination} · {shown.outputs.length} {words('output cells')} ·{' '}
                {shown.collisions} {words('occupied cells will be replaced')}
              </p>
              {!!shown.rejections && (
                <p role="alert" aria-label={words('Rejected conversions')}>
                  {shown.rejections}{' '}
                  {words('rejected conversions. Change the field types or correct the source before applying.')}
                </p>
              )}
              <table aria-label={words('Split preview')}>
                <thead>
                  <tr>
                    <th>{words('Source cell')}</th>
                    <th>{words('Original text')}</th>
                    {Array.from({ length: Math.min(10, shown.columns - columnPage * 10) }, (_, index) => (
                      <th key={index}>
                        {words('Field')} {columnPage * 10 + index + 1}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {shown.rows.slice(rowPage * 20, (rowPage + 1) * 20).map((row, rowIndex) => (
                    <tr key={row.source}>
                      <th>{row.source}</th>
                      <td>{JSON.stringify(row.text)}</td>
                      {Array.from({ length: Math.min(10, shown.columns - columnPage * 10) }, (_, index) => {
                        const field = columnPage * 10 + index,
                          output = shown.outputs[(rowPage * 20 + rowIndex) * shown.columns + field];
                        return (
                          <td key={index}>
                            <strong>{output.reference}</strong>
                            <output aria-label={words('Output value') + ' ' + output.reference}>
                              {outputText(output.value)}
                            </output>
                            <span aria-label={words('Output type') + ' ' + output.reference}>
                              {words(output.conversion.status === 'converted' ? output.conversion.kind : 'rejected')}
                            </span>
                            {output.conversion.status === 'rejected' && (
                              <span role="alert" aria-label={words('Conversion issue') + ' ' + output.reference}>
                                {words(output.conversion.reason)}
                              </span>
                            )}
                            {output.conversion.notices.map((notice, i) => (
                              <small key={i}>{words(notice)}</small>
                            ))}
                            {output.padded ? ' · ' + words('padded empty text') : ''}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
              <button disabled={rowPage === 0} onClick={() => setRowPage((page) => page - 1)}>
                {words('Previous preview rows')}
              </button>
              <button disabled={(rowPage + 1) * 20 >= shown.rows.length} onClick={() => setRowPage((page) => page + 1)}>
                {words('More preview rows')}
              </button>
              <button disabled={columnPage === 0} onClick={() => setColumnPage((page) => page - 1)}>
                {words('Previous preview fields')}
              </button>
              <button
                disabled={(columnPage + 1) * 10 >= shown.columns}
                onClick={() => setColumnPage((page) => page + 1)}
              >
                {words('More preview fields')}
              </button>
              {!!collisions.length && (
                <div aria-label={words('Destination replacements')}>
                  {collisions.slice(collisionPage * 20, (collisionPage + 1) * 20).map((output) => (
                    <p key={output.reference}>
                      <strong>{output.reference}</strong> {prior(output)} → {outputText(output.value)}
                      {output.padded ? ' · ' + words('padded empty text') : ''}
                    </p>
                  ))}
                  <button disabled={collisionPage === 0} onClick={() => setCollisionPage((page) => page - 1)}>
                    {words('Previous replacements')}
                  </button>
                  <button
                    disabled={(collisionPage + 1) * 20 >= collisions.length}
                    onClick={() => setCollisionPage((page) => page + 1)}
                  >
                    {words('More replacements')}
                  </button>
                </div>
              )}
              <button
                type="button"
                disabled={busy || !fresh || !!shown.rejections || access?.readOnly || access?.busy || !context.accept}
                onClick={() => void apply()}
              >
                {words(
                  shown.collisions
                    ? 'Split and replace ' + shown.collisions + ' occupied cells'
                    : 'Split into ' + shown.destination
                )}
              </button>
            </>
          )}
          {outcome && (
            <p role="status">
              {words('Split accepted')} · {outcome.destination} · {outcome.positions}{' '}
              {words('output cells. Use Undo to restore the previous cells.')}
            </p>
          )}
        </>
      )}
      {error && runtime && <p role="alert">{error}</p>}
    </section>
  );
}
