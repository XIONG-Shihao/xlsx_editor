import { useLayoutEffect, useRef, useState } from 'react';
import { replacementPreviewPage, type ReplacePreview } from './replacePreview';
import { captureHomeBinding, homeBindingMatches, type HomeBinding } from './homeCommands';

export interface ReplaceContext {
  readonly key: string;
  readonly binding: HomeBinding;
  readonly caption: string;
  isCurrent(): boolean;
  isRuntimeCurrent(): boolean;
  disabledReason(action: 'preview' | 'apply'): string | undefined;
  preview(signal: AbortSignal): Promise<ReplacePreview>;
  apply(preview: ReplacePreview): Promise<void>;
}
export interface WorkbookReplaceProps {
  context: ReplaceContext | null;
  visible: boolean;
  words?: (text: string) => string;
}
const identity = (text: string) => text;

/** The original owner re-reads the complete search before one acknowledged
 * batch. Pages are views of a captured proposal, never independent commits. */
export function WorkbookReplace({ context, visible, words = identity }: WorkbookReplaceProps): JSX.Element {
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
  const [preview, setPreview] = useState<{ context: ReplaceContext; value: ReplacePreview } | null>(null),
    [offset, setOffset] = useState(0),
    [error, setError] = useState(''),
    [accepted, setAccepted] = useState(false);
  const current = () => mounted.current && live.current.key === key && live.current.visible;
  const fresh = !!context && context.isCurrent(),
    runtime = !!context && context.isRuntimeCurrent(),
    busy = phase !== 'idle';
  const previewReason = context?.disabledReason('preview'),
    applyReason = context?.disabledReason('apply');
  const shown =
    visible && fresh && preview && preview.context.key === key && preview.context.isCurrent() && !accepted
      ? preview
      : null;
  const page = shown ? replacementPreviewPage(shown.value, offset, 100) : null;

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
      const value = structuredClone(await captured.preview(controller.signal));
      if (!current() || generation.current !== serial || controller.signal.aborted || !captured.isCurrent()) return;
      if (!homeBindingMatches(value.binding, captured.binding))
        throw Error('The replacement preview belongs to a different workbook or selection.');
      replacementPreviewPage(value, 0, 100);
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
      !shown.value.editable ||
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
    if (current() && context?.isCurrent() && shown?.context.isCurrent() && !busy && !physical.current && !accepted)
      setOffset(next);
  };

  return (
    <section
      className="sheet-replace"
      aria-label={words('Replacement preview')}
      data-formula-draft-zone
      data-replace-state={context?.binding.state}
    >
      {visible && context && runtime && (
        <>
          <p>
            <strong>{context.caption}</strong>
          </p>
          <p>{words('Preview all matches in the chosen scope, then apply the editable replacements together.')}</p>
        </>
      )}
      {visible && !context && <p role="status">{words('Replacement preview is unavailable here.')}</p>}
      {visible && context && !fresh && !accepted && (
        <p role="status">{words('The workbook, selection or search changed. Reopen replacement preview.')}</p>
      )}
      {visible && context && runtime && (fresh || accepted) && (
        <>
          <div className="sheet-replace-actions">
            <button
              type="button"
              data-home-command="home.replace-preview"
              disabled={busy || !fresh || accepted || !!previewReason}
              title={previewReason ? words(previewReason) : undefined}
              onClick={() => void inspect()}
            >
              {words('Preview all replacements')}
            </button>
            {shown && page && (
              <button
                type="button"
                data-home-command="home.replace-all"
                disabled={busy || !fresh || accepted || !shown.value.editable || !!applyReason}
                title={applyReason ? words(applyReason) : undefined}
                onClick={() => void apply()}
              >
                {words('Apply replacements')}
              </button>
            )}
            <button type="button" disabled={!fresh || accepted || phase === 'accepting'} onClick={cancel}>
              {words('Cancel replacements')}
            </button>
          </div>
          {previewReason && fresh && !accepted && <p role="status">{words(previewReason)}</p>}
          {applyReason && applyReason !== previewReason && fresh && !accepted && (
            <p role="status">{words(applyReason)}</p>
          )}
          {phase === 'reading' && fresh && <p role="status">{words('Reading all matching cells…')}</p>}
          {phase === 'stopping' && <p role="status">{words('Finishing the replacement preview…')}</p>}
          {phase === 'accepting' && <p role="status">{words('Finishing the replacements…')}</p>}
          {shown && page && (
            <>
              <dl aria-label={words('Replacement summary')}>
                <dt>{words('Matched cells')}</dt>
                <dd>{shown.value.matched}</dd>
                <dt>{words('Editable replacements')}</dt>
                <dd>{shown.value.editable}</dd>
                <dt>{words('Skipped cells')}</dt>
                <dd>{shown.value.skipped}</dd>
                <dt>{words('Search content')}</dt>
                <dd>
                  {words(
                    shown.value.configuration.mode === 'formula'
                      ? 'Formulas'
                      : shown.value.configuration.mode === 'raw'
                        ? 'Raw values'
                        : 'Displayed values'
                  )}
                </dd>
                <dt>{words('Worksheets')}</dt>
                <dd>{shown.value.configuration.sheets.join(', ')}</dd>
                <dt>{words('Find')}</dt>
                <dd>
                  <pre>{shown.value.configuration.text}</pre>
                </dd>
                <dt>{words('Replace with')}</dt>
                <dd>
                  <pre>{shown.value.configuration.replacement}</pre>
                </dd>
              </dl>
              <p>{words('Text keeps its type. Protected or unsupported cells may prevent Apply.')}</p>
              <div className="sheet-replace-rows">
                <table aria-label={words('Replacement targets')}>
                  <thead>
                    <tr>
                      <th>{words('Cell')}</th>
                      <th>{words('Before')}</th>
                      <th>{words('After / reason')}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {page.rows.map((row) => (
                      <tr key={`${row.sheet}\0${row.cell}`}>
                        <td>
                          {row.sheet}!{row.cell}
                        </td>
                        <td>
                          <pre>{row.before}</pre>
                        </td>
                        <td>
                          {row.after !== null && <pre>{row.after}</pre>}
                          {row.reason && <span>{words(row.reason)}</span>}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <p role="status" aria-label={words('Replacement cells shown')}>
                {page.total ? `${page.offset + 1}–${page.offset + page.rows.length}` : '0'} {words('of')} {page.total}
              </p>
              <button
                type="button"
                disabled={busy || page.offset === 0}
                onClick={() => move(Math.max(0, page.offset - 100))}
              >
                {words('Previous replacement cells')}
              </button>
              <button
                type="button"
                disabled={busy || page.nextOffset === null}
                onClick={() => {
                  if (page.nextOffset !== null) move(page.nextOffset);
                }}
              >
                {words('More replacement cells')}
              </button>
            </>
          )}
          {accepted && runtime && (
            <p role="status">{words('Replacements applied. Use Undo to restore the previous cells.')}</p>
          )}
        </>
      )}
      {error && fresh && visible && <p role="alert">{words(error)}</p>}
    </section>
  );
}
