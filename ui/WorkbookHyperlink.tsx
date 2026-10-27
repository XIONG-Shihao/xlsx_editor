import { useEffect, useLayoutEffect, useRef, useState, useSyncExternalStore } from 'react';
import { createHyperlinkController, type HyperlinkFrame } from './hyperlinkController';
export interface WorkbookHyperlinkProps {
  frame: HyperlinkFrame | null;
  words?: (text: string) => string;
  sheets: string[];
}
export function WorkbookHyperlink({ frame, words = (text: string) => text, sheets }: WorkbookHyperlinkProps) {
  const [controller] = useState(createHyperlinkController),
    lifetime = useRef(0);
  const snapshot = useSyncExternalStore(controller.subscribe, controller.getSnapshot);
  useLayoutEffect(() => {
    controller.setFrame(frame);
  });
  useEffect(() => {
    const serial = ++lifetime.current;
    return () => {
      queueMicrotask(() => {
        if (lifetime.current === serial) controller.dispose();
      });
    };
  }, [controller]);
  const [draft, setDraft] = useState({
      key: '',
      kind: 'url' as 'url' | 'internal',
      target: '',
      sheet: '',
      cell: 'A1',
      tooltip: '',
      text: '',
    }),
    [error, setError] = useState({ key: '', text: '' });
  const definition = snapshot.key === frame?.key ? snapshot.definition : null;
  useEffect(() => {
    if (!frame || !definition) return;
    const link = definition.link;
    setDraft({
      key: frame.key,
      kind: link?.kind ?? 'url',
      target: link?.kind === 'url' ? link.target : '',
      sheet: link?.kind === 'internal' ? link.sheet : frame.sheet,
      cell: link?.kind === 'internal' ? link.cell : 'A1',
      tooltip: definition.tooltip,
      text: typeof snapshot.cell?.value === 'string' ? snapshot.cell.value : '',
    });
    setError({ key: '', text: '' });
  }, [definition, frame?.key]);
  const current =
    draft.key === frame?.key
      ? draft
      : {
          key: frame?.key ?? '',
          kind: 'url' as const,
          target: '',
          sheet: frame?.sheet ?? '',
          cell: 'A1',
          tooltip: '',
          text: '',
        };
  const cell = snapshot.cell,
    labelEditable =
      snapshot.labelAvailable &&
      (!cell ||
        (!cell.formula &&
          !cell.formulaType &&
          !cell.spill &&
          (cell.dataType === 'text' || cell.dataType === 'blank' || cell.dataType === undefined) &&
          (cell.value == null || typeof cell.value === 'string')));
  const disabled = !frame?.active || !frame.writable || !definition?.canEdit || snapshot.loading || snapshot.pending;
  const edit = (patch: Partial<typeof draft>) => {
    if (frame?.active && snapshot.key === frame.key) {
      setDraft({ ...current, ...patch, key: frame.key });
      setError({ key: '', text: '' });
    }
  };
  const run = async (operation: () => Promise<unknown>) => {
    const key = frame?.key;
    try {
      setError({ key: '', text: '' });
      await operation();
    } catch (cause) {
      if (key === frame?.key)
        setError({ key: key ?? '', text: String(cause instanceof Error ? cause.message : cause) });
    }
  };
  return (
    <section aria-label={words('Cell hyperlink')}>
      <p>
        {words('Selected cell')}: {frame ? frame.sheet + '!' + frame.cell : words('Select one regular cell.')}
      </p>
      <p>
        {words(
          'Formatting is preserved. Plain-text labels can be edited; a new link in an empty cell uses its address.'
        )}
      </p>
      {frame?.reason && <p role="status">{words(frame.reason)}</p>}
      {definition?.protectedReason && <p role="status">{words(definition.protectedReason)}</p>}
      {snapshot.loading && <p role="status">{words('Loading cell link…')}</p>}
      <form
        onSubmit={(event) => {
          event.preventDefault();
          if (!frame || disabled) return;
          void run(() =>
            controller.write(
              frame,
              current.kind === 'url'
                ? { kind: 'url', target: current.target }
                : { kind: 'internal', sheet: current.sheet, cell: current.cell },
              current.tooltip,
              labelEditable ? current.text : undefined
            )
          );
        }}
      >
        <fieldset disabled={disabled}>
          <legend>{words(definition?.link ? 'Edit link' : 'Insert link')}</legend>
          {labelEditable ? (
            <label>
              {words('Text to display')}
              <input
                aria-label={words('Text to display')}
                value={current.text}
                maxLength={32767}
                placeholder={words('Use the address for an empty cell')}
                onChange={(event) => edit({ text: event.target.value })}
              />
            </label>
          ) : (
            <small>{words('Edit numeric or formula cell text in the worksheet.')}</small>
          )}
          <label>
            {words('Link destination')}
            <select
              aria-label={words('Link destination')}
              value={current.kind}
              onChange={(event) => edit({ kind: event.target.value as 'url' | 'internal' })}
            >
              <option value="url">{words('Web address')}</option>
              <option value="internal">{words('Worksheet cell')}</option>
            </select>
          </label>
          {current.kind === 'url' ? (
            <label>
              {words('Web address')}
              <input
                aria-label={words('Web address')}
                type="url"
                value={current.target}
                maxLength={2048}
                placeholder="https://example.com/"
                required
                onChange={(event) => edit({ target: event.target.value })}
              />
            </label>
          ) : (
            <>
              <label>
                {words('Target worksheet')}
                <select
                  aria-label={words('Target worksheet')}
                  value={current.sheet}
                  onChange={(event) => edit({ sheet: event.target.value })}
                >
                  {sheets.map((sheet) => (
                    <option key={sheet}>{sheet}</option>
                  ))}
                </select>
              </label>
              <label>
                {words('Target cell')}
                <input
                  aria-label={words('Target cell')}
                  value={current.cell}
                  maxLength={10}
                  required
                  onChange={(event) => edit({ cell: event.target.value })}
                />
              </label>
            </>
          )}
          <label>
            {words('Screen tip')}
            <input
              aria-label={words('Screen tip')}
              value={current.tooltip}
              maxLength={256}
              onChange={(event) => edit({ tooltip: event.target.value })}
            />
          </label>
          <button type="submit">{words(definition?.link ? 'Apply link changes' : 'Insert link')}</button>
        </fieldset>
      </form>
      <div>
        <button
          type="button"
          disabled={disabled || !definition?.link}
          onClick={() => {
            if (frame) void run(() => controller.write(frame, null));
          }}
        >
          {words('Remove link')}
        </button>
        <button
          type="button"
          disabled={
            !frame?.active ||
            snapshot.loading ||
            snapshot.pending ||
            !definition?.link ||
            !definition.canEdit ||
            !frame.follow
          }
          onClick={() => {
            if (frame) void run(() => controller.follow(frame));
          }}
        >
          {words('Open link')}
        </button>
        <button
          type="button"
          disabled={!frame?.active || snapshot.loading || snapshot.pending}
          onClick={() => {
            setError({ key: '', text: '' });
            controller.refresh();
          }}
        >
          {words('Reset link form')}
        </button>
      </div>
      {snapshot.pending && <p role="status">{words('Applying link operation…')}</p>}
      {((snapshot.key === frame?.key ? snapshot.error : '') || (error.key === frame?.key ? error.text : '')) && (
        <p role="alert">
          {words((snapshot.key === frame?.key ? snapshot.error : '') || (error.key === frame?.key ? error.text : ''))}
        </p>
      )}
      {snapshot.key === frame?.key && snapshot.notice && <p role="status">{words(snapshot.notice)}</p>}
    </section>
  );
}
