import { useEffect, useLayoutEffect, useRef, useState, useSyncExternalStore } from 'react';
import { booleanCellValue, symbolCellEdit, symbolSamples } from '../core/cell-insertions.mjs';
import { createCellInsertionController, type CellInsertionFrame } from './cellInsertionController';
export function WorkbookCellInsertions({
  frame,
  mode,
  words = (text: string) => text,
}: {
  frame: CellInsertionFrame | null;
  mode: 'checkbox' | 'symbols';
  words?: (text: string) => string;
}) {
  const [controller] = useState(createCellInsertionController),
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
  const [error, setError] = useState({ key: '', text: '' });
  let value: boolean | null = null,
    reason = '';
  const current = snapshot.key === frame?.key && snapshot.loaded;
  if (current)
    try {
      if (mode === 'checkbox') value = booleanCellValue(snapshot.cell);
      else symbolCellEdit(frame!.sheet, frame!.cell, snapshot.cell, 'omega');
    } catch (e) {
      reason = String(e instanceof Error ? e.message : e);
    }
  const disabled = !frame?.active || !frame.writable || !current || snapshot.loading || snapshot.pending || !!reason;
  const run = async (operation: () => Promise<unknown>) => {
    const key = frame?.key;
    setError({ key: '', text: '' });
    try {
      await operation();
    } catch (e) {
      setError({ key: key ?? '', text: String(e instanceof Error ? e.message : e) });
    }
  };
  return (
    <section aria-label={words(mode === 'checkbox' ? 'Boolean-cell checkbox' : 'Literal Symbols samples')}>
      <p>
        {words('Selected cell')}: {frame ? frame.sheet + '!' + frame.cell : words('Select one regular cell.')}
      </p>
      {mode === 'checkbox' ? (
        <>
          <p>
            {words(
              'This editor uses ordinary TRUE/FALSE cells as checkboxes. Excel control metadata and imported form controls are protected.'
            )}
          </p>
          <fieldset disabled={disabled}>
            <legend>{words('Boolean-cell checkbox')}</legend>
            {value !== null ? (
              <>
                <label>
                  <input
                    type="checkbox"
                    aria-label={words('Selected cell checked')}
                    checked={value}
                    onChange={() => {
                      if (frame && !disabled) void run(() => controller.boolean(frame, 'toggle'));
                    }}
                  />
                  {words(value ? 'TRUE — checked' : 'FALSE — unchecked')}
                </label>
                <button
                  type="button"
                  onClick={() => {
                    if (frame && !disabled) void run(() => controller.boolean(frame, 'remove'));
                  }}
                >
                  {words('Remove checkbox')}
                </button>
              </>
            ) : (
              <button
                type="button"
                onClick={() => {
                  if (frame && !disabled) void run(() => controller.boolean(frame, 'insert'));
                }}
              >
                {words('Insert checkbox')}
              </button>
            )}
          </fieldset>
        </>
      ) : (
        <>
          <p>
            {words(
              'Append a literal symbol to a blank or plain-text cell. These three samples keep existing formatting.'
            )}
          </p>
          <fieldset disabled={disabled}>
            <legend>{words('Symbols samples')}</legend>
            {symbolSamples.map((sample) => (
              <button
                type="button"
                key={sample.id}
                aria-label={words(sample.label) + ' (' + sample.value + ')'}
                onClick={() => {
                  if (frame && !disabled) void run(() => controller.symbol(frame, sample.id));
                }}
              >
                <span aria-hidden="true">{sample.value}</span> {words(sample.label)}
              </button>
            ))}
          </fieldset>
        </>
      )}
      <button
        type="button"
        disabled={!frame?.active || snapshot.loading || snapshot.pending}
        onClick={() => controller.refresh()}
      >
        {words('Refresh selected cell')}
      </button>
      {(frame?.reason || reason) && <p role="status">{words(frame?.reason ?? reason)}</p>}
      {snapshot.loading && <p role="status">{words('Loading selected cell…')}</p>}
      {snapshot.pending && <p role="status">{words('Applying cell insertion…')}</p>}
      {(snapshot.error || (error.key === frame?.key ? error.text : '')) && (
        <p role="alert">{words(snapshot.error || (error.key === frame?.key ? error.text : ''))}</p>
      )}
      {snapshot.notice && <p role="status">{words(snapshot.notice)}</p>}
    </section>
  );
}
