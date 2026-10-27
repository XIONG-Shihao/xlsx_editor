import { useEffect, useLayoutEffect, useMemo, useRef, useSyncExternalStore } from 'react';
import type { LocalExecutionAction } from '../core/local-action-execution.mjs';
import { createLocalExecutionController, type LocalExecutionFrame } from './localExecutionController';
export interface WorkbookLocalExecutionProps {
  frame: LocalExecutionFrame | null;
  intent: { serial: number; action: LocalExecutionAction } | null;
  words?: (s: string) => string;
}
export function WorkbookLocalExecution({ frame, intent, words = (s) => s }: WorkbookLocalExecutionProps): JSX.Element {
  const controller = useMemo(() => createLocalExecutionController(), []),
    retirement = useRef<ReturnType<typeof setTimeout> | null>(null),
    seen = useRef(-1),
    action = useRef<LocalExecutionAction>('unhide-all');
  const state = useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot);
  useLayoutEffect(() => {
    controller.setFrame(frame);
  });
  useEffect(() => {
    if (retirement.current !== null) clearTimeout(retirement.current);
    return () => {
      controller.setFrame(null);
      retirement.current = setTimeout(() => controller.dispose(), 0);
    };
  }, [controller]);
  useEffect(() => {
    if (!intent || !frame?.active || intent.serial <= seen.current) return;
    seen.current = intent.serial;
    action.current = intent.action;
    controller.preview(intent.action);
  }, [controller, intent?.serial, intent?.action, frame?.active]);
  const fresh = !!frame?.active && state.ownerKey === frame.ownerKey,
    preview = fresh && state.key === frame?.key ? state.preview : null;
  return (
    <section className="sheet-local-execution" aria-label={words('Run local action')} hidden={!frame?.active}>
      <p>
        {words(
          'These built-in actions run only when you press Run. Opening a workbook or saving a private definition does not run anything.'
        )}
      </p>
      {preview && (
        <div data-local-action-preview={preview.action}>
          <strong>{words(preview.label)}</strong>
          <p>
            {words('Captured worksheet')}: {preview.sheet}
          </p>
          <p>
            {words('Captured selection')}: {preview.selection}
          </p>
          {preview.action === 'freeze-selection' ? (
            <p>
              {words('Freeze preceding rows')}:{' '}
              {preview.commands[0].operation === 'freeze' ? preview.commands[0].rows : 0} ·{' '}
              {words('Freeze preceding columns')}:{' '}
              {preview.commands[0].operation === 'freeze' ? preview.commands[0].columns : 0}
            </p>
          ) : (
            <p>
              {words(
                'Show all rows and columns, keeping positive custom sizes. Worksheets with tables, filters or protection are refused.'
              )}
            </p>
          )}
        </div>
      )}
      {fresh && state.message && <p role="status">{words(state.message)}</p>}
      {frame?.reason && <p role="status">{words(frame.reason)}</p>}
      <div>
        <button
          type="button"
          disabled={!frame?.active || !!frame.reason || state.pending}
          onClick={() => controller.preview(action.current)}
        >
          {words('Preview action')}
        </button>{' '}
        <button
          type="button"
          disabled={!preview || !!frame?.reason || state.pending || state.phase !== 'preview'}
          onClick={() => void controller.run()}
        >
          {words('Run action')}
        </button>{' '}
        <button type="button" disabled={state.pending || !preview} onClick={() => controller.cancel()}>
          {words('Cancel action preview')}
        </button>
      </div>
    </section>
  );
}
