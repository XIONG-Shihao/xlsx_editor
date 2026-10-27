import { useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react';
import {
  createWatchWindowController,
  MAX_CELL_WATCHES,
  type WatchReader,
  type WatchScope,
  type WatchTarget,
  type WatchValue,
  type WatchRow,
} from './watchWindow';
import { watchScopeKey } from './watchAdapter';
export interface WorkbookWatchProps {
  scope: WatchScope;
  read: WatchReader;
  selected: WatchTarget | null;
  selectionReason?: string;
  words?: (text: string) => string;
}
const identity = (value: string) => value;
function valueText(value: WatchValue | undefined): string {
  if (!value) return '—';
  if (value.kind === 'blank') return '(blank)';
  if (value.kind === 'text') return JSON.stringify(value.value);
  if (value.kind === 'number') return Object.is(value.value, -0) ? '-0' : String(value.value);
  if (value.kind === 'boolean') return value.value ? 'TRUE' : 'FALSE';
  if (value.kind === 'error') return value.error;
  if (value.kind === 'protected') return 'Protected number: ' + value.raw;
  if (value.kind === 'stored') return 'Stored result; not recalculated.';
  return value.reason;
}
export function WorkbookWatch({
  scope,
  read,
  selected,
  selectionReason,
  words = identity,
}: WorkbookWatchProps): JSX.Element {
  const key = watchScopeKey(scope),
    live = useRef({ scope, key, read });
  live.current = { scope, key, read };
  const controller = useMemo(
    () => createWatchWindowController((captured, target, signal) => live.current.read(captured, target, signal)),
    []
  );
  const snapshot = useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot);
  const [error, setError] = useState('');
  const refreshIntent = useRef(0);
  const lifetime = useRef(0);
  const fresh = !!snapshot.scope && watchScopeKey(snapshot.scope) === key && scope.active;
  const sameRuntime =
    snapshot.scope?.documentId === scope.documentId && snapshot.scope?.ownerIdentity === scope.ownerIdentity;
  // Render precedes layout cleanup on a changed head/tab. Keep target names for
  // the same runtime, but never flash accepted values or another owner's rows.
  const rows: readonly WatchRow[] = fresh
    ? snapshot.rows
    : sameRuntime
      ? snapshot.rows.map(({ id, sheet, reference }) => ({
          id,
          sheet,
          reference,
          status: scope.active ? ('waiting' as const) : ('paused' as const),
        }))
      : [];
  const actionCurrent = (): boolean => live.current.key === key && live.current.scope.active;
  const refresh = async (): Promise<void> => {
    const captured = key;
    if (!actionCurrent()) return;
    const intent = ++refreshIntent.current;
    try {
      let outcome = await controller.refresh();
      // A retired physical read may ignore abort. Join it first, then make one
      // fresh pass for the still-current visible scope; never overlap workers.
      if (
        outcome.outcome === 'retired' &&
        refreshIntent.current === intent &&
        live.current.key === captured &&
        live.current.scope.active
      )
        outcome = await controller.refresh();
      if (refreshIntent.current === intent && live.current.key === captured && outcome.failed)
        setError('Some cells could not be read at this workbook state.');
    } catch (failure) {
      if (refreshIntent.current === intent && live.current.key === captured)
        setError(failure instanceof Error ? failure.message : String(failure));
    }
  };
  useLayoutEffect(() => {
    setError('');
    try {
      controller.setScope(scope);
      if (scope.active && controller.getSnapshot().rows.length) void refresh();
    } catch {
      setError('Watch Window is unavailable for this workbook.');
    }
    return () => {
      refreshIntent.current++;
      controller.cancelRefresh();
    };
  }, [key]);
  useLayoutEffect(() => {
    const epoch = ++lifetime.current;
    return () => {
      refreshIntent.current++;
      queueMicrotask(() => {
        if (lifetime.current === epoch) controller.dispose();
      });
    };
  }, []);
  const add = (): void => {
    if (!actionCurrent() || !selected || selectionReason) return;
    try {
      const added = controller.add(selected);
      setError('');
      if (added) void refresh();
    } catch (failure) {
      setError(failure instanceof Error ? failure.message : String(failure));
    }
  };
  return (
    <section
      className="sheet-watch"
      aria-label={words('Watch Window')}
      data-watch-document={scope.documentId}
      data-watch-owner={scope.ownerIdentity}
      data-watch-state={scope.state}
    >
      <p>
        {words('Watch cell values and formulas without changing the workbook. Watches stay with this open document.')}
      </p>
      <div>
        <button
          type="button"
          disabled={!scope.active || !selected || !!selectionReason || rows.length >= MAX_CELL_WATCHES}
          title={selectionReason ? words(selectionReason) : undefined}
          onClick={add}
        >
          {words('Add selected cell')}
        </button>
        <button
          type="button"
          disabled={!scope.active || !rows.length || snapshot.phase === 'reading'}
          onClick={() => void refresh()}
        >
          {words('Refresh watches')}
        </button>
        <button
          type="button"
          disabled={!scope.active || !rows.length}
          onClick={() => {
            if (actionCurrent()) {
              refreshIntent.current++;
              controller.clear();
              setError('');
            }
          }}
        >
          {words('Clear watches')}
        </button>
        {snapshot.phase === 'reading' && (
          <button
            type="button"
            onClick={() => {
              if (actionCurrent()) {
                refreshIntent.current++;
                controller.cancelRefresh();
              }
            }}
          >
            {words('Cancel refresh')}
          </button>
        )}
      </div>
      {selected && (
        <p>
          {selected.sheet}!{selected.reference}
        </p>
      )}
      {selectionReason && <p role="status">{words(selectionReason)}</p>}
      <p role="status">
        {rows.length} / {MAX_CELL_WATCHES} {words('watched cells')}
        {!scope.active
          ? ' · ' + words('Paused')
          : snapshot.phase === 'reading'
            ? ' · ' + words('Reading accepted cells…')
            : ''}
      </p>
      {error && <p role="alert">{words(error)}</p>}
      {!rows.length && <p>{words('Select one cell and add it to the watch list.')}</p>}
      <ol aria-label={words('Watched cells')}>
        {rows.map((row) => (
          <li key={row.id} data-watch-target={row.sheet + '!' + row.reference} data-watch-status={row.status}>
            <strong>
              {row.sheet}!{row.reference}
            </strong>
            <span>{words(row.status === 'ready' ? (row.value?.kind ?? 'ready') : row.status)}</span>
            <output aria-label={words('Watch value') + ' ' + row.sheet + '!' + row.reference}>
              {valueText(row.value)}
            </output>
            {row.formula && (
              <output aria-label={words('Watch formula') + ' ' + row.sheet + '!' + row.reference}>
                {'=' + row.formula}
              </output>
            )}
            {row.status === 'failed' && <p role="status">{words('Value unavailable at this accepted state.')}</p>}
            <button
              type="button"
              disabled={!scope.active}
              aria-label={words('Remove watch') + ' ' + row.sheet + '!' + row.reference}
              onClick={() => {
                if (actionCurrent()) {
                  controller.remove({ sheet: row.sheet, reference: row.reference });
                  setError('');
                  if (controller.getSnapshot().rows.length) void refresh();
                }
              }}
            >
              {words('Remove')}
            </button>
          </li>
        ))}
      </ol>
    </section>
  );
}
