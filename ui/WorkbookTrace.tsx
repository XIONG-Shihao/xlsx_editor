import { useLayoutEffect, useRef, useState, useSyncExternalStore } from 'react';
import {
  formulaTraceScopeKey,
  type FormulaTraceController,
  type FormulaTraceScope,
  type TraceDirection,
  type TraceEndpoint,
} from './formulaTrace';
export interface WorkbookTraceProps {
  controller: FormulaTraceController;
  scope: FormulaTraceScope;
  onNavigate: (target: TraceEndpoint) => void;
  disabledReason?: string;
  words?: (text: string) => string;
}
const identity = (text: string) => text;
export function WorkbookTrace({
  controller,
  scope,
  onNavigate,
  disabledReason,
  words = identity,
}: WorkbookTraceProps): JSX.Element {
  const key = formulaTraceScopeKey(scope),
    live = useRef({ key, scope, onNavigate, controller, disabledReason });
  live.current = { key, scope, onNavigate, controller, disabledReason };
  const snapshot = useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot),
    intent = useRef(0),
    [feedback, setFeedback] = useState('');
  const fresh = scope.active && !disabledReason && !!snapshot.scope && formulaTraceScopeKey(snapshot.scope) === key;
  const edges = fresh ? snapshot.edges : [],
    markers = fresh ? snapshot.markers : [],
    busy = fresh && snapshot.phase === 'reading';
  const current = () =>
    live.current.key === key &&
    live.current.scope.active &&
    live.current.controller === controller &&
    !live.current.disabledReason;
  useLayoutEffect(() => {
    intent.current++;
    setFeedback('');
    controller.setScope(scope);
    return () => {
      intent.current++;
      controller.cancel(scope);
    };
  }, [key, controller]);
  const trace = async (direction: TraceDirection): Promise<void> => {
    if (!current()) return;
    const action = ++intent.current;
    setFeedback('');
    try {
      const outcome = await controller.trace(direction);
      if (current() && intent.current === action && outcome.outcome === 'retired')
        setFeedback('The previous trace stopped. Trace again to read this workbook state.');
    } catch {
      if (current() && intent.current === action) setFeedback('Relationships could not be read. Trace again to retry.');
    }
  };
  const go = (target: TraceEndpoint): void => {
    if (!current() || !fresh || controller.getSnapshot() !== snapshot) return;
    try {
      onNavigate(target);
    } catch {
      if (current()) setFeedback('The related cell could not be selected.');
    }
  };
  const status = disabledReason
    ? disabledReason
    : !scope.active
      ? 'Trace paused. Return to Normal view in this document.'
      : busy
        ? 'Reading known formula relationships…'
        : !fresh || !snapshot.direction
          ? 'Trace the selected cell to show its known relationships.'
          : snapshot.complete
            ? 'No further known static relationships.'
            : snapshot.canExpand
              ? 'More levels are available. Trace again to expand the next level.'
              : 'Some relationships are not covered by this trace.';
  return (
    <section
      className="sheet-trace"
      aria-label={words('Formula tracing')}
      data-trace-document={scope.documentId}
      data-trace-owner={scope.ownerIdentity}
      data-trace-state={scope.state}
      data-trace-root={scope.sheet + '!' + scope.cell}
    >
      <p>
        <strong>
          {scope.sheet}!{scope.cell}
        </strong>
      </p>
      <p>
        {words(
          'Arrows show known references. Dynamic, external and unsupported relationships are listed separately. This does not change the workbook.'
        )}
      </p>
      <div className="sheet-trace-actions" style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
        <button
          type="button"
          disabled={!scope.active || !!disabledReason || busy}
          onClick={() => void trace('precedents')}
        >
          {words('Trace Precedents')}
        </button>
        <button
          type="button"
          disabled={!scope.active || !!disabledReason || busy}
          onClick={() => void trace('dependents')}
        >
          {words('Trace Dependents')}
        </button>
        <button
          type="button"
          disabled={!scope.active || !fresh || !snapshot.direction}
          onClick={() => {
            if (current()) {
              intent.current++;
              controller.clear(scope);
              setFeedback('');
            }
          }}
        >
          {words('Remove Arrows')}
        </button>
        {busy && (
          <button
            type="button"
            onClick={() => {
              if (current()) {
                intent.current++;
                controller.cancel(scope);
                setFeedback('');
              }
            }}
          >
            {words('Cancel trace')}
          </button>
        )}
      </div>
      <p role="status">{words(status)}</p>
      {fresh && snapshot.direction && (
        <p role="status">
          {snapshot.depth} {words('levels shown')} · {edges.length} {words('known arrows')} ·{' '}
          {words(snapshot.direction === 'precedents' ? 'Precedents' : 'Dependents')}
        </p>
      )}
      {fresh && snapshot.message && <p role="status">{words(snapshot.message)}</p>}
      {feedback && scope.active && <p role="alert">{words(feedback)}</p>}
      <ol aria-label={words('Known formula relationships')}>
        {edges.map((edge) => (
          <li key={edge.id} data-trace-edge={edge.id} style={{ paddingBlock: 6 }}>
            <span>
              {edge.from.sheet}!{edge.from.range} → {edge.to.sheet}!{edge.to.range}
            </span>
            <span>
              {edge.kind === 'spill-origin'
                ? words(' · Spill origin')
                : edge.kind === 'spill-range'
                  ? words(' · Spill range')
                  : ''}
            </span>
            <div>
              <button
                type="button"
                disabled={!scope.active}
                aria-label={words('Go to trace source') + ' ' + edge.from.sheet + '!' + edge.from.range}
                onClick={() => go(edge.from)}
              >
                {words('Go to source')}
              </button>
              <button
                type="button"
                disabled={!scope.active}
                aria-label={words('Go to trace formula') + ' ' + edge.to.sheet + '!' + edge.to.range}
                onClick={() => go(edge.to)}
              >
                {words('Go to formula')}
              </button>
            </div>
          </li>
        ))}
      </ol>
      {markers.length > 0 && (
        <div>
          <h4>{words('Coverage limits')}</h4>
          <ul aria-label={words('Trace coverage limits')}>
            {markers.map((marker) => (
              <li key={marker.id} data-trace-marker-kind={marker.kind}>
                <strong>
                  {marker.sheet}
                  {marker.cell ? '!' + marker.cell : ''}:{' '}
                </strong>
                {words(marker.reason)}
              </li>
            ))}
          </ul>
        </div>
      )}
    </section>
  );
}
