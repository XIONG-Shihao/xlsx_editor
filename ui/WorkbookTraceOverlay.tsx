import { useId, useRef, useSyncExternalStore } from 'react';
import {
  formulaTraceScopeKey,
  type FormulaTraceController,
  type FormulaTraceScope,
  type TraceEndpoint,
} from './formulaTrace';
import { formulaTraceArrows } from './formulaTraceGeometry';
import type { ViewGeometry, ViewGeometrySheet, ViewViewport } from './viewGeometry';
export interface WorkbookTraceOverlayProps {
  controller: FormulaTraceController;
  scope: FormulaTraceScope;
  geometry: ViewGeometry;
  viewport: ViewViewport;
  sheet?: ViewGeometrySheet;
  onNavigate: (target: TraceEndpoint) => void;
  onOpenList?: () => void;
  words?: (text: string) => string;
}
const identity = (text: string) => text;
export function WorkbookTraceOverlay({
  controller,
  scope,
  geometry,
  viewport,
  sheet,
  onNavigate,
  onOpenList,
  words = identity,
}: WorkbookTraceOverlayProps): JSX.Element | null {
  const snapshot = useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot),
    key = formulaTraceScopeKey(scope),
    id = useId().replace(/:/g, '');
  const live = useRef({ key, scope, controller });
  live.current = { key, scope, controller };
  if (!scope.active || !snapshot.scope || formulaTraceScopeKey(snapshot.scope) !== key || !snapshot.edges.length)
    return null;
  const arrows = formulaTraceArrows(snapshot.edges, scope.sheet, geometry, viewport, sheet);
  const navigation = new Map<string, { target: TraceEndpoint; x: number; y: number; kind: string }>();
  for (const arrow of arrows)
    if (arrow.draw)
      for (const endpoint of [arrow.from, arrow.to])
        if (!endpoint.visible) {
          const k = JSON.stringify(endpoint.endpoint);
          if (!navigation.has(k))
            navigation.set(k, {
              target: endpoint.endpoint,
              x: endpoint.point.x,
              y: endpoint.point.y,
              kind: endpoint.kind,
            });
        }
  const current = () =>
    live.current.key === key &&
    live.current.scope.active &&
    live.current.controller === controller &&
    controller.getSnapshot() === snapshot;
  const grouped = new Map<string, Array<{ target: TraceEndpoint; x: number; y: number; kind: string }>>();
  for (const marker of navigation.values()) {
    const position = Math.round(marker.x / 20) + ':' + Math.round(marker.y / 20),
      items = grouped.get(position) ?? [];
    items.push(marker);
    grouped.set(position, items);
  }
  const markers = [...grouped.values()].slice(0, 24),
    offscreenOnly = arrows.filter((a) => !a.draw).length;
  return (
    <div
      className="sheet-trace-overlay"
      role="group"
      aria-label={words('Formula trace arrows')}
      data-trace-overlay-state={scope.state}
      data-trace-root={scope.sheet + '!' + scope.cell}
      style={{ position: 'absolute', inset: 0, pointerEvents: 'none', overflow: 'hidden', zIndex: 6 }}
    >
      <svg aria-hidden="true" width={viewport.width} height={viewport.height} style={{ display: 'block' }}>
        <defs>
          <marker
            id={id + '-arrow'}
            viewBox="0 0 10 10"
            refX="9"
            refY="5"
            markerWidth="6"
            markerHeight="6"
            orient="auto-start-reverse"
          >
            <path d="M 0 0 L 10 5 L 0 10 z" fill="#1765cc" />
          </marker>
          <clipPath id={id + '-clip'}>
            <rect
              x={geometry.rowHeaderWidth}
              y={geometry.columnHeaderHeight}
              width={Math.max(0, viewport.width - geometry.rowHeaderWidth)}
              height={Math.max(0, viewport.height - geometry.columnHeaderHeight)}
            />
          </clipPath>
        </defs>
        <g clipPath={'url(#' + id + '-clip)'}>
          {arrows
            .filter((a) => a.draw)
            .map((arrow) => (
              <g
                key={arrow.id}
                data-trace-arrow={arrow.id}
                data-trace-from={arrow.edge.from.sheet + '!' + arrow.edge.from.range}
                data-trace-to={arrow.edge.to.sheet + '!' + arrow.edge.to.range}
              >
                <path
                  d={arrow.path}
                  fill="none"
                  stroke="#1765cc"
                  strokeWidth="1.8"
                  strokeDasharray={arrow.from.visible && arrow.to.visible ? undefined : '4 3'}
                  markerEnd={'url(#' + id + '-arrow)'}
                />
                <circle cx={arrow.from.point.x} cy={arrow.from.point.y} r="2.5" fill="#1765cc" />
              </g>
            ))}
        </g>
      </svg>
      {markers.map((items) => {
        const marker = items[0],
          multiple = items.length > 1,
          label = multiple
            ? words('Open trace relationship list')
            : words('Go to trace') + ' ' + marker.target.sheet + '!' + marker.target.range;
        return (
          <button
            key={marker.target.sheet + '!' + marker.target.range}
            type="button"
            data-trace-boundary-kind={multiple ? 'multiple' : marker.kind}
            data-trace-boundary-count={items.length}
            disabled={multiple && !onOpenList}
            aria-label={label}
            title={
              multiple
                ? items.length + ' ' + words('related cells; choose one in the list')
                : marker.target.sheet +
                  '!' +
                  marker.target.range +
                  ' · ' +
                  words(
                    marker.kind === 'cross-sheet'
                      ? 'Other worksheet'
                      : marker.kind === 'hidden'
                        ? 'Hidden cells'
                        : marker.kind === 'split-merge'
                          ? 'Range spans frozen panes'
                          : 'Outside the visible cells'
                  )
            }
            style={{
              position: 'absolute',
              left: Math.max(geometry.rowHeaderWidth, marker.x - 8),
              top: Math.max(geometry.columnHeaderHeight, marker.y - 8),
              width: 20,
              height: 20,
              padding: 0,
              border: '1px solid #1765cc',
              borderRadius: 3,
              background: '#fff',
              color: '#1765cc',
              pointerEvents: 'auto',
              fontSize: 12,
            }}
            onClick={() => {
              if (current()) {
                if (multiple) onOpenList?.();
                else onNavigate(marker.target);
              }
            }}
          >
            {multiple ? items.length : marker.kind === 'cross-sheet' ? '▦' : '↗'}
          </button>
        );
      })}
      {(offscreenOnly > 0 || grouped.size > markers.length) && onOpenList && (
        <button
          type="button"
          aria-label={words('Open trace relationship list')}
          style={{ position: 'absolute', right: 8, top: geometry.columnHeaderHeight + 8, pointerEvents: 'auto' }}
          onClick={() => {
            if (current()) onOpenList();
          }}
        >
          {words('Related cells outside view')}
        </button>
      )}
    </div>
  );
}
