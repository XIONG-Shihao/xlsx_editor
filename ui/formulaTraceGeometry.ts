import {
  cellRectangle,
  mergedCellBounds,
  type ViewGeometry,
  type ViewGeometrySheet,
  type ViewViewport,
  type ViewRectangle,
} from './viewGeometry';
import { traceBounds, type TraceEndpoint, type TraceEdgeRow } from './formulaTrace';
export interface TracePoint {
  x: number;
  y: number;
}
export interface TraceEndpointPosition {
  endpoint: TraceEndpoint;
  point: TracePoint;
  visible: boolean;
  kind: 'cell' | 'offscreen' | 'cross-sheet' | 'hidden' | 'outside-grid' | 'split-merge';
  rectangle?: ViewRectangle;
}
export interface TraceArrowGeometry {
  id: string;
  edge: TraceEdgeRow;
  from: TraceEndpointPosition;
  to: TraceEndpointPosition;
  draw: boolean;
  path: string;
}
function clamp(value: number, first: number, last: number): number {
  return Math.min(Math.max(value, first), Math.max(first, last));
}
/** Coordinates are local to the scroller's visible viewport, never table extent.
 * Nonfrozen cells are clipped behind sticky strips; headers are not cell space. */
export function traceEndpointPosition(
  endpoint: TraceEndpoint,
  sheetName: string,
  g: ViewGeometry,
  v: ViewViewport,
  sheet?: ViewGeometrySheet
): TraceEndpointPosition {
  if (![v.width, v.height, v.scrollLeft, v.scrollTop].every((n) => Number.isFinite(n) && n >= 0))
    throw Error('Invalid trace viewport.');
  const left = g.rowHeaderWidth,
    top = g.columnHeaderHeight,
    right = Math.max(left, v.width),
    bottom = Math.max(top, v.height),
    marker = { x: clamp(left + 16, left, right), y: clamp(top + 16, top, bottom) };
  if (endpoint.sheet !== sheetName) return { endpoint, point: marker, visible: false, kind: 'cross-sheet' };
  const b = traceBounds(endpoint.range);
  if (b.first.row >= g.rows || b.first.col >= g.columns || b.last.row >= g.rows || b.last.col >= g.columns)
    return { endpoint, point: marker, visible: false, kind: 'outside-grid' };
  const merged =
    b.first.row === b.last.row && b.first.col === b.last.col ? mergedCellBounds(sheet, b.first.row, b.first.col) : b;
  const r = cellRectangle(g, merged);
  const frozenX = merged.last.col < g.frozenColumns,
    frozenY = merged.last.row < g.frozenRows;
  const x = left + r.left - (frozenX ? 0 : v.scrollLeft),
    y = top + r.top - (frozenY ? 0 : v.scrollTop);
  const clipLeft = frozenX ? left : left + g.frozenWidth,
    clipTop = frozenY ? top : top + g.frozenHeight;
  const clipRight = frozenX ? Math.min(right, left + g.frozenWidth) : right,
    clipBottom = frozenY ? Math.min(bottom, top + g.frozenHeight) : bottom;
  const rectangle = { left: x, top: y, width: r.width, height: r.height };
  const point = {
    x: clamp(x + r.width / 2, Math.min(right - 8, clipLeft + 8), right - 8),
    y: clamp(y + r.height / 2, Math.min(bottom - 8, clipTop + 8), bottom - 8),
  };
  if (!r.width || !r.height) return { endpoint, point, visible: false, kind: 'hidden', rectangle };
  if (
    (merged.first.col < g.frozenColumns && merged.last.col >= g.frozenColumns) ||
    (merged.first.row < g.frozenRows && merged.last.row >= g.frozenRows)
  )
    return { endpoint, point, visible: false, kind: 'split-merge', rectangle };
  const x0 = Math.max(x, clipLeft),
    x1 = Math.min(x + r.width, clipRight),
    y0 = Math.max(y, clipTop),
    y1 = Math.min(y + r.height, clipBottom);
  if (x1 <= x0 || y1 <= y0) return { endpoint, point, visible: false, kind: 'offscreen', rectangle };
  return {
    endpoint,
    point: { x: (x0 + x1) / 2, y: (y0 + y1) / 2 },
    visible: true,
    kind: 'cell',
    rectangle: { left: x0, top: y0, width: x1 - x0, height: y1 - y0 },
  };
}
export function formulaTraceArrows(
  edges: readonly TraceEdgeRow[],
  sheetName: string,
  g: ViewGeometry,
  v: ViewViewport,
  sheet?: ViewGeometrySheet
): TraceArrowGeometry[] {
  if (edges.length > 1024) throw Error('Too many trace arrows.');
  return edges.map((edge) => {
    const from = traceEndpointPosition(edge.from, sheetName, g, v, sheet),
      to = traceEndpointPosition(edge.to, sheetName, g, v, sheet);
    // Foreign-sheet endpoints have no invented grid coordinate. Attach the marker
    // near the known on-sheet end; its list action performs actual navigation.
    if (from.kind === 'cross-sheet' && to.visible)
      from.point = {
        x: Math.max(g.rowHeaderWidth + 10, to.point.x - 44),
        y: Math.max(g.columnHeaderHeight + 10, to.point.y - 28),
      };
    if (to.kind === 'cross-sheet' && from.visible)
      to.point = { x: Math.min(v.width - 10, from.point.x + 44), y: Math.min(v.height - 10, from.point.y + 28) };
    const draw = (from.visible || to.visible) && v.width > g.rowHeaderWidth && v.height > g.columnHeaderHeight;
    const path =
      Math.abs(from.point.x - to.point.x) + Math.abs(from.point.y - to.point.y) < 1 && from.visible && to.visible
        ? `M ${from.point.x} ${from.point.y} C ${from.point.x + 24} ${from.point.y - 24} ${from.point.x + 24} ${from.point.y + 24} ${from.point.x} ${from.point.y}`
        : `M ${from.point.x} ${from.point.y} L ${to.point.x} ${to.point.y}`;
    return { id: edge.id, edge, from, to, draw, path };
  });
}
