/** Screen coordinates share physical XLSX dimensions with accepted print. */
import {
  createWorksheetDimensions,
  worksheetRowOffsets,
  worksheetColumnOffsets,
  type GeometryProfile,
  type WorksheetGeometrySource,
} from '../core/worksheet-geometry.mjs';
export interface ViewGeometrySheet extends WorksheetGeometrySource {
  rowHeights?: Array<{ row: number; height: number }>;
  hiddenRows?: number[];
  widths?: Array<{ min: number; max: number; width: number; hidden: boolean }>;
  freeze?: { rows: number; columns: number };
  merges?: string[];
}
export interface ViewCellPosition {
  row: number;
  col: number;
}
export interface ViewCellBounds {
  first: ViewCellPosition;
  last: ViewCellPosition;
}
export interface ViewRectangle {
  left: number;
  top: number;
  width: number;
  height: number;
}
export interface ViewViewport {
  width: number;
  height: number;
  scrollLeft: number;
  scrollTop: number;
}
export interface ViewGeometryOptions {
  rows: number;
  columns: number;
  zoom?: number;
  headings?: boolean;
  geometryProfile?: GeometryProfile;
  literal?: boolean;
}
export interface ViewGeometry {
  zoom: number;
  scale: number;
  rows: number;
  columns: number;
  headings: boolean;
  rowHeaderWidth: number;
  columnHeaderHeight: number;
  rowOffsets: Float64Array;
  columnOffsets: Float64Array;
  frozenRows: number;
  frozenColumns: number;
  frozenHeight: number;
  frozenWidth: number;
}
export const VIEW_MAX_ROWS = 1_048_576,
  VIEW_MAX_COLUMNS = 16_384;
function integer(value: number, minimum: number, maximum: number, label: string): number {
  if (!Number.isInteger(value) || value < minimum || value > maximum) throw Error('Invalid ' + label);
  return value;
}
function finite(value: number, label: string): number {
  if (!Number.isFinite(value) || value < 0) throw Error('Invalid ' + label);
  return value;
}
export function viewScale(zoom = 100): number {
  return integer(zoom, 10, 400, 'view zoom') / 100;
}
export function viewHeaders(zoom = 100, headings = true): { rowHeaderWidth: number; columnHeaderHeight: number } {
  if (typeof headings !== 'boolean') throw Error('Invalid heading visibility');
  const scale = viewScale(zoom);
  return { rowHeaderWidth: headings ? 48 * scale : 0, columnHeaderHeight: headings ? 32 * scale : 0 };
}
export interface ViewDimensionOptions {
  geometryProfile?: GeometryProfile;
  literal?: boolean;
}
export function rowHeight(
  sheet: ViewGeometrySheet | undefined,
  row: number,
  zoom = 100,
  options: ViewDimensionOptions = {}
): number {
  integer(row, 0, VIEW_MAX_ROWS - 1, 'row');
  const scale = viewScale(zoom);
  if (options.literal)
    return sheet?.hiddenRows?.includes(row)
      ? 0
      : finite(
          (((sheet?.rowHeights?.find((item) => item.row === row)?.height ?? 43.5) * 96) / 72) * scale,
          'row height'
        );
  return ((createWorksheetDimensions(sheet, options.geometryProfile).rowHeightPt(row) * 96) / 72) * scale;
}
export function columnWidth(
  sheet: ViewGeometrySheet | undefined,
  col: number,
  zoom = 100,
  options: ViewDimensionOptions = {}
): number {
  integer(col, 0, VIEW_MAX_COLUMNS - 1, 'column');
  const scale = viewScale(zoom);
  if (options.literal) {
    const item = sheet?.widths?.find((width) => col + 1 >= width.min && col + 1 <= width.max);
    return item?.hidden
      ? 0
      : finite((item ? Math.max(60, finite(item.width, 'column width') * 8) : 100) * scale, 'column width');
  }
  return createWorksheetDimensions(sheet, options.geometryProfile).columnPixels(col) * scale;
}
/** Prefix coordinates start at content origin. Sticky header extents are separate. */
export function rowOffsets(
  sheet: ViewGeometrySheet | undefined,
  count: number,
  zoom = 100,
  options: ViewDimensionOptions = {}
): Float64Array {
  integer(count, 0, VIEW_MAX_ROWS, 'row count');
  const scale = viewScale(zoom);
  if (!options.literal) {
    const offsets = worksheetRowOffsets(sheet, count, options.geometryProfile);
    for (let i = 1; i < offsets.length; i++) offsets[i] *= scale;
    return offsets;
  }
  const heights = new Map<number, number>(),
    hidden = new Set<number>();
  for (const item of sheet?.rowHeights ?? []) {
    integer(item.row, 0, VIEW_MAX_ROWS - 1, 'row height index');
    finite(item.height, 'row height');
    if (!heights.has(item.row)) heights.set(item.row, item.height);
  }
  for (const row of sheet?.hiddenRows ?? []) hidden.add(integer(row, 0, VIEW_MAX_ROWS - 1, 'hidden row'));
  const result = new Float64Array(count + 1);
  let unscaled = 0;
  for (let row = 0; row < count; row++) {
    unscaled += hidden.has(row) ? 0 : ((heights.get(row) ?? 43.5) * 96) / 72;
    result[row + 1] = finite(unscaled * scale, 'row offset');
  }
  return result;
}
export function columnOffsets(
  sheet: ViewGeometrySheet | undefined,
  count: number,
  zoom = 100,
  options: ViewDimensionOptions = {}
): Float64Array {
  integer(count, 0, VIEW_MAX_COLUMNS, 'column count');
  const scale = viewScale(zoom);
  if (!options.literal) {
    const offsets = worksheetColumnOffsets(sheet, count, options.geometryProfile);
    for (let i = 1; i < offsets.length; i++) offsets[i] *= scale;
    return offsets;
  }
  const sizes = new Float64Array(count).fill(100),
    next = new Uint32Array(count + 1);
  for (let index = 0; index <= count; index++) next[index] = index;
  const free = (start: number): number => {
    let index = start;
    while (next[index] !== index) index = next[index];
    while (next[start] !== start) {
      const old = next[start];
      next[start] = index;
      start = old;
    }
    return index;
  };
  for (const item of sheet?.widths ?? []) {
    integer(item.min, 1, VIEW_MAX_COLUMNS, 'column interval');
    integer(item.max, item.min, VIEW_MAX_COLUMNS, 'column interval');
    finite(item.width, 'column width');
    if (typeof item.hidden !== 'boolean') throw Error('Invalid column visibility');
    const end = Math.min(count, item.max),
      width = item.hidden ? 0 : Math.max(60, item.width * 8);
    for (let index = free(Math.min(count, item.min - 1)); index < end; index = free(index)) {
      sizes[index] = width;
      next[index] = free(index + 1);
    }
  }
  const result = new Float64Array(count + 1);
  let unscaled = 0;
  for (let col = 0; col < count; col++) {
    unscaled += sizes[col];
    result[col + 1] = finite(unscaled * scale, 'column offset');
  }
  return result;
}
export function offsetForRow(
  sheet: ViewGeometrySheet | undefined,
  row: number,
  zoom = 100,
  options: ViewDimensionOptions = {}
): number {
  integer(row, 0, VIEW_MAX_ROWS, 'row offset index');
  return options.literal
    ? rowOffsets(sheet, row, zoom, options)[row]
    : ((createWorksheetDimensions(sheet, options.geometryProfile).rowDistancePt(0, row) * 96) / 72) * viewScale(zoom);
}
export function offsetForColumn(
  sheet: ViewGeometrySheet | undefined,
  col: number,
  zoom = 100,
  options: ViewDimensionOptions = {}
): number {
  integer(col, 0, VIEW_MAX_COLUMNS, 'column offset index');
  return columnOffsets(sheet, col, zoom, options)[col];
}
export function createViewGeometry(sheet: ViewGeometrySheet | undefined, options: ViewGeometryOptions): ViewGeometry {
  const zoom = options.zoom ?? 100,
    scale = viewScale(zoom),
    headings = options.headings ?? true,
    headers = viewHeaders(zoom, headings);
  const rows = integer(options.rows, 0, VIEW_MAX_ROWS, 'row count'),
    columns = integer(options.columns, 0, VIEW_MAX_COLUMNS, 'column count'),
    r = rowOffsets(sheet, rows, zoom, options),
    c = columnOffsets(sheet, columns, zoom, options);
  const frozenRows = Math.min(rows, integer(sheet?.freeze?.rows ?? 0, 0, VIEW_MAX_ROWS, 'frozen rows')),
    frozenColumns = Math.min(columns, integer(sheet?.freeze?.columns ?? 0, 0, VIEW_MAX_COLUMNS, 'frozen columns'));
  return {
    zoom,
    scale,
    rows,
    columns,
    headings,
    ...headers,
    rowOffsets: r,
    columnOffsets: c,
    frozenRows,
    frozenColumns,
    frozenHeight: r[frozenRows],
    frozenWidth: c[frozenColumns],
  };
}
function upperBound(offsets: Float64Array, value: number): number {
  let first = 0,
    last = offsets.length;
  while (first < last) {
    const middle = (first + last) >>> 1;
    if (offsets[middle] <= value) first = middle + 1;
    else last = middle;
  }
  return first;
}
function lowerBound(offsets: Float64Array, value: number): number {
  let first = 0,
    last = offsets.length;
  while (first < last) {
    const middle = (first + last) >>> 1;
    if (offsets[middle] < value) first = middle + 1;
    else last = middle;
  }
  return first;
}
function axisIndex(offsets: Float64Array, value: number): number | null {
  if (value < 0 || value >= offsets[offsets.length - 1]) return null;
  const index = upperBound(offsets, value) - 1;
  return index >= 0 && index + 1 < offsets.length && offsets[index + 1] > offsets[index] ? index : null;
}
function visibleIndices(offsets: Float64Array, first: number, last: number): number[] {
  const values: number[] = [];
  for (let index = first; index < last; ) {
    if (offsets[index + 1] > offsets[index]) {
      values.push(index);
      index++;
    } else index = Math.max(index + 1, upperBound(offsets, offsets[index]) - 1);
  }
  return values;
}
export interface ViewVirtualWindow {
  first: number;
  last: number;
  indices: number[];
  frozenIndices: number[];
}
/** last is exclusive. Input prefixes must come from the validated builders above. */
export function virtualWindow(
  offsets: Float64Array,
  scroll: number,
  viewport: number,
  options: { headerExtent?: number; frozenCount?: number; overscan?: number } = {}
): ViewVirtualWindow {
  const count = offsets.length - 1;
  if (count < 0 || offsets[0] !== 0) throw Error('Invalid view offsets');
  finite(scroll, 'scroll');
  finite(viewport, 'viewport');
  const header = finite(options.headerExtent ?? 0, 'header extent'),
    frozen = integer(options.frozenCount ?? 0, 0, count, 'frozen count'),
    overscan = integer(options.overscan ?? 2, 0, 100, 'overscan');
  const available = Math.max(0, viewport - header),
    frozenExtent = offsets[frozen];
  const frozenEnd = Math.min(frozen, lowerBound(offsets, available)),
    frozenIndices = visibleIndices(offsets, 0, frozenEnd);
  if (available <= frozenExtent || count === 0) return { first: frozen, last: frozen, indices: [], frozenIndices };
  const begin = Math.min(count, Math.max(frozen, upperBound(offsets, scroll + frozenExtent) - 1)),
    end = Math.min(count, Math.max(begin, lowerBound(offsets, scroll + available)));
  const first = Math.max(frozen, begin - overscan),
    last = Math.min(count, end + overscan);
  return { first, last, indices: visibleIndices(offsets, first, last), frozenIndices };
}
function validateBounds(g: ViewGeometry, bounds: ViewCellBounds): void {
  integer(bounds.first.row, 0, g.rows - 1, 'first row');
  integer(bounds.last.row, bounds.first.row, g.rows - 1, 'last row');
  integer(bounds.first.col, 0, g.columns - 1, 'first column');
  integer(bounds.last.col, bounds.first.col, g.columns - 1, 'last column');
}
export function cellRectangle(g: ViewGeometry, bounds: ViewCellBounds): ViewRectangle {
  validateBounds(g, bounds);
  const left = g.columnOffsets[bounds.first.col],
    top = g.rowOffsets[bounds.first.row];
  return {
    left,
    top,
    width: g.columnOffsets[bounds.last.col + 1] - left,
    height: g.rowOffsets[bounds.last.row + 1] - top,
  };
}
function position(reference: string): ViewCellPosition {
  const match = /^\$?([A-Z]{1,3})\$?([1-9][0-9]{0,6})$/i.exec(reference);
  if (!match) throw Error('Invalid merged reference');
  let col = 0;
  for (const character of match[1].toUpperCase()) col = col * 26 + character.charCodeAt(0) - 64;
  return {
    row: integer(Number(match[2]) - 1, 0, VIEW_MAX_ROWS - 1, 'merged row'),
    col: integer(col - 1, 0, VIEW_MAX_COLUMNS - 1, 'merged column'),
  };
}
export function mergedCellBounds(sheet: ViewGeometrySheet | undefined, row: number, col: number): ViewCellBounds {
  integer(row, 0, VIEW_MAX_ROWS - 1, 'row');
  integer(col, 0, VIEW_MAX_COLUMNS - 1, 'column');
  for (const range of sheet?.merges ?? []) {
    const parts = range.split(':');
    if (parts.length > 2) throw Error('Invalid merged range');
    const first = position(parts[0]),
      last = position(parts[1] ?? parts[0]);
    if (last.row < first.row || last.col < first.col) throw Error('Invalid merged range');
    if (row >= first.row && row <= last.row && col >= first.col && col <= last.col) return { first, last };
  }
  return { first: { row, col }, last: { row, col } };
}
export function hitTestCell(
  g: ViewGeometry,
  x: number,
  y: number,
  scroll: { scrollLeft: number; scrollTop: number },
  sheet?: ViewGeometrySheet
): ViewCellPosition | null {
  if (!Number.isFinite(x) || !Number.isFinite(y)) throw Error('Invalid pointer position');
  finite(scroll.scrollLeft, 'scroll');
  finite(scroll.scrollTop, 'scroll');
  const left = x - g.rowHeaderWidth,
    top = y - g.columnHeaderHeight;
  if (left < 0 || top < 0) return null;
  const col = axisIndex(g.columnOffsets, left + (left < g.frozenWidth ? 0 : scroll.scrollLeft)),
    row = axisIndex(g.rowOffsets, top + (top < g.frozenHeight ? 0 : scroll.scrollTop));
  if (row === null || col === null) return null;
  return mergedCellBounds(sheet, row, col).first;
}
function viewport(v: ViewViewport): void {
  finite(v.width, 'viewport width');
  finite(v.height, 'viewport height');
  finite(v.scrollLeft, 'scroll');
  finite(v.scrollTop, 'scroll');
}
export function revealCellScroll(
  g: ViewGeometry,
  v: ViewViewport,
  row: number,
  col: number,
  sheet?: ViewGeometrySheet,
  padding = 8
): { scrollLeft: number; scrollTop: number } {
  viewport(v);
  finite(padding, 'reveal padding');
  const bounds = mergedCellBounds(sheet, row, col),
    rectangle = cellRectangle(g, bounds);
  const reveal = (
    start: number,
    end: number,
    first: number,
    last: number,
    frozenCount: number,
    frozenExtent: number,
    extent: number,
    current: number
  ): number => {
    if (last < frozenCount || extent <= frozenExtent) return current;
    const targetStart = Math.max(start, frozenExtent),
      remaining = extent - frozenExtent;
    if (end - targetStart >= remaining) return Math.max(0, targetStart - frozenExtent);
    if (targetStart < current + frozenExtent) return Math.max(0, targetStart - frozenExtent - padding);
    if (end > current + extent) return Math.max(0, end - extent + padding);
    return current;
  };
  const scrollLeft = reveal(
    rectangle.left,
    rectangle.left + rectangle.width,
    bounds.first.col,
    bounds.last.col,
    g.frozenColumns,
    g.frozenWidth,
    Math.max(0, v.width - g.rowHeaderWidth),
    v.scrollLeft
  );
  const scrollTop = reveal(
    rectangle.top,
    rectangle.top + rectangle.height,
    bounds.first.row,
    bounds.last.row,
    g.frozenRows,
    g.frozenHeight,
    Math.max(0, v.height - g.columnHeaderHeight),
    v.scrollTop
  );
  return {
    scrollLeft: Math.min(scrollLeft, Math.max(0, g.columnOffsets[g.columns] - Math.max(0, v.width - g.rowHeaderWidth))),
    scrollTop: Math.min(scrollTop, Math.max(0, g.rowOffsets[g.rows] - Math.max(0, v.height - g.columnHeaderHeight))),
  };
}
/** Fit includes visible headers and the complete frozen strips, without counting
 * their overlap with the selected cells twice. Returned zoom is an integer. */
export function zoomToSelection(
  g: ViewGeometry,
  bounds: ViewCellBounds,
  v: Pick<ViewViewport, 'width' | 'height'>
): number {
  finite(v.width, 'viewport width');
  finite(v.height, 'viewport height');
  const rectangle = cellRectangle(g, bounds);
  if (!v.width || !v.height || !rectangle.width || !rectangle.height) return g.zoom;
  const horizontal =
    (g.rowHeaderWidth +
      g.frozenWidth +
      Math.max(0, rectangle.left + rectangle.width - Math.max(rectangle.left, g.frozenWidth))) /
    g.scale;
  const vertical =
    (g.columnHeaderHeight +
      g.frozenHeight +
      Math.max(0, rectangle.top + rectangle.height - Math.max(rectangle.top, g.frozenHeight))) /
    g.scale;
  return Math.max(10, Math.min(400, Math.floor(Math.min(v.width / horizontal, v.height / vertical) * 100)));
}
