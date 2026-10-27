import { BULK_CELL_LIMIT } from './gridLimits';
import type { GridInternalClipboard, GridRanges, GridSetViewport } from './gridTypes';
import type { RangeWindow } from '../rangeCells';
import { type ViewGeometry, virtualWindow } from '../viewGeometry';
import type { ViewPreferences } from '../viewPreferences';
import type { WorkbookSheet } from '../../core/contracts';
import { colName } from '../sheetModel';
import { selectedBounds } from './rangeAddress';
import { useEffect, useLayoutEffect, useMemo } from 'react';

export interface GridWindowsInputs {
  activeRange: () => string;
  columnOffsets: Float64Array<ArrayBufferLike>;
  geometry: ViewGeometry;
  gridScroll: React.RefObject<HTMLDivElement>;
  internalClipboard: GridInternalClipboard;
  normalReturnPending: React.MutableRefObject<boolean>;
  normalViewportBookmark: React.MutableRefObject<{ key: string; left: number; top: number; zoom: number } | null>;
  normalViewportKey: string;
  ranges: GridRanges;
  rowOffsets: Float64Array<ArrayBufferLike>;
  selected: string;
  setViewport: GridSetViewport;
  sheet: WorkbookSheet;
  view: ViewPreferences;
  viewMode: 'normal' | 'pageLayout' | 'pageBreak';
  viewport: { width: number; height: number; scrollTop: number; scrollLeft: number };
}

export function useGridWindows({
  activeRange,
  columnOffsets,
  geometry,
  gridScroll,
  internalClipboard,
  normalReturnPending,
  normalViewportBookmark,
  normalViewportKey,
  ranges,
  rowOffsets,
  selected,
  setViewport,
  sheet,
  view,
  viewMode,
  viewport,
}: GridWindowsInputs) {
  useLayoutEffect(() => {
    if (viewMode !== 'normal' || !normalReturnPending.current) return;
    normalReturnPending.current = false;
    const saved = normalViewportBookmark.current;
    normalViewportBookmark.current = null;
    const scroll = gridScroll.current;
    if (!saved || saved.key !== normalViewportKey || !scroll) return;
    const scale = view.zoom / saved.zoom;
    scroll.scrollLeft = saved.left * scale;
    scroll.scrollTop = saved.top * scale;
    setViewport({
      width: scroll.clientWidth,
      height: scroll.clientHeight,
      scrollLeft: scroll.scrollLeft,
      scrollTop: scroll.scrollTop,
    });
  }, [viewMode, normalViewportKey, view.zoom, geometry]);

  const rowWindow = useMemo(
    () =>
      virtualWindow(rowOffsets, viewport.scrollTop, viewport.height, {
        headerExtent: geometry.columnHeaderHeight,
        frozenCount: geometry.frozenRows,
        overscan: 6,
      }),
    [geometry, viewport.scrollTop, viewport.height]
  );
  const columnWindow = useMemo(
    () =>
      virtualWindow(columnOffsets, viewport.scrollLeft, viewport.width, {
        headerExtent: geometry.rowHeaderWidth,
        frozenCount: geometry.frozenColumns,
        overscan: 3,
      }),
    [geometry, viewport.scrollLeft, viewport.width]
  );
  const firstRenderedRow = rowWindow.first,
    lastRenderedRow = rowWindow.last;
  const frozenRowCount = geometry.frozenRows;
  const renderedRows = useMemo(() => [...rowWindow.frozenIndices, ...rowWindow.indices], [rowWindow]);
  const firstRenderedColumn = columnWindow.first,
    lastRenderedColumn = columnWindow.last;
  const frozenColumnCount = geometry.frozenColumns;
  const renderedColumnIndices = useMemo(() => [...columnWindow.frozenIndices, ...columnWindow.indices], [columnWindow]);
  const frozenRowGap = rowWindow.indices.length
    ? 0
    : rowOffsets[frozenRowCount] - rowOffsets[(rowWindow.frozenIndices.at(-1) ?? -1) + 1];
  const frozenColumnGap = columnWindow.indices.length
    ? 0
    : columnOffsets[frozenColumnCount] - columnOffsets[(columnWindow.frozenIndices.at(-1) ?? -1) + 1];
  const windows: RangeWindow[] = [];
  if (sheet) {
    // Zero-height/hidden runs occupy no pixels and must not turn a viewport read
    // into a request for every hidden row between sparse visible overrides.
    const bands = (indices: number[]): Array<[number, number]> => {
      const result: Array<[number, number]> = [];
      for (const index of indices) {
        const previous = result.at(-1);
        if (previous && previous[1] + 1 === index) previous[1] = index;
        else result.push([index, index]);
      }
      return result;
    };
    const rowBands = bands(renderedRows),
      colBands = bands(renderedColumnIndices);
    for (const [first, last] of rowBands)
      for (const [left, right] of colBands) {
        if (last < first || right < left) continue;
        // At small zooms a real viewport can exceed the owner's per-read bound.
        // Split its bands rather than silently leaving visible cells unloaded.
        const step = Math.max(1, Math.floor(BULK_CELL_LIMIT / (right - left + 1)));
        for (let row = first; row <= last; row += step)
          windows.push({
            sheet: sheet.name,
            range: `${colName(left)}${row + 1}:${colName(right)}${Math.min(last + 1, row + step)}`,
          });
      }
    windows.push({ sheet: sheet.name, range: selected });
    try {
      const selection = activeRange();
      selectedBounds(selection, BULK_CELL_LIMIT);
      windows.push({ sheet: sheet.name, range: selection });
    } catch {
      /* A whole-sheet selection does not load the whole sheet. */
    }
    if (internalClipboard.current) windows.push(internalClipboard.current);
  }
  const rangeKey = ranges.request(windows);
  useEffect(() => ranges.load(rangeKey), [rangeKey]);
  return {
    rowWindow,
    columnWindow,
    firstRenderedRow,
    lastRenderedRow,
    frozenRowCount,
    renderedRows,
    firstRenderedColumn,
    lastRenderedColumn,
    frozenColumnCount,
    renderedColumnIndices,
    frozenRowGap,
    frozenColumnGap,
    rangeKey,
  } as const;
}
