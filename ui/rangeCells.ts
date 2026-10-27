import { useMemo, useRef, useState, useEffect } from 'react';
import type { WorkbookCell, WorkbookSheet } from '../core/contracts';
import { addressPosition, colName } from './sheetModel';
import { diagnosticPayload, registerWorkbookDiagnostics } from './workbookDiagnostics';
import { RangeSummaryCache } from './rangeSummaryCache';
export type ReadRange = (sheet: string, range: string) => Promise<Record<string, WorkbookCell>>;
export interface RangeWindow {
  sheet: string;
  range: string;
}
/** A viewport/selection cache, independent of workbook history. Empty cells are
 * covered by loaded tiles too; absence of an object does not mean unloaded. */
export function useRangeCells(metadata: WorkbookSheet[], read: ReadRange | undefined, revision?: string) {
  const reader = useRef(read);
  reader.current = read;
  const [cache, setCache] = useState<{
    revision?: string;
    key: string;
    cells: Map<string, Record<string, WorkbookCell>>;
    duplicates?: Map<string, Record<string, string[]>>;
  }>({ key: '', cells: new Map() });
  const [requested, setRequested] = useState<{ key: string; revision?: string; tiles: RangeWindow[] }>({
    key: '',
    tiles: [],
  });
  const [error, setError] = useState('');
  const retained = useRef(new Map<string, Record<string, WorkbookCell>>());
  const duplicateCache = useRef(new RangeSummaryCache());
  const inFlight = useRef<Map<string, Record<string, WorkbookCell>> | undefined>(undefined);
  const published = useRef(cache);
  published.current = cache;
  const requestedCount = useRef(0);
  requestedCount.current = requested.tiles.length;
  const source = useRef(metadata);
  source.current = metadata;
  const owner = useRef(revision);
  if (owner.current !== revision) {
    owner.current = revision;
    retained.current = new Map();
    duplicateCache.current = new RangeSummaryCache();
  }
  useEffect(
    () =>
      registerWorkbookDiagnostics('viewport', () => ({
        retainedTiles: retained.current.size,
        publishedTiles: published.current.cells.size,
        requestedTiles: requestedCount.current,
        inFlightTiles: inFlight.current?.size ?? 0,
        ...duplicateCache.current.diagnostics(),
        ...diagnosticPayload([
          retained.current,
          published.current.cells,
          inFlight.current,
          duplicateCache.current.diagnosticEntries(),
          published.current.duplicates,
        ]),
      })),
    []
  );
  const request = (windows: RangeWindow[]): string => {
    if (!read) return '';
    const tiles = new Map<string, RangeWindow>();
    for (const window of windows) {
      const [start, end = start] = window.range.split(':'),
        a = addressPosition(start),
        b = addressPosition(end);
      if ((b.row - a.row + 1) * (b.col - a.col + 1) > 10000) continue; // Existing UI bulk operations validate separately.
      for (let row = Math.floor(a.row / 32) * 32; row <= b.row; row += 32)
        for (let col = Math.floor(a.col / 32) * 32; col <= b.col; col += 32) {
          const range = `${colName(col)}${row + 1}:${colName(Math.min(16383, col + 31))}${Math.min(1048576, row + 32)}`;
          tiles.set(`${window.sheet}\0${range}`, { sheet: window.sheet, range });
        }
    }
    return JSON.stringify({ revision, tiles: [...tiles].sort(([a], [b]) => a.localeCompare(b)) });
  };
  const load = (key: string): void => {
    if (!key) return;
    const data = JSON.parse(key) as { revision?: string; tiles: Array<[string, RangeWindow]> };
    setRequested((previous) =>
      previous.key === key ? previous : { key, revision: data.revision, tiles: data.tiles.map(([, tile]) => tile) }
    );
  };
  useEffect(() => {
    // A revision render precedes the grid's effect that submits its new key.
    // Do not start reads for the old key using the new revision's reader.
    if (!read || !requested.key || requested.revision !== revision) return;
    let active = true;
    const current = revision;
    const next = new Map<string, Record<string, WorkbookCell>>();
    inFlight.current = next;
    void (async () => {
      try {
        for (const tile of requested.tiles) {
          const key = `${tile.sheet}\0${tile.range}`;
          const cells = retained.current.get(key) ?? (await reader.current!(tile.sheet, tile.range));
          if (!active || owner.current !== current) return;
          next.set(key, cells);
        }
        const duplicates = new Map<string, Record<string, string[]>>();
        for (const name of new Set(requested.tiles.map((tile) => tile.sheet))) {
          const summaries: Record<string, string[]> = {};
          const visible = requested.tiles
            .filter((tile) => tile.sheet === name)
            .map((tile) => {
              const [first, last = first] = tile.range.split(':');
              return { first: addressPosition(first), last: addressPosition(last) };
            });
          for (const rule of source.current.find((sheet) => sheet.name === name)?.conditionalRules ?? []) {
            if (rule.rule !== 'duplicateValues' || summaries[rule.range]) continue;
            const [first, last = first] = rule.range.split(':'),
              a = addressPosition(first),
              b = addressPosition(last);
            if ((b.row - a.row + 1) * (b.col - a.col + 1) > 10000) continue;
            // A visible cell needs duplicate counts for its entire rule range.
            // Offscreen rules can wait until their tiles are actually requested.
            if (
              !visible.some(
                (tile) =>
                  a.row <= tile.last.row && b.row >= tile.first.row && a.col <= tile.last.col && b.col >= tile.first.col
              )
            )
              continue;
            const cached = duplicateCache.current.get(name, rule.range);
            if (cached) {
              summaries[rule.range] = cached;
              continue;
            }
            const values = await reader.current!(name, rule.range),
              counts = new Map<string, number>();
            if (!active || owner.current !== current) return;
            for (const cell of Object.values(values)) {
              const value = cell.formula
                ? cell.calculation?.status === 'calculated'
                  ? cell.calculation.value
                  : null
                : cell.value;
              if (value !== undefined && value !== null && value !== '') {
                const key = JSON.stringify(value);
                counts.set(key, (counts.get(key) ?? 0) + 1);
              }
            }
            summaries[rule.range] = [...counts].filter(([, count]) => count > 1).map(([key]) => key);
            duplicateCache.current.set(name, rule.range, summaries[rule.range]);
          }
          duplicates.set(name, summaries);
        }
        retained.current = next;
        setCache({ revision: current, key: requested.key, cells: next, duplicates });
        setError('');
      } catch (cause) {
        if (active) setError(String(cause));
      } finally {
        if (inFlight.current === next) inFlight.current = undefined;
      }
    })();
    return () => {
      active = false;
      if (inFlight.current === next) inFlight.current = undefined;
    };
  }, [requested, revision, !!read]);
  const sheets = useMemo(() => {
    if (!read) return metadata;
    return metadata.map((sheet) => {
      const cells: Record<string, WorkbookCell> = {};
      if (cache.revision === revision)
        for (const [key, values] of cache.cells) if (key.startsWith(`${sheet.name}\0`)) Object.assign(cells, values);
      return {
        ...sheet,
        cells,
        duplicateValues: cache.revision === revision ? cache.duplicates?.get(sheet.name) : undefined,
      };
    });
  }, [metadata, cache, revision, !!read]);
  return {
    sheets,
    request,
    load,
    error,
    ready: (key: string) => !read || (cache.revision === revision && cache.key === key),
  };
}

/** Read a single row/column in pages for Ctrl/Cmd-arrow navigation. */
export async function indexedBoundary(
  read: ReadRange,
  sheet: WorkbookSheet,
  position: { row: number; col: number },
  vertical: boolean,
  step: number
): Promise<number> {
  const limit = vertical ? Math.max(29, sheet.rows - 1) : Math.max(11, sheet.columns - 1);
  let page = -1,
    cells: Record<string, WorkbookCell> = {};
  const filled = async (index: number): Promise<boolean> => {
    const next = Math.floor(index / 1024);
    if (next !== page) {
      const a = next * 1024,
        b = Math.min(limit, a + 1023);
      cells = await read(
        sheet.name,
        vertical
          ? `${colName(position.col)}${a + 1}:${colName(position.col)}${b + 1}`
          : `${colName(a)}${position.row + 1}:${colName(b)}${position.row + 1}`
      );
      page = next;
    }
    const cell = cells[vertical ? `${colName(position.col)}${index + 1}` : `${colName(index)}${position.row + 1}`];
    return (
      !!(cell?.formula || cell?.formulaType) || (cell?.value !== undefined && cell.value !== null && cell.value !== '')
    );
  };
  let index = vertical ? position.row : position.col;
  const next = index + step;
  if (next < 0 || next > limit) return index;
  const contiguous = (await filled(index)) && (await filled(next));
  while (index + step >= 0 && index + step <= limit) {
    const occupied = await filled(index + step);
    if (contiguous && !occupied) break;
    index += step;
    if (!contiguous && occupied) break;
  }
  return index;
}
