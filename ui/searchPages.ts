import type { WorkbookCell } from '../core/contracts';

export type SearchMode = 'display' | 'raw' | 'formula';
export type SearchOptions = { range?: string; matchLimit?: number };
export type SearchQuery = (
  sheet: string,
  text: string,
  mode: SearchMode,
  after?: string,
  options?: SearchOptions
) => Promise<{
  cells: Record<string, WorkbookCell>;
  matches: number;
  nextCursor: string | null;
}>;
export type SearchCursor = { sheet: number; after?: string };
export type SearchMatch = { sheet: string; cell: string; value: WorkbookCell };
export type SearchPage = { matches: SearchMatch[]; next: SearchCursor | null };
export const SEARCH_PAGE_LIMIT = 1000;

function position(ref: string): number {
  const parsed = /^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(ref);
  if (!parsed) throw Error('Invalid search cell cursor.');
  let column = 0;
  for (const char of parsed[1]) column = column * 26 + char.charCodeAt(0) - 64;
  const row = Number(parsed[2]);
  if (column > 16384 || row > 1048576) throw Error('Search cell cursor is outside the worksheet.');
  return (row - 1) * 16384 + column - 1;
}

/** A page keeps every returned match, never just a subset of a counted batch. */
export async function loadSearchPage(options: {
  sheets: string[];
  text: string;
  mode: SearchMode;
  query: SearchQuery;
  cursor?: SearchCursor;
  cancelled: () => boolean;
  yieldWork?: () => Promise<void>;
}): Promise<SearchPage | null> {
  let sheet = options.cursor?.sheet ?? 0,
    after = options.cursor?.after;
  if (!Number.isInteger(sheet) || sheet < 0 || sheet > options.sheets.length)
    throw Error('Invalid search sheet cursor.');
  const matches: SearchMatch[] = [];
  while (sheet < options.sheets.length) {
    if (options.cancelled()) return null;
    const limit = SEARCH_PAGE_LIMIT - matches.length;
    const reply = await options.query(options.sheets[sheet], options.text, options.mode, after, { matchLimit: limit });
    if (options.cancelled()) return null;
    const cells = Object.entries(reply.cells);
    if (cells.length > limit || reply.matches !== cells.length)
      throw Error('Search owner omitted matches or exceeded the requested page limit.');
    let previous = after ? position(after) : -1;
    for (const [cell, value] of cells) {
      const current = position(cell);
      if (current <= previous) throw Error('Search owner returned duplicate or unordered cells.');
      previous = current;
      matches.push({ sheet: options.sheets[sheet], cell, value });
    }
    if (reply.nextCursor !== null) {
      const next = position(reply.nextCursor);
      if (next <= (after ? position(after) : -1) || next < previous)
        throw Error('Search owner returned a non-advancing cursor.');
      after = reply.nextCursor;
    } else {
      sheet++;
      after = undefined;
    }
    if (matches.length === SEARCH_PAGE_LIMIT)
      return { matches, next: sheet < options.sheets.length ? { sheet, after } : null };
    if (options.yieldWork) await options.yieldWork();
  }
  return options.cancelled() ? null : { matches, next: null };
}

/** Adapter for an in-memory projection or an existing bounded scan owner. */
export function scanSearchQuery(options: {
  cells: (sheet: string) => Record<string, WorkbookCell>;
  scan?: (sheet: string, after?: string) => Promise<{ cells: Record<string, WorkbookCell>; nextCursor: string | null }>;
  content: (sheet: string, cell: string, value: WorkbookCell, mode: SearchMode) => string;
}): SearchQuery {
  return async (sheet, text, mode, after, paging) => {
    const limit = paging?.matchLimit ?? SEARCH_PAGE_LIMIT;
    const lower = text.toLocaleLowerCase(),
      afterPosition = after ? position(after) : -1;
    const source = options.scan ? await options.scan(sheet, after) : { cells: options.cells(sheet), nextCursor: null };
    const entries = Object.entries(source.cells)
      .filter(([ref]) => position(ref) > afterPosition)
      .sort(([a], [b]) => position(a) - position(b));
    const cells: Record<string, WorkbookCell> = {};
    let matches = 0;
    for (let i = 0; i < entries.length; i++) {
      const [ref, value] = entries[i];
      if (options.content(sheet, ref, value, mode).toLocaleLowerCase().includes(lower)) {
        cells[ref] = value;
        if (++matches === limit)
          return { cells, matches, nextCursor: i + 1 < entries.length ? ref : source.nextCursor };
      }
    }
    return { cells, matches, nextCursor: source.nextCursor };
  };
}
