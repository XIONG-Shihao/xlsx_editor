import type { WorkbookPrintBoundary } from '../core/contracts';
import type { AcceptedPrintJob } from './useAcceptedPrintPreview';
import type { ViewGeometry, ViewViewport, ViewRectangle } from './viewGeometry';
import { addressPosition, colName as columnName } from './sheetModel';
import { validatePrintSettings } from '../core/print-model.mjs';
const rangeBounds = (range: string) => {
  const [a, z = a] = range.split(':');
  return { first: addressPosition(a), last: addressPosition(z) };
};
export interface PageBoundaryModel {
  area: string;
  rows: number;
  columns: number;
  pages: Array<{
    number: number;
    range: string;
    firstRow: number;
    lastRow: number;
    firstColumn: number;
    lastColumn: number;
  }>;
  repeatRows: string | null;
  repeatColumns: string | null;
}
/** Map accepted pagination to source axes; repeated prefixes are shown once. */
export function pageBoundaryModel(job: AcceptedPrintJob): PageBoundaryModel {
  const settings = validatePrintSettings(job.settings),
    bounds = rangeBounds(settings.area!),
    items = job.boundaries;
  if (!Array.isArray(items) || items.length !== job.pageCount || items.length < 1 || items.length > 100)
    throw Error('Page-boundary data is unavailable for this print runtime.');
  const span = (value: unknown, first: number, last: number, optional = false): value is [number, number] | null =>
    (optional && value === null) ||
    (Array.isArray(value) &&
      value.length === 2 &&
      value.every(Number.isSafeInteger) &&
      value[0] >= first &&
      value[1] <= last &&
      value[0] <= value[1]);
  items.forEach((item: WorkbookPrintBoundary, index: number) => {
    if (
      !item ||
      Object.keys(item).sort().join(',') !== 'bodyColumns,bodyRows,columns,number,rows' ||
      item.number !== index + 1 ||
      !span(item.rows, bounds.first.row, bounds.last.row) ||
      !span(item.columns, bounds.first.col, bounds.last.col) ||
      !span(item.bodyRows, item.rows[0], item.rows[1], true) ||
      !span(item.bodyColumns, item.columns[0], item.columns[1], true)
    )
      throw Error('Invalid accepted page boundaries.');
    const repeat = job.settings.repeatRows,
      columns = job.settings.repeatColumns;
    if (
      (item.bodyRows && repeat && item.bodyRows[0] < repeat.last) ||
      (item.bodyColumns && columns && item.bodyColumns[0] < rangeBounds(columns.last + '1').first.col + 1)
    )
      throw Error('Page boundaries overlap repeated titles.');
  });
  // The first accepted page is already carried by the lease; summaries must agree.
  const first = job.pages.find((page) => page.number === 1),
    summary = items[0];
  if (
    !first ||
    !first.rowIndices.length ||
    !first.columnIndices.length ||
    summary.rows[0] !== first.rowIndices[0] ||
    summary.rows[1] !== first.rowIndices.at(-1) ||
    summary.columns[0] !== first.columnIndices[0] ||
    summary.columns[1] !== first.columnIndices.at(-1)
  )
    throw Error('Page boundaries disagree with the accepted page.');
  const firstBodyRow = Math.min(...items.map((item) => item.bodyRows?.[0] ?? Infinity)),
    firstBodyColumn = Math.min(...items.map((item) => item.bodyColumns?.[0] ?? Infinity));
  const pages = items.map((item) => {
    const firstRow = !item.bodyRows || item.bodyRows[0] === firstBodyRow ? bounds.first.row : item.bodyRows[0],
      lastRow = item.bodyRows?.[1] ?? item.rows[1],
      firstColumn =
        !item.bodyColumns || item.bodyColumns[0] === firstBodyColumn ? bounds.first.col : item.bodyColumns[0],
      lastColumn = item.bodyColumns?.[1] ?? item.columns[1];
    return {
      number: item.number,
      firstRow,
      lastRow,
      firstColumn,
      lastColumn,
      range: columnName(firstColumn) + (firstRow + 1) + ':' + columnName(lastColumn) + (lastRow + 1),
    };
  });
  const keys = new Set(pages.map((p) => [p.firstRow, p.lastRow, p.firstColumn, p.lastColumn].join(',')));
  if (keys.size !== pages.length) throw Error('Duplicate accepted page boundaries.');
  return {
    area: job.settings.area!,
    rows: bounds.last.row + 1,
    columns: bounds.last.col + 1,
    pages,
    repeatRows: job.settings.repeatRows ? job.settings.repeatRows.first + ':' + job.settings.repeatRows.last : null,
    repeatColumns: job.settings.repeatColumns
      ? job.settings.repeatColumns.first + ':' + job.settings.repeatColumns.last
      : null,
  };
}
export function projectPageBoundary(
  page: PageBoundaryModel['pages'][number],
  geometry: ViewGeometry,
  viewport: ViewViewport
): ViewRectangle {
  if (page.lastRow + 1 >= geometry.rowOffsets.length || page.lastColumn + 1 >= geometry.columnOffsets.length)
    throw Error('Page boundaries exceed the shared worksheet axes.');
  return {
    left: geometry.rowHeaderWidth + geometry.columnOffsets[page.firstColumn] - viewport.scrollLeft,
    top: geometry.columnHeaderHeight + geometry.rowOffsets[page.firstRow] - viewport.scrollTop,
    width: geometry.columnOffsets[page.lastColumn + 1] - geometry.columnOffsets[page.firstColumn],
    height: geometry.rowOffsets[page.lastRow + 1] - geometry.rowOffsets[page.firstRow],
  };
}
