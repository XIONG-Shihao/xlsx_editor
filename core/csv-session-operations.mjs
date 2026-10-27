import { readCsvFileIndex } from './csv-file.mjs';
import { csvCellInfo, csvReadRange, csvPosition, csvAddress, applyCsvCommands } from './csv-workbook.mjs';
import { writeCsvIndex } from './csv-file-save.mjs';
import { refusal } from './refusal-codes.mjs';

export async function openCsv(source, delimiter, expectedRevision) {
  // The document owner holds an immutable snapshot for the entire session.
  // Keep source bytes there; range/search requests decode bounded windows only.
  return readCsvFileIndex(source, { delimiter }, expectedRevision);
}
export const applyCsv = (index, commands) => applyCsvCommands(index, commands).index;
export function csvView(index, revision) {
  return {
    revision,
    readOnly: false,
    limitations: [],
    calculation: 'literal',
    sheets: [
      {
        name: 'CSV',
        state: 'visible',
        path: '',
        cells: {},
        rows: index.rows,
        columns: index.columns,
        merges: [],
        widths: [],
        storedCells: index.storedCells,
        storedFormulaCount: 0,
      },
    ],
  };
}
export function readCsv(index, range) {
  const [first, last = first] = String(range).split(':'),
    a = csvPosition(first),
    b = csvPosition(last);
  if (b.row < a.row || b.column < a.column || (b.row - a.row + 1) * (b.column - a.column + 1) > 10000)
    throw refusal('csv-limit', 'Read at most 10,000 positions at a time.');
  return {
    cells: Object.fromEntries(
      csvReadRange(index, range, { limit: 10000 }).items.map((item) => [
        item.cell,
        { value: item.value, rawValue: item.rawValue, dataType: 'text' },
      ])
    ),
  };
}
export function scanCsv(index, after, query, mode = 'raw', options = {}) {
  const cells = {};
  let cursor = after === undefined ? 0 : Number(after),
    scanned = 0,
    matches = 0;
  if (!Number.isSafeInteger(cursor) || cursor < 0 || !['raw', 'display', 'formula'].includes(mode))
    throw refusal('internal-invalid-field', 'Invalid CSV query cursor or mode.');
  const matchLimit = options.matchLimit ?? 1000;
  if (!Number.isSafeInteger(matchLimit) || matchLimit < 1 || matchLimit > 1000)
    throw refusal('internal-invalid-field', 'Invalid CSV search match limit');
  let first = { row: 0, column: 0 },
    last = { row: index.rows - 1, column: index.columns - 1 };
  if (options.range !== undefined) {
    const parts = String(options.range).split(':');
    if (parts.length > 2) throw refusal('internal-invalid-field', 'Invalid CSV query range');
    first = csvPosition(parts[0]);
    last = csvPosition(parts[1] ?? parts[0]);
    if (last.row < first.row || last.column < first.column)
      throw refusal('internal-invalid-field', 'Invalid CSV query range');
    last = { row: Math.min(last.row, index.rows - 1), column: Math.min(last.column, index.columns - 1) };
  }
  const columns = Math.max(0, last.column - first.column + 1),
    total = Math.max(0, last.row - first.row + 1) * columns;
  if (!Number.isSafeInteger(total)) throw refusal('csv-limit', 'CSV query exceeds address precision.');
  const normalized = query?.toLocaleLowerCase('en');
  for (; cursor < total && scanned < 50000 && matches < matchLimit; cursor++, scanned++) {
    const row = first.row + Math.floor(cursor / columns),
      column = first.column + (cursor % columns);
    const width = index.widthOverrides[row] ?? index.rowWidths[row] ?? 1;
    if (column >= width) {
      cursor = (row - first.row + 1) * columns - 1;
      continue;
    }
    const cell = csvCellInfo(index, row, column);
    if (
      cell &&
      (query === undefined || (mode !== 'formula' && String(cell.value).toLocaleLowerCase('en').includes(normalized)))
    ) {
      cells[csvAddress(row, column)] = cell;
      matches++;
    }
  }
  return { cells, matches: Object.keys(cells).length, scanned, nextCursor: cursor < total ? String(cursor) : null };
}
export function summarizeCsv(index, range) {
  const parts = String(range).split(':');
  if (parts.length > 2) throw refusal('internal-invalid-field', 'Invalid CSV summary range');
  const [first, last = first] = parts,
    a = csvPosition(first),
    b = csvPosition(last);
  const positions = (b.row - a.row + 1) * (b.column - a.column + 1);
  if (b.row < a.row || b.column < a.column || positions > 1000000)
    throw refusal('csv-limit', 'CSV summaries support at most one million positions.');
  let count = 0;
  for (let row = a.row; row <= Math.min(b.row, index.rows - 1); row++)
    for (let column = a.column; column <= Math.min(b.column, index.columns - 1); column++) {
      const cell = csvCellInfo(index, row, column);
      if (cell && cell.value !== '') count++;
    }
  return {
    count,
    numericCount: 0,
    sum: null,
    average: null,
    min: null,
    max: null,
    errors: 0,
    unresolved: 0,
    unsafeNumbers: 0,
  };
}
export async function exportCsv(index, source, destination) {
  writeCsvIndex(index, source, destination);
}
