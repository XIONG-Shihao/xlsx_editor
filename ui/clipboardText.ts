import { parseDelimited } from '../core/delimited';
import { parseCellInput } from '../core/cell-input.mjs';
import { addressPosition as cellPosition, colName as columnName } from './sheetModel';
import { PASTE_BYTE_LIMIT, pasteBytes, validatePasteCells } from '../core/paste-protocol.mjs';
import { inspectPaste } from './largePaste';

export const CLIPBOARD_TEXT_POSITION_LIMIT = 10_000;
const fieldShape = { sheet: 'Clipboard', start: 'A1', rows: 1, columns: 1 };

function validateField(value: unknown): asserts value is string {
  if (typeof value !== 'string') throw Error('Clipboard fields must be text');
  // Match accepted setCells, whose text bound is one character below staged pages.
  if (value.length > 32767) throw Error('Cell text exceeds 32767 characters');
  validatePasteCells([[0, value]], fieldShape);
}

/** Quote every field so a copied blank cell is distinct from an empty clipboard. */
export function serializeClipboardText(rows: readonly (readonly string[])[]): string {
  if (!Array.isArray(rows) || !rows.length || !Array.isArray(rows[0]) || !rows[0].length)
    throw Error('Clipboard rows require a nonempty rectangle');
  const columns = rows[0].length;
  if (rows.length * columns > CLIPBOARD_TEXT_POSITION_LIMIT) throw Error('Clipboard text exceeds 10,000 positions');
  let bytes = rows.length - 1 + rows.length * (columns - 1);
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    if (!Object.hasOwn(rows, r) || !Array.isArray(row) || row.length !== columns)
      throw Error('Clipboard rows require a dense rectangle');
    for (let c = 0; c < columns; c++) {
      if (!Object.hasOwn(row, c)) throw Error('Clipboard rows require a dense rectangle');
      const value = row[c];
      validateField(value);
      let quotes = 0;
      for (const char of value) if (char === '"') quotes++;
      bytes += pasteBytes(value) + quotes + 2;
      if (bytes > PASTE_BYTE_LIMIT) throw Error('Clipboard text exceeds 64 MiB');
    }
  }
  return rows
    .map((row: readonly string[]) => row.map((value: string) => '"' + value.replaceAll('"', '""') + '"').join('\t'))
    .join('\n');
}

/** Literal paste writes the entire rectangle, including empty ragged-row holes.
 * The caller may treat an empty external clipboard as a no-op before invoking. */
export function literalClipboardCells(
  text: string,
  options: { sheet: string; start: string }
): {
  range: string;
  cells: Array<{ cell: string; value: string }>;
} {
  if (
    !options ||
    typeof options !== 'object' ||
    Array.isArray(options) ||
    Object.keys(options).length !== 2 ||
    !Object.hasOwn(options, 'sheet') ||
    !Object.hasOwn(options, 'start')
  )
    throw Error('Clipboard target requires sheet and start');
  const inspection = inspectPaste(text, { ...options, policy: 'text' });
  const { rows, columns } = inspection.shape;
  if (rows * columns > CLIPBOARD_TEXT_POSITION_LIMIT) throw Error('Clipboard text exceeds 10,000 positions');
  const parsed = parseDelimited(text, '\t').rows;
  if (
    parsed.length !== rows ||
    Math.max(...parsed.map((row) => row.length)) !== columns ||
    parsed.reduce((count, row) => count + row.length, 0) !== inspection.count
  )
    throw Error('Clipboard parsers disagree on the rectangle');
  const first = cellPosition(options.start),
    cells: Array<{ cell: string; value: string }> = [];
  for (let r = 0; r < rows; r++)
    for (let c = 0; c < columns; c++) {
      const value = parsed[r][c] ?? '';
      validateField(value);
      const literal = parseCellInput(value, true);
      if (typeof literal.value !== 'string') throw Error('Literal clipboard field must remain text');
      cells.push({ cell: columnName(first.col + c) + (first.row + r + 1), value: literal.value });
    }
  const last = columnName(first.col + columns - 1) + (first.row + rows);
  return { range: options.start === last ? last : options.start + ':' + last, cells };
}
