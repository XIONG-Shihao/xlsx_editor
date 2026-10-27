import {
  PASTE_BYTE_LIMIT,
  PASTE_PAGE_BYTES,
  PASTE_PAGE_CELLS,
  PASTE_POSITION_LIMIT,
  pasteBytes,
  validatePastePlan,
  validatePasteCells,
  type LegacyPastePlan,
  type PastePage,
} from '../core/paste-protocol.mjs';

export interface PasteInspection extends LegacyPastePlan {
  rawBytes: number;
}
/** One field at a time: quoted tabs/newlines and ragged row holes remain exact.
 * The browser clipboard API supplies one string; no second rows/cells array is built. */
function* fields(text: string): Generator<{ row: number; col: number; text: string }> {
  let row = 0,
    col = 0,
    value = '',
    quoted = false,
    closed = false,
    fieldStarted = false,
    finishedRows = 0;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quoted) {
      if (c === '"') {
        if (text[i + 1] === '"') {
          value += '"';
          i++;
        } else {
          quoted = false;
          closed = true;
        }
      } else value += c;
    } else if (c === '"' && !fieldStarted) {
      quoted = true;
      fieldStarted = true;
    } else if (c === '\t' || c === '\r' || c === '\n') {
      yield { row, col, text: value };
      value = '';
      closed = false;
      fieldStarted = false;
      if (c === '\t') col++;
      else {
        row++;
        col = 0;
        finishedRows++;
        if (c === '\r' && text[i + 1] === '\n') i++;
      }
    } else {
      if (closed) throw Error('Malformed delimited file: text follows a closing quote');
      value += c;
      fieldStarted = true;
    }
    if (value.length > 32768) throw Error('Paste field exceeds cell text limits');
    if (row > PASTE_POSITION_LIMIT || col >= 16384) throw Error('Paste exceeds workbook position limits');
  }
  if (quoted) throw Error('Malformed delimited file: unterminated quote');
  if (fieldStarted || value || col || !finishedRows) yield { row, col, text: value };
}
export function inspectPaste(
  text: string,
  options: { sheet: string; start: string; policy?: 'grid' | 'text' }
): PasteInspection {
  if (typeof text !== 'string') throw Error('Clipboard text is required');
  const rawBytes = pasteBytes(text);
  if (rawBytes > PASTE_BYTE_LIMIT) throw Error('Clipboard text exceeds 64 MiB');
  let rows = 0,
    columns = 0,
    count = 0;
  for (const field of fields(text)) {
    rows = field.row + 1;
    columns = Math.max(columns, field.col + 1);
    count++;
    if (rows * columns > PASTE_POSITION_LIMIT) throw Error('Paste exceeds 1,000,000 positions');
    validatePasteCells([[0, field.text]], { sheet: options.sheet, start: options.start, rows: 1, columns: 1 });
  }
  return {
    ...(validatePastePlan({
      shape: { sheet: options.sheet, start: options.start, rows, columns },
      count,
      policy: options.policy ?? 'grid',
    }) as LegacyPastePlan),
    rawBytes,
  };
}
export function* pastePages(text: string, inspection: PasteInspection): Generator<PastePage> {
  const plan = validatePastePlan({ shape: inspection.shape, count: inspection.count, policy: inspection.policy });
  if (pasteBytes(text) !== inspection.rawBytes || inspection.rawBytes > PASTE_BYTE_LIMIT)
    throw Error('Clipboard text changed after inspection');
  let cells: [number, string][] = [],
    bytes = 0,
    sequence = 0,
    count = 0,
    rows = 0,
    columns = 0,
    previous = -1;
  for (const field of fields(text)) {
    const entry: [number, string] = [field.row * plan.shape.columns + field.col, field.text];
    const size = pasteBytes(JSON.stringify(entry) + '\n');
    if (cells.length && (cells.length >= PASTE_PAGE_CELLS || bytes + size > PASTE_PAGE_BYTES)) {
      validatePasteCells(cells, plan.shape, previous);
      previous = cells.at(-1)![0];
      yield { sequence: sequence++, cells };
      cells = [];
      bytes = 0;
    }
    if (size > PASTE_PAGE_BYTES) throw Error('Paste field exceeds page byte limit');
    cells.push(entry);
    bytes += size;
    count++;
    rows = field.row + 1;
    columns = Math.max(columns, field.col + 1);
  }
  if (cells.length) {
    validatePasteCells(cells, plan.shape, previous);
    yield { sequence, cells };
  }
  if (count !== plan.count || rows !== plan.shape.rows || columns !== plan.shape.columns)
    throw Error('Clipboard shape changed after inspection');
}
