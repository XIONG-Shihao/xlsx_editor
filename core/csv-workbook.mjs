/** Portable indexed CSV model. One immutable text source and compact row
 * offsets own the file; cells are decoded on demand, and edits stay sparse.
 * Cache bounds limit retained decoded rows, never the file or address space. */
import { refusal } from './refusal-codes.mjs';

const rowCaches = new WeakMap();
const ROW_CACHE_BYTES = 2 * 1024 * 1024;
const own = (object, key) => Object.hasOwn(object, key);
export const CSV_SHEET = 'CSV';
function position(address) {
  const match = /^\$?([A-Z]+)\$?([1-9]\d*)$/i.exec(String(address));
  if (!match) throw refusal('internal-invalid-field', `Invalid CSV cell address: ${address}`);
  let column = 0;
  for (const letter of match[1].toUpperCase()) column = column * 26 + letter.charCodeAt(0) - 64;
  const row = Number(match[2]) - 1;
  column--;
  if (!Number.isSafeInteger(row) || !Number.isSafeInteger(column) || column < 0)
    throw refusal('csv-limit', 'CSV address exceeds JavaScript integer precision');
  return { row, column };
}
function address(row, column) {
  let name = '';
  for (let n = column + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name + (row + 1);
}
function rangeBounds(range) {
  const parts = String(range).split(':');
  if (parts.length > 2) throw refusal('internal-invalid-field', 'Invalid CSV range');
  const first = position(parts[0]),
    last = position(parts[1] ?? parts[0]);
  if (first.row > last.row || first.column > last.column) throw refusal('csv-malformed', 'Reversed CSV range');
  return { first, last };
}
function effectiveWidth(index, row) {
  return index.widthOverrides?.[row] ?? index.rowWidths[row] ?? (row < index.rows ? 1 : 0);
}
function sourceEnd(index, row) {
  let end = index.rowOffsets[row + 1];
  const tail = index.text.slice(Math.max(index.rowOffsets[row], end - 2), end);
  if (tail.endsWith('\n')) {
    end--;
    if (tail.endsWith('\r\n')) end--;
  } else if (tail.endsWith('\r')) end--;
  return end;
}
function parsedRow(index, row, decode = true) {
  if (row >= index.rowWidths.length) return [];
  let cache = rowCaches.get(index.rowOffsets);
  if (!cache) {
    cache = { rows: new Map(), bytes: 0 };
    rowCaches.set(index.rowOffsets, cache);
  }
  if (cache.rows.has(row)) {
    const entry = cache.rows.get(row);
    cache.rows.delete(row);
    cache.rows.set(row, entry);
    return entry.fields;
  }
  const start = index.rowOffsets[row],
    end = sourceEnd(index, row),
    fields = [];
  if (index.maxRowBytes !== undefined && (end - start) * 4 > index.maxRowBytes)
    throw refusal('csv-limit', 'CSV row exceeds the decode projection budget');
  // The desktop source may decode bounded compressed blocks. Read the row once,
  // not the entire document and not a block for every character/field.
  const text = index.text.slice(start, end);
  let fieldStart = start,
    quoted = false;
  const add = (stop) => {
    if (index.maxRowBytes !== undefined && (end - start) * 4 + (fields.length + 1) * 80 > index.maxRowBytes)
      throw refusal('csv-limit', 'CSV row exceeds the decode projection budget');
    const wasQuoted = text[fieldStart - start] === '"',
      field = { start: fieldStart, end: stop, quoted: wasQuoted };
    if (decode) {
      const raw = text.slice(fieldStart - start, stop - start);
      field.value = wasQuoted ? raw.slice(1, -1).replace(/""/g, '"') : raw;
    }
    fields.push(field);
  };
  for (let i = start; i < end; i++) {
    const c = text[i - start];
    if (c === '"') {
      if (quoted && text[i + 1 - start] === '"') i++;
      else if (quoted || i === fieldStart) quoted = !quoted;
    } else if (c === index.delimiter && !quoted) {
      add(i);
      fieldStart = i + 1;
    }
  }
  add(end);
  // Account for row backing text, decoded values, and per-field metadata.
  // A row-count cap alone could retain 128 near-2MiB rows in the owner.
  const bytes = (end - start) * 4 + fields.length * 80;
  if (decode && bytes <= ROW_CACHE_BYTES) {
    cache.rows.set(row, { fields, bytes });
    cache.bytes += bytes;
    while (cache.rows.size > 128 || cache.bytes > ROW_CACHE_BYTES) {
      const oldest = cache.rows.keys().next().value;
      cache.bytes -= cache.rows.get(oldest).bytes;
      cache.rows.delete(oldest);
    }
  }
  return fields;
}
/** Content-free owner diagnostics; estimated cache payload is bounded even
 * when source records are unusually wide. Does not decode additional rows. */
export function csvCacheStatistics(index) {
  const cache = rowCaches.get(index.rowOffsets);
  return {
    rows: cache?.rows.size ?? 0,
    estimatedBytes: cache?.bytes ?? 0,
    maxRows: 128,
    maxEstimatedBytes: ROW_CACHE_BYTES,
  };
}
export function detectCsvDelimiter(text) {
  const counts = { ',': 0, ';': 0, '\t': 0 };
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === '"') {
      if (quoted && text[i + 1] === '"') i++;
      else quoted = !quoted;
    } else if (!quoted) {
      if (c === '\r' || c === '\n') break;
      if (own(counts, c)) counts[c]++;
    }
  }
  return Object.entries(counts).sort((a, b) => b[1] - a[1])[0][0];
}
export function indexCsv(text, options = {}) {
  if (typeof text !== 'string') throw refusal('csv-bad-input', 'CSV input must be decoded text');
  const delimiter =
    options.delimiter === 'auto' ? detectCsvDelimiter(text) : options.delimiter === undefined ? ',' : options.delimiter;
  return { ...indexCsvChunks([text], { ...options, delimiter }), text };
}
/** The same scanner serves memory and file-backed callers. Keep one character
 * between chunks for escaped quotes and CRLF; offsets remain UTF-16 positions. */
export function indexCsvChunks(chunks, { delimiter = ',', encoding = 'utf-8', bom = false, maxIndexBytes } = {}) {
  if (![',', ';', '\t'].includes(delimiter)) throw refusal('csv-unsupported', 'Unsupported CSV delimiter');
  if (maxIndexBytes !== undefined && (!Number.isSafeInteger(maxIndexBytes) || maxIndexBytes < 20))
    throw refusal('csv-malformed', 'Invalid CSV row-index budget');
  const offsets = [0],
    widths = [];
  let recordNewline;
  let width = 1,
    columns = 1,
    nonemptyCells = 0,
    storedCells = 0,
    quoted = false,
    closed = false,
    fieldStart = true,
    hasValue = false;
  const cell = () => {
    if (hasValue) nonemptyCells++;
    hasValue = false;
    closed = false;
    fieldStart = true;
  };
  const row = (end) => {
    if (maxIndexBytes !== undefined && 8 + (widths.length + 1) * 12 > maxIndexBytes)
      throw refusal('csv-limit', 'CSV row index exceeds its metadata budget');
    if (width > 0xffffffff || !Number.isSafeInteger(storedCells + width))
      throw refusal('csv-limit', 'CSV index exceeds address precision');
    cell();
    widths.push(width);
    columns = Math.max(columns, width);
    storedCells += width;
    offsets.push(end);
    width = 1;
  };
  let offset = 0,
    pending = '',
    last = '';
  // Cache each next structural position. A failed search scans the remainder
  // only once, and quoted fields need only the next quote. This avoids both
  // per-character dispatch and allocating a RegExp match for every delimiter.
  const scan = (chunk, final) => {
    const text = pending + chunk,
      end = text.length - (final ? 0 : 1);
    // Short fields make repeated string searches more expensive than the small
    // state machine. A bounded sample selects that original path; this choice
    // changes only scanning strategy, never quote or newline semantics.
    let sampleQuoted = quoted,
      structuralCount = 0;
    const sampleEnd = Math.min(end, 4096);
    for (let at = 0; at < sampleEnd; at++) {
      const c = text[at];
      if (c === '"') {
        sampleQuoted = !sampleQuoted;
        structuralCount++;
      } else if (!sampleQuoted && (c === delimiter || c === '\r' || c === '\n')) structuralCount++;
    }
    const useSearch = structuralCount * 8 <= sampleEnd;
    let i = 0,
      quote = -1,
      separator = -1,
      cr = -1,
      lf = -1;
    const find = (character, previous) => {
      if (previous >= i) return previous;
      const at = text.indexOf(character, i);
      return at < 0 ? end : Math.min(at, end);
    };
    while (i < end) {
      if (useSearch) {
        quote = find('"', quote);
        let next = quote;
        if (!quoted) {
          separator = find(delimiter, separator);
          cr = find('\r', cr);
          lf = find('\n', lf);
          next = Math.min(next, separator, cr, lf);
        }
        if (next > i) {
          if (!quoted && closed)
            throw refusal('csv-malformed', `Malformed CSV: text follows a closing quote near offset ${offset + i}`);
          hasValue = true;
          if (!quoted) fieldStart = false;
          i = next;
        }
        if (i >= end) break;
      }
      const c = text[i];
      if (quoted) {
        if (c === '"') {
          if (text[i + 1] === '"') {
            hasValue = true;
            i++;
          } else {
            quoted = false;
            closed = true;
          }
        } else hasValue = true;
      } else if (c === '"' && fieldStart) {
        quoted = true;
        fieldStart = false;
      } else if (c === delimiter) {
        cell();
        width++;
      } else if (c === '\r' || c === '\n') {
        recordNewline ??= c === '\r' && text[i + 1] === '\n' ? '\r\n' : c;
        if (c === '\r' && text[i + 1] === '\n') i++;
        row(offset + i + 1);
      } else {
        if (closed)
          throw refusal('csv-malformed', `Malformed CSV: text follows a closing quote near offset ${offset + i}`);
        hasValue = true;
        fieldStart = false;
      }
      i++;
    }
    pending = text.slice(i);
    offset += i;
  };
  for (const chunk of chunks) {
    if (typeof chunk !== 'string') throw refusal('csv-bad-input', 'CSV input must be decoded text');
    if (chunk.length) last = chunk.at(-1);
    scan(chunk, false);
  }
  scan('', true);
  if (quoted) throw refusal('csv-malformed', 'Malformed CSV: unterminated quoted field');
  if (offsets.at(-1) < offset || !widths.length) row(offset);
  return {
    kind: 'csv',
    sheet: CSV_SHEET,
    rowOffsets: Float64Array.from(offsets),
    rowWidths: Uint32Array.from(widths),
    rows: widths.length,
    columns,
    nonemptyCells,
    storedCells,
    delimiter,
    encoding,
    bom,
    newline: recordNewline ?? '\n',
    terminalNewline: /[\r\n]$/.test(last),
    changes: {},
    widthOverrides: {},
  };
}
export function csvCell(index, row, column) {
  if (
    !Number.isSafeInteger(row) ||
    !Number.isSafeInteger(column) ||
    row < 0 ||
    column < 0 ||
    row >= index.rows ||
    column >= effectiveWidth(index, row)
  )
    return null;
  const key = address(row, column);
  if (own(index.changes, key)) return index.changes[key];
  return parsedRow(index, row)[column]?.value ?? '';
}
export function csvCellInfo(index, row, column) {
  const value = csvCell(index, row, column);
  return value === null ? null : { value, rawValue: value, dataType: 'text' };
}
export function csvReadRange(index, range, { offset = 0, limit = 100 } = {}) {
  if (!Number.isSafeInteger(offset) || offset < 0 || !Number.isSafeInteger(limit) || limit < 0)
    throw refusal('internal-invalid-field', 'Invalid CSV page');
  const { first, last } = rangeBounds(range),
    width = last.column - first.column + 1,
    total = (last.row - first.row + 1) * width;
  if (!Number.isSafeInteger(total)) throw refusal('csv-limit', 'CSV range exceeds JavaScript integer precision');
  const items = [];
  for (let ordinal = offset; ordinal < Math.min(total, offset + limit); ordinal++) {
    const row = first.row + Math.floor(ordinal / width),
      column = first.column + (ordinal % width);
    items.push({
      sheet: CSV_SHEET,
      cell: address(row, column),
      ...(csvCellInfo(index, row, column) ?? { value: null, rawValue: null, dataType: 'blank' }),
    });
  }
  return { items, total, offset, nextOffset: offset + items.length < total ? offset + items.length : null };
}
function literal(value) {
  if (value === null) return '';
  if (!['string', 'number', 'boolean'].includes(typeof value) || (typeof value === 'number' && !Number.isFinite(value)))
    throw refusal('internal-invalid-field', 'Invalid CSV literal value');
  return String(value);
}
export function applyCsvCommands(index, commands) {
  if (!Array.isArray(commands) || !commands.length)
    throw refusal('csv-bad-input', 'CSV command batch must not be empty');
  if (commands.some((command) => command?.type === 'clearHyperlinks'))
    throw refusal(
      'csv-bad-input',
      'Clear Hyperlinks requires the native XLSX owner; CSV stores literal text and has no hyperlink metadata.'
    );
  const next = { ...index, changes: { ...index.changes }, widthOverrides: { ...index.widthOverrides } },
    edits = [],
    affected = new Map();
  const write = (raw) => {
    if (raw.sheet !== CSV_SHEET) throw refusal('csv-refused', `CSV has one worksheet named ${CSV_SHEET}`);
    if (raw.operation || ['formula', 'bold', 'italic', 'numberFormat', 'styleId'].some((key) => raw[key] !== undefined))
      throw refusal(
        'csv-bad-input',
        'CSV stores literal text; formula, formatting and workbook structure require XLSX'
      );
    if (!own(raw, 'value')) throw refusal('csv-bad-input', 'CSV cell edit requires a literal value');
    const { row, column } = position(raw.cell),
      ref = address(row, column),
      previous = csvCell(next, row, column),
      value = literal(raw.value);
    if (own(raw, 'expectedValue') && !Object.is(previous, raw.expectedValue))
      throw refusal('csv-refused', `Conflict: unexpected value in CSV!${ref}`);
    if (own(raw, 'expectedRawValue') && !Object.is(previous, raw.expectedRawValue))
      throw refusal('csv-refused', `Conflict: unexpected stored value in CSV!${ref}`);
    if (own(raw, 'expectedFormula') && raw.expectedFormula !== null)
      throw refusal('csv-bad-input', `Conflict: CSV!${ref} has no formula`);
    const original = row < index.rowWidths.length ? (parsedRow(index, row)[column]?.value ?? null) : null;
    if (value === original) delete next.changes[ref];
    else next.changes[ref] = value;
    const oldWidth = effectiveWidth(next, row),
      oldRows = next.rows;
    if (row >= oldRows) {
      next.storedCells += row - oldRows + 1;
      next.rows = row + 1;
    }
    if (column >= oldWidth) {
      next.widthOverrides[row] = column + 1;
      next.storedCells += column + 1 - Math.max(1, oldWidth);
    }
    next.columns = Math.max(next.columns, column + 1);
    next.nonemptyCells += Number(value !== '') - Number(previous !== null && previous !== '');
    const edit = { ...raw, sheet: CSV_SHEET, cell: ref, value };
    edits.push(edit);
    affected.set(ref, { sheet: CSV_SHEET, cell: ref, ...csvCellInfo(next, row, column) });
  };
  for (const command of commands) {
    if (command.type === 'setCells') {
      if (!Array.isArray(command.cells) || !command.cells.length)
        throw refusal('csv-bad-input', 'CSV setCells requires cells');
      for (const cell of command.cells) write({ ...cell, sheet: command.sheet });
    } else if (command.type === 'updateRows') {
      if (command.sheet !== CSV_SHEET) throw refusal('csv-refused', `CSV has one worksheet named ${CSV_SHEET}`);
      const { first, last } = rangeBounds(command.range),
        keyColumn = position(command.keyColumn + '1').column;
      if (keyColumn < first.column || keyColumn > last.column)
        throw refusal('csv-unsupported', 'CSV key column is outside the selected range');
      if (!Array.isArray(command.updates) || !command.updates.length)
        throw refusal('csv-bad-input', 'CSV keyed updates must not be empty');
      const requested = new Map();
      for (const update of command.updates) {
        if (requested.has(update.key)) throw refusal('csv-duplicate', `Duplicate requested key: ${update.key}`);
        requested.set(update.key, { update, row: null });
      }
      for (let row = first.row; row <= Math.min(last.row, next.rows - 1); row++) {
        const match = requested.get(csvCell(next, row, keyColumn));
        if (match) {
          if (match.row !== null) throw refusal('csv-duplicate', `Duplicate key: ${csvCell(next, row, keyColumn)}`);
          match.row = row;
        }
      }
      for (const [key, { update, row }] of requested) {
        if (row === null) throw refusal('internal-missing-field', `Missing key: ${key}`);
        if (!Array.isArray(update.cells) || !update.cells.length)
          throw refusal('csv-bad-input', 'CSV keyed update requires cells');
        for (const { column, ...cell } of update.cells) {
          const targetColumn = position(column + '1').column;
          if (targetColumn < first.column || targetColumn > last.column)
            throw refusal('csv-unsupported', 'CSV keyed update is outside its range');
          if (!own(cell, 'expectedValue') && !own(cell, 'expectedRawValue'))
            throw refusal('csv-bad-input', 'CSV keyed updates require an expected old value');
          write({ ...cell, sheet: CSV_SHEET, cell: address(row, targetColumn) });
        }
      }
    } else if (command.type !== undefined)
      throw refusal('csv-unsupported', `CSV command is unsupported: ${command.type}`);
    else write(command);
  }
  return { index: next, edits, affected: [...affected.values()] };
}
function encodeField(value, delimiter, quoted) {
  return quoted || value.includes(delimiter) || /[,;\t"\r\n]/.test(value)
    ? '"' + value.replaceAll('"', '""') + '"'
    : value;
}
/** Changed fields in UTF-16 source coordinates; untouched fields/separators remain
 * in the source. Node can translate these spans to byte offsets without
 * materializing and encoding another complete output string. */
export function csvTextEdits(index) {
  const changedRows = new Map();
  for (const [ref, value] of Object.entries(index.changes)) {
    const { row, column } = position(ref);
    if (!changedRows.has(row)) changedRows.set(row, new Map());
    changedRows.get(row).set(column, value);
  }
  for (const row of Object.keys(index.widthOverrides))
    if (!changedRows.has(Number(row))) changedRows.set(Number(row), new Map());
  const patches = [];
  for (const row of [...changedRows.keys()].filter((row) => row < index.rowWidths.length).sort((a, b) => a - b)) {
    const fields = parsedRow(index, row, false),
      width = effectiveWidth(index, row),
      changes = changedRows.get(row) ?? new Map();
    let previous;
    for (const column of [...changes.keys()].filter((column) => column < fields.length).sort((a, b) => a - b)) {
      const field = fields[column],
        text = encodeField(changes.get(column), index.delimiter, field.quoted);
      // Adjacent fields share one patch. Retain the delimiter's exact spelling
      // and never rebuild or decode unchanged field values for serialization.
      if (previous?.column === column - 1) {
        previous.patch.end = field.end;
        previous.patch.text += index.delimiter + text;
        previous.column = column;
      } else {
        const patch = { start: field.start, end: field.end, text };
        patches.push(patch);
        previous = { column, patch };
      }
    }
    if (width > fields.length) {
      const values = [];
      for (let column = fields.length; column < width; column++)
        values.push(encodeField(changes.get(column) ?? '', index.delimiter, false));
      const end = fields.at(-1)?.end ?? sourceEnd(index, row),
        text = index.delimiter + values.join(index.delimiter);
      if (previous?.patch.end === end) previous.patch.text += text;
      else patches.push({ start: end, end, text });
    }
  }
  if (index.rows > index.rowWidths.length) {
    const chunks = [];
    if (!index.terminalNewline) chunks.push(index.newline);
    for (let row = index.rowWidths.length; row < index.rows; row++) {
      const values = [];
      for (let column = 0; column < effectiveWidth(index, row); column++)
        values.push(encodeField(csvCell(index, row, column) ?? '', index.delimiter, false));
      chunks.push(values.join(index.delimiter));
      if (row < index.rows - 1 || index.terminalNewline) chunks.push(index.newline);
    }
    patches.push({ start: index.text.length, end: index.text.length, text: chunks.join('') });
  }
  return patches;
}
export function serializeCsv(index, edits) {
  if (edits?.length) index = applyCsvCommands(index, edits).index;
  const chunks = [];
  let cursor = 0;
  for (const patch of csvTextEdits(index)) {
    chunks.push(index.text.slice(cursor, patch.start), patch.text);
    cursor = patch.end;
  }
  chunks.push(index.text.slice(cursor));
  return chunks.join('');
}

// Share address arithmetic with the protocol adapter; CSV has no Excel grid bounds.
export const csvPosition = position;
export const csvAddress = address;
