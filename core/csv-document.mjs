import { createHash } from 'node:crypto';
import { Buffer } from 'node:buffer';
import { decodeCsvBytes, encodeCsvText } from './csv-codec.mjs';
import { indexCsv, applyCsvCommands, serializeCsv, csvReadRange, csvTextEdits } from './csv-workbook.mjs';
import { refusal } from './refusal-codes.mjs';

const revision = (bytes) => createHash('sha256').update(bytes).digest('hex');

export function readCsvWorkbook(bytes, options = {}) {
  const decoded = decodeCsvBytes(bytes);
  return { ...indexCsv(decoded.text, { ...decoded, delimiter: 'auto', ...options }), revision: revision(bytes) };
}

/** Publish only after the complete command batch has passed its guards. The
 * caller owns file authorization, live revision checks and atomic publication. */
function prepareCsvEdit(bytes, commands, expectedRevision, options = {}) {
  const original = readCsvWorkbook(bytes, options);
  if (expectedRevision && expectedRevision.replace(/^sha256:/, '') !== original.revision)
    throw refusal('csv-stale', 'Conflict: CSV revision is stale');
  const result = applyCsvCommands(original, commands);
  return { original, result };
}

/** Desktop save needs bytes/revision only; do not rebuild a discarded output index. */
export function editCsvBytes(bytes, commands, expectedRevision, options = {}) {
  const { original, result } = prepareCsvEdit(bytes, commands, expectedRevision, options);
  const patches = csvTextEdits(result.index);
  const prefix = original.bom ? (original.encoding === 'utf-8' ? 3 : 2) : 0;
  let character = 0,
    offset = prefix;
  const byteOffset = (next) => {
    if (next < character || next > original.text.length) throw refusal('csv-malformed', 'Invalid CSV source span');
    offset +=
      original.encoding === 'utf-8'
        ? Buffer.byteLength(original.text.slice(character, next), 'utf8')
        : (next - character) * 2;
    character = next;
    return offset;
  };
  let size = bytes.length;
  const replacements = patches.map((patch) => {
    const start = byteOffset(patch.start),
      end = byteOffset(patch.end);
    const data = encodeCsvText(patch.text, original.encoding, false);
    size += data.length - (end - start);
    return { start, end, data };
  });
  const output = new Uint8Array(size);
  let from = 0,
    to = 0;
  for (const patch of replacements) {
    output.set(bytes.subarray(from, patch.start), to);
    to += patch.start - from;
    output.set(patch.data, to);
    to += patch.data.length;
    from = patch.end;
  }
  output.set(bytes.subarray(from), to);
  return { bytes: output, revision: revision(output) };
}

export function editCsvWorkbook(bytes, commands, expectedRevision, options = {}) {
  const { original, result } = prepareCsvEdit(bytes, commands, expectedRevision, options);
  const output = encodeCsvText(serializeCsv(result.index), original.encoding, original.bom);
  const index = readCsvWorkbook(output, { delimiter: original.delimiter });
  const affected = result.affected.flatMap((item) => {
    const before = csvReadRange(original, item.cell).items[0];
    return before.value === item.value
      ? []
      : [{ ...item, before: { value: before.value, formula: null }, after: { value: item.value, formula: null } }];
  });
  return { bytes: output, index, edits: result.edits, affected };
}

export function createCsvWorkbook(rows) {
  const literal = (value) => {
    if (value !== null && !['string', 'number', 'boolean'].includes(typeof value))
      throw refusal('csv-bad-input', 'CSV creation requires literal scalar values');
    if (typeof value === 'number' && !Number.isFinite(value))
      throw refusal('csv-bad-input', 'CSV numbers must be finite');
    const text = value === null ? '' : String(value);
    return /[",;\t\r\n]/.test(text) ? '"' + text.replaceAll('"', '""') + '"' : text;
  };
  return encodeCsvText(rows.map((row) => row.map(literal).join(',')).join('\r\n') + (rows.length ? '\r\n' : ''));
}
