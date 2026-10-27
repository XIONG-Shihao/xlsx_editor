/** Fresh writer output only: expand adjacent short cells into full cell records.
 * SheetJS documents ids12–18 as Excel-supported but absent from MS-XLSB.
 * This closed transform lets independent readers inspect every exported value. */
import { unzipSync, zipSync } from 'fflate';
import { refusal } from './refusal-codes.mjs';
const fail = (message) => {
  throw refusal('fill-bad-input', 'Data-only XLSB normalization refused: ' + message);
};
function records(bytes) {
  const result = [];
  let at = 0;
  while (at < bytes.length) {
    let id = bytes[at++] & 127;
    if (bytes[at - 1] & 128) {
      if (at >= bytes.length || bytes[at] & 128) fail('record type encoding');
      id |= bytes[at++] << 7;
    }
    let length = 0,
      shift = 0,
      byte;
    do {
      if (at >= bytes.length || shift >= 28) fail('record length encoding');
      byte = bytes[at++];
      length |= (byte & 127) << shift;
      shift += 7;
    } while (byte & 128);
    if (length < 0 || at + length > bytes.length) fail('record body truncation');
    result.push({ id, body: bytes.subarray(at, at + length) });
    at += length;
    if (result.length > 2000100) fail('record count budget');
  }
  return result;
}
function header(id, length) {
  const result = [id < 128 ? id : (id & 127) | 128];
  if (id >= 128) result.push(id >>> 7);
  do {
    const byte = length & 127;
    length >>>= 7;
    result.push(byte | (length ? 128 : 0));
  } while (length);
  return Uint8Array.from(result);
}
/** Returns only final bytes; no source mutation and no unsupported graph reuse. */
export function normalizeDataOnlyXlsb(input) {
  if (!(input instanceof Uint8Array) || input.length > 64 * 1024 * 1024) fail('input byte budget');
  const parts = unzipSync(input);
  let changed = 0,
    positions = 0,
    total = 0;
  for (const bytes of Object.values(parts)) {
    total += bytes.length;
    if (total > 128 * 1024 * 1024 || bytes.length > 64 * 1024 * 1024) fail('expanded byte budget');
  }
  for (const [name, bytes] of Object.entries(parts)) {
    if (!/^xl\/worksheets\/sheet[1-9]\d*\.bin$/.test(name)) continue;
    const chunks = [];
    let row = -1,
      column = -1,
      lastRow = -1,
      size = 0;
    for (let { id, body } of records(bytes)) {
      if (id === 0) {
        if (body.length < 17) fail('row header length');
        row = new DataView(body.buffer, body.byteOffset, body.byteLength).getUint32(0, true);
        if (row <= lastRow || row >= 1048576) fail('row ordering');
        lastRow = row;
        column = -1;
      } else if (id >= 8 && id <= 11) fail('formula in fresh scalar output');
      else if ((id >= 1 && id <= 7) || (id >= 12 && id <= 18)) {
        if (row < 0 || ++positions > 1000000) fail('cell position budget');
        let nextColumn;
        if (id >= 12) {
          if (column < 0 || body.length < 4) fail('short cell has no preceding column');
          nextColumn = column + 1;
          if (id === 14) {
            if (body.length !== 8 || body[5] || body[6] || body[7]) fail('short error padding');
            body = body.subarray(0, 5);
          }
          const replacement = new Uint8Array(4 + body.length);
          new DataView(replacement.buffer).setUint32(0, nextColumn, true);
          replacement.set(body, 4);
          body = replacement;
          id -= 11;
          changed++;
        } else {
          if (body.length < 8) fail('full cell prefix');
          nextColumn = new DataView(body.buffer, body.byteOffset, body.byteLength).getUint32(0, true);
        }
        if (nextColumn <= column || nextColumn >= 16384) fail('cell ordering');
        column = nextColumn;
        if (
          (id === 1 && body.length !== 8) ||
          ([2, 7].includes(id) && body.length !== 12) ||
          ([3, 4].includes(id) && body.length !== 9) ||
          (id === 5 && body.length !== 16)
        )
          fail('scalar record length');
        if (id === 6) {
          if (body.length < 12) fail('inline text length');
          const length = new DataView(body.buffer, body.byteOffset, body.byteLength).getUint32(8, true);
          if (length > 32767 || body.length !== 12 + length * 2) fail('inline text budget');
        }
      }
      const prefix = header(id, body.length);
      size += prefix.length + body.length;
      chunks.push(prefix, body);
    }
    const output = new Uint8Array(size);
    let at = 0;
    for (const chunk of chunks) {
      output.set(chunk, at);
      at += chunk.length;
    }
    parts[name] = output;
  }
  if (!changed) return input;
  const output = zipSync(parts, { level: 6 });
  if (output.length > 64 * 1024 * 1024) fail('output byte budget');
  return output;
}
