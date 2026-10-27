/** Bounded Node CSV edits. The input must be a caller-owned immutable snapshot;
 * bytes are rehashed while copying so a changed snapshot cannot be accepted. */
import { openSync, closeSync, readSync, writeSync, fstatSync, statSync, fsyncSync, unlinkSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { readCsvFileIndex, checkCsvFileIndexSource } from './csv-file.mjs';
import { applyCsvCommands, csvTextEdits } from './csv-workbook.mjs';
import { encodeCsvText } from './csv-codec.mjs';
import { refusal } from './refusal-codes.mjs';

export function editCsvFile(source, commands, expectedRevision, options, destination) {
  const index = readCsvFileIndex(source, options);
  if (expectedRevision && expectedRevision.replace(/^sha256:/, '') !== index.revision)
    throw refusal('csv-stale', 'Conflict: CSV revision is stale');
  const edited = Array.isArray(commands) && commands.length === 0 ? index : applyCsvCommands(index, commands).index;
  return { revision: writeCsvIndex(edited, source, destination, index.revision) };
}

export function writeCsvIndex(index, source, destination, expectedRevision = index.revision, { maxOutputBytes } = {}) {
  if (maxOutputBytes !== undefined && (!Number.isSafeInteger(maxOutputBytes) || maxOutputBytes < 0))
    throw refusal('internal-invalid-field', 'Invalid CSV output disk budget');
  checkCsvFileIndexSource(index, source);
  const prefix = index.bom ? (index.encoding === 'utf-8' ? 3 : 2) : 0;
  let character = 0,
    offset = prefix;
  function byteOffset(next) {
    if (next < character || next > index.text.length) throw refusal('csv-malformed', 'Invalid CSV source span.');
    if (typeof index.text.byteOffset === 'function') return index.text.byteOffset(next);
    if (index.encoding !== 'utf-8') {
      offset += (next - character) * 2;
      character = next;
      return offset;
    }
    while (character < next) {
      let end = Math.min(next, character + 1024 * 1024);
      if (end < next) {
        const last = index.text.slice(end - 1, end).charCodeAt(0);
        if (last >= 0xd800 && last <= 0xdbff) end++;
      }
      offset += Buffer.byteLength(index.text.slice(character, end), 'utf8');
      character = end;
    }
    return offset;
  }
  const patches = csvTextEdits(index).map((p) => ({
    start: byteOffset(p.start),
    end: byteOffset(p.end),
    bytes: encodeCsvText(p.text, index.encoding, false),
  }));
  const input = openSync(source, 'r');
  let output,
    created = false;
  try {
    const sourceInfo = fstatSync(input),
      sourceStamp = stamp(fstatSync(input, { bigint: true }));
    if (!sourceInfo.isFile()) throw refusal('csv-bad-input', 'CSV source must be a regular file');
    checkCsvFileIndexSource(index, source, input);
    const size = sourceInfo.size;
    let cursor = 0,
      outputBytes = size;
    for (const p of patches) {
      if (p.start < cursor || p.end < p.start || p.end > size)
        throw refusal('internal-invalid-field', 'Invalid CSV byte patch.');
      cursor = p.end;
    }
    for (const p of patches) outputBytes += p.bytes.length - (p.end - p.start);
    if (!Number.isSafeInteger(outputBytes) || (maxOutputBytes !== undefined && outputBytes > maxOutputBytes))
      throw refusal('csv-limit', 'CSV output exceeds the session disk budget');
    output = openSync(destination, 'wx', 0o600);
    created = true;
    const sourceHash = createHash('sha256'),
      outputHash = createHash('sha256'),
      buffer = Buffer.allocUnsafe(1024 * 1024);
    const write = (bytes) => {
      outputHash.update(bytes);
      for (let written = 0; written < bytes.length;) {
        const n = writeSync(output, bytes, written, bytes.length - written);
        if (!n) throw refusal('csv-unsupported', 'Incomplete CSV write');
        written += n;
      }
    };
    // Dense edits otherwise issue one write/hash update for each tiny source or
    // replacement span. Keep a bounded buffer for those spans; large unchanged
    // blocks still go directly to the file without another full-block copy.
    const small = Buffer.allocUnsafe(64 * 1024);
    let buffered = 0;
    const flush = () => {
      if (buffered) {
        write(small.subarray(0, buffered));
        buffered = 0;
      }
    };
    const emit = (bytes) => {
      if (bytes.length >= small.length) {
        flush();
        write(bytes);
        return;
      }
      for (let at = 0; at < bytes.length;) {
        const n = Math.min(bytes.length - at, small.length - buffered);
        small.set(bytes.subarray(at, at + n), buffered);
        buffered += n;
        at += n;
        if (buffered === small.length) flush();
      }
    };
    // cursor is the next source byte to emit. It may be ahead of the current
    // read chunk when a replaced span crosses chunks; every input byte is still
    // hashed, including replaced spans, before the candidate can be accepted.
    let position = 0,
      patch = 0;
    cursor = 0;
    for (let length; (length = readSync(input, buffer, 0, buffer.length, position)) > 0;) {
      const bytes = buffer.subarray(0, length),
        end = position + length;
      sourceHash.update(bytes);
      let at = Math.max(0, cursor - position);
      while (at < length) {
        const next = patches[patch];
        if (!next || next.start >= end) {
          emit(bytes.subarray(at));
          cursor = end;
          break;
        }
        if (next.start > position + at) emit(bytes.subarray(at, next.start - position));
        emit(next.bytes);
        cursor = next.end;
        patch++;
        at = Math.max(at, cursor - position);
      }
      position = end;
    }
    for (; patch < patches.length; patch++) {
      const next = patches[patch];
      if (next.start !== position || next.end !== position)
        throw refusal('csv-stale', 'CSV source changed while saving');
      emit(next.bytes);
    }
    flush();
    if (
      position !== size ||
      (expectedRevision && sourceHash.digest('hex') !== expectedRevision.replace(/^sha256:/, ''))
    )
      throw refusal('csv-stale', 'Workbook recovery base changed');
    if (
      stamp(fstatSync(input, { bigint: true })) !== sourceStamp ||
      stamp(statSync(source, { bigint: true })) !== sourceStamp
    )
      throw refusal('csv-stale', 'CSV source changed while saving');
    checkCsvFileIndexSource(index, source, input);
    fsyncSync(output);
    return outputHash.digest('hex');
  } catch (error) {
    if (output !== undefined) {
      closeSync(output);
      output = undefined;
    }
    if (created) unlinkSync(destination);
    throw error;
  } finally {
    if (output !== undefined) closeSync(output);
    closeSync(input);
  }
}
const stamp = (value) => [value.dev, value.ino, value.size, value.mtimeNs, value.ctimeNs].join(':');
