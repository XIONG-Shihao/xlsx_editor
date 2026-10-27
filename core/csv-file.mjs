/** Node-only CSV opening. Retain compressed blocks and row metadata, never the
 * complete byte buffer or decoded text. Both editor hosts use this component API. */
import { openSync, readSync, closeSync, fstatSync, statSync } from 'node:fs';
import { createHash } from 'node:crypto';
import path from 'node:path';
import { gzipSync } from 'node:zlib';
import { csvEncoding } from './csv-codec.mjs';
import { indexCsvChunks } from './csv-workbook.mjs';
import { refusal } from './refusal-codes.mjs';
const fileSources = new WeakMap();

/** A file-backed index and all its sparse versions belong to exactly one
 * immutable snapshot. Checking the open descriptor and current path catches
 * replacement even when its bytes happen to have the same digest. */
export function checkCsvFileIndexSource(index, source, descriptor) {
  const identity = index.text && typeof index.text === 'object' ? fileSources.get(index.text) : undefined;
  if (!identity) return false;
  const file = source === undefined ? identity.file : path.resolve(source);
  if (
    file !== identity.file ||
    index.rowOffsets !== identity.rowOffsets ||
    index.rowWidths !== identity.rowWidths ||
    index.revision !== identity.revision ||
    index.encoding !== identity.encoding ||
    index.delimiter !== identity.delimiter ||
    index.bom !== identity.bom
  )
    throw refusal('csv-stale', 'CSV index belongs to another source');
  if (
    stampValue(statSync(file, { bigint: true })) !== identity.stamp ||
    (descriptor !== undefined && stamp(descriptor) !== identity.stamp)
  )
    throw refusal('csv-stale', 'CSV source changed; workbook recovery base changed');
  return true;
}

export function readCsvFileTransfer(file, options, expectedRevision) {
  return readCsvFile(file, options, expectedRevision, false);
}
/** Private immutable files can supply bounded text windows without compression
 * or a whole-document string. The caller owns the source file's lifetime. */
export function readCsvFileIndex(file, options, expectedRevision) {
  return readCsvFile(file, options, expectedRevision, true);
}
function readCsvFile(file, { delimiter = 'auto', maxIndexBytes } = {}, expectedRevision, fileBacked) {
  file = path.resolve(file);
  const fd = openSync(file, 'r');
  try {
    if (!fstatSync(fd).isFile()) throw refusal('csv-bad-input', 'CSV source must be a regular file');
    const sourceStamp = stamp(fd);
    const prefix = Buffer.alloc(4),
      count = readSync(fd, prefix, 0, 4, 0);
    const { encoding, bom, start } = csvEncoding(prefix.subarray(0, count));
    function* chunks(digest) {
      // Smaller file-backed decoding windows reduce transient string retention;
      // compressed transport keeps its established block size.
      const buffer = Buffer.allocUnsafe(fileBacked ? 256 * 1024 : 1024 * 1024),
        decoder = new TextDecoder(encoding, { fatal: true, ignoreBOM: true });
      let offset = 0;
      for (let bytes; (bytes = readSync(fd, buffer, 0, buffer.length, offset)) > 0;) {
        digest?.update(buffer.subarray(0, bytes));
        let text;
        try {
          text = decoder.decode(buffer.subarray(offset === 0 ? start : 0, bytes), { stream: true });
        } catch {
          throw refusal(
            'csv-unsupported',
            'CSV encoding is unsupported or malformed. Use valid UTF-8 or BOM-marked UTF-16; no characters were replaced.'
          );
        }
        offset += bytes;
        if (text) yield text;
      }
      try {
        const tail = decoder.decode();
        if (tail) yield tail;
      } catch {
        throw refusal(
          'csv-unsupported',
          'CSV encoding is unsupported or malformed. Use valid UTF-8 or BOM-marked UTF-16; no characters were replaced.'
        );
      }
    }
    if (delimiter === 'auto') {
      const counts = { ',': 0, ';': 0, '\t': 0 };
      let quoted = false;
      header: for (const text of chunks())
        for (const c of text) {
          // Each escaped pair toggles twice, preserving the quoted state.
          if (c === '"') quoted = !quoted;
          else if (!quoted) {
            if (c === '\r' || c === '\n') break header;
            if (Object.hasOwn(counts, c)) counts[c]++;
          }
        }
      delimiter = Object.entries(counts).sort((a, b) => b[1] - a[1])[0][0];
    }
    const digest = createHash('sha256'),
      blocks = [],
      offsets = [0],
      byteOffsets = [start];
    function* compressedChunks() {
      for (const text of chunks(digest)) {
        if (fileBacked)
          byteOffsets.push(
            byteOffsets.at(-1) + (encoding === 'utf-8' ? Buffer.byteLength(text, 'utf8') : text.length * 2)
          );
        else blocks.push(new Uint8Array(gzipSync(text, { level: 1 })));
        offsets.push(offsets.at(-1) + text.length);
        yield text;
      }
    }
    const index = indexCsvChunks(compressedChunks(), { delimiter, encoding, bom, maxIndexBytes }),
      revision = digest.digest('hex');
    if (expectedRevision && revision !== expectedRevision.replace(/^sha256:/, ''))
      throw refusal('csv-stale', 'Workbook recovery base changed');
    if (stamp(fd) !== sourceStamp || stampValue(statSync(file, { bigint: true })) !== sourceStamp)
      throw refusal('csv-stale', 'CSV source changed while indexing');
    if (fileBacked) {
      const text = fileText(file, encoding, offsets, byteOffsets, sourceStamp);
      fileSources.set(text, {
        file,
        stamp: sourceStamp,
        revision,
        rowOffsets: index.rowOffsets,
        rowWidths: index.rowWidths,
        encoding,
        delimiter,
        bom,
      });
      return { ...index, revision, text };
    }
    if (blocks.length <= 1)
      return {
        ...index,
        revision,
        transport: 'csv-gzip-v1',
        compressedText: blocks[0] ?? new Uint8Array(gzipSync('')),
      };
    return { ...index, revision, transport: 'csv-gzip-blocks-v2', blocks, textOffsets: Float64Array.from(offsets) };
  } finally {
    closeSync(fd);
  }
}

function stamp(fd) {
  return stampValue(fstatSync(fd, { bigint: true }));
}
function stampValue(value) {
  return [value.dev, value.ino, value.size, value.mtimeNs, value.ctimeNs].join(':');
}
function fileText(file, encoding, offsets, byteOffsets, sourceStamp) {
  const length = offsets.at(-1),
    cache = new Map();
  let previousPosition = 0,
    previousByte = byteOffsets[0],
    previousBlock = 0;
  function position(value) {
    const integer = Math.trunc(value) || 0;
    return integer < 0 ? Math.max(length + integer, 0) : Math.min(integer, length);
  }
  function blockAt(at) {
    let low = 0,
      high = offsets.length - 1;
    while (low + 1 < high) {
      const mid = (low + high) >>> 1;
      if (offsets[mid] <= at) low = mid;
      else high = mid;
    }
    return low;
  }
  function block(i) {
    let text = cache.get(i);
    if (text === undefined) {
      const bytes = Buffer.allocUnsafe(byteOffsets[i + 1] - byteOffsets[i]),
        fd = openSync(file, 'r');
      try {
        if (stamp(fd) !== sourceStamp) throw refusal('csv-stale', 'CSV source changed while reading');
        for (let count = 0; count < bytes.length;) {
          const read = readSync(fd, bytes, count, bytes.length - count, byteOffsets[i] + count);
          if (!read) throw refusal('csv-stale', 'CSV source changed while reading');
          count += read;
        }
        if (stamp(fd) !== sourceStamp) throw refusal('csv-stale', 'CSV source changed while reading');
      } finally {
        closeSync(fd);
      }
      text = new TextDecoder(encoding, { fatal: true, ignoreBOM: true }).decode(bytes);
      if (text.length !== offsets[i + 1] - offsets[i]) throw refusal('csv-stale', 'CSV source changed while reading');
    }
    cache.delete(i);
    cache.set(i, text);
    while (cache.size > 2) cache.delete(cache.keys().next().value);
    return text;
  }
  return {
    length,
    slice(start = 0, end = length) {
      start = position(start);
      end = position(end);
      if (end <= start) return '';
      const pieces = [];
      for (let i = blockAt(start); i < offsets.length - 1 && offsets[i] < end; i++)
        pieces.push(block(i).slice(Math.max(0, start - offsets[i]), end - offsets[i]));
      const result = pieces.join('');
      // Short row/cell strings must not pin evicted decoded blocks in V8.
      return pieces.length === 1 ? Buffer.from(result, 'utf16le').toString('utf16le') : result;
    },
    byteOffset(at) {
      if (!Number.isSafeInteger(at) || at < 0 || at > length)
        throw refusal('csv-malformed', 'Invalid CSV source span.');
      if (at === length) return byteOffsets.at(-1);
      const i = blockAt(at),
        count = at - offsets[i];
      if (encoding !== 'utf-8') return byteOffsets[i] + count * 2;
      const text = block(i);
      if (
        count > 0 &&
        count < text.length &&
        text.charCodeAt(count - 1) >= 0xd800 &&
        text.charCodeAt(count - 1) <= 0xdbff
      )
        throw refusal('csv-refused', 'CSV byte span splits a Unicode character');
      // Patch coordinates are normally monotonic. Count only the new prefix
      // segment, so dense edits do not rescan up to a full block per cell.
      const reuse = previousBlock === i && at >= previousPosition;
      const from = reuse ? previousPosition - offsets[i] : 0;
      const byte = (reuse ? previousByte : byteOffsets[i]) + Buffer.byteLength(text.slice(from, count), 'utf8');
      previousPosition = at;
      previousByte = byte;
      previousBlock = i;
      return byte;
    },
  };
}
