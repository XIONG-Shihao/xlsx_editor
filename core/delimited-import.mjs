import { FORMAT } from './format.mjs';
/** Portable one-shot delimited import. Keep one scalar and one page, never a
 * document string or row/cell matrix. Hosts supply immutable byte sources. */
import { sha256 } from '@noble/hashes/sha2.js';
import { bytesToHex } from '@noble/hashes/utils.js';
import { csvEncoding } from './csv-codec.mjs';
import {
  PASTE_BYTE_LIMIT,
  PASTE_POSITION_LIMIT,
  PASTE_PAGE_CELLS,
  PASTE_PAGE_BYTES,
  pasteBytes,
  pasteId,
  validatePasteState,
  validatePastePlan,
  validatePasteCells,
  validatePasteDescriptor,
} from './paste-protocol.mjs';
import { refusal } from './refusal-codes.mjs';

export const DELIMITED_READ_BYTES = 256 * 1024;
export const DELIMITED_IMPORT_SCHEMA = FORMAT.SCHEMA_DELIMITED_FILE_IMPORT_V1;
const delimiters = [',', ';', '\t'];
const encodingError = () =>
  refusal(
    'csv-unsupported',
    'Delimited file encoding is unsupported or malformed. Use valid UTF-8 or BOM-marked UTF-16; no characters were replaced.'
  );
async function check(source, signal) {
  signal?.throwIfAborted();
  await source.check?.();
  signal?.throwIfAborted();
}
function sourceInfo(source) {
  if (
    !source ||
    !pasteId(source.id) ||
    typeof source.name !== 'string' ||
    !source.name ||
    source.name.length > 255 ||
    /[\u0000/\\]/.test(source.name) ||
    !/\.(csv|tsv)$/i.test(source.name) ||
    !Number.isSafeInteger(source.size) ||
    source.size < 0 ||
    source.size > PASTE_BYTE_LIMIT ||
    typeof source.chunks !== 'function'
  )
    throw refusal('internal-invalid-field', 'Invalid bounded immutable delimited source');
  return { id: source.id, name: source.name, bytes: source.size };
}

/** SHA covers every raw byte, including BOM, independently of decoded fields. */
async function* decoded(source, { signal, proof, progress = () => {} } = {}) {
  sourceInfo(source);
  let bytes = 0,
    head = new Uint8Array(0),
    decoder,
    encoding,
    bom,
    start,
    emptyChunks = 0;
  const hash = sha256.create();
  try {
    for await (const input of source.chunks()) {
      await check(source, signal);
      if (!(input instanceof Uint8Array) || input.length > DELIMITED_READ_BYTES)
        throw refusal('csv-limit', 'Delimited source exceeded its 256 KiB byte window');
      if (!input.length) {
        if (++emptyChunks > 8) throw refusal('csv-bad-input', 'Delimited source is not advancing');
        continue;
      }
      emptyChunks = 0;
      bytes += input.length;
      if (bytes > source.size || bytes > PASTE_BYTE_LIMIT) throw refusal('csv-stale', 'Delimited source size changed');
      hash.update(input);
      let body = input;
      if (!decoder) {
        const used = Math.min(4 - head.length, input.length),
          prefix = new Uint8Array(head.length + used);
        prefix.set(head);
        prefix.set(input.subarray(0, used), head.length);
        head = prefix;
        body = input.subarray(used);
        if (head.length < 4) continue;
        ({ encoding, bom, start } = csvEncoding(head));
        decoder = new TextDecoder(encoding, { fatal: true, ignoreBOM: true });
        let text;
        try {
          text = decoder.decode(head.subarray(start), { stream: true });
        } catch {
          throw encodingError();
        }
        if (text) yield text;
      }
      let text;
      try {
        text = decoder.decode(body, { stream: true });
      } catch {
        throw encodingError();
      }
      if (text) yield text;
      progress(bytes, source.size);
    }
    await check(source, signal);
    if (bytes !== source.size) throw refusal('csv-refused', 'Delimited source ended before its declared size');
    if (!decoder) {
      ({ encoding, bom, start } = csvEncoding(head));
      decoder = new TextDecoder(encoding, { fatal: true, ignoreBOM: true });
      let text;
      try {
        text = decoder.decode(head.subarray(start), { stream: true });
      } catch {
        throw encodingError();
      }
      if (text) yield text;
    }
    let tail;
    try {
      tail = decoder.decode();
    } catch {
      throw encodingError();
    }
    if (tail) yield tail;
    await check(source, signal);
    progress(bytes, source.size);
    if (proof) Object.assign(proof, { bytes, sha256: bytesToHex(hash.digest()), encoding, bom });
  } finally {
    hash.destroy();
  }
}

/** Same field-start quotes, doubled quotes, strict post-quote and CRLF semantics
 * as csv-workbook; unlike its row projection this scanner emits one field. */
export async function* delimitedFields(source, delimiter, { signal, proof, progress } = {}) {
  if (!delimiters.includes(delimiter)) throw refusal('csv-unsupported', 'Unsupported delimited separator');
  let row = 0,
    col = 0,
    value = '',
    quoted = false,
    quotePending = false,
    closed = false,
    started = false,
    finishedRows = 0,
    skipLf = false;
  const scalar = () => {
    if (value.length > 32768) throw refusal('csv-limit', 'Delimited field exceeds cell text limits');
  };
  for await (const chunk of decoded(source, { signal, proof, progress })) {
    for (let i = 0; i < chunk.length; i++) {
      const c = chunk[i];
      if (skipLf) {
        skipLf = false;
        if (c === '\n') continue;
      }
      if (quotePending) {
        quotePending = false;
        if (c === '"') {
          value += '"';
          quoted = true;
          scalar();
          continue;
        }
        quoted = false;
        closed = true;
      }
      if (quoted) {
        if (c === '"') quotePending = true;
        else value += c;
        scalar();
        continue;
      }
      if (c === '"' && !started) {
        quoted = true;
        started = true;
      } else if (c === delimiter || c === '\r' || c === '\n') {
        yield { row, col, text: value };
        value = '';
        closed = false;
        started = false;
        if (c === delimiter) col++;
        else {
          row++;
          col = 0;
          finishedRows++;
          skipLf = c === '\r';
        }
      } else {
        if (closed) throw refusal('csv-malformed', 'Malformed delimited file: text follows a closing quote');
        value += c;
        started = true;
      }
      scalar();
      if (row > PASTE_POSITION_LIMIT || col >= 16384)
        throw refusal('csv-limit', 'Delimited file exceeds workbook position limits');
    }
    await check(source, signal);
  }
  if (quoted && !quotePending) throw refusal('csv-malformed', 'Malformed delimited file: unterminated quoted field');
  if (started || value || col || !finishedRows) yield { row, col, text: value };
}

/** Auto is only a declared first-record separator heuristic, matching the CSV
 * reader's comma-first tie policy. It adds a bounded prefix scan, not a buffer. */
async function detectDelimiter(source, { signal } = {}) {
  const counts = { ',': 0, ';': 0, '\t': 0 };
  let quoted = false,
    pendingQuote = false;
  for await (const text of decoded(source, { signal }))
    for (const c of text) {
      if (pendingQuote) {
        pendingQuote = false;
        if (c === '"') continue;
        quoted = false;
      }
      if (c === '"') {
        if (quoted) pendingQuote = true;
        else quoted = true;
      } else if (!quoted) {
        if (c === '\r' || c === '\n')
          return delimiters.reduce((best, key) => (counts[key] > counts[best] ? key : best), ',');
        if (Object.hasOwn(counts, c)) counts[c]++;
      }
    }
  return delimiters.reduce((best, key) => (counts[key] > counts[best] ? key : best), ',');
}

export async function inspectDelimitedSource(
  source,
  { sheet, start, delimiter = 'auto', signal, onProgress = () => {} }
) {
  const info = sourceInfo(source);
  validatePastePlan({ shape: { sheet, start, rows: 1, columns: 1 }, count: 1, policy: 'text' });
  await check(source, signal);
  if (delimiter === 'auto') delimiter = await detectDelimiter(source, { signal });
  if (!delimiters.includes(delimiter)) throw refusal('csv-unsupported', 'Unsupported delimited separator');
  let rows = 0,
    columns = 0,
    count = 0;
  const proof = {};
  for await (const field of delimitedFields(source, delimiter, {
    signal,
    proof,
    progress: (done, total) => onProgress({ phase: 'inspecting-file', completed: done, total }),
  })) {
    rows = field.row + 1;
    columns = Math.max(columns, field.col + 1);
    count++;
    if (rows * columns > PASTE_POSITION_LIMIT) throw refusal('csv-limit', 'Delimited file exceeds 1,000,000 positions');
    validatePasteCells([[0, field.text]], { sheet, start, rows: 1, columns: 1 });
  }
  await check(source, signal);
  if (source.sha256 && source.sha256 !== proof.sha256)
    throw refusal('csv-stale', 'Delimited snapshot hash changed before inspection');
  return {
    schema: DELIMITED_IMPORT_SCHEMA,
    source: { ...info, sha256: proof.sha256 },
    encoding: proof.encoding,
    bom: proof.bom,
    delimiter,
    ...validatePastePlan({ shape: { sheet, start, rows, columns }, count, policy: 'text' }),
  };
}
function inspectionFor(source, inspection) {
  const info = sourceInfo(source);
  if (
    !inspection ||
    inspection.schema !== DELIMITED_IMPORT_SCHEMA ||
    inspection.source?.id !== info.id ||
    inspection.source?.name !== info.name ||
    inspection.source?.bytes !== info.bytes ||
    !/^[a-f0-9]{64}$/.test(inspection.source.sha256) ||
    !delimiters.includes(inspection.delimiter) ||
    !['utf-8', 'utf-16le', 'utf-16be'].includes(inspection.encoding) ||
    typeof inspection.bom !== 'boolean' ||
    inspection.policy !== 'text'
  )
    throw refusal('csv-stale', 'Delimited inspection belongs to another source');
  return validatePastePlan({ shape: inspection.shape, count: inspection.count, policy: inspection.policy });
}
export async function* delimitedImportPages(source, inspection, { signal, onProgress = () => {} } = {}) {
  const plan = inspectionFor(source, inspection),
    proof = {};
  let cells = [],
    bytes = 0,
    sequence = 0,
    count = 0,
    rows = 0,
    columns = 0,
    previous = -1,
    totalBytes = 0;
  for await (const field of delimitedFields(source, inspection.delimiter, { signal, proof })) {
    const entry = [field.row * plan.shape.columns + field.col, field.text],
      size = pasteBytes(JSON.stringify(entry) + '\n');
    if (field.col >= plan.shape.columns || field.row >= plan.shape.rows || ++count > plan.count)
      throw refusal('csv-stale', 'Delimited shape changed after inspection');
    if (cells.length && (cells.length >= PASTE_PAGE_CELLS || bytes + size > PASTE_PAGE_BYTES)) {
      validatePasteCells(cells, plan.shape, previous);
      previous = cells.at(-1)[0];
      await check(source, signal);
      yield { sequence: sequence++, cells };
      cells = [];
      bytes = 0;
    }
    if (size > PASTE_PAGE_BYTES || (totalBytes += size) > PASTE_BYTE_LIMIT)
      throw refusal('csv-limit', 'Delimited attachment exceeds its 64 MiB journal budget');
    cells.push(entry);
    bytes += size;
    rows = field.row + 1;
    columns = Math.max(columns, field.col + 1);
    onProgress({ phase: 'reading-file', completed: count, total: plan.count });
  }
  await check(source, signal);
  if (
    count !== plan.count ||
    rows !== plan.shape.rows ||
    columns !== plan.shape.columns ||
    proof.bytes !== inspection.source.bytes ||
    proof.sha256 !== inspection.source.sha256 ||
    proof.encoding !== inspection.encoding ||
    proof.bom !== inspection.bom
  )
    throw refusal('csv-stale', 'Delimited source changed after inspection');
  // Hold the final page until the raw hash and complete shape fence succeeds.
  if (cells.length) {
    validatePasteCells(cells, plan.shape, previous);
    yield { sequence, cells };
  }
}

/** Uploads are backpressured by one awaited acknowledgement. Adoption is left
 * to the host's existing critical/cancellable descriptor apply transaction. */
export async function stageDelimitedImport(
  source,
  inspection,
  { state, signal, checkCurrent = () => {}, sink, onProgress = () => {} }
) {
  const plan = inspectionFor(source, inspection);
  validatePasteState(state);
  let id,
    sealed = false;
  if (!sink || ['begin', 'upload', 'seal', 'cancel', 'prune'].some((key) => typeof sink[key] !== 'function'))
    throw refusal('csv-bad-input', 'Delimited import requires a trusted bounded staging owner');
  const guard = async () => {
    await check(source, signal);
    await checkCurrent();
    signal?.throwIfAborted();
  };
  try {
    await guard();
    const begin = await sink.begin({ state, ...plan });
    id = begin?.id ?? begin?.uploadId;
    if (!pasteId(id) || (begin.sequence ?? begin.nextSequence ?? 0) !== 0)
      throw refusal('csv-refused', 'Delimited staging returned an invalid upload identity or sequence');
    let count = 0,
      sequence = 0;
    for await (const page of delimitedImportPages(source, inspection, { signal })) {
      await guard();
      const receipt = await sink.upload({ id, state, ...page });
      sequence++;
      count += page.cells.length;
      if (
        (receipt?.sequence ?? receipt?.nextSequence) !== sequence ||
        (receipt?.count ?? receipt?.accepted) !== count ||
        (receipt.id ?? id) !== id
      )
        throw refusal('csv-bad-input', 'Delimited upload acknowledgement did not match its bounded page');
      onProgress({ phase: 'uploading-file', completed: count, total: plan.count });
      await guard();
    }
    await guard();
    const result = await sink.seal({ id, state });
    sealed = true;
    const descriptor = validatePasteDescriptor(result?.descriptor ?? result);
    if (
      descriptor.id !== id ||
      descriptor.count !== plan.count ||
      descriptor.policy !== 'text' ||
      JSON.stringify(descriptor.shape) !== JSON.stringify(plan.shape)
    )
      throw refusal('csv-bad-input', 'Delimited sealed plan did not match inspection');
    await guard();
    return {
      descriptor,
      source: inspection.source,
      delimiter: inspection.delimiter,
      encoding: inspection.encoding,
      bom: inspection.bom,
    };
  } catch (error) {
    if (id) {
      try {
        if (sealed) await sink.prune({ state });
        else
          try {
            await sink.cancel({ id, state });
          } catch {
            await sink.prune({ state });
          }
      } catch {
        /* Cleanup refusal never permits deletion of accepted attachments. */
      }
    }
    throw error;
  }
}

/** Native browser File/Blob snapshots expose no path and avoid arrayBuffer() on
 * the complete object. The same raw-byte digest fence also covers second reads. */
export function createBlobDelimitedSource(
  blob,
  { name = blob?.name ?? 'import.csv', chunkBytes = 64 * 1024, signal } = {}
) {
  if (
    !blob ||
    typeof blob.slice !== 'function' ||
    !Number.isSafeInteger(blob.size) ||
    !Number.isInteger(chunkBytes) ||
    chunkBytes < 1 ||
    chunkBytes > DELIMITED_READ_BYTES
  )
    throw refusal('csv-refused', 'A bounded user-selected File/Blob is required');
  const id = globalThis.crypto.randomUUID(),
    size = blob.size;
  let closed = false;
  sourceInfo({ id, name, size, chunks() {} });
  const check = () => {
    signal?.throwIfAborted();
    if (closed || blob.size !== size)
      throw refusal('session-closed', 'Delimited browser snapshot is closed or changed');
  };
  return {
    id,
    name,
    size,
    check,
    async *chunks() {
      check();
      for (let at = 0; at < size; at += chunkBytes) {
        check();
        const stop = Math.min(at + chunkBytes, size),
          bytes = new Uint8Array(await blob.slice(at, stop).arrayBuffer());
        check();
        if (bytes.length !== stop - at) throw refusal('csv-stale', 'Delimited browser snapshot changed');
        yield bytes;
      }
    },
    async dispose() {
      closed = true;
    },
  };
}
