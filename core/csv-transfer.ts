import { gunzipSync, strFromU8 } from 'fflate';
import type { CsvDocumentView } from './contracts.js';
import { refusal } from './refusal-codes.mjs';

/** Large CSV source text crosses worker/main/bridge boundaries compressed.
 * New large transfers decode only requested blocks after the worker retires. */
export type CsvTransferPayload =
  | { transport: 'csv-gzip-v1'; compressedText: Uint8Array }
  | { transport: 'csv-gzip-blocks-v2'; blocks: Uint8Array[]; textOffsets: Float64Array };
export type CsvTransfer = Omit<CsvDocumentView, 'text'> & CsvTransferPayload;
/** Source offsets are UTF-16 positions, identical to the raw CSV index. Only two
 * decoded blocks survive; sparse edits/history share this immutable source. */
function blockText(blocks: Uint8Array[], offsets: Float64Array): CsvDocumentView['text'] {
  if (offsets.length !== blocks.length + 1 || offsets[0] !== 0)
    throw refusal('internal-invalid-field', 'Invalid CSV block offsets');
  for (let i = 1; i < offsets.length; i++)
    if (
      !Number.isSafeInteger(offsets[i]) ||
      offsets[i] <= offsets[i - 1] ||
      offsets[i] - offsets[i - 1] > 1024 * 1024 + 1
    )
      throw refusal('internal-invalid-field', 'Invalid CSV block size');
  const length = offsets[offsets.length - 1],
    cache = new Map<number, string>();
  const block = (i: number): string => {
    let text = cache.get(i);
    if (text === undefined) {
      const bytes = gunzipSync(blocks[i]);
      text = strFromU8(bytes);
      if (text.length !== offsets[i + 1] - offsets[i]) throw refusal('internal-mismatch', 'CSV block length mismatch');
    }
    cache.delete(i);
    cache.set(i, text);
    while (cache.size > 2) cache.delete(cache.keys().next().value!);
    return text;
  };
  const position = (n: number): number => {
    const integer = Math.trunc(n) || 0;
    return integer < 0 ? Math.max(length + integer, 0) : Math.min(integer, length);
  };
  return {
    length,
    slice(start = 0, end = length): string {
      start = position(start);
      end = position(end);
      if (end <= start) return '';
      let low = 0,
        high = blocks.length;
      while (low + 1 < high) {
        const mid = (low + high) >>> 1;
        if (offsets[mid] <= start) low = mid;
        else high = mid;
      }
      const pieces: string[] = [];
      for (let i = low; i < blocks.length && offsets[i] < end; i++)
        pieces.push(block(i).slice(Math.max(0, start - offsets[i]), end - offsets[i]));
      const result = pieces.join('');
      // A short sliced string can retain its entire decoded block in V8. Detach
      // single-block slices so row/cell caches cannot pin evicted source blocks.
      // JSON preserves even a requested slice ending inside a surrogate pair.
      return pieces.length === 1 ? (JSON.parse(JSON.stringify(result)) as string) : result;
    },
  };
}
export function decodeCsvTransfer(value: CsvDocumentView | CsvTransfer): CsvDocumentView {
  if (!('transport' in value)) return value;
  if (value.transport === 'csv-gzip-blocks-v2') {
    const { transport: _transport, blocks, textOffsets, ...metadata } = value;
    return { ...metadata, text: blockText(blocks, textOffsets) };
  }
  if (value.transport !== 'csv-gzip-v1') throw refusal('csv-unsupported', 'Unsupported CSV transport');
  const { transport: _transport, compressedText, ...metadata } = value;
  return { ...metadata, text: strFromU8(gunzipSync(compressedText)) };
}
