import { gzipSync } from 'node:zlib';
import type { CsvDocumentView } from './contracts.js';
import type { CsvTransferPayload } from './csv-transfer.js';
import { refusal } from './refusal-codes.mjs';

/** Compress inside the disposable worker using native zlib. The renderer keeps
 * the portable decoder; no Node dependency crosses the preload boundary. */
export function encodeCsvTransfer<T extends Omit<CsvDocumentView, 'sourceSnapshot'>>(
  view: T
): Omit<T, 'text'> & CsvTransferPayload {
  const { text, ...metadata } = view;
  if (typeof text !== 'string') throw refusal('csv-bad-input', 'CSV transfer requires the original decoded source');
  const blockSize = 1024 * 1024;
  if (text.length <= blockSize)
    return { ...metadata, transport: 'csv-gzip-v1', compressedText: gzipSync(text, { level: 1 }) };
  const blocks: Uint8Array[] = [],
    offsets = [0];
  for (let start = 0; start < text.length; ) {
    let end = Math.min(start + blockSize, text.length);
    // Never UTF-8 encode half a surrogate pair at a block boundary.
    if (end < text.length && text.charCodeAt(end - 1) >= 0xd800 && text.charCodeAt(end - 1) <= 0xdbff) end++;
    blocks.push(new Uint8Array(gzipSync(text.slice(start, end), { level: 1 })));
    offsets.push(end);
    start = end;
  }
  return { ...metadata, transport: 'csv-gzip-blocks-v2', blocks, textOffsets: Float64Array.from(offsets) };
}
