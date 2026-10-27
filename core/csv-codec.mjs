/** CSV has no implicit legacy code page. Preserve a declared Unicode BOM and
 * refuse undecodable input instead of replacing stored characters. */
import { refusal } from './refusal-codes.mjs';

export function csvEncoding(input) {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  if (
    (bytes[0] === 0xff && bytes[1] === 0xfe && bytes[2] === 0 && bytes[3] === 0) ||
    (bytes[0] === 0 && bytes[1] === 0 && bytes[2] === 0xfe && bytes[3] === 0xff)
  )
    throw refusal(
      'csv-unsupported',
      'CSV encoding UTF-32 is unsupported. Convert to UTF-8 or BOM-marked UTF-16 without replacing characters.'
    );
  let encoding = 'utf-8',
    bom = false,
    start = 0;
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    bom = true;
    start = 3;
  } else if (bytes[0] === 0xff && bytes[1] === 0xfe) {
    encoding = 'utf-16le';
    bom = true;
    start = 2;
  } else if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    encoding = 'utf-16be';
    bom = true;
    start = 2;
  }
  return { encoding, bom, start };
}
export function decodeCsvBytes(input) {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  const { encoding, bom, start } = csvEncoding(bytes);
  try {
    return {
      text: new TextDecoder(encoding, { fatal: true, ignoreBOM: true }).decode(bytes.subarray(start)),
      encoding,
      bom,
    };
  } catch {
    throw refusal(
      'csv-unsupported',
      'CSV encoding is unsupported or malformed. Use valid UTF-8 or BOM-marked UTF-16; no characters were replaced.'
    );
  }
}

export function encodeCsvText(text, encoding = 'utf-8', bom = false) {
  if (typeof text !== 'string' || /[\uD800-\uDFFF]/u.test(text))
    throw refusal('csv-malformed', 'CSV text contains an unpaired Unicode surrogate');
  if (encoding === 'utf-8') {
    const body = new TextEncoder().encode(text);
    if (!bom) return body;
    const bytes = new Uint8Array(body.length + 3);
    bytes.set([0xef, 0xbb, 0xbf]);
    bytes.set(body, 3);
    return bytes;
  }
  if (!['utf-16le', 'utf-16be'].includes(encoding)) throw refusal('csv-unsupported', 'Unsupported CSV encoding');
  const offset = bom ? 2 : 0,
    bytes = new Uint8Array(text.length * 2 + offset),
    little = encoding === 'utf-16le';
  if (bom) bytes.set(little ? [0xff, 0xfe] : [0xfe, 0xff]);
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i),
      at = offset + i * 2;
    bytes[at] = little ? code & 255 : code >>> 8;
    bytes[at + 1] = little ? code >>> 8 : code & 255;
  }
  return bytes;
}
