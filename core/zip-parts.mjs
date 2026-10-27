/** Portable lazy ZIP members; preserve unchanged compressed payloads verbatim.
 * Resource limits belong to the caller; this is deliberately not a Buffer shim. */
import { inflateSync, zipSync, strFromU8 } from 'fflate';
import { refusal } from './refusal-codes.mjs';
const sources = new WeakMap();
/** Fatal UTF-8 decoding for XML graphs, without a host TextDecoder dependency. */
export function decodeXmlUtf8(bytes) {
  for (let at = 0; at < bytes.length;) {
    const first = bytes[at++];
    if (first < 128) continue;
    let count,
      lower = 128,
      upper = 191;
    if (first >= 194 && first <= 223) count = 1;
    else if (first >= 224 && first <= 239) {
      count = 2;
      if (first === 224) lower = 160;
      if (first === 237) upper = 159;
    } else if (first >= 240 && first <= 244) {
      count = 3;
      if (first === 240) lower = 144;
      if (first === 244) upper = 143;
    } else throw refusal('internal-invalid-field', 'Invalid UTF-8 package metadata');
    if (at + count > bytes.length || bytes[at] < lower || bytes[at] > upper)
      throw refusal('internal-invalid-field', 'Invalid UTF-8 package metadata');
    at++;
    for (let next = 1; next < count; next++, at++)
      if (bytes[at] < 128 || bytes[at] > 191) throw refusal('internal-invalid-field', 'Invalid UTF-8 package metadata');
  }
  const text = strFromU8(bytes);
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}
export function equalBytes(a, b) {
  return !!a && !!b && a.length === b.length && a.every((byte, index) => byte === b[index]);
}
const view = (bytes) => new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
const crcTable = Uint32Array.from({ length: 256 }, (_, index) => {
  let value = index;
  for (let bit = 0; bit < 8; bit++) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
  return value >>> 0;
});
function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) crc = crcTable[(crc ^ byte) & 255] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}
function directory(input, maxExpandedBytes = Infinity) {
  if (!ArrayBuffer.isView(input) || input.BYTES_PER_ELEMENT !== 1)
    throw refusal('package-bad-input', 'ZIP input must be a byte array');
  if (typeof maxExpandedBytes !== 'number' || Number.isNaN(maxExpandedBytes) || maxExpandedBytes < 0)
    throw refusal('package-malformed', 'Invalid ZIP expansion limit');
  const bytes = new Uint8Array(input.buffer, input.byteOffset, input.byteLength),
    data = view(bytes);
  if (bytes.length < 22) throw refusal('package-malformed', 'Invalid ZIP package');
  const u16 = (at) => data.getUint16(at, true),
    u32 = (at) => data.getUint32(at, true);
  let end = -1;
  for (let at = bytes.length - 22; at >= Math.max(0, bytes.length - 65557); at--)
    if (u32(at) === 0x06054b50 && at + 22 + u16(at + 20) === bytes.length) {
      end = at;
      break;
    }
  if (end < 0 || u16(end + 4) || u16(end + 6) || u16(end + 8) !== u16(end + 10))
    throw refusal('package-unsupported', 'Unsupported ZIP package');
  const count = u16(end + 10),
    size = u32(end + 12),
    offset = u32(end + 16);
  if (count === 65535 || size === 0xffffffff || offset + size > end)
    throw refusal('package-unsupported', 'Unsupported ZIP directory');
  const entries = new Map();
  let cursor = offset,
    expanded = 0;
  for (let index = 0; index < count; index++) {
    if (cursor + 46 > offset + size || u32(cursor) !== 0x02014b50)
      throw refusal('package-malformed', 'Invalid ZIP entry');
    const flags = u16(cursor + 8),
      method = u16(cursor + 10),
      compressed = u32(cursor + 20),
      raw = u32(cursor + 24),
      crc = u32(cursor + 16),
      nameLength = u16(cursor + 28);
    const length = 46 + nameLength + u16(cursor + 30) + u16(cursor + 32),
      local = u32(cursor + 42);
    if (
      cursor + length > offset + size ||
      local + 30 > offset ||
      u32(local) !== 0x04034b50 ||
      flags & 1 ||
      ![0, 8].includes(method)
    )
      throw refusal('package-unsupported', 'Unsupported ZIP entry');
    const localNameLength = u16(local + 26),
      start = local + 30 + localNameLength + u16(local + 28);
    if (
      start + compressed > offset ||
      u16(local + 6) !== flags ||
      u16(local + 8) !== method ||
      !equalBytes(
        bytes.subarray(cursor + 46, cursor + 46 + nameLength),
        bytes.subarray(local + 30, local + 30 + localNameLength)
      )
    )
      throw refusal('internal-mismatch', 'ZIP local header disagrees with its directory');
    if (!(flags & 8) && (u32(local + 14) !== crc || u32(local + 18) !== compressed || u32(local + 22) !== raw))
      throw refusal('internal-mismatch', 'ZIP local sizes or CRC disagree with its directory');
    expanded += raw;
    if (expanded > maxExpandedBytes)
      throw refusal('package-limit', 'Expanded document package exceeds its editing limit');
    const name = strFromU8(bytes.subarray(cursor + 46, cursor + 46 + nameLength));
    if (entries.has(name)) throw refusal('package-duplicate', 'Duplicate ZIP member');
    entries.set(name, {
      name,
      method,
      raw,
      crc,
      central: bytes.subarray(cursor, cursor + length),
      local: bytes.subarray(local, start),
      compressed: bytes.subarray(start, start + compressed),
    });
    cursor += length;
  }
  return { entries, comment: bytes.subarray(end + 22) };
}
export function openZipParts(bytes, maxExpandedBytes) {
  const source = directory(bytes, maxExpandedBytes),
    changed = new Map(),
    decoded = new Map(),
    removed = new Set();
  const parts = new Proxy(Object.create(null), {
    get: (_target, key) => {
      if (changed.has(key)) return changed.get(key);
      const entry = source.entries.get(key);
      if (!entry || removed.has(key)) return undefined;
      if (!decoded.has(key)) {
        const data =
          entry.method === 0 ? entry.compressed : inflateSync(entry.compressed, { out: new Uint8Array(entry.raw) });
        if (data.length !== entry.raw) throw refusal('internal-mismatch', 'ZIP member size mismatch');
        if (crc32(data) !== entry.crc) throw refusal('internal-mismatch', 'ZIP member CRC mismatch');
        decoded.set(key, data);
      }
      return decoded.get(key);
    },
    set: (_target, key, value) => {
      changed.set(key, value);
      removed.delete(key);
      return true;
    },
    deleteProperty: (_target, key) => {
      removed.add(key);
      changed.delete(key);
      return true;
    },
    ownKeys: () => [...new Set([...source.entries.keys(), ...changed.keys()])].filter((key) => !removed.has(key)),
    getOwnPropertyDescriptor: (_target, key) =>
      !removed.has(key) && (changed.has(key) || source.entries.has(key))
        ? { enumerable: true, configurable: true }
        : undefined,
  });
  sources.set(parts, { source, changed, removed });
  return parts;
}
/** Inspect a member's declared size without allocating/inflating its payload. */
export function zipPartSize(parts, name) {
  const inventory = sources.get(parts);
  if (!inventory) return parts[name]?.byteLength;
  if (inventory.removed.has(name)) return undefined;
  return inventory.changed.has(name)
    ? inventory.changed.get(name)?.byteLength
    : inventory.source.entries.get(name)?.raw;
}
export function writeZipParts(parts) {
  const { source, changed, removed } = sources.get(parts);
  const replacements = changed.size
    ? directory(zipSync(Object.fromEntries(changed), { mtime: new Date(2000, 0, 1) })).entries
    : new Map();
  const names = Object.keys(parts),
    chunks = [],
    centrals = [];
  let offset = 0,
    centralSize = 0;
  if (names.length >= 65535) throw refusal('package-bad-input', 'ZIP64 output requires a separate writer');
  for (const name of names) {
    if (removed.has(name)) continue;
    const entry = replacements.get(name) ?? source.entries.get(name),
      local = Uint8Array.from(entry.local),
      central = Uint8Array.from(entry.central),
      l = view(local),
      c = view(central);
    // Resolve data descriptors using the authoritative directory fields.
    l.setUint16(6, l.getUint16(6, true) & ~8, true);
    c.setUint16(8, c.getUint16(8, true) & ~8, true);
    local.set(central.subarray(16, 28), 14);
    c.setUint32(42, offset, true);
    chunks.push(local, entry.compressed);
    centrals.push(central);
    offset += local.length + entry.compressed.length;
    centralSize += central.length;
    if (offset > 0xfffffffe || centralSize > 0xfffffffe)
      throw refusal('package-bad-input', 'ZIP64 output requires a separate writer');
  }
  const end = new Uint8Array(22),
    data = view(end);
  data.setUint32(0, 0x06054b50, true);
  data.setUint16(8, centrals.length, true);
  data.setUint16(10, centrals.length, true);
  data.setUint32(12, centralSize, true);
  data.setUint32(16, offset, true);
  data.setUint16(20, source.comment.length, true);
  const output = new Uint8Array(offset + centralSize + 22 + source.comment.length);
  let cursor = 0;
  for (const bytes of [...chunks, ...centrals, end, source.comment]) {
    output.set(bytes, cursor);
    cursor += bytes.length;
  }
  return output;
}
/** Compare lazy preserved members without decoding unrelated media. */
export function sameZipMember(left, right, name) {
  const a = sources.get(left),
    b = sources.get(right);
  if (!a || !b || a.removed.has(name) || b.removed.has(name)) return false;
  if (a.changed.has(name) || b.changed.has(name)) return equalBytes(left[name], right[name]);
  const x = a.source.entries.get(name),
    y = b.source.entries.get(name);
  return (
    !!x && !!y && x.method === y.method && x.raw === y.raw && x.crc === y.crc && equalBytes(x.compressed, y.compressed)
  );
}
