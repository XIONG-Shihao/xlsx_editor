/** Dedicated bounded metadata ZIP reader. Worksheet and unrelated payloads never inflate. */
import fs from 'node:fs';
import { inflateRawSync } from 'node:zlib';
import { refusal } from './refusal-codes.mjs';
export const NOTE_XML_PART_BYTES = 16 * 1024 * 1024;
export const NOTE_XML_TOTAL_BYTES = 64 * 1024 * 1024;
const utf8 = new TextDecoder('utf-8', { fatal: true });
const crcTable = Uint32Array.from({ length: 256 }, (_, n) => {
  for (let i = 0; i < 8; i++) n = n & 1 ? 0xedb88320 ^ (n >>> 1) : n >>> 1;
  return n >>> 0;
});
const crc32 = (bytes) => {
  let crc = 0xffffffff;
  for (const value of bytes) crc = crcTable[(crc ^ value) & 255] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
};
const fail = (message, code = 'malformed-package') => {
  const error = refusal('package-malformed', message);
  error.noteMembershipCode = code;
  throw error;
};
const resource = (message) => fail(message, 'resource-limit');
const partName = (name) =>
  typeof name === 'string' &&
  name.length &&
  name.length <= 1024 &&
  !name.startsWith('/') &&
  !name.includes('\\') &&
  !/[\0?#]/.test(name) &&
  name.split('/').every((part) => part && part !== '.' && part !== '..');
export async function openNoteZipFile(file, guard = () => {}) {
  guard();
  const handle = await fs.promises.open(file, fs.constants.O_RDONLY | (fs.constants.O_NOFOLLOW ?? 0));
  let closed = false;
  const close = async () => {
    if (!closed) {
      closed = true;
      await handle.close();
    }
  };
  const read = async (at, length) => {
    guard();
    if (closed) fail('Note metadata ZIP is closed.');
    if (!Number.isSafeInteger(at) || !Number.isSafeInteger(length) || at < 0 || length < 0)
      fail('Invalid note metadata ZIP bounds.');
    const bytes = Buffer.alloc(length);
    let offset = 0;
    while (offset < length) {
      guard();
      const result = await handle.read(bytes, offset, length - offset, at + offset);
      guard();
      if (!result.bytesRead) fail('Truncated note metadata ZIP.');
      offset += result.bytesRead;
    }
    return bytes;
  };
  try {
    guard();
    const stat = await handle.stat();
    guard();
    if (!stat.isFile() || !Number.isSafeInteger(stat.size) || stat.size < 22) fail('Invalid note metadata ZIP.');
    const tail = await read(Math.max(0, stat.size - 65557), Math.min(stat.size, 65557));
    let end = -1;
    for (let i = tail.length - 22; i >= 0; i--)
      if (tail.readUInt32LE(i) === 0x06054b50 && i + 22 + tail.readUInt16LE(i + 20) === tail.length) {
        end = i;
        break;
      }
    if (
      end < 0 ||
      tail.readUInt16LE(end + 4) ||
      tail.readUInt16LE(end + 6) ||
      tail.readUInt16LE(end + 8) !== tail.readUInt16LE(end + 10)
    )
      fail('Unsupported note metadata ZIP directory.');
    const count = tail.readUInt16LE(end + 10),
      size = tail.readUInt32LE(end + 12),
      offset = tail.readUInt32LE(end + 16),
      endPosition = stat.size - tail.length + end;
    if (count === 65535 || size === 0xffffffff || offset === 0xffffffff || offset + size !== endPosition)
      fail('Unsupported note metadata ZIP directory.');
    if (size > NOTE_XML_PART_BYTES) resource('Note metadata ZIP directory exceeds its 16 MiB resource limit.');
    const entries = new Map(),
      intervals = [];
    let cursor = offset;
    for (let i = 0; i < count; i++) {
      const central = await read(cursor, 46);
      if (central.readUInt32LE(0) !== 0x02014b50) fail('Invalid note metadata ZIP entry.');
      const flags = central.readUInt16LE(8),
        method = central.readUInt16LE(10),
        compressed = central.readUInt32LE(20),
        raw = central.readUInt32LE(24),
        nameLength = central.readUInt16LE(28),
        length = 46 + nameLength + central.readUInt16LE(30) + central.readUInt16LE(32),
        local = central.readUInt32LE(42),
        crc = central.readUInt32LE(16);
      if (
        cursor + length > offset + size ||
        local + 30 > offset ||
        flags & ~0x808 ||
        ![0, 8].includes(method) ||
        central.readUInt16LE(34) ||
        [compressed, raw, local].includes(0xffffffff)
      )
        fail('Unsupported note metadata ZIP entry.');
      const nameBytes = await read(cursor + 46, nameLength),
        name = utf8.decode(nameBytes);
      if (!partName(name.endsWith('/') ? name.slice(0, -1) : name)) fail('Invalid note metadata ZIP part name.');
      if (entries.has(name)) fail('Duplicate note metadata ZIP member.');
      const header = await read(local, 30);
      if (
        header.readUInt32LE(0) !== 0x04034b50 ||
        header.readUInt16LE(6) !== flags ||
        header.readUInt16LE(8) !== method ||
        header.readUInt16LE(26) !== nameLength
      )
        fail('Conflicting note metadata ZIP local header.');
      const localName = await read(local + 30, nameLength);
      if (!localName.equals(nameBytes)) fail('Conflicting note metadata ZIP member name.');
      if (
        !(flags & 8) &&
        (header.readUInt32LE(14) !== crc || header.readUInt32LE(18) !== compressed || header.readUInt32LE(22) !== raw)
      )
        fail('Conflicting note metadata ZIP sizes.');
      const start = local + 30 + nameLength + header.readUInt16LE(28);
      if (start + compressed > offset) fail('Invalid note metadata ZIP member bounds.');
      entries.set(name, { start, compressed, raw, method, crc });
      intervals.push([local, start + compressed]);
      cursor += length;
    }
    if (cursor !== offset + size) fail('Conflicting note metadata ZIP directory size.');
    intervals.sort((a, b) => a[0] - b[0]);
    for (let i = 1; i < intervals.length; i++)
      if (intervals[i][0] < intervals[i - 1][1]) fail('Overlapping note metadata ZIP members.');
    let expanded = 0,
      compressedTotal = 0;
    const decoded = new Map();
    return {
      names: [...entries.keys()],
      has: (name) => entries.has(name),
      close,
      read: async (name) => {
        guard();
        if (!entries.has(name)) fail('Missing note metadata part.', 'missing-part');
        if (decoded.has(name)) return decoded.get(name);
        const entry = entries.get(name);
        // Both declared selected sizes are checked before either allocation or inflation.
        if (entry.raw > NOTE_XML_PART_BYTES || entry.compressed > NOTE_XML_PART_BYTES)
          resource('Selected note metadata part exceeds its 16 MiB compressed or expanded resource limit.');
        if (expanded + entry.raw > NOTE_XML_TOTAL_BYTES || compressedTotal + entry.compressed > NOTE_XML_TOTAL_BYTES)
          resource('Selected note metadata exceeds its 64 MiB compressed or expanded resource limit.');
        expanded += entry.raw;
        compressedTotal += entry.compressed;
        const packed = await read(entry.start, entry.compressed);
        guard();
        let bytes;
        try {
          bytes = entry.method === 0 ? packed : inflateRawSync(packed, { maxOutputLength: Math.max(1, entry.raw) });
        } catch {
          fail('Invalid or oversized compressed note metadata.');
        }
        guard();
        if (bytes.length !== entry.raw || crc32(bytes) !== entry.crc)
          fail('Note metadata ZIP size or checksum mismatch.');
        decoded.set(name, bytes);
        return bytes;
      },
    };
  } catch (error) {
    await close();
    throw error;
  }
}
