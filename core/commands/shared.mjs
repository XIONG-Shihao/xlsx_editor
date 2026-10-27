import { FORMAT } from '../format.mjs';
/** Helpers shared by the session coordinator and the command modules. */
import fs from 'node:fs/promises';
import { constants as fileConstants } from 'node:fs';
import { inflateSync } from 'fflate';
import { refusal } from '../refusal-codes.mjs';

export const nativeValue = (value) =>
  JSON.parse(
    JSON.stringify(value, (_key, item) =>
      typeof item === 'number' && (!Number.isFinite(item) || Object.is(item, -0))
        ? { $workbookNumber: Object.is(item, -0) ? '-0' : String(item) }
        : item
    )
  );
export const empty = (view) => ({ ...view, sheets: view.sheets.map((sheet) => ({ ...sheet, cells: {} })) });
export const own = (object, key) => Object.hasOwn(object, key);
export const FACTS_V2 = FORMAT.SCHEMA_USED_RANGE_FACTS_V2;
export const uuidV4 = /^[a-f0-9]{8}-[a-f0-9]{4}-4[a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}$/;
/** Read one private immutable PNG/GLB part, bounding compressed AND raw bytes before
 * allocation/inflation. The part name comes from owner metadata, never IPC. */
export async function boundedObjectMedia(file, part, maxBytes, signal) {
  signal.throwIfAborted();
  const handle = await fs.open(file, fileConstants.O_RDONLY | (fileConstants.O_NOFOLLOW ?? 0));
  const read = async (at, length) => {
    signal.throwIfAborted();
    const bytes = Buffer.alloc(length);
    let offset = 0;
    while (offset < length) {
      signal.throwIfAborted();
      const result = await handle.read(bytes, offset, length - offset, at + offset);
      if (!result.bytesRead) throw refusal('session-malformed', 'Truncated object package');
      offset += result.bytesRead;
    }
    return bytes;
  };
  try {
    const stat = await handle.stat();
    if (!stat.isFile() || stat.size < 22) throw refusal('internal-invalid-field', 'Invalid object package');
    const tail = await read(Math.max(0, stat.size - 65557), Math.min(stat.size, 65557));
    let end = -1;
    for (let i = tail.length - 22; i >= 0; i--)
      if (tail.readUInt32LE(i) === 0x06054b50 && i + 22 + tail.readUInt16LE(i + 20) === tail.length) {
        end = i;
        break;
      }
    if (end < 0 || tail.readUInt16LE(end + 4) || tail.readUInt16LE(end + 6))
      throw refusal('session-unsupported', 'Unsupported object ZIP package');
    const count = tail.readUInt16LE(end + 10),
      size = tail.readUInt32LE(end + 12),
      start = tail.readUInt32LE(end + 16),
      endPosition = stat.size - tail.length + end;
    if (count === 65535 || size === 0xffffffff || start + size > endPosition)
      throw refusal('session-unsupported', 'Unsupported object ZIP directory');
    let cursor = start,
      selected;
    const names = new Set();
    for (let i = 0; i < count; i++) {
      const entry = await read(cursor, 46);
      if (entry.readUInt32LE(0) !== 0x02014b50) throw refusal('session-malformed', 'Invalid object ZIP directory');
      const length = 46 + entry.readUInt16LE(28) + entry.readUInt16LE(30) + entry.readUInt16LE(32);
      if (cursor + length > start + size) throw refusal('session-malformed', 'Invalid object ZIP entry bounds');
      const name = (await read(cursor + 46, entry.readUInt16LE(28))).toString('utf8');
      if (names.has(name)) throw refusal('session-duplicate', 'Duplicate object ZIP part');
      names.add(name);
      if (name === part)
        selected = {
          raw: entry.readUInt32LE(24),
          compressed: entry.readUInt32LE(20),
          local: entry.readUInt32LE(42),
          flags: entry.readUInt16LE(8),
          method: entry.readUInt16LE(10),
          crc: entry.readUInt32LE(16),
        };
      cursor += length;
    }
    if (!selected) throw refusal('internal-missing-field', 'Object media is missing');
    const { raw, compressed, local, flags, method } = selected;
    if (
      raw > maxBytes ||
      compressed > maxBytes + 65536 ||
      flags & 1 ||
      ![0, 8].includes(method) ||
      (method === 0 && raw !== compressed) ||
      local + 30 > start
    )
      throw refusal('session-limit', 'Object media exceeds its bound or uses unsupported ZIP data');
    const header = await read(local, 30);
    if (header.readUInt32LE(0) !== 0x04034b50 || header.readUInt16LE(8) !== method || header.readUInt16LE(6) !== flags)
      throw refusal('internal-mismatch', 'Object ZIP local header disagrees with its directory');
    const nameLength = header.readUInt16LE(26),
      dataStart = local + 30 + nameLength + header.readUInt16LE(28);
    if (dataStart + compressed > start || (await read(local + 30, nameLength)).toString('utf8') !== part)
      throw refusal('session-malformed', 'Invalid object ZIP local part bounds');
    const packed = await read(dataStart, compressed),
      bytes = method === 0 ? packed : inflateSync(packed, { out: new Uint8Array(raw + 1) });
    if (bytes.length !== raw) throw refusal('session-unsupported', 'Object media inflated outside its declared size');
    let crc = 0xffffffff;
    for (const byte of bytes) {
      crc ^= byte;
      for (let bit = 0; bit < 8; bit++) crc = crc & 1 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1;
    }
    if ((crc ^ 0xffffffff) >>> 0 !== selected.crc)
      throw refusal('internal-mismatch', 'Object ZIP media checksum mismatch');
    signal.throwIfAborted();
    return bytes;
  } finally {
    await handle.close();
  }
}
/** Does this one-command list belong to a module that handles these command types or operations? */
export function matchesHandles(handles, commands) {
  if (commands.length !== 1) return false;
  const [command] = commands;
  return handles.type ? handles.type.includes(command.type) : handles.operation.includes(command?.operation);
}
