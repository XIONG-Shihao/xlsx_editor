/** File-backed, read-only ZIP directory. Only requested metadata members inflate. */
import fs from 'node:fs';
import { inflateSync } from 'fflate';
import { refusal } from './refusal-codes.mjs';
export function openZipFile(file, maxExpandedBytes = Infinity) {
  const fd = fs.openSync(file, fs.constants.O_RDONLY | (fs.constants.O_NOFOLLOW ?? 0));
  let closed = false;
  const close = () => {
    if (!closed) {
      closed = true;
      fs.closeSync(fd);
    }
  };
  const read = (at, n) => {
    if (closed) throw refusal('session-closed', 'ZIP file is closed');
    const b = Buffer.alloc(n);
    let offset = 0;
    while (offset < n) {
      const got = fs.readSync(fd, b, offset, n - offset, at + offset);
      if (!got) throw refusal('package-malformed', 'Truncated ZIP file');
      offset += got;
    }
    return b;
  };
  try {
    const st = fs.fstatSync(fd);
    if (!st.isFile() || st.size < 22) throw refusal('package-malformed', 'Invalid ZIP package');
    const tail = read(Math.max(0, st.size - 65557), Math.min(st.size, 65557));
    let end = -1;
    for (let i = tail.length - 22; i >= 0; i--)
      if (tail.readUInt32LE(i) === 0x06054b50 && i + 22 + tail.readUInt16LE(i + 20) === tail.length) {
        end = i;
        break;
      }
    if (end < 0 || tail.readUInt16LE(end + 4) || tail.readUInt16LE(end + 6))
      throw refusal('package-unsupported', 'Unsupported ZIP package');
    const count = tail.readUInt16LE(end + 10),
      size = tail.readUInt32LE(end + 12),
      offset = tail.readUInt32LE(end + 16),
      endPosition = st.size - tail.length + end;
    if (count === 65535 || size === 0xffffffff || offset + size > endPosition)
      throw refusal('package-unsupported', 'Unsupported ZIP directory');
    const entries = new Map();
    let cursor = offset,
      expanded = 0;
    for (let i = 0; i < count; i++) {
      const central = read(cursor, 46);
      if (central.readUInt32LE(0) !== 0x02014b50) throw refusal('package-malformed', 'Invalid ZIP entry');
      const flags = central.readUInt16LE(8),
        method = central.readUInt16LE(10),
        compressed = central.readUInt32LE(20),
        raw = central.readUInt32LE(24),
        length = 46 + central.readUInt16LE(28) + central.readUInt16LE(30) + central.readUInt16LE(32),
        local = central.readUInt32LE(42);
      if (cursor + length > offset + size || local + 30 > offset || flags & 1 || ![0, 8].includes(method))
        throw refusal('package-unsupported', 'Unsupported ZIP entry');
      const name = read(cursor + 46, central.readUInt16LE(28)).toString('utf8'),
        header = read(local, 30);
      if (header.readUInt32LE(0) !== 0x04034b50) throw refusal('package-malformed', 'Invalid ZIP local header');
      const start = local + 30 + header.readUInt16LE(26) + header.readUInt16LE(28);
      if (start + compressed > offset) throw refusal('package-malformed', 'Invalid ZIP member bounds');
      expanded += raw;
      if (expanded > maxExpandedBytes)
        throw refusal('package-limit', 'Expanded document package exceeds its editing limit');
      if (entries.has(name)) throw refusal('package-duplicate', 'Duplicate ZIP member');
      entries.set(name, { start, compressed, raw, method });
      cursor += length;
    }
    const decoded = new Map();
    const parts = new Proxy(Object.create(null), {
      ownKeys: () => [...entries.keys()],
      getOwnPropertyDescriptor: (_, k) => (entries.has(k) ? { enumerable: true, configurable: true } : undefined),
      get: (_, k) => {
        const entry = entries.get(k);
        if (!entry) return undefined;
        if (!decoded.has(k)) {
          const packed = read(entry.start, entry.compressed),
            bytes = entry.method === 0 ? packed : inflateSync(packed, { out: new Uint8Array(entry.raw) });
          if (bytes.length !== entry.raw) throw refusal('internal-mismatch', 'ZIP member size mismatch');
          decoded.set(k, bytes);
        }
        return decoded.get(k);
      },
    });
    return { parts, close };
  } catch (error) {
    close();
    throw error;
  }
}
