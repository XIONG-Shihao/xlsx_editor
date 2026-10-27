/** Passive MS-OVBA source decoding. No evaluation, compilation or macro dispatch. */
import { refusal } from './refusal-codes.mjs';

export const VBA_SOURCE_LIMITS = Object.freeze({
  projectBytes: 32 * 1024 ** 2,
  directoryBytes: 1024 ** 2,
  sourceBytes: 4 * 1024 ** 2,
  totalSourceBytes: 8 * 1024 ** 2,
  modules: 128,
  chunkText: 32768,
});
const bad = (message) => {
  throw refusal('workbook-refused', 'VBA source: ' + message);
};
const bounded = (value, max, label) => {
  if (!Number.isSafeInteger(value) || value < 1 || value > max) bad('invalid ' + label);
  return value;
};
/** MS-OVBA 2.4.1: per-chunk dictionaries, overlapping references and fixed raw
 * chunks. Strict source boundaries refuse truncation instead of partial output. */
export function decompressVba(input, { maxBytes = VBA_SOURCE_LIMITS.sourceBytes } = {}) {
  bounded(maxBytes, VBA_SOURCE_LIMITS.totalSourceBytes, 'output budget');
  if (
    !(input instanceof Uint8Array) ||
    input.length < 4 ||
    input.length > VBA_SOURCE_LIMITS.projectBytes ||
    input[0] !== 1
  )
    bad('invalid compressed container');
  const view = new DataView(input.buffer, input.byteOffset, input.byteLength),
    chunks = [];
  let at = 1,
    total = 0;
  while (at < input.length) {
    if (at + 2 > input.length) bad('truncated chunk header');
    const header = view.getUint16(at, true),
      size = (header & 0x0fff) + 3,
      end = at + size;
    if ((header & 0x7000) !== 0x3000 || end > input.length) bad('invalid chunk signature or boundary');
    const compressed = !!(header & 0x8000),
      output = new Uint8Array(4096);
    let count = 0;
    at += 2;
    if (!compressed) {
      if (size !== 4098) bad('raw chunk must contain4096 bytes');
      output.set(input.subarray(at, end));
      count = 4096;
      at = end;
    } else {
      while (at < end) {
        const flags = input[at++];
        let tokens = 0;
        for (let bit = 0; bit < 8 && at < end; bit++) {
          tokens++;
          if (flags & (1 << bit)) {
            if (at + 2 > end || count === 0) bad('invalid copy token');
            const token = view.getUint16(at, true);
            at += 2;
            const offsetBits = Math.max(4, Math.ceil(Math.log2(count))),
              lengthBits = 16 - offsetBits,
              mask = (1 << lengthBits) - 1;
            const offset = (token >>> lengthBits) + 1,
              length = (token & mask) + 3;
            if (offset > count || count + length > 4096) bad('copy token leaves its chunk');
            for (let i = 0; i < length; i++) {
              output[count] = output[count - offset];
              count++;
            }
          } else {
            if (count === 4096) bad('literal leaves its chunk');
            output[count++] = input[at++];
          }
        }
        // The retained Excel12 dir stream ends in a zero flag after its final
        // complete token. It contributes no bytes; any partial token or a flag
        // that claims a missing copy remains invalid.
        if (!tokens && !(flags === 0 && count > 0 && at === end)) bad('empty token sequence');
      }
    }
    if (!count || (count < 4096 && at < input.length)) bad('nonfinal chunk is incomplete');
    total += count;
    if (total > maxBytes) bad('source exceeds its output budget');
    chunks.push(output.subarray(0, count));
  }
  const result = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    result.set(chunk, offset);
    offset += chunk.length;
  }
  return result;
}
class Records {
  constructor(bytes) {
    this.bytes = bytes;
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    this.at = 0;
  }
  next() {
    return this.at + 2 <= this.bytes.length ? this.view.getUint16(this.at, true) : -1;
  }
  record(id, size, actualSize = size) {
    if (this.at + 6 > this.bytes.length || this.next() !== id) bad('unsupported or malformed directory record');
    const count = this.view.getUint32(this.at + 2, true);
    if ((size !== undefined && count !== size) || count > VBA_SOURCE_LIMITS.directoryBytes)
      bad('directory record size disagrees');
    const length = actualSize ?? count,
      start = this.at + 6,
      end = start + length;
    if (end > this.bytes.length) bad('truncated directory record');
    this.at = end;
    return this.bytes.subarray(start, end);
  }
}
const number = (bytes) => new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(0, true);
function decoder(codepage) {
  if (codepage === 1252) return new TextDecoder('windows-1252', { fatal: true });
  if (codepage === 65001) return new TextDecoder('utf-8', { fatal: true });
  bad('unsupported source codepage ' + codepage);
}
function unicode(bytes) {
  if (bytes.length % 2) bad('invalid Unicode directory value');
  return new TextDecoder('utf-16le', { fatal: true }).decode(bytes);
}
function identifier(value) {
  if (typeof value !== 'string' || !value || value.length > 128 || /[\x00-\x1f/\\]/.test(value))
    bad('invalid module identity');
  return value;
}
/** Closed directory profile: ordinary registered-library references and standard/
 * document/class modules. Unqualified controls/project-reference records refuse
 * source inspection while the outer workbook remains available for passive Copy. */
export function vbaDirectory(input) {
  const bytes = decompressVba(input, { maxBytes: VBA_SOURCE_LIMITS.directoryBytes }),
    r = new Records(bytes);
  r.record(1, 4);
  if (r.next() === 0x4a) r.record(0x4a, 4);
  r.record(2, 4);
  r.record(0x14, 4);
  const cp = r.record(3, 2),
    codepage = cp[0] | (cp[1] << 8),
    decode = decoder(codepage);
  const projectName = identifier(decode.decode(r.record(4)));
  r.record(5);
  unicode(r.record(0x40));
  r.record(6);
  r.record(0x3d);
  r.record(7, 4);
  r.record(8, 4);
  r.record(9, 4, 6);
  r.record(0x0c);
  unicode(r.record(0x3c));
  let references = 0;
  while (r.next() !== 0x0f) {
    if (++references > 128) bad('reference count exceeds its budget');
    if (r.next() === 0x16) {
      r.record(0x16);
      unicode(r.record(0x3e));
    }
    if (r.next() !== 0x0d) bad('control or project references are not qualified for source inspection');
    const reference = r.record(0x0d);
    if (
      reference.length < 10 ||
      number(reference) !== reference.length - 10 ||
      reference.subarray(-6).some((value) => value !== 0)
    )
      bad('invalid registered-library reference');
  }
  const countBytes = r.record(0x0f, 2),
    count = countBytes[0] | (countBytes[1] << 8);
  bounded(count, VBA_SOURCE_LIMITS.modules, 'module count');
  r.record(0x13, 2);
  const modules = [],
    names = new Set(),
    streams = new Set();
  for (let i = 0; i < count; i++) {
    let name = decode.decode(r.record(0x19));
    if (r.next() === 0x47) name = unicode(r.record(0x47));
    identifier(name);
    const legacyStream = decode.decode(r.record(0x1a)),
      stream = identifier(unicode(r.record(0x32)) || legacyStream);
    identifier(legacyStream);
    r.record(0x1c);
    unicode(r.record(0x48));
    const offset = number(r.record(0x31, 4));
    r.record(0x1e, 4);
    r.record(0x2c, 2);
    const type = r.next();
    if (![0x21, 0x22].includes(type)) bad('unsupported module kind');
    r.record(type, 0);
    let readOnly = false,
      privateModule = false;
    if (r.next() === 0x25) {
      r.record(0x25, 0);
      readOnly = true;
    }
    if (r.next() === 0x28) {
      r.record(0x28, 0);
      privateModule = true;
    }
    r.record(0x2b, 0);
    if (names.has(name.toLowerCase()) || streams.has(stream.toLowerCase())) bad('ambiguous module name or stream');
    names.add(name.toLowerCase());
    streams.add(stream.toLowerCase());
    modules.push({
      name,
      stream,
      legacyStream,
      offset,
      kind: type === 0x21 ? 'standard' : 'document-or-class',
      readOnly,
      privateModule,
    });
  }
  r.record(0x10, 0);
  if (r.at !== bytes.length) bad('unexpected directory suffix');
  return { projectName, codepage, modules };
}
/** The caller supplies a validated maintained-parser CFB model, never a path.
 * Offsets come from the dir records and must point inside their exact stream. */
export function inspectVbaSources(cfb) {
  const entries = new Map();
  for (let i = 0; i < cfb.FullPaths.length; i++) {
    if (cfb.FileIndex[i].type !== 2) continue;
    const key = cfb.FullPaths[i].toLowerCase();
    if (entries.has(key)) bad('duplicate compound source stream');
    entries.set(key, { path: cfb.FullPaths[i], bytes: cfb.FileIndex[i].content });
  }
  const dirs = [...entries.values()].filter((item) => /\/vba\/dir$/i.test(item.path));
  if (dirs.length !== 1) bad('project has no unique VBA directory');
  const dir = dirs[0],
    prefix = dir.path.slice(0, -3),
    metadata = vbaDirectory(dir.bytes),
    decode = decoder(metadata.codepage),
    modules = [];
  let total = 0;
  for (const module of metadata.modules) {
    const entry =
      entries.get((prefix + module.stream).toLowerCase()) ?? entries.get((prefix + module.legacyStream).toLowerCase());
    if (!entry || module.offset >= entry.bytes.length) bad('module source offset or stream is missing');
    const source = decompressVba(entry.bytes.subarray(module.offset));
    total += source.length;
    if (total > VBA_SOURCE_LIMITS.totalSourceBytes) bad('project sources exceed their budget');
    const code = decode.decode(source);
    if (code.includes('\0')) bad('module source contains NUL');
    modules.push({
      name: module.name,
      kind: module.kind,
      readOnly: module.readOnly,
      privateModule: module.privateModule,
      code,
      sourceBytes: source.length,
    });
  }
  return { projectName: metadata.projectName, codepage: metadata.codepage, modules, execution: 'disabled' };
}
