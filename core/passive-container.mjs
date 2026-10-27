/** Bounded passive container admission. Original bytes are never converted. */
import { unzipSync } from 'fflate';
import { DOMParser } from '@xmldom/xmldom';
import { refusal } from './refusal-codes.mjs';
export const PASSIVE_FORMAT_LIMITS = Object.freeze({
  sourceBytes: 64 * 1024 ** 2,
  expandedBytes: 128 * 1024 ** 2,
  partBytes: 32 * 1024 ** 2,
  parts: 4096,
  sheets: 64,
  storedCells: 1000000,
  readPositions: 1000,
  replyBytes: 1024 ** 2,
  text: 32768,
  heapMiB: 256,
  openTimeoutMs: 30000,
});
const TYPES = 'http://schemas.openxmlformats.org/package/2006/content-types',
  RELS = 'http://schemas.openxmlformats.org/package/2006/relationships',
  OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const fail = (message) => {
  throw refusal('package-refused', 'Passive workbook: ' + message);
};
const safePart = (name) =>
  typeof name === 'string' &&
  name.length <= 4096 &&
  !name.startsWith('/') &&
  !/[\\\x00-\x1f]/.test(name) &&
  !name.split('/').some((p) => !p || p === '.' || p === '..');
const children = (node) => Array.from(node.childNodes).filter((n) => n.nodeType === 1);
const crcTable = Uint32Array.from({ length: 256 }, (_, i) => {
  let n = i;
  for (let k = 0; k < 8; k++) n = n & 1 ? 0xedb88320 ^ (n >>> 1) : n >>> 1;
  return n >>> 0;
});
const crc = (bytes) => {
  let value = 0xffffffff;
  for (const byte of bytes) value = crcTable[(value ^ byte) & 255] ^ (value >>> 8);
  return (value ^ 0xffffffff) >>> 0;
};
function xml(bytes, namespace, root) {
  if (!(bytes instanceof Uint8Array) || bytes.length > PASSIVE_FORMAT_LIMITS.partBytes)
    fail('missing or oversized XML part');
  const raw = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  if (/<!DOCTYPE|<!ENTITY/i.test(raw)) fail('XML entities are unsupported');
  const bad = () => fail('malformed XML graph'),
    doc = new DOMParser({ errorHandler: { warning: bad, error: bad, fatalError: bad } }).parseFromString(
      raw,
      'application/xml'
    );
  if (doc.documentElement.namespaceURI !== namespace || doc.documentElement.localName !== root)
    fail('incorrect XML namespace or root');
  return doc.documentElement;
}
function zipParts(input) {
  const view = new DataView(input.buffer, input.byteOffset, input.byteLength);
  let end = -1;
  for (let i = input.length - 22; i >= Math.max(0, input.length - 65557); i--)
    if (view.getUint32(i, true) === 0x06054b50 && i + 22 + view.getUint16(i + 20, true) === input.length) {
      end = i;
      break;
    }
  if (end < 0 || view.getUint16(end + 4, true) || view.getUint16(end + 6, true))
    fail('unsupported or incomplete ZIP container');
  const count = view.getUint16(end + 10, true),
    size = view.getUint32(end + 12, true),
    start = view.getUint32(end + 16, true);
  if (!count || count > PASSIVE_FORMAT_LIMITS.parts || count !== view.getUint16(end + 8, true) || start + size !== end)
    fail('ZIP directory exceeds its closed bounds');
  const entries = new Map(),
    canonical = new Set(),
    intervals = [];
  let at = start,
    total = 0;
  for (let i = 0; i < count; i++) {
    if (at + 46 > end || view.getUint32(at, true) !== 0x02014b50) fail('invalid ZIP directory');
    const flags = view.getUint16(at + 8, true),
      method = view.getUint16(at + 10, true),
      checksum = view.getUint32(at + 16, true),
      compressed = view.getUint32(at + 20, true),
      original = view.getUint32(at + 24, true),
      length = view.getUint16(at + 28, true),
      extra = view.getUint16(at + 30, true),
      comment = view.getUint16(at + 32, true),
      local = view.getUint32(at + 42, true);
    if (
      flags & 0x2041 ||
      ![0, 8].includes(method) ||
      view.getUint16(at + 34, true) ||
      original > PASSIVE_FORMAT_LIMITS.partBytes ||
      at + 46 + length + extra + comment > end
    )
      fail('encrypted, oversized or unsupported ZIP member');
    const nameBytes = input.subarray(at + 46, at + 46 + length);
    if (!(flags & 0x800) && nameBytes.some((b) => b >= 128)) fail('unsupported ZIP filename encoding');
    const name = new TextDecoder('utf-8', { fatal: true }).decode(nameBytes),
      directory = name.endsWith('/'),
      part = directory ? name.slice(0, -1) : name,
      key = part.normalize('NFC').toLowerCase();
    if (!safePart(part) || canonical.has(key)) fail('ambiguous or escaping ZIP member');
    canonical.add(key);
    if (
      local + 30 > start ||
      view.getUint32(local, true) !== 0x04034b50 ||
      view.getUint16(local + 6, true) !== flags ||
      view.getUint16(local + 8, true) !== method
    )
      fail('ZIP member header disagrees');
    const localName = view.getUint16(local + 26, true),
      localExtra = view.getUint16(local + 28, true),
      data = local + 30 + localName + localExtra;
    if (
      localName !== length ||
      data + compressed > start ||
      !input.subarray(local + 30, local + 30 + localName).every((b, j) => b === nameBytes[j])
    )
      fail('ZIP member boundary disagrees');
    if (
      !(flags & 8) &&
      (view.getUint32(local + 14, true) !== checksum ||
        view.getUint32(local + 18, true) !== compressed ||
        view.getUint32(local + 22, true) !== original)
    )
      fail('ZIP local sizes or checksum disagree');
    if ((method === 0 && compressed !== original) || (directory && (original || compressed)))
      fail('invalid stored ZIP member');
    total += original;
    if (total > PASSIVE_FORMAT_LIMITS.expandedBytes) fail('ZIP expansion exceeds its budget');
    intervals.push([local, data + compressed]);
    if (!directory) entries.set(name, { original, checksum });
    at += 46 + length + extra + comment;
  }
  if (at !== end) fail('unexpected ZIP directory bytes');
  intervals.sort((a, b) => a[0] - b[0]);
  for (let i = 1; i < intervals.length; i++) if (intervals[i][0] < intervals[i - 1][1]) fail('overlapping ZIP members');
  const parts = unzipSync(input, { filter: (file) => entries.has(file.name) });
  if (Object.keys(parts).length !== entries.size) fail('ZIP membership changed while decoding');
  for (const [name, entry] of entries)
    if (parts[name]?.length !== entry.original || crc(parts[name]) !== entry.checksum)
      fail('ZIP bytes or checksum disagree');
  return parts;
}
function target(owner, value) {
  if (
    typeof value !== 'string' ||
    !value ||
    value.includes('\\') ||
    /[\x00-\x1f?#]/.test(value) ||
    /^[a-z][a-z\d+.-]*:/i.test(value)
  )
    fail('invalid internal OPC target');
  const stack = value.startsWith('/') ? [] : owner.split('/').slice(0, -1);
  for (const part of value.split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') {
      if (!stack.length) fail('escaping OPC target');
      stack.pop();
    } else stack.push(part);
  }
  const result = stack.join('/');
  if (!safePart(result)) fail('invalid OPC target');
  return result;
}
function relationOwner(part) {
  if (part === '_rels/.rels') return '';
  const match = /^(?:(.*\/))?_rels\/([^/]+)\.rels$/.exec(part);
  if (!match) fail('unsupported OPC relationship path');
  return (match[1] ?? '') + match[2];
}
export function admitPassiveContainer(input, format) {
  if (
    !(input instanceof Uint8Array) ||
    !input.length ||
    input.length > PASSIVE_FORMAT_LIMITS.sourceBytes ||
    !['xlsm', 'xlsb', 'xls'].includes(format)
  )
    fail('invalid source format or byte budget');
  if (format === 'xls') {
    const magic = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
    if (input.length < 512 || !magic.every((b, i) => input[i] === b)) fail('XLS requires a compound binary container');
    const view = new DataView(input.buffer, input.byteOffset, input.byteLength),
      major = view.getUint16(26, true),
      shift = view.getUint16(30, true);
    if (
      view.getUint16(28, true) !== 0xfffe ||
      ![3, 4].includes(major) ||
      shift !== (major === 3 ? 9 : 12) ||
      view.getUint16(32, true) !== 6 ||
      view.getUint32(56, true) !== 4096 ||
      input.length % (1 << shift) ||
      (major === 3 && view.getUint32(40, true) !== 0)
    )
      fail('unsupported compound binary header');
    return {
      format,
      container: 'cfb',
      parts: null,
      vbaParts: [],
      notice: 'Read-only stored cells; binary structures and macros are preserved by exact original-file Copy.',
    };
  }
  if (input.length < 22 || input[0] !== 80 || input[1] !== 75) fail('OOXML format requires a ZIP container');
  const parts = zipParts(input),
    typeRoot = xml(parts['[Content_Types].xml'], TYPES, 'Types'),
    types = new Map(),
    defaults = new Map();
  for (const node of children(typeRoot)) {
    const mime = node.getAttribute('ContentType');
    if (node.namespaceURI !== TYPES || children(node).length || !mime || mime !== mime.trim() || /\s/.test(mime))
      fail('invalid OPC content type');
    if (node.localName === 'Override') {
      const name = node.getAttribute('PartName');
      if (!name.startsWith('/') || !safePart(name.slice(1)) || types.has(name.slice(1))) fail('ambiguous OPC override');
      types.set(name.slice(1), mime);
    } else if (node.localName === 'Default') {
      const extension = node.getAttribute('Extension').toLowerCase();
      if (!/^[a-z\d]+$/.test(extension) || defaults.has(extension)) fail('ambiguous OPC default');
      defaults.set(extension, mime);
    } else fail('unknown OPC type entry');
  }
  const relations = new Map();
  let external = 0;
  for (const [part, bytes] of Object.entries(parts))
    if (part.endsWith('.rels')) {
      const owner = relationOwner(part),
        root = xml(bytes, RELS, 'Relationships'),
        ids = new Set(),
        entries = [];
      for (const node of children(root)) {
        const id = node.getAttribute('Id'),
          type = node.getAttribute('Type'),
          mode = node.getAttribute('TargetMode'),
          rawTarget = node.getAttribute('Target');
        if (
          node.namespaceURI !== RELS ||
          node.localName !== 'Relationship' ||
          children(node).length ||
          !id ||
          ids.has(id) ||
          !/^https?:\/\/[^\s]+$/.test(type) ||
          !rawTarget ||
          (mode && !['External'].includes(mode))
        )
          fail('invalid OPC relationship');
        ids.add(id);
        if (mode === 'External') {
          external++;
          entries.push({ type, external: true });
          continue;
        }
        const resolved = target(owner, rawTarget);
        if (!parts[resolved]) fail('missing internal OPC target');
        entries.push({ type, target: resolved });
      }
      relations.set(owner, entries);
    }
  const main = (relations.get('') ?? []).filter((r) => r.type === OFFICE + 'officeDocument');
  if (main.length !== 1 || main[0].external) fail('ambiguous workbook root relationship');
  const workbookPart = main[0].target,
    expectedPart = format === 'xlsm' ? 'xl/workbook.xml' : 'xl/workbook.bin',
    expectedType =
      format === 'xlsm'
        ? 'application/vnd.ms-excel.sheet.macroEnabled.main+xml'
        : 'application/vnd.ms-excel.sheet.binary.macroEnabled.main';
  if (
    workbookPart !== expectedPart ||
    (types.get(workbookPart) ?? defaults.get(workbookPart.split('.').at(-1))) !== expectedType
  )
    fail('source extension and workbook MIME disagree');
  const vbaParts = (relations.get(workbookPart) ?? [])
    .filter((r) => r.type === 'http://schemas.microsoft.com/office/2006/relationships/vbaProject')
    .map((r) => {
      if (r.external) fail('external VBA project is unsupported');
      if ((types.get(r.target) ?? defaults.get(r.target.split('.').at(-1))) !== 'application/vnd.ms-office.vbaProject')
        fail('VBA project MIME disagrees');
      return r.target;
    });
  if (vbaParts.length > 1) fail('ambiguous VBA project');
  return {
    format,
    container: 'opc',
    parts,
    workbookPart,
    vbaParts,
    externalRelationships: external,
    notice:
      'Read-only stored cells; macros and external data are not run. Original package bytes are retained for same-format Copy.',
  };
}
export function validatePassiveCfb(cfb, { workbook = false } = {}) {
  if (
    !Array.isArray(cfb?.FullPaths) ||
    !Array.isArray(cfb?.FileIndex) ||
    cfb.FullPaths.length !== cfb.FileIndex.length ||
    cfb.FullPaths.length > PASSIVE_FORMAT_LIMITS.parts
  )
    fail('invalid compound directory');
  const seen = new Set();
  let total = 0,
    stream = null;
  for (let i = 0; i < cfb.FullPaths.length; i++) {
    const name = cfb.FullPaths[i],
      entry = cfb.FileIndex[i];
    if (!entry || ![0, 1, 2, 5].includes(entry.type) || (entry.type === 5 && i !== 0))
      fail('invalid compound directory entry kind');
    if (entry.type === 0) {
      if (entry.content?.length) fail('free compound entry carries bytes');
      continue;
    }
    if (
      typeof name !== 'string' ||
      name.length > 4096 ||
      name.includes('\\') ||
      name.includes('\0') ||
      name.split('/').some((p) => p === '.' || p === '..' || /[\x00-\x1f]/.test(p.replace(/^[\x01\x05]/, '')))
    )
      fail('invalid compound stream path');
    const key = name.toLowerCase();
    if (seen.has(key)) fail('duplicate compound stream');
    seen.add(key);
    if (entry.type === 2) {
      if (
        !(entry.content instanceof Uint8Array) ||
        entry.content.length > PASSIVE_FORMAT_LIMITS.partBytes ||
        !Number.isSafeInteger(entry.size) ||
        entry.size !== entry.content.length
      )
        fail('oversized or truncated compound stream');
      total += entry.content.length;
      if (total > PASSIVE_FORMAT_LIMITS.expandedBytes) fail('compound streams exceed their budget');
      if (/\/(?:EncryptionInfo|EncryptedPackage)$/i.test(name)) fail('encrypted workbook is unsupported');
      if (/\/(?:Workbook|Book)$/i.test(name)) {
        if (stream) fail('ambiguous binary Workbook stream');
        stream = entry.content;
      }
    }
  }
  if (workbook) {
    if (!stream || stream.length < 20) fail('missing binary Workbook stream');
    const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
    if (
      view.getUint16(0, true) !== 0x0809 ||
      view.getUint16(2, true) < 16 ||
      view.getUint16(4, true) !== 0x0600 ||
      view.getUint16(6, true) !== 5
    )
      fail('only BIFF8 XLS passive admission is qualified');
    let at = 0,
      ended = false;
    while (at + 4 <= stream.length) {
      const id = view.getUint16(at, true),
        size = view.getUint16(at + 2, true);
      if (at + 4 + size > stream.length) fail('truncated BIFF global record');
      if (id === 0x002f) fail('encrypted XLS is unsupported');
      at += 4 + size;
      if (id === 0x000a) {
        ended = true;
        break;
      }
    }
    if (!ended) fail('missing BIFF global terminator');
  }
  return cfb;
}
