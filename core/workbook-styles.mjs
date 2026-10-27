import { XMLSerializer } from '@xmldom/xmldom';
import { refusal } from './refusal-codes.mjs';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const serializer = new XMLSerializer();
const children = (element, name) =>
  Array.from(element?.childNodes ?? []).filter(
    (node) => node.nodeType === 1 && node.namespaceURI === NS && (!name || node.localName === name)
  );
const child = (element, name) => children(element, name)[0];
const create = (document, name, attributes = {}) => {
  const result = document.createElementNS(NS, name);
  for (const [key, value] of Object.entries(attributes)) result.setAttribute(key, String(value));
  return result;
};
function replaceChild(parent, name, replacement) {
  for (const old of children(parent, name)) parent.removeChild(old);
  if (replacement) parent.appendChild(replacement);
}
const MAX_INDEX_BYTES = 4 * 1024 * 1024,
  MAX_INDEX_ENTRIES = 16384;
function disableIndex(cache) {
  cache.disabled = true;
  cache.tables = new WeakMap();
  cache.bytes = 0;
  cache.entries = 0;
}
function reserveIndex(cache, signature) {
  const bytes = signature.length * 2 + 96;
  if (cache.bytes + bytes > cache.maxBytes || cache.entries >= cache.maxEntries) {
    disableIndex(cache);
    return false;
  }
  cache.bytes += bytes;
  cache.entries++;
  cache.peakBytes = Math.max(cache.peakBytes, cache.bytes);
  return true;
}
function styleIndex(parent, name, cache) {
  if (!parent || !cache || cache.disabled) return null;
  let index = cache.tables.get(parent);
  if (index) return index;
  const entries = children(parent, name),
    signatures = new Map();
  if (entries.length > cache.maxEntries - cache.entries) {
    disableIndex(cache);
    return null;
  }
  for (let i = 0; i < entries.length; i++) {
    const signature = serializer.serializeToString(entries[i]);
    if (!reserveIndex(cache, signature)) return null;
    // Preserve the old writer's first matching record when a source has duplicates.
    if (!signatures.has(signature)) signatures.set(signature, i);
  }
  index = { entries, signatures };
  cache.tables.set(parent, index);
  return index;
}
const styleEntries = (parent, name, cache) => styleIndex(parent, name, cache)?.entries ?? children(parent, name);
function appendUnique(parent, name, value, cache) {
  const signature = serializer.serializeToString(value);
  const indexed = styleIndex(parent, name, cache);
  const entries = indexed?.entries ?? children(parent, name);
  const index = indexed
    ? (indexed.signatures.get(signature) ?? -1)
    : entries.findIndex((entry) => serializer.serializeToString(entry) === signature);
  if (index >= 0) return index;
  const next = entries.length;
  parent.appendChild(value);
  parent.setAttribute('count', String(next + 1));
  if (indexed) {
    const storedSignature = serializer.serializeToString(value);
    if (reserveIndex(cache, storedSignature)) {
      entries.push(value);
      if (!indexed.signatures.has(storedSignature)) indexed.signatures.set(storedSignature, next);
    }
  }
  return next;
}

/** Synchronous exclusive batch over one privately parsed style document. The
 * writer only appends/clones records; callers must not mutate existing records
 * during the batch. Signature accounting is an estimate, not a heap cap. */
export function createWorkbookStyleWriter(styles, { maxBytes = MAX_INDEX_BYTES, maxEntries = MAX_INDEX_ENTRIES } = {}) {
  if (
    !Number.isSafeInteger(maxBytes) ||
    maxBytes < 0 ||
    maxBytes > MAX_INDEX_BYTES ||
    !Number.isSafeInteger(maxEntries) ||
    maxEntries < 0 ||
    maxEntries > MAX_INDEX_ENTRIES
  )
    throw refusal('style-malformed', 'Invalid style index budget');
  const cache = { tables: new WeakMap(), bytes: 0, entries: 0, peakBytes: 0, disabled: false, maxBytes, maxEntries };
  // Small or repeatedly identical batches already avoid most scanning. Delay
  // indexing so one Bold edit does not index a large pre-existing style table.
  let writes = 0;
  const apply = (cell, edit) => applyStyle(styles, cell, edit, ++writes < 32 ? undefined : cache);
  apply.diagnostics = () => ({
    estimatedIndexBytes: cache.bytes,
    peakEstimatedIndexBytes: cache.peakBytes,
    indexedEntries: cache.entries,
    styleWrites: writes,
    disabled: cache.disabled,
  });
  return apply;
}
function ensureContainer(styles, name, before) {
  const existing = child(styles.documentElement, name);
  if (existing) return existing;
  const result = create(styles, name, { count: 0 });
  styles.documentElement.insertBefore(
    result,
    before.map((tag) => child(styles.documentElement, tag)).find(Boolean) ?? null
  );
  return result;
}
function numberFormatId(styles, name) {
  const builtIn = {
    general: 0,
    number: 2,
    percent: 10,
    date: 14,
    text: 49,
    accounting: 44,
    thousands: 3,
    decimal0: 1,
    decimal2: 2,
  };
  if (Object.hasOwn(builtIn, name)) return builtIn[name];
  const code = { currency: '"$"#,##0.00;("$"#,##0.00)', decimal1: '0.0' }[name];
  if (!code) throw refusal('style-unsupported', 'Unsupported number format');
  const formats = ensureContainer(styles, 'numFmts', ['fonts', 'fills', 'borders', 'cellXfs']);
  const entries = children(formats, 'numFmt');
  const prior = entries.find((entry) => entry.getAttribute('formatCode') === code);
  if (prior) return Number(prior.getAttribute('numFmtId'));
  const id = Math.max(163, ...entries.map((entry) => Number(entry.getAttribute('numFmtId')) || 0)) + 1;
  formats.appendChild(create(styles, 'numFmt', { numFmtId: id, formatCode: code }));
  formats.setAttribute('count', String(entries.length + 1));
  return id;
}

/** Clone only the selected style records; never mutate a style used by another cell. */
export function applyWorkbookStyle(styles, cell, edit) {
  return applyStyle(styles, cell, edit);
}
function applyStyle(styles, cell, edit, cache) {
  const xfs = child(styles.documentElement, 'cellXfs');
  const base = styleEntries(xfs, 'xf', cache)[Number(cell.getAttribute('s') || 0)];
  if (!base) throw refusal('style-unsupported', 'Unsupported cell style');
  const xf = base.cloneNode(true);
  const fontFields = ['bold', 'italic', 'underline', 'fontFamily', 'fontSize', 'fontColor'];
  if (fontFields.some((field) => edit[field] !== undefined)) {
    const fonts = child(styles.documentElement, 'fonts');
    const baseFont = styleEntries(fonts, 'font', cache)[Number(xf.getAttribute('fontId') || 0)];
    if (!baseFont) throw refusal('style-unsupported', 'Unsupported font style');
    const font = baseFont.cloneNode(true);
    for (const [field, tag] of [
      ['bold', 'b'],
      ['italic', 'i'],
      ['underline', 'u'],
    ])
      if (edit[field] !== undefined) replaceChild(font, tag, edit[field] ? create(styles, tag) : null);
    for (const [field, tag] of [
      ['fontFamily', 'name'],
      ['fontSize', 'sz'],
    ])
      if (edit[field] !== undefined) replaceChild(font, tag, create(styles, tag, { val: edit[field] }));
    // Choosing a literal Home font detaches its inherited theme scheme.
    if (edit.fontFamily !== undefined) replaceChild(font, 'scheme', null);
    if (edit.fontColor !== undefined)
      replaceChild(font, 'color', create(styles, 'color', { rgb: `FF${edit.fontColor.toUpperCase()}` }));
    xf.setAttribute('fontId', String(appendUnique(fonts, 'font', font, cache)));
    xf.setAttribute('applyFont', '1');
  }
  if (edit.fillColor !== undefined) {
    const fills = ensureContainer(styles, 'fills', ['borders', 'cellStyleXfs', 'cellXfs']);
    const fill = create(styles, 'fill');
    const pattern = create(styles, 'patternFill', { patternType: 'solid' });
    pattern.appendChild(create(styles, 'fgColor', { rgb: `FF${edit.fillColor.toUpperCase()}` }));
    pattern.appendChild(create(styles, 'bgColor', { indexed: 64 }));
    fill.appendChild(pattern);
    xf.setAttribute('fillId', String(appendUnique(fills, 'fill', fill, cache)));
    xf.setAttribute('applyFill', '1');
  }
  if (edit.border !== undefined) {
    const borders = ensureContainer(styles, 'borders', ['cellStyleXfs', 'cellXfs']);
    const border = create(styles, 'border');
    for (const side of ['left', 'right', 'top', 'bottom'])
      border.appendChild(create(styles, side, edit.border === 'none' ? {} : { style: 'thin' }));
    border.appendChild(create(styles, 'diagonal'));
    xf.setAttribute('borderId', String(appendUnique(borders, 'border', border, cache)));
    xf.setAttribute('applyBorder', '1');
  }
  if (edit.numberFormat !== undefined) {
    xf.setAttribute('numFmtId', String(numberFormatId(styles, edit.numberFormat)));
    xf.setAttribute('applyNumberFormat', '1');
  }
  if (['horizontalAlign', 'verticalAlign', 'wrapText'].some((field) => edit[field] !== undefined)) {
    const alignment = child(xf, 'alignment') ?? create(styles, 'alignment');
    if (edit.horizontalAlign !== undefined) alignment.setAttribute('horizontal', edit.horizontalAlign);
    if (edit.verticalAlign !== undefined)
      alignment.setAttribute('vertical', edit.verticalAlign === 'center' ? 'center' : edit.verticalAlign);
    if (edit.wrapText !== undefined) alignment.setAttribute('wrapText', edit.wrapText ? '1' : '0');
    if (!alignment.parentNode) xf.appendChild(alignment);
    xf.setAttribute('applyAlignment', '1');
  }
  cell.setAttribute('s', String(appendUnique(xfs, 'xf', xf, cache)));
}

/** Differential fills are referenced by conditional rules, never by ordinary XFs. */
export function appendDifferentialFill(styles, color) {
  const dxfs = ensureContainer(styles, 'dxfs', ['tableStyles', 'colors', 'extLst']);
  const dxf = create(styles, 'dxf');
  const fill = create(styles, 'fill');
  const pattern = create(styles, 'patternFill', { patternType: 'solid' });
  pattern.appendChild(create(styles, 'fgColor', { rgb: `FF${color.toUpperCase()}` }));
  pattern.appendChild(create(styles, 'bgColor', { rgb: `FF${color.toUpperCase()}` }));
  fill.appendChild(pattern);
  dxf.appendChild(fill);
  return appendUnique(dxfs, 'dxf', dxf);
}
