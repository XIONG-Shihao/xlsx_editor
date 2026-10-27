import { FORMAT } from './format.mjs';
/** Owned filtering metadata. Visibility provenance is internal, never a public
 * criterion or renderer-supplied restoration plan. */
import { sha256 } from '@noble/hashes/sha2.js';
import { bytesToHex } from '@noble/hashes/utils.js';
import { refusal } from './refusal-codes.mjs';
export const FILTER_URI = FORMAT.URN_OWNED_FILTER_V1;
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const children = (n) => Array.from(n?.childNodes ?? []).filter((c) => c.nodeType === 1);
const attrs = (n) =>
  Object.fromEntries(
    Array.from(n?.attributes ?? [])
      .filter((a) => !a.name.startsWith('xmlns'))
      .map((a) => [a.name, a.value])
  );
const exact = (v, keys) =>
  v &&
  typeof v === 'object' &&
  !Array.isArray(v) &&
  Object.keys(v).length === keys.length &&
  keys.every((k) => Object.hasOwn(v, k));
const encoder = new TextEncoder(),
  byteLength = (text) => encoder.encode(text).byteLength;
const hash = (text) => bytesToHex(sha256(encoder.encode(text)));
const encode = (v) => (Object.is(v, -0) ? { $workbookNumber: '-0' } : v);
const decode = (v) => (exact(v, ['$workbookNumber']) && v.$workbookNumber === '-0' ? -0 : v);
function bounds(range) {
  const m = /^([A-Z]{1,3})([1-9]\d*):([A-Z]{1,3})([1-9]\d*)$/.exec(range);
  if (!m) throw refusal('internal-invalid-field', 'Invalid filter range');
  const col = (s) => [...s].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
  const b = { first: Number(m[2]) - 1, last: Number(m[4]) - 1, left: col(m[1]), right: col(m[3]) };
  if (b.first >= b.last || b.last >= 1048576 || b.left > b.right || b.right >= 16384)
    throw refusal('internal-invalid-field', 'Invalid filter range');
  return b;
}
export function validateFilterCriteria(range, column, values) {
  const b = bounds(range);
  if (typeof column !== 'string' || !/^[A-Z]{1,3}$/.test(column))
    throw refusal('internal-invalid-field', 'Invalid filter column');
  const c = [...column].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  if (c < b.left || c > b.right) throw refusal('filter-unsupported', 'Filter column is outside the requested range');
  if (
    !Array.isArray(values) ||
    values.length > 1000 ||
    values.some(
      (v) =>
        (v !== null && !(['string', 'boolean'].includes(typeof v) || (typeof v === 'number' && Number.isFinite(v)))) ||
        (typeof v === 'string' && (byteLength(v) > 32768 || /[\uD800-\uDFFF]/u.test(v)))
    )
  )
    throw refusal('internal-invalid-field', 'Invalid filter values');
  return b;
}
export function validateFilterState(model) {
  if (
    !exact(model, ['schema', 'range', 'column', 'values', 'baseline']) ||
    model.schema !== 1 ||
    !exact(model.baseline, ['format', 'rows'])
  )
    throw refusal('internal-invalid-field', 'Invalid owned filter provenance');
  const values = model.values.map(decode),
    b = validateFilterCriteria(model.range, model.column, values);
  const f = model.baseline.format;
  if (
    !f ||
    typeof f !== 'object' ||
    Array.isArray(f) ||
    Object.entries(f).some(([k, v]) => typeof v !== 'string' || !/^(?:\w+:)?\w+$/.test(k)) ||
    (Object.hasOwn(f, 'zeroHeight') && !['0', '1', 'true', 'false'].includes(f.zeroHeight))
  )
    throw refusal('internal-invalid-field', 'Invalid filter source defaults');
  if (!Array.isArray(model.baseline.rows) || model.baseline.rows.length > 1048576)
    throw refusal('internal-invalid-field', 'Invalid filter visibility baseline');
  let previous = b.first;
  for (const item of model.baseline.rows) {
    if (
      !Array.isArray(item) ||
      item.length !== 2 ||
      !Number.isSafeInteger(item[0]) ||
      item[0] <= previous ||
      item[0] > b.last ||
      (item[1] !== null && !['0', '1', 'true', 'false'].includes(item[1]))
    )
      throw refusal('internal-invalid-field', 'Invalid filter visibility baseline');
    previous = item[0];
  }
  return { model, values, bounds: b };
}
export function filterStateRecord(model) {
  validateFilterState(model);
  const text = JSON.stringify(model);
  if (byteLength(text) > 16 * 1024 ** 2) throw refusal('filter-limit', 'Filter provenance exceeds its metadata budget');
  return { model, text, revision: hash(text) };
}
export function filterDefinition(record) {
  const { model, values } = validateFilterState(record.model),
    active = values.length > 0;
  return {
    range: model.range,
    column: model.column,
    values,
    owned: true,
    revision: record.revision,
    canReapply: active,
    canClear: active,
    ...(!active ? { reason: 'No active filter criteria' } : {}),
  };
}
export function readFilterRecord(filter) {
  if (!filter) return null;
  const nodes = children(filter),
    exts = nodes.filter((n) => n.namespaceURI === NS && n.localName === 'extLst');
  // Not criteria: a column that only hides or shows its dropdown button, and
  // the record of how the range was last sorted. They are kept as they are.
  const buttonOnly = (n) =>
    n.namespaceURI === NS &&
    n.localName === 'filterColumn' &&
    !children(n).length &&
    Object.keys(attrs(n)).every((key) => ['colId', 'hiddenButton', 'showButton'].includes(key) || key.includes(':'));
  const passive = (n) => buttonOnly(n) || (n.namespaceURI === NS && n.localName === 'sortState');
  if (!exts.length) {
    if (nodes.some((n) => !passive(n)))
      throw refusal('filter-unsupported', 'Imported filter criteria have no owned visibility provenance');
    return null;
  }
  if (exts.length !== 1 || children(exts[0]).length !== 1)
    throw refusal('internal-unknown-reference', 'Unknown filter extensions');
  const ext = children(exts[0])[0],
    records = children(ext);
  if (
    ext.namespaceURI !== NS ||
    ext.localName !== 'ext' ||
    ext.getAttribute('uri') !== FILTER_URI ||
    records.length !== 1
  )
    throw refusal('internal-unknown-reference', 'Unknown filter extensions');
  const state = records[0];
  if (
    state.namespaceURI !== FILTER_URI ||
    state.localName !== 'filterState' ||
    Object.keys(attrs(state)).join(',') !== 'revision' ||
    children(state).length
  )
    throw refusal('internal-invalid-field', 'Invalid owned filter record');
  const text = state.textContent,
    revision = state.getAttribute('revision');
  if (byteLength(text) > 16 * 1024 ** 2 || !/^[a-f0-9]{64}$/.test(revision) || hash(text) !== revision)
    throw refusal('filter-bad-input', 'Filter revision does not match its provenance');
  const model = JSON.parse(text);
  validateFilterState(model);
  if (model.range !== filter.getAttribute('ref'))
    throw refusal('internal-mismatch', 'Filter range differs from owned criteria');
  const ordinary = nodes.filter((n) => n !== exts[0] && !passive(n));
  const expected = filterCriteriaXml(model);
  const actual = ordinary.map((n) => {
    if (
      n.namespaceURI !== NS ||
      n.localName !== 'filterColumn' ||
      !n.hasAttribute('colId') ||
      Object.keys(attrs(n)).some((key) => !['colId', 'hiddenButton', 'showButton'].includes(key)) ||
      children(n).length !== 1
    )
      throw refusal('internal-unknown-reference', 'Unknown filter criteria');
    const fs = children(n)[0];
    if (
      fs.namespaceURI !== NS ||
      fs.localName !== 'filters' ||
      Object.keys(attrs(fs)).some((k) => k !== 'blank') ||
      children(fs).some(
        (v) =>
          v.namespaceURI !== NS ||
          v.localName !== 'filter' ||
          Object.keys(attrs(v)).join(',') !== 'val' ||
          children(v).length
      )
    )
      throw refusal('internal-unknown-reference', 'Unknown filter criteria');
    return {
      column: n.getAttribute('colId'),
      blank: fs.getAttribute('blank') || null,
      values: children(fs).map((v) => v.getAttribute('val')),
    };
  });
  const selected = model.values.map(decode).filter((v) => v !== null && v !== '');
  if (
    actual.length !== expected.length ||
    actual.some(
      (a, i) =>
        a.column !== expected[i].column ||
        a.blank !== expected[i].blank ||
        a.values.length !== selected.length ||
        a.values.some((v, j) =>
          typeof selected[j] === 'number'
            ? !/^-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(v) || !Object.is(Number(v), selected[j])
            : v !== expected[i].values[j]
        )
    )
  )
    throw refusal('internal-mismatch', 'Saved filter criteria differ from typed owned criteria');
  return { model, text, revision };
}
function filterCriteriaXml(model) {
  if (!model.values.length) return [];
  const { left } = bounds(model.range),
    col = [...model.column].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1,
    values = model.values.map(decode);
  return [
    {
      column: String(col - left),
      blank: values.some((v) => v === null || v === '') ? '1' : null,
      values: values
        .filter((v) => v !== null && v !== '')
        .map((v) => (typeof v === 'boolean' ? String(Number(v)) : String(v))),
    },
  ];
}
export function inspectFilter(filter) {
  if (!filter) return { definition: undefined, record: null };
  try {
    const record = readFilterRecord(filter);
    return record
      ? { definition: filterDefinition(record), record }
      : {
          definition: {
            range: filter.getAttribute('ref'),
            owned: false,
            revision: null,
            canReapply: false,
            canClear: false,
            reason: 'No active owned filter criteria',
          },
          record: null,
        };
  } catch (error) {
    return {
      definition: {
        range: filter.getAttribute('ref') || '',
        owned: false,
        revision: null,
        canReapply: false,
        canClear: false,
        reason: String(error.message),
      },
      record: null,
    };
  }
}
export function captureFilterBaseline(document, range) {
  const b = bounds(range),
    root = document.documentElement,
    format = children(root).filter((n) => n.namespaceURI === NS && n.localName === 'sheetFormatPr');
  if (format.length > 1) throw refusal('filter-ambiguous', 'Ambiguous worksheet defaults');
  const data = children(root).find((n) => n.namespaceURI === NS && n.localName === 'sheetData'),
    rows = children(data)
      .filter((n) => n.namespaceURI === NS && n.localName === 'row')
      .map((n) => [Number(n.getAttribute('r')) - 1, n.hasAttribute('hidden') ? n.getAttribute('hidden') : null])
      .filter(([r]) => r > b.first && r <= b.last)
      .sort((a, b) => a[0] - b[0]);
  return { format: attrs(format[0]), rows };
}
export function prepareFilterState(command, record, baseline) {
  const follow = ['reapplyFilter', 'clearFilter', 'removeFilter'].includes(command.operation);
  if (follow) {
    if (
      !exact(command, ['operation', 'sheet', 'expectedFilterRevision']) ||
      !record ||
      command.expectedFilterRevision !== record.revision
    )
      throw refusal('filter-stale', 'Filter definition is stale or has no owned provenance');
    if (command.operation !== 'removeFilter' && !record.model.values.length)
      throw refusal('filter-bad-input', 'No active filter criteria');
  } else {
    validateFilterCriteria(command.range, command.column, command.values);
    if (record && record.model.range !== command.range)
      throw refusal('filter-unsupported', 'Remove the existing owned filter before changing its range');
  }
  const model = follow
    ? structuredClone(record.model)
    : {
        schema: 1,
        range: command.range,
        column: command.column,
        values: command.values.map(encode),
        baseline: record ? structuredClone(record.model.baseline) : baseline,
      };
  if (['clearFilter', 'removeFilter'].includes(command.operation)) model.values = [];
  return filterStateRecord(model);
}
export function baselineHidden(record, row, lookup) {
  const hidden = (lookup ?? new Map(record.model.baseline.rows)).get(row) ?? null;
  return (
    hidden === '1' ||
    hidden === 'true' ||
    (hidden === null && ['1', 'true'].includes(record.model.baseline.format.zeroHeight))
  );
}
export function appendFilterRecord(document, filter, record) {
  const extLst = document.createElementNS(NS, 'extLst'),
    ext = document.createElementNS(NS, 'ext'),
    state = document.createElementNS(FILTER_URI, FORMAT.FILTERSTATE);
  ext.setAttribute('uri', FILTER_URI);
  state.setAttribute('revision', record.revision);
  state.appendChild(document.createTextNode(record.text));
  ext.appendChild(state);
  extLst.appendChild(ext);
  filter.appendChild(extLst);
}
