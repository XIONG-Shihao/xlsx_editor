import { FORMAT, matchesFormat } from './format.mjs';
/** Closed read-only source/effective geometry. Display dimensions grant no authority. */
import { refusal } from './refusal-codes.mjs';

const plain = (value) => value !== null && typeof value === 'object' && !Array.isArray(value);
const bytes = (value) => new TextEncoder().encode(value).length;
const exact = (value, keys, label) => {
  if (!plain(value) || Object.keys(value).length !== keys.length || keys.some((key) => !Object.hasOwn(value, key)))
    throw refusal('internal-invalid-field', 'Invalid ' + label + ' fields.');
};
const bounded = (value, max, label) => {
  if (typeof value !== 'string' || !value || value.includes('\0') || bytes(value) > max)
    throw refusal('internal-invalid-field', 'Invalid ' + label + '.');
  return value;
};
const fields = [
  'schema',
  'policy',
  'status',
  'sheet',
  'native',
  'scopeMinMax',
  'range',
  'lastCell',
  'empty',
  'provenance',
  'reasons',
];
const nativeFields = [
  'schema',
  'policy',
  'status',
  'version',
  'sourceRevision',
  'sheet',
  'sheetPath',
  'scopeMinMax',
  'range',
  'lastCell',
  'empty',
  'provenance',
  'reasons',
];
const boundsFields = ['minRow', 'minColumn', 'maxRow', 'maxColumn'];
const integer = (value, max) => Number.isSafeInteger(value) && !Object.is(value, -0) && value >= 0 && value <= max;
const column = (index) => {
  let text = '';
  for (let value = index + 1; value; value = Math.floor((value - 1) / 26))
    text = String.fromCharCode(65 + ((value - 1) % 26)) + text;
  return text;
};
const cell = (row, col) => column(col) + (row + 1);
const boundsRange = (bounds) => {
  const first = cell(bounds.minRow, bounds.minColumn),
    last = cell(bounds.maxRow, bounds.maxColumn);
  return first === last ? first : first + ':' + last;
};
const contains = (outer, inner) =>
  inner === null ||
  (outer.minRow <= inner.minRow &&
    outer.minColumn <= inner.minColumn &&
    outer.maxRow >= inner.maxRow &&
    outer.maxColumn >= inner.maxColumn);
function bounds(value, label, nullable = false) {
  if (nullable && value === null) return;
  exact(value, boundsFields, label);
  if (
    !integer(value.minRow, 1048575) ||
    !integer(value.maxRow, 1048575) ||
    !integer(value.minColumn, 16383) ||
    !integer(value.maxColumn, 16383) ||
    value.minRow > value.maxRow ||
    value.minColumn > value.maxColumn
  )
    throw refusal('internal-invalid-field', 'Invalid ' + label + ' bounds.');
}
export function usedRangeSheet(sheet) {
  bounded(sheet, 124, 'UsedRange worksheet');
  if (sheet.length > 31) throw refusal('internal-invalid-field', 'Invalid UsedRange worksheet.');
  return sheet;
}
export function usedRangeFields(request, { transport = false } = {}) {
  exact(
    request,
    ['operation', 'documentId', 'state', 'sheet', ...(transport ? ['kind', 'sessionId'] : [])],
    'UsedRange request'
  );
  if (request.operation !== 'usedRange' || (transport && request.kind !== 'indexed'))
    throw refusal('internal-invalid-field', 'Invalid UsedRange operation.');
  bounded(request.documentId, 128, 'UsedRange document');
  bounded(request.state, 128, 'UsedRange state');
  usedRangeSheet(request.sheet);
  if (transport) bounded(request.sessionId, 128, 'UsedRange session');
  return request;
}
export function usedRangeReceipt(receipt) {
  exact(receipt, ['documentId', 'ownerId', 'state', 'revision'], 'UsedRange receipt');
  for (const key of ['documentId', 'ownerId', 'state', 'revision'])
    bounded(receipt[key], key === 'revision' ? 512 : 128, 'UsedRange ' + key);
  return receipt;
}
function nativeIdentity(value) {
  exact(value, ['version', 'sourceRevision', 'sheetPath'], 'UsedRange native identity');
  if (
    !integer(value.version, Number.MAX_SAFE_INTEGER) ||
    typeof value.sourceRevision !== 'string' ||
    !/^[a-f0-9]{64}$/.test(value.sourceRevision)
  )
    throw refusal('internal-invalid-field', 'Invalid UsedRange native revision.');
  bounded(value.sheetPath, 1024, 'UsedRange source worksheet path');
  if (
    value.sheetPath.startsWith('/') ||
    value.sheetPath.includes('\\') ||
    value.sheetPath.split('/').some((part) => !part || part === '.' || part === '..')
  )
    throw refusal('layout-malformed', 'Invalid UsedRange source worksheet path.');
}
function provenance(value) {
  exact(
    value,
    [
      'factsSchema',
      'declaration',
      'sourceCells',
      'sourceMerges',
      'effectiveCells',
      'sourceAxisFormatting',
      'acceptedHistory',
      'sourceDefaultFormatting',
    ],
    'UsedRange provenance'
  );
  if (
    !['SCHEMA_USED_RANGE_FACTS_V1', 'SCHEMA_USED_RANGE_FACTS_V2'].some((key) =>
      matchesFormat(value.factsSchema, key)
    ) ||
    typeof value.sourceDefaultFormatting !== 'boolean'
  )
    throw refusal('internal-invalid-field', 'Invalid UsedRange source facts.');
  exact(value.declaration, ['status', 'ref', 'bounds'], 'UsedRange declaration');
  const declaration = value.declaration;
  if (!['absent', 'valid', 'invalid', 'duplicate'].includes(declaration.status))
    throw refusal('internal-invalid-field', 'Invalid UsedRange declaration status.');
  if (declaration.ref !== null) bounded(declaration.ref, 128, 'UsedRange dimension declaration');
  bounds(declaration.bounds, 'UsedRange declaration', true);
  if (
    (declaration.status === 'absent' && (declaration.ref !== null || declaration.bounds !== null)) ||
    (declaration.status === 'valid' && (declaration.ref === null || declaration.bounds === null))
  )
    throw refusal('internal-invalid-field', 'Invalid UsedRange declaration provenance.');
  if (declaration.status === 'valid') {
    const canonical = boundsRange(declaration.bounds),
      single = canonical.includes(':') ? canonical : canonical + ':' + canonical;
    if (declaration.ref !== canonical && declaration.ref !== single)
      throw refusal('layout-limit', 'UsedRange declaration disagrees with its bounds.');
  }
  for (const key of ['sourceCells', 'sourceMerges', 'effectiveCells', 'sourceAxisFormatting', 'acceptedHistory'])
    bounds(value[key], 'UsedRange ' + key, true);
}
export function usedRangeInner(reply, expected) {
  exact(reply, fields, 'UsedRange reply');
  if (reply.schema !== 1 || reply.policy !== 'xlsx-used-range-v1' || !['complete', 'incomplete'].includes(reply.status))
    throw refusal('internal-invalid-field', 'Invalid UsedRange authority.');
  usedRangeSheet(reply.sheet);
  nativeIdentity(reply.native);
  provenance(reply.provenance);
  if (
    (expected?.sheet !== undefined && reply.sheet !== expected.sheet) ||
    (expected?.native && Object.keys(expected.native).some((key) => reply.native[key] !== expected.native[key]))
  )
    throw refusal('layout-stale', 'UsedRange source/version changed.');
  if (!Array.isArray(reply.reasons) || reply.reasons.length > 16)
    throw refusal('internal-invalid-field', 'Invalid UsedRange reasons.');
  for (const reason of reply.reasons) {
    exact(reason, ['code', 'message'], 'UsedRange reason');
    bounded(reason.code, 64, 'UsedRange reason code');
    if (!/^[a-z][a-z0-9-]*$/.test(reason.code))
      throw refusal('internal-invalid-field', 'Invalid UsedRange reason code.');
    bounded(reason.message, 1024, 'UsedRange reason');
  }
  if (reply.status === 'complete') {
    bounds(reply.scopeMinMax, 'UsedRange scope');
    if (
      reply.range !== boundsRange(reply.scopeMinMax) ||
      reply.lastCell !== cell(reply.scopeMinMax.maxRow, reply.scopeMinMax.maxColumn) ||
      typeof reply.empty !== 'boolean' ||
      reply.reasons.length
    )
      throw refusal('internal-invalid-field', 'Invalid complete UsedRange geometry.');
    const authoritative = [
      'effectiveCells',
      'sourceMerges',
      'sourceAxisFormatting',
      'acceptedHistory',
      ...(matchesFormat(reply.provenance.factsSchema, 'SCHEMA_USED_RANGE_FACTS_V2') ? ['sourceCells'] : []),
    ];
    if (authoritative.some((key) => !contains(reply.scopeMinMax, reply.provenance[key])))
      throw refusal('layout-limit', 'Complete UsedRange excludes its authoritative bounds.');
    if (
      reply.empty &&
      (reply.range !== 'A1' ||
        reply.provenance.effectiveCells !== null ||
        reply.provenance.sourceMerges !== null ||
        reply.provenance.sourceAxisFormatting !== null ||
        reply.provenance.acceptedHistory !== null ||
        reply.provenance.sourceDefaultFormatting ||
        (matchesFormat(reply.provenance.factsSchema, 'SCHEMA_USED_RANGE_FACTS_V2') &&
          reply.provenance.sourceCells !== null))
    )
      throw refusal('internal-invalid-field', 'Invalid empty UsedRange authority.');
  } else if (
    reply.scopeMinMax !== null ||
    reply.range !== null ||
    reply.lastCell !== null ||
    reply.empty !== null ||
    !reply.reasons.length
  )
    throw refusal('layout-unsupported', 'Incomplete UsedRange cannot authorize geometry.');
  if (bytes(JSON.stringify(reply)) > 16384)
    throw refusal('layout-limit', 'UsedRange reply exceeds its transport bound.');
  return reply;
}
export function normalizeNativeUsedRange(reply, expected) {
  exact(reply, nativeFields, 'native UsedRange reply');
  if (!matchesFormat(reply.schema, 'SCHEMA_USED_RANGE_V1'))
    throw refusal('internal-invalid-field', 'Invalid native UsedRange schema.');
  const { version, sourceRevision, sheetPath, ...data } = reply;
  return usedRangeInner({ ...data, schema: 1, native: { version, sourceRevision, sheetPath } }, expected);
}
export function bindUsedRange(reply, receipt, expected) {
  usedRangeInner(reply, expected);
  usedRangeReceipt(receipt);
  return { ...reply, receipt: { ...receipt } };
}
export function validateUsedRange(reply, expected) {
  exact(reply, [...fields, 'receipt'], 'bound UsedRange reply');
  const { receipt, ...inner } = reply;
  usedRangeReceipt(receipt);
  usedRangeInner(inner, expected);
  if (expected?.receipt && Object.keys(receipt).some((key) => receipt[key] !== expected.receipt[key]))
    throw refusal('layout-stale', 'UsedRange owner/head has retired.');
  if (bytes(JSON.stringify(reply)) > 16384)
    throw refusal('layout-limit', 'UsedRange reply exceeds its transport bound.');
  return reply;
}
