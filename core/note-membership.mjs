/** Closed read-only source-backed legacy-note membership. */
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
const fields = ['schema', 'policy', 'status', 'sheet', 'source', 'notes', 'threadedPlaceholders', 'reasons'];
export const NOTE_MEMBERSHIP_MAX_ANNOTATIONS = 100000;
export const NOTE_MEMBERSHIP_MAX_REPLY_BYTES = 4 * 1024 * 1024;
export function noteMembershipSheet(sheet) {
  bounded(sheet, 124, 'note membership worksheet');
  if (sheet.length > 31) throw refusal('internal-invalid-field', 'Invalid note membership worksheet.');
  return sheet;
}
export function noteMembershipFields(request, { transport = false } = {}) {
  exact(
    request,
    ['operation', 'documentId', 'state', 'sheet', ...(transport ? ['kind', 'sessionId'] : [])],
    'note membership request'
  );
  if (request.operation !== 'noteMembership' || (transport && request.kind !== 'indexed'))
    throw refusal('internal-invalid-field', 'Invalid note membership operation.');
  bounded(request.documentId, 128, 'note membership document');
  bounded(request.state, 128, 'note membership state');
  noteMembershipSheet(request.sheet);
  if (transport) bounded(request.sessionId, 128, 'note membership session');
  return request;
}
export function noteMembershipReceipt(receipt) {
  exact(receipt, ['documentId', 'ownerId', 'state', 'revision'], 'note membership receipt');
  for (const key of ['documentId', 'ownerId', 'state', 'revision'])
    bounded(receipt[key], key === 'revision' ? 512 : 128, 'note membership ' + key);
  return receipt;
}
export function noteMembershipSource(source) {
  exact(source, ['version', 'sourceRevision', 'sheetPath'], 'note membership source');
  if (
    !Number.isSafeInteger(source.version) ||
    Object.is(source.version, -0) ||
    source.version < 0 ||
    typeof source.sourceRevision !== 'string' ||
    !/^[a-f0-9]{64}$/.test(source.sourceRevision)
  )
    throw refusal('internal-invalid-field', 'Invalid note membership source revision.');
  bounded(source.sheetPath, 1024, 'note membership source worksheet path');
  if (
    source.sheetPath.startsWith('/') ||
    source.sheetPath.includes('\\') ||
    source.sheetPath.split('/').some((part) => !part || part === '.' || part === '..')
  )
    throw refusal('note-malformed', 'Invalid note membership source worksheet path.');
  return source;
}
/** Canonical A1 membership references have no range, dollars, or leading zeroes. */
export function noteMembershipCell(ref) {
  if (typeof ref !== 'string') throw refusal('internal-invalid-field', 'Invalid note membership cell reference.');
  const match = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(ref);
  if (!match) throw refusal('internal-invalid-field', 'Invalid note membership cell reference.');
  const column = [...match[1]].reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0),
    row = Number(match[2]);
  if (column > 16384 || row > 1048576) throw refusal('note-limit', 'Note membership cell is outside worksheet limits.');
  return (row - 1) * 16384 + column - 1;
}
function inventory(value) {
  if (!Array.isArray(value) || value.length > NOTE_MEMBERSHIP_MAX_ANNOTATIONS)
    throw refusal('internal-invalid-field', 'Invalid note membership inventory.');
  let prior = -1;
  for (const ref of value) {
    const position = noteMembershipCell(ref);
    if (position <= prior) throw refusal('note-bad-input', 'Note membership inventory must be sorted and unique.');
    prior = position;
  }
}
export function noteMembershipInner(reply, expected) {
  exact(reply, fields, 'note membership reply');
  if (
    reply.schema !== 1 ||
    reply.policy !== 'xlsx-legacy-note-membership-v1' ||
    !['complete', 'incomplete'].includes(reply.status)
  )
    throw refusal('internal-invalid-field', 'Invalid note membership authority.');
  noteMembershipSheet(reply.sheet);
  noteMembershipSource(reply.source);
  if (
    (expected?.sheet !== undefined && reply.sheet !== expected.sheet) ||
    (expected?.source && Object.keys(expected.source).some((key) => reply.source[key] !== expected.source[key]))
  )
    throw refusal('note-stale', 'Note membership source/version changed.');
  if (!Array.isArray(reply.reasons) || reply.reasons.length > 16)
    throw refusal('internal-invalid-field', 'Invalid note membership reasons.');
  for (const reason of reply.reasons) {
    exact(reason, ['code', 'message'], 'note membership reason');
    bounded(reason.code, 64, 'note membership reason code');
    if (!/^[a-z][a-z0-9-]*$/.test(reason.code))
      throw refusal('internal-invalid-field', 'Invalid note membership reason code.');
    bounded(reason.message, 1024, 'note membership reason');
  }
  if (reply.status === 'complete') {
    inventory(reply.notes);
    inventory(reply.threadedPlaceholders);
    if (reply.notes.length + reply.threadedPlaceholders.length > NOTE_MEMBERSHIP_MAX_ANNOTATIONS)
      throw refusal('note-limit', 'Note membership inventory exceeds its resource limit.');
    const notes = new Set(reply.notes);
    if (reply.threadedPlaceholders.some((ref) => notes.has(ref)) || reply.reasons.length)
      throw refusal('internal-invalid-field', 'Invalid complete note membership.');
  } else if (reply.notes !== null || reply.threadedPlaceholders !== null || !reply.reasons.length)
    throw refusal('note-unsupported', 'Incomplete note membership cannot authorize partial inventory.');
  if (bytes(JSON.stringify(reply)) > NOTE_MEMBERSHIP_MAX_REPLY_BYTES)
    throw refusal('note-limit', 'Note membership reply exceeds its transport resource limit.');
  return reply;
}
export function bindNoteMembership(reply, receipt, expected) {
  noteMembershipInner(reply, expected);
  noteMembershipReceipt(receipt);
  return validateNoteMembership({ ...reply, receipt: { ...receipt } }, expected);
}
export function validateNoteMembership(reply, expected) {
  exact(reply, [...fields, 'receipt'], 'bound note membership reply');
  const { receipt, ...inner } = reply;
  noteMembershipReceipt(receipt);
  noteMembershipInner(inner, expected);
  if (expected?.receipt && Object.keys(receipt).some((key) => receipt[key] !== expected.receipt[key]))
    throw refusal('note-stale', 'Note membership owner/head has retired.');
  if (bytes(JSON.stringify(reply)) > NOTE_MEMBERSHIP_MAX_REPLY_BYTES)
    throw refusal('note-limit', 'Note membership reply exceeds its transport resource limit.');
  return reply;
}
