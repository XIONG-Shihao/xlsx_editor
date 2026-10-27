/** Closed readonly draft envelopes. No accepted calculation or history authority. */
import { refusal } from './refusal-codes.mjs';

const plain = (value) => value !== null && typeof value === 'object' && !Array.isArray(value);
const size = (value) => new TextEncoder().encode(value).length;
const own = (value, keys, label) => {
  if (!plain(value) || Object.keys(value).length !== keys.length || keys.some((key) => !Object.hasOwn(value, key)))
    throw refusal('internal-invalid-field', 'Invalid ' + label + ' fields.');
};
const bounded = (value, max, label) => {
  if (typeof value !== 'string' || !value || size(value) > max)
    throw refusal('internal-invalid-field', 'Invalid ' + label + '.');
  return value;
};
export function formulaPreviewRequestId(value) {
  bounded(value, 36, 'formula preview request ID');
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[47][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/.test(value))
    throw refusal('internal-invalid-field', 'Invalid formula preview request ID.');
  return value;
}
export function formulaPreviewTarget(sheet, cell, formula) {
  bounded(sheet, 124, 'formula preview worksheet');
  if (sheet.length > 31) throw refusal('internal-invalid-field', 'Invalid formula preview worksheet.');
  const match = typeof cell === 'string' && /^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(cell);
  if (!match || Number(match[2]) > 1048576 || [...match[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) > 16384)
    throw refusal('formula-bad-input', 'Choose one canonical cell for formula preview.');
  bounded(formula, 8192, 'formula draft');
  if (!formula.trim() || formula.startsWith('=') || formula.includes('\0'))
    throw refusal('internal-invalid-field', 'Invalid formula draft.');
  return { sheet, cell, formula };
}
export function formulaPreviewFields(request, { transport = false, cancel = false } = {}) {
  const fields = [
    'operation',
    'documentId',
    'state',
    'requestId',
    ...(!cancel ? ['sheet', 'cell', 'formula'] : []),
    ...(transport ? ['kind', 'sessionId'] : []),
  ];
  own(request, fields, 'formula preview request');
  if (
    request.operation !== (cancel ? 'cancelFormulaPreview' : 'previewFormula') ||
    (transport && request.kind !== 'indexed')
  )
    throw refusal('internal-invalid-field', 'Invalid formula preview operation.');
  bounded(request.documentId, 128, 'formula preview document');
  bounded(request.state, 128, 'formula preview state');
  formulaPreviewRequestId(request.requestId);
  if (transport) bounded(request.sessionId, 128, 'formula preview session');
  if (!cancel) formulaPreviewTarget(request.sheet, request.cell, request.formula);
  return request;
}
export function formulaPreviewReceipt(receipt) {
  own(receipt, ['documentId', 'ownerId', 'state', 'revision'], 'formula preview receipt');
  for (const key of ['documentId', 'ownerId', 'state', 'revision'])
    bounded(receipt[key], key === 'revision' ? 512 : 128, 'formula preview ' + key);
  return receipt;
}
const codes = new Set(['#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#NULL!', '#SPILL!', '#CALC!']);
export function formulaPreviewInner(reply, expected) {
  own(reply, ['schema', 'requestId', 'sheet', 'cell', 'formula', 'authority', 'result'], 'formula preview reply');
  if (reply.schema !== 1 || reply.authority !== 'draft-preview')
    throw refusal('internal-invalid-field', 'Invalid formula draft authority.');
  formulaPreviewRequestId(reply.requestId);
  formulaPreviewTarget(reply.sheet, reply.cell, reply.formula);
  if (expected && ['requestId', 'sheet', 'cell', 'formula'].some((key) => reply[key] !== expected[key]))
    throw refusal('formula-stale', 'Formula preview reply belongs to another draft.');
  const result = reply.result;
  if (result?.status === 'value') {
    own(result, ['status', 'type', 'value'], 'formula preview scalar');
    if (result.type === 'number') {
      if (typeof result.value !== 'number' || !Number.isFinite(result.value))
        throw refusal('formula-unsupported', 'Formula preview number is unavailable.');
    } else if (result.type === 'boolean') {
      if (typeof result.value !== 'boolean')
        throw refusal('internal-invalid-field', 'Invalid formula preview Boolean.');
    } else if (result.type === 'text') {
      if (typeof result.value !== 'string' || size(result.value) > 32768)
        throw refusal('formula-limit', 'Formula preview text exceeds its limit.');
    } else if (result.type === 'error') {
      if (typeof result.value !== 'string' || !codes.has(result.value))
        throw refusal('internal-invalid-field', 'Invalid formula preview error.');
    } else throw refusal('internal-invalid-field', 'Invalid formula preview scalar type.');
  } else {
    own(result, ['status', 'reason'], 'formula preview outcome');
    if (!['invalid', 'unavailable', 'cancelled'].includes(result.status))
      throw refusal('internal-invalid-field', 'Invalid formula preview outcome.');
    bounded(result.reason, 2048, 'formula preview reason');
  }
  return reply;
}
export function bindFormulaPreview(reply, receipt, expected) {
  formulaPreviewInner(reply, expected);
  formulaPreviewReceipt(receipt);
  return { ...reply, receipt: { ...receipt } };
}
export function validateFormulaPreview(reply, expected) {
  own(
    reply,
    ['schema', 'requestId', 'receipt', 'sheet', 'cell', 'formula', 'authority', 'result'],
    'bound formula preview'
  );
  const { receipt, ...inner } = reply;
  formulaPreviewReceipt(receipt);
  formulaPreviewInner(inner, expected);
  if (expected?.receipt && Object.keys(receipt).some((key) => receipt[key] !== expected.receipt[key]))
    throw refusal('formula-stale', 'Formula preview owner/head has retired.');
  return reply;
}
export function formulaPreviewCancellation(reply, requestId) {
  own(reply, ['schema', 'requestId', 'cancelled'], 'formula preview cancellation');
  if (
    reply.schema !== 1 ||
    reply.requestId !== formulaPreviewRequestId(requestId) ||
    typeof reply.cancelled !== 'boolean'
  )
    throw refusal('internal-invalid-field', 'Invalid formula preview cancellation.');
  return reply;
}
