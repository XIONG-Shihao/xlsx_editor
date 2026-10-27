import { FORMAT, matchesFormat } from './format.mjs';
/** Closed hyperlink commands and bound read DTOs; no file, shell or network access. */
import { refusal } from './refusal-codes.mjs';

const record = (value, keys, label) => {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.keys(value).sort().join(',') !== [...keys].sort().join(',')
  )
    throw refusal('internal-invalid-field', 'Invalid ' + label + ' fields.');
};
const text = (value, max, label) => {
  if (
    typeof value !== 'string' ||
    value.length > max ||
    (typeof value.isWellFormed === 'function' && !value.isWellFormed()) ||
    /[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/.test(value)
  )
    throw refusal('internal-invalid-field', 'Invalid ' + label + '.');
  return value;
};
const hash = (value) => typeof value === 'string' && /^[a-f0-9]{64}$/.test(value);
export function hyperlinkCell(value) {
  if (typeof value !== 'string' || !/^([A-Z]{1,3})([1-9]\d*)$/.test(value))
    throw refusal('hyperlink-bad-input', 'Select a canonical single cell for a hyperlink.');
  const [, letters, row] = /^([A-Z]{1,3})([1-9]\d*)$/.exec(value);
  let column = 0;
  for (const c of letters) column = column * 26 + c.charCodeAt(0) - 64;
  if (column > 16384 || Number(row) > 1048576)
    throw refusal('hyperlink-unsupported', 'Hyperlink cell is outside the worksheet.');
  return value;
}
function sheet(value) {
  text(value, 31, 'hyperlink worksheet');
  if (!value.trim() || /[\[\]:*?\\/]/.test(value))
    throw refusal('internal-invalid-field', 'Invalid hyperlink worksheet.');
  return value;
}
export function hyperlinkDestination(value) {
  if (value?.kind === 'url') {
    record(value, ['kind', 'target'], 'web hyperlink');
    text(value.target, 2048, 'web hyperlink');
    if (!/^https?:\/\//.test(value.target) || /[\s\\]/u.test(value.target))
      throw refusal('hyperlink-bad-input', 'Use an ordinary HTTP(S) web address.');
    let url;
    try {
      url = new URL(value.target);
    } catch {
      throw refusal('internal-invalid-field', 'Invalid web address.');
    }
    if (!url.hostname || url.username || url.password || !['http:', 'https:'].includes(url.protocol))
      throw refusal('hyperlink-bad-input', 'Use a web address without embedded credentials.');
    if (/%(?![a-f0-9]{2})/i.test(value.target)) throw refusal('hyperlink-malformed', 'Invalid web address escape.');
    return { kind: 'url', target: value.target };
  }
  if (value?.kind === 'internal') {
    record(value, ['kind', 'sheet', 'cell'], 'internal hyperlink');
    return { kind: 'internal', sheet: sheet(value.sheet), cell: hyperlinkCell(value.cell) };
  }
  throw refusal('hyperlink-bad-input', 'Choose a web or worksheet-cell destination.');
}
export function hyperlinkCommand(value) {
  const remove = value?.type === 'removeHyperlink';
  const keys = [
    'type',
    'sheet',
    'cell',
    'expectedHyperlinkRevision',
    ...(!remove ? ['link', ...(Object.hasOwn(value ?? {}, 'tooltip') ? ['tooltip'] : [])] : []),
  ];
  record(value, keys, 'hyperlink command');
  if (!remove && value.type !== 'setHyperlink')
    throw refusal('hyperlink-unsupported', 'Unsupported hyperlink command.');
  if (value.expectedHyperlinkRevision !== null && !hash(value.expectedHyperlinkRevision))
    throw refusal('internal-invalid-field', 'Invalid prior hyperlink revision.');
  return {
    type: value.type,
    sheet: sheet(value.sheet),
    cell: hyperlinkCell(value.cell),
    expectedHyperlinkRevision: value.expectedHyperlinkRevision,
    ...(!remove
      ? {
          link: hyperlinkDestination(value.link),
          ...(Object.hasOwn(value, 'tooltip') ? { tooltip: text(value.tooltip, 256, 'hyperlink tooltip') } : {}),
        }
      : {}),
  };
}
export function normalizeHyperlinkCommands(commands, view) {
  if (!Array.isArray(commands) || !commands.length)
    throw refusal('internal-missing-field', 'Missing hyperlink command batch.');
  let found = false;
  const normalized = commands.map((value) => {
    if (!['setHyperlink', 'removeHyperlink'].includes(value?.type)) return structuredClone(value);
    found = true;
    const command = hyperlinkCommand(value);
    if (
      view &&
      (!view.sheets.some((s) => s.name === command.sheet) ||
        (command.link?.kind === 'internal' && !view.sheets.some((s) => s.name === command.link.sheet)))
    )
      throw refusal('internal-missing-field', 'Hyperlink worksheet is missing.');
    return command;
  });
  if (found && new TextEncoder().encode(JSON.stringify(normalized)).byteLength > 2 * 1024 * 1024)
    throw refusal('hyperlink-limit', 'Hyperlink batch exceeds the 2 MiB action bound.');
  return normalized;
}
export function hyperlinkFields(request, { transport = false } = {}) {
  record(
    request,
    ['operation', 'documentId', 'state', 'sheet', 'cell', ...(transport ? ['kind', 'sessionId'] : [])],
    'hyperlink request'
  );
  if (request.operation !== 'hyperlink' || (transport && request.kind !== 'indexed'))
    throw refusal('internal-invalid-field', 'Invalid hyperlink operation.');
  for (const key of ['documentId', 'state', ...(transport ? ['sessionId'] : [])])
    if (!text(request[key], 128, 'hyperlink ' + key).trim())
      throw refusal('internal-missing-field', 'Missing hyperlink identity.');
  sheet(request.sheet);
  hyperlinkCell(request.cell);
  return request;
}
const fields = ['schema', 'sheet', 'cell', 'hyperlinkRevision', 'link', 'tooltip', 'canEdit', 'protectedReason'];
function definition(value, expected) {
  record(value, fields, 'hyperlink definition');
  if (
    !matchesFormat(value.schema, 'SCHEMA_HYPERLINK_DEFINITION_V1') ||
    value.sheet !== expected.sheet ||
    value.cell !== expected.cell
  )
    throw refusal('hyperlink-stale', 'Hyperlink reply changed the captured target.');
  if (
    (value.hyperlinkRevision !== null && !hash(value.hyperlinkRevision)) ||
    typeof value.canEdit !== 'boolean' ||
    (value.protectedReason !== null && typeof value.protectedReason !== 'string')
  )
    throw refusal('internal-invalid-field', 'Invalid hyperlink definition.');
  if (
    value.canEdit !== (value.protectedReason === null) ||
    (typeof value.protectedReason === 'string' && !value.protectedReason.trim())
  )
    throw refusal('hyperlink-unsupported', 'Hyperlink edit authority is inconsistent.');
  text(value.tooltip, 256, 'hyperlink tooltip');
  if (value.protectedReason !== null) text(value.protectedReason, 512, 'hyperlink refusal');
  if (value.link !== null) hyperlinkDestination(value.link);
  if (
    (value.link !== null && value.hyperlinkRevision === null) ||
    (value.canEdit && value.hyperlinkRevision !== null && value.link === null)
  )
    throw refusal('internal-mismatch', 'Hyperlink identity and target disagree.');
  return value;
}
function receipt(value) {
  record(value, ['documentId', 'ownerId', 'state', 'revision'], 'hyperlink receipt');
  for (const key of Object.keys(value))
    if (!text(value[key], key === 'revision' ? 512 : 128, 'hyperlink ' + key).trim())
      throw refusal('internal-missing-field', 'Missing hyperlink receipt.');
  return value;
}
export function bindHyperlink(value, binding, expected) {
  receipt(binding);
  definition(value, expected);
  return { ...structuredClone(value), ...binding };
}
export function validateHyperlink(value, binding, expected) {
  receipt(binding);
  record(value, [...fields, 'documentId', 'ownerId', 'state', 'revision'], 'bound hyperlink');
  for (const key of Object.keys(binding))
    if (value[key] !== binding[key])
      throw refusal('hyperlink-stale', 'Hyperlink reply belongs to a retired owner/head.');
  const raw = Object.fromEntries(fields.map((key) => [key, value[key]]));
  definition(raw, expected);
  return value;
}

export function hyperlinkDisplayText(value) {
  return text(value, 32767, 'link display text');
}
