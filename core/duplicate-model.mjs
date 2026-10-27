/** Closed duplicate configuration/receipt. Row maps remain native-only. */
import { refusal } from './refusal-codes.mjs';

const exact = (value, keys) =>
  !!value &&
  typeof value === 'object' &&
  !Array.isArray(value) &&
  Object.keys(value).length === keys.length &&
  keys.every((k) => Object.hasOwn(value, k));
const integer = (value, max = 1000000) => Number.isSafeInteger(value) && value >= 0 && value <= max;
const col = (name) => [...name].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
export function duplicateConfiguration(sheet, range, hasHeaders, columns) {
  if (typeof sheet !== 'string' || !sheet || sheet.length > 31 || typeof range !== 'string')
    throw refusal('internal-invalid-field', 'Invalid duplicate worksheet or range');
  const m = /^([A-Z]{1,3})([1-9]\d*)(?::([A-Z]{1,3})([1-9]\d*))?$/.exec(range);
  if (!m) throw refusal('internal-invalid-field', 'Invalid duplicate range');
  const first = { row: Number(m[2]) - 1, col: col(m[1]) },
    last = { row: Number(m[4] ?? m[2]) - 1, col: col(m[3] ?? m[1]) };
  if (
    first.row > last.row ||
    first.col > last.col ||
    last.row >= 1048576 ||
    last.col >= 16384 ||
    (last.row - first.row + 1) * (last.col - first.col + 1) > 1000000
  )
    throw refusal('duplicates-duplicate', 'Duplicate removal is limited to 1,000,000 selected positions');
  if (typeof hasHeaders !== 'boolean' || first.row + Number(hasHeaders) > last.row)
    throw refusal('duplicates-bad-input', 'Choose a header and at least one data row');
  if (
    !Array.isArray(columns) ||
    columns.length < 1 ||
    columns.length > 16 ||
    new Set(columns).size !== columns.length ||
    columns.some((c) => typeof c !== 'string' || !/^[A-Z]{1,3}$/.test(c) || col(c) < first.col || col(c) > last.col)
  )
    throw refusal('duplicates-duplicate', 'Choose 1–16 distinct in-range duplicate key columns');
  return { sheet, range, hasHeaders, columns: [...columns], first, last, body: first.row + Number(hasHeaders) };
}
export function duplicateCommand(command) {
  if (
    !exact(command, ['operation', 'sheet', 'range', 'hasHeaders', 'columns', 'expectedFingerprint']) ||
    command.operation !== 'removeDuplicates' ||
    typeof command.expectedFingerprint !== 'string' ||
    !/^[a-f0-9]{64}$/.test(command.expectedFingerprint)
  )
    throw refusal('internal-invalid-field', 'Invalid duplicate removal command fields');
  duplicateConfiguration(command.sheet, command.range, command.hasHeaders, command.columns);
  return { ...command, columns: [...command.columns] };
}
export function duplicatePage(reply, config, offset = 0) {
  const c = duplicateConfiguration(config.sheet, config.range, config.hasHeaders, config.columns),
    total = c.last.row - c.body + 1;
  if (
    !integer(offset) ||
    !exact(reply, [
      'fingerprint',
      'totalRows',
      'keptRows',
      'removedRows',
      'comparison',
      'samples',
      'offset',
      'nextOffset',
    ]) ||
    !/^[a-f0-9]{64}$/.test(reply.fingerprint) ||
    reply.comparison !== 'exact-typed-values-v1' ||
    reply.totalRows !== total ||
    !integer(reply.keptRows, total) ||
    reply.keptRows < 1 ||
    reply.removedRows !== total - reply.keptRows ||
    reply.offset !== offset ||
    !Array.isArray(reply.samples) ||
    reply.samples.length > 100 ||
    (offset > reply.removedRows && reply.samples.length) ||
    reply.samples.length !== Math.min(100, Math.max(0, reply.removedRows - offset))
  )
    throw refusal('internal-invalid-field', 'Invalid duplicate preview receipt');
  let previous = 0;
  for (const sample of reply.samples) {
    if (
      !exact(sample, ['row', 'keptRow']) ||
      !integer(sample.row, 1048576) ||
      sample.row < c.body + 1 ||
      sample.row > c.last.row + 1 ||
      sample.row <= previous ||
      !integer(sample.keptRow, 1048576) ||
      sample.keptRow < c.body + 1 ||
      sample.keptRow >= sample.row
    )
      throw refusal('internal-invalid-field', 'Invalid duplicate preview sample');
    previous = sample.row;
  }
  const next = offset + reply.samples.length < reply.removedRows ? offset + reply.samples.length : null;
  if (reply.nextOffset !== next) throw refusal('internal-invalid-field', 'Invalid duplicate preview continuation');
  return reply;
}
export function bindDuplicatePreview(raw, receipt, config, offset = 0) {
  duplicatePage(raw, config, offset);
  if (
    !exact(receipt, ['documentId', 'ownerId', 'state', 'revision']) ||
    Object.values(receipt).some((v) => typeof v !== 'string' || !v)
  )
    throw refusal('internal-invalid-field', 'Invalid duplicate preview owner receipt');
  return {
    schema: 1,
    ...receipt,
    sheet: config.sheet,
    range: config.range,
    hasHeaders: config.hasHeaders,
    columns: [...config.columns],
    ...raw,
  };
}
export function validateDuplicatePreview(reply, receipt, config, offset = 0) {
  if (
    !exact(reply, [
      'schema',
      'documentId',
      'ownerId',
      'state',
      'revision',
      'sheet',
      'range',
      'hasHeaders',
      'columns',
      'fingerprint',
      'totalRows',
      'keptRows',
      'removedRows',
      'comparison',
      'samples',
      'offset',
      'nextOffset',
    ]) ||
    reply.schema !== 1 ||
    Object.keys(receipt).some((k) => reply[k] !== receipt[k]) ||
    reply.sheet !== config.sheet ||
    reply.range !== config.range ||
    reply.hasHeaders !== config.hasHeaders ||
    !Array.isArray(reply.columns) ||
    JSON.stringify(reply.columns) !== JSON.stringify(config.columns)
  )
    throw refusal('duplicates-duplicate', 'Duplicate preview owner/configuration changed');
  const raw = Object.fromEntries(
    ['fingerprint', 'totalRows', 'keptRows', 'removedRows', 'comparison', 'samples', 'offset', 'nextOffset'].map(
      (k) => [k, reply[k]]
    )
  );
  duplicatePage(raw, config, offset);
  return reply;
}
