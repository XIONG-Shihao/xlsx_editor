import { FORMAT, matchesFormat } from './format.mjs';
import { TEXT_CONVERSION_POLICIES } from './text-conversion.mjs';
import { cellPosition, columnName, assertXmlText } from './workbook-commands.mjs';
import { refusal } from './refusal-codes.mjs';

export const PASTE_POSITION_LIMIT = 1_000_000;
export const PASTE_BYTE_LIMIT = 64 * 1024 * 1024;
export const PASTE_PAGE_CELLS = 1000;
export const PASTE_PAGE_BYTES = 1024 * 1024;
export const PASTE_SCHEMA = FORMAT.SCHEMA_STAGED_PASTE_V1;
export const TYPED_PASTE_SCHEMA = FORMAT.SCHEMA_STAGED_PASTE_V2;
export const TYPED_PASTE_POLICY = 'typed-text-v1';
const exactKeys = (value, keys) =>
  value &&
  typeof value === 'object' &&
  !Array.isArray(value) &&
  Object.keys(value).every((key) => keys.includes(key)) &&
  keys.every((key) => Object.hasOwn(value, key));
/** UTF-8 byte accounting without allocating an encoded copy of clipboard text. */
export function pasteBytes(text) {
  let bytes = 0;
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    if (c < 0x80) bytes++;
    else if (c < 0x800) bytes += 2;
    else if (c >= 0xd800 && c <= 0xdbff && text.charCodeAt(i + 1) >= 0xdc00 && text.charCodeAt(i + 1) <= 0xdfff) {
      bytes += 4;
      i++;
    } else bytes += 3;
  }
  return bytes;
}
export const pasteId = (id) =>
  typeof id === 'string' && /^[a-f0-9]{8}-[a-f0-9]{4}-4[a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}$/.test(id);
export function validatePasteState(state) {
  if (typeof state !== 'string' || !state || state.length > 256 || /[\u0000-\u001f]/.test(state))
    throw refusal('internal-invalid-field', 'Invalid paste parent state');
  return state;
}
export function validatePasteShape(shape) {
  if (
    !exactKeys(shape, ['sheet', 'start', 'rows', 'columns']) ||
    typeof shape.sheet !== 'string' ||
    !shape.sheet ||
    shape.sheet.length > 31
  )
    throw refusal('internal-invalid-field', 'Invalid paste shape');
  assertXmlText(shape.sheet);
  if (typeof shape.start !== 'string') throw refusal('internal-invalid-field', 'Invalid paste origin');
  const first = cellPosition(shape.start);
  const start = columnName(first.col) + (first.row + 1);
  if (
    shape.start !== start ||
    !Number.isInteger(shape.rows) ||
    !Number.isInteger(shape.columns) ||
    shape.rows < 1 ||
    shape.columns < 1 ||
    shape.rows * shape.columns > PASTE_POSITION_LIMIT ||
    first.row + shape.rows > 1048576 ||
    first.col + shape.columns > 16384
  )
    throw refusal('paste-limit', 'Paste exceeds workbook or 1,000,000 position limits');
  return { ...shape };
}
/** A dense typed query includes its true output bounds, distinct from the
 * union rectangle used to clear the previous accepted result. */
export function validateTypedPasteConversion(value, shape) {
  if (
    !exactKeys(value, ['schema', 'policies', 'date1904', 'keepHeader', 'outputRows', 'outputColumns']) ||
    value.schema !== 1 ||
    !Array.isArray(value.policies) ||
    value.policies.length < 1 ||
    value.policies.length > 256 ||
    value.policies.some((policy) => !TEXT_CONVERSION_POLICIES.includes(policy)) ||
    typeof value.date1904 !== 'boolean' ||
    typeof value.keepHeader !== 'boolean' ||
    !Number.isInteger(value.outputRows) ||
    value.outputRows < 0 ||
    value.outputRows > shape.rows ||
    !Number.isInteger(value.outputColumns) ||
    value.outputColumns < 1 ||
    value.outputColumns > shape.columns ||
    value.outputColumns !== value.policies.length
  )
    throw refusal('internal-invalid-field', 'Invalid typed query conversion metadata');
  return {
    schema: 1,
    policies: [...value.policies],
    date1904: value.date1904,
    keepHeader: value.keepHeader,
    outputRows: value.outputRows,
    outputColumns: value.outputColumns,
  };
}
export function validatePastePlan(plan) {
  const typed = plan?.policy === TYPED_PASTE_POLICY;
  if (!exactKeys(plan, typed ? ['shape', 'count', 'policy', 'conversion'] : ['shape', 'count', 'policy']))
    throw refusal('internal-invalid-field', 'Invalid paste plan');
  const shape = validatePasteShape(plan.shape);
  if (
    !['grid', 'text', TYPED_PASTE_POLICY].includes(plan.policy) ||
    !Number.isInteger(plan.count) ||
    plan.count < 1 ||
    plan.count > shape.rows * shape.columns ||
    (typed && plan.count !== shape.rows * shape.columns)
  )
    throw refusal('internal-invalid-field', 'Invalid paste count or input policy');
  return {
    shape,
    count: plan.count,
    policy: plan.policy,
    ...(typed ? { conversion: validateTypedPasteConversion(plan.conversion, shape) } : {}),
  };
}
export function validatePasteDescriptor(descriptor) {
  const typed = descriptor?.policy === TYPED_PASTE_POLICY;
  if (
    !exactKeys(
      descriptor,
      typed
        ? ['type', 'id', 'sha256', 'count', 'bytes', 'shape', 'policy', 'schema', 'conversion']
        : ['type', 'id', 'sha256', 'count', 'bytes', 'shape', 'policy']
    ) ||
    descriptor.type !== 'importCells' ||
    !pasteId(descriptor.id) ||
    (typed && !matchesFormat(descriptor.schema, 'SCHEMA_STAGED_PASTE_V2')) ||
    typeof descriptor.sha256 !== 'string' ||
    !/^[a-f0-9]{64}$/.test(descriptor.sha256) ||
    !Number.isSafeInteger(descriptor.bytes) ||
    descriptor.bytes < 1 ||
    descriptor.bytes > PASTE_BYTE_LIMIT
  )
    throw refusal('paste-malformed', 'Invalid staged paste descriptor');
  return {
    ...descriptor,
    ...validatePastePlan({
      shape: descriptor.shape,
      count: descriptor.count,
      policy: descriptor.policy,
      ...(typed ? { conversion: descriptor.conversion } : {}),
    }),
  };
}
/** New headers use neutral tags; the descriptor field ordering stays stable. */
export function pasteHeader(plan, documentId, state) {
  const checked = validatePastePlan(plan);
  return {
    schema: checked.policy === TYPED_PASTE_POLICY ? TYPED_PASTE_SCHEMA : PASTE_SCHEMA,
    documentId,
    state,
    ...checked,
  };
}
export function pasteDescriptorIdentity(value) {
  const d = validatePasteDescriptor(value),
    legacy = [
      d.type,
      d.id,
      d.sha256,
      d.count,
      d.bytes,
      d.shape.sheet,
      d.shape.start,
      d.shape.rows,
      d.shape.columns,
      d.policy,
    ];
  return JSON.stringify(d.policy === TYPED_PASTE_POLICY ? [...legacy, d.schema, d.conversion] : legacy);
}
export function validatePasteCells(cells, shape, previous = -1) {
  if (!Array.isArray(cells) || cells.length < 1 || cells.length > PASTE_PAGE_CELLS)
    throw refusal('paste-bad-input', 'Paste pages require 1 to 1,000 cells');
  let bytes = 0;
  for (const entry of cells) {
    if (
      !Array.isArray(entry) ||
      entry.length !== 2 ||
      !Number.isInteger(entry[0]) ||
      entry[0] <= previous ||
      entry[0] >= shape.rows * shape.columns ||
      typeof entry[1] !== 'string'
    )
      throw refusal('paste-bad-input', 'Paste offsets must be increasing inside the declared shape');
    if (entry[1].length > 32768) throw refusal('paste-limit', 'Paste field exceeds cell text limits');
    assertXmlText(entry[1]);
    bytes += pasteBytes(JSON.stringify(entry) + '\n');
    if (bytes > PASTE_PAGE_BYTES) throw refusal('paste-limit', 'Paste page exceeds 1 MiB');
    previous = entry[0];
  }
  return { lastOffset: previous, bytes };
}
