/** Browser-safe descriptor validation. Storage is not permission or a route to execute. */
import { assertXmlText, rangeBounds, columnName } from './workbook-commands.mjs';
import { normalizeSortKeys } from './sort-keys.mjs';
import { refusal } from './refusal-codes.mjs';
export const LOCAL_ACTION_LIMITS = Object.freeze({
  entries: 128,
  registryBytes: 1024 * 1024,
  descriptorBytes: 64 * 1024,
  nameCharacters: 128,
  nameBytes: 512,
  freezeRows: 100,
  freezeColumns: 30,
  formatPositions: 10000,
  sortPositions: 1000000,
  sortKeys: 16,
});
const bytes = (value) => new TextEncoder().encode(value).length,
  own = (v, k) => Object.hasOwn(v, k);
const plain = (v) =>
  v !== null &&
  typeof v === 'object' &&
  !Array.isArray(v) &&
  [Object.prototype, null].includes(Object.getPrototypeOf(v));
function exact(value, required, optional = [], label = 'local action') {
  if (
    !plain(value) ||
    required.some((key) => !own(value, key)) ||
    Object.keys(value).some((key) => ![...required, ...optional].includes(key))
  )
    throw refusal('internal-invalid-field', 'Invalid ' + label + ' fields');
}
function integer(value, min, max, label) {
  if (!Number.isSafeInteger(value) || value < min || value > max)
    throw refusal('internal-invalid-field', 'Invalid ' + label);
  return value;
}
function name(value) {
  if (typeof value !== 'string' || !value.trim() || value.length > 128 || bytes(value) > 512 || /[\r\n]/.test(value))
    throw refusal('internal-invalid-field', 'Invalid action name');
  assertXmlText(value);
  return value;
}
function sheet(value) {
  if (
    typeof value !== 'string' ||
    !value.trim() ||
    value.length > 31 ||
    /[\\/:?*\[\]]/.test(value) ||
    value.startsWith("'") ||
    value.endsWith("'")
  )
    throw refusal('internal-invalid-field', 'Invalid action worksheet');
  assertXmlText(value);
  return value;
}
function range(value, max) {
  if (typeof value !== 'string' || value.length > 64) throw refusal('internal-invalid-field', 'Invalid action range');
  const b = rangeBounds(value);
  if (b.size > max) throw refusal('query-limit', 'Action range exceeds its position budget');
  const first = columnName(b.first.col) + (b.first.row + 1),
    last = columnName(b.last.col) + (b.last.row + 1);
  return first === last ? first : first + ':' + last;
}
const formatKeys = [
  'bold',
  'italic',
  'underline',
  'fontFamily',
  'fontSize',
  'fontColor',
  'fillColor',
  'horizontalAlign',
  'verticalAlign',
  'wrapText',
  'border',
  'numberFormat',
];
function format(value) {
  exact(value, [], formatKeys, 'action format');
  if (!Object.keys(value).length) throw refusal('query-bad-input', 'Action format requires a change');
  for (const key of ['bold', 'italic', 'underline', 'wrapText'])
    if (own(value, key) && typeof value[key] !== 'boolean')
      throw refusal('internal-invalid-field', 'Invalid action ' + key);
  if (own(value, 'fontFamily')) {
    if (typeof value.fontFamily !== 'string' || !value.fontFamily.trim() || value.fontFamily.length > 100)
      throw refusal('internal-invalid-field', 'Invalid action font family');
    assertXmlText(value.fontFamily);
  }
  if (own(value, 'fontSize') && (!Number.isFinite(value.fontSize) || value.fontSize < 6 || value.fontSize > 72))
    throw refusal('internal-invalid-field', 'Invalid action font size');
  for (const key of ['fontColor', 'fillColor'])
    if (own(value, key) && (typeof value[key] !== 'string' || !/^\p{ASCII_Hex_Digit}{6}$/u.test(value[key])))
      throw refusal('internal-invalid-field', 'Invalid action color');
  for (const [key, allowed] of Object.entries({
    horizontalAlign: ['left', 'center', 'right'],
    verticalAlign: ['top', 'center', 'bottom'],
    border: ['none', 'all'],
    numberFormat: [
      'general',
      'number',
      'percent',
      'date',
      'text',
      'currency',
      'accounting',
      'thousands',
      'decimal0',
      'decimal1',
      'decimal2',
    ],
  }))
    if (own(value, key) && !allowed.includes(value[key]))
      throw refusal('internal-invalid-field', 'Invalid action ' + key);
  return { ...value };
}
function catalog(id, label, command, readScope, writeScope, budgets, prerequisites) {
  return Object.freeze({
    id,
    label,
    command,
    descriptorSchema: 1,
    supportedProfiles: Object.freeze(['xlsx-writable']),
    requiresWrite: true,
    readScope,
    writeScope,
    budgets: Object.freeze(budgets),
    prerequisites: Object.freeze(prerequisites),
    capabilityReason:
      'Stored local definition only; execution is not enabled by this catalog. An executing owner must independently check its captured source, head, profile and prerequisites.',
  });
}
export const LOCAL_ACTION_CATALOG = Object.freeze([
  catalog(
    'freeze-panes',
    'Freeze panes',
    'freeze',
    'accepted worksheet layout',
    'worksheet pane metadata',
    { rows: 100, columns: 30 },
    [
      'Existing worksheet; writable XLSX owner; qualified source layout. Current owner projection enforces 100 rows / 30 columns.',
    ]
  ),
  catalog(
    'format-range',
    'Format range',
    'formatRange',
    'accepted worksheet cells and styles',
    'selected cell formatting',
    { positions: 10000 },
    [
      'Existing worksheet; writable XLSX owner; source/special-cell and style guards. Workbook-relative style IDs are not saved parameters.',
    ]
  ),
  catalog(
    'recalculate-workbook',
    'Recalculate workbook',
    'recalculate',
    'accepted workbook calculation inputs',
    'workbook calculation results',
    { ownerSessionBytes: 4 * 1024 ** 3, projectionBytes: 64 * 1024 ** 2 },
    [
      'Writable XLSX owner; supported calculation policy and complete qualified dependencies under existing native/compatibility budgets.',
    ]
  ),
  catalog(
    'sort-rows',
    'Sort rows',
    'sortRows',
    'accepted typed rows and reference graph',
    'one stable row permutation',
    { positions: 1000000, keys: 16 },
    [
      'Writable XLSX owner; whole used-column span; no active/imported autoFilter; existing source, precision, row metadata and relative-reference qualification. Range explicitly describes body rows.',
    ]
  ),
]);
export function localActionCatalog(command) {
  const entry = LOCAL_ACTION_CATALOG.find((item) => item.command === command);
  if (!entry) throw refusal('query-unsupported', 'Unsupported local action command');
  return entry;
}
export function validateLocalActionDescriptor(value) {
  exact(value, ['schema', 'command', 'parameters'], [], 'action descriptor');
  if (value.schema !== 1) throw refusal('query-unsupported', 'Unsupported action descriptor version');
  localActionCatalog(value.command);
  let parameters;
  if (value.command === 'freeze') {
    exact(value.parameters, ['sheet', 'rows', 'columns'], [], 'freeze parameters');
    parameters = {
      sheet: sheet(value.parameters.sheet),
      rows: integer(value.parameters.rows, 0, 100, 'freeze rows'),
      columns: integer(value.parameters.columns, 0, 30, 'freeze columns'),
    };
  } else if (value.command === 'formatRange') {
    exact(value.parameters, ['sheet', 'range', 'format'], [], 'format parameters');
    parameters = {
      sheet: sheet(value.parameters.sheet),
      range: range(value.parameters.range, 10000),
      format: format(value.parameters.format),
    };
  } else if (value.command === 'recalculate') {
    exact(value.parameters, [], [], 'recalculate parameters');
    parameters = {};
  } else {
    exact(value.parameters, ['sheet', 'range', 'keys'], [], 'sort parameters');
    const selected = range(value.parameters.range, 1000000);
    parameters = {
      sheet: sheet(value.parameters.sheet),
      range: selected,
      keys: normalizeSortKeys({ range: selected, keys: value.parameters.keys }).map((key) => ({ ...key })),
    };
  }
  const result = { schema: 1, command: value.command, parameters };
  if (bytes(JSON.stringify(result)) > LOCAL_ACTION_LIMITS.descriptorBytes)
    throw refusal('query-limit', 'Action descriptor exceeds its byte budget');
  return result;
}
export function validateLocalActionDraft(value) {
  exact(value, ['name', 'descriptor'], [], 'action draft');
  return { name: name(value.name), descriptor: validateLocalActionDescriptor(value.descriptor) };
}
export function localActionId(value) {
  if (typeof value !== 'string' || !/^[a-f0-9]{8}-[a-f0-9]{4}-4[a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}$/.test(value))
    throw refusal('internal-invalid-field', 'Invalid local action id');
  return value;
}
export function localActionRevision(value) {
  return integer(value, 0, Number.MAX_SAFE_INTEGER, 'action registry revision');
}
export function validateLocalActionEntry(value) {
  exact(value, ['id', 'revision', 'name', 'descriptor', 'author', 'createdAt', 'updatedAt'], [], 'action entry');
  if (value.author !== 'local-profile') throw refusal('query-unsupported', 'Unsupported local action author');
  const createdAt = integer(value.createdAt, 0, Number.MAX_SAFE_INTEGER, 'creation timestamp'),
    updatedAt = integer(value.updatedAt, createdAt, Number.MAX_SAFE_INTEGER, 'update timestamp');
  return {
    id: localActionId(value.id),
    revision: integer(value.revision, 1, Number.MAX_SAFE_INTEGER, 'entry revision'),
    ...validateLocalActionDraft({ name: value.name, descriptor: value.descriptor }),
    author: 'local-profile',
    createdAt,
    updatedAt,
  };
}
export function validateLocalActionList(value) {
  exact(value, ['schema', 'scope', 'revision', 'entries'], ['warning'], 'action list');
  if (
    value.schema !== 1 ||
    value.scope !== 'private-profile' ||
    !Array.isArray(value.entries) ||
    value.entries.length > 128
  )
    throw refusal('internal-invalid-field', 'Invalid local action registry');
  const seen = new Set(),
    entries = Array.from(value.entries, (entry) => {
      const checked = validateLocalActionEntry(entry);
      if (seen.has(checked.id)) throw refusal('query-duplicate', 'Duplicate local action id');
      seen.add(checked.id);
      return checked;
    });
  const revision = localActionRevision(value.revision);
  if (entries.some((entry) => entry.revision > revision))
    throw refusal('query-limit', 'Entry revision exceeds registry revision');
  const result = { schema: 1, scope: 'private-profile', revision, entries };
  if (own(value, 'warning')) {
    if (typeof value.warning !== 'string' || !value.warning || bytes(value.warning) > 512)
      throw refusal('internal-invalid-field', 'Invalid registry warning');
    assertXmlText(value.warning);
    result.warning = value.warning;
  }
  if (bytes(JSON.stringify(result)) > LOCAL_ACTION_LIMITS.registryBytes)
    throw refusal('query-limit', 'Action registry exceeds its 1 MiB byte budget');
  return result;
}
