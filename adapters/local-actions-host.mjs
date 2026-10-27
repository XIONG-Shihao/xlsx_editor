/** Closed main-route CRUD admission. Trusted owner checks and private storage
 * selection belong to the host; library management never executes a workbook. */
import { validateLocalActionDraft, localActionId, localActionRevision } from '../core/local-actions.mjs';
import { refusal } from '../core/refusal-codes.mjs';

export const LOCAL_ACTION_REQUEST_BYTES = 16 * 1024;
const uuid = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/;
const hash = /^[0-9a-f]{64}$/;
const fields = {
  list: ['operation', 'documentId', 'state'],
  create: ['operation', 'documentId', 'state', 'expectedRevision', 'draft'],
  update: ['operation', 'documentId', 'state', 'expectedRevision', 'actionId', 'draft'],
  remove: ['operation', 'documentId', 'state', 'expectedRevision', 'actionId'],
};
const plain = (value) =>
  value !== null &&
  typeof value === 'object' &&
  !Array.isArray(value) &&
  [Object.prototype, null].includes(Object.getPrototypeOf(value));

// Inspect only own data properties, then copy before any shell/registry queue.
// Counting escaped keys/scalars and JSON punctuation bounds the actual UTF-8
// envelope without invoking getters, toJSON, inherited aliases or a serializer.
function capture(body) {
  let bytes = 0;
  const ancestors = new Set();
  const add = (n) => {
    bytes += n;
    if (bytes > LOCAL_ACTION_REQUEST_BYTES)
      throw refusal('query-limit', 'Local action request exceeds its 16 KiB UTF-8 byte budget');
  };
  const encoded = (value) => {
    if (typeof value === 'string' && value.length > LOCAL_ACTION_REQUEST_BYTES)
      throw refusal('query-limit', 'Local action request exceeds its 16 KiB UTF-8 byte budget');
    add(new TextEncoder().encode(JSON.stringify(value)).length);
  };
  function copy(value, depth = 0) {
    if (
      value === null ||
      typeof value === 'boolean' ||
      typeof value === 'string' ||
      (typeof value === 'number' && Number.isFinite(value))
    ) {
      encoded(value);
      return Object.is(value, -0) ? 0 : value;
    }
    if (typeof value !== 'object' || depth > 16 || ancestors.has(value))
      throw refusal('query-bad-input', 'Local action request must contain independent JSON data only');
    const array = Array.isArray(value);
    if (array ? Object.getPrototypeOf(value) !== Array.prototype : !plain(value))
      throw refusal('internal-invalid-field', 'Invalid local action request prototype');
    const keys = Reflect.ownKeys(value),
      descriptors = Object.getOwnPropertyDescriptors(value);
    if (
      keys.some(
        (key) =>
          typeof key !== 'string' ||
          (key !== 'length' && (!descriptors[key].enumerable || !Object.hasOwn(descriptors[key], 'value')))
      )
    )
      throw refusal('query-refused', 'Local action request contains a hidden, symbolic or accessor field');
    if (
      array &&
      (keys.length !== value.length + 1 ||
        keys.some((key) => key !== 'length' && (!/^(0|[1-9][0-9]*)$/.test(key) || Number(key) >= value.length)))
    )
      throw refusal('query-bad-input', 'Local action request arrays must be dense and contain no aliases');
    ancestors.add(value);
    add(2);
    const result = array ? [] : Object.create(null),
      names = array ? Array.from({ length: value.length }, (_, i) => String(i)) : keys;
    for (const [index, key] of names.entries()) {
      if (index) add(1);
      if (!array) {
        encoded(key);
        add(1);
      }
      result[key] = copy(descriptors[key].value, depth + 1);
    }
    ancestors.delete(value);
    return result;
  }
  return copy(body);
}
function freeze(value) {
  if (value && typeof value === 'object') {
    for (const child of Object.values(value)) freeze(child);
    Object.freeze(value);
  }
  return value;
}

export function localActionRequest(body) {
  if (!plain(body)) throw refusal('internal-invalid-field', 'Invalid closed local action request');
  const input = capture(body),
    required =
      typeof input.operation === 'string' && Object.hasOwn(fields, input.operation) ? fields[input.operation] : null;
  if (!required || Object.keys(input).length !== required.length || required.some((key) => !Object.hasOwn(input, key)))
    throw refusal('internal-invalid-field', 'Invalid closed local action request fields');
  if (typeof input.documentId !== 'string' || !(uuid.test(input.documentId) || hash.test(input.documentId)))
    throw refusal('internal-invalid-field', 'Invalid local action document id');
  if (typeof input.state !== 'string' || !(input.state === 'base' || uuid.test(input.state) || hash.test(input.state)))
    throw refusal('internal-invalid-field', 'Invalid local action accepted state');
  const request = { operation: input.operation, documentId: input.documentId, state: input.state };
  if (input.operation !== 'list') request.expectedRevision = localActionRevision(input.expectedRevision);
  if (input.operation === 'update' || input.operation === 'remove') request.actionId = localActionId(input.actionId);
  if (input.operation === 'create' || input.operation === 'update')
    request.draft = validateLocalActionDraft(input.draft);
  return freeze(request);
}

export function invokeLocalActionRequest(store, body) {
  const request = localActionRequest(body);
  if (request.operation === 'list') return store.list();
  if (request.operation === 'create') return store.create(request.draft, request.expectedRevision);
  if (request.operation === 'update') return store.update(request.actionId, request.draft, request.expectedRevision);
  return store.remove(request.actionId, request.expectedRevision);
}
