/** Closed, copied envelopes captured before the main shell queue. */
import { refusal } from '../core/refusal-codes.mjs';

const uuid = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/;
const hash = /^[0-9a-f]{64}$/;
export function queryRefreshAllRequest(body) {
  if (
    !body ||
    typeof body !== 'object' ||
    Array.isArray(body) ||
    ![Object.prototype, null].includes(Object.getPrototypeOf(body))
  )
    throw refusal('internal-invalid-field', 'Invalid Refresh All request');
  const keys = Reflect.ownKeys(body),
    descriptors = Object.getOwnPropertyDescriptors(body);
  if (
    keys.some(
      (key) => typeof key !== 'string' || !descriptors[key].enumerable || !Object.hasOwn(descriptors[key], 'value')
    )
  )
    throw refusal('query-bad-input', 'Refresh All requires plain data fields');
  const operation = descriptors.operation?.value,
    fields =
      operation === 'preview'
        ? ['operation', 'documentId', 'state', 'requestId']
        : operation === 'refresh'
          ? ['operation', 'documentId', 'state', 'requestId', 'previewId']
          : null;
  if (!fields || keys.length !== fields.length || fields.some((key) => !Object.hasOwn(descriptors, key)))
    throw refusal('internal-invalid-field', 'Invalid Refresh All request fields');
  const input = Object.fromEntries(fields.map((key) => [key, descriptors[key].value]));
  if (typeof input.documentId !== 'string' || !(uuid.test(input.documentId) || hash.test(input.documentId)))
    throw refusal('query-malformed', 'Invalid Refresh All document identity');
  if (typeof input.state !== 'string' || !(input.state === 'base' || uuid.test(input.state) || hash.test(input.state)))
    throw refusal('internal-invalid-field', 'Invalid Refresh All accepted state');
  if (
    typeof input.requestId !== 'string' ||
    !uuid.test(input.requestId) ||
    (operation === 'refresh' && (typeof input.previewId !== 'string' || !uuid.test(input.previewId)))
  )
    throw refusal('query-malformed', 'Invalid Refresh All request/preview identity');
  return Object.freeze(input);
}
