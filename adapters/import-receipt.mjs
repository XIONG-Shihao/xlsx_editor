import { validatePasteDescriptor, validatePasteState, pasteDescriptorIdentity } from '../core/paste-protocol.mjs';
import { refusal } from '../core/refusal-codes.mjs';
const identity = pasteDescriptorIdentity;
export function validateImportReceiptRequest(request) {
  if (
    !request ||
    Object.keys(request).some(
      (key) => !['kind', 'operation', 'sessionId', 'documentId', 'state', 'descriptor', 'parentState'].includes(key)
    ) ||
    request.operation !== 'importReceipt'
  )
    throw refusal('internal-invalid-field', 'Invalid scoped import receipt request');
  validatePasteState(request.parentState);
  return validatePasteDescriptor(request.descriptor);
}
/** Callers supply only hash/schema/parent-verified accepted history, including
 * its retained redo tip. Absence is deliberately not a rejection receipt. */
export function findImportReceipt(nodes, request, { baseState, currentState, tipState }) {
  const expected = validateImportReceiptRequest(request);
  let found;
  for (const { head, parent, commands } of nodes)
    for (const command of commands) {
      if (command?.type !== 'importCells' || command.id !== expected.id) continue;
      const descriptor = validatePasteDescriptor(command),
        parentState = parent ?? baseState;
      if (identity(descriptor) !== identity(expected) || parentState !== request.parentState || found)
        throw refusal('package-stale', 'Accepted import receipt identity, hash or parent changed');
      found = { found: true, head, parentState, descriptor };
    }
  return { ...(found ?? { found: false, reason: 'not-retained' }), currentState, tipState };
}
