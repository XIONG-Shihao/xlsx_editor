import test from 'node:test';
import assert from 'node:assert/strict';
import { traceReceipt, readTraceCursor } from '../core/indexed-trace.mjs';
const key = Buffer.alloc(32, 7),
  binding = {
    ownerId: 'owner',
    state: '0',
    version: 0,
    revision: 'revision',
    sheet: 'Sheet1',
    cell: 'A1',
    direction: 'dependents',
    sheets: ['Sheet1'],
  };
const reply = {
  version: 0,
  sheet: 'Sheet1',
  cell: 'A1',
  direction: 'dependents',
  edges: [],
  markers: [],
  next: { sheetIndex: 0, after: 'B1', offset: 0 },
  scannedFormulas: 1,
  unknown: false,
  limited: false,
};
test('signed trace cursor admits exact receipt and rejects equivalent noncanonical signature bytes', () => {
  const receipt = traceReceipt(key, binding, { pages: 0, scanned: 0, unknown: false }, reply, 100),
    token = receipt.nextCursor,
    parts = token.split('.');
  assert.equal(readTraceCursor(key, token, binding).cursor.after, 'B1');
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_',
    last = parts[1].at(-1),
    position = alphabet.indexOf(last);
  assert.equal(position % 4, 0);
  const altered = parts[1].slice(0, -1) + alphabet[position + 1];
  assert.deepEqual(
    Buffer.from(altered, 'base64url'),
    Buffer.from(parts[1], 'base64url'),
    'Unused base64 padding bits preserve the same HMAC bytes'
  );
  assert.throws(() => readTraceCursor(key, parts[0] + '.' + altered, binding), {
    code: /^(internal-invalid-field|rules-limit|spill-unsupported)$/,
  });
  assert.throws(() => readTraceCursor(key, token, { ...binding, state: '1' }), /stale/);
  assert.throws(() => readTraceCursor(Buffer.alloc(32, 8), token, binding), { code: /^(formula-stale|note-refused)$/ });
});
