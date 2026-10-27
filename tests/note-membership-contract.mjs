/** Closed source-inventory contract; no Excel selector observations. */
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  noteMembershipFields,
  noteMembershipSheet,
  noteMembershipInner,
  noteMembershipReceipt,
  bindNoteMembership,
  validateNoteMembership,
} from '../core/note-membership.mjs';
const source = { version: 7, sourceRevision: 'a'.repeat(64), sheetPath: 'xl/worksheets/sheet1.xml' };
const receipt = {
  documentId: 'document',
  ownerId: 'original-owner',
  state: 'accepted-head',
  revision: 'public:owner:3',
};
const complete = () => ({
  schema: 1,
  policy: 'xlsx-legacy-note-membership-v1',
  status: 'complete',
  sheet: 'Data',
  source: { ...source },
  notes: ['Z1', 'A2', 'XFD1048576'],
  threadedPlaceholders: ['B2'],
  reasons: [],
});
const expected = { sheet: 'Data', source, receipt };
test('membership requests have a closed read-only envelope', () => {
  const request = { operation: 'noteMembership', documentId: receipt.documentId, state: receipt.state, sheet: 'Data' };
  assert.equal(noteMembershipFields(request), request);
  noteMembershipFields({ ...request, kind: 'indexed', sessionId: receipt.ownerId }, { transport: true });
  for (const field of [
    'range',
    'cell',
    'path',
    'version',
    'source',
    'revision',
    'ownerId',
    'commands',
    'metadata',
    'receipt',
  ])
    assert.throws(() => noteMembershipFields({ ...request, [field]: 'injected' }), /fields/);
  assert.throws(() => noteMembershipFields({ ...request, kind: 'indexed' }), /fields/);
  assert.throws(() => noteMembershipFields({ ...request, operation: 'read' }), /operation/);
  assert.throws(
    () => noteMembershipFields({ ...request, kind: 'csv', sessionId: 'owner' }, { transport: true }),
    /operation/
  );
  for (const sheet of ['', 'a'.repeat(32), 'Data\0', null])
    assert.throws(() => noteMembershipSheet(sheet), /worksheet/);
  for (const key of ['documentId', 'state'])
    assert.throws(() => noteMembershipFields({ ...request, [key]: 'x'.repeat(129) }));
});
test('selected source identity is separate from exact original owner/head receipt', () => {
  const raw = complete(),
    result = bindNoteMembership(raw, receipt, { sheet: 'Data', source });
  validateNoteMembership(result, expected);
  assert.deepEqual(result.source, source);
  assert.deepEqual(result.receipt, receipt);
  assert.notEqual(result.receipt, receipt);
  for (const key of Object.keys(source))
    assert.throws(() =>
      noteMembershipInner({ ...complete(), source: { ...source, [key]: key === 'version' ? 8 : 'other' } }, { source })
    );
  assert.throws(() => noteMembershipInner({ ...complete(), sheet: 'Other' }, expected), /changed/);
  for (const key of Object.keys(receipt))
    assert.throws(() => validateNoteMembership(result, { receipt: { ...receipt, [key]: 'other' } }), /retired/);
  for (const key of Object.keys(receipt)) {
    const invalid = { ...receipt };
    delete invalid[key];
    assert.throws(() => noteMembershipReceipt(invalid), /fields/);
  }
  assert.throws(() => noteMembershipReceipt({ ...receipt, source }), /fields/);
  assert.throws(() => noteMembershipInner({ ...complete(), receipt }), /fields/);
  assert.throws(() => validateNoteMembership({ ...result, cache: {} }), /fields/);
});
test('complete inventories are canonical, row-major, unique, and disjoint', () => {
  noteMembershipInner(complete());
  for (const notes of [
    ['A2', 'Z1'],
    ['A1', 'A1'],
    ['a1'],
    ['$A$1'],
    ['A0'],
    ['A01'],
    ['XFE1'],
    ['A1048577'],
    ['A1:B2'],
    ['Data!A1'],
    [1],
  ])
    assert.throws(() => noteMembershipInner({ ...complete(), notes }));
  assert.throws(() => noteMembershipInner({ ...complete(), threadedPlaceholders: ['A2'] }), /complete/);
  for (const patch of [
    { schema: 2 },
    { policy: 'notes' },
    { status: 'partial' },
    { source: { ...source, extra: true } },
    { source: { ...source, version: -0 } },
    { source: { ...source, version: -1 } },
    { source: { ...source, sourceRevision: 'A'.repeat(64) } },
    { source: { ...source, sheetPath: '../sheet.xml' } },
  ])
    assert.throws(() => noteMembershipInner({ ...complete(), ...patch }));
  assert.throws(
    () => noteMembershipInner({ ...complete(), reasons: [{ code: 'some-reason', message: 'incomplete' }] }),
    /complete/
  );
});
test('incomplete receipts carry reasons and never authorize empty or partial arrays', () => {
  const raw = {
    ...complete(),
    status: 'incomplete',
    notes: null,
    threadedPlaceholders: null,
    reasons: [{ code: 'missing-person', message: 'The persons graph is missing.' }],
  };
  const result = bindNoteMembership(raw, receipt);
  validateNoteMembership(result, expected);
  for (const patch of [{ notes: [] }, { threadedPlaceholders: [] }, { notes: ['A1'] }, { reasons: [] }])
    assert.throws(() => noteMembershipInner({ ...raw, ...patch }), { code: /^(layout-unsupported|note-unsupported)$/ });
  for (const reasons of [
    null,
    [{ code: 'Bad_code', message: 'reason' }],
    [{ code: 'reason', message: '' }],
    [{ code: 'reason', message: 'x'.repeat(1025) }],
    [{ code: 'reason', message: 'reason', partial: true }],
    Array.from({ length: 17 }, () => ({ code: 'reason', message: 'reason' })),
  ])
    assert.throws(() => noteMembershipInner({ ...raw, reasons }));
  assert.equal(noteMembershipInner({ ...complete(), notes: [], threadedPlaceholders: [] }).status, 'complete');
});
test('full inventories have an explicit 100000 annotation resource bound', () => {
  const refs = Array.from({ length: 100000 }, (_, i) => 'A' + (i + 1));
  noteMembershipInner({ ...complete(), notes: refs, threadedPlaceholders: [] });
  assert.throws(
    () => noteMembershipInner({ ...complete(), notes: [...refs, 'A100001'], threadedPlaceholders: [] }),
    /inventory/
  );
  assert.throws(() => noteMembershipInner({ ...complete(), notes: refs, threadedPlaceholders: ['B1'] }), {
    code: /^(note-limit|package-limit)$/,
  });
});
