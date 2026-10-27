/** Controlled host-envelope checks; actual native/app qualification is separate. */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createWorkbookReviewControls } from '../ui/reviewHostAdapter.ts';
const base = { documentId: 'document-A', ownerId: 'owner-A', state: 'head-A', revision: 'revision-A' };
const wordRequest = () => ({ requestId: crypto.randomUUID(), locale: 'en-US', words: ['helo'] });
test('host callbacks capture their original receipt and copy only spelling fields', async () => {
  const supplied = { ...base },
    calls = [];
  const controls = createWorkbookReviewControls(supplied, async (request) => {
    calls.push(request);
    return {
      spellingCheck: {
        schema: 1,
        receipt: { ...base },
        result: { schema: 1, requestId: request.requestId, locale: request.locale, results: [] },
      },
    };
  });
  supplied.ownerId = 'replacement-owner';
  const words = wordRequest();
  await controls.check({ ...words, ownerId: 'injected-owner', state: 'injected-head' });
  assert.deepEqual(controls.receipt, base);
  assert(Object.isFrozen(controls.receipt));
  assert.deepEqual(calls, [{ operation: 'spellingCheck', ...base, ...words }]);
  assert.notEqual(calls[0].words, words.words);
});
test('late or malformed owner, request and locale replies refuse', async () => {
  for (const changed of [
    { receipt: { ...base, ownerId: 'owner-B' } },
    { receipt: { ...base, state: 'head-B' } },
    { receipt: { ...base, revision: 'revision-B' } },
    { receipt: { ...base, extra: true } },
    { schema: 2 },
    { result: { requestId: 'another-request', locale: 'en-US' } },
    { result: { locale: 'en-GB' } },
  ]) {
    const request = wordRequest(),
      controls = createWorkbookReviewControls(base, async () => ({
        spellingCheck: {
          schema: 1,
          receipt: { ...base },
          result: { requestId: request.requestId, locale: request.locale },
          ...changed,
        },
      }));
    await assert.rejects(controls.check(request));
  }
  for (const receipt of [
    { ...base, extra: true },
    { ...base, ownerId: '' },
    { ...base, revision: null },
  ])
    assert.throws(() => createWorkbookReviewControls(receipt, async () => ({})));
});
test('availability and cancellation use original owner and exact request acknowledgements', async () => {
  const calls = [],
    availability = { schema: 1, provider: 'nspell' },
    controls = createWorkbookReviewControls(base, async (request) => {
      calls.push(request);
      return request.operation === 'spellingAvailability'
        ? { spellingAvailability: { schema: 1, requestId: request.requestId, receipt: { ...base }, availability } }
        : { spellingCancellation: { schema: 1, requestId: request.requestId, cancelled: false } };
    });
  assert.deepEqual(await controls.availability(), availability);
  const requestId = crypto.randomUUID();
  await controls.cancel(requestId);
  for (const call of calls) for (const [key, value] of Object.entries(base)) assert.equal(call[key], value);
  assert.equal(calls[1].requestId, requestId);
  assert.equal(calls[1].operation, 'spellingCancel');
  for (const bad of [{ requestId: 'another' }, { cancelled: 'false' }, { schema: 2 }])
    await assert.rejects(
      createWorkbookReviewControls(base, async () => ({
        spellingCancellation: { schema: 1, requestId, cancelled: false, ...bad },
      })).cancel(requestId)
    );
});
