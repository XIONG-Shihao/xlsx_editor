import { test } from 'node:test';
import assert from 'node:assert/strict';
import { queryRefreshAllRequest } from '../adapters/query-refresh-all-request.mjs';
const id = '11111111-1111-4111-8111-111111111111',
  base = () => ({ operation: 'preview', documentId: 'a'.repeat(64), state: 'base', requestId: id });
test('closed preview/run envelopes are copied and immutable before queueing', () => {
  const body = base(),
    p = queryRefreshAllRequest(body);
  body.state = 'later';
  assert.equal(p.state, 'base');
  assert(Object.isFrozen(p));
  assert.equal(queryRefreshAllRequest({ ...base(), operation: 'refresh', previewId: id }).previewId, id);
});
test('forged paths, actor/owner/target fields, missing identities and wrong operations refuse', () => {
  for (const body of [
    { ...base(), filename: '/private/foreign.csv' },
    { ...base(), actor: 'admin' },
    { ...base(), ownerId: id },
    { ...base(), connectionId: id },
    { ...base(), operation: 'run' },
    { ...base(), operation: 'refresh' },
    { ...base(), previewId: id },
    { ...base(), requestId: 'guess' },
    { ...base(), state: NaN },
    { ...base(), documentId: 'forged' },
  ])
    assert.throws(() => queryRefreshAllRequest(body));
});
test('getters/prototypes/symbols/nonenumerable fields never execute or enter a captured request', () => {
  let reads = 0;
  const accessor = { ...base() };
  Object.defineProperty(accessor, 'state', {
    get() {
      reads++;
      return 'base';
    },
    enumerable: true,
  });
  assert.throws(() => queryRefreshAllRequest(accessor));
  assert.equal(reads, 0);
  for (const body of [
    Object.assign(Object.create({ actor: 'admin' }), base()),
    Object.assign([], base()),
    Object.defineProperty({ ...base() }, 'hidden', { value: 1 }),
    { ...base(), [Symbol('actor')]: 'admin' },
    { ...base(), toJSON: () => base() },
  ])
    assert.throws(() => queryRefreshAllRequest(body));
});
