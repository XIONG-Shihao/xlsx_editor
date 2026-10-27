import { test } from 'node:test';
import assert from 'node:assert/strict';
import { randomUUID } from 'node:crypto';
import { WorkbookTabRegistry } from '../adapters/workbook-tab-registry.mjs';
const record = (filename, documentId = randomUUID()) => ({
  id: randomUUID(),
  owner: {},
  filename,
  status: { documentId },
});
test('four bounded owners retain order; duplicate path/identity and fifth admission preserve all records', () => {
  const registry = new WorkbookTabRegistry(),
    records = Array.from({ length: 4 }, (_, index) => registry.add(record('/' + index + '.xlsx')));
  assert.throws(() => registry.add(record('/fifth.xlsx')), { code: 'session-refused' });
  assert.deepEqual(registry.records(), records);
  assert.equal(registry.active, records[3]);
  const small = new WorkbookTabRegistry();
  small.add(records[0]);
  assert.throws(() => small.add(record(records[0].filename)), /already/);
  assert.throws(() => small.add(record('/other.xlsx', records[0].status.documentId)), /already/);
  assert.equal(small.size, 1);
});
test('replacement renews handle/epoch without changing tab; stale handles cannot act on a later worker', () => {
  const registry = new WorkbookTabRegistry(),
    first = registry.add(record('/first.xlsx')),
    second = registry.add(record('/second.xlsx'));
  registry.activate(first.id);
  const next = registry.replace(first, record(first.filename, first.status.documentId));
  assert.equal(next.tabId, first.tabId);
  assert.equal(next.epoch, 2);
  assert.equal(registry.active, next);
  assert.throws(() => registry.selected(first.id), /not granted/);
  assert.equal(registry.selected(second.id), second);
  assert.throws(() => registry.replace(first, record('/third.xlsx')), /changed/);
  assert.equal(registry.size, 2);
});
test('destination ownership and close-one are independent of active tab', () => {
  const registry = new WorkbookTabRegistry(),
    first = registry.add(record('/first.xlsx')),
    second = registry.add(record('/second.xlsx'));
  assert.throws(() => registry.assertDestination(first, second.filename), {
    code: /^(formula-stale|session-unsupported)$/,
  });
  registry.assertDestination(first, first.filename);
  registry.assertDestination(first, '/copy.xlsx');
  registry.remove(first);
  assert.equal(registry.active, second);
  assert.equal(registry.size, 1);
  assert.throws(() => registry.selected(first.id), /not granted/);
  registry.remove(second);
  assert.equal(registry.active, null);
  assert.equal(registry.size, 0);
});
