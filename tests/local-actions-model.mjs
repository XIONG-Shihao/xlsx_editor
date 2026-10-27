import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  LOCAL_ACTION_CATALOG,
  validateLocalActionDescriptor as descriptor,
  validateLocalActionDraft as draft,
  validateLocalActionEntry as entry,
  validateLocalActionList as list,
} from '../core/local-actions.mjs';
const id = '11111111-1111-4111-8111-111111111111',
  recalc = { schema: 1, command: 'recalculate', parameters: {} };
test('catalog local labels declare source prerequisites without an execution mapper', () => {
  assert.deepEqual(
    LOCAL_ACTION_CATALOG.map((row) => row.id),
    ['freeze-panes', 'format-range', 'recalculate-workbook', 'sort-rows']
  );
  for (const row of LOCAL_ACTION_CATALOG) {
    assert.deepEqual(row.supportedProfiles, ['xlsx-writable']);
    assert(row.requiresWrite);
    assert(row.readScope && row.writeScope && row.capabilityReason.includes('not enabled'));
    assert(row.prerequisites.length);
    assert(Object.isFrozen(row));
    assert.equal(row.run, undefined);
  }
});
test('four closed descriptors preserve false formatting and ordered mixed-direction keys', () => {
  const freeze = descriptor({
    schema: 1,
    command: 'freeze',
    parameters: { sheet: "O'Brien Data", rows: 100, columns: 30 },
  });
  assert.equal(freeze.parameters.rows, 100);
  const input = {
    schema: 1,
    command: 'formatRange',
    parameters: {
      sheet: 'Data',
      range: '$a$1:$b$2',
      format: { bold: false, fontSize: 12.5, fontColor: 'Aa00fF', numberFormat: 'date' },
    },
  };
  const value = descriptor(input);
  assert.equal(value.parameters.range, 'A1:B2');
  assert.equal(value.parameters.format.bold, false);
  value.parameters.format.bold = true;
  assert.equal(input.parameters.format.bold, false);
  const sort = descriptor({
    schema: 1,
    command: 'sortRows',
    parameters: {
      sheet: 'Data',
      range: 'a1:p62500',
      keys: [
        { column: 'b', direction: 'desc' },
        { column: 'a', direction: 'asc' },
      ],
    },
  });
  assert.equal(sort.parameters.range, 'A1:P62500');
  assert.deepEqual(sort.parameters.keys, [
    { column: 'B', direction: 'desc' },
    { column: 'A', direction: 'asc' },
  ]);
  assert.deepEqual(descriptor(recalc), recalc);
});
test('unsupported versions, script/path/actor fields, untyped format and ambiguous sort refuse', () => {
  const invalid = [
    { ...recalc, schema: 2 },
    { ...recalc, command: 'eval' },
    { ...recalc, path: '/tmp/run.js' },
    { ...recalc, parameters: { actor: 'local-profile' } },
    { schema: 1, command: 'freeze', parameters: { sheet: 'S', rows: 101, columns: 0 } },
    { schema: 1, command: 'freeze', parameters: { sheet: 'S', rows: 0, columns: 31 } },
    { schema: 1, command: 'formatRange', parameters: { sheet: 'S', range: 'A1', format: { styleId: 1 } } },
    { schema: 1, command: 'formatRange', parameters: { sheet: 'S', range: 'A1', format: { bold: 'true' } } },
    {
      schema: 1,
      command: 'sortRows',
      parameters: {
        sheet: 'S',
        range: 'A1:C5',
        keys: [
          { column: 'A', direction: 'asc' },
          { column: 'a', direction: 'desc' },
        ],
      },
    },
    {
      schema: 1,
      command: 'sortRows',
      parameters: { sheet: 'S', range: 'A1:C5', keys: [{ column: 'D', direction: 'asc' }] },
    },
    {
      schema: 1,
      command: 'sortRows',
      parameters: { sheet: 'S', range: 'A1:C5', keys: [{ column: 'A', direction: 'asc', source: 'caller' }] },
    },
    {
      schema: 1,
      command: 'sortRows',
      parameters: {
        sheet: 'S',
        range: 'A1:C5',
        keys: [{ column: 'A', direction: 'asc' }],
        column: 'A',
        direction: 'asc',
      },
    },
  ];
  for (const value of invalid) assert.throws(() => descriptor(value));
});
test('range and key boundaries are exact and names preserve Unicode without malformed text', () => {
  assert(
    descriptor({
      schema: 1,
      command: 'formatRange',
      parameters: { sheet: 'S', range: 'A1:J1000', format: { wrapText: true } },
    })
  );
  assert.throws(() =>
    descriptor({
      schema: 1,
      command: 'formatRange',
      parameters: { sheet: 'S', range: 'A1:J1001', format: { wrapText: true } },
    })
  );
  const keys = Array.from({ length: 16 }, (_, i) => ({ column: String.fromCharCode(65 + i), direction: 'asc' }));
  assert(descriptor({ schema: 1, command: 'sortRows', parameters: { sheet: 'S', range: 'A1:P62500', keys } }));
  assert.throws(() =>
    descriptor({
      schema: 1,
      command: 'sortRows',
      parameters: { sheet: 'S', range: 'A1:Q1', keys: [...keys, { column: 'Q', direction: 'desc' }] },
    })
  );
  assert.throws(() =>
    descriptor({ schema: 1, command: 'sortRows', parameters: { sheet: 'S', range: 'A1:P62501', keys } })
  );
  assert.equal(draft({ name: 'Résumé 🙂', descriptor: recalc }).name, 'Résumé 🙂');
  for (const name of [' ', 'a'.repeat(129), 'x\u0000', 'x\ud800'])
    assert.throws(() => draft({ name, descriptor: recalc }));
});
test('entry/list identity, author, timestamps and closed schema reject forged grants', () => {
  const record = {
    id,
    revision: 1,
    name: 'Recalculate',
    descriptor: recalc,
    author: 'local-profile',
    createdAt: 10,
    updatedAt: 11,
  };
  assert.deepEqual(entry(record), record);
  const saved = { schema: 1, scope: 'private-profile', revision: 1, entries: [record] };
  assert.deepEqual(list(saved), saved);
  for (const bad of [
    { ...record, author: 'project-admin' },
    { ...record, createdAt: 12 },
    { ...record, filename: '/tmp/a.xlsx' },
    { ...record, id: 'forged' },
    { ...record, revision: 0 },
  ])
    assert.throws(() => entry(bad));
  for (const bad of [
    { ...saved, ownerId: id },
    { ...saved, scope: 'shared-project' },
    { ...saved, schema: 2 },
    { ...saved, entries: [record, record] },
    { ...saved, revision: 0 },
    { ...saved, entries: Array(129).fill(record) },
  ])
    assert.throws(() => list(bad));
  const copy = list(saved);
  copy.entries[0].descriptor.parameters.actor = 'forged';
  assert.deepEqual(saved.entries[0].descriptor.parameters, {});
});
