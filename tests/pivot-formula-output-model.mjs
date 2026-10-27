/** Pure pivot admission regression; native producer/independent admission is separate. */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createPivotModel } from '../core/pivot-model.mjs';
const input = () => ({
  sheet: 'Source',
  range: 'A2:B4',
  rowField: 'Group',
  valueField: 'Amount',
  cells: {
    A2: { value: 'Group' },
    B2: { value: 'Amount' },
    A3: { value: 'A' },
    B3: { value: 2 },
    A4: { value: 'B' },
    B4: { value: 3 },
  },
});
test('known cached header/body outputs refuse even when their master lies outside the pivot range', () => {
  for (const target of ['A2', 'A3', 'B3'])
    for (const calculation of [undefined, { status: 'calculated', value: 2, revision: 'accepted' }]) {
      const source = input();
      source.cells[target] = {
        ...source.cells[target],
        formulaOutput: { schema: 1, master: 'A1', range: 'A1:B4' },
        ...(calculation ? { calculation } : {}),
      };
      const before = structuredClone(source);
      assert.throws(() => createPivotModel(source), { code: 'pivot-bad-input' });
      assert.deepEqual(source, before);
    }
});
test('ordinary current sources retain the same qualified groups and exact totals', () => {
  const source = input(),
    before = structuredClone(source),
    result = createPivotModel(source);
  assert.equal(result.total, 5);
  assert.deepEqual(
    result.groups.map(({ label, value }) => ({ label, value })),
    [
      { label: 'A', value: 2 },
      { label: 'B', value: 3 },
    ]
  );
  assert.deepEqual(source, before);
});
