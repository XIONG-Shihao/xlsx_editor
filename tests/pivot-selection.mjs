import { test } from 'node:test';
import assert from 'node:assert/strict';
import { selectedPivot, pivotContextKey, pivotRefreshReason, refreshPivotContext } from '../ui/pivotSelection.ts';
const pivot = {
  name: 'SalesPivot',
  sheet: 'Summary',
  sourceSheet: 'Data',
  sourceRange: 'A1:B5',
  outputRange: 'A1:B4',
  rowField: 0,
  valueField: 1,
  fieldNames: ['Category', 'Amount'],
  aggregate: 'sum',
  revision: 'pivot-1',
  canEdit: true,
};
const frame = (key = 'owner/head/epoch0') => ({
  key,
  pivot: structuredClone(pivot),
  active: true,
  normal: true,
  readOnly: false,
  busy: false,
  accept: async () => {},
});
test('selection maps only a complete regular rectangle in one bound output, including offscreen cells', () => {
  for (const range of ['A1', 'B4', 'A2:B3']) assert.equal(selectedPivot([pivot], 'Summary', range), pivot);
  for (const [sheet, range, regular] of [
    ['Data', 'A1', true],
    ['Summary', 'C1', true],
    ['Summary', 'A1:C1', true],
    ['Summary', 'A1', false],
    ['Summary', 'A0', true],
  ])
    assert.equal(selectedPivot([pivot], sheet, range, regular), null);
  assert.equal(selectedPivot([pivot, { ...pivot, name: 'Other' }], 'Summary', 'A1'), null);
  assert.equal(selectedPivot([{ ...pivot, outputRange: 'Z900:AA902' }], 'Summary', 'AA901')?.name, 'SalesPivot');
});
test('known protected output retains inspectable selection and an explicit refresh reason', () => {
  const protectedPivot = { ...pivot, canEdit: false, revision: null, reason: 'Imported pivot is preserved.' };
  assert.equal(selectedPivot([protectedPivot], 'Summary', 'B2'), protectedPivot);
  assert.equal(pivotRefreshReason({ ...frame(), pivot: protectedPivot }), protectedPivot.reason);
  for (const [patch, reason] of [
    [{ readOnly: true }, /read-only/],
    [{ busy: true }, /Wait/],
    [{ active: false }, /Activate/],
    [{ normal: false }, /Normal/],
    [{ accept: undefined }, /unavailable/],
  ])
    assert.match(pivotRefreshReason({ ...frame(), ...patch }), reason);
});
test('context identity includes owner/head/activation/navigation and the displayed definition', () => {
  const a = pivotContextKey(['owner', 'head', 'A1', 0], pivot);
  for (const [owner, p] of [
    [['owner', 'head', 'A1', 1], pivot],
    [['other', 'head', 'A1', 0], pivot],
    [['owner', 'next', 'A1', 0], pivot],
    [['owner', 'head', 'A1', 0], { ...pivot, revision: 'pivot-2' }],
    [['owner', 'head', 'A1', 0], { ...pivot, fieldNames: ['Changed', 'Amount'] }],
  ])
    assert.notEqual(pivotContextKey(owner, p), a);
});
test('refresh awaits one captured strict batch; retirement after acceptance keeps the original accepted truth', async () => {
  let resolve, current;
  const writes = [],
    captured = frame();
  captured.accept = (commands) => {
    writes.push(commands);
    return new Promise((r) => (resolve = r));
  };
  current = captured;
  const work = refreshPivotContext(captured, () => current);
  assert.deepEqual(writes, [
    [{ operation: 'refreshPivot', sheet: 'Summary', name: 'SalesPivot', expectedPivotRevision: 'pivot-1' }],
  ]);
  current = frame('other');
  resolve();
  assert.deepEqual(await work, { accepted: true, key: captured.key, current: false });
});
test('retired, permission-changed and unknown-definition callbacks refuse before any batch', async () => {
  let calls = 0;
  const original = frame();
  original.accept = async () => calls++;
  for (const current of [null, frame('changed'), { ...original, readOnly: true }, { ...original, busy: true }])
    await assert.rejects(refreshPivotContext(original, () => current));
  assert.equal(calls, 0);
  const protectedFrame = { ...original, pivot: { ...pivot, canEdit: false } };
  await assert.rejects(refreshPivotContext(protectedFrame, () => protectedFrame));
  assert.equal(calls, 0);
});
