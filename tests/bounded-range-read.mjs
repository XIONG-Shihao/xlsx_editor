import { test } from 'node:test';
import assert from 'node:assert/strict';
import { boundedRangeRead } from '../ui/boundedRangeRead.ts';
import { rangeBounds } from '../core/workbook-commands.mjs';
test('passive viewport tile and long navigation reads use bounded pages with exact typed values', async () => {
  const calls = [],
    read = boundedRangeRead(async (sheet, range) => {
      const b = rangeBounds(range);
      assert(b.size <= 1000);
      calls.push(range);
      return range.startsWith('A1:')
        ? {
            A1: { formula: 'MYFUNC()', calculation: { status: 'stored', value: 42 } },
            B1: { value: -0 },
            C1: { value: '00123' },
          }
        : {};
    });
  const cells = await read('Sheet1', 'A1:AF32');
  assert.equal(calls.length, 2);
  assert.equal(cells.A1.calculation.value, 42);
  assert.equal(cells.C1.value, '00123');
  assert(Object.is(cells.B1.value, -0));
  calls.length = 0;
  await read('Sheet1', 'A1:A1024');
  assert.equal(calls.length, 2);
  calls.length = 0;
  await read('Sheet1', 'A1:AMJ1');
  assert.equal(calls.length, 2);
});
test('passive combined reply and geometry bounds refuse without silently changing reader limits', async () => {
  let calls = 0;
  const oversized = boundedRangeRead(async () => {
    calls++;
    return { A1: { value: 'x'.repeat(1024 * 1024) } };
  });
  await assert.rejects(oversized('Sheet1', 'A1:AF32'), /1 MiB/);
  assert.equal(calls, 1);
  await assert.rejects(oversized('Sheet1', 'A1:ZZ100'), /position/);
  assert.equal(calls, 1);
});
