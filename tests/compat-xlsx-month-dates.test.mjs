import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { readWorkbook } from '../core/workbook-package.mjs';
import { calculateWorkbook, createCalculationSession } from '../core/calculation-node.mjs';
import { calculateDateScalar } from '../core/date-calculation.mjs';
for (const system of [1900, 1904])
  test(`portable EDATE/EOMONTH exactly match fresh ${system} numeric scripting reads and separately observed error labels`, async () => {
    const oracle = JSON.parse(
        await fs.readFile(new URL(`../fixtures/excel-month-${system}-reference.json`, import.meta.url), 'utf8')
      ),
      source = await fs.readFile(new URL(`../fixtures/${oracle.source}`, import.meta.url));
    assert.equal(createHash('sha256').update(source).digest('hex'), oracle.sourceSha256);
    const input = readWorkbook(source),
      before = structuredClone(input),
      output = await calculateWorkbook(input);
    for (const [ref, item] of Object.entries(oracle.cells)) {
      const result = output.sheets[0].cells[ref].calculation;
      if (item.dataType === 'number') {
        assert.equal(item.valueClass, 'real');
        assert.equal(result.status, 'calculated', ref);
        assert.equal(result.value, item.expected, ref);
      } else {
        assert.equal(item.qualification, 'missing scripting value plus displayed error; no saved OOXML typed cache');
        assert.equal(result.status, 'error', ref);
        assert.equal(result.error, item.observedError, ref);
      }
    }
    assert.deepEqual(input, before);
  });
const view = (cells) => ({
  revision: 'months',
  sheets: [{ name: 'Cases', date1904: false, rows: 40, columns: 10, cells }],
});
test('monthly unqualified coercion/array/origin modes propagate through lazy branches, names and dependencies', async () => {
  const input = view({
    A1: { value: '45306', dataType: 'text' },
    B1: { formula: 'EDATE(A1,1)' },
    B2: { formula: 'IFERROR(B1,99)' },
    B3: { formula: 'EDATE(TRUE,1)' },
    B4: { formula: 'INDEX(EDATE({45306,45307},1),1,1)' },
    B5: { formula: 'EOMONTH(0,1)' },
    B6: { formula: 'EDATE(60,0)' },
    B7: { formula: 'IFERROR(EDATE(45306,120001),99)' },
    C1: { formula: 'MonthName+1' },
    D1: { formula: 'EDATE(DATE(2024,1,15),1)' },
    D2: { formula: 'D1+1' },
  });
  input.namedRanges = [{ name: 'MonthName', formula: 'EOMONTH(Cases!$A$1,1)' }];
  const output = await calculateWorkbook(input);
  for (const ref of ['B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'C1'])
    assert.equal(output.sheets[0].cells[ref].calculation.status, 'unresolved', ref);
  assert.equal(output.sheets[0].cells.D1.calculation.value, 45337);
  assert.equal(output.sheets[0].cells.D2.calculation.value, 45338);
  assert.match(calculateDateScalar('EDATE', [30, -2], true).unresolved, /before calendar origin/);
});
test('incremental month-date owners keep old states and match cold results after numeric/text transitions', async () => {
  const input = view({
      A1: { value: 45306, dataType: 'number' },
      B1: { formula: 'EDATE(A1,1)' },
      C1: { formula: 'EOMONTH(A1,1)' },
      D1: { formula: 'B1+C1' },
      E1: { formula: 'IFERROR(B1,99)' },
    }),
    session = await createCalculationSession(input),
    accepted = structuredClone(session.view);
  try {
    for (const value of [45322, '45322', 45306, 0, 60, 45306]) {
      const next = structuredClone(input);
      next.revision = 'monthly:' + value + typeof value;
      next.sheets[0].cells.A1 = { value, dataType: typeof value === 'number' ? 'number' : 'text' };
      const patches = await session.update(next, [{ sheet: 'Cases', cell: 'A1' }]);
      for (const patch of patches) next.sheets[0].cells[patch.cell] = patch.value;
      const cold = await calculateWorkbook(next);
      for (const ref of ['B1', 'C1', 'D1', 'E1'])
        assert.deepEqual(next.sheets[0].cells[ref].calculation, cold.sheets[0].cells[ref].calculation, ref);
    }
    assert.deepEqual(session.view, accepted);
  } finally {
    session.dispose();
  }
});
