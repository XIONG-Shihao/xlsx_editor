/** Incremental recalculation matches a cold calculation where financial formulas overlap protected shapes. */
import assert from 'node:assert/strict';
import { test as check } from 'node:test';
import { calculateWorkbook, createCalculationSession } from '../core/calculation-node.mjs';
const cases = [
  { name: '1904-date-overlap', calendar: true, formula: 'PMT(A1,12,100)+WEEKDAY(D1)' },
  {
    name: 'array-geometry-overlap',
    formula: 'PMT(A1,12,100)',
    special: { formulaType: 'array', formulaReference: 'B1:C3' },
    consumer: 'SUM(C2)+A1',
  },
  {
    name: 'invalid-financial-array-geometry',
    initial: { value: '0', dataType: 'text' },
    formula: 'PMT(A1,12,100)',
    special: { formulaType: 'array', formulaReference: 'B1:C3' },
    consumer: 'SUM(C2)+A1',
  },
  { name: 'volatile-overlap', formula: 'PMT(A1,12,100)+RAND()' },
  { name: 'quoted-date-overlap', formula: 'PMT(A1,12,100)+DATE("1900",1,1)' },
  { name: 'invalid-valid-financial-restoration', initial: { value: '0', dataType: 'text' }, formula: 'PMT(A1,12,100)' },
  {
    name: 'invalid-valid-financial-date-preservation',
    calendar: true,
    initial: { value: '0', dataType: 'text' },
    formula: 'PMT(A1,12,100)+WEEKDAY(D1)',
  },
];
for (const test of cases) check(test.name, async () => {
  let session;
  try {
    const input = {
      revision: 'initial:' + test.name,
      sheets: [
        {
          name: 'Cases',
          date1904: !!test.calendar,
          rows: 8,
          columns: 8,
          cells: {
            A1: test.initial ?? { value: 0, dataType: 'number' },
            D1: { value: 100, dataType: 'number' },
            B1: { formula: test.formula, ...test.special },
            C2: { value: 5, dataType: 'number' },
            E1: { formula: test.consumer ?? 'B1*12' },
            F1: { formula: 'IFERROR(B1,99)' },
            G1: { formula: 'SUM(2,3,4)' },
          },
        },
      ],
    };
    const inputBefore = structuredClone(input);
    session = await createCalculationSession(input);
    const original = structuredClone(session.view);
    assert.deepEqual(input, inputBefore);
    let current = structuredClone(session.view);
    for (const [index, value] of [0.005, 0].entries()) {
      const next = structuredClone(current);
      next.revision = test.name + ':edit' + index;
      next.sheets[0].cells.A1 = { value, dataType: 'number' };
      const before = structuredClone(next);
      const patches = await session.update(next, [{ sheet: 'Cases', cell: 'A1' }]);
      assert.deepEqual(next, before, 'session must not mutate supplied next view');
      const actual = structuredClone(next);
      for (const p of patches) actual.sheets[0].cells[p.cell] = p.value;
      const cold = await calculateWorkbook(next);
      for (const ref of ['B1', 'E1', 'F1']) {
        const semantic = (c) => {
          const { revision, ...result } = c ?? {};
          return result;
        };
        assert.deepEqual(
          semantic(actual.sheets[0].cells[ref].calculation),
          semantic(cold.sheets[0].cells[ref].calculation),
          test.name + ':' + ref
        );
      }
      current = actual;
      assert.equal(cold.sheets[0].cells.G1.calculation.value, 9);
      assert.equal(original.sheets[0].cells.G1.calculation.value, 9);
      if (test.name === 'invalid-valid-financial-restoration')
        assert.equal(cold.sheets[0].cells.B1.calculation.status, 'calculated');
      else assert.equal(cold.sheets[0].cells.B1.calculation.status, 'unresolved');
      assert.deepEqual(session.view, original, 'retained published old view unchanged');
    }
  } finally {
    session?.dispose();
  }
});
