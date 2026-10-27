import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import { calculateWorkbook, createCalculationSession } from '../core/calculation-node.mjs';
import { readWorkbook } from '../core/workbook-package.mjs';

const view = (cells) => ({ revision: 'rounding', sheets: [{ name: 'Main', rows: 20, columns: 8, cells }] });
const number = (value) => ({ value, dataType: 'number' });
const calculated = (value, revision = 'rounding') => ({ status: 'calculated', value, revision });

test('ROUNDUP exactly replays retained Excel scalar results and feeds dependent formulas', async () => {
  const oracle = JSON.parse(await fs.readFile(new URL('../fixtures/excel-b6-runtime.json', import.meta.url), 'utf8'));
  const source = readWorkbook(await fs.readFile(new URL('../fixtures/excel-b6-runtime.xlsx', import.meta.url)));
  const cases = oracle.cases.filter((item) => item.formula.startsWith('ROUNDUP('));
  assert.deepEqual(
    cases.map((item) => item.ref),
    ['D13', 'D15', 'D17']
  );
  const input = view(Object.fromEntries(cases.map((item) => [item.ref, source.sheets[0].cells[item.ref]])));
  input.sheets[0].cells.F1 = { formula: 'D13*2' };
  input.sheets[0].cells.F2 = { formula: 'SUM(D13)' };
  const before = structuredClone(input),
    result = (await calculateWorkbook(input)).sheets[0].cells;
  for (const item of cases) assert.deepEqual(result[item.ref].calculation, calculated(item.expected), item.formula);
  assert.deepEqual(result.F1.calculation, calculated(-5.6000000000000005));
  assert.deepEqual(result.F2.calculation, calculated(-2.8000000000000003));
  assert.deepEqual(input, before, 'Authored formulas, stored caches and caller-owned input remain intact');
});

test('bounded numeric ROUNDUP handles both signs, zero and positive/negative places', async () => {
  // Only D13/D15/D17 above are retained exact Excel runtime oracles. These
  // surrounding cases check the declared scalar rescaling and documented rule.
  const cases = [
    [2.71, 1, 2.8000000000000003],
    [-2.71, 1, -2.8000000000000003],
    [1.005, 2, 1.01],
    [-1.005, 2, -1.01],
    [3.14159, 3, 3.142],
    [123.4, -1, 130],
    [-123.4, -1, -130],
    [31415.92654, -2, 31500],
    [3.2, 0, 4],
    [-3.2, 0, -4],
    [0, 1, -0],
    [0, -1, -0],
  ];
  const cells = {};
  cases.forEach(([value, digits], index) => {
    const row = index + 1;
    cells[`A${row}`] = number(value);
    cells[`B${row}`] = number(digits);
    cells[`C${row}`] = { formula: `ROUNDUP(A${row},B${row})`, cachedValue: 999 };
  });
  const output = (await calculateWorkbook(view(cells))).sheets[0].cells;
  cases.forEach(([, , expected], index) => assert.deepEqual(output[`C${index + 1}`].calculation, calculated(expected)));
});

test('ROUNDUP keeps upstream coercion, multi-cell arrays and unqualified scalar precision paths', async () => {
  const cases = [
    ['ROUNDUP("-2.71",1)', -2.8],
    ['ROUNDUP(-2.71,"1")', -2.8],
    // The large-number result is retained upstream behavior, not an Excel
    // oracle claim: its epsilon/unit selection is outside this correction.
    ['ROUNDUP(-2.71,1.9)', -2.8],
    ['ROUNDUP(1E20,1)', 99999999999999970000],
    ['INDEX(ROUNDUP({-2.71,-1.005},1),1,1)', -2.8],
    ['ROUNDUP(TRUE,1)', 1],
  ];
  const cells = Object.fromEntries(cases.map(([formula], index) => [`A${index + 1}`, { formula }]));
  cells.B1 = { formula: 'ROUNDUP(1/0,1)' };
  cells.B2 = { formula: 'IFERROR(ROUNDUP("invalid",1),99)' };
  cells.B3 = { formula: 'ROUNDUP(UNKNOWN_LOCAL(1),1)' };
  const result = (await calculateWorkbook(view(cells))).sheets[0].cells;
  cases.forEach(([, expected], index) => assert.deepEqual(result[`A${index + 1}`].calculation, calculated(expected)));
  assert.deepEqual(result.B1.calculation, { status: 'error', error: '#DIV/0!', revision: 'rounding' });
  assert.deepEqual(result.B2.calculation, calculated(99));
  assert.equal(result.B3.calculation.status, 'unresolved');
});

test('incremental ROUNDUP literal changes agree exactly with cold results and preserve the accepted branch', async () => {
  const input = view({ A1: number(-2.71), A2: number(1), B1: { formula: 'ROUNDUP(A1,A2)' }, C1: { formula: 'B1*2' } });
  const session = await createCalculationSession(input),
    accepted = structuredClone(session.view);
  try {
    for (const [value, digits, expected] of [
      [2.71, 1, 2.8000000000000003],
      [-1.005, 2, -1.01],
      [-123.4, -1, -130],
      [0, 0, -0],
    ]) {
      const next = structuredClone(accepted);
      next.revision = `updated:${value}:${digits}`;
      next.sheets[0].cells.A1 = number(value);
      next.sheets[0].cells.A2 = number(digits);
      const changes = [
        { sheet: 'Main', cell: 'A1' },
        { sheet: 'Main', cell: 'A2' },
      ];
      for (const patch of await session.update(next, changes)) next.sheets[0].cells[patch.cell] = patch.value;
      const cold = await calculateWorkbook(next);
      assert.deepEqual(next.sheets[0].cells.B1.calculation, calculated(expected, next.revision));
      for (const ref of ['B1', 'C1'])
        assert.deepEqual(next.sheets[0].cells[ref].calculation, cold.sheets[0].cells[ref].calculation);
    }
    assert.deepEqual(session.view, accepted);
    assert.deepEqual(accepted.sheets[0].cells.B1.calculation, calculated(-2.8000000000000003));
  } finally {
    session.dispose();
  }
});
