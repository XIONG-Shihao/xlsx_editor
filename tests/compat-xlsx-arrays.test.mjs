import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { readWorkbook } from '../core/workbook-package.mjs';
import {
  calculateWorkbook,
  createCalculationSession,
  calculationInputRanges,
  calculationDependencies,
} from '../core/calculation-node.mjs';
import { arrayFormulaContext, ARRAY_CALCULATION_REASON } from '../core/array-calculation.mjs';
const view = (cells) => ({ revision: 'array-protection', sheets: [{ name: 'Cases', rows: 20, columns: 10, cells }] });
const inputs = { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 2 } };

test('actual Excel dynamic workbooks retain opaque caches while every unqualified master stays unresolved', async () => {
  const oracle = JSON.parse(
    await fs.readFile(new URL('../fixtures/excel-dynamic-reference.json', import.meta.url), 'utf8')
  );
  assert.equal(oracle.excel, '16.113.3');
  for (const [name, digest] of [
    [oracle.source, oracle.sourceSha256],
    [oracle.oracleWorkbook, oracle.oracleWorkbookSha256],
  ]) {
    const bytes = await fs.readFile(new URL(`../fixtures/${name}`, import.meta.url));
    assert.equal(createHash('sha256').update(bytes).digest('hex'), digest);
    const input = readWorkbook(bytes),
      original = structuredClone(input),
      out = await calculateWorkbook(input);
    for (const ref of Object.keys(oracle.authoredFormulas)) {
      assert.equal(out.sheets[0].cells[ref].calculation.status, 'unresolved', `${name}: ${ref}`);
      assert.equal(
        out.sheets[0].cells[ref].cachedValue,
        input.sheets[0].cells[ref].cachedValue,
        `${name}: retained ${ref}`
      );
    }
    assert.deepEqual(input, original);
  }
  // Rich errors carry a separate legacy cache spelling. No oracle expectation
  // is normalized into a guessed scalar/function error by this protection.
  assert.equal(oracle.cells.A41.richError.errorType, 8);
  assert.equal(oracle.cells.A41.expected, '#SPILL!');
  assert.equal(oracle.cells.A41.rawValue, '#VALUE!');
});

test('authored modern array functions and spill syntax never publish a top-left result', async () => {
  for (const formula of [
    'SEQUENCE(3)',
    'sequence(1)',
    'FILTER(A1:A3,A1:A3>0)',
    'UNIQUE(A1:A3)',
    'SUM(SEQUENCE(3))',
    'IF(TRUE,9,SEQUENCE(3))',
    'IFERROR(SEQUENCE(3),99)',
    '_xlfn.SEQUENCE(3)',
    '_xlfn._xlws.FILTER(A1:A3,A1:A3>0)',
    'SUM(C1#)',
    '@C1',
  ]) {
    const input = view({
        ...inputs,
        C1: { formula, value: 777, cachedValue: 777 },
        E1: { formula: 'IFERROR(C1,99)' },
        E2: { formula: 'C2+1' },
        F1: { formula: 'SUM(A1:A3)' },
      }),
      original = structuredClone(input);
    const cells = (await calculateWorkbook(input)).sheets[0].cells;
    assert.equal(cells.C1.calculation.status, 'unresolved', formula);
    assert.equal(cells.C1.calculation.error, ARRAY_CALCULATION_REASON, formula);
    for (const ref of ['E1', 'E2']) assert.equal(cells[ref].calculation.status, 'unresolved', `${formula}: ${ref}`);
    assert.equal(cells.F1.calculation.value, 5);
    assert.equal(cells.C1.cachedValue, 777);
    assert.equal(cells.C2, undefined);
    assert.equal(cells.C3, undefined);
    assert.deepEqual(input, original);
  }
});

test('multi-cell result shape protects ordinary range/array roots and follower consumers', async () => {
  for (const formula of ['A1:A3', 'A1:A3+1', '{1;2;3}', 'INDEX(A1:A3,0)', 'ROW(A1:A3)', 'TRANSPOSE(A1:A3)']) {
    const input = view({
      ...inputs,
      C1: { formula },
      E5: { formula: 'IFERROR(C1,99)' },
      E6: { formula: 'C2+1' },
      E7: { formula: 'SUM(C1:D3)' },
      F8: { formula: 'SUM(A1:A3)' },
    });
    const cells = (await calculateWorkbook(input)).sheets[0].cells;
    assert.equal(cells.C1.calculation.status, 'unresolved', formula);
    assert.equal(cells.C1.calculation.error, ARRAY_CALCULATION_REASON, formula);
    assert.equal(cells.E5.calculation.status, 'unresolved', formula);
    assert.equal(cells.E7.calculation.status, 'unresolved', formula);
    if (formula !== 'TRANSPOSE(A1:A3)') assert.equal(cells.E6.calculation.status, 'unresolved', formula);
    assert.equal(cells.F8.calculation.value, 5);
    assert.equal(cells.C2, undefined);
    assert.equal(cells.D1, undefined);
  }
  const collision = view({
    ...inputs,
    C1: { formula: 'A1:A3' },
    C2: { value: 'preserved blocker' },
    E1: { formula: 'IFERROR(C1,99)' },
    F1: { formula: 'SUM(A1:A3)' },
  });
  const cells = (await calculateWorkbook(collision)).sheets[0].cells;
  assert.equal(cells.C1.calculation.status, 'unresolved');
  assert.equal(cells.E1.calculation.status, 'unresolved');
  assert.equal(cells.C2.value, 'preserved blocker');
  assert.equal(cells.F1.calculation.value, 5);
  const declared = view({
    ...inputs,
    C1: { formula: 'A1:A3', formulaType: 'array', formulaReference: 'C1:C3', value: 1, cachedValue: 1 },
    C2: { value: 777, cachedValue: 777 },
    C3: { value: 888, cachedValue: 888 },
    E1: { formula: 'C2+1' },
    F1: { formula: 'SUM(A1:A3)' },
  });
  const opaque = (await calculateWorkbook(declared)).sheets[0].cells;
  assert.equal(opaque.C1.calculation.status, 'unresolved');
  assert.equal(opaque.E1.calculation.status, 'unresolved', 'Declared followers cannot lend stale cached values');
  assert.equal(opaque.C2.value, 777);
  assert.equal(opaque.F1.calculation.value, 5);
});

test('names and lazy/error-catching branches cannot hide unqualified array contexts', async () => {
  const input = view({
    ...inputs,
    C1: { formula: 'ArrayName' },
    C5: { formula: 'SUM(ChainedName)' },
    C9: { formula: 'IFERROR(C5,99)' },
    C12: { formula: 'IF(TRUE,9,SEQUENCE(3))' },
    C15: { formula: 'IFERROR(UNIQUE(A1:A3),99)' },
    F1: { formula: 'SUM(A1:A3)' },
    F2: { formula: 'LEN("SEQUENCE(3) and A1#")' },
  });
  input.namedRanges = [
    { name: 'ArrayName', formula: 'SEQUENCE(3)' },
    { name: 'ChainedName', formula: 'ArrayName+1' },
  ];
  const cells = (await calculateWorkbook(input)).sheets[0].cells;
  for (const ref of ['C1', 'C5', 'C9', 'C12', 'C15']) assert.equal(cells[ref].calculation.status, 'unresolved', ref);
  assert.equal(cells.F1.calculation.value, 5);
  assert.equal(cells.F2.calculation.value, 19);
});

test('array dependency hints remain conservative while scalar aggregates and quoted selectors stay qualified', async () => {
  for (const formula of ['SEQUENCE(3)', 'IF(A1=1,A1:A3,A1)', 'INDEX(A1:A3,0)', '{1;2;3}', 'SUM(C1#)']) {
    const input = view({ ...inputs, C1: { formula } });
    assert.equal(calculationInputRanges(input).mode, 'full', formula);
    assert.equal(calculationDependencies(input), null, formula);
    const proof = calculationDependencies(input, { conservative: true });
    assert.deepEqual(
      proof[0].ranges,
      [{ sheet: 'Cases', startRow: 0, endRow: 1048575, startColumn: 0, endColumn: 16383 }],
      formula
    );
  }
  const input = view({
    ...inputs,
    C1: { formula: 'SUM({1;2;3})' },
    C2: { formula: 'SUM(A1:A3)' },
    C3: { formula: 'LEN("FILTER(A1:A3) # @")' },
    C4: { formula: 'INDEX(A1:A3,2)' },
  });
  const cells = (await calculateWorkbook(input)).sheets[0].cells;
  assert.equal(cells.C1.calculation.value, 6);
  assert.equal(cells.C2.calculation.value, 5);
  assert.equal(cells.C4.calculation.value, 2);
  for (const formula of [
    'SUM({1;2;3})',
    'SUM(A1:A3)',
    "SUM('SEQUENCE(3)'!A1:A3)",
    'SUM(Prices[[#Data],[SEQUENCE(3)]])',
    'LEN("A1# @ SEQUENCE(3)")',
    '#N/A',
  ])
    assert.equal(arrayFormulaContext(formula).blocked, false, formula);
});

test('incremental scalar-to-array-to-scalar transitions refresh follower consumers and match cold views', async () => {
  let accepted = view({
    A1: { value: 0 },
    B1: { value: 2 },
    B2: { value: 3 },
    B3: { value: 4 },
    C1: { formula: 'IF(A1=0,99,B1:B3)' },
    D1: { formula: 'C2+1' },
    E1: { formula: 'SUM(C1:C3)' },
    F1: { formula: 'SUM(B1:B3)' },
  });
  const session = await createCalculationSession(accepted);
  accepted = structuredClone(session.view);
  try {
    for (const value of [1, 0, 1, 0]) {
      const next = structuredClone(accepted);
      next.revision = `array-state:${value}`;
      next.sheets[0].cells.A1.value = value;
      for (const patch of await session.update(next, [{ sheet: 'Cases', cell: 'A1' }]))
        next.sheets[0].cells[patch.cell] = patch.value;
      const cold = await calculateWorkbook(next);
      for (const ref of ['C1', 'D1', 'E1'])
        assert.deepEqual(
          next.sheets[0].cells[ref].calculation,
          cold.sheets[0].cells[ref].calculation,
          `${value}: ${ref}`
        );
      assert.equal(next.sheets[0].cells.C1.calculation.status, value ? 'unresolved' : 'calculated');
      assert.equal(next.sheets[0].cells.C2, undefined);
      accepted = next;
    }
  } finally {
    session.dispose();
  }
});

test('disposing an array owner preserves another scalar owner with identical authored formulas', async () => {
  const make = (value) =>
    view({
      A1: { value },
      B1: { value: 2 },
      B2: { value: 3 },
      C1: { formula: 'IF(A1=1,B1:B2,99)' },
      D1: { formula: 'C2+1' },
    });
  const array = await createCalculationSession(make(1)),
    scalar = await createCalculationSession(make(0));
  try {
    assert.equal(array.view.sheets[0].cells.C1.calculation.status, 'unresolved');
    assert.equal(scalar.view.sheets[0].cells.C1.calculation.value, 99);
    array.dispose();
    const next = make(0);
    next.revision = 'scalar-survivor';
    next.sheets[0].cells.B1.value = 7;
    const patches = await scalar.update(next, [{ sheet: 'Cases', cell: 'B1' }]);
    assert.equal(patches.find((p) => p.cell === 'C1').value.calculation.value, 99);
  } finally {
    array.dispose();
    scalar.dispose();
  }
});
