import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import { calculateWorkbook, calculationDependencies, calculationInputRanges } from '../core/calculation-node.mjs';

const view = (cells) => ({ revision: 'union-guard', sheets: [{ name: 'Cases', rows: 20, columns: 8, cells }] });

test('actual Excel union references remain explicit native requirements in the portable evaluator', async () => {
  const oracle = JSON.parse(
    await fs.readFile(new URL('../fixtures/excel-union-reference.json', import.meta.url), 'utf8')
  );
  assert.equal(oracle.excel, '16.113.3');
  const input = view({
    A1: { value: 1e16 },
    A2: { value: 1 },
    A3: { value: -1e16 },
    B1: { value: 4 },
    B2: { value: 5 },
    B3: { value: 6 },
    C1: { value: '#DIV/0!', dataType: 'error' },
    C2: { value: '#N/A', dataType: 'error' },
    ...Object.fromEntries(
      oracle.cases.map((item) => [
        item.ref,
        { formula: item.formula, value: item.expected, cachedValue: item.expected },
      ])
    ),
    F1: { formula: 'IFERROR(E1,99)' },
    G1: { formula: 'SUM(B1,B2)' },
    H1: { formula: '(B1+B2)' },
  });
  input.sheets.push({ name: 'Other Sheet', cells: { A1: { value: 7 } } });
  const original = structuredClone(input),
    out = await calculateWorkbook(input);
  for (const item of oracle.cases) {
    const cell = out.sheets[0].cells[item.ref];
    assert.equal(cell.calculation.status, 'unresolved', item.ref);
    assert.match(cell.calculation.error, /unions require qualified native/);
    assert.equal(cell.cachedValue, item.expected);
  }
  assert.equal(
    out.sheets[0].cells.F1.calculation.status,
    'unresolved',
    'IFERROR cannot hide an unsupported union predecessor'
  );
  for (const ref of ['G1', 'H1']) assert.equal(out.sheets[0].cells[ref].calculation.value, 9);
  assert.deepEqual(input, original);
});

test('truncated lexer prefixes cannot establish static union dependencies', () => {
  const input = view({ A1: { value: 1 }, A2: { value: 2 }, B1: { formula: 'SUM((A1,A2))' } });
  assert.equal(calculationInputRanges(input).mode, 'full');
  assert.equal(calculationDependencies(input), null);
  const conservative = calculationDependencies(input, { conservative: true });
  assert.equal(conservative.length, 1);
  assert.deepEqual(conservative[0].ranges, [
    { sheet: 'Cases', startRow: 0, endRow: 1048575, startColumn: 0, endColumn: 16383 },
  ]);
});

test('function arguments, array constants, quoted commas and table selectors remain distinct from unions', async () => {
  const input = view({
    A1: { value: 2 },
    A2: { value: 3 },
    B1: { formula: 'SUM(A1,A2)' },
    B2: { formula: 'SUM({1,2})' },
    B3: { formula: 'LEN("a,b")' },
    B4: { formula: 'ROUNDUP(1.005,2)' },
    B5: { formula: 'SUM((A1+A2))' },
    B6: { formula: 'LEN("escaped ""comma, quote"" ")' },
  });
  const out = await calculateWorkbook(input),
    cells = out.sheets[0].cells;
  for (const [ref, value] of Object.entries({ B1: 5, B2: 3, B3: 3, B4: 1.01, B5: 5 }))
    assert.deepEqual(cells[ref].calculation, { status: 'calculated', value, revision: 'union-guard' }, ref);
  assert.equal(cells.B6.calculation.status, 'calculated');
  const table = view({
    A1: { value: 'Price' },
    A2: { value: 2 },
    A3: { value: 3 },
    B1: { formula: 'SUM(Prices[[#Data],[Price]])' },
  });
  table.tables = [{ name: 'Prices', sheet: 'Cases', ref: 'A1:A3', columns: [{ name: 'Price' }] }];
  const calculated = (await calculateWorkbook(table)).sheets[0].cells.B1.calculation;
  assert.equal(calculated.status, 'calculated');
  assert.equal(calculated.value, 5);
});

test('named and lazy union contexts remain conservative without blocking unrelated formulas', async () => {
  const input = view({
    A1: { value: 1 },
    A2: { value: 2 },
    B1: { formula: 'SUM(UnionArea)' },
    B2: { formula: 'IF(TRUE,9,SUM((A1,A2)))' },
    B3: { formula: 'A1+A2' },
    B4: { formula: 'IFERROR(B1,99)' },
  });
  input.namedRanges = [{ name: 'UnionArea', formula: '(Cases!A1,Cases!A2)' }];
  const out = await calculateWorkbook(input);
  for (const ref of ['B1', 'B2', 'B4']) assert.equal(out.sheets[0].cells[ref].calculation.status, 'unresolved', ref);
  assert.equal(out.sheets[0].cells.B3.calculation.value, 3);
  const graph = calculationDependencies(input, { conservative: true });
  for (const ref of ['B1', 'B2'])
    assert(graph.find((item) => item.ref === ref).ranges.some((range) => range.endRow === 1048575));
});
