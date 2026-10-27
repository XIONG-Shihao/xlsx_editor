/** Native and portable financial functions agree with exact mathematics, refuse unsafe inputs and
 * keep incremental results equal to cold calculation. */
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import { calculateWorkbook, createCalculationSession } from '../core/calculation-node.mjs';
import { calculateFinancialScalar, financialCalls } from '../core/financial-calculation.mjs';
import { nativeFormula } from '../core/native-formula.mjs';
import { createWorkbookNativeClient } from '../core/native-client.mjs';
const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-financial-semantics-'));
const base = fileURLToPath(new URL('../', import.meta.url)),
  planPath = path.join(base, 'tests/support/financial-engine-cases.json'),
  plan = JSON.parse(await fs.readFile(planPath, 'utf8')),
  helper = path.join(base, 'native/bin/darwin-arm64/workbook-engine'),
  client = createWorkbookNativeClient(helper),
  native = (request) => client.call(request);
const bits = (value) => {
  const b = Buffer.alloc(8);
  b.writeDoubleBE(value);
  return b.toString('hex');
};
const view = (cells, calendar = 1900) => ({
  revision: 'financial-engine:' + calendar,
  sheets: [
    { name: 'Cases', path: 'xl/worksheets/sheet1.xml', date1904: calendar === 1904, rows: 60, columns: 10, cells },
  ],
});
try {
  const mathematics = JSON.parse(
    execFileSync('python3', [path.join(base, 'tests/support/financial-engine-binary-input-math.py'), planPath], {
      encoding: 'utf8',
    })
  );
  for (const calendar of [1900, 1904]) {
    const input = view(Object.fromEntries(plan.cases.map((c) => [c.cell, { formula: c.formula }])), calendar),
      before = structuredClone(input),
      root = path.join(directory, 'native-' + calendar);
    await fs.mkdir(root);
    const analyze = await nativeFormula(native, root, 'analyze', { view: input }),
      calculated = await nativeFormula(native, root, 'evaluate', { view: input }),
      portable = await calculateWorkbook(input);
    assert.equal(analyze.supported, true);
    assert.equal(calculated.supported, true, calculated.reason);
    assert.deepEqual(input, before);
    const values = plan.cases.map((c) => {
      const actual = calculated.results.find((r) => r.ref === c.cell)?.value,
        compat = portable.sheets[0].cells[c.cell].calculation,
        math = mathematics.cases.find((m) => m.id === c.id);
      assert.equal(compat.status, 'calculated', c.id);
      assert.equal(typeof actual, 'number');
      assert.equal(bits(actual), bits(compat.value), c.id + ' native/portable candidate agreement');
      if (math.exactRationalVector)
        assert.equal(bits(actual), math.nearestBinary64Bits, c.id + ' exact rational mathematics');
      else {
        const residual = Math.abs(actual - math.nearestBinary64Value),
          bound = Math.max(1e-10, Math.abs(math.nearestBinary64Value) * 2 ** -42);
        assert(residual <= bound, c.id + ' independently bounded mathematical residual');
      }
      return {
        id: c.id,
        formula: c.formula,
        native: actual,
        portable: compat.value,
        bits: bits(actual),
        mathematicalNearestBits: math.nearestBinary64Bits,
        exactRationalMathematics: math.exactRationalVector,
        actualExcelReference: null,
      };
    });
  }
  const unsupported = [
    'PMT("0",12,100)',
    'PMT(TRUE,12,100)',
    'PMT(0,2.5,100)',
    'PMT(0,0,100)',
    'PMT(0,1201,100)',
    'PMT(-0,12,100)',
    'PMT(-0.1,12,100)',
    'PMT(1.1,12,100)',
    'PMT(0,12,100,0,2)',
    'PMT(0,12,1000000000001)',
    'PMT(1,1200,100)',
    'PMT(0,12,100,,0)',
    'PMT(0,12,A1:A2)',
    'PMT(0,12,Other!A1)',
    'PMT(0,12,Cash)',
    'PMT(0.1234567890123456789012345,12,100)',
    'PMT(0,12,900719925474.0993)',
    'PMT(0,12,9007199254740993)',
    'PMT(0,12,{1,2})',
    '_xlfn.PMT(0,12,100)',
  ];
  for (const [index, formula] of unsupported.entries()) {
    const root = path.join(directory, 'miss-' + index);
    await fs.mkdir(root);
    const input = view({
      A1: { value: 0, dataType: 'number' },
      A2: { value: 12, dataType: 'number' },
      B1: { formula },
      B2: { formula: 'IFERROR(B1,99)' },
      B3: { formula: 'IF(TRUE,7,' + formula + ')' },
      B4: { formula: 'IFERROR(' + formula + ',99)' },
      C1: { formula: 'B2+1' },
      D1: { formula: 'SUM(2,3,4)' },
    });
    input.namedRanges = [{ name: 'Cash', formula: 'Cases!$A$1' }];
    const result = await nativeFormula(native, root, 'evaluate', { view: input });
    assert.equal(result.supported, false, formula);
    assert.equal(result.results, undefined, 'No partial projection cache');
    const out = await calculateWorkbook(input);
    for (const ref of ['B1', 'B2', 'B3', 'B4', 'C1'])
      assert.equal(out.sheets[0].cells[ref].calculation.status, 'unresolved', formula + ':' + ref);
    assert.equal(out.sheets[0].cells.D1.calculation.value, 9, formula + ' independent sibling');
  }
  const references = [
    { value: '0', dataType: 'text' },
    { value: null, dataType: 'blank' },
    { value: true, dataType: 'boolean' },
    { value: '#N/A', dataType: 'text' },
    { value: '#N/A', dataType: 'error' },
    { value: 0.12345678901234568, dataType: 'number', rawValue: '0.1234567890123456789012345' },
    { value: 0, dataType: 'number', rawValue: '-0' },
    { value: 0, dataType: 'number', spill: { schema: 1 } },
    { value: 0, dataType: 'number', scenarioResult: true },
    { formula: '1-1', calculation: { status: 'stored', value: 999 } },
    { formula: 'NO_SUCH_FUNCTION()', calculation: { status: 'stored', value: 999 } },
    { formula: 'A1' },
    { formula: '0', formulaType: 'array', formulaReference: 'A1:A1' },
  ];
  for (const [index, source] of references.entries()) {
    const root = path.join(directory, 'ref-' + index);
    await fs.mkdir(root);
    const input = view({
      A1: source,
      A2: { value: 12, dataType: 'number' },
      A3: { value: 1200, dataType: 'number' },
      B1: { formula: 'PMT(A1,A2,A3)' },
      B2: { formula: 'IFERROR(B1,99)' },
      B3: { formula: 'IF(TRUE,7,PMT(A1,A2,A3))' },
      C1: { formula: 'SUM(2,3)' },
    });
    const result = await nativeFormula(native, root, 'evaluate', { view: input });
    assert.equal(result.supported, false, JSON.stringify(source));
    const out = await calculateWorkbook(input);
    for (const ref of ['B1', 'B2', 'B3'])
      assert.equal(
        out.sheets[0].cells[ref].calculation.status,
        index === 11 && ref === 'B2' ? 'circular' : 'unresolved',
        index + ':' + ref
      );
    assert.equal(out.sheets[0].cells.C1.calculation.value, 5);
  }
  const defaults = [];
  for (const name of ['PMT', 'PV', 'FV'])
    for (const args of [
      [0, 12, 100],
      [0.5, 2, -100],
    ]) {
      const base = calculateFinancialScalar(name, args).value;
      assert.equal(bits(base), bits(calculateFinancialScalar(name, [...args, 0]).value));
      assert.equal(bits(base), bits(calculateFinancialScalar(name, [...args, 0, 0]).value));
      assert.equal(calculateFinancialScalar(name, [args[0], args[1], -args[2]]).value, -base);
      defaults.push({ name, args, value: base });
    }
  assert.equal(bits(calculateFinancialScalar('FV', [0, 12, -100, 1200]).value), '0000000000000000');
  const input = view({
    A1: { value: 0, dataType: 'number' },
    A2: { value: 12, dataType: 'number' },
    A3: { value: 1200, dataType: 'number' },
    B1: { formula: 'PMT(A1,A2,A3)' },
    B2: { formula: 'IFERROR(B1,99)' },
    B3: { formula: 'IF(TRUE,7,PMT(A1,A2,A3))' },
    C1: { formula: 'B1*12' },
    D1: { formula: 'IFERROR(Finance,99)' },
  });
  input.namedRanges = [{ name: 'Finance', formula: 'PMT(0,12,100)' }];
  const session = await createCalculationSession(input),
    initial = structuredClone(session.view);
  try {
    for (const [index, source] of [
      { value: 0.5, dataType: 'number' },
      { value: '0', dataType: 'text' },
      { value: 0, dataType: 'number', rawValue: '-0' },
      { value: 0, dataType: 'number' },
    ].entries()) {
      const next = structuredClone(input);
      next.revision = 'financial-incremental:' + index;
      next.sheets[0].cells.A1 = source;
      const patches = await session.update(next, [{ sheet: 'Cases', cell: 'A1' }]);
      for (const patch of patches) next.sheets[0].cells[patch.cell] = patch.value;
      const cold = await calculateWorkbook(next);
      for (const ref of ['B1', 'B2', 'B3', 'C1'])
        assert.deepEqual(
          next.sheets[0].cells[ref].calculation,
          cold.sheets[0].cells[ref].calculation,
          index + ':' + ref
        );
      assert.equal(cold.sheets[0].cells.D1.calculation.status, 'unresolved');
    }
    assert.deepEqual(session.view, initial);
  } finally {
    session.dispose();
  }
} finally {
  client.close();
  await fs.rm(directory, { recursive: true, force: true });
}
