import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { readWorkbook } from '../core/workbook-package.mjs';
import { calculateWorkbook, createCalculationSession } from '../core/calculation-node.mjs';
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const view = (cells, date1904 = false) => ({
  revision: 'dates',
  sheets: [{ name: 'Cases', rows: 20, columns: 8, date1904, cells }],
});

for (const system of [1900, 1904])
  test(`portable numeric dates exactly match the actual Excel ${system} oracle and protect text diagnostics`, async () => {
    const source = await fs.readFile(new URL(`../fixtures/excel-date-${system}-source.xlsx`, import.meta.url));
    const reference = await fs.readFile(new URL(`../fixtures/excel-date-${system}-reference.xlsx`, import.meta.url));
    const oracle = JSON.parse(
      await fs.readFile(new URL(`../fixtures/excel-date-${system}-reference.json`, import.meta.url), 'utf8')
    );
    assert.equal(hash(source), oracle.sourceSha256);
    assert.equal(hash(reference), oracle.oracleWorkbookSha256);
    const input = readWorkbook(source),
      before = structuredClone(input),
      output = await calculateWorkbook(input),
      cells = output.sheets[0].cells;
    assert.equal(input.sheets[0].date1904, system === 1904);
    let qualified = 0,
      protectedCases = 0;
    for (const item of oracle.cases) {
      const result = cells[item.ref].calculation;
      if (Number(item.ref.slice(1)) >= 78) {
        assert.equal(result.status, 'unresolved', item.ref);
        protectedCases++;
        continue;
      }
      if (item.dataType === 'error') {
        assert.equal(result.status, 'error', item.ref);
        assert.equal(result.error, item.expected, item.ref);
        assert.equal(result.value, undefined);
      } else {
        assert.equal(result.status, 'calculated', item.ref);
        assert.equal(result.value, item.expected, item.ref);
      }
      assert.equal(result.revision, input.revision, item.ref);
      qualified++;
    }
    assert.equal(qualified, 248);
    assert.equal(protectedCases, 9);
    assert.deepEqual(input, before);
  });

test('unqualified date inputs stay unresolved through IFERROR, siblings, names and dependents', async () => {
  const input = view(
    {
      A1: { value: '2024', dataType: 'text' },
      A2: { value: 2024, dataType: 'number' },
      B1: { formula: 'DATE(A1,1,1)' },
      B2: { formula: 'IFERROR(B1,99)' },
      C1: { formula: 'IFERROR(DATE(A1,1,1),99)+DATE(A2,1,1)' },
      C2: { formula: 'C1+1' },
      D1: { formula: 'IFERROR(YEAR("2024-01-01"),99)' },
      D2: { formula: 'IFERROR(DATE(2024,1201,1),99)' },
      E1: { formula: 'INDEX(DATE({2024,2025},1,1),1,1)' },
      E2: { formula: 'YEAR(TRUE)' },
      F1: { formula: 'DATEVALUE("2024-01-01")' },
      F2: { formula: 'F1+1' },
      G1: { formula: 'DateName+1' },
      H1: { formula: 'DATE(2024,1,1)' },
    },
    true
  );
  input.namedRanges = [{ name: 'DateName', formula: 'DATE(Cases!$A$1,1,1)' }];
  const output = (await calculateWorkbook(input)).sheets[0].cells;
  for (const ref of ['B1', 'B2', 'C1', 'C2', 'D1', 'D2', 'E1', 'E2', 'F1', 'F2', 'G1'])
    assert.equal(output[ref].calculation.status, 'unresolved', ref);
  assert.equal(output.H1.calculation.value, 43830);
});

test('incremental date misses retire only affected proofs and match cold calculations', async () => {
  const input = view(
    {
      A1: { value: '2024', dataType: 'text' },
      A2: { value: 2024, dataType: 'number' },
      B1: { formula: 'IFERROR(DATE(A1,1,1),99)+DATE(A2,1,1)' },
      C1: { formula: 'B1+1' },
      D1: { formula: 'YEAR(A2)' },
      E1: { formula: 'DATE("unqualified",1,1)' },
      E2: { formula: 'IFERROR(E1,99)' },
    },
    true
  );
  const session = await createCalculationSession(input),
    accepted = structuredClone(session.view);
  try {
    for (const value of [2024, '2024', 2025, -1, 2024]) {
      const next = structuredClone(accepted);
      next.revision = `changed:${value}:${typeof value}`;
      next.sheets[0].cells.A1 = { value, dataType: typeof value === 'number' ? 'number' : 'text' };
      for (const patch of await session.update(next, [{ sheet: 'Cases', cell: 'A1' }]))
        next.sheets[0].cells[patch.cell] = patch.value;
      const cold = await calculateWorkbook(next);
      for (const ref of ['B1', 'C1'])
        assert.deepEqual(next.sheets[0].cells[ref].calculation, cold.sheets[0].cells[ref].calculation, ref);
      assert.equal(cold.sheets[0].cells.E1.calculation.status, 'unresolved');
      assert.equal(cold.sheets[0].cells.E2.calculation.status, 'unresolved');
      assert.equal(next.sheets[0].cells.B1.calculation.status, typeof value === 'number' ? 'calculated' : 'unresolved');
    }
    assert.deepEqual(session.view, accepted);
  } finally {
    session.dispose();
  }
});

test('simultaneous date systems remain isolated when another owner is disposed', async () => {
  const make = (date1904) =>
    view(
      {
        A1: { value: 0, dataType: 'number' },
        B1: { formula: 'YEAR(A1)' },
        C1: { formula: 'MONTH(A1)' },
        D1: { formula: 'DAY(A1)' },
        E1: { formula: 'DATE(1904,1,1)' },
      },
      date1904
    );
  const a = await createCalculationSession(make(false)),
    b = await createCalculationSession(make(true));
  try {
    assert.equal(a.view.sheets[0].cells.B1.calculation.value, 1900);
    assert.equal(a.view.sheets[0].cells.D1.calculation.value, 0);
    assert.equal(b.view.sheets[0].cells.B1.calculation.value, 1904);
    assert.equal(b.view.sheets[0].cells.D1.calculation.value, 1);
    a.dispose();
    const next = make(true);
    next.revision = 'survivor';
    next.sheets[0].cells.A1.value = 59.5;
    const patches = await b.update(next, [{ sheet: 'Cases', cell: 'A1' }]);
    assert.equal(patches.find((p) => p.cell === 'D1').value.calculation.value, 29);
  } finally {
    a.dispose();
    b.dispose();
  }
});

test('portable owners refuse changed or malformed calendar metadata before literal mutation', async () => {
  const input = view({
      A1: { value: 60.75, dataType: 'number' },
      B1: { formula: 'DAY(A1)' },
      C1: { formula: 'DATE(1904,1,1)' },
    }),
    session = await createCalculationSession(input);
  const accepted = structuredClone(session.view);
  try {
    for (const date1904 of [true, 'true', 1, null]) {
      const next = structuredClone(input);
      next.sheets[0].date1904 = date1904;
      next.sheets[0].cells.A1.value = 61.99;
      await assert.rejects(session.update(next, [{ sheet: 'Cases', cell: 'A1' }]), { code: 'formula-unsupported' });
      if (typeof date1904 !== 'boolean') {
        assert.equal((await calculateWorkbook(next)).sheets[0].cells.B1.calculation.status, 'unresolved');
        await assert.rejects(createCalculationSession(next), { code: /^(formula-unsupported|internal-invalid-field)$/ });
      }
      assert.deepEqual(session.view, accepted);
    }
    const next = structuredClone(input);
    next.revision = 'after-refusal';
    next.sheets[0].cells.A1.value = 61.99;
    const patches = await session.update(next, [{ sheet: 'Cases', cell: 'A1' }]);
    assert.equal(patches.find((p) => p.cell === 'B1').value.calculation.value, 1);
    assert.equal(session.view.sheets[0].cells.C1.calculation.value, 1462);
  } finally {
    session.dispose();
  }
});
