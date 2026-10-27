import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import { createRequire } from 'node:module';
import { calculateWorkbook, createCalculationSession } from '../core/calculation-node.mjs';
const view = (cells) => ({ revision: 'typed-errors', sheets: [{ name: 'Main', rows: 20, columns: 8, cells }] });
const error = (code) => ({ status: 'error', error: code, revision: 'typed-errors' });

test('a real REPT ErrorValueObject matches the retained Excel oracle and propagates as an ordinary error', async () => {
  const oracle = JSON.parse(
    await fs.readFile(new URL('../fixtures/excel-native-text22.json', import.meta.url), 'utf8')
  );
  const retained = oracle.cases.find((item) => item.ref === 'B29');
  assert.equal(retained.formula, 'REPT("a",-0.2)');
  assert.equal(retained.expected, '#VALUE!');
  const input = view({
    A1: { formula: retained.formula, value: 999, cachedValue: 999 },
    B1: { formula: 'A1+1' },
    C1: { formula: 'IFERROR(A1,"caught")' },
    D1: { formula: 'ISERROR(A1)' },
  });
  const before = structuredClone(input),
    cells = (await calculateWorkbook(input)).sheets[0].cells;
  assert.deepEqual(cells.A1.calculation, error(retained.expected));
  assert.deepEqual(cells.B1.calculation, error('#VALUE!'));
  assert.deepEqual(cells.C1.calculation, { status: 'calculated', value: 'caught', revision: 'typed-errors' });
  assert.deepEqual(cells.D1.calculation, { status: 'calculated', value: true, revision: 'typed-errors' });
  assert.deepEqual(input, before);
});

test('typed final errors do not certify same-looking formula text, caught nested errors or unsafe literal inputs', async () => {
  const cells = (
    await calculateWorkbook(
      view({
        A1: { formula: '"#VALUE!"' },
        A2: { formula: 'SUM(A1)' },
        B1: { formula: 'IF(TRUE,"#VALUE!",REPT("a",-0.2))' },
        B2: { formula: 'IFERROR(B1,99)' },
        C1: { formula: 'IFERROR(REPT("a",-0.2),"#VALUE!")' },
        C2: { formula: 'ISERROR(C1)' },
        D1: { value: '#VALUE!', dataType: 'text' },
        D2: { formula: 'REPT(D1,1)' },
        E1: { formula: 'IFERROR(UNKNOWN_LOCAL(1),"#VALUE!")' },
        E2: { formula: 'ISERROR(E1)' },
        F1: { formula: 'PLATFORM_TYPED_RESULT(REPT("a",-0.2))' },
      })
    )
  ).sheets[0].cells;
  assert.deepEqual(cells.A1.calculation, { status: 'calculated', value: '#VALUE!', revision: 'typed-errors' });
  for (const ref of ['A2', 'B1', 'B2', 'C1', 'C2', 'D2', 'E1', 'E2', 'F1'])
    assert.equal(cells[ref].calculation.status, 'unresolved', ref);
});

test('quoted arguments can yield proven ordinary errors while circular, unsupported and array semantics retain guards', async () => {
  const cells = (
    await calculateWorkbook(
      view({
        A1: { formula: 'REPT("a",-1)' },
        A2: { formula: 'SUBSTITUTE("a","a","b",0)' },
        B1: { formula: 'IFERROR(REPT("a",-0.2),"bad")' },
        C1: { formula: 'IFERROR(UNKNOWN_LOCAL("a"),99)' },
        D1: { formula: 'D1+LEN("x")' },
        E1: { formula: 'REPT("a",-0.2)', formulaType: 'array', formulaReference: 'E1:E2' },
        E2: { formulaType: 'array' },
      })
    )
  ).sheets[0].cells;
  for (const ref of ['A1', 'A2']) assert.deepEqual(cells[ref].calculation, error('#VALUE!'));
  assert.deepEqual(cells.B1.calculation, { status: 'calculated', value: 'bad', revision: 'typed-errors' });
  assert.equal(cells.C1.calculation.status, 'unresolved');
  assert.equal(cells.D1.calculation.status, 'circular');
  for (const ref of ['E1', 'E2']) assert.equal(cells[ref].calculation.status, 'unresolved');
});

test('persistent error-to-text-to-error transitions retire final-error proofs and match cold closure', async () => {
  const input = view({
    A1: { value: -1, dataType: 'number' },
    B1: { formula: 'IF(A1<0,REPT("a",A1),"#VALUE!")' },
    C1: { formula: 'IFERROR(B1,99)' },
    D1: { formula: 'B1+1' },
  });
  const session = await createCalculationSession(input),
    accepted = structuredClone(session.view);
  try {
    assert.equal(accepted.sheets[0].cells.B1.calculation.status, 'error');
    for (const value of [1, -0.2, 2, -1]) {
      const next = structuredClone(accepted);
      next.revision = `changed:${value}`;
      next.sheets[0].cells.A1 = { value, dataType: 'number' };
      for (const patch of await session.update(next, [{ sheet: 'Main', cell: 'A1' }]))
        next.sheets[0].cells[patch.cell] = patch.value;
      const cold = await calculateWorkbook(next);
      for (const ref of ['B1', 'C1', 'D1'])
        assert.deepEqual(next.sheets[0].cells[ref].calculation, cold.sheets[0].cells[ref].calculation, ref);
      assert.equal(next.sheets[0].cells.B1.calculation.status, value < 0 ? 'error' : 'unresolved');
    }
    assert.deepEqual(session.view, accepted);
  } finally {
    session.dispose();
  }
});

test('repeated cold disposal and simultaneous owners preserve separate error proofs and active updates', async () => {
  const input = (value) =>
    view({
      A1: { value, dataType: 'number' },
      B1: { formula: 'IF(A1<0,REPT("a",A1),"#VALUE!")' },
      C1: { formula: 'IFERROR(B1,99)' },
    });
  const a = await createCalculationSession(input(-1)),
    b = await createCalculationSession(input(1));
  const acceptedA = structuredClone(a.view),
    acceptedB = structuredClone(b.view);
  async function update(session, value, label) {
    const next = input(value);
    next.revision = label;
    for (const patch of await session.update(next, [{ sheet: 'Main', cell: 'A1' }]))
      next.sheets[0].cells[patch.cell] = patch.value;
    assert.equal(next.sheets[0].cells.B1.calculation.status, value < 0 ? 'error' : 'unresolved');
    assert.equal(next.sheets[0].cells.C1.calculation.status, value < 0 ? 'calculated' : 'unresolved');
    if (value < 0) assert.equal(next.sheets[0].cells.C1.calculation.value, 99);
    return next;
  }
  try {
    for (let cycle = 0; cycle < 12; cycle++) {
      const cold = await calculateWorkbook(input(cycle % 2 ? -0.2 : 1));
      assert.equal(cold.sheets[0].cells.B1.calculation.status, cycle % 2 ? 'error' : 'unresolved');
      await update(a, cycle % 2 ? 1 : -0.2, `a:${cycle}`);
      await update(b, cycle % 2 ? -1 : 2, `b:${cycle}`);
    }
    a.dispose();
    await assert.rejects(a.update(input(-1), []), { code: 'formula-refused' });
    await update(b, -0.2, 'b:after-other-disposal');
    await update(b, 1, 'b:text-after-other-disposal');
    assert.deepEqual(a.view, acceptedA);
    assert.deepEqual(b.view, acceptedB);
  } finally {
    a.dispose();
    b.dispose();
  }
});

test('separately loaded calculation copies cannot share cached observer closures', async () => {
  const require = createRequire(import.meta.url);
  const libs = Object.assign(
    {},
    ...['core', 'engine-formula', 'sheets', 'sheets-formula', 'sheets-table'].map((name) =>
      require('@univerjs/' + name)
    ),
    require('@univerjs/core/facade')
  );
  for (const name of ['engine-formula', 'sheets', 'sheets-formula', 'sheets-table'])
    require('@univerjs/' + name + '/facade');
  const copyA = await import('../core/calculation.mjs?owner-copy=a'),
    copyB = await import('../core/calculation.mjs?owner-copy=b');
  const input = (value) =>
    view({ A1: { value, dataType: 'number' }, B1: { formula: 'IF(A1<0,REPT("a",A1),"#VALUE!")' } });
  const a = await copyA.createCalculationSessionWith(libs, input(-1));
  let b;
  try {
    b = await copyB.createCalculationSessionWith(libs, input(-0.2));
    assert.equal(a.view.sheets[0].cells.B1.calculation.status, 'error');
    assert.equal(b.view.sheets[0].cells.B1.calculation.status, 'error');
    const next = input(1);
    next.revision = 'second-copy:text';
    const patches = await b.update(next, [{ sheet: 'Main', cell: 'A1' }]);
    assert.equal(patches.find((patch) => patch.cell === 'B1').value.calculation.status, 'unresolved');
    b.dispose();
    const remaining = input(-0.2);
    remaining.revision = 'first-copy:after-disposal';
    const live = await a.update(remaining, [{ sheet: 'Main', cell: 'A1' }]);
    assert.equal(live.find((patch) => patch.cell === 'B1').value.calculation.status, 'error');
  } finally {
    a.dispose();
    b?.dispose();
  }
});
