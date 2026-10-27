import test from 'node:test';
import assert from 'node:assert/strict';
import { calculateWorkbook, shiftFormula, transformWorkbook } from '../core/calculation-node.mjs';
const view = (cells) => ({ revision: 'source-revision', sheets: [{ name: 'Main', cells }] });
const result = async (cells) => (await calculateWorkbook(view(cells))).sheets[0].cells;
test('IF dependencies calculate15, stale cache stays separate, inputs and results do not share formula state', async () => {
  const input = view({
    A1: { formula: 'IF(1=1,10,0)', value: 999, cachedValue: 999, rawValue: '999' },
    A2: { value: 5 },
    B1: { formula: 'SUM(A1:A2)' },
  });
  const before = structuredClone(input),
    output = await calculateWorkbook(input);
  assert.deepEqual(input, before);
  assert.equal(output.sheets[0].cells.A1.value, 999);
  assert.equal(output.sheets[0].cells.A1.cachedValue, 999);
  assert.deepEqual(output.sheets[0].cells.B1.calculation, {
    status: 'calculated',
    value: 15,
    revision: 'source-revision',
  });
  output.sheets[0].cells.B1.calculation.value = 123;
  assert.equal(input.sheets[0].cells.B1.calculation, undefined);
});
test('unknown functions, ordinary errors, stored errors and circular dependencies never become partial numbers', async () => {
  const cells = await result({
    A1: { formula: 'UNKNOWN_LOCAL(1)' },
    A2: { value: 5 },
    A3: { formula: 'SUM(A1:A2)' },
    B1: { formula: '1/0' },
    B2: { formula: 'B1+1' },
    C1: { formula: 'C2+1' },
    C2: { formula: 'C1+1' },
    C3: { formula: 'SUM(C1:C2)' },
    D1: { dataType: 'error', value: '#N/A' },
    D2: { formula: 'SUM(D1)' },
    E1: { formula: 'SUM(E1:E2)' },
    E2: { value: 5 },
  });
  for (const ref of ['A1', 'A3']) assert.equal(cells[ref].calculation.status, 'unresolved');
  for (const ref of ['B1', 'B2', 'D2']) assert.equal(cells[ref].calculation.status, 'error');
  for (const ref of ['C1', 'C2', 'C3', 'E1']) assert.equal(cells[ref].calculation.status, 'circular');
  for (const ref of ['A3', 'B2', 'C3', 'D2', 'E1']) assert.equal(cells[ref].calculation.value, undefined);
});
test('cross-sheet, mixed references, workbook names and table references recalculate after shared commands', async () => {
  const input = {
    revision: 'before',
    sheets: [
      {
        name: 'Main',
        cells: {
          A1: { formula: "'Prices'!B2*$A$2" },
          A2: { value: 2 },
          B1: { formula: 'SUM(PriceList[Total])' },
          C1: { formula: 'TaxRate*2' },
        },
      },
      {
        name: 'Prices',
        cells: {
          A1: { value: 'Quantity' },
          B1: { value: 'Price' },
          C1: { value: 'Total' },
          A2: { value: 2 },
          B2: { value: 125 },
          C2: { formula: 'A2*B2' },
          A3: { value: 4 },
          B3: { value: 20 },
          C3: { formula: 'A3*B3' },
        },
      },
    ],
    namedRanges: [{ name: 'TaxRate', formula: 'Main!$A$2' }],
    tables: [
      {
        name: 'PriceList',
        sheet: 'Prices',
        ref: 'A1:C3',
        columns: [{ name: 'Quantity' }, { name: 'Price' }, { name: 'Total' }],
      },
    ],
  };
  let output = await calculateWorkbook(input);
  assert.equal(output.sheets[0].cells.A1.calculation.value, 250);
  assert.equal(output.sheets[0].cells.B1.calculation.value, 330);
  assert.equal(output.sheets[0].cells.C1.calculation.value, 4);
  input.sheets[1].cells.A2.value = 3;
  input.sheets[1].cells.B2.value = 150;
  input.revision = 'after';
  output = await calculateWorkbook(input);
  assert.equal(output.sheets[0].cells.B1.calculation.value, 530);
  assert.equal(output.sheets[1].cells.C2.calculation.value, 450);
  assert.equal(output.sheets[1].cells.C2.calculation.revision, 'after');
});
test('stored precision and literal types survive; unsafe dependencies are explicitly unresolved', async () => {
  const cells = await result({
    A1: { value: '00123', dataType: 'text' },
    A2: { value: '=SUM(A1)', dataType: 'text' },
    A3: { value: true, dataType: 'boolean' },
    A4: { value: '東京 / 中文', dataType: 'text' },
    B1: { value: 9007199254740992, rawValue: '9007199254740993', dataType: 'number' },
    B2: { formula: 'B1+1' },
    B3: { formula: 'SUM(B2)' },
    C1: { value: '#DIV/0!', dataType: 'text' },
    C2: { formula: 'SUM(C1)' },
    D1: { formula: '1=1' },
    E1: { formula: '"#DIV/0!"' },
    E2: { formula: 'SUM(E1)' },
    F1: { formula: 'IF(TRUE,"#NAME?",0)' },
  });
  assert.equal(cells.A1.value, '00123');
  assert.equal(cells.A2.formula, undefined);
  assert.equal(cells.A3.value, true);
  assert.equal(cells.A4.value, '東京 / 中文');
  assert.equal(cells.B1.rawValue, '9007199254740993');
  for (const ref of ['B2', 'B3', 'C2', 'E2', 'F1']) assert.equal(cells[ref].calculation.status, 'unresolved');
  assert.equal(cells.D1.calculation.value, true);
  assert.deepEqual(cells.E1.calculation, { status: 'calculated', value: '#DIV/0!', revision: 'source-revision' });
});
test('parser shifts mixed references and cross-sheet references without editing quoted text', () => {
  assert.equal(shiftFormula('$B2*C$2+Data!A1+"A1"', 1, 2), '$B4*D$2+Data!B3+"A1"');
});
test('reference-aware insert updates all sheets, names and table bounds without mutating source', async () => {
  const input = {
    revision: 'old',
    sheets: [
      {
        name: 'Data',
        rows: 20,
        columns: 8,
        cells: {
          A1: { value: 'Qty' },
          B1: { value: 'Price' },
          C1: { value: 'Total' },
          A2: { value: 2 },
          B2: { value: 125 },
          C2: { formula: '$A2*B$2' },
          A3: { value: 4 },
          B3: { value: 20 },
          C3: { formula: 'A3*B3', bold: true },
          C4: { formula: 'SUM(C2:C3)' },
        },
      },
      { name: 'Other', cells: { A1: { formula: 'Data!C4' } } },
    ],
    namedRanges: [{ name: 'Costs', formula: 'Data!$C$2:$C$3' }],
    tables: [
      { sheet: 'Data', name: 'Budget', ref: 'A1:C3', columns: [{ name: 'Qty' }, { name: 'Price' }, { name: 'Total' }] },
    ],
  };
  const before = structuredClone(input),
    output = await transformWorkbook(input, [{ sheet: 'Data', operation: 'insertRow', index: 2 }]);
  assert.deepEqual(input, before);
  assert.equal(output.revision, 'old');
  assert.equal(output.sheets[0].cells.C4.formula, 'A4*B4');
  assert.equal(output.sheets[0].cells.C4.bold, true);
  assert.equal(output.sheets[0].cells.C5.formula, 'SUM(C2:C4)');
  assert.equal(output.sheets[1].cells.A1.formula, 'Data!C5');
  assert.equal(output.namedRanges[0].formula, 'Data!$C$2:$C$4');
  assert.equal(output.tables[0].ref, 'A1:C4');
});
test('unqualified table header/totals semantics fail explicitly instead of yielding partial results', async () => {
  const input = view({ A1: { formula: 'SUM(Budget[Price])' } });
  input.tables = [{ name: 'Budget', sheet: 'Main', ref: 'B1:B3', totalsRowCount: 1 }];
  assert.equal((await calculateWorkbook(input)).sheets[0].cells.A1.calculation.status, 'unresolved');
});
test('shared/array/data-table formulas and followers remain unresolved without borrowing stale caches', async () => {
  const input = view({
    A1: {
      formula: 'B1*2',
      formulaType: 'shared',
      sharedFormulaId: '0',
      formulaReference: 'A1:A2',
      value: 250,
      cachedValue: 250,
    },
    A2: { formulaType: 'shared', sharedFormulaId: '0', value: 300, cachedValue: 300 },
    B1: { value: 125 },
    C1: { formula: 'SUM(A1:A2)' },
    D1: { formula: 'B1+1' },
    E1: { formula: 'B1:B2*2', formulaType: 'array', value: 250 },
    F1: { formulaType: 'dataTable', value: 42 },
  });
  const before = structuredClone(input),
    output = await calculateWorkbook(input),
    cells = output.sheets[0].cells;
  assert.deepEqual(input, before);
  assert.equal(cells.A2.formula, undefined);
  assert.equal(cells.A2.cachedValue, 300);
  assert.equal(cells.A1.formula, 'B1*2');
  for (const ref of ['A1', 'A2', 'C1', 'E1', 'F1']) assert.equal(cells[ref].calculation.status, 'unresolved');
  assert.equal(cells.D1.calculation.value, 126);
  await assert.rejects(
    () => transformWorkbook(input, [{ sheet: 'Main', operation: 'insertRow', index: 1 }]),
    /not qualified/
  );
});
test('qualified 1904 dates and serial arithmetic work while unsafe decimal precision remains protected', async () => {
  const input = view({
    A1: { value: 0 },
    A2: { formula: 'A1+1' },
    B1: { formula: 'DATE(1904,1,1)' },
    B2: { formula: 'B1+1' },
    C1: { value: 0.12345678901234568, rawValue: '0.123456789012345678901', dataType: 'number' },
    C2: { formula: 'C1+1' },
    D1: { value: 0.1, rawValue: '0.100000000000000000', dataType: 'number' },
    D2: { formula: 'D1+1' },
  });
  input.sheets[0].date1904 = true;
  const cells = (await calculateWorkbook(input)).sheets[0].cells;
  assert.equal(cells.A2.calculation.value, 1);
  assert.equal(cells.B1.calculation.value, 0);
  assert.equal(cells.B2.calculation.value, 1);
  assert.equal(cells.C2.calculation.status, 'unresolved');
  assert.equal(cells.C1.rawValue, '0.123456789012345678901');
  assert.equal(cells.D2.calculation.value, 1.1);
});
test('insert immediately above a total expands aggregate/name/table boundaries and fills formulas without copying literals', async () => {
  const input = {
    revision: 'old',
    sheets: [
      {
        name: 'Budget',
        rows: 20,
        columns: 8,
        cells: {
          A1: { value: 'Quantity' },
          B1: { value: 'Price' },
          C1: { value: 'Total' },
          A2: { value: 2, styleId: 3 },
          B2: { value: 10, format: 2, bold: true },
          C2: { formula: 'A2*B2', styleId: 4 },
          C3: { formula: 'SUM(C2:C2)' },
        },
      },
    ],
    namedRanges: [{ name: 'Costs', formula: 'Budget!$C$2:$C$2' }],
    tables: [
      {
        name: 'Items',
        sheet: 'Budget',
        ref: 'A1:C2',
        columns: [{ name: 'Quantity' }, { name: 'Price' }, { name: 'Total' }],
      },
    ],
  };
  const out = await transformWorkbook(input, [{ sheet: 'Budget', operation: 'insertRow', index: 2 }]);
  const cells = out.sheets[0].cells;
  assert.equal(cells.A3.value, null);
  assert.equal(cells.A3.styleId, 3);
  assert.equal(cells.B3.value, null);
  assert.equal(cells.B3.format, 2);
  assert.equal(cells.B3.bold, true);
  assert.equal(cells.C3.formula, 'A3*B3');
  assert.equal(cells.C3.cachedValue, null);
  assert.equal(cells.C4.formula, 'SUM(C2:C3)');
  assert.equal(out.tables[0].ref, 'A1:C3');
  assert.equal(out.namedRanges[0].formula, 'Budget!$C$2:$C$3');
  cells.A3.value = 3;
  cells.B3.value = 7;
  assert.equal((await calculateWorkbook(out)).sheets[0].cells.C4.calculation.value, 41);
  input.sheets[0].cells.C3.formula = 'SUM(C2:C2)*1.1';
  await assert.rejects(() => transformWorkbook(input, [{ sheet: 'Budget', operation: 'insertRow', index: 2 }]), {
    code: 'formula-unsupported',
  });
});
test('volatile formulas and dependents are unresolved and reproducible in the1900 system', async () => {
  const input = view({
    A1: { formula: 'NOW()' },
    A2: { formula: 'TODAY()' },
    A3: { formula: 'RAND()' },
    A4: { formula: 'RANDBETWEEN(1,10)' },
    B1: { formula: 'SUM(A1:A4)' },
    C1: { formula: 'IF(TRUE,15,0)' },
    D1: { formula: '"RAND()"' },
  });
  const before = structuredClone(input),
    a = await calculateWorkbook(input),
    b = await calculateWorkbook(input);
  assert.deepEqual(a, b);
  assert.deepEqual(input, before);
  for (const ref of ['A1', 'A2', 'A3', 'A4', 'B1'])
    assert.equal(a.sheets[0].cells[ref].calculation.status, 'unresolved');
  assert.equal(a.sheets[0].cells.C1.calculation.value, 15);
  assert.equal(a.sheets[0].cells.D1.calculation.value, 'RAND()');
});
test('persistent literal batches match full calculation for errors, text, names and cycles without mutating accepted results', async () => {
  const { createCalculationSession } = await import('../core/calculation-node.mjs');
  const input = view({
    A1: { value: 2, dataType: 'number' },
    B1: { formula: '10/A1' },
    C1: { formula: 'B1+1' },
    D1: { formula: 'SUM(A1)' },
    E1: { formula: 'E2+1' },
    E2: { formula: 'E1+1' },
    F1: { formula: 'RAND()+A1' },
    G1: { formula: 'NamedInput+1' },
    H1: { formula: 'IF(A1=0,"#N/A",15)' },
    I1: { formula: 'SUM(H1)' },
  });
  input.namedRanges = [{ name: 'NamedInput', formula: 'Main!$A$1' }];
  const original = structuredClone(input),
    session = await createCalculationSession(input),
    accepted = structuredClone(session.view);
  try {
    for (const value of [0, '#DIV/0!', 4]) {
      const next = structuredClone(accepted);
      next.revision = `value:${value}`;
      next.sheets[0].cells.A1 = {
        value,
        dataType: typeof value === 'number' ? 'number' : 'text',
        rawValue: String(value),
      };
      const patches = await session.update(next, [{ sheet: 'Main', cell: 'A1' }]);
      for (const patch of patches) next.sheets[0].cells[patch.cell] = patch.value;
      const complete = await calculateWorkbook(next);
      for (const ref of ['B1', 'C1', 'D1', 'G1', 'H1', 'I1'])
        assert.deepEqual(next.sheets[0].cells[ref].calculation, complete.sheets[0].cells[ref].calculation, ref);
      assert.equal(complete.sheets[0].cells.E1.calculation.status, 'circular');
      assert.equal(complete.sheets[0].cells.F1.calculation.status, 'unresolved');
    }
    assert.deepEqual(input, original);
    assert.deepEqual(session.view, accepted);
    const changed = structuredClone(accepted);
    changed.sheets[0].cells.A1 = { formula: 'B1+1' };
    await assert.rejects(() => session.update(changed, [{ sheet: 'Main', cell: 'A1' }]), /literal/);
  } finally {
    session.dispose();
  }
  await assert.rejects(() => session.update(input, []), { code: 'formula-refused' });
});
test('error-catching functions cannot mask unsupported functions or names as fresh values', async () => {
  const cells = await result({
    A1: { formula: 'UNKNOWN_LOCAL(1)' },
    B1: { formula: 'IFERROR(A1,99)' },
    C1: { formula: 'ISERROR(A1)' },
    D1: { formula: 'IFERROR(UNKNOWN_LOCAL(1),99)' },
    E1: { formula: 'IFERROR(1/0,99)' },
    F1: { formula: 'IFERROR(UnknownName,99)' },
  });
  for (const ref of ['A1', 'B1', 'C1', 'D1', 'F1']) assert.equal(cells[ref].calculation.status, 'unresolved', ref);
  assert.equal(cells.E1.calculation.status, 'calculated');
  assert.equal(cells.E1.calculation.value, 99);
});
test('genuine local broken references calculate typed errors and catchable results without changing authored formulas or caches', async () => {
  const formulas = [
    '#REF!',
    'Sheet1!#REF!',
    "'O''Brien Inputs'!#REF!",
    'SUM(#REF!:A2)',
    'SUM(A2:#REF!)',
    'SUM(#REF!:#REF!)',
    'SUM(Sheet1!#REF!)',
    'SUM(Sheet1!A2:#REF!)',
  ];
  const cells = { A2: { value: 7, dataType: 'number' } };
  formulas.forEach((formula, i) => (cells['B' + (i + 1)] = { formula, value: 913, cachedValue: 913, rawValue: '913' }));
  Object.assign(cells, {
    C1: { formula: 'IFERROR(Sheet1!#REF!,99)' },
    C2: { formula: 'IFERROR(SUM(#REF!:A2),99)' },
    C3: { formula: 'IFERROR(B3,99)' },
    C4: { formula: 'ISERROR(Sheet1!#REF!)' },
  });
  const input = {
      revision: 'broken-reference',
      sheets: [
        { name: 'Main', cells },
        { name: 'Sheet1', cells: { A2: { value: 8 } } },
        { name: "O'Brien Inputs", cells: { A2: { value: 9 } } },
      ],
    },
    before = structuredClone(input),
    out = (await calculateWorkbook(input)).sheets[0].cells;
  assert.deepEqual(input, before);
  formulas.forEach((formula, i) => {
    const cell = out['B' + (i + 1)];
    assert.equal(cell.formula, formula);
    assert.equal(cell.value, 913);
    assert.equal(cell.cachedValue, 913);
    assert.equal(cell.rawValue, '913');
    assert.deepEqual(cell.calculation, { status: 'error', error: '#REF!', revision: 'broken-reference' });
  });
  for (const ref of ['C1', 'C2', 'C3'])
    assert.deepEqual(out[ref].calculation, { status: 'calculated', value: 99, revision: 'broken-reference' });
  assert.equal(out.C4.calculation.value, true);
});
test('broken-reference admission cannot launder unknown calls/names, external or three-dimensional qualifiers, token junk or error-looking text', async () => {
  const formulas = [
    'IFERROR(UnknownName+#REF!,99)',
    'IFERROR(UNKNOWN_LOCAL(1)+Sheet1!#REF!,99)',
    'IFERROR(#REF!(),99)',
    'IFERROR(Sheet1!#REF!(),99)',
    'IFERROR(#REF!junk,99)',
    "IFERROR('[Book.xlsx]Sheet1'!#REF!,99)",
    "IFERROR('Sheet1:Main'!#REF!,99)",
    'IFERROR(Sheet1:Main!#REF!,99)',
  ];
  const cells = {};
  formulas.forEach((formula, i) => (cells['A' + (i + 1)] = { formula, value: 913, cachedValue: 913 }));
  Object.assign(cells, {
    B1: { formula: '"Sheet1!#REF!"' },
    B2: { formula: '"#REF!:A2"' },
    B3: { formula: '"#REF!"' },
    B4: { formula: 'IF(TRUE,"#REF!",0)' },
    B5: { formula: 'IFERROR(Sheet1!#REF!,"#REF!")' },
    C1: { value: '#REF!', dataType: 'text' },
    C2: { formula: 'IFERROR(C1,99)' },
  });
  const out = await result(cells);
  formulas.forEach((_, i) => {
    assert.equal(out['A' + (i + 1)].calculation.status, 'unresolved');
    assert.equal(out['A' + (i + 1)].calculation.value, undefined);
  });
  for (const [ref, value] of [
    ['B1', 'Sheet1!#REF!'],
    ['B2', '#REF!:A2'],
    ['B3', '#REF!'],
  ])
    assert.deepEqual(out[ref].calculation, { status: 'calculated', value, revision: 'source-revision' });
  for (const ref of ['B4', 'B5', 'C2']) {
    assert.equal(out[ref].calculation.status, 'unresolved', ref);
    assert.equal(out[ref].calculation.value, undefined, ref);
  }
  assert.equal(out.C1.value, '#REF!');
  assert.equal(out.C1.dataType, 'text');
});
test('persistent compatibility updates keep broken-reference errors and their IFERROR dependencies consistent with fresh calculation', async () => {
  const { createCalculationSession } = await import('../core/calculation-node.mjs');
  const input = view({
      A1: { value: 2, dataType: 'number' },
      B1: { formula: 'Main!#REF!', value: 913, cachedValue: 913 },
      C1: { formula: 'IFERROR(B1,A1)' },
      D1: { formula: 'SUM(#REF!:A2)' },
      E1: { formula: 'IFERROR(D1,A1*2)' },
    }),
    before = structuredClone(input),
    session = await createCalculationSession(input),
    accepted = structuredClone(session.view);
  try {
    assert.equal(accepted.sheets[0].cells.B1.calculation.status, 'error');
    assert.equal(accepted.sheets[0].cells.C1.calculation.value, 2);
    const next = structuredClone(accepted);
    next.revision = 'broken-reference-update';
    next.sheets[0].cells.A1 = { value: 7, dataType: 'number' };
    const patches = await session.update(next, [{ sheet: 'Main', cell: 'A1' }]);
    for (const patch of patches) next.sheets[0].cells[patch.cell] = patch.value;
    const fresh = await calculateWorkbook(next);
    for (const ref of ['C1', 'E1'])
      assert.deepEqual(next.sheets[0].cells[ref].calculation, fresh.sheets[0].cells[ref].calculation, ref);
    assert.equal(fresh.sheets[0].cells.C1.calculation.value, 7);
    assert.equal(fresh.sheets[0].cells.E1.calculation.value, 14);
    assert.deepEqual(input, before);
    assert.deepEqual(session.view, accepted);
  } finally {
    session.dispose();
  }
});
