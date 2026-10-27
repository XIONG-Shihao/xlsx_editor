import assert from 'node:assert/strict';
import { test, before, after } from 'node:test';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {
  nativeWorkbookOperationAsync as native,
  closeNativeWorkbookWorker as closeNativeDocumentWorker,
} from '@office-editors/xlsx-editor/native';
import { calculateWorkbook, transformWorkbook } from '@office-editors/xlsx-editor/core/calculation-node.mjs';
import { nativeFormula } from '@office-editors/xlsx-editor/core/native-formula.mjs';
let root;
before(async () => {
  root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-native-formula-'));
});
after(async () => {
  closeNativeDocumentWorker();
  if (root) await fs.rm(root, { recursive: true, force: true });
});
const view = (cells, extra = []) => ({
  revision: 'test',
  sheets: [{ name: 'Data', path: 'xl/worksheets/sheet1.xml', rows: 40, columns: 20, cells }, ...extra],
});
const number = (value) => ({ value, dataType: 'number' }),
  formula = (f) => ({ value: null, dataType: 'number', formula: f });
test('native arithmetic and aggregate results agree with the compatibility engine and independent values', async () => {
  const cells = { A1: number(3), A2: number(5), A4: number(-2) };
  const forms = [
    'A1*2',
    'SUM(A1:A4)',
    'AVERAGE(A1:A4)',
    'COUNT(A1:A4)',
    'MIN(A1:A4)',
    'MAX(A1:A4)',
    'SUM(A1,A2,A3,5)',
    '-(A1+2)/4',
    'A1+25%',
    '1.25e2+A1',
    'A3+1',
    'SUM(A3:A3)',
    'COUNT(A3:A3)',
    "'Other Sheet'!$A$1+A1",
  ];
  forms.forEach((f, i) => (cells[`B${i + 1}`] = formula(f)));
  for (let i = 1; i <= 120; i++) cells[`C${i}`] = formula(`(${i}/4+A1)*2-SUM(A1:A4)`);
  const input = view(cells, [{ name: 'Other Sheet', rows: 5, columns: 5, cells: { A1: number(10) } }]);
  input.sheets[0].rows = 130;
  const result = await nativeFormula(native, root, 'evaluate', { view: input });
  assert.equal(result.supported, true, result.reason);
  const expected = [6, 6, 2, 3, -2, 5, 13, -1.25, 3.25, 128, 1, 0, 0, 13];
  const compatible = await calculateWorkbook(structuredClone(input));
  for (const item of result.results) {
    const old = compatible.sheets.find((s) => s.name === item.sheet).cells[item.ref].calculation;
    assert.equal(old.status, 'calculated');
    assert.equal(item.value, old.value, `${item.ref}: ${item.value} vs ${old.value}`);
    if (item.ref.startsWith('B')) assert.equal(item.value, expected[Number(item.ref.slice(1)) - 1]);
  }
  const parsed = await nativeFormula(native, root, 'analyze', { view: input });
  assert.equal(parsed.supported, true);
  assert.equal(parsed.items.length, 134);
  assert(parsed.scope.ranges.some((r) => r.sheet === 'Other Sheet' && r.ref === 'A1'));
});
test('unsupported semantics return an atomic miss, with no partial results', async () => {
  for (const f of [
    'SUM("4",A1)',
    'SUM(1:1)',
    'A1^0.5',
    'SUM(Items[Value])',
    'OFFSET(A1,-1,0)',
    'INDIRECT("R1C1",FALSE)',
    'RAND()',
    'MissingName+1',
    'B1+1',
    'SUM()',
    "'Data:Other'!A1",
    "'[other.xlsx]Data'!A1",
    '-0',
    '0/-1',
    '-A3',
  ]) {
    const result = await nativeFormula(native, root, 'evaluate', {
      view: view({ A1: number(2), B1: formula(f), C1: formula('A1+1') }),
    });
    assert.equal(result.supported, false, f);
    assert.equal(result.results, undefined, f);
  }
  for (const value of [Infinity, NaN])
    assert.equal(
      (await nativeFormula(native, root, 'evaluate', { view: view({ A1: { value }, B1: formula('A1+1') }) })).supported,
      false
    );
  for (const cell of [{ value: 'unproved' }, { value: 'unproved', dataType: 'text' }]) {
    const textOperand = await nativeFormula(native, root, 'evaluate', {
      view: view({ A1: cell, B1: formula('A1+1') }),
    });
    assert.equal(textOperand.supported, true);
    assert.equal(textOperand.results[0].error, '#VALUE!');
    assert.equal(Object.hasOwn(textOperand.results[0], 'value'), false);
  }
  const unprovedError = await nativeFormula(native, root, 'evaluate', {
    view: view({ A1: { value: '#DIV/0!', dataType: 'error' }, B1: formula('A1+1') }),
  });
  assert.equal(unprovedError.supported, false);
  assert.equal(
    (await nativeFormula(native, root, 'evaluate', { view: view({ A1: formula('SUM(A1:A1048576)') }) })).supported,
    false
  );
  assert.equal(
    (await nativeFormula(native, root, 'analyze', { view: view({ A1: formula('1+'.repeat(1000) + '1') }) })).supported,
    false
  );
  assert.equal(
    (
      await nativeFormula(native, root, 'evaluate', {
        view: view({ A1: { value: 'x'.repeat(2 * 1024 * 1024) }, B1: formula('1+2') }),
      })
    ).supported,
    false
  );
});
// Inserting rows (names, totals, references that cannot be adjusted) is the engine's
// own work now; tests/structural-references.mjs checks it through the real session.
test('native formula work bounds and cancellation do not return partial results', async () => {
  const bounded = await nativeFormula(native, root, 'evaluate', {
    view: view({ A1: number(2), B1: formula('SUM(A1:A1048576)') }),
  });
  assert.equal(bounded.supported, false);
  assert.match(bounded.reason, /work bound/);
  assert.equal(bounded.results, undefined);
  await fs.writeFile(path.join(root, 'cancelled'), '');
  try {
    await assert.rejects(nativeFormula(native, root, 'evaluate', { view: view({ A1: formula('1+2') }) }), { code: 'operation-cancelled' });
  } finally {
    await fs.rm(path.join(root, 'cancelled'));
  }
});

test('native subtraction retains the verified binary64 eight-ULP boundary', async () => {
  const input = view({ A1: number(1_000_000_000 + 2 ** -20), B1: formula('A1-1000000000') });
  const result = await nativeFormula(native, root, 'evaluate', { view: input });
  assert.equal(result.supported, true);
  assert.equal(result.results[0].value, 2 ** -20);
  const compatible = await calculateWorkbook(input);
  assert.equal(compatible.sheets[0].cells.B1.calculation.value, 2 ** -20);
  assert.equal(compatible.sheets[0].cells.B1.calculation.value, 2 ** -20);
});
test('native decimal operations and common functions exactly match the pinned evaluator', async () => {
  const cells = {};
  let seed = 0x5a17;
  const random = () => {
    seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
    return seed;
  };
  for (let row = 1; row <= 240; row++) {
    const a = ((random() % 2000001) - 1000000) / 10 ** (random() % 7),
      b = ((random() % 999999) + 1) / 10 ** (random() % 7);
    cells[`A${row}`] = number(a);
    cells[`B${row}`] = number(b);
    for (const [col, op] of [
      ['C', '+'],
      ['D', '-'],
      ['E', '*'],
      ['F', '/'],
    ])
      cells[`${col}${row}`] = formula(`A${row}${op}B${row}`);
  }
  const cases = [
    '0.1+0.2',
    '1/3',
    'ABS(-1.25)',
    'ROUND(1.005,2)',
    'ROUND(-2.675,2)',
    'IF(A1>0,ROUND(A1,2),ABS(A1))',
    'IF(TRUE,12,1/0)',
    'IF(FALSE,1/0,13)',
    'IFERROR(1/0,14)',
    'SUM(10000000000000000,A1:A3,-10000000000000000)',
    'SUM(0.1,A1:A10)',
    'AVERAGE(A1:A10,B1:B10)',
    '(1000000000.000001-1000000000)',
    '1000000000000000100/3',
    '0.0000000000000001/3',
  ];
  cases.forEach((f, i) => (cells[`G${i + 1}`] = formula(f)));
  const input = view(cells);
  input.sheets[0].rows = 240;
  input.sheets[0].columns = 7;
  const result = await nativeFormula(native, root, 'evaluate', { view: input });
  assert.equal(result.supported, true, result.reason);
  const compatible = await calculateWorkbook(structuredClone(input));
  for (const item of result.results) {
    const expected = compatible.sheets[0].cells[item.ref].calculation;
    assert.equal(expected.status, 'calculated', item.ref);
    assert.equal(item.value, expected.value, `${item.ref}: ${cells[item.ref].formula}`);
  }
  assert.equal(result.results.length, 975);
});
test('referenced text, booleans and blanks have aggregate semantics, not scalar coercion', async () => {
  const input = view({
    A1: { value: 'header', dataType: 'text' },
    A2: number(0.1),
    A3: { value: true, dataType: 'boolean' },
    A5: number(0.2),
    B1: formula('SUM(A1:A6)'),
    B2: formula('AVERAGE(A1:A6)'),
    B3: formula('COUNT(A1:A6)'),
    B4: formula('MIN(A1:A6)'),
    B5: formula('MAX(A1:A6)'),
    B6: formula('SUM(A1)'),
  });
  const result = await nativeFormula(native, root, 'evaluate', { view: input });
  assert.equal(result.supported, true, result.reason);
  const expected = await calculateWorkbook(input);
  for (const item of result.results)
    assert.equal(item.value, expected.sheets[0].cells[item.ref].calculation.value, item.ref);
});
test('native static names, table columns and logical predicates match the pinned evaluator', async () => {
  const input = view({
    A1: { value: 'Value', dataType: 'text' },
    B1: { value: 'Other', dataType: 'text' },
    A2: number(0.1),
    A3: number(0.2),
    A4: number(0.3),
    B2: formula('IF(AND(A2>0,NOT(FALSE)),Items[[#This Row],[Value]]*Tax,0)'),
    B3: formula('Items[[#This Row],[Value]]*Tax'),
    B4: formula('Items[[#This Row],[Value]]*Tax'),
    D1: formula('SUM(Items[Value])'),
    D2: formula('SUM(Amounts)'),
    D3: formula('SUM(Nested)'),
    D4: formula('INT(-1.25)+SIGN(-4)'),
    D5: formula('IF(OR(FALSE,TRUE),7,1/0)'),
  });
  input.tables = [
    {
      name: 'Items',
      sheet: 'Data',
      ref: 'A1:B4',
      columns: [{ name: 'Value' }, { name: 'Other' }],
      headerRowCount: 1,
      totalsRowCount: 0,
    },
  ];
  input.namedRanges = [
    { name: 'Tax', formula: '2' },
    { name: 'Amounts', formula: 'Data!$A$2:$A$4' },
    { name: 'Nested', formula: 'Amounts' },
  ];
  const result = await nativeFormula(native, root, 'evaluate', { view: input });
  assert.equal(result.supported, true, result.reason);
  const expected = await calculateWorkbook(structuredClone(input));
  for (const item of result.results) {
    assert.equal(expected.sheets[0].cells[item.ref].calculation.status, 'calculated', item.ref);
    assert.equal(item.value, expected.sheets[0].cells[item.ref].calculation.value, item.ref);
  }
  for (const shorthand of ['Items[@Value]', '[@Value]']) {
    const unsafe = structuredClone(input);
    unsafe.sheets[0].cells.B2 = formula(shorthand);
    assert.equal(
      (await nativeFormula(native, root, 'evaluate', { view: unsafe })).supported,
      false,
      'Do not qualify shorthand before its OOXML serialization is supported'
    );
  }
  const parsed = await nativeFormula(native, root, 'analyze', { view: input });
  assert(parsed.supported, parsed.reason);
  assert(parsed.items.find((i) => i.ref === 'D2').ranges.some((r) => r.startRow === 1 && r.endRow === 3));
  for (const namedRanges of [
    [{ name: 'Tax', formula: 'Tax' }],
    [{ name: 'Tax', formula: 'Data!A2' }],
    [{ name: 'Tax', formula: '$A$2' }],
  ])
    assert.equal((await nativeFormula(native, root, 'evaluate', { view: { ...input, namedRanges } })).supported, false);
});
test('expanded native numeric subset matches pinned evaluator and independent expectations', async () => {
  const cases = [
    ['PRODUCT(A1:A5)', 6],
    ['PRODUCT(A3:A3)', 0],
    ['PRODUCT(A1:A5,2)', 12],
    ['MOD(7,3)', 1],
    ['MOD(-7,3)', 2],
    ['MOD(7,-3)', -2],
    ['SQRT(9)', 3],
    ['POWER(2,10)', 1024],
    ['A1^3', 8],
    ['TRUNC(-2.7)', -2],
    ['IFERROR(SQRT(-1),9)', 9],
    ['IFERROR(MOD(1,0),10)', 10],
    ['2^3^2', 64],
    ['-2^2', 4],
  ];
  const cells = {
    A1: number(2),
    A2: number(3),
    A4: { value: 'ignore', dataType: 'text' },
    A5: { value: true, dataType: 'boolean' },
  };
  cases.forEach(([f], i) => (cells[`B${i + 1}`] = formula(f)));
  const input = view(cells),
    compatible = await calculateWorkbook(structuredClone(input));
  const result = await nativeFormula(native, root, 'evaluate', { view: input });
  assert(result.supported, result.reason);
  for (const [i, [f, expected]] of cases.entries()) {
    const found = result.results.find((r) => r.ref === `B${i + 1}`);
    assert.equal(found.value, expected, f);
    assert.equal(found.value, compatible.sheets[0].cells[`B${i + 1}`].calculation.value, f);
  }
});
test('reference rewriting handles static unknown calls, strings, error literals and both copy axes', async () => {
  const { shiftFormulas } = await import('@office-editors/xlsx-editor/core/calculation-node.mjs');
  const forms = [
    'IF(A1>0,"A1 ""quoted""",B$2)&"B2"',
    'VLOOKUP(A1,$B$2:C9,2,FALSE)',
    'SIN(A1)+A1^2',
    'IFERROR(#REF!,A1)',
  ];
  const result = await nativeFormula(native, root, 'rewrite', {
    items: forms.map((formula) => ({ formula, copyRows: 2, copyColumns: 1 })),
  });
  assert(result.supported, result.reason);
  assert.equal(result.formulas[0], 'IF(B3>0,"A1 ""quoted""",C$2)&"B2"'); // Preserve escaped literals; pinned shifter corrupts this spelling.
  assert.deepEqual(
    result.formulas.slice(1),
    shiftFormulas(forms.slice(1).map((formula) => ({ formula, rowOffset: 2, columnOffset: 1 })))
  );
  for (const formula of ['INDIRECT("A1")', 'OFFSET(A1,1,0)', 'UNREGISTERED(A1)'])
    assert.equal((await nativeFormula(native, root, 'analyze', { view: view({ B1: { formula } }) })).supported, false);
});
test('explicit local names resolve against their named sheet and native dependency ranges', async () => {
  const input = view({ A1: number(2), B1: formula("'Other Sheet'!Rate+A1") }, [
    { name: 'Other Sheet', path: 'xl/worksheets/sheet2.xml', rows: 4, columns: 4, cells: { A1: number(7) } },
  ]);
  input.namedRanges = [
    { name: 'Rate', formula: '99' },
    { name: 'Rate', sheet: 'Other Sheet', formula: '$A$1' },
  ];
  const result = await nativeFormula(native, root, 'evaluate', { view: input });
  assert(result.supported, result.reason);
  assert.equal(result.results[0].value, 9);
  const plan = await nativeFormula(native, root, 'analyze', { view: input });
  assert(plan.supported, plan.reason);
  assert(plan.items[0].ranges.some((r) => r.sheet === 'Other Sheet' && r.ref === 'A1'));
  const expected = await calculateWorkbook(structuredClone(input));
  assert.equal(expected.sheets[0].cells.B1.calculation.status, 'unresolved'); // Qualified local names extend a previously unresolved input, checked independently and against Excel.
});

test('expanded common functions match the pinned engine on signed decimal and range cases', async () => {
  const cells = {};
  for (let r = 1; r <= 80; r++) {
    cells[`A${r}`] = number((r - 40) / 10);
    cells[`B${r}`] = number((((r % 9) + 1) / 10) * (r % 2 ? -1 : 1));
    cells[`C${r}`] = formula(`MOD(A${r},B${r})`);
    cells[`D${r}`] = formula(`SQRT(ABS(A${r}))`);
    cells[`E${r}`] = formula(`PRODUCT(A${r}:B${r},0.25)`);
    cells[`F${r}`] = formula(`TRUNC(A${r})`);
  }
  const input = view(cells);
  input.sheets[0].rows = 80;
  const expected = await calculateWorkbook(structuredClone(input));
  // Individual calls distinguish a declared semantic miss from a wrong native answer.
  let accepted = 0;
  for (const column of ['C', 'D', 'E', 'F'])
    for (let r = 1; r <= 80; r++) {
      const subset = view({
        A1: cells[`A${r}`],
        B1: cells[`B${r}`],
        C1: formula(cells[`${column}${r}`].formula.replace(/([AB])\d+/g, '$11')),
      });
      const nativeResult = await nativeFormula(native, root, 'evaluate', { view: subset });
      if (nativeResult.supported) {
        accepted++;
        assert.equal(
          nativeResult.results[0].value,
          expected.sheets[0].cells[`${column}${r}`].calculation.value,
          `${column}${r}: ${cells[`${column}${r}`].formula}`
        );
      } else assert.match(nativeResult.reason, /Negative zero/);
    }
  assert(accepted >= 300, `Expected broad native coverage, got ${accepted}/320`);
});

test('native and partial evaluation refuse rounded stored numeric inputs', async () => {
  const input = view({
    A1: { value: 12345678901234567000, rawValue: '12345678901234567890', dataType: 'number' },
    B1: formula('A1+1'),
  });
  const result = await nativeFormula(native, root, 'evaluate', { view: input });
  assert.equal(result.supported, false);
  assert.match(result.reason, /precision/);
  assert.equal(result.results, undefined);
});

test('duplicate defined-name bindings cannot silently choose the first value', async () => {
  const input = view({ B1: formula('Tax+1') });
  input.namedRanges = [
    { name: 'Tax', formula: '2' },
    { name: 'Tax', formula: '3' },
  ];
  const result = await nativeFormula(native, root, 'evaluate', { view: input });
  assert.equal(result.supported, false);
  assert.match(result.reason, /Ambiguous/);
});

test('approved Excel exceptions reach native, fallback and dependent formulas', async () => {
  const input = view({
    A1: number(-3),
    B1: formula('TRUNC(A1)'),
    B2: formula('DATE(1900,2,29)'),
    C1: formula('B1+10'),
    C2: formula('B2+1'),
  });
  const expected = { B1: -3, B2: 60, C1: 7, C2: 61 };
  const compatible = await calculateWorkbook(structuredClone(input));
  const result = await nativeFormula(native, root, 'evaluate', { view: input });
  assert(result.supported, result.reason);
  for (const r of result.results) {
    assert.equal(r.value, expected[r.ref]);
    assert.equal(compatible.sheets[0].cells[r.ref].calculation.value, expected[r.ref]);
  }
  // B4 also corrects scalar numeric nonzero-digit TRUNC's negative sign handling.
  const zero = await calculateWorkbook(
    view({
      A1: formula('TRUNC(-3,0)'),
      A2: formula('DATE(1900,3,1)'),
      A3: formula('DATE(2024,1,1)'),
      A4: formula('TRUNC(-2.7,1)'),
      A5: formula('IFERROR(TRUNC(1/0),8)'),
    })
  );
  for (const [ref, value] of Object.entries({ A1: -3, A2: 61, A3: 45292, A4: -2.7, A5: 8 }))
    assert.equal(zero.sheets[0].cells[ref].calculation.value, value, ref);
  const guarded = view({ A1: formula('DATE(1900,2,29)') });
  guarded.sheets[0].date1904 = true;
  const typed = await nativeFormula(native, root, 'evaluate', { view: guarded });
  assert.equal(typed.supported, true);
  assert.equal(typed.results[0].error, '#NUM!');
  assert.equal(Object.hasOwn(typed.results[0], 'value'), false);
  assert.deepEqual((await calculateWorkbook(guarded)).sheets[0].cells.A1.calculation, {
    status: 'error',
    error: '#NUM!',
    revision: guarded.revision,
  });
});
test('nonzero-digit TRUNC preserves magnitude at decimal boundaries and feeds dependents', async () => {
  const cases = [
    [-2.7, 1, -2.7],
    [-2.79, 1, -2.7],
    [-1.23, 2, -1.23],
    [-1.005, 2, -1],
    [-120, -1, -120],
    [-129, -1, -120],
    [-1234, -2, -1200],
    [-0.009, 2, 0],
    [-2.79, 1.9, -2.7],
    [-129, -1.9, -120],
    [-1e-20, 22, -1e-20],
  ];
  const cells = {};
  cases.forEach(([value, digits, expected], i) => {
    const r = i + 1;
    cells[`A${r}`] = number(value);
    cells[`B${r}`] = number(digits);
    cells[`C${r}`] = formula(`TRUNC(A${r},B${r})`);
    cells[`D${r}`] = formula(`C${r}*2`);
    cells[`E${r}`] = formula(`TRUNC(ABS(A${r}),B${r})`);
  });
  const result = await calculateWorkbook(view(cells));
  cases.forEach(([, , expected], i) => {
    for (const [col, want] of [
      ['C', expected],
      ['D', expected * 2],
      ['E', Math.abs(expected)],
    ])
      assert.equal(result.sheets[0].cells[`${col}${i + 1}`].calculation.value, want, `${col}${i + 1}`);
  });
  const errors = await calculateWorkbook(
    view({ A1: formula('IFERROR(TRUNC(-2.7,1/0),99)'), A2: formula('IFERROR(TRUNC(1/0,1),98)') })
  );
  assert.equal(errors.sheets[0].cells.A1.calculation.value, 99);
  assert.equal(errors.sheets[0].cells.A2.calculation.value, 98);
});
test('whole-column structural references preserve anchors, cross-sheet scope and deletion errors', async () => {
  const rewrite = async (extra, items) => {
    const r = await nativeFormula(native, root, 'rewrite', {
      target: 'Data',
      index: 1,
      count: 1,
      axis: 'column',
      ...extra,
      items,
    });
    assert(r.supported, r.reason);
    return r.formulas;
  };
  assert.deepEqual(
    await rewrite({}, [
      { owner: 'Data', formula: 'SUM(A:C)+SUM($B:$D)' },
      { owner: 'Other', formula: 'SUM(Data!A:C)+SUM(A:C)' },
      { owner: '', nameBoundary: true, formula: 'Data!$A:$C' },
    ]),
    ['SUM(A:D)+SUM($C:$E)', 'SUM(Data!A:D)+SUM(A:C)', 'Data!$A:$D']
  );
  assert.deepEqual(await rewrite({ remove: true }, [{ owner: 'Data', formula: 'SUM(A:C)+SUM(B:B)+SUM(C:D)' }]), [
    'SUM(A:B)+SUM(#REF!)+SUM(B:C)',
  ]);
  assert.deepEqual(await rewrite({ axis: 'row', index: 0, count: 2 }, [{ owner: 'Data', formula: 'SUM(A:C)' }]), [
    'SUM(A:C)',
  ]);
  assert.deepEqual(await rewrite({}, [{ formula: 'SUM(A:$C)+SUM($B:D)', copyRows: 5, copyColumns: 1 }]), [
    'SUM(B:$C)+SUM($B:E)',
  ]);
  const underflow = await nativeFormula(native, root, 'rewrite', { items: [{ formula: 'SUM(A:A)', copyColumns: -1 }] });
  assert.equal(underflow.supported, true, underflow.reason);
  assert.deepEqual(underflow.formulas, ['SUM(#REF!)']);
});
test('copying formulas across worksheet boundaries preserves syntax and emits reference errors', async () => {
  const vectors = [
    ['=A1', -1, 0, '=#REF!'],
    ['A1', 0, -1, '#REF!'],
    ['$A1+A$1+$A$1', -1, 0, '#REF!+A$1+$A$1'],
    ['$A1+A$1+$A$1', 0, -1, '$A1+#REF!+$A$1'],
    ["'Other Sheet'!A1", -1, 0, "'Other Sheet'!#REF!"],
    ['IFERROR(A1,99)', -1, 0, 'IFERROR(#REF!,99)'],
    ['"A1"&$A$1', -1, -1, '"A1"&$A$1'],
    ['SUM(1:1)', -1, 0, 'SUM(#REF!)'],
    ['SUM(A:A)', -1, 0, 'SUM(A:A)'],
    ['SUM(1:1)', 0, -1, 'SUM(1:1)'],
    ['XFD1048576', 1, 0, '#REF!'],
    ['XFD1048576', 0, 1, '#REF!'],
  ];
  const items = vectors.map(([formula, copyRows, copyColumns]) => ({ formula, copyRows, copyColumns })),
    before = structuredClone(items),
    result = await nativeFormula(native, root, 'rewrite', { items });
  assert.equal(result.supported, true, result.reason);
  assert.deepEqual(
    result.formulas,
    vectors.map((x) => x[3])
  );
  assert.deepEqual(items, before);
  for (const offset of ['-1', -0.5, null, true]) {
    const invalid = await nativeFormula(native, root, 'rewrite', { items: [{ formula: '1+1', copyRows: offset }] });
    assert.equal(invalid.supported, false);
    assert.match(invalid.reason, /Invalid copy row offset/);
    assert.equal(invalid.formulas, undefined);
  }
});
test('actual Excel Fill Up oracle collapses every partly invalid reference range', async () => {
  // Owned Book2 source B2:J2 filled up to B1:J1; exact formula-bar AX captures.
  const vectors = [
    ['SUM(A1:A2)', 'SUM(#REF!)'],
    ['SUM($A1:A$3)', 'SUM(#REF!)'],
    ['SUM(Sheet1!A1:A2)', 'SUM(Sheet1!#REF!)'],
    ['SUM($A$1:A1)', 'SUM(#REF!)'],
    ['SUM(A:A)', 'SUM(A:A)'],
    ['SUM(1:2)', 'SUM(#REF!)'],
    ['IFERROR(A1,99)', 'IFERROR(#REF!,99)'],
    ['"A1"&$A$1', '"A1"&$A$1'],
    ['A1', '#REF!'],
  ];
  const result = await nativeFormula(native, root, 'rewrite', {
    items: vectors.map(([formula]) => ({ formula, copyRows: -1 })),
  });
  assert.equal(result.supported, true, result.reason);
  assert.deepEqual(
    result.formulas,
    vectors.map((x) => x[1])
  );
});
test('actual Excel Fill Left oracle distinguishes normal and pre-existing broken ranges', async () => {
  // Owned A4:B8 Fill Left: Sheet1 prefix was normalized at entry in B6.
  const vectors = [
    ['SUM(A1:B1)', 'SUM(#REF!)'],
    ['SUM(A:B)', 'SUM(#REF!)'],
    ['Sheet1!A1', 'Sheet1!#REF!'],
    ['SUM(#REF!:A2)', 'SUM(#REF!:#REF!)'],
    ['SUM(A2:#REF!)', 'SUM(#REF!:#REF!)'],
    ['SUM(#REF!:B2)', 'SUM(#REF!:A2)'],
    ['SUM(B2:#REF!)', 'SUM(A2:#REF!)'],
  ];
  const result = await nativeFormula(native, root, 'rewrite', {
    items: vectors.map(([formula]) => ({ formula, copyColumns: -1 })),
  });
  assert.equal(result.supported, true, result.reason);
  assert.deepEqual(
    result.formulas,
    vectors.map((x) => x[1])
  );
});
test('sheet rename updates qualified reference errors without changing bare errors or literal text', async () => {
  const vectors = [
    ['Sheet1!#REF!', "'New Sheet'!#REF!"],
    ['SUM(A2:Sheet1!#REF!)', "SUM(A2:'New Sheet'!#REF!)"],
    ['Sheet1!#REF!+Other!#REF!+#REF!+"Sheet1!#REF!"', '\'New Sheet\'!#REF!+Other!#REF!+#REF!+"Sheet1!#REF!"'],
    ['#REF!+"Sheet1!#REF!"', '#REF!+"Sheet1!#REF!"'],
  ];
  const result = await nativeFormula(native, root, 'rename', {
    old: 'Sheet1',
    new: 'New Sheet',
    items: vectors.map(([formula]) => ({ formula })),
  });
  assert.equal(result.supported, true, result.reason);
  assert.deepEqual(
    result.formulas,
    vectors.map((x) => x[1])
  );
  const apostrophe = await nativeFormula(native, root, 'rename', {
    old: "O'Brien",
    new: "Changed'Sheet",
    items: [{ formula: "SUM('O''Brien'!#REF!:A2)" }],
  });
  assert.equal(apostrophe.supported, true, apostrophe.reason);
  assert.deepEqual(apostrophe.formulas, ["SUM('Changed''Sheet'!#REF!:A2)"]);
  const calculated = await nativeFormula(native, root, 'evaluate', {
    view: view({ B1: formula('COUNTA(Sheet1!#REF!)') }),
  });
  assert.equal(calculated.supported, true, calculated.reason);
  assert.equal(calculated.results[0].value, 1);
  const array = await nativeFormula(native, root, 'rewrite', { items: [{ formula: '{Sheet1!#REF!}', copyRows: 1 }] });
  assert.equal(array.supported, false);
  assert.equal(array.formulas, undefined);
});
test('qualified errors and broken range endpoints can be rewritten and caught by IFERROR', async () => {
  const vectors = [
    ['SUM(#REF!:A2)', 1, 'SUM(#REF!:A3)'],
    ['SUM(A2:#REF!)', -1, 'SUM(A1:#REF!)'],
    ["SUM('Other Sheet'!#REF!:A2)", 1, "SUM('Other Sheet'!#REF!:A3)"],
    ["IFERROR('Other Sheet'!#REF!,99)", 1, "IFERROR('Other Sheet'!#REF!,99)"],
    ['SUM(#REF!:#REF!)', 1, 'SUM(#REF!:#REF!)'],
  ];
  const result = await nativeFormula(native, root, 'rewrite', {
    items: vectors.map(([formula, copyRows]) => ({ formula, copyRows })),
  });
  assert.equal(result.supported, true, result.reason);
  assert.deepEqual(
    result.formulas,
    vectors.map((x) => x[2])
  );
  const cells = { A2: formula('A2') };
  ['#REF!', "'Other Sheet'!#REF!", 'SUM(#REF!:A2)', 'SUM(A2:#REF!)', 'SUM(#REF!:#REF!)'].forEach(
    (f, i) => (cells[`B${i + 1}`] = formula(`IFERROR(${f},99)`))
  );
  const input = view(cells);
  delete input.sheets[0].cells.A2;
  const calculated = await nativeFormula(native, root, 'evaluate', { view: input });
  assert.equal(calculated.supported, true, calculated.reason);
  assert.deepEqual(
    calculated.results.map((x) => x.value),
    [99, 99, 99, 99, 99]
  );
  for (const formula of ['A1:#DIV/0!', '#REF!:SUM(A1)', 'A1:MissingName', 'A1:B2:#REF!', 'SUM((#REF!:Other!A1,A2))'])
    assert.equal(
      (await nativeFormula(native, root, 'rewrite', { items: [{ formula, copyRows: -1 }] })).supported,
      false,
      formula
    );
});

test('unused metadata/external names do not block unrelated fallback calculation', async () => {
  const input = view({ A1: number(3), B1: formula('LEN("abc")+A1'), B2: formula('Used*2'), B3: formula('External+1') });
  input.namedRanges = [
    { name: '_xlnm.Print_Area', sheet: 'Data', formula: 'Data!$A$1:$B$9' },
    { name: 'Unused', formula: "'[missing.xlsx]Data'!$A$1" },
    { name: 'External', formula: "'[missing.xlsx]Data'!$A$1" },
    { name: 'Used', formula: 'Nested+1' },
    { name: 'Nested', formula: 'Data!$A$1' },
  ];
  // A referenced external name must remain unresolved, including its dependents.
  const output = await calculateWorkbook(input);
  assert.equal(output.sheets[0].cells.B1.calculation.value, 6);
  assert.equal(output.sheets[0].cells.B2.calculation.value, 8);
  assert.equal(output.sheets[0].cells.B3.calculation.status, 'unresolved');
  assert.deepEqual(output.namedRanges, input.namedRanges);
});

test('scalar/reference division uses direct binary64 division in both evaluators', async () => {
  const cells = { A1: number(64827.36), B1: number(971171.12) };
  const forms = [
    'SUM(A1)/B1',
    'A1/B1',
    '64827.36/B1',
    'SUM(A1)/(B1+1)',
    'SUM(A1)/ABS(B1)',
    '(-SUM(A1))/B1',
    'ROUND(A1,2)/B1',
    'IF(TRUE,A1,0)/B1',
    'IFERROR(A1,0)/B1',
  ];
  forms.forEach((formula, i) => (cells[`C${i + 1}`] = { formula }));
  const input = view(cells),
    expected = await calculateWorkbook(structuredClone(input)),
    actual = await nativeFormula(native, root, 'evaluate', { view: input });
  assert(actual.supported, actual.reason);
  for (const cell of actual.results) {
    if (cell.ref === 'C6') {
      assert.equal(expected.sheets[0].cells.C6.calculation.status, 'unresolved');
      assert.equal(
        cell.value,
        -actual.results.find((r) => r.ref === 'C1').value,
        'native unary minus remains independently correct where the compatibility parser refuses'
      );
    } else
      assert.equal(
        cell.value,
        expected.sheets[0].cells[cell.ref].calculation.value,
        `${cell.ref}: ${cells[cell.ref].formula}`
      );
  }
});

test('native LEN reads qualified literal references and preserves coercion guards', async () => {
  const cells = {
    A1: { value: '中文', dataType: 'text' },
    A2: { value: true, dataType: 'boolean' },
    A3: number(-123),
    A4: { value: 'a"b', dataType: 'text' },
  };
  const cases = [
    ['LEN(A1)', 2],
    ['LEN(A2)', 4],
    ['LEN(A3)', 4],
    ['LEN(A4)', 3],
    ['LEN(A5)', 0],
    ['LEN("a""b")', 3],
    ['LEN("中文")', 2],
    ['LEN(FALSE)', 5],
  ];
  cases.forEach(([f], i) => (cells[`B${i + 1}`] = formula(f)));
  const input = view(cells),
    actual = await nativeFormula(native, root, 'evaluate', { view: input }),
    expected = await calculateWorkbook(structuredClone(input));
  assert(actual.supported, actual.reason);
  for (const [i, [f, n]] of cases.entries()) {
    const ref = `B${i + 1}`;
    assert.equal(actual.results.find((r) => r.ref === ref).value, n, f);
    assert.equal(expected.sheets[0].cells[ref].calculation.value, n, f);
  }
  for (const [f, cell] of [
    ['LEN("🙂")', number(1)],
    ['LEN(A1)', number(0.1)],
    ['LEN(A1)', { value: '#REF!', dataType: 'error' }],
  ])
    assert.equal(
      (await nativeFormula(native, root, 'evaluate', { view: view({ A1: cell, B1: formula(f) }) })).supported,
      false,
      f
    );
});
test('native two-argument TRUNC exactly follows qualified signed precision semantics', async () => {
  const cells = {};
  let row = 0;
  for (const n of [-120, -3.9, -2.7, -1.005, -0.00000007, 0, 0.00000007, 1.005, 2.7, 3.9, 120])
    for (const d of [-6, -2, -1, -0.9, 0, 0.9, 1, 1.9, 2, 6])
      if (!(d === 0 && n < 0 && n > -1)) cells[`A${++row}`] = formula(`TRUNC(${n},${d})`);
  const input = view(cells);
  input.sheets[0].rows = row;
  const actual = await nativeFormula(native, root, 'evaluate', { view: input }),
    expected = await calculateWorkbook(structuredClone(input));
  assert(actual.supported, actual.reason);
  for (const item of actual.results)
    assert.equal(item.value, expected.sheets[0].cells[item.ref].calculation.value, cells[item.ref].formula);
  for (const f of ['TRUNC(-0.01,0)', 'TRUNC(1,100)', 'TRUNC(1,-100)', 'TRUNC(1E20,2)', 'TRUNC("3",1)'])
    assert.equal(
      (await nativeFormula(native, root, 'evaluate', { view: view({ A1: formula(f) }) })).supported,
      false,
      f
    );
});

test('numeric SUM uses the actual Excel oracle for range, scalar and named-table inputs', async () => {
  const oracle = JSON.parse(await fs.readFile(new URL('../fixtures/excel-b5-oracle.json', import.meta.url), 'utf8'));
  const cells = { A1: { value: 'Value', dataType: 'text' }, A2: number(0.1), A3: number(0.2), A4: number(0.3) };
  const cases = {
    D1: 'SUM(Prices[Value])*Rate',
    D2: 'SUM(A2:A4)*2',
    D3: 'SUM(0.1,0.2,0.3)*2',
    D6: 'SUM(A2:A3)',
    D7: 'SUM(0.1,0.2)',
    D8: 'SUM(A2:A4)',
    D9: 'SUM(10000000000000000,1,-10000000000000000)',
    D10: 'SUM(0.1,0.2,0.3)',
  };
  for (const [ref, f] of Object.entries(cases)) cells[ref] = formula(f);
  const input = view(cells);
  input.tables = [
    { name: 'Prices', sheet: 'Data', ref: 'A1:A4', columns: [{ name: 'Value' }], headerRowCount: 1, totalsRowCount: 0 },
  ];
  input.namedRanges = [{ name: 'Rate', formula: '2' }];
  const compatibility = structuredClone(input);
  compatibility.namedRanges[0].name = 'Multiplier';
  compatibility.sheets[0].cells.D1.formula = 'SUM(Prices[Value])*Multiplier';
  const result = await nativeFormula(native, root, 'evaluate', { view: input }),
    fallback = await calculateWorkbook(compatibility);
  assert(result.supported, result.reason);
  for (const item of result.results) {
    const expected = oracle.values.Numbers[Number(item.ref.slice(1)) - 1][3];
    assert.equal(item.value, expected, item.ref);
    assert.equal(fallback.sheets[0].cells[item.ref].calculation.value, expected, item.ref);
  }
});

test('native TRUNC precision factors agree across the declared -12..12 interval', async () => {
  const cells = {};
  let seed = 0x7b51,
    row = 0;
  for (let i = 0; i < 100; i++) {
    seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
    const n = ((seed % 2000001) - 1000000) / 10 ** (i % 7);
    for (let d = -12; d <= 12; d++) {
      if (Math.abs(n) * 10 ** d > 9e15 || (d === 0 && n < 0 && n > -1)) continue;
      cells[`A${++row}`] = formula(`TRUNC(${n},${d})`);
    }
  }
  const input = view(cells);
  input.sheets[0].rows = row;
  const expected = await calculateWorkbook(structuredClone(input)),
    actual = await nativeFormula(native, root, 'evaluate', { view: input });
  assert(actual.supported, actual.reason);
  for (const item of actual.results)
    assert.equal(item.value, expected.sheets[0].cells[item.ref].calculation.value, cells[item.ref].formula);
});

test('SUM accumulates continuously across ranges and matches saved Excel cancellation cases', async () => {
  const cases = JSON.parse(await fs.readFile(new URL('../fixtures/excel-b5-sum-order.json', import.meta.url), 'utf8'));
  for (const { formula: f, value } of cases) {
    const input = view({ A1: number(-1e16), A2: number(1), B1: number(1e16), B2: number(-1), D1: formula(f.slice(1)) });
    const actual = await nativeFormula(native, root, 'evaluate', { view: input }),
      compatible = await calculateWorkbook(structuredClone(input));
    assert.equal(compatible.sheets[0].cells.D1.calculation.value, value, f);
    if (!f.includes('{') && !f.includes('TRUE')) {
      assert(actual.supported, actual.reason);
      assert.equal(actual.results[0].value, value, f);
    } else assert.equal(actual.supported, false, 'Array constants/scalar coercion remain compatibility cases');
  }
});

test('B6 numeric multiplication and division match saved Excel scalars and references', async () => {
  const oracle = JSON.parse(await fs.readFile(new URL('../fixtures/excel-b6-runtime.json', import.meta.url))),
    cells = { A1: number(0.1), A2: number(0.2), A3: number(0.3) };
  for (const item of oracle.cases.filter((c) => ['D4', 'D5', 'D6', 'D7'].includes(c.ref)))
    cells[item.ref] = formula(item.formula);
  const input = view(cells),
    actual = await nativeFormula(native, root, 'evaluate', { view: input }),
    compatible = await calculateWorkbook(structuredClone(input));
  assert(actual.supported, actual.reason);
  for (const item of actual.results) {
    const expected = oracle.cases.find((c) => c.ref === item.ref).expected;
    assert.equal(item.value, expected);
    assert.equal(compatible.sheets[0].cells[item.ref].calculation.value, expected);
  }
});
test('native dimensions, COUNTA and dynamic references retain values and safe refusal', async () => {
  const cases = [
    ['ROWS(A1:B3)', 3],
    ['COLUMNS(A1:B3)', 2],
    ['COUNTA(A1:A5)', 4],
    ['COUNTBLANK(A1:A5)', 2],
    ['COUNTBLANK(A1:A1048576)', 1048573],
    ['COUNTA("",TRUE,#REF!)', 3],
    ['INDIRECT("A1")', 2],
    ['INDIRECT(B1)', 3],
    ['SUM(INDIRECT("A1:A2"))', 5],
    ['OFFSET(A1,1,0)', 3],
    ['SUM(OFFSET(A1,0,0,2,1))', 5],
    ['IF(TRUE,INDIRECT("A1"),INDIRECT("A2"))', 2],
  ];
  const cells = {
    A1: number(2),
    A2: number(3),
    A3: { value: '', dataType: 'text' },
    A4: { value: false, dataType: 'boolean' },
    B1: { value: 'A2', dataType: 'text' },
  };
  cases.forEach(([f], i) => (cells[`D${i + 1}`] = formula(f)));
  const actual = await nativeFormula(native, root, 'evaluate', { view: view(cells) });
  assert(actual.supported, actual.reason);
  cases.forEach(([, n], i) => assert.equal(actual.results.find((c) => c.ref === `D${i + 1}`).value, n));
  for (const f of [
    'INDIRECT("R1C1",FALSE)',
    'INDIRECT("[external.xlsx]Data!A1")',
    'OFFSET(A1,-1,0)',
    'OFFSET(A1,0,0,0,1)',
    'INDIRECT("D1")',
    'IF(TRUE,1,OFFSET(A1,-1,0))',
  ]) {
    const input = view({ A1: number(2), B1: formula('"A1"'), D1: formula(f) });
    assert.equal((await nativeFormula(native, root, 'evaluate', { view: input })).supported, false, f);
  }
});

test('B7 authored grouping and final addition/subtraction match the saved Excel matrix', async () => {
  const oracle = JSON.parse(await fs.readFile(new URL('../fixtures/excel-b7-add-sub.json', import.meta.url))),
    cells = { A1: number(0.1), A2: number(0.2), A3: number(0.3) };
  for (const item of oracle.cases.filter((c) => !/^ROUND(?:UP|DOWN)/.test(c.formula)))
    cells[item.ref] = formula(item.formula);
  const input = view(cells);
  input.sheets[0].rows = 150;
  const actual = await nativeFormula(native, root, 'evaluate', { view: input });
  assert(actual.supported, actual.reason);
  const shared = await calculateWorkbook(structuredClone(input));
  for (const item of actual.results) {
    const expected = oracle.cases.find((c) => c.ref === item.ref).expected;
    assert.equal(item.value, expected, 'native ' + cells[item.ref].formula);
    if (!cells[item.ref].formula.includes('INDIRECT'))
      assert.equal(shared.sheets[0].cells[item.ref].calculation.value, expected, 'shared ' + cells[item.ref].formula);
  }
});
test('native typed strings and formula-derived selectors preserve types and reject cycles', async () => {
  const input = view({
    A1: formula('LEFT("Bxxx",1)&2'),
    A2: formula('IF("yes"="YES",A1,"B3")'),
    B2: number(17),
    B3: number(23),
    C1: formula('INDIRECT(A2)'),
    C2: formula('LEN(A1)'),
    C3: formula('TRIM("  x   y ")'),
    C4: formula('IF(TRUE,"ok",123)'),
    C5: formula('IF(FALSE,"ok",123)'),
    C6: formula('AND(TRUE,NOT(FALSE))'),
  });
  const actual = await nativeFormula(native, root, 'evaluate', { view: input });
  assert(actual.supported, actual.reason);
  const values = Object.fromEntries(actual.results.map((r) => [r.ref, r.value]));
  assert.deepEqual(values, { A1: 'B2', A2: 'B2', C1: 17, C2: 2, C3: 'x y', C4: 'ok', C5: 123, C6: true });
  for (const cells of [
    { A1: formula('A2&"x"'), A2: formula('A1&"y"') },
    { A1: formula('INDIRECT(A2)'), A2: formula('"A1"') },
    { A1: formula('UPPER("ß")') },
    { A1: formula('LEFT("🙂",1)') },
  ])
    assert.equal((await nativeFormula(native, root, 'evaluate', { view: view(cells) })).supported, false);
});

test('calculation-only arithmetic wrappers never leak into authored formulas or become public functions', async () => {
  const input = view({
    A1: number(0.1),
    A2: number(0.2),
    B1: formula('A1+A2'),
    B2: formula('PLATFORM_B7_ADD(A1,A2)'),
    B3: formula('B2+1'),
  });
  const calculated = await calculateWorkbook(input);
  assert.equal(calculated.sheets[0].cells.B1.formula, 'A1+A2');
  assert.equal(calculated.sheets[0].cells.B1.calculation.value, 0.1 + 0.2);
  assert.equal(calculated.sheets[0].cells.B2.calculation.status, 'unresolved');
  assert.equal(calculated.sheets[0].cells.B3.calculation.status, 'unresolved');
  const transformed = await transformWorkbook(input, [{ sheet: 'Data', operation: 'insertRow', index: 0 }]);
  assert.equal(transformed.sheets[0].cells.B2.formula, 'A2+A3');
});

test('typed formula results cannot authorize nonfinite wire literals or excessive text', async () => {
  for (const value of [{ $workbookNumber: 'Infinity' }, { $workbookNumber: 'NaN' }, ['unexpected']]) {
    const input = view({ A1: { value }, B1: formula('A1') });
    assert.equal((await nativeFormula(native, root, 'evaluate', { view: input })).supported, false);
  }
  assert.equal(
    (
      await nativeFormula(native, root, 'evaluate', {
        view: view({ A1: { value: 'x'.repeat(32768) }, B1: formula('A1') }),
      })
    ).supported,
    false
  );
});

test('native scalar text and explicit coercion functions agree with independent values and shared engine', async () => {
  const entries = [
    ['REPT("ab",2.9)', 'abab'],
    ['REPT("",1000000000)', ''],
    ['SUBSTITUTE("aba aba","aba","X")', 'X X'],
    ['SUBSTITUTE("aaa","aa","b")', 'ba'],
    ['SUBSTITUTE("abc","","X")', 'abc'],
    ['SUBSTITUTE("a-a-a","a","X",2.9)', 'a-X-a'],
    ['SUBSTITUTE("a","a","X",9)', 'a'],
    ['EXACT("a","A")', false],
    ['EXACT(12,"12")', true],
    ['N(TRUE)', 1],
    ['N(FALSE)', 0],
    ['N("12")', 0],
    ['N(A1)', 0],
    ['N(A2)', 1],
    ['N(A3)', 0],
    ['T(12)', ''],
    ['T(A1)', '12'],
    ['T(A3)', ''],
    ['ISNUMBER(A1)', false],
    ['ISNUMBER(12)', true],
    ['ISTEXT(A1)', true],
    ['ISLOGICAL(A2)', true],
    ['ISBLANK(A3)', true],
    ['ISBLANK("")', false],
    ['ISNUMBER(1/0)', false],
    ['IF(ISNUMBER(12),N(TRUE)+2,0)', 3],
    ['IFERROR(REPT("a",-1),"bad")', 'bad'],
    ['IFERROR(SUBSTITUTE("a","a","b",0),"bad")', 'bad'],
  ];
  const cells = { A1: { value: '12', dataType: 'text' }, A2: { value: true, dataType: 'boolean' } };
  entries.forEach(([f], i) => (cells[`B${i + 1}`] = formula(f)));
  const input = view(cells),
    actual = await nativeFormula(native, root, 'evaluate', { view: input }),
    shared = await calculateWorkbook(structuredClone(input));
  assert(actual.supported, actual.reason);
  for (const item of actual.results) {
    const expected = entries[Number(item.ref.slice(1)) - 1][1];
    assert.equal(item.value, expected, item.ref);
    assert.equal(shared.sheets[0].cells[item.ref].calculation.value, expected, 'shared ' + item.ref);
  }
  for (const f of ['REPT("ab",20000)', 'SUBSTITUTE(REPT("a",1000),"a",REPT("b",1000))', 'N(A1:A2)', 'T(A1:A2)'])
    assert.equal(
      (await nativeFormula(native, root, 'evaluate', { view: view({ B1: formula(f) }) })).supported,
      false,
      f
    );
});

test('native text/type results match the saved Excel 16.113 oracle including negative fractional REPT', async () => {
  const oracle = JSON.parse(await fs.readFile(new URL('../fixtures/excel-native-text22.json', import.meta.url))),
    cells = { A1: { value: '12', dataType: 'text' }, A2: { value: true, dataType: 'boolean' } };
  for (const c of oracle.cases)
    cells[c.ref] = formula(c.expected === '#VALUE!' ? `IFERROR(${c.formula},"bad")` : c.formula);
  const input = view(cells),
    actual = await nativeFormula(native, root, 'evaluate', { view: input }),
    shared = await calculateWorkbook(structuredClone(input));
  assert(actual.supported, actual.reason);
  for (const c of oracle.cases) {
    const expected = c.expected === '#VALUE!' ? 'bad' : c.expected;
    assert.equal(actual.results.find((r) => r.ref === c.ref).value, expected, c.ref);
    assert.equal(shared.sheets[0].cells[c.ref].calculation.value, expected, 'shared ' + c.ref);
  }
});

test('B8 scalar arithmetic coercion matches 35 saved Excel cases and shared evaluation', async () => {
  const oracle = JSON.parse(await fs.readFile(new URL('../fixtures/excel-scenarios-coercion24.json', import.meta.url)));
  for (const c of oracle.cases) {
    const input = view({
        A1: { value: '12', dataType: 'text' },
        A2: { value: true, dataType: 'boolean' },
        [c.ref]: formula(c.formula),
      }),
      result = await nativeFormula(native, root, 'evaluate', { view: input });
    assert(result.supported, result.reason);
    assert.equal(result.results[0].value, c.expected, c.formula);
    const old = await calculateWorkbook(input);
    assert.equal(old.sheets[0].cells[c.ref].calculation.value, c.expected, c.formula);
  }
  for (const f of [
    '"9007199254740991"+1',
    '"1"+9007199254740991',
    'SUM("12",1)',
    '"2026/1/1"+1',
    '"€12"+1',
    '"$1,200"+1',
    '"1,234e2"+1',
    '"2.5"^0.3',
  ])
    assert.equal(
      (await nativeFormula(native, root, 'evaluate', { view: view({ A1: formula(f) }) })).supported,
      false,
      f
    );
});

test('whole-row reference rewriting preserves anchors, sheets, axis semantics and boundary errors', async () => {
  for (const [formula, copyRows, copyColumns, expected] of [
    ['SUM(1:3)', 2, 5, 'SUM(3:5)'],
    ['SUM($1:3)', 2, -5, 'SUM($1:5)'],
    ["SUM('Other Sheet'!$2:$4)", 3, 4, "SUM('Other Sheet'!$2:$4)"],
    ['=SUM(1:1)', 2, 0, '=SUM(3:3)'],
  ]) {
    const result = await nativeFormula(native, root, 'rewrite', { items: [{ formula, copyRows, copyColumns }] });
    assert(result.supported, result.reason);
    assert.equal(result.formulas[0], expected);
  }
  for (const [axis, index, count, remove, expected] of [
    ['column', 0, 3, false, 'SUM(1:3)'],
    ['row', 1, 2, true, 'SUM(1:1)'],
    ['row', 0, 3, true, 'SUM(#REF!)'],
  ]) {
    const result = await nativeFormula(native, root, 'rewrite', {
      target: 'Data',
      axis,
      index,
      count,
      remove,
      items: [{ formula: 'SUM(1:3)', owner: 'Data' }],
    });
    assert(result.supported, result.reason);
    assert.equal(result.formulas[0], expected);
  }
  const overflow = await nativeFormula(native, root, 'rewrite', {
    items: [{ formula: 'SUM(1048576:1048576)', copyRows: 1 }],
  });
  assert.equal(overflow.supported, true, overflow.reason);
  assert.deepEqual(overflow.formulas, ['SUM(#REF!)']);
});

test('genuine native errors have an explicit error field while error-looking text stays a value', async () => {
  const input = view({
    A1: number(2),
    B1: formula('A1/0'),
    C1: formula('IFERROR(B1,7)'),
    D1: { value: '#N/A', dataType: 'text' },
    E1: formula('IFERROR(D1,99)'),
  });
  const result = await nativeFormula(native, root, 'evaluate', { view: input });
  assert.equal(result.supported, true);
  assert.equal(result.results.find((c) => c.ref === 'B1').error, '#DIV/0!');
  assert.equal(
    Object.hasOwn(
      result.results.find((c) => c.ref === 'B1'),
      'value'
    ),
    false
  );
  assert.equal(result.results.find((c) => c.ref === 'C1').value, 7);
  assert.equal(result.results.find((c) => c.ref === 'E1').value, '#N/A');
  const emptyAverage = await nativeFormula(native, root, 'evaluate', { view: view({ B1: formula('AVERAGE(A3:A4)') }) });
  assert.equal(emptyAverage.results[0].error, '#DIV/0!');
});
