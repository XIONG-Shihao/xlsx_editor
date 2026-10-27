import test, { after } from 'node:test';
import os from 'node:os';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { calculateWorkbook, createCalculationSession, transformWorkbook } from '../core/calculation-node.mjs';
import { readWorkbook } from '../core/workbook-package.mjs';

const corpus = fs.mkdtempSync(path.join(os.tmpdir(), 'xlsx-reference-corpus-'));
after(() => fs.rmSync(corpus, { recursive: true, force: true }));
execFileSync('python3', [fileURLToPath(new URL('../fixtures/build-corpus.py', import.meta.url)), '--output', corpus]);
const sheet = (name, cells) => ({ name, rows: 80, columns: 30, cells });
const book = (cells) => ({ revision: 'reference-input:0', sheets: [sheet('Main', cells)] });
function freeze(value) {
  if (value && typeof value === 'object' && !Object.isFrozen(value)) {
    Object.freeze(value);
    for (const child of Object.values(value)) freeze(child);
  }
  return value;
}
async function equivalent(input, expectedMode) {
  const original = structuredClone(input);
  freeze(input);
  let projection;
  const projected = await calculateWorkbook(input, { onInputProjection: (value) => (projection = value) });
  const complete = await calculateWorkbook(input, { inputProjection: 'full' });
  assert.deepEqual(
    projected,
    complete,
    'Projected input must preserve every logical value, formula, status and revision'
  );
  assert.deepEqual(input, original, 'Input remains immutable');
  if (expectedMode) assert.equal(projection.mode, expectedMode, projection.reason);
  return { output: projected, projection };
}

for (const name of [
  'pricing-schedule',
  'budget',
  'calculation-states',
  'dates-1900',
  'dates-1904',
  'preservation-rich',
  'response-matrix',
  'protected',
]) {
  test(`referenced and complete engine inputs agree for independent corpus ${name}`, async () => {
    const bytes = fs.readFileSync(path.join(corpus, `${name}.xlsx`));
    const input = readWorkbook(bytes);
    for (const current of input.sheets) current.cells.Z70 = { value: 'Unrelated original text' };
    await equivalent(input);
  });
}

test('point/range/whole-column/whole-row references preserve all occupied inputs and dimensions', async () => {
  const cases = [
    book({ A1: { value: 2 }, A2: { value: 4 }, B1: { formula: 'SUM(A:A)' }, Z70: { value: 'untouched' } }),
    book({
      A1: { value: 2 },
      Z1: { value: 7 },
      A2: { value: 4 },
      B4: { formula: 'SUM(1:2)' },
      Z70: { value: 'untouched' },
    }),
    book({
      A1: { value: 2 },
      A2: { value: 4 },
      B1: { value: 3 },
      B2: { value: 6 },
      C1: { formula: 'SUMPRODUCT(A1:A2,B1:B2)' },
      D1: { formula: 'INDEX(A1:B2,2,2)' },
      E1: { formula: 'IF($A$1=2,SUM(A1:B2),0)' },
      Z70: { value: 'untouched' },
    }),
  ];
  for (const input of cases) {
    const { projection } = await equivalent(input, 'static-references');
    assert(projection.sheets[0].engineCells < projection.sheets[0].sourceCells);
    assert.equal(projection.sheets[0].rowCount, 80);
    assert.equal(projection.sheets[0].columnCount, 30);
  }
});

test('scoped/recursive names and table ranges retain cross-sheet dependency closure', async () => {
  const input = {
    revision: 'names',
    sheets: [
      sheet('Main', {
        A1: { value: 2 },
        B1: { formula: 'AdjustedSeed*3' },
        C1: { formula: 'SUM(Prices[Amount])' },
        D1: { formula: "'Other Sheet'!B2+$A$1" },
        Z70: { value: 'untouched' },
      }),
      sheet('Other Sheet', {
        A1: { value: 'Item' },
        B1: { value: 'Amount' },
        A2: { value: 'one' },
        B2: { value: 5 },
        A3: { value: 'two' },
        B3: { formula: 'LocalInput*2' },
        Z70: { value: 'untouched' },
      }),
    ],
    namedRanges: [
      { name: 'SeedValue', formula: 'Main!$A$1' },
      { name: 'AdjustedSeed', formula: 'SeedValue' },
      { name: 'LocalInput', sheet: 'Other Sheet', formula: '$B$2' },
    ],
    tables: [{ name: 'Prices', sheet: 'Other Sheet', ref: 'A1:B3', columns: [{ name: 'Item' }, { name: 'Amount' }] }],
  };
  const { output, projection } = await equivalent(input, 'static-references');
  assert.equal(output.sheets[0].cells.B1.calculation.value, 6);
  assert.equal(output.sheets[0].cells.C1.calculation.value, 15);
  assert(projection.sheets.every((item) => item.engineCells < item.sourceCells));
});

test('cycles, ordinary errors and unsafe literal precision preserve status propagation', async () => {
  await equivalent(
    book({
      A1: { value: 0 },
      A2: { formula: '1/A1' },
      A3: { formula: 'IFERROR(A2,99)' },
      B1: { formula: 'B2+1' },
      B2: { formula: 'B1+1' },
      C1: { value: 9007199254740992, rawValue: '9007199254740993', dataType: 'number' },
      C2: { formula: 'SUM(C1)' },
      D1: { value: '#NAME?', dataType: 'text' },
      D2: { formula: 'IFERROR(D1,99)' },
      Z70: { value: 'untouched' },
    }),
    'static-references'
  );
});

test('defined expressions preserve arithmetic precedence without changing authored names', async () => {
  const input = {
    revision: 'named-expressions',
    sheets: [
      sheet('Main', {
        A1: { value: 2 },
        B1: { formula: 'AdjustedSeed*3' },
        C1: { formula: '10/ScaleValue' },
        D1: { formula: 'NestedSeed*2' },
        Z70: { value: 'untouched' },
      }),
      sheet('Local', { A1: { value: 4 }, B1: { formula: 'LocalAdjustedSeed*3' }, Z70: { value: 'untouched' } }),
    ],
    namedRanges: [
      { name: 'SeedValue', formula: 'Main!$A$1' },
      { name: 'AdjustedSeed', formula: 'SeedValue+1' },
      { name: 'ScaleValue', formula: 'Main!$A$1*2' },
      { name: 'NestedSeed', formula: 'AdjustedSeed+1' },
      { name: 'LocalAdjustedSeed', sheet: 'Local', formula: '$A$1+2' },
    ],
  };
  const { output } = await equivalent(input, 'static-references');
  assert.equal(output.sheets[0].cells.B1.calculation.value, 9);
  assert.equal(output.sheets[0].cells.C1.calculation.value, 2.5);
  assert.equal(output.sheets[0].cells.D1.calculation.value, 8);
  assert.equal(output.sheets[1].cells.B1.calculation.value, 18);
  assert.deepEqual(output.namedRanges, input.namedRanges);
  const session = await createCalculationSession(input);
  try {
    assert.equal(session.view.sheets[0].cells.B1.calculation.value, 9);
    assert.equal(session.view.sheets[1].cells.B1.calculation.value, 18);
  } finally {
    session.dispose();
  }
});

test('unsupported shadowed defined names stay explicit instead of producing a false result', async () => {
  const input = {
    revision: 'shadowed-names',
    sheets: [sheet('Main', { A1: { value: 2 }, B1: { formula: 'SeedValue*3' } })],
    namedRanges: [
      { name: 'SeedValue', formula: 'Main!$A$1' },
      { name: 'SeedValue', sheet: 'Main', formula: '$A$1+1' },
    ],
  };
  const { output } = await equivalent(input);
  assert.equal(output.sheets[0].cells.B1.calculation.status, 'unresolved');
  assert.match(output.sheets[0].cells.B1.calculation.error, /nameDuplicate/);
  assert.equal(output.sheets[0].cells.B1.calculation.value, undefined);
});

test('owned calculation sessions transfer only an explicitly private view', async () => {
  const initial = book({ A1: { value: 2 }, B1: { formula: 'A1*3' }, Z70: { value: 'untouched' } });
  const original = structuredClone(initial),
    source = freeze(initial);
  const immutable = await createCalculationSession(source);
  try {
    assert.notEqual(immutable.view, source);
    assert.notEqual(immutable.view.sheets[0].cells.B1, source.sheets[0].cells.B1);
    assert.deepEqual(source, original);
    assert.equal(immutable.view.sheets[0].cells.B1.calculation.value, 6);
  } finally {
    immutable.dispose();
  }
  const privateView = structuredClone(source),
    literalRecord = privateView.sheets[0].cells.Z70;
  const owned = await createCalculationSession(privateView, { takeOwnership: true });
  try {
    assert.equal(owned.view, privateView);
    assert.equal(owned.view.sheets[0].cells.Z70, literalRecord);
    assert.equal(owned.view.sheets[0].cells.B1.calculation.value, 6);
    assert.deepEqual(source, original);
  } finally {
    owned.dispose();
  }
});

test('unproved contexts retain full inputs while already-blocked dynamic formulas need no literal projection', async () => {
  for (const formula of [
    'IFERROR(UNKNOWN_LOCAL(A1),99)',
    'INDIRECT("A1")',
    'SUM(A1:INDEX(C1:C3,2,1))',
    'SUM(Main:Last!A1)',
    'A1#',
  ]) {
    await equivalent(
      book({
        A1: { value: 2 },
        B1: { value: 3 },
        C1: { value: 4 },
        C2: { value: 5 },
        C3: { value: 6 },
        D1: { formula },
        Z70: { value: 'untouched' },
      }),
      formula === 'INDIRECT("A1")' ? 'static-references' : 'full'
    );
  }
});

test('functions that expand implicit result ranges keep the complete input', async () => {
  for (const [formula, expected] of [
    ['SUMIF(A1:A3,">0",B1)', 60],
    ['SUMIF(A1:A3,">0",B1:B2)', 60],
    ['AVERAGEIF(A1:A3,">0",B1)', 20],
  ]) {
    const { output } = await equivalent(
      book({
        A1: { value: 1 },
        A2: { value: 2 },
        A3: { value: 3 },
        B1: { value: 10 },
        B2: { value: 20 },
        B3: { value: 30 },
        D1: { formula },
        Z70: { value: 'untouched' },
      }),
      'full'
    );
    assert.equal(output.sheets[0].cells.D1.calculation.value, expected);
  }
});

test('persistent edits to omitted or previously blank literals agree with complete recalculation', async () => {
  const initial = book({ A1: { value: 2 }, B1: { formula: 'A1+A2' }, Z70: { value: 10 } });
  const session = await createCalculationSession(initial);
  try {
    let next = structuredClone(session.view);
    for (const [ref, value] of [
      ['Z70', 20],
      ['A2', 5],
      ['A1', 4],
    ]) {
      next = structuredClone(next);
      next.revision = `changed:${ref}`;
      next.sheets[0].cells[ref] = { value };
      const patches = await session.update(next, [{ sheet: 'Main', cell: ref }]);
      for (const patch of patches) next.sheets[0].cells[patch.cell] = patch.value;
      const full = await calculateWorkbook(next, { inputProjection: 'full' });
      assert.equal(next.sheets[0].cells.B1.calculation.value, full.sheets[0].cells.B1.calculation.value);
      if (ref === 'Z70') assert.equal(patches.length, 0);
    }
    next.sheets[0].cells.C1 = { formula: 'Z70*2' };
    const rebuilt = await createCalculationSession(next);
    try {
      assert.equal(rebuilt.view.sheets[0].cells.C1.calculation.value, 40);
    } finally {
      rebuilt.dispose();
    }
    assert.equal(initial.sheets[0].cells.Z70.value, 10);
  } finally {
    session.dispose();
  }
});

test('adding the first formula to a large literal view avoids copying unrelated engine inputs losslessly', async () => {
  const cells = {};
  for (let row = 1; row <= 80000; row++) cells[`A${row}`] = { value: row, rawValue: String(row), dataType: 'number' };
  cells.B1 = { formula: 'A1+A80000' };
  const input = { revision: 'large-first-formula', sheets: [{ name: 'Main', rows: 80000, columns: 2, cells }] };
  const { output, projection } = await equivalent(input, 'static-references');
  assert.equal(output.sheets[0].cells.B1.calculation.value, 80001);
  assert.equal(projection.sheets[0].sourceCells, 80001);
  assert.equal(projection.sheets[0].engineCells, 3);
  assert.equal(Object.keys(output.sheets[0].cells).length, 80001);
});

test('row insertion has no arbitrary batch ceiling and retains the XLSX format boundary', async () => {
  const input = {
    revision: 'rows',
    sheets: [{ name: 'Main', rows: 2, columns: 1, cells: { A1: { value: 'first' }, A2: { value: 'last' } } }],
  };
  const output = await transformWorkbook(input, [{ sheet: 'Main', operation: 'insertRow', index: 0, count: 10001 }]);
  assert.equal(output.sheets[0].cells.A10002.value, 'first');
  assert.equal(output.sheets[0].cells.A10003.value, 'last');
  await assert.rejects(
    () => transformWorkbook(input, [{ sheet: 'Main', operation: 'insertRow', index: 0, count: 1048576 }]),
    { code: 'formula-limit' }
  );
  assert.equal(input.sheets[0].cells.A1.value, 'first');
});
