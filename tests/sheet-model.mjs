import assert from 'node:assert/strict';
import {
  applyGridEdit,
  displayValue,
  formulaDisplay,
  maxColumnCount,
  plainSheets,
  parseGridInput,
} from '@office-editors/xlsx-editor/ui/sheetModel';
import * as before from '../fixtures/spreadsheet-model-before.ts';
import { calculateWorkbook } from '@office-editors/xlsx-editor/core/calculation-node.mjs';
const sheet = (cells) => ({
  name: 'Sheet1',
  state: 'visible',
  path: '',
  cells,
  rows: 20,
  columns: 8,
  merges: [],
  widths: [],
});
assert.equal(maxColumnCount(Array.from({ length: 200_000 }, () => ['value'])), 1);
const diagnostic = sheet({
  A1: { formula: 'IF(1=1,10,0)', value: 999, cachedValue: 999 },
  A2: { value: 5 },
  B1: { formula: 'SUM(A1:A2)' },
});
assert.equal(
  before.formulaDisplay(diagnostic, 'B1'),
  5,
  'frozen historical evaluator reproduces the partial-total bug'
);
assert.equal(formulaDisplay(diagnostic, 'A1'), 'Calculating…', 'stored999 is never presented as a fresh result');
const calculated = (await calculateWorkbook({ revision: 'r1', sheets: [diagnostic] })).sheets[0];
assert.equal(formulaDisplay(calculated, 'B1'), 15);
assert.equal(calculated.cells.A1.cachedValue, 999);
const changed = applyGridEdit(calculated, { sheet: 'Sheet1', cell: 'A2', value: 7 });
const updated = (await calculateWorkbook({ revision: 'r2', sheets: [changed] })).sheets[0];
assert.equal(formulaDisplay(updated, 'B1'), 17);
assert.equal(formulaDisplay(calculated, 'B1'), 15, 'undo frame remains unchanged');
for (const [status, expected] of [
  ['unresolved', 'Unresolved'],
  ['circular', '#CYCLE!'],
  ['error', '#DIV/0!'],
])
  assert.equal(
    formulaDisplay(sheet({ A1: { formula: '1/0', calculation: { status, error: '#DIV/0!' } } }), 'A1'),
    expected
  );
assert.equal(displayValue(sheet({ A1: { value: 0.25, format: 10 } }), 'A1'), '25.00%');
assert.equal(displayValue(sheet({ A1: { value: 60, format: 14 } }), 'A1'), '1900-02-29');
assert.equal(displayValue({ ...sheet({ A1: { value: 0, format: 14 } }), date1904: true }, 'A1'), '1904-01-01');
let current = sheet({ A1: { value: 2 } });
for (let index = 0; index < 1000; index++)
  current = applyGridEdit(current, { sheet: 'Sheet1', cell: 'Z100', value: index });
assert.equal(current.cells.Z100.value, 999);
assert.equal(current.rows, 100);
assert.equal(current.columns, 26);
assert.equal(Object.keys(current.cells).length, 2);
assert.deepEqual(structuredClone(plainSheets([current])), plainSheets([current]));
for (const operation of ['insertRow', 'deleteRow', 'insertColumn', 'deleteColumn'])
  assert.deepEqual(
    JSON.parse(JSON.stringify(applyGridEdit(current, { operation, index: 0 }))),
    before.applyGridEdit(JSON.parse(JSON.stringify(current)), { operation, index: 0 })
  );
for (const [text, expected] of [
  ['00123', { value: '00123' }],
  ["'=SUM(A1)", { value: '=SUM(A1)' }],
  ['=SUM(A1)', { formula: 'SUM(A1)' }],
  ['TRUE', { value: true }],
  ['false', { value: false }],
  ['', { value: null }],
  ['9007199254740993', { value: '9007199254740993' }],
  ['0.25', { value: 0.25 }],
  ['東京', { value: '東京' }],
])
  assert.deepEqual(parseGridInput(text), expected);
assert.deepEqual(parseGridInput('=SUM(A1)', true), { value: '=SUM(A1)' });
console.log(
  'PASS authoritative calculation display, historical IF15 regression, explicit formula states, typed inputs, immutable frames, dates and recovery projection'
);
const { inputDisplay, applyCalculatedPatches, prepareCellRecord } = await import('@office-editors/xlsx-editor/ui/sheetModel');
const unsafe = { value: 9007199254740992, rawValue: '9007199254740993', dataType: 'number', format: 2 };
assert.equal(inputDisplay(unsafe), '9007199254740993');
assert.equal(displayValue(sheet({ A1: unsafe }), 'A1'), '9007199254740993');
const immutable = sheet({
  A1: { value: 1 },
  B1: { formula: 'A1+1', calculation: { status: 'calculated', value: 2, revision: 'old' } },
});
prepareCellRecord(immutable.cells);
const patched = applyCalculatedPatches(
  immutable,
  { A1: { value: 3 }, B1: { formula: 'A1+1', calculation: { status: 'calculated', value: 4, revision: 'new' } } },
  'new'
);
assert.equal(immutable.cells.A1.value, 1);
assert.equal(immutable.cells.B1.calculation.value, 2);
assert.equal(patched.cells.B1.calculation.value, 4);
assert.equal(patched.calculationRevision, 'new');

assert.equal(
  inputDisplay({
    formulaType: 'shared',
    value: 999,
    rawValue: '999',
    calculation: { status: 'calculated', value: 12 },
  }),
  '12',
  'Shared followers expose fresh results to the formula bar and clipboard'
);
assert.equal(
  inputDisplay({ scenarioResult: true, value: 6, rawValue: '6', calculation: { status: 'calculated', value: 21 } }),
  '21'
);
console.log('PASS fresh special-formula edit and clipboard projection');
