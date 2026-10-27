import { test } from 'node:test';
import assert from 'node:assert/strict';
import { literalFillSnapshot } from '../core/workbook-commands.mjs';
import { projectWorkbookCommands } from '../core/workbook-projection.mjs';
const engine = { shiftFormulas: (items) => items.map((item) => item.formula + '+0') };
test('Authored literal fill preserves declared errors and exact literal text without type inference', () => {
  for (const value of ['#N/A', '#DIV/0!', '=SUM(A1)', '00123', "'literal", 'TRUE', '雪😀', ''])
    assert.deepEqual(literalFillSnapshot({ value, dataType: 'text' }), { value });
  assert.deepEqual(literalFillSnapshot({ value: '#N/A', dataType: 'error', error: '#N/A' }), {
    value: '#N/A',
    valueType: 'error',
  });
  for (const value of [null, true, false, 7.5, -0]) assert(Object.is(literalFillSnapshot({ value }).value, value));
  assert.throws(() => literalFillSnapshot({ value: '#not-an-error', dataType: 'error' }), {
    code: /^(internal-invalid-field|session-unsupported)$/,
  });
  assert.throws(() => literalFillSnapshot({ value: 9007199254740992, rawValue: '9007199254740993' }), /precision/);
});
test('Literal fill rejects every known derived/current/stored record rather than borrowing Values policy', () => {
  for (const marker of [
    { formula: '1' },
    { formulaType: 'array' },
    { formulaOutput: { master: 'A1', range: 'A1:A2' } },
    { spill: { master: 'A1' } },
    { scenarioResult: true },
    { calculation: { status: 'calculated', value: 913 } },
    { calculation: { status: 'stored', pending: true } },
  ])
    assert.throws(() => literalFillSnapshot({ value: 913, ...marker }), {
      code: /^(fill-unsupported|formula-unsupported)$/,
    });
});
test('Portable down/right and ordinary internal Paste keep error versus text and source style intent', async () => {
  const input = {
      sheets: [
        {
          name: 'S',
          cells: {
            A1: { value: '#N/A', dataType: 'error', error: '#N/A', styleId: 1, format: 2, bold: true, italic: true },
            A2: { value: '#N/A', dataType: 'text', styleId: 0, format: 0 },
            B1: { value: 8, styleId: 0 },
            B2: { value: 9, styleId: 0 },
          },
        },
      ],
    },
    before = structuredClone(input);
  const result = await projectWorkbookCommands(
      input,
      [
        { type: 'fillRange', sheet: 'S', source: 'A1:A2', range: 'B1:B2' },
        { type: 'fillRange', sheet: 'S', source: 'A1', range: 'C1:E1' },
      ],
      engine
    ),
    cells = result.view.sheets[0].cells;
  for (const ref of ['B1', 'C1', 'D1', 'E1']) {
    assert.equal(cells[ref].dataType, 'error');
    assert.equal(cells[ref].error, '#N/A');
    assert.equal(cells[ref].value, '#N/A');
    assert.equal(cells[ref].styleId, 1);
    assert.equal(cells[ref].format, 2);
    assert.equal(cells[ref].bold, true);
    assert.equal(cells[ref].italic, true);
  }
  assert.equal(cells.B2.dataType, 'text');
  assert.equal(cells.B2.error, undefined);
  assert.deepEqual(input, before);
  assert.equal(result.edits.filter((edit) => edit.valueType === 'error').length, 4);
});
test('Formula branch retains authored formula semantics even when its original cache is pending/error', async () => {
  const input = {
    sheets: [
      {
        name: 'S',
        cells: {
          A1: {
            formula: 'B1+1',
            value: '#N/A',
            dataType: 'error',
            calculation: { status: 'stored', pending: true },
            styleId: 1,
          },
        },
      },
    ],
  };
  const result = await projectWorkbookCommands(
    input,
    [{ type: 'fillRange', sheet: 'S', source: 'A1', range: 'A2' }],
    engine
  );
  assert.equal(result.view.sheets[0].cells.A2.formula, 'B1+1+0');
  assert.equal(result.edits[0].valueType, undefined);
  assert.equal(result.edits[0].value, undefined);
});
test('One internal Cut plan clears only its source and later protected source refuses without a prefix', async () => {
  const input = {
      sheets: [
        {
          name: 'S',
          cells: {
            A1: { value: '#REF!', dataType: 'error', error: '#REF!', styleId: 1 },
            A2: { value: 914, formulaOutput: { schema: 1, master: 'Z1', range: 'Z1:Z2' } },
          },
        },
      ],
    },
    before = structuredClone(input);
  const moved = await projectWorkbookCommands(
    input,
    [
      { type: 'fillRange', sheet: 'S', source: 'A1', range: 'B1' },
      { type: 'clearRange', sheet: 'S', range: 'A1' },
    ],
    engine
  );
  assert.equal(moved.view.sheets[0].cells.B1.dataType, 'error');
  assert.equal(moved.view.sheets[0].cells.A1.value, null);
  assert.equal(moved.view.sheets[0].cells.A1.error, undefined);
  assert.equal(moved.edits.length, 2);
  await assert.rejects(
    projectWorkbookCommands(
      input,
      [
        { type: 'fillRange', sheet: 'S', source: 'A1', range: 'B1' },
        { type: 'fillRange', sheet: 'S', source: 'A2', range: 'B2' },
      ],
      engine
    ),
    { code: /^(fill-unsupported|formula-unsupported)$/ }
  );
  assert.deepEqual(input, before);
});
