import { test } from 'node:test';
import assert from 'node:assert/strict';
import { normalizeValuePasteCell as snapshot, validateValueType } from '../core/value-paste.mjs';
import { resolveWorkbookEdits } from '../core/workbook-commands.mjs';
import { projectWorkbookCommands } from '../core/workbook-projection.mjs';
import { createWorkbook, editWorkbook, readWorkbook } from '../core/workbook-package.mjs';
import { unzipSync, strFromU8 } from 'fflate';
test('Literal Values preserves types, precision, signed zero and text without guessing', () => {
  for (const value of ['00123', '=SUM(A1)', 'TRUE', "'word", '#N/A', '', ' 雪😀\t\r\n'])
    assert.deepEqual(snapshot({ value, dataType: 'text' }), { value });
  for (const value of [null, true, false, 7.5, -0]) assert(Object.is(snapshot({ value }).value, value));
  assert.deepEqual(snapshot({ value: '#N/A', dataType: 'error' }), { value: '#N/A', valueType: 'error' });
  assert.throws(() => snapshot({ value: 9007199254740992, rawValue: '9007199254740993' }), /precision/);
  for (const value of [Infinity, NaN, {}, [], 'x'.repeat(32768), '\ud800']) assert.throws(() => snapshot({ value }));
});
test('Current scalar results only; pending overrides stale calculated and error fields', () => {
  for (const value of [2, 'computed', false]) {
    const cell = Object.freeze({
      value: 913,
      formula: 'A1',
      calculation: Object.freeze({ status: 'calculated', value }),
    });
    assert.deepEqual(snapshot(cell), { value });
  }
  for (const marker of [
    { formula: '1/0' },
    { spill: { master: 'A1' } },
    { formulaOutput: { schema: 1, master: 'A1', range: 'A1:A2' } },
    { scenarioResult: true },
  ])
    assert.deepEqual(snapshot({ ...marker, calculation: { status: 'error', error: '#DIV/0!' } }), {
      value: '#DIV/0!',
      valueType: 'error',
    });
  for (const status of ['stored', 'unresolved', 'circular', 'missing'])
    assert.throws(
      () =>
        snapshot({ formulaOutput: { master: 'A1' }, value: 913, calculation: { status, value: 913, error: '#N/A' } }),
      /current/
    );
  assert.throws(() => snapshot({ formulaOutput: { master: 'A1' }, value: 913 }), /current/);
  for (const status of ['calculated', 'error'])
    assert.throws(
      () => snapshot({ formula: '1', calculation: { status, value: 913, error: '#N/A', pending: true } }),
      /pending/
    );
});
test('Declared error is a closed narrow value type and cannot smuggle other cell datatypes', () => {
  for (const code of ['#N/A', '#DIV/0!', '#REF!', '#SPILL!']) validateValueType({ value: code, valueType: 'error' });
  for (const edit of [
    { value: '#N/A', valueType: 'text' },
    { value: 'word', valueType: 'error' },
    { value: 1, valueType: 'error' },
    { valueType: 'error' },
    { value: '#N/A', valueType: 'error', formula: '1' },
  ])
    assert.throws(() => validateValueType(edit), { code: /^(internal-invalid-field|session-unsupported)$/ });
  for (const edit of [
    { value: 'word', valueType: 'error' },
    { value: '#N/A', valueType: 'number' },
  ])
    assert.throws(
      () => resolveWorkbookEdits({ sheets: [{ name: 'S', cells: {} }] }, [{ sheet: 'S', cell: 'A1', ...edit }]),
      { code: /^(internal-invalid-field|session-unsupported)$/ }
    );
});
test('Portable Values preserves destination styles and error types without implicit calculation', async () => {
  const view = {
      revision: 'old',
      sheets: [
        {
          name: 'S',
          cells: {
            A1: { value: '#N/A', dataType: 'error' },
            A2: { value: '#N/A', dataType: 'text' },
            A3: { formula: '1/0', value: 913, calculation: { status: 'error', error: '#DIV/0!' } },
            B1: { value: 2, styleId: 7, bold: true },
            B2: { value: 3, styleId: 7 },
            B3: { value: 4, styleId: 7 },
          },
        },
      ],
    },
    before = structuredClone(view),
    engine = {
      calculateWorkbook() {
        assert.fail('No implicit calculation');
      },
    };
  const result = await projectWorkbookCommands(
      view,
      [{ type: 'pasteValues', sheet: 'S', source: 'A1:A3', range: 'B1:B3' }],
      engine
    ),
    cells = result.view.sheets[0].cells;
  assert.equal(cells.B1.dataType, 'error');
  assert.equal(cells.B1.error, '#N/A');
  assert.equal(cells.B2.dataType, 'text');
  assert.equal(cells.B3.dataType, 'error');
  assert.equal(cells.B1.styleId, 7);
  assert.equal(cells.B1.bold, true);
  assert.deepEqual(view, before);
  const pending = {
    ...view,
    sheets: [
      {
        ...view.sheets[0],
        cells: {
          ...view.sheets[0].cells,
          A3: { formula: '1', calculation: { status: 'calculated', value: 913, pending: true } },
        },
      },
    ],
  };
  await assert.rejects(
    projectWorkbookCommands(pending, [{ type: 'pasteValues', sheet: 'S', source: 'A1:A3', range: 'B1:B3' }], engine),
    /pending/
  );
  assert.deepEqual(view, before);
});
test('Portable OOXML error write differs exactly from the same literal error-looking text', () => {
  const original = createWorkbook(),
    sheet = readWorkbook(original).sheets[0].name,
    bytes = editWorkbook(original, [
      { sheet, cell: 'A1', value: '#N/A', valueType: 'error' },
      { sheet, cell: 'A2', value: '#N/A' },
      { sheet, cell: 'A3', value: 1, bold: true },
    ]),
    cells = readWorkbook(bytes).sheets[0].cells;
  assert.equal(cells.A1.dataType, 'error');
  assert.equal(cells.A1.value, '#N/A');
  assert.equal(cells.A2.dataType, 'text');
  assert.equal(cells.A2.value, '#N/A');
  assert.equal(cells.A3.bold, true);
  const xml = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']);
  assert.match(xml, /t="e"[^>]*><v>#N\/A<\/v>/);
  assert.match(xml, /t="inlineStr"/);
  assert.equal(
    readWorkbook(editWorkbook(bytes, [{ sheet, cell: 'A1', value: '#N/A' }])).sheets[0].cells.A1.dataType,
    'text'
  );
});
test('Portable target merge/table barriers refuse atomically while read-only source geometry remains usable', async () => {
  for (const kind of ['merge', 'table']) {
    const view = {
        revision: 'base',
        sheets: [
          { name: 'S', cells: { A1: { value: 5 }, B1: { value: 8 } }, merges: kind === 'merge' ? ['B1:C1'] : [] },
        ],
        tables: kind === 'table' ? [{ sheet: 'S', ref: 'B1:C2' }] : [],
      },
      before = structuredClone(view);
    await assert.rejects(
      projectWorkbookCommands(view, [{ type: 'pasteValues', sheet: 'S', source: 'A1', range: 'B1' }], {}),
      kind === 'merge' ? /merged/ : /table/
    );
    assert.deepEqual(view, before);
    const result = await projectWorkbookCommands(
      view,
      [{ type: 'pasteValues', sheet: 'S', source: 'B1', range: 'D1' }],
      {}
    );
    assert.equal(result.view.sheets[0].cells.D1.value, 8);
  }
});
test('Actual portable source array follower cannot acquire literal authority from its cached scalar', async () => {
  const { valuePasteFixture, valueSheet } = await import('./support/value-paste-fixture.mjs');
  const view = readWorkbook(valuePasteFixture({ boundary: 'unknown-array' }));
  const sheet = view.sheets.find((sheet) => sheet.name === valueSheet);
  assert.equal(sheet.cells.R1100.formulaType, 'array');
  assert.equal(sheet.cells.R1100.formulaReference, 'R1100:R1101');
  assert.equal(sheet.cells.R1101.value, 914);
  assert.equal(sheet.cells.R1101.formulaOutput, undefined);
  await assert.rejects(
    projectWorkbookCommands(view, [{ type: 'pasteValues', sheet: valueSheet, source: 'R1101', range: 'D11' }], {}),
    { code: /^(fill-bad-input|fill-unsupported|paste-bad-input)$/ }
  );
});
test('Captured source geometry refuses apparent-current unowned outputs and known incomplete inventory', async () => {
  const { createValuePasteSourceContext } = await import('../core/value-paste.mjs');
  const master = { formula: 'UNKNOWN()', formulaType: 'array', formulaReference: 'A1:A2' },
    sheet = { name: 'S', storedFormulaCount: 1, cells: { A1: master, A2: { value: 913 } } },
    cell = { value: 913, calculation: { status: 'calculated', value: 913 } };
  assert.throws(() => snapshot(cell, { sheet, cell: 'A2' }), { code: /^(fill-bad-input|paste-bad-input)$/ });
  assert.deepEqual(snapshot({ value: 'ordinary' }, { sheet, cell: 'C1' }), { value: 'ordinary' });
  assert.throws(() => snapshot(cell, { sheet: { ...sheet, cells: { A2: cell } }, cell: 'A2' }), /incomplete/);
  const owned = {
      ...sheet,
      cells: {
        A1: { ...master, spill: { schema: 1, master: 'A1', range: 'A1:A2', role: 'master' } },
        A2: {
          value: 2,
          spill: { schema: 1, master: 'A1', range: 'A1:A2', role: 'follower' },
          calculation: { status: 'calculated', value: 2 },
        },
      },
    },
    context = createValuePasteSourceContext(owned);
  assert.deepEqual(snapshot(owned.cells.A2, { ...context, cell: 'A2' }), { value: 2 });
  assert.throws(
    () =>
      snapshot(
        { ...owned.cells.A2, calculation: { status: 'calculated', value: 2, pending: true } },
        { ...context, cell: 'A2' }
      ),
    /pending/
  );
});
