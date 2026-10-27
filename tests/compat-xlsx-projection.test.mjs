import test from 'node:test';
import assert from 'node:assert/strict';
import { projectWorkbookCommands } from '../core/workbook-projection.mjs';
import * as engine from '../core/calculation-node.mjs';
const book = (cells) => ({
  revision: 'old',
  readOnly: false,
  limitations: [],
  sheets: [{ name: 'Main', rows: 20, columns: 8, cells, merges: [], widths: [] }],
});
test('paste/setCells is atomic and literal types are preserved', async () => {
  const input = book({ A1: { value: 1 } }),
    before = structuredClone(input);
  await assert.rejects(
    () =>
      projectWorkbookCommands(
        input,
        [
          {
            type: 'setCells',
            sheet: 'Main',
            cells: [
              { cell: 'A1', value: 2 },
              { cell: 'B1', value: { bad: 1 } },
            ],
          },
        ],
        engine
      ),
    { code: 'internal-invalid-field' }
  );
  assert.deepEqual(input, before);
  const out = await projectWorkbookCommands(
    input,
    [
      {
        type: 'setCells',
        sheet: 'Main',
        cells: [
          { cell: 'A1', value: '00123' },
          { cell: 'B1', value: '=literal' },
          { cell: 'C1', value: true },
          { cell: 'D1', value: null },
        ],
      },
    ],
    engine
  );
  assert.deepEqual(
    ['A1', 'B1', 'C1', 'D1'].map((ref) => out.view.sheets[0].cells[ref].value),
    ['00123', '=literal', true, null]
  );
  assert.equal(out.edits.length, 4);
});
test('fill uses parsed relative/mixed references and commits a complete batch', async () => {
  const input = book({ A1: { formula: '$B1+C$2+"A1"', bold: true } }),
    out = await projectWorkbookCommands(
      input,
      [{ type: 'fillRange', sheet: 'Main', source: 'A1', range: 'A2:A3' }],
      engine
    );
  assert.equal(out.view.sheets[0].cells.A2.formula, '$B2+C$2+"A1"');
  assert.equal(out.view.sheets[0].cells.A3.formula, '$B3+C$2+"A1"');
  assert.equal(out.view.sheets[0].cells.A3.bold, true);
  assert.equal(out.edits.length, 2);
});
test('sort preserves row associations and relative row formulas; totals outside keep their text', async () => {
  const input = book({
    A1: { value: 'B' },
    B1: { value: 2 },
    C1: { formula: 'B1*10' },
    A2: { value: 'A' },
    B2: { value: 3 },
    C2: { formula: 'B2*10' },
  });
  const out = await projectWorkbookCommands(
    input,
    [{ sheet: 'Main', operation: 'sortRows', range: 'A1:C2', column: 'A', direction: 'asc' }],
    engine
  );
  assert.equal(out.view.sheets[0].cells.A1.value, 'A');
  assert.equal(out.view.sheets[0].cells.B1.value, 3);
  assert.equal(out.view.sheets[0].cells.C1.formula, 'B1*10');
  assert.equal((await engine.calculateWorkbook(out.view)).sheets[0].cells.C1.calculation.value, 30);
  await assert.rejects(
    () =>
      projectWorkbookCommands(
        input,
        [{ sheet: 'Main', operation: 'sortRows', range: 'A1:B2', column: 'A', direction: 'asc' }],
        engine
      ),
    { code: 'sort-bad-input' }
  );
  // A total outside the sorted rows keeps its text, as in Excel.
  input.sheets[0].cells.D3 = { formula: 'SUM(B1:B2)' };
  const totalled = await projectWorkbookCommands(
    input,
    [{ sheet: 'Main', operation: 'sortRows', range: 'A1:C2', column: 'A', direction: 'asc' }],
    engine
  );
  assert.equal(totalled.view.sheets[0].cells.D3.formula, 'SUM(B1:B2)');
});
test('filter retains cells, resize splits width spans, freeze is bounded', async () => {
  const input = book({ A1: { value: 'keep' }, A2: { value: 'hide' }, B2: { value: 99 } });
  input.sheets[0].widths = [{ min: 1, max: 4, width: 10, hidden: false }];
  const out = await projectWorkbookCommands(
    input,
    [
      { sheet: 'Main', operation: 'filterRows', range: 'A1:B2', column: 'A', values: ['keep'] },
      { sheet: 'Main', operation: 'resizeColumn', index: 1, width: 25 },
      { sheet: 'Main', operation: 'freeze', rows: 1, columns: 1 },
    ],
    engine
  );
  assert.deepEqual(out.view.sheets[0].hiddenRows, [1]);
  assert.equal(out.view.sheets[0].cells.B2.value, 99);
  assert.deepEqual(
    out.view.sheets[0].widths.map((w) => [w.min, w.max, w.width]),
    [
      [1, 1, 10],
      [2, 2, 25],
      [3, 4, 10],
    ]
  );
  assert.deepEqual(out.view.sheets[0].freeze, { rows: 1, columns: 1 });
  await assert.rejects(
    () => projectWorkbookCommands(input, [{ sheet: 'Main', operation: 'freeze', rows: 1000, columns: 0 }], engine),
    { code: 'layout-malformed' }
  );
});
test('fill retains the source style identity and refuses raw numeric precision loss atomically', async () => {
  const input = book({
    A1: { value: 12.5, styleId: 2, format: '$0.00', bold: true, italic: false },
    B1: { value: 9007199254740992, rawValue: '9007199254740993', dataType: 'number' },
  });
  const out = await projectWorkbookCommands(
    input,
    [{ type: 'fillRange', sheet: 'Main', source: 'A1', range: 'A2:A3' }],
    engine
  );
  assert.equal(out.edits[0].styleId, 2);
  assert.equal(out.edits[0].bold, undefined);
  assert.equal(out.view.sheets[0].cells.A2.format, '$0.00');
  assert.equal(out.view.sheets[0].cells.A2.bold, true);
  const before = structuredClone(input);
  await assert.rejects(
    () =>
      projectWorkbookCommands(input, [{ type: 'fillRange', sheet: 'Main', source: 'A1:B1', range: 'A2:B2' }], engine),
    /precision/
  );
  assert.deepEqual(input, before);
});
test('row sizing uses bounded points and preserves other row dimensions', async () => {
  const input = book({ A1: { value: 'heading' } });
  input.sheets[0].rowHeights = [
    { row: 0, height: 20 },
    { row: 4, height: 40 },
  ];
  const out = await projectWorkbookCommands(
    input,
    [{ sheet: 'Main', operation: 'resizeRow', index: 0, height: 60 }],
    engine
  );
  assert.deepEqual(out.view.sheets[0].rowHeights, [
    { row: 0, height: 60 },
    { row: 4, height: 40 },
  ]);
  assert.deepEqual(input.sheets[0].rowHeights, [
    { row: 0, height: 20 },
    { row: 4, height: 40 },
  ]);
  await assert.rejects(
    () => projectWorkbookCommands(input, [{ sheet: 'Main', operation: 'resizeRow', index: 0, height: 201 }], engine),
    { code: /^(internal-invalid-field|print-bad-input)$/ }
  );
});
test('sort and insertion move row heights and hidden state with existing rows', async () => {
  const input = book({ A1: { value: 'B' }, A2: { value: 'A' } });
  input.sheets[0].rowHeights = [
    { row: 0, height: 20 },
    { row: 1, height: 60 },
  ];
  input.sheets[0].hiddenRows = [0];
  const sorted = await projectWorkbookCommands(
    input,
    [{ sheet: 'Main', operation: 'sortRows', range: 'A1:A2', column: 'A', direction: 'asc' }],
    engine
  );
  assert.deepEqual(sorted.view.sheets[0].rowHeights, [
    { row: 0, height: 60 },
    { row: 1, height: 20 },
  ]);
  assert.deepEqual(sorted.view.sheets[0].hiddenRows, [1]);
  const inserted = await projectWorkbookCommands(input, [{ sheet: 'Main', operation: 'insertRow', index: 0 }], engine);
  assert.deepEqual(inserted.view.sheets[0].rowHeights, [
    { row: 1, height: 20 },
    { row: 2, height: 60 },
  ]);
  assert.deepEqual(inserted.view.sheets[0].hiddenRows, [1]);
});
test('bounded cell projection shares normalization and guards without touching accepted history', async () => {
  const { projectCellCommands } = await import('../core/workbook-projection.mjs');
  const original = {
      sheets: [
        {
          name: 'Main',
          rows: 1,
          columns: 2,
          cells: {
            A1: { value: 1, rawValue: '1' },
            B1: { formula: 'A1+1', calculation: { value: 2, status: 'calculated' } },
          },
        },
      ],
    },
    before = structuredClone(original);
  const projected = projectCellCommands(original, [
    { sheet: 'Main', cell: 'A1', value: 2, expectedValue: 1 },
    { type: 'setCells', sheet: 'Main', cells: [{ cell: 'A1', value: 3, expectedValue: 2 }] },
  ]);
  assert.equal(projected.changes.length, 1);
  assert.equal(projected.changes[0].value.value, 3);
  assert.equal(projected.edits.length, 2);
  assert.deepEqual(original, before);
  assert.throws(
    () =>
      projectCellCommands(original, [
        { sheet: 'Main', cell: 'A1', value: 2 },
        { sheet: 'Main', cell: 'A1', value: 3, expectedValue: 1 },
      ]),
    /Conflict/
  );
  assert.deepEqual(original, before);
  assert.equal(projectCellCommands(original, [{ sheet: 'Main', operation: 'insertRow', index: 0 }]), null);
});
test('sort and filter refuse raw numeric key precision loss atomically', async () => {
  const source = {
      sheets: [
        {
          name: 'Main',
          rows: 3,
          columns: 1,
          merges: [],
          cells: {
            A1: { value: 'Key' },
            A2: { value: 9007199254740992, rawValue: '9007199254740993', dataType: 'number' },
            A3: { value: 2 },
          },
        },
      ],
    },
    before = structuredClone(source);
  for (const edit of [
    { sheet: 'Main', operation: 'sortRows', range: 'A2:A3', column: 'A', direction: 'asc' },
    { sheet: 'Main', operation: 'filterRows', range: 'A1:A3', column: 'A', values: [9007199254740992] },
  ])
    await assert.rejects(() => projectWorkbookCommands(source, [edit], engine), { code: 'workbook-unsupported' });
  assert.deepEqual(source, before);
});

test('column resizing preserves hidden state in the view just as in the saved XML', async () => {
  const input = book({ A1: { value: 'keep' } });
  input.sheets[0].widths = [{ min: 1, max: 4, width: 12, hidden: true }];
  const out = await projectWorkbookCommands(
    input,
    [{ operation: 'resizeColumn', sheet: 'Main', index: 1, width: 27 }],
    engine
  );
  assert.deepEqual(out.view.sheets[0].widths, [
    { min: 1, max: 1, width: 12, hidden: true },
    { min: 2, max: 2, width: 27, hidden: true },
    { min: 3, max: 4, width: 12, hidden: true },
  ]);
  assert.deepEqual(input.sheets[0].widths, [{ min: 1, max: 4, width: 12, hidden: true }]);
});
