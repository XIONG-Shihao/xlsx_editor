import test from 'node:test';
import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { fileURLToPath } from 'node:url';
import { projectWorkbookCommands } from '../core/workbook-projection.mjs';
const bundle = await build({
  entryPoints: [fileURLToPath(new URL('../ui/fillSeriesPreview.ts', import.meta.url))],
  bundle: true,
  write: false,
  platform: 'browser',
  format: 'esm',
  logLevel: 'silent',
});
const { createFillSeriesPreview, fillSeriesPreviewPage, FILL_SERIES_LIMITS, FILL_SERIES_POLICY } = await import(
  'data:text/javascript;base64,' + Buffer.from(bundle.outputFiles[0].contents).toString('base64')
);
const binding = (range) => ({
  documentId: 'book',
  ownerIdentity: 'native-owner',
  state: 'accepted-head',
  sheet: 'Data',
  selection: range,
  selectedCell: range.split(':')[0],
  activationEpoch: 1,
  navigationGeneration: 2,
  modeGeneration: 3,
  specialSelectionEpoch: 4,
  mode: 'normal',
  specialSelection: null,
});
const number = (value, rawValue = String(value)) => ({ value, rawValue, dataType: 'number' });
function input(range = 'A1:A6', cells = { A1: number(1), A2: number(3) }, extra = {}) {
  const sheet = {
    name: 'Data',
    path: 'xl/worksheets/sheet1.xml',
    state: 'visible',
    rows: 10000,
    columns: 30,
    cells: {},
    merges: [],
    widths: [],
    ...extra.sheet,
  };
  return { binding: binding(range), sheet, cells, readRange: range, ...extra, sheet };
}
function ready(value) {
  const result = createFillSeriesPreview(value);
  assert.equal(result.status, 'ready', result.reason);
  return result.preview;
}
function refused(value, code) {
  const result = createFillSeriesPreview(value);
  assert.equal(result.status, 'refused');
  assert.equal(result.code, code);
  assert(result.reason.length > 0);
  return result;
}

test('vertical two-seed policy preserves seeds and previews independently declared exact destinations', () => {
  const source = input(),
    before = structuredClone(source),
    preview = ready(source);
  assert.equal(preview.policy, 'linear-two-seeds-15-significant-digits');
  assert.equal(FILL_SERIES_POLICY, preview.policy);
  assert.equal(preview.orientation, 'vertical');
  assert.equal(preview.positions, 6);
  assert.equal(preview.destinationCount, 4);
  assert.equal(preview.step, 2);
  assert.deepEqual(preview.seeds, [
    { cell: 'A1', value: 1, rawValue: '1' },
    { cell: 'A2', value: 3, rawValue: '3' },
  ]);
  assert.deepEqual(preview.destinations, [
    { cell: 'A3', value: 5 },
    { cell: 'A4', value: 7 },
    { cell: 'A5', value: 9 },
    { cell: 'A6', value: 11 },
  ]);
  assert.deepEqual(preview.command, { type: 'fillSeries', sheet: 'Data', range: 'A1:A6' });
  assert.deepEqual(fillSeriesPreviewPage(preview, 0, 3), {
    offset: 0,
    total: 4,
    rows: [
      { cell: 'A3', value: 5 },
      { cell: 'A4', value: 7 },
      { cell: 'A5', value: 9 },
    ],
    nextOffset: 3,
  });
  assert.deepEqual(fillSeriesPreviewPage(preview, 3, 3), {
    offset: 3,
    total: 4,
    rows: [{ cell: 'A6', value: 11 }],
    nextOffset: null,
  });
  assert.deepEqual(source, before);
});

test('horizontal descending, constant, decimals, binary64 ties and zero match the bounded native arithmetic policy', () => {
  const descending = ready(input('B5:F5', { B5: number(10), C5: number(7) }));
  assert.equal(descending.orientation, 'horizontal');
  assert.equal(descending.step, -3);
  assert.deepEqual(descending.destinations, [
    { cell: 'D5', value: 4 },
    { cell: 'E5', value: 1 },
    { cell: 'F5', value: -2 },
  ]);
  assert.deepEqual(ready(input('A1:D1', { A1: number(3), B1: number(3) })).destinations, [
    { cell: 'C1', value: 3 },
    { cell: 'D1', value: 3 },
  ]);
  assert.deepEqual(ready(input('A1:A5', { A1: number(0.1, '0.100'), A2: number(0.2, '2e-1') })).destinations, [
    { cell: 'A3', value: 0.3 },
    { cell: 'A4', value: 0.4 },
    { cell: 'A5', value: 0.5 },
  ]);
  const tie = ready(input('A1:A3', { A1: number(123456789012342.5), A2: number(123456789012343.5) }));
  assert.equal(tie.destinations[0].value, 123456789012345);
  const zero = ready(input('A1:A4', { A1: number(-0, '-0'), A2: number(-0, '-0') }));
  assert(Object.is(zero.seeds[0].value, -0));
  assert(Object.is(zero.seeds[1].value, -0));
  assert(zero.destinations.every((cell) => Object.is(cell.value, 0)));
});

test('style-only blank destinations are eligible and preview agrees with the existing portable series projection', async () => {
  const source = input('D9:D13', {
      D9: { ...number(1), styleId: 2, format: 14 },
      D10: { ...number(3), bold: true },
      D11: { value: null, styleId: 5, format: 14, bold: true },
      D13: { value: null, rawValue: null, cachedValue: null, dataType: 'blank', italic: true },
    }),
    before = structuredClone(source);
  const preview = ready(source),
    view = {
      revision: 'accepted',
      readOnly: false,
      limitations: [],
      sheets: [{ ...source.sheet, cells: source.cells }],
    };
  const result = await projectWorkbookCommands(view, [preview.command], {});
  for (const destination of preview.destinations)
    assert.equal(result.view.sheets[0].cells[destination.cell].value, destination.value);
  assert.equal(result.view.sheets[0].cells.D11.styleId, 5);
  assert.equal(result.view.sheets[0].cells.D11.format, 14);
  assert.equal(result.view.sheets[0].cells.D11.bold, true);
  assert.deepEqual(result.view.sheets[0].cells.D9, source.cells.D9);
  assert.deepEqual(result.view.sheets[0].cells.D10, source.cells.D10);
  assert.deepEqual(source, before);
});

test('numeric seeds refuse inference, precision loss, formulas/current results, pending/stored and known special outputs', () => {
  const variants = [
    { value: '1' },
    { value: true },
    { value: null },
    number(Infinity),
    number(NaN),
    number(9007199254740992, '9007199254740993'),
    { ...number(1), formula: 'SUM(B1:B2)', calculation: { status: 'calculated', value: 1 } },
    { ...number(1), calculation: { status: 'calculated', value: 1 } },
    { ...number(1), calculation: { status: 'stored', pending: true } },
    { ...number(1), cachedResult: { authority: 'stored', status: 'calculated', value: 1 } },
    { ...number(1), cachedValue: 1 },
    { ...number(1), formulaType: 'shared' },
    { ...number(1), formulaOutput: { schema: 1, master: 'A10', range: 'A1:A10' } },
    { ...number(1), spill: { schema: 1, master: 'A1', range: 'A1:A5', role: 'master' } },
    { ...number(1), scenarioResult: true },
    { ...number(1), dataType: 'error' },
  ];
  for (const seed of variants) {
    const source = input('A1:A4', { A1: seed, A2: number(3) }),
      before = structuredClone(source);
    refused(source, 'invalid-seeds');
    assert.deepEqual(source, before);
  }
  refused(input('A1:A4', { A2: number(3) }), 'invalid-seeds');
  refused(input('A1:A4', { A1: number(1) }), 'invalid-seeds');
});

test('occupied and known protected targets refuse all destinations before any plan becomes ready', () => {
  for (const value of ['', 0, false, '#N/A', '=SUM(A1)'])
    refused(input('A1:A4', { A1: number(1), A2: number(3), A4: { value } }), 'occupied-target');
  for (const target of [
    { value: null, formula: 'UNKNOWN()' },
    { value: null, formulaType: 'array' },
    { value: null, formulaOutput: { schema: 1, master: 'A20', range: 'A3:A20' } },
    { value: null, spill: { schema: 1, master: 'B1', range: 'A1:C6', role: 'follower' } },
    { value: null, calculation: { status: 'stored', pending: true } },
    { value: null, cachedResult: { authority: 'stored', status: 'unresolved' } },
    { value: null, cachedValue: 913 },
    { value: null, dataType: 'error' },
  ])
    refused(input('A1:A4', { A1: number(1), A2: number(3), A4: target }), 'protected-target');
  refused(input('A1:A5', undefined, { sheet: { merges: ['A3:B3'] } }), 'protected-target');
  refused(
    input('A1:A5', undefined, { sheet: { tableRanges: [{ name: 'Table1', ref: 'A4:B8' }] } }),
    'protected-target'
  );
  refused(
    input('A1:A5', undefined, { tables: [{ name: 'Table1', sheet: 'Data', ref: 'A4:B8', columns: [] }] }),
    'protected-target'
  );
  refused(input('A1:A5', undefined, { sheet: { filter: { range: 'A4:B8' } } }), 'protected-target');
  refused(input('A1:A5', undefined, { protectionReason: 'Input-rule inventory is protected.' }), 'protected-target');
  assert(
    ready(
      input('A1:A5', undefined, {
        sheet: { merges: ['C1:D3'] },
        tables: [{ name: 'Other', sheet: 'Other', ref: 'A1:A5', columns: [] }],
      })
    )
  );
});

test('captured owner, canonical range, complete read and malformed inventory refusals are fallible and specific', () => {
  const valid = input();
  refused({ ...valid, binding: { ...valid.binding, ownerIdentity: null } }, 'retired-target');
  refused({ ...valid, binding: { ...valid.binding, sheet: 'Other' } }, 'retired-target');
  refused({ ...valid, binding: { ...valid.binding, specialSelection: ['A1'] } }, 'unsupported-selection');
  refused({ ...valid, binding: { ...valid.binding, mode: 'pageLayout' } }, 'unsupported-selection');
  refused({ ...valid, readRange: 'A1:A2' }, 'read-incomplete');
  refused({ ...valid, cells: { ...valid.cells, Z99: number(1) } }, 'read-incomplete');
  refused({ ...valid, cells: { ...valid.cells, A3: { bold: true } } }, 'read-incomplete');
  refused({ ...valid, sheet: { ...valid.sheet, merges: undefined } }, 'protected-target');
  refused({ ...valid, tables: [{ sheet: 'Data' }] }, 'invalid-input');
  refused({ ...valid, foreignPath: '/private/foreign' }, 'invalid-input');
  for (const range of ['A1:A2', 'A1:B3', 'A6:A1', 'XFE1:XFE3', 'A1048575:A1048577', 'a1:a4', '$A$1:$A$4']) {
    const result = createFillSeriesPreview(input(range));
    assert.equal(result.status, 'refused');
    assert.equal(result.code === 'unsupported-selection' || result.code === 'invalid-input', true);
  }
});

test('10,000-position paging visits each destination once, respects bounds and never exposes mutable arrays', () => {
  assert.deepEqual(FILL_SERIES_LIMITS, { positions: 10000, page: 100 });
  const preview = ready(input('A1:A10000'));
  let offset = 0;
  const rows = [];
  do {
    const page = fillSeriesPreviewPage(preview, offset);
    assert(page.rows.length <= 100);
    assert.equal(page.offset, offset);
    assert.equal(page.total, 9998);
    rows.push(...page.rows);
    offset = page.nextOffset;
  } while (offset !== null);
  assert.equal(rows.length, 9998);
  assert.equal(new Set(rows.map((row) => row.cell)).size, 9998);
  assert.deepEqual(rows[0], { cell: 'A3', value: 5 });
  assert.deepEqual(rows.at(-1), { cell: 'A10000', value: 19999 });
  assert.deepEqual(fillSeriesPreviewPage(preview, 9998), { offset: 9998, total: 9998, rows: [], nextOffset: null });
  for (const [offset, limit] of [
    [-1, 100],
    [9999, 100],
    [0, 0],
    [0, 101],
    [1.5, 10],
    [0, 1.5],
  ])
    assert.throws(() => fillSeriesPreviewPage(preview, offset, limit), /Invalid/);
  refused(input('A1:A10001'), 'unsupported-selection');
  const edge = ready(input('XFD1048574:XFD1048576', { XFD1048574: number(1), XFD1048575: number(3) }));
  assert.deepEqual(edge.destinations, [{ cell: 'XFD1048576', value: 5 }]);
  for (const object of [
    preview,
    preview.binding,
    preview.seeds,
    preview.seeds[0],
    preview.destinations,
    preview.destinations[0],
    preview.command,
    fillSeriesPreviewPage(preview),
    fillSeriesPreviewPage(preview).rows,
  ])
    assert(Object.isFrozen(object));
  assert.throws(() => preview.destinations.push({ cell: 'A10001', value: 1 }), TypeError);
  assert.throws(() => {
    preview.command.range = 'B1:B3';
  }, TypeError);
});

test('seed difference and late generated overflow refuse the complete preview without partial outputs', () => {
  refused(input('A1:A3', { A1: number(-1e308), A2: number(1e308) }), 'number-overflow');
  const source = input('A1:A4', { A1: number(1e308), A2: number(1.3e308) }),
    before = structuredClone(source);
  const result = refused(source, 'number-overflow');
  assert(!Object.hasOwn(result, 'preview'));
  assert.deepEqual(source, before);
});
