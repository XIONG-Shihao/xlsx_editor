// The small pure parts of the spreadsheet screen that live under ui/grid.
import test from 'node:test';
import assert from 'node:assert/strict';
import { rectangle, selectedBounds, rangesOverlap, inRange } from '../ui/grid/rangeAddress.ts';
import { mergeAnchor, nextNavigableCell } from '../ui/grid/cellNavigation.ts';
import { editableNumberFormat, editableCellFormat } from '../ui/grid/editableFormats.ts';
import { conditionalFill, tableAppearance } from '../ui/grid/cellAppearance.ts';
import { acceptGridCommandBatch, normalizeGridCommandBatch } from '../ui/grid/commandBatch.ts';

const sheet = (patch = {}) => ({ name: 'Sheet1', cells: {}, merges: [], ...patch });

test('rectangle orders corners and collapses a single cell', () => {
  assert.equal(rectangle('C3', 'A1'), 'A1:C3');
  assert.equal(rectangle('B2', 'B2'), 'B2');
});

test('selectedBounds measures a range and refuses what it cannot hold', () => {
  const bounds = selectedBounds('B2:D5', 100);
  assert.deepEqual([bounds.rows, bounds.columns], [4, 3]);
  assert.throws(() => selectedBounds('A1:B2:C3', 100), /one rectangular range/);
  assert.throws(() => selectedBounds('A1:Z100', 100), /exceeds the 100-cell limit/);
  assert.throws(() => selectedBounds('A1:XFE1', 1e9), /outside the worksheet limits/);
});

test('rangesOverlap and inRange agree on edges', () => {
  assert.equal(rangesOverlap('A1:B2', 'B2:C3'), true);
  assert.equal(rangesOverlap('A1:B2', 'C3:D4'), false);
  assert.equal(inRange('A1:B2', 'B2'), true);
  assert.equal(inRange('A1:B2', 'C2'), false);
  assert.equal(inRange('not a range', 'A1'), false);
});

test('mergeAnchor returns the top-left cell of a merge and leaves other cells alone', () => {
  const merged = sheet({ merges: ['B2:C3'] });
  assert.equal(mergeAnchor(merged, 'C3'), 'B2');
  assert.equal(mergeAnchor(merged, 'D3'), 'D3');
  assert.equal(mergeAnchor(undefined, 'A1'), 'A1');
});

test('nextNavigableCell steps over merges and stops at the sheet edge', () => {
  const merged = sheet({ merges: ['B1:C1'] });
  assert.deepEqual(nextNavigableCell(merged, { row: 0, col: 0 }, 0, 1), { row: 0, col: 1, ref: 'B1' });
  assert.deepEqual(nextNavigableCell(merged, { row: 0, col: 1 }, 0, 1), { row: 0, col: 3, ref: 'D1' });
  assert.deepEqual(nextNavigableCell(merged, { row: 0, col: 0 }, -1, 0), { row: 0, col: 0, ref: 'A1' });
});

test('editable number formats map the supported built-ins and refuse the rest', () => {
  assert.equal(editableNumberFormat(undefined), 'general');
  assert.equal(editableNumberFormat(10), 'percent');
  assert.equal(editableNumberFormat('$#,##0'), 'currency');
  assert.equal(editableNumberFormat(999), undefined);
  assert.equal(editableCellFormat({ bold: true, format: 14 }).numberFormat, 'date');
  assert.equal(editableCellFormat(undefined).styleId, 0);
});

test('table banding and conditional fills read the sheet model', () => {
  const table = sheet({ tableRanges: [{ ref: 'A1:B3' }] });
  assert.equal(tableAppearance(table, 'A1').header, true);
  assert.equal(tableAppearance(table, 'A2').background, '#FFFFFF');
  assert.equal(tableAppearance(table, 'A3').background, '#D9E5F3');
  assert.equal(tableAppearance(table, 'C1'), undefined);
  const ruled = sheet({
    cells: { A1: { value: 5 }, A2: { value: 1 }, A3: { value: 5 } },
    conditionalRules: [{ range: 'A1:A3', rule: 'greaterThan', threshold: 3, fillColor: 'FFF2CC' }],
  });
  assert.equal(conditionalFill(ruled, 'A1'), 'FFF2CC');
  assert.equal(conditionalFill(ruled, 'A2'), undefined);
  const dup = sheet({
    cells: { A1: { value: 'x' }, A2: { value: 'y' }, A3: { value: 'x' } },
    conditionalRules: [{ range: 'A1:A3', rule: 'duplicateValues', fillColor: 'C6EFCE' }],
  });
  assert.equal(conditionalFill(dup, 'A1'), 'C6EFCE');
  assert.equal(conditionalFill(dup, 'A2'), undefined);
});

test('command batches split large cell writes and refuse unsafe states in a fixed order', async () => {
  const cells = Array.from({ length: 2500 }, (_, index) => ({ ref: 'A' + (index + 1), value: index }));
  const parts = normalizeGridCommandBatch([{ type: 'setCells', sheet: 'Sheet1', cells }, { type: 'recalculate' }]);
  assert.deepEqual(
    parts.map((part) => (part.type === 'setCells' ? part.cells.length : part.type)),
    [1000, 1000, 500, 'recalculate']
  );
  const ready = { active: true, readOnly: false, busy: false, specialSelection: false, onCommandBatch: async () => {} };
  await assert.rejects(acceptGridCommandBatch([], { ...ready, active: false }), /Activate this workbook/);
  await assert.rejects(acceptGridCommandBatch([], { ...ready, readOnly: true }), /read-only/);
  await assert.rejects(acceptGridCommandBatch([], { ...ready, busy: true }), /Wait for the workbook/);
  await assert.rejects(acceptGridCommandBatch([], { ...ready, specialSelection: true }), /Special selection/);
  await acceptGridCommandBatch([], { ...ready, specialSelection: true }, true);
  await assert.rejects(acceptGridCommandBatch([], { ...ready, onCommandBatch: undefined }), /unavailable/);
  const seen = [];
  await acceptGridCommandBatch([{ type: 'recalculate' }], {
    ...ready,
    onCommandBatch: async (batch) => seen.push(batch),
  });
  assert.deepEqual(seen, [[{ type: 'recalculate' }]]);
});
