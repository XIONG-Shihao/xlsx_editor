import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  booleanCellCommand,
  normalizeBooleanCellCommands,
  symbolCellEdit,
  symbolSamples,
} from '../core/cell-insertions.mjs';
import { resolveWorkbookEdits } from '../core/workbook-commands.mjs';
const command = { type: 'booleanCell', action: 'insert', sheet: 'Main', cell: 'D1', expectedValue: null };
test('closed Boolean-cell intent captures values and rejects forged/coercible profiles', () => {
  assert.deepEqual(booleanCellCommand(command), command);
  for (const patch of [
    { value: false },
    { expectedValue: 'FALSE' },
    { cell: 'd1' },
    { cell: 'XFE1' },
    { cell: 'A1048577' },
    { cell: 'A1:B1' },
    { cell: 'A01' },
    { sheet: 'foreign/name' },
    { action: 'toggle' },
    { action: 'script' },
    { expectedValue: false },
  ])
    assert.throws(() => booleanCellCommand({ ...command, ...patch }));
  assert.equal(booleanCellCommand({ ...command, action: 'toggle', expectedValue: false }).action, 'toggle');
});
test('queued batch snapshot is independent of caller mutation including mixed literal edits', () => {
  const cell = { sheet: 'Main', cell: 'A1', value: 'original' },
    commands = [{ ...command }, cell],
    captured = normalizeBooleanCellCommands(commands);
  commands[0].cell = 'A1';
  cell.value = 'changed';
  assert.equal(captured[0].cell, 'D1');
  assert.equal(captured[1].value, 'original');
  assert.throws(
    () => normalizeBooleanCellCommands([{ ...command }, { sheet: 'Main', cell: 'A1', value: 'Ω'.repeat(1048576) }]),
    { code: /^(session-limit|workbook-limit)$/ }
  );
});
test('semantic checkbox refuses portable authoring instead of losing its profile gate', () => {
  assert.throws(() => resolveWorkbookEdits({ sheets: [] }, [command]), /native XLSX owner/);
});
test('three frozen Symbols samples produce literal Unicode appends with original-value guards', () => {
  assert.equal(symbolSamples.length, 3);
  assert(Object.isFrozen(symbolSamples) && symbolSamples.every(Object.isFrozen));
  let source = null;
  for (const s of symbolSamples) {
    const edit = symbolCellEdit('Main', 'A1', source, s.id);
    assert.equal(edit.expectedValue, source?.value ?? null);
    assert.equal(edit.expectedFormula, null);
    source = { value: edit.value, dataType: 'text' };
  }
  assert.equal(source.value, 'Ω±°');
  assert.equal(symbolCellEdit('Main', 'A1', { value: '=SUM(', dataType: 'text' }, 'omega').value, '=SUM(Ω');
});
test('Symbols refuse formula, derived, numeric and error cells and text overflow', () => {
  for (const source of [
    { value: 1 },
    { value: false },
    { value: null, formula: 'A1' },
    { value: 'x', formulaOutput: {} },
    { value: 'x', spill: {} },
    { value: 'x', calculation: {} },
    { value: '#N/A', dataType: 'error' },
    { value: 'x'.repeat(32767), dataType: 'text' },
  ])
    assert.throws(() => symbolCellEdit('Main', 'A1', source, 'degree'));
  assert.throws(() => symbolCellEdit('Main', 'A1', null, 'full-gallery'));
  assert.equal(symbolCellEdit('Main', 'A1', { value: '#N/A', dataType: 'text' }, 'degree').value, '#N/A°');
});
