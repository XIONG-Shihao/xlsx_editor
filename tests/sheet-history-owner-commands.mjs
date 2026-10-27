import { test } from 'node:test';
import assert from 'node:assert/strict';
import { SheetHistory } from '../ui/sheetHistory.ts';
const initial = {
  name: 'S',
  path: '',
  state: 'visible',
  rows: 3,
  columns: 2,
  cells: { A1: { value: 'Key' }, A2: { value: 'a' }, A3: { value: 'a' } },
  merges: [],
  widths: [],
};
const view = (sheet, state) => ({
  sheets: [sheet],
  revision: state,
  storage: { kind: 'rust-indexed-v1', sessionId: 'owner', state },
});
test('duplicate/filter history refuses browser inference and retains exact accepted views through Undo/Redo', () => {
  const duplicate = {
      operation: 'removeDuplicates',
      sheet: 'S',
      range: 'A1:B3',
      hasHeaders: true,
      columns: ['A'],
      expectedFingerprint: 'a'.repeat(64),
    },
    remove = { operation: 'removeFilter', sheet: 'S', expectedFilterRevision: 'b'.repeat(64) },
    history = new SheetHistory([initial], [], false, view(initial, 'base'));
  for (const command of [duplicate, remove]) {
    assert.throws(() => history.batch([command]), /requires its accepted workbook view/);
    assert.throws(() => history.batch([command], [initial]), /requires its accepted workbook view/);
    assert.throws(() => new SheetHistory([initial], [command], true), /requires its accepted workbook view/);
  }
  assert.equal(history.canUndo, false);
  assert.equal(history.dirty, false);
  const accepted = { ...initial, cells: { A1: { value: 'Key' }, A2: { value: 'a' }, A3: { value: null } } },
    acceptedView = view(accepted, 'after-duplicate');
  history.batch([duplicate], undefined, acceptedView);
  assert.equal(history.current.sheets[0], accepted);
  assert.equal(history.current.view, acceptedView);
  assert.deepEqual(history.current.edits, [duplicate]);
  duplicate.columns.push('B');
  assert.deepEqual(history.current.edits[0].columns, ['A']);
  const noFilter = { ...accepted, hiddenRows: [] },
    filterView = view(noFilter, 'after-remove');
  history.batch([remove], undefined, filterView);
  assert.equal(history.current.view, filterView);
  assert.equal(history.current.sheets[0], noFilter);
  assert.equal(history.current.edits.length, 2);
  assert.equal(history.undo().view, acceptedView);
  assert.equal(history.undo().view.revision, 'base');
  assert.equal(history.redo().view, acceptedView);
  assert.equal(history.redo().view, filterView);
  assert.deepEqual(history.current.edits[1], remove);
});
