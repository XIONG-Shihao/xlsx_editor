import { test } from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { SheetHistory } from '../ui/sheetHistory.ts';
const component = fileURLToPath(new URL('../', import.meta.url));
const sheet = (value) => ({
  name: 'Sheet1',
  path: 'xl/worksheets/sheet1.xml',
  state: 'visible',
  rows: 2,
  columns: 2,
  cells: { A1: { value }, B1: { formula: 'A1+1', calculation: { status: 'calculated', value: value + 1 } } },
  merges: [],
  widths: [],
});
const view = (state, value, policy = { mode: 'auto', recalculateBeforeSave: true, pendingFormulaCells: 0 }) => ({
  revision: 'revision-' + state,
  sheets: [sheet(value)],
  storage: { kind: 'rust-indexed-v1', sessionId: 'owner-A', state },
  calculationPolicy: { schema: 1, canCalculateSheet: true, canChangeOptions: true, ...policy },
});
const check = (name, fn) =>
  test(name, async () => {
    await fn();
  });
check(
  'options/full/sheet calculation use owner-supplied accepted sheets/policy and keep exact commands through Undo/Redo',
  () => {
    const initial = view('base', 1),
      h = new SheetHistory(initial.sheets, [], false, initial);
    h.edit({ sheet: 'Sheet1', cell: 'A1', value: 2 });
    h.markSaved();
    const before = h.current;
    const manual = view('manual', 7, { mode: 'manual', recalculateBeforeSave: false, pendingFormulaCells: 1 }),
      command = { type: 'setCalculationOptions', mode: 'manual', recalculateBeforeSave: false };
    h.batch([command], undefined, manual);
    assert.equal(h.current.sheets, manual.sheets);
    assert.equal(h.current.view, manual);
    assert.deepEqual(h.current.edits.at(-1), command);
    assert.equal(h.current.focus, before.focus);
    assert(h.dirty);
    assert.equal(h.undo().view, before.view);
    assert.equal(h.current.sheets, before.sheets);
    assert.equal(h.dirty, false);
    assert.equal(h.redo().view, manual);
    assert.deepEqual(h.current.edits.at(-1), command);
    const calculated = view('calculated', 9, { mode: 'manual', recalculateBeforeSave: false, pendingFormulaCells: 0 });
    h.batch([{ type: 'recalculate' }], undefined, calculated);
    assert.equal(h.current.view, calculated);
    assert.equal(h.current.sheets[0].cells.B1.calculation.value, 10);
    assert.deepEqual(h.current.edits.at(-1), { type: 'recalculate' });
    h.undo();
    assert.equal(h.current.view, manual);
    h.redo();
    assert.equal(h.current.view, calculated);
    const selected = view('sheet', 10);
    h.batch([{ type: 'calculateSheet', sheet: 'Sheet1' }], undefined, selected);
    assert.equal(h.current.sheets, selected.sheets);
    assert.deepEqual(h.current.edits.at(-1), { type: 'calculateSheet', sheet: 'Sheet1' });
    assert.deepEqual(h.current.focus, before.focus);
  }
);
check(
  'metadata-only history refuses without an accepted view before changing Undo/Redo/journal; sheets alone never authorize it',
  () => {
    for (const command of [
      { type: 'recalculate' },
      { type: 'calculateSheet', sheet: 'Sheet1' },
      { type: 'setCalculationOptions', mode: 'manual', recalculateBeforeSave: false },
    ]) {
      const h = new SheetHistory([sheet(1)]);
      h.edit({ sheet: 'Sheet1', cell: 'A1', value: 2 });
      h.undo();
      const current = h.current,
        edits = h.current.edits,
        redo = h.canRedo;
      assert.throws(() => h.batch([command]), /accepted workbook view/);
      assert.throws(() => h.batch([command], [sheet(99)]), /accepted workbook view/);
      assert.equal(h.current, current);
      assert.deepEqual(h.current.edits, edits);
      assert.equal(h.canRedo, redo);
      assert.throws(() => new SheetHistory([sheet(1)], [command], true), /accepted workbook view/);
    }
  }
);
check(
  'calculation barriers preserve surrounding repeated cell values and false options across checkpoint compaction/replay',
  () => {
    const h = new SheetHistory([sheet(1)]),
      commands = [];
    for (let n = 0; n < 210; n++) {
      const point = { sheet: 'Sheet1', cell: 'A1', value: n };
      h.edit(point);
      if (n % 15 === 0) {
        const command =
          n % 30 === 0
            ? { type: 'setCalculationOptions', mode: 'manual', recalculateBeforeSave: false }
            : { type: 'recalculate' };
        commands.push(command);
        h.batch([command], undefined, view('v-' + n, n));
      }
    }
    const retained = h.current.edits.filter((e) => ['setCalculationOptions', 'recalculate'].includes(e.type));
    assert.deepEqual(retained, commands);
    const first = h.current.edits.findIndex((e) => e.type === 'setCalculationOptions');
    assert.equal(h.current.edits[first - 1].value, 0);
    assert.equal(h.current.edits[first + 1].value, 15);
    assert.equal(h.current.edits[first].recalculateBeforeSave, false);
    for (let n = 0; n < 30; n++) h.undo();
    for (let n = 0; n < 30; n++) h.redo();
    assert.deepEqual(
      h.current.edits.filter((e) => ['setCalculationOptions', 'recalculate'].includes(e.type)),
      commands
    );
  }
);
check(
  'accepted recovery metadata is retained without browser calculation and caller mutation cannot rewrite a journal command',
  () => {
    const command = { type: 'setCalculationOptions', mode: 'manual', recalculateBeforeSave: false },
      accepted = view('recovered', 12),
      h = new SheetHistory([sheet(0)], [command], true, accepted);
    command.mode = 'auto';
    command.recalculateBeforeSave = true;
    assert.deepEqual(h.current.edits, [
      { type: 'setCalculationOptions', mode: 'manual', recalculateBeforeSave: false },
    ]);
    assert.equal(h.current.view, accepted);
    assert.equal(h.current.sheets, accepted.sheets);
    assert.equal(h.canUndo, false);
    assert(h.dirty);
    const full = { type: 'recalculate' };
    h.batch([full], undefined, view('calculated', 13));
    full.force = true;
    assert.deepEqual(h.current.edits.at(-1), { type: 'recalculate' });
    assert.throws(() => h.batch([{ type: 'recalculate', force: true }], undefined, accepted), /Invalid calculation/);
  }
);
check(
  'calculation accepted views are never silently replaced by contradictory supplied sheet arrays and remain document isolated',
  () => {
    const accepted = view('A-new', 6),
      a = new SheetHistory([sheet(1)]),
      b = new SheetHistory([sheet(2)]);
    a.batch([{ type: 'calculateSheet', sheet: 'Sheet1' }], [sheet(999)], accepted);
    assert.equal(a.current.sheets, accepted.sheets);
    assert.equal(b.current.sheets[0].cells.A1.value, 2);
    assert.equal(b.current.edits.length, 0);
    assert.deepEqual(a.retainedNativeStates('owner-A'), ['A-new']);
    assert.deepEqual(b.retainedNativeStates('owner-A'), []);
  }
);
