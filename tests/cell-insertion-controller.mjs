import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createCellInsertionController } from '../ui/cellInsertionController.ts';
const frame = (patch) => ({
  key: 'A/head1/epoch1',
  family: 'A/cell/epoch1',
  sheet: 'Main',
  cell: 'D1',
  active: true,
  writable: true,
  read: async () => null,
  apply: async () => {},
  ...patch,
});
const tick = () => new Promise((r) => setTimeout(r, 0));
async function ready(c, f) {
  c.setFrame(f);
  await tick();
  assert(c.getSnapshot().loaded);
}
test('one shared physical read lane ignores retired cells and serves only the latest target', async () => {
  const c = createCellInsertionController();
  let release,
    active = 0,
    peak = 0;
  const old = frame({
    read: async () => {
      peak = Math.max(peak, ++active);
      await new Promise((r) => (release = r));
      active--;
      return { value: false, dataType: 'boolean' };
    },
  });
  c.setFrame(old);
  c.setFrame(
    frame({
      key: 'other',
      cell: 'A1',
      read: async () => {
        peak = Math.max(peak, ++active);
        active--;
        return { value: 'text', dataType: 'text' };
      },
    })
  );
  assert.equal(c.getSnapshot().loaded, false);
  release();
  await tick();
  await tick();
  assert.equal(peak, 1);
  assert.equal(c.getSnapshot().cell.value, 'text');
  c.dispose();
});
test('Boolean insert/toggle/remove captures the original immutable cell and one action', async () => {
  const c = createCellInsertionController(),
    writes = [];
  let release;
  const source = { value: false, dataType: 'boolean' },
    old = frame({
      read: async () => source,
      apply: async (commands) => {
        writes.push(commands);
        await new Promise((r) => (release = r));
      },
    });
  await ready(c, old);
  source.value = true;
  const pending = c.boolean(old, 'toggle');
  await assert.rejects(c.boolean(old, 'toggle'), /Wait/);
  c.setFrame(frame({ key: 'head2', read: async () => ({ value: true, dataType: 'boolean' }) }));
  release();
  const result = await pending;
  assert.equal(result.accepted, true);
  assert.equal(result.current, false);
  assert.deepEqual(writes, [
    [{ type: 'booleanCell', action: 'toggle', sheet: 'Main', cell: 'D1', expectedValue: false }],
  ]);
  c.dispose();
});
test('pane/owner retirement discards late errors, stale callbacks and disabled writes', async () => {
  const c = createCellInsertionController();
  let reject;
  const old = frame({ apply: () => new Promise((_, r) => (reject = r)) });
  await ready(c, old);
  const pending = c.boolean(old, 'insert');
  c.setFrame({ ...old, active: false });
  c.setFrame(frame({ key: 'new-pane', family: 'new-pane' }));
  reject(Error('retired failure'));
  await assert.rejects(pending);
  await tick();
  assert.equal(c.getSnapshot().error, '');
  await assert.rejects(c.boolean(old, 'insert'), /changed/);
  c.setFrame(frame({ key: 'new-pane', family: 'new-pane', writable: false, reason: 'Read-only' }));
  await assert.rejects(c.boolean(frame({ key: 'new-pane' }), 'insert'), /Read-only/);
  c.dispose();
});
test('Symbols capture plain-text inputs and append exact literal samples without formula conversion', async () => {
  const c = createCellInsertionController(),
    writes = [],
    source = { value: '=SUM(', dataType: 'text' };
  const f = frame({ read: async () => source, apply: async (commands) => writes.push(commands) });
  await ready(c, f);
  source.value = 'changed';
  await c.symbol(f, 'omega');
  assert.deepEqual(writes, [
    [{ sheet: 'Main', cell: 'D1', value: '=SUM(Ω', expectedValue: '=SUM(', expectedFormula: null }],
  ]);
  c.dispose();
});
test('formula, invalid Boolean values, malformed replies and missing reads never grant writes', async () => {
  const c = createCellInsertionController(),
    writes = [];
  let f = frame({ read: async () => ({ value: false, formula: 'TRUE()' }), apply: async (x) => writes.push(x) });
  await ready(c, f);
  await assert.rejects(c.boolean(f, 'toggle'), { code: 'workbook-bad-input' });
  await assert.rejects(c.symbol(f, 'degree'), { code: /^(internal-invalid-field|workbook-bad-input)$/ });
  f = frame({ key: 'invalid', read: async () => [] });
  c.setFrame(f);
  await tick();
  assert.equal(c.getSnapshot().loaded, false);
  await assert.rejects(c.boolean(f, 'insert'), /Load/);
  assert.deepEqual(writes, []);
  c.dispose();
});
