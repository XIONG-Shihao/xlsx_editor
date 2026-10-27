import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createInkPreferenceStore, INK_PREFERENCE_KEY } from '../ui/inkPreferences.ts';
import { dispatchRibbonCommand } from '../ui/ribbonCommands.ts';
import { createInkGestureController } from '../ui/inkGestureController.ts';
import { createObjectAxes } from '../ui/worksheetObjectGeometry.ts';
function storage() {
  const values = new Map(),
    writes = [];
  return {
    values,
    writes,
    getItem: (k) => values.get(k) ?? null,
    setItem: (k, v) => {
      writes.push([k, v]);
      values.set(k, v);
    },
  };
}
test('all three choices persist as one small private value and restore cold without workbook commands', () => {
  const s = storage(),
    store = createInkPreferenceStore(s);
  assert.equal(store.read(), 'black');
  assert.equal(s.writes.length, 0);
  for (const pen of ['black', 'red', 'yellow-highlighter']) {
    assert.equal(store.write(pen), pen);
    assert.deepEqual(JSON.parse(s.values.get(INK_PREFERENCE_KEY)), { schema: 1, pen });
    assert.equal(createInkPreferenceStore(s).read(), pen);
  }
  assert.equal(s.values.size, 1);
  assert(s.values.get(INK_PREFERENCE_KEY).length < 256);
  assert.equal(store.warning, null);
});
test('missing/deleted private choice defaults to Black without implicit writes', () => {
  const s = storage(),
    store = createInkPreferenceStore(s);
  store.write('red');
  s.values.delete(INK_PREFERENCE_KEY);
  assert.equal(store.read(), 'black');
  assert.equal(s.writes.length, 1);
});
test('invalid/versioned/oversized stored data is preserved with a warning until deliberate replacement', () => {
  for (const raw of [
    'bad JSON',
    'null',
    '[]',
    '"red"',
    '{"schema":2,"pen":"red"}',
    '{"schema":1,"pen":"foreign"}',
    '{"schema":1,"pen":"red","path":"private"}',
    'x'.repeat(257),
  ]) {
    const s = storage();
    s.values.set(INK_PREFERENCE_KEY, raw);
    const store = createInkPreferenceStore(s);
    assert.equal(store.read(), 'black');
    assert(store.warning);
    assert.equal(s.values.get(INK_PREFERENCE_KEY), raw);
    assert.equal(s.writes.length, 0);
    store.write('yellow-highlighter');
    assert.equal(store.warning, null);
    assert.equal(createInkPreferenceStore(s).read(), 'yellow-highlighter');
  }
});
test('quota, absent storage and denied reads keep a session choice usable without pretending it persisted', () => {
  for (const s of [
    null,
    {
      getItem: () => null,
      setItem: () => {
        throw Error('quota');
      },
    },
    {
      getItem: () => {
        throw Error('denied');
      },
      setItem: () => {
        throw Error('denied');
      },
    },
  ]) {
    const store = createInkPreferenceStore(s);
    store.read();
    assert.equal(store.write('red'), 'red');
    assert.equal(store.read(), 'red');
    assert.match(store.warning, /session only/);
    assert.equal(store.write('yellow-highlighter'), 'yellow-highlighter');
    assert.equal(store.read(), 'yellow-highlighter');
  }
});
test('unknown pen input cannot mutate the previous preference or storage', () => {
  const s = storage(),
    store = createInkPreferenceStore(s);
  store.write('red');
  assert.throws(() => store.write('foreign'), { code: 'object-bad-input' });
  assert.equal(store.read(), 'red');
  assert.equal(s.writes.length, 1);
});
const binding = { documentId: 'doc', state: 'head', sheet: 'Sheet1', selection: 'A1', activationEpoch: 1 },
  scope = {
    captureBinding: () => binding,
    isCurrent: () => true,
    availability: () => ({ profile: 'xlsx', active: true, canEdit: true, busy: false }),
  };
test('ribbon preset arguments are detached/closed and inactive, stale, readonly and busy callbacks do not run', async () => {
  let calls = 0;
  const args = { pen: 'red' };
  const result = await dispatchRibbonCommand(
    'draw.preset',
    scope,
    (_b, a) => {
      calls++;
      assert(Object.isFrozen(a));
      assert.notEqual(a, args);
      assert.equal(a.pen, 'red');
    },
    args
  );
  assert.equal(result.outcome, 'completed');
  for (const invalid of [undefined, { pen: 'unknown' }, { pen: 'red', width: 100 }, { pen: null }])
    await assert.rejects(dispatchRibbonCommand('draw.preset', scope, () => calls++, invalid));
  for (const patch of [{ active: false }, { canEdit: false }, { busy: true }])
    assert.equal(
      (
        await dispatchRibbonCommand(
          'draw.preset',
          { ...scope, availability: () => ({ ...scope.availability(), ...patch }) },
          () => calls++,
          args
        )
      ).outcome,
      'refused'
    );
  assert.equal(
    (await dispatchRibbonCommand('draw.preset', { ...scope, isCurrent: () => false }, () => calls++, args)).outcome,
    'stale'
  );
  assert.equal(calls, 1);
});
test('changing the selected preset retires a captured draft; the next stroke uses the new pen', async () => {
  const c = createInkGestureController(),
    writes = [],
    axes = createObjectAxes({ sheetFormat: { defaultColumnWidth: 10, defaultRowHeight: 15 } }),
    f = {
      key: 'owner/head/black',
      active: true,
      sheet: 'Sheet1',
      pen: 'black',
      axes,
      view: { scale: 1, rowHeaderWidth: 48, columnHeaderHeight: 32, scrollLeft: 0, scrollTop: 0 },
      accept: async (a) => writes.push(a),
    };
  c.setFrame(f);
  c.start(1, 140, 124);
  c.move(1, 180, 144);
  c.setFrame({ ...f, key: 'owner/head/red', pen: 'red' });
  assert.equal((await c.finish(1, 240, 164)).accepted, false);
  assert.equal(writes.length, 0);
  c.start(1, 140, 124);
  await c.finish(1, 240, 164);
  assert.equal(writes[0][0].stroke.pen, 'red');
  c.dispose();
});
