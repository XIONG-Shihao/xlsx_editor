import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createInkGestureController } from '../ui/inkGestureController.ts';
import { createObjectAxes } from '../ui/worksheetObjectGeometry.ts';
const axes = createObjectAxes({ sheetFormat: { defaultColumnWidth: 10, defaultRowHeight: 15 } });
function harness() {
  const writes = [],
    c = createInkGestureController();
  let f = {
    key: 'doc/owner/head/sheet/activation/viewport',
    sheet: 'Sheet1',
    active: true,
    pen: 'black',
    axes,
    view: { scale: 1, rowHeaderWidth: 40, columnHeaderHeight: 24, scrollLeft: 0, scrollTop: 0 },
    accept: async (commands) => {
      writes.push(commands);
    },
  };
  c.setFrame(f);
  return {
    c,
    writes,
    get f() {
      return f;
    },
    set(p) {
      f = { ...f, ...p };
      c.setFrame(f);
    },
  };
}
const stroke = (h) => {
  assert(h.c.start(1, 140, 124));
  assert(h.c.move(1, 180, 144));
  return h.c.finish(1, 240, 164);
};
test('actual source axes produce one black vector command on release; draft does not write', async () => {
  const h = harness();
  assert(h.c.start(1, 140, 124));
  h.c.move(1, 180, 144);
  assert.equal(h.writes.length, 0);
  assert.equal(h.c.getSnapshot().points.length, 2);
  assert.deepEqual(await h.c.finish(1, 240, 164), { accepted: true, current: true });
  assert.equal(h.writes.length, 1);
  const c = h.writes[0][0];
  assert.equal(c.operation, 'createInkStroke');
  assert.match(c.objectId, /^[a-f0-9-]{36}$/);
  assert.equal(c.expectedObjectRevision, null);
  assert.equal(c.sheet, 'Sheet1');
  assert.deepEqual(c.stroke.points, [
    [9525, 9525],
    [390525, 200025],
    [962025, 390525],
  ]);
  assert.deepEqual(c.anchor, {
    kind: 'oneCell',
    cell: 'B5',
    dxEmu: 276225,
    dyEmu: 180975,
    widthEmu: 971550,
    heightEmu: 400050,
  });
  assert.equal(h.c.getSnapshot().points.length, 0);
});
test('ordinary zoom/scroll converts once to the same model and placement', async () => {
  const a = harness(),
    b = harness();
  await stroke(a);
  b.set({ view: { scale: 1.25, rowHeaderWidth: 40, columnHeaderHeight: 24, scrollLeft: 50, scrollTop: 40 } });
  assert(b.c.start(1, 115, 109));
  b.c.move(1, 165, 134);
  await b.c.finish(1, 240, 159);
  assert.deepEqual(b.writes[0][0].stroke, a.writes[0][0].stroke);
  assert.deepEqual(b.writes[0][0].anchor, a.writes[0][0].anchor);
});
test('cancel and mismatched pointer leave workbook unchanged', async () => {
  const h = harness();
  h.c.start(1, 140, 124);
  assert.equal(h.c.move(2, 180, 144), false);
  assert.equal((await h.c.finish(2, 240, 164)).accepted, false);
  h.c.cancel();
  assert.equal((await h.c.finish(1, 240, 164)).accepted, false);
  assert.equal(h.writes.length, 0);
});
test('owner/head/viewport changes and inactive/readonly states retire the draft before any write', async () => {
  for (const patch of [
    { key: 'new owner/head/activation' },
    { key: 'zoom-scroll-epoch' },
    { active: false },
    { reason: 'Read-only' },
  ]) {
    const h = harness();
    h.c.start(1, 140, 124);
    h.c.move(1, 180, 144);
    h.set(patch);
    assert.equal(h.c.getSnapshot().points.length, 0);
    assert.equal((await h.c.finish(1, 240, 164)).accepted, false);
    assert.equal(h.writes.length, 0);
  }
});
test('one physical acceptance lane remains reserved across head/owner retirement and duplicate releases', async () => {
  const h = harness();
  let done;
  const pending = new Promise((resolve) => (done = resolve));
  h.set({
    accept: async (commands) => {
      h.writes.push(commands);
      await pending;
    },
  });
  h.c.start(1, 140, 124);
  const flight = h.c.finish(1, 240, 164);
  assert(h.c.getSnapshot().pending);
  assert.equal(h.c.start(2, 140, 124), false);
  assert.equal((await h.c.finish(1, 240, 164)).accepted, false);
  h.set({ key: 'replacement' });
  assert.equal(h.c.start(2, 140, 124), false);
  done();
  assert.deepEqual(await flight, { accepted: true, current: false });
  assert.equal(h.writes.length, 1);
  assert(!h.c.getSnapshot().pending);
  assert(h.c.start(2, 140, 124));
});
test('tap, worksheet edge, frozen pane and point budget are refused without acceptance', async () => {
  for (const mode of ['tap', 'edge', 'frozen', 'budget']) {
    const h = harness();
    if (mode === 'frozen') h.set({ view: { ...h.f.view, frozenRows: 1 } });
    if (mode === 'edge') {
      assert.equal(h.c.start(1, 39, 124), false);
    } else if (mode === 'frozen') {
      assert.equal(h.c.start(1, 140, 124), false);
    } else {
      h.c.start(1, 140, 124);
      if (mode === 'budget') for (let i = 1; i <= 512; i++) h.c.move(1, 140 + i, 124 + i);
      await h.c.finish(1, 140, 124);
    }
    assert.equal(h.writes.length, 0);
    assert(h.c.getSnapshot().error);
  }
});
test('acceptance failure is visible only to the current frame and releases the lane for a deliberate retry', async () => {
  const h = harness();
  h.set({
    accept: async () => {
      throw Error('Native refusal');
    },
  });
  assert.equal((await stroke(h)).accepted, false);
  assert.equal(h.c.getSnapshot().error, 'Native refusal');
  assert(!h.c.getSnapshot().pending);
  h.set({ accept: async (commands) => h.writes.push(commands) });
  assert((await stroke(h)).accepted);
  assert.equal(h.writes.length, 1);
  assert.equal(h.c.getSnapshot().error, null);
});
test('retired failure never publishes into a replacement owner and disposal cannot accept', async () => {
  const h = harness();
  let fail;
  h.set({ accept: () => new Promise((_, reject) => (fail = reject)) });
  h.c.start(1, 140, 124);
  const flight = h.c.finish(1, 240, 164);
  h.set({ key: 'replacement' });
  fail(Error('Old refusal'));
  await flight;
  assert.equal(h.c.getSnapshot().error, null);
  assert.equal(h.writes.length, 0);
  h.c.dispose();
  assert.equal(h.c.start(1, 140, 124), false);
});
