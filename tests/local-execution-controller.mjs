import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createLocalExecutionController } from '../ui/localExecutionController.ts';
function harness() {
  const c = createLocalExecutionController(),
    writes = [];
  let frame = {
    key: 'owner/head/sheet/selection',
    ownerKey: 'owner/sheet/selection/activation',
    sheet: 'Sheet1',
    selection: 'B3:C7',
    active: true,
    apply: async (commands) => writes.push(commands),
  };
  c.setFrame(frame);
  return {
    c,
    writes,
    get frame() {
      return frame;
    },
    set(p) {
      frame = { ...frame, ...p };
      c.setFrame(frame);
    },
  };
}
test('Preview captures scope without writing, and Run submits its original single command once', async () => {
  const h = harness();
  assert(h.c.preview('freeze-selection'));
  assert.equal(h.writes.length, 0);
  assert.deepEqual(h.c.getSnapshot().preview.commands, [{ operation: 'freeze', sheet: 'Sheet1', rows: 2, columns: 1 }]);
  h.c.getSnapshot().preview.commands[0].sheet = 'forged';
  assert.deepEqual(await h.c.run(), { accepted: true, current: true });
  assert.equal(h.writes.length, 1);
  assert.equal(h.writes[0][0].sheet, 'Sheet1');
  assert.equal(h.c.getSnapshot().phase, 'completed');
  assert.equal((await h.c.run()).accepted, false);
});
test('Cancel discards a prepared action without a command or automatic reprepare', async () => {
  const h = harness();
  h.c.preview('unhide-all');
  assert(h.c.cancel());
  assert.equal((await h.c.run()).accepted, false);
  assert.equal(h.writes.length, 0);
  h.set({ key: 'new-head' });
  assert.equal(h.c.getSnapshot().preview, null);
});
test('selection/head/owner/activation/inactive and readonly changes retire previews before Run', async () => {
  for (const patch of [
    { key: 'new-head' },
    { key: 'new-selection', ownerKey: 'new-selection' },
    { key: 'new-owner', ownerKey: 'new-owner' },
    { ownerKey: 'new-activation' },
    { active: false },
    { reason: 'Read-only' },
  ]) {
    const h = harness();
    h.c.preview('unhide-all');
    h.set(patch);
    assert.equal(h.c.getSnapshot().preview, null);
    assert.equal((await h.c.run()).accepted, false);
    assert.equal(h.writes.length, 0);
  }
});
test('one physical write remains reserved across replacement and masks old success', async () => {
  const h = harness();
  let done;
  h.set({
    apply: async (commands) => {
      h.writes.push(commands);
      await new Promise((r) => (done = r));
    },
  });
  h.c.preview('unhide-all');
  const pending = h.c.run();
  assert(h.c.getSnapshot().pending);
  assert.equal((await h.c.run()).accepted, false);
  h.set({ key: 'replacement', ownerKey: 'replacement' });
  assert.equal(h.c.preview('freeze-selection'), false);
  done();
  assert.deepEqual(await pending, { accepted: true, current: false });
  assert.equal(h.writes.length, 1);
  assert.equal(h.c.getSnapshot().phase, 'idle');
  assert(!h.c.getSnapshot().pending);
});
test('expected accepted-head advancement preserves acknowledgment in the same owner/target', async () => {
  const h = harness();
  h.set({
    apply: async (commands) => {
      h.writes.push(commands);
      h.set({ key: 'accepted-head' });
    },
  });
  h.c.preview('unhide-all');
  assert.deepEqual(await h.c.run(), { accepted: true, current: true });
  assert.match(h.c.getSnapshot().message, /completed.*Sheet1.*Undo/);
  assert.equal(h.c.getSnapshot().key, 'accepted-head');
});
test('failure in unchanged scope is visible despite transient host busy; stale-head errors stay masked', async () => {
  const h = harness();
  h.set({
    apply: async () => {
      h.set({ reason: 'Wait for operation' });
      throw Error('Source has filters');
    },
  });
  h.c.preview('unhide-all');
  assert.equal((await h.c.run()).accepted, false);
  assert.equal(h.c.getSnapshot().message, 'Source has filters');
  h.set({
    reason: undefined,
    apply: async () => {
      h.set({ key: 'other-head' });
      throw Error('Old head failure');
    },
  });
  h.c.preview('unhide-all');
  await h.c.run();
  assert(!h.c.getSnapshot().message.includes('Old head failure'));
});
test('invalid A1/limits and readonly/busy scopes do not produce a runnable preview', () => {
  for (const patch of [
    { selection: 'A1' },
    { selection: 'AF102' },
    { reason: 'Read-only' },
    { reason: 'Wait for operation' },
    { active: false },
  ]) {
    const h = harness();
    h.set(patch);
    assert.equal(h.c.preview('freeze-selection'), false);
    assert.equal(h.c.getSnapshot().preview, null);
    assert.equal(h.writes.length, 0);
  }
});
test('hidden/disposed scopes do not auto-execute and cannot revive an old prepared target', async () => {
  const h = harness();
  h.c.preview('freeze-selection');
  h.set({ active: false });
  h.set({ active: true });
  assert.equal((await h.c.run()).accepted, false);
  h.c.preview('freeze-selection');
  h.c.dispose();
  assert.equal((await h.c.run()).accepted, false);
  assert.equal(h.writes.length, 0);
});
