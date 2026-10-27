import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createQueryRefreshAllController } from '../ui/queryRefreshAllController.ts';
const ids = [
  '11111111-1111-4111-8111-111111111111',
  '22222222-2222-4222-8222-222222222222',
  '33333333-3333-4333-8333-333333333333',
];
const scope = () => ({ documentId: 'a'.repeat(64), ownerId: ids[0], state: 'base' });
const plan = (s = scope()) => ({
  schema: 1,
  id: ids[2],
  scope: s,
  expiresAt: Date.now() + 60000,
  undoPolicy: 'One Undo step for each accepted connection',
  entries: ids.slice(0, 2).map((id, i) => ({
    connectionId: id,
    revision: 1,
    name: 'Connection ' + i,
    target: { sheet: 'S', start: i ? 'D1' : 'A1' },
    status: 'ready',
  })),
});
const result = (p) => ({
  schema: 1,
  previewId: p.id,
  scope: p.scope,
  lastAcknowledgedState: ids[1],
  cancelled: false,
  stopped: false,
  outcomes: p.entries.map((e, i) => ({
    ...e,
    status: 'refreshed',
    undoUnits: 1,
    parentState: i ? ids[0] : 'base',
    state: ids[i],
  })),
  accepted: 2,
  uncertain: 0,
  undoUnits: 2,
});
const defer = () => {
  let resolve, reject;
  const promise = new Promise((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
};
function harness(patch = {}) {
  const calls = [],
    c = createQueryRefreshAllController();
  let frame = {
    key: 'owner/base',
    ownerKey: 'owner/activation/dock',
    name: 'Budget.xlsx',
    active: true,
    controls: {
      scope: scope(),
      preview: async (requestId) => {
        calls.push(['preview', requestId]);
        return { refreshAllPreview: plan() };
      },
      refresh: async (id, requestId) => {
        calls.push(['refresh', id, requestId]);
        return { refreshAllResult: result(plan()) };
      },
      cancel: async (requestId) => {
        calls.push(['cancel', requestId]);
        return { outcome: 'cancelled' };
      },
    },
    ...patch,
  };
  c.setFrame(frame);
  return {
    c,
    calls,
    get frame() {
      return frame;
    },
    set(p) {
      frame = { ...frame, ...p };
      c.setFrame(frame);
    },
  };
}
test('explicit preview is read-only; private token survives public mutation and Run submits once', async () => {
  const h = harness();
  assert.equal(h.calls.length, 0);
  await h.c.preview();
  assert.equal(h.calls.filter((c) => c[0] === 'refresh').length, 0);
  h.c.getSnapshot().preview.id = ids[0];
  h.c.getSnapshot().preview.entries[0].target.start = 'Z9';
  assert(await h.c.run());
  assert.equal(h.calls[1][1], ids[2]);
  assert.equal(h.c.getSnapshot().result.accepted, 2);
  assert.equal(await h.c.run(), false);
});
test('own busy render and accepted head advancement keep acknowledgment without adopting a foreign result', async () => {
  const p = defer(),
    r = defer(),
    h = harness();
  h.set({ controls: { ...h.frame.controls, preview: () => p.promise, refresh: () => r.promise } });
  const first = h.c.preview();
  h.set({ busy: true });
  p.resolve({ refreshAllPreview: plan() });
  assert(await first);
  h.set({ busy: false });
  const run = h.c.run();
  h.set({ key: 'owner/accepted', controls: { ...h.frame.controls, scope: { ...scope(), state: ids[1] } } });
  r.resolve({ refreshAllResult: result(plan()) });
  assert(await run);
  assert.equal(h.c.getSnapshot().result.accepted, 2);
});
test('head/owner/mode/closing retirement invalidates prepared previews and never automatically reprepares', async () => {
  for (const patch of [
    { key: 'new-head' },
    { ownerKey: 'new-owner' },
    { ownerKey: 'new-mode' },
    { active: false },
    { reason: 'Read-only' },
  ]) {
    const h = harness();
    await h.c.preview();
    h.set(patch);
    assert.equal(h.c.getSnapshot().preview, null);
    assert.equal(await h.c.run(), false);
    assert.equal(h.calls.filter((c) => c[0] === 'refresh').length, 0);
  }
});
test('one physical request remains reserved across owner replacement and masks late success/error', async () => {
  for (const failed of [false, true]) {
    const gate = defer(),
      h = harness();
    h.set({ controls: { ...h.frame.controls, refresh: () => gate.promise } });
    await h.c.preview();
    const run = h.c.run();
    assert.equal(await h.c.run(), false);
    h.set({ ownerKey: 'replacement', key: 'replacement' });
    assert.equal(await h.c.preview(), false);
    failed ? gate.reject(Error('Old owner failure')) : gate.resolve({ refreshAllResult: result(plan()) });
    await run;
    assert.equal(h.c.getSnapshot().result, null);
    assert(!h.c.getSnapshot().message.includes('Old owner'));
    assert(!h.c.getSnapshot().pending);
  }
});
test('preview cancellation owns the original request ID and masks a late preview without writes', async () => {
  const gate = defer(),
    h = harness();
  h.set({
    controls: {
      ...h.frame.controls,
      preview: async (id) => {
        h.calls.push(['preview', id]);
        return gate.promise;
      },
    },
  });
  const pending = h.c.preview();
  await h.c.cancel();
  assert.equal(h.calls[1][1], h.calls[0][1]);
  gate.resolve({ refreshAllPreview: plan() });
  assert.equal(await pending, false);
  assert.equal(h.c.getSnapshot().preview, null);
  assert.equal(await h.c.run(), false);
});
test('Run cancellation preserves and displays the acknowledged prefix rather than discarding its result', async () => {
  const gate = defer(),
    h = harness();
  h.set({ controls: { ...h.frame.controls, refresh: () => gate.promise } });
  await h.c.preview();
  const pending = h.c.run();
  await h.c.cancel();
  const p = plan(),
    r = result(p);
  Object.assign(r, { cancelled: true, stopped: true, accepted: 1, undoUnits: 1, lastAcknowledgedState: ids[0] });
  r.outcomes[1] = { ...p.entries[1], status: 'cancelled', undoUnits: 0 };
  gate.resolve({ refreshAllResult: r });
  await pending;
  assert.equal(h.c.getSnapshot().result.accepted, 1);
  assert(h.c.getSnapshot().result.cancelled);
});
test('invalid/foreign preview summaries and inconsistent outcomes never enable a runnable result', async () => {
  for (const bad of [
    { ...plan(), scope: { ...scope(), documentId: 'b'.repeat(64) } },
    { ...plan(), entries: [{ ...plan().entries[0], sample: [['ok'], ['bad', 1]] }] },
    { ...plan(), expiresAt: 0 },
  ]) {
    const h = harness();
    h.set({ controls: { ...h.frame.controls, preview: async () => ({ refreshAllPreview: bad }) } });
    assert.equal(await h.c.preview(), false);
    assert.equal(h.c.getSnapshot().phase, 'failed');
    assert.equal(await h.c.run(), false);
  }
  const h = harness();
  h.set({
    controls: {
      ...h.frame.controls,
      refresh: async () => ({ refreshAllResult: { ...result(plan()), undoUnits: 99 } }),
    },
  });
  await h.c.preview();
  assert.equal(await h.c.run(), false);
  assert.equal(h.c.getSnapshot().result, null);
});
test('uncertain acceptance preserves unknown Undo count and never reports the uncertain row as refused', async () => {
  const h = harness(),
    p = plan();
  h.set({
    controls: {
      ...h.frame.controls,
      refresh: async () => ({
        refreshAllResult: {
          schema: 1,
          previewId: p.id,
          scope: p.scope,
          lastAcknowledgedState: 'base',
          cancelled: false,
          stopped: true,
          accepted: 0,
          uncertain: 1,
          undoUnits: null,
          outcomes: [
            { ...p.entries[0], status: 'uncertain', undoUnits: null, reason: 'Receipt pending.' },
            { ...p.entries[1], status: 'not-run', undoUnits: 0 },
          ],
        },
      }),
    },
  });
  await h.c.preview();
  assert(await h.c.run());
  assert.equal(h.c.getSnapshot().result.undoUnits, null);
  assert.equal(h.c.getSnapshot().result.uncertain, 1);
  assert.match(h.c.getSnapshot().message, /uncertain/);
});

test('a delayed Cancel reply releases its UI flag and blocks a successor until cancellation settles', async () => {
  const reply = defer(),
    cancelReply = defer(),
    h = harness();
  h.set({ controls: { ...h.frame.controls, refresh: () => reply.promise, cancel: () => cancelReply.promise } });
  await h.c.preview();
  const run = h.c.run(),
    cancel = h.c.cancel();
  reply.resolve({ refreshAllResult: result(plan()) });
  await run;
  assert.equal(
    await h.c.preview(),
    false,
    'Do not start another request while its predecessor cancellation is pending'
  );
  cancelReply.resolve({ outcome: 'completed' });
  await cancel;
  assert.equal(h.c.getSnapshot().pending, false);
  assert.equal(h.c.getSnapshot().cancelPending, false);
  assert.equal(h.c.getSnapshot().result.accepted, 2);
  assert(await h.c.preview());
  assert(await h.c.cancel());
  assert.equal(h.c.getSnapshot().preview, null);
});
test('a delayed old-owner Cancel reply cannot keep cancellation disabled in the replacement pane', async () => {
  const reply = defer(),
    cancelReply = defer(),
    h = harness();
  h.set({ controls: { ...h.frame.controls, preview: () => reply.promise, cancel: () => cancelReply.promise } });
  const preview = h.c.preview(),
    cancel = h.c.cancel();
  h.set({ key: 'other', ownerKey: 'other' });
  reply.resolve({ refreshAllPreview: plan() });
  await preview;
  cancelReply.resolve({ outcome: 'completed' });
  await cancel;
  assert.equal(h.c.getSnapshot().cancelPending, false);
  assert.equal(h.c.getSnapshot().preview, null);
  assert.equal(h.c.getSnapshot().result, null);
});
