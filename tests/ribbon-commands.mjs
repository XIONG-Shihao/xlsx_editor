import { test } from 'node:test';
import assert from 'node:assert/strict';
import { ribbonCommands, ribbonCommand, ribbonCapability, dispatchRibbonCommand } from '../ui/ribbonCommands.ts';

const initial = () => ({
  documentId: 'doc-A',
  state: 'head-1',
  sheet: 'Sheet1',
  selection: 'B2:C4',
  activationEpoch: 1,
});
function harness(options = {}) {
  let current = initial();
  const scope = {
    captureBinding: () => current,
    isCurrent: (binding) => Object.keys(initial()).every((key) => current[key] === binding[key]),
    availability: () => ({ profile: 'xlsx', active: true, canEdit: true, busy: false, ...options }),
  };
  return {
    scope,
    current: () => current,
    set: (patch) => {
      current = { ...current, ...patch };
    },
  };
}
function deferred() {
  let resolve, reject;
  const promise = new Promise((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
}

test('registered action IDs are unique, carry well-formed reference ids and never advertise the whole planned inventory', async () => {
  assert.equal(new Set(ribbonCommands.map((c) => c.id)).size, ribbonCommands.length);
  // The reference ids once pointed at rows in a documentation table; they are kept as stable labels.
  for (const command of ribbonCommands) for (const id of command.matrixIds) assert.match(id, /^XL-[A-Z]+\d+$/, id);
  assert.deepEqual(
    ribbonCommands.filter((c) => c.tab === 'Draw').map((c) => c.id),
    ['draw.preset', 'draw.pen', 'draw.select-strokes', 'draw.eraser', 'draw.select-cells']
  );
  assert.deepEqual(
    ribbonCommands
      .filter((c) => c.tab === 'Automate')
      .map((c) => c.id)
      .sort(),
    ['actions.freeze-selection', 'actions.library', 'actions.new', 'actions.unhide-all']
  );
  assert.deepEqual(
    ribbonCommands
      .filter((c) => c.tab === 'Review')
      .map((c) => c.id)
      .sort(),
    ['review.accessibility', 'review.notes', 'review.spelling', 'review.statistics']
  );
  assert(!ribbonCommands.some((c) => /Forms|Office Scripts|Stocks|Currencies/.test(c.label)));
  assert.deepEqual(ribbonCommand('view.show-formulas').matrixIds, ['XL-F10', 'XL-V04']);
  assert.deepEqual(ribbonCommand('data.refresh-all').matrixIds, ['XL-DA02']);
  assert.equal(
    ribbonCapability('data.refresh-all', {
      profile: 'xlsx',
      active: true,
      canEdit: false,
      busy: false,
      hasHandler: true,
    }).state,
    'disabled'
  );
});
test('capabilities preserve read-only inspection/local display while refusing actual writes and missing handlers', () => {
  const scope = { profile: 'xlsx', active: true, canEdit: false, busy: true, hasHandler: true };
  assert.equal(ribbonCapability('names.manager', scope).state, 'enabled');
  assert.equal(ribbonCapability('view.zoom', scope).state, 'enabled');
  assert.match(ribbonCapability('view.freeze', scope).reason, /read-only/);
  assert.equal(ribbonCapability('data.import-file', { ...scope, canEdit: true }).state, 'disabled');
  assert.equal(
    ribbonCapability('view.zoom', { ...scope, hasHandler: false, capability: { state: 'enabled' } }).state,
    'hidden'
  );
  assert.equal(ribbonCapability('data.queries', { ...scope, profile: 'delimited' }).state, 'hidden');
  assert.match(ribbonCapability('names.manager', { ...scope, active: false }).reason, /Activate/);
  assert.deepEqual(
    ribbonCapability('view.zoom', { ...scope, capability: { state: 'disabled', reason: 'No geometry available.' } }),
    { state: 'disabled', reason: 'No geometry available.' }
  );
});
test('async dispatch captures immutable initiating identity/selection and retains accepted success after state advancement', async () => {
  const h = harness(),
    job = deferred();
  let captured;
  const pending = dispatchRibbonCommand('formulas.auto-sum', h.scope, (binding) => {
    captured = binding;
    return job.promise;
  });
  assert.deepEqual(captured, initial());
  assert(Object.isFrozen(captured));
  h.set({ state: 'head-2', selection: 'H9' });
  job.resolve({ accepted: true, state: 'head-2' });
  const result = await pending;
  assert.equal(result.outcome, 'completed');
  assert.equal(result.current, false);
  assert.deepEqual(result.value, { accepted: true, state: 'head-2' });
  assert.equal(result.binding.selection, 'B2:C4');
});
test('stale before dispatch skips handler, and A→B→A activation cannot accept a late reply as current', async () => {
  const h = harness();
  let calls = 0;
  const stale = { ...h.scope, isCurrent: () => false };
  assert.equal((await dispatchRibbonCommand('names.manager', stale, () => calls++)).outcome, 'stale');
  assert.equal(calls, 0);
  const job = deferred(),
    pending = dispatchRibbonCommand('names.manager', h.scope, () => job.promise);
  h.set({ documentId: 'doc-B', activationEpoch: 2 });
  h.set({ documentId: 'doc-A', activationEpoch: 3 });
  job.resolve('old inventory');
  const result = await pending;
  assert.equal(result.outcome, 'completed');
  assert.equal(result.current, false);
  assert.equal(result.binding.activationEpoch, 1);
});
test('late failures retain captured owner and current=false instead of being surfaced as the active workbook error', async () => {
  const h = harness(),
    job = deferred(),
    error = new Error('initiating owner refused'),
    pending = dispatchRibbonCommand('data.import-file', h.scope, () => job.promise);
  h.set({ documentId: 'doc-B', state: 'other', activationEpoch: 2 });
  job.reject(error);
  const result = await pending;
  assert.equal(result.outcome, 'failed');
  assert.equal(result.current, false);
  assert.equal(result.error, error);
  assert.equal(result.binding.documentId, 'doc-A');
});
test('closed arguments reject invalid zoom/toggles/freeze and disabled callbacks never run', async () => {
  const h = harness();
  let calls = 0;
  for (const args of [{ zoom: 9 }, { zoom: 401 }, { zoom: 125.5 }, { zoom: NaN }, { zoom: 125, path: '/ungranted' }])
    await assert.rejects(
      dispatchRibbonCommand('view.zoom', h.scope, () => calls++, args),
      /Invalid ribbon/
    );
  await assert.rejects(
    dispatchRibbonCommand('view.gridlines', h.scope, () => calls++, { value: 'false' }),
    /Invalid ribbon/
  );
  await assert.rejects(
    dispatchRibbonCommand('view.freeze', h.scope, () => calls++, { mode: 'all' }),
    /Invalid ribbon/
  );
  assert.equal(calls, 0);
  const args = { zoom: 125 };
  let received;
  const ok = await dispatchRibbonCommand(
    'view.zoom',
    h.scope,
    (_binding, value) => {
      received = value;
    },
    args
  );
  assert.equal(ok.outcome, 'completed');
  assert(Object.isFrozen(received));
  assert.notEqual(received, args);
  assert.equal(received.zoom, 125);
  const readonly = harness({ canEdit: false });
  assert.equal(
    (await dispatchRibbonCommand('view.freeze', readonly.scope, () => calls++, { mode: 'row' })).outcome,
    'refused'
  );
  assert.equal(calls, 0);
  await assert.rejects(
    dispatchRibbonCommand(
      'view.zoom',
      { ...h.scope, captureBinding: () => ({ ...initial(), state: '' }) },
      () => calls++,
      { zoom: 125 }
    ),
    /accepted document binding/
  );
});
