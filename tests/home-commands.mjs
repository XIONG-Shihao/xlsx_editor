import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  homeCommandDefinitions,
  homeCommandDefinition,
  homeCommandCapability,
  normalizeHomeArguments,
  captureHomeBinding,
  homeBindingMatches,
  dispatchHomeCommand,
  HomeCommandRetired,
  normalizeHomeBatch,
} from '../ui/homeCommands.ts';
import { ribbonCommands, ribbonCommand, dispatchRibbonCommand } from '../ui/ribbonCommands.ts';

const initial = () => ({
  documentId: 'book-A',
  ownerIdentity: 'session-A',
  state: 'accepted-1',
  sheet: 'Sheet1',
  selection: 'B2:C4',
  selectedCell: 'B2',
  activationEpoch: 1,
  navigationGeneration: 0,
  modeGeneration: 0,
  specialSelectionEpoch: 0,
  mode: 'normal',
  specialSelection: null,
});
function harness(patch = {}) {
  const original = {
    binding: initial(),
    profile: 'xlsx',
    active: true,
    canEdit: true,
    busy: false,
    acceptBatch: async () => {},
    ...patch,
  };
  let current = original;
  return {
    original,
    scope: { captureFrame: () => original, currentFrame: () => current },
    set: (patch) => {
      current = { ...current, ...patch, binding: { ...current.binding, ...patch.binding } };
    },
    current: () => current,
  };
}
const deferred = () => {
  let resolve, reject;
  const promise = new Promise((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
};
const batch = () => [{ type: 'formatRange', sheet: 'Sheet1', range: 'B2:C4', format: { bold: true } }];

test('Home registry inventories all seven existing groups, individual routes and no invented screenshot IDs', () => {
  assert.deepEqual(
    [...new Set(homeCommandDefinitions.map((c) => c.group))],
    ['Clipboard', 'Font', 'Alignment', 'Number', 'Styles', 'Cells', 'Editing']
  );
  assert.equal(new Set(homeCommandDefinitions.map((c) => c.id)).size, homeCommandDefinitions.length);
  for (const id of [
    'home.copy',
    'home.cut',
    'home.paste',
    'home.cancel-paste',
    'home.paint-format',
    'home.font-family',
    'home.font-size-increase',
    'home.bold',
    'home.borders',
    'home.align-top',
    'home.align-right',
    'home.wrap',
    'home.merge-center',
    'home.number-format',
    'home.accounting',
    'home.decimals-decrease',
    'styles.conditional',
    'home.conditional-apply',
    'home.create-table',
    'home.style-heading',
    'home.insert-row',
    'home.delete-column',
    'home.column-width',
    'home.fill-right',
    'home.clear-all',
    'home.find-page',
    'home.replace-all',
    'home.find-draft',
    'home.select-formulas',
    'home.undo',
    'home.redo',
  ])
    assert(homeCommandDefinition(id), id);
  for (const item of homeCommandDefinitions) {
    const registered = ribbonCommand(item.id);
    assert.equal(registered.tab, 'Home');
    assert.deepEqual(registered.matrixIds, []);
    assert.equal(ribbonCommands.filter((c) => c.id === item.id).length, 1);
  }
  // Match action-name segments: "hyperlinks" is an observed Clear command,
  // not an invented Ink action merely because it contains those three letters.
  assert(!homeCommandDefinitions.some((c) => /(?:^|[.-])(?:refresh|goal|ink|theme|advanced)(?:[.-]|$)/i.test(c.id)));
});
test('closed arguments preserve intentional typed values and incomplete local drafts, refuse coercion and extra fields', () => {
  for (const [id, args] of [
    ['home.bold', { value: false }],
    ['home.font-size', { value: 6 }],
    ['home.column-width', { value: 0.01 }],
    ['home.row-height', { value: 409.5 }],
    ['home.paste', { text: '00123\t=SUM(A1)\t#N/A' }],
    ['home.conditional-draft', { field: 'threshold', value: '-' }],
    ['home.table-draft', { name: '' }],
    ['home.find-draft', { field: 'replacement', value: '00123' }],
    ['home.find-page', { direction: 'next' }],
    ['home.find-match', { index: 999 }],
    ['home.dimension-draft', { dimension: 'row-height', value: '-' }],
    ['home.menu', { value: 'format' }],
  ]) {
    const normalized = normalizeHomeArguments(id, args);
    assert.deepEqual(normalized, args);
    assert.notEqual(normalized, args);
    assert(Object.isFrozen(normalized));
  }
  assert.equal(normalizeHomeArguments('home.find-cancel'), undefined);
  for (const [id, args] of [
    ['home.copy', {}],
    ['home.bold', { value: 'false' }],
    ['home.font-size', { value: NaN }],
    ['home.font-size', { value: 73 }],
    ['home.column-width', { value: 0 }],
    ['home.row-height', { value: 410 }],
    ['home.font-color', { value: '#ff0000' }],
    ['home.number-format', { value: 'arbitrary' }],
    ['home.menu', { value: 'new-feature' }],
    ['home.find-draft', { field: 'scope', value: 'profile' }],
    ['home.find-draft', { field: 'mode', value: 'cached' }],
    ['home.font-family', { value: 'Arial', path: '/foreign' }],
    ['home.create-table', { name: '' }],
    ['home.conditional-apply', { rule: 'greaterThan', fillColor: 'FF0000' }],
    ['home.conditional-apply', { rule: 'duplicateValues', threshold: 0, fillColor: 'FF0000' }],
  ])
    assert.throws(() => normalizeHomeArguments(id, args), /Invalid Home/);
  assert.throws(
    () => normalizeHomeArguments('home.bold', Object.assign(Object.create({ value: true }), {})),
    /Invalid Home/
  );
  for (const [id, args] of [
    ['home.find-match', { index: 1000 }],
    ['home.find-match', { index: -1 }],
    ['home.find-match', { index: 0.5 }],
    ['home.find-match', { index: 0, sheet: 'foreign' }],
    ['home.dimension-draft', { dimension: 'workbook', value: '15' }],
    ['home.dimension-draft', { dimension: 'row-height', value: 15 }],
  ])
    assert.throws(() => normalizeHomeArguments(id, args), /Invalid Home/);
  assert.deepEqual(normalizeHomeArguments('home.fill-range', { source: 'A1:A1000', range: 'B1:B10000' }), {
    source: 'A1:A1000',
    range: 'B1:B10000',
  });
  for (const args of [
    { source: 'A1:A1001', range: 'B1:B2' },
    { source: 'A1', range: 'B1:B10001' },
    { source: 'Sheet1!A1', range: 'B1:B2' },
    { source: 'a1', range: 'B1:B2' },
    { source: 'A2:A1', range: 'B1:B2' },
    { source: 'A1', range: 'XFE1' },
    { source: 'A1', range: 'B1', sheet: 'other' },
  ])
    assert.throws(() => normalizeHomeArguments('home.fill-range', args), /Invalid Home/);
  const original = { format: { bold: true, numberFormat: 'text', styleId: 0 } },
    normalized = normalizeHomeArguments('home.paint-format', original);
  original.format.bold = false;
  assert.equal(normalized.format.bold, true);
  assert(Object.isFrozen(normalized.format));
  assert.throws(() => normalizeHomeArguments('home.paint-format', { format: { value: 1 } }), /Invalid Home formatting/);
});
test('Up/Left fill are closed captured workbook routes between Right and Series with identical write barriers', () => {
  const h = harness(),
    right = homeCommandDefinition('home.fill-right'),
    series = homeCommandDefinition('home.fill-series');
  for (const id of ['home.fill-up', 'home.fill-left']) {
    const command = homeCommandDefinition(id);
    assert.equal(command.kind, 'workbook');
    assert.equal(command.requiresEdit, true);
    assert.equal(command.requiresBatch, true);
    assert.equal(command.allowSpecial, false);
    assert(command.order > right.order && command.order < series.order);
    assert.equal(normalizeHomeArguments(id), undefined);
    assert.throws(() => normalizeHomeArguments(id, { source: 'foreign' }), /Invalid Home/);
    assert.equal(homeCommandCapability(id, h.original).state, 'enabled');
    for (const patch of [
      { canEdit: false },
      { busy: true },
      { active: false },
      { acceptBatch: undefined },
      { binding: { ...initial(), mode: 'pageLayout' } },
      { binding: { ...initial(), specialSelection: ['B2'] } },
    ])
      assert.equal(homeCommandCapability(id, { ...h.original, ...patch }).state, 'disabled');
    for (const profile of ['delimited', 'passive'])
      assert.equal(homeCommandCapability(id, { ...h.original, profile }).state, 'hidden');
    assert.equal(ribbonCommand(id).tab, 'Home');
    assert.deepEqual(ribbonCommand(id).matrixIds, []);
  }
  assert(homeCommandDefinition('home.fill-up').order < homeCommandDefinition('home.fill-left').order);
});
test('Go To Special preview/select are selection-only routes and cannot submit a workbook batch', async () => {
  const h = harness({ canEdit: false, acceptBatch: undefined });
  for (const id of ['home.special-open', 'home.special-preview', 'home.special-select']) {
    assert.equal(homeCommandDefinition(id).requiresEdit, false);
    assert.equal(homeCommandDefinition(id).requiresBatch, false);
    assert.equal(homeCommandDefinition(id).allowSpecial, false);
    assert.equal(homeCommandCapability(id, h.original).state, 'enabled');
    assert.equal(homeCommandCapability(id, { ...h.original, profile: 'passive' }).state, 'hidden');
    const result = await dispatchHomeCommand(id, h.scope, (lease) => lease.accept(batch()));
    assert.equal(result.outcome, 'failed');
    assert.match(result.error.message, /unavailable for this action/);
  }
  assert.equal(ribbonCommand('home.special-open').tab, 'Home');
});
test('readonly/passive inspection survives jobs while writes, unsupported profiles and missing strict batch refuse', () => {
  const h = harness({ canEdit: false, busy: true });
  for (const id of ['home.copy', 'home.find-open', 'home.find-search', 'home.menu', 'styles.conditional'])
    assert.equal(homeCommandCapability(id, h.original).state, 'enabled', id);
  assert.match(homeCommandCapability('home.bold', h.original).reason, /read-only/);
  assert.equal(homeCommandCapability('home.copy', { ...h.original, profile: 'passive' }).state, 'enabled');
  assert.equal(
    homeCommandCapability('home.clear-contents', { ...h.original, profile: 'passive', canEdit: true }).state,
    'hidden'
  );
  assert.equal(
    homeCommandCapability('home.bold', { ...h.original, profile: 'delimited', canEdit: true }).state,
    'hidden'
  );
  assert.match(
    homeCommandCapability('home.bold', { ...h.original, canEdit: true, busy: false, acceptBatch: undefined }).reason,
    /unavailable/
  );
  assert.equal(homeCommandCapability('home.copy', h.original, false).state, 'hidden');
  assert.match(
    homeCommandCapability('home.clear-contents', {
      ...h.original,
      canEdit: true,
      busy: false,
      binding: { ...initial(), mode: 'pageLayout' },
    }).reason,
    /Normal/
  );
});
test('special targets are immutable, formatting/Clear admitted, ordinary rectangle actions refused and epoch prevents revival', async () => {
  const source = { ...initial(), specialSelection: ['C4', 'B2'], specialSelectionEpoch: 3 },
    snapshot = captureHomeBinding(source);
  source.specialSelection.push('D5');
  assert.deepEqual(snapshot.specialSelection, ['B2', 'C4']);
  assert(Object.isFrozen(snapshot.specialSelection));
  assert(homeBindingMatches(snapshot, { ...snapshot, specialSelection: ['C4', 'B2'] }));
  const h = harness({ binding: snapshot });
  assert.equal(homeCommandCapability('home.bold', h.original).state, 'enabled');
  assert.equal(homeCommandCapability('home.clear-contents', h.original).state, 'enabled');
  assert.equal(homeCommandCapability('home.paste', h.original).state, 'disabled');
  for (const id of ['home.replace-one', 'home.replace-all'])
    assert.equal(homeCommandCapability(id, h.original).state, 'disabled');
  const held = deferred();
  let writes = 0;
  const pending = dispatchHomeCommand(
    'home.bold',
    h.scope,
    async (lease) => {
      await lease.read(() => held.promise);
      await lease.accept(batch());
      writes++;
    },
    { value: true }
  );
  h.set({ binding: { specialSelection: null, specialSelectionEpoch: 4 } });
  h.set({ binding: { specialSelection: ['C4', 'B2'], specialSelectionEpoch: 5 } });
  held.resolve('actual read');
  const result = await pending;
  assert.equal(result.outcome, 'failed');
  assert(result.error instanceof HomeCommandRetired);
  assert.equal(result.current, false);
  assert.equal(writes, 0);
});
test('read preflight is retired by every owner/head/target/mode/navigation fence without calling captured batch', async () => {
  for (const patch of [
    { ownerIdentity: 'session-B' },
    { documentId: 'book-B' },
    { state: 'accepted-2' },
    { sheet: 'Other' },
    { selection: 'D9', selectedCell: 'D9' },
    { activationEpoch: 2 },
    { navigationGeneration: 1 },
    { modeGeneration: 1 },
    { mode: 'pageLayout' },
    { specialSelectionEpoch: 1 },
  ]) {
    const held = deferred();
    let calls = 0;
    const h = harness({
      acceptBatch: async () => {
        calls++;
      },
    });
    const pending = dispatchHomeCommand('home.paste', h.scope, async (lease) => {
      await lease.read(() => held.promise);
      await lease.accept([{ type: 'setCells', sheet: 'Sheet1', cells: [{ cell: 'B2', value: '00123' }] }]);
    });
    h.set({ binding: patch });
    held.resolve('00123');
    const result = await pending;
    assert.equal(result.outcome, 'failed', JSON.stringify(patch));
    assert(result.error instanceof HomeCommandRetired);
    assert.equal(calls, 0);
    assert.equal(result.current, false);
  }
});
test('retained render handlers never capture a replacement frame or changed availability', async () => {
  let oldCalls = 0,
    newCalls = 0;
  const h = harness({
    acceptBatch: async () => {
      oldCalls++;
    },
  });
  h.set({
    binding: { ownerIdentity: 'session-B', state: 'accepted-B' },
    acceptBatch: async () => {
      newCalls++;
    },
  });
  const result = await dispatchHomeCommand('home.bold', h.scope, (lease) => lease.accept(batch()), { value: true });
  assert.equal(result.outcome, 'stale');
  assert.equal(oldCalls + newCalls, 0);
  const pending = harness(),
    hold = deferred();
  const rejected = dispatchHomeCommand(
    'home.bold',
    pending.scope,
    async (lease) => {
      await lease.read(() => hold.promise);
      await lease.accept(batch());
    },
    { value: true }
  );
  pending.set({ canEdit: false });
  hold.resolve('cells');
  const refusal = await rejected;
  assert.equal(refusal.outcome, 'failed');
  assert.match(refusal.error.message, /read-only/);
});
test('actual Promise acknowledgement stays pending and accepted head advance remains a truthful completion', async () => {
  const owner = deferred();
  let captured,
    settled = false;
  const h = harness({
    acceptBatch: (commands) => {
      captured = commands;
      return owner.promise;
    },
  });
  const pending = dispatchHomeCommand('home.bold', h.scope, (lease) => lease.accept(batch()), { value: true });
  pending.then(() => {
    settled = true;
  });
  await Promise.resolve();
  assert.equal(settled, false);
  assert.deepEqual(captured, batch());
  h.set({ binding: { state: 'accepted-2' } });
  owner.resolve();
  const result = await pending;
  assert.equal(result.outcome, 'completed');
  assert.equal(result.current, false);
  assert.equal(result.runtimeCurrent, true);
  assert.equal(result.binding.state, 'accepted-1');
});
test('late accepted writes and late failures retain original owner without new-owner adoption', async () => {
  for (const accepted of [true, false]) {
    const owner = deferred(),
      h = harness({ acceptBatch: () => owner.promise });
    const pending = dispatchHomeCommand('home.clear-contents', h.scope, (lease) =>
      lease.accept([{ type: 'clearRange', sheet: 'Sheet1', range: 'B2:C4' }])
    );
    h.set({ binding: { documentId: 'book-B', ownerIdentity: 'session-B', state: 'other', activationEpoch: 2 } });
    if (accepted) owner.resolve();
    else owner.reject(Error('Original owner refused'));
    const result = await pending;
    assert.equal(result.outcome, accepted ? 'completed' : 'failed');
    assert.equal(result.current, false);
    assert.equal(result.runtimeCurrent, false);
    assert.equal(result.binding.documentId, 'book-A');
  }
});
test('strict CSV callback supports literal paste/Clear without XLSX-only admission or coercion', async () => {
  const accepted = [];
  const h = harness({ profile: 'delimited', acceptBatch: async (commands) => accepted.push(commands) });
  const cells = [
    { cell: 'B2', value: '00123' },
    { cell: 'B3', value: '=SUM(A1)' },
    { cell: 'B4', value: '#N/A' },
  ];
  assert.equal(
    (
      await dispatchHomeCommand(
        'home.paste',
        h.scope,
        (lease) => lease.accept([{ type: 'setCells', sheet: 'Sheet1', cells }]),
        { text: '00123\n=SUM(A1)\n#N/A' }
      )
    ).outcome,
    'completed'
  );
  assert.equal(
    (
      await dispatchHomeCommand('home.clear-contents', h.scope, (lease) =>
        lease.accept([{ type: 'setCells', sheet: 'Sheet1', cells: [{ cell: 'B2', value: null }] }])
      )
    ).outcome,
    'completed'
  );
  assert.deepEqual(accepted[0][0].cells, cells);
  assert.equal(accepted.length, 2);
  assert.equal(
    (await dispatchHomeCommand('home.bold', h.scope, () => assert.fail('CSV styling'), { value: true })).outcome,
    'refused'
  );
});
test('portable CSV bound callback is explicit and its identity cannot change during an asynchronous preflight', async () => {
  let calls = 0;
  const captured = async () => {
      calls++;
    },
    h = harness({
      profile: 'delimited',
      binding: { ...initial(), ownerIdentity: null },
      portableOwnerBound: true,
      acceptBatch: captured,
    });
  assert.equal(homeCommandCapability('home.clear-contents', h.original).state, 'enabled');
  assert.equal(
    homeCommandCapability('home.clear-contents', { ...h.original, portableOwnerBound: false }).state,
    'disabled'
  );
  assert.equal(homeCommandCapability('home.clear-contents', { ...h.original, profile: 'xlsx' }).state, 'disabled');
  const held = deferred(),
    pending = dispatchHomeCommand('home.paste', h.scope, async (lease) => {
      await lease.read(() => held.promise);
      await lease.accept([{ type: 'setCells', sheet: 'Sheet1', cells: [{ cell: 'B2', value: 'x' }] }]);
    });
  h.set({ acceptBatch: async () => assert.fail('new callback') });
  held.resolve('x');
  const result = await pending;
  assert.equal(result.outcome, 'failed');
  assert(result.error instanceof HomeCommandRetired);
  assert.equal(calls, 0);
});
test('large paste and strict history jobs await their own actual Promise and preserve finishing outcomes', async () => {
  for (const id of ['home.paste', 'home.undo', 'home.redo']) {
    const job = deferred(),
      h = harness();
    let started = 0,
      finished = false;
    const pending = dispatchHomeCommand(id, h.scope, (lease) =>
      lease.perform(() => {
        started++;
        return job.promise;
      })
    );
    pending.then(() => {
      finished = true;
    });
    await Promise.resolve();
    assert.equal(started, 1);
    assert.equal(finished, false);
    h.set({ binding: { state: 'accepted-2' } });
    job.resolve('durable owner result');
    const result = await pending;
    assert.equal(result.outcome, 'completed');
    assert.equal(result.value, 'durable owner result');
    assert.equal(result.runtimeCurrent, true);
  }
  const result = await dispatchHomeCommand('home.undo', harness().scope, (lease) => lease.perform(() => undefined));
  assert.equal(result.outcome, 'failed');
  assert.match(result.error.message, /awaited acknowledgement/);
});
test('copy/Find/menu/cancel inspection cannot call the workbook batch; invalid/disabled actions never invoke handlers', async () => {
  const h = harness({ busy: true });
  let cancelled = 0;
  assert.equal(
    (
      await dispatchHomeCommand('home.cancel-paste', h.scope, () => {
        cancelled++;
      })
    ).outcome,
    'completed'
  );
  assert.equal(cancelled, 1);
  assert.equal((await dispatchHomeCommand('home.copy', h.scope, (lease) => lease.accept(batch()))).outcome, 'failed');
  let calls = 0;
  assert.equal(
    (
      await dispatchHomeCommand(
        'home.bold',
        h.scope,
        () => {
          calls++;
        },
        { value: true }
      )
    ).outcome,
    'refused'
  );
  assert.equal(calls, 0);
  const invalid = await dispatchHomeCommand(
    'home.bold',
    harness().scope,
    () => {
      calls++;
    },
    { value: 'true' }
  );
  assert.equal(invalid.outcome, 'failed');
  assert.equal(calls, 0);
  await assert.rejects(
    dispatchHomeCommand(
      'home.copy',
      {
        captureFrame: () => ({ ...h.original, binding: { ...initial(), specialSelection: ['B2', 'B2'] } }),
        currentFrame: () => h.original,
      },
      () => {}
    ),
    /special selection/
  );
});
test('batch normalization preserves source arguments, signed zero and one atomic callback while splitting transport pages', () => {
  const source = [
      {
        type: 'setCells',
        sheet: 'Sheet1',
        cells: Array.from({ length: 1201 }, (_, i) => ({ cell: 'A' + (i + 1), value: i === 0 ? -0 : i })),
      },
      { type: 'formatRange', sheet: 'Sheet1', range: 'A1:A1201', format: { bold: true } },
    ],
    copy = structuredClone(source),
    normalized = normalizeHomeBatch(source);
  assert.deepEqual(
    normalized.filter((c) => c.type === 'setCells').map((c) => c.cells.length),
    [1000, 201]
  );
  assert.equal(normalized.length, 3);
  assert(Object.is(normalized[0].cells[0].value, -0));
  normalized[0].cells[1].value = 99;
  assert.deepEqual(source, copy);
  assert.throws(() => normalizeHomeBatch([]), /no workbook changes/);
});
test('shared ribbon dispatcher recognizes closed Home arguments without weakening its existing argument profiles', async () => {
  const b = { documentId: 'book-A', state: 'accepted-1', sheet: 'Sheet1', selection: 'B2:C4', activationEpoch: 1 },
    scope = {
      captureBinding: () => b,
      isCurrent: () => true,
      availability: () => ({ profile: 'xlsx', active: true, canEdit: true, busy: false }),
    };
  let received;
  assert.equal(
    (
      await dispatchRibbonCommand(
        'home.bold',
        scope,
        (_binding, args) => {
          received = args;
        },
        { value: true }
      )
    ).outcome,
    'completed'
  );
  assert.deepEqual(received, { value: true });
  assert(Object.isFrozen(received));
  await assert.rejects(
    dispatchRibbonCommand('home.bold', scope, () => assert.fail('invalid handler'), { value: true, mode: 'foreign' }),
    /Invalid Home/
  );
  await assert.rejects(
    dispatchRibbonCommand('view.zoom', scope, () => assert.fail('old argument handler'), { value: true }),
    /Invalid ribbon/
  );
});
