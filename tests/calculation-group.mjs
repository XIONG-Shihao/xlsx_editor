/** Calculation-group protocol: the real coordinator against simulated trusted owner commits. */
import assert from 'node:assert/strict';
import { randomUUID } from 'node:crypto';
import { runCalculationGroup, validateCalculationAction } from '../adapters/workbook-calculation-group.mjs';
// Cases run one after another; each awaits the coordinator to completion.
const check = async (name, body) => {
  await body();
};
function owner(name, extra = {}) {
  return {
    documentId: 'document-' + name,
    ownerId: 'owner-' + name,
    state: 'base-' + name,
    revision: 'revision-' + name,
    name: name + '.xlsx',
    profile: 'xlsx',
    readOnly: false,
    canChangeOptions: true,
    canCalculateSheet: true,
    ...extra,
  };
}
function host(targets) {
  const slots = new Map(
    targets.map((target) => [
      target.ownerId,
      {
        target,
        latest: {
          documentId: target.documentId,
          ownerId: target.ownerId,
          state: target.state,
          revision: target.revision,
        },
        alive: true,
        commits: [],
      },
    ])
  );
  const calls = [],
    authorization = [],
    feedback = [],
    progress = [];
  const current = (target, state) => {
    const slot = slots.get(target.ownerId);
    return (
      !!slot?.alive &&
      slot.latest.documentId === target.documentId &&
      slot.latest.ownerId === target.ownerId &&
      slot.latest.state === state
    );
  };
  const authorize = async (target) => {
    authorization.push(target.ownerId);
  };
  const apply = async (target, commands) => {
    const slot = slots.get(target.ownerId);
    assert(current(target, slot.latest.state));
    calls.push({ ownerId: target.ownerId, commands: structuredClone(commands) });
    slot.commits.push(structuredClone(commands));
    slot.latest = {
      ...slot.latest,
      state: 'accepted-' + target.ownerId + '-' + slot.commits.length,
      revision: 'accepted-revision-' + slot.commits.length,
    };
    return structuredClone(slot.latest);
  };
  const args = {
    requestId: randomUUID(),
    initiator: targets[0],
    action: { type: 'calculateAllOpen' },
    targets,
    current,
    authorize,
    apply,
    progress: (value) => progress.push(structuredClone(value)),
    onAccepted: (target, reply) => feedback.push({ ownerId: target.ownerId, reply: structuredClone(reply) }),
  };
  return { slots, calls, authorization, feedback, progress, args, current, apply };
}
await check('Closed actions/bindings/job identity/roster budgets refuse before any owner call', async () => {
  const good = [
    { type: 'calculateAllOpen' },
    { type: 'calculateSheet', sheet: 'Report' },
    { type: 'setOptions', mode: 'manual' },
    { type: 'setOptions', mode: 'auto', recalculateBeforeSave: false },
  ];
  for (const action of good) assert.deepEqual(validateCalculationAction(action), action);
  const bad = [
    null,
    [],
    { type: 'calculateAllOpen', participants: [] },
    { type: 'calculateSheet', sheet: '' },
    { type: 'calculateSheet', sheet: 'bad\0sheet' },
    { type: 'setOptions', mode: 'autoNoTable' },
    { type: 'setOptions', mode: 'manual', recalculateBeforeSave: 1 },
    { type: 'setOptions', mode: 'auto', force: true },
    Object.create({ type: 'calculateAllOpen' }),
  ];
  for (const action of bad) assert.throws(() => validateCalculationAction(action));
  const h = host([owner('A')]);
  for (const overrides of [
    { requestId: 'bad' },
    { targets: [] },
    { maxParticipants: 0 },
    { maxParticipants: 129 },
    { initiator: { ...h.args.initiator, ownerId: 'foreign' } },
    { initiator: { ...h.args.initiator, state: 'changed' } },
    { targets: [h.args.targets[0], h.args.targets[0]] },
    { targets: Array.from({ length: 33 }, (_, i) => owner(String(i))) },
  ])
    await assert.rejects(() => runCalculationGroup({ ...h.args, ...overrides }));
  assert.deepEqual(h.calls, []);
  return { validActions: good, refusedActions: bad.length, rosterRefusals: 8 };
});
await check('All-open captures ordered native owners and explicitly skips readonly/delimited/passive', async () => {
  const targets = [
      owner('A'),
      owner('B'),
      owner('readonly', { readOnly: true }),
      owner('literal', { profile: 'delimited' }),
      owner('passive', { profile: 'passive' }),
    ],
    h = host(targets),
    receipt = await runCalculationGroup(h.args);
  assert.equal(receipt.schema, 1);
  assert.equal(receipt.consistency, 'independent-workbook-commits');
  assert.equal(receipt.outcome, 'completed');
  assert.deepEqual(
    receipt.results.map((r) => r.status),
    ['accepted', 'accepted', 'skipped', 'skipped', 'skipped']
  );
  assert.deepEqual(
    h.calls.map((c) => c.ownerId),
    ['owner-A', 'owner-B']
  );
  for (const call of h.calls) assert.deepEqual(call.commands, [{ type: 'recalculate' }]);
  for (const target of targets.slice(2)) assert.equal(h.slots.get(target.ownerId).commits.length, 0);
  assert.equal(h.feedback.length, 2);
  assert.equal(h.progress.at(-1).completed, 5);
  assert.equal(h.progress.at(-1).cancellable, false);
  return { receipt, calls: h.calls };
});
await check('Options send one exact owner command per eligible captured workbook, preserving omitted/explicit Save default', async () => {
  const targets = [
      owner('A'),
      owner('B'),
      owner('unsupported', { canChangeOptions: false, reason: 'Protected original calculation policy' }),
    ],
    h = host(targets),
    receipt = await runCalculationGroup({ ...h.args, action: { type: 'setOptions', mode: 'manual' } });
  assert.deepEqual(
    receipt.results.map((r) => r.status),
    ['accepted', 'accepted', 'skipped']
  );
  for (const call of h.calls) assert.deepEqual(call.commands, [{ type: 'setCalculationOptions', mode: 'manual' }]);
  const explicit = host([owner('false-save')]);
  const second = await runCalculationGroup({
    ...explicit.args,
    action: { type: 'setOptions', mode: 'manual', recalculateBeforeSave: false },
  });
  assert.equal(second.results[0].status, 'accepted');
  assert.deepEqual(explicit.calls[0].commands, [
    { type: 'setCalculationOptions', mode: 'manual', recalculateBeforeSave: false },
  ]);
  return { receipt, explicit: second, commands: explicit.calls };
});
await check('Calculate Sheet visits only the captured initiator and requires actual capability', async () => {
  const h = host([owner('A'), owner('B')]),
    receipt = await runCalculationGroup({ ...h.args, action: { type: 'calculateSheet', sheet: 'Report' } });
  assert.equal(receipt.results.length, 1);
  assert.deepEqual(h.calls, [{ ownerId: 'owner-A', commands: [{ type: 'calculateSheet', sheet: 'Report' }] }]);
  assert.equal(h.slots.get('owner-B').commits.length, 0);
  const denied = host([owner('opaque', { canCalculateSheet: false })]),
    refused = await runCalculationGroup({ ...denied.args, action: { type: 'calculateSheet', sheet: 'Report' } });
  assert.equal(refused.results[0].status, 'skipped');
  assert.deepEqual(denied.calls, []);
  return { receipt, unavailable: refused };
});
await check('Changed queued head and revoked authorization refuse without overwriting earlier accepted owners', async () => {
  const h = host([owner('A'), owner('stale'), owner('revoked')]);
  const authorize = async (target) => {
    await h.args.authorize(target);
    if (target.ownerId === 'owner-stale') h.slots.get(target.ownerId).latest.state = 'external-edit';
    if (target.ownerId === 'owner-revoked') throw Error('Project edit access revoked');
  };
  const receipt = await runCalculationGroup({ ...h.args, authorize });
  assert.equal(receipt.outcome, 'partial');
  assert.deepEqual(
    receipt.results.map((r) => r.status),
    ['accepted', 'refused', 'refused']
  );
  assert.equal(h.calls.length, 1);
  assert.equal(h.slots.get('owner-stale').latest.state, 'external-edit');
  assert.equal(h.slots.get('owner-revoked').commits.length, 0);
  return { receipt, calls: h.calls };
});
await check('One failed calculation remains a per-owner failure while another owner can independently commit', async () => {
  const h = host([owner('A'), owner('B')]),
    receipt = await runCalculationGroup({
      ...h.args,
      apply: async (target, commands) => {
        if (target.ownerId === 'owner-A') throw Error('Unqualified prerequisite graph');
        return h.apply(target, commands);
      },
    });
  assert.equal(receipt.outcome, 'partial');
  assert.deepEqual(
    receipt.results.map((r) => r.status),
    ['failed', 'accepted']
  );
  assert.equal(h.slots.get('owner-A').commits.length, 0);
  assert.equal(h.slots.get('owner-B').commits.length, 1);
  return { receipt };
});
await check('Cancellation before preparation never starts an owner and during granted failure keeps all prior heads', async () => {
  const h = host([owner('A'), owner('B')]),
    abort = new AbortController();
  abort.abort();
  const receipt = await runCalculationGroup({ ...h.args, signal: abort.signal });
  assert.equal(receipt.outcome, 'cancelled');
  assert.deepEqual(
    receipt.results.map((r) => r.status),
    ['notRun', 'notRun']
  );
  assert.deepEqual(h.calls, []);
  const during = host([owner('during'), owner('after')]),
    cancel = new AbortController();
  const denied = await runCalculationGroup({
    ...during.args,
    signal: cancel.signal,
    apply: async () => {
      cancel.abort();
      const error = Error('Cancellation granted before owner commit');
      error.code = 'OWNER_COMMAND_CANCELLED';
      throw error;
    },
  });
  assert.deepEqual(
    denied.results.map((r) => r.status),
    ['cancelled', 'notRun']
  );
  for (const slot of during.slots.values()) assert.equal(slot.commits.length, 0);
  return { before: receipt, during: denied };
});
await check('Cancel after a trusted durable reply keeps that owner accepted and prevents later owners from starting', async () => {
  const h = host([owner('A'), owner('B')]),
    cancel = new AbortController();
  const receipt = await runCalculationGroup({
    ...h.args,
    signal: cancel.signal,
    apply: async (target, commands) => {
      const reply = await h.apply(target, commands);
      cancel.abort();
      return reply;
    },
  });
  assert.equal(receipt.outcome, 'cancelled');
  assert.deepEqual(
    receipt.results.map((r) => r.status),
    ['accepted', 'notRun']
  );
  assert.equal(receipt.results[0].afterState, h.slots.get('owner-A').latest.state);
  assert.equal(h.slots.get('owner-A').commits.length, 1);
  assert.equal(h.slots.get('owner-B').commits.length, 0);
  return { receipt };
});
await check('Post-commit owner retirement still reports acceptance and suppresses foreign view adoption', async () => {
  const h = host([owner('A'), owner('B')]);
  const receipt = await runCalculationGroup({
    ...h.args,
    apply: async (target, commands) => {
      const reply = await h.apply(target, commands);
      if (target.ownerId === 'owner-A') h.slots.get(target.ownerId).alive = false;
      return reply;
    },
  });
  assert.equal(receipt.outcome, 'completed');
  assert.deepEqual(
    receipt.results.map((r) => r.status),
    ['accepted', 'accepted']
  );
  assert.match(receipt.results[0].reason, /original workbook.*retired/);
  assert.equal(h.feedback.length, 1);
  assert.equal(h.feedback[0].ownerId, 'owner-B');
  assert.equal(h.slots.get('owner-A').commits.length, 1);
  return { receipt, adopted: h.feedback };
});
await check('Simultaneous post-commit retirement and Cancel cannot turn an accepted journal receipt into cancellation', async () => {
  const h = host([owner('A'), owner('B')]),
    cancel = new AbortController();
  const receipt = await runCalculationGroup({
    ...h.args,
    signal: cancel.signal,
    apply: async (target, commands) => {
      const reply = await h.apply(target, commands);
      h.slots.get(target.ownerId).alive = false;
      cancel.abort();
      return reply;
    },
  });
  assert.deepEqual(
    receipt.results.map((r) => r.status),
    ['accepted', 'notRun']
  );
  assert.equal(receipt.results[0].afterState, h.slots.get('owner-A').latest.state);
  assert.deepEqual(h.feedback, []);
  return { receipt };
});
await check('Foreign/malformed accepted receipts never adopt as the original owner', async () => {
  const h = host([owner('A'), owner('B')]);
  const receipt = await runCalculationGroup({
    ...h.args,
    apply: async (target) =>
      target.ownerId === 'owner-A'
        ? { documentId: 'foreign', ownerId: 'foreign', state: 'foreign', revision: 'foreign' }
        : null,
  });
  assert.equal(receipt.outcome, 'partial');
  assert.deepEqual(
    receipt.results.map((r) => r.status),
    ['failed', 'failed']
  );
  assert.deepEqual(h.feedback, []);
  for (const slot of h.slots.values()) assert.equal(slot.commits.length, 0);
  return { receipt };
});
await check('Progress/adoption feedback exceptions cannot undo trusted owner acceptance', async () => {
  const h = host([owner('A')]);
  const receipt = await runCalculationGroup({
    ...h.args,
    progress: () => {
      throw Error('Detached feedback consumer');
    },
    onAccepted: () => {
      throw Error('Detached UI view');
    },
  });
  assert.equal(receipt.results[0].status, 'accepted');
  assert.equal(receipt.outcome, 'completed');
  assert.match(receipt.results[0].reason, /Accepted; host feedback failed/);
  assert.equal(h.slots.get('owner-A').commits.length, 1);
  return { receipt };
});
await check('Protected source canRecalculate false skips without authorization/apply while its qualified neighbor commits', async () => {
  const targets = [
      owner('A', { canRecalculate: true }),
      owner('protected', { canRecalculate: false, reason: 'Protected unqualified source calculation policy' }),
    ],
    h = host(targets),
    before = structuredClone(h.slots.get('owner-protected').latest),
    receipt = await runCalculationGroup(h.args);
  assert.equal(receipt.outcome, 'completed');
  assert.deepEqual(
    receipt.results.map((r) => r.status),
    ['accepted', 'skipped']
  );
  assert.equal(receipt.results[1].reason, targets[1].reason);
  assert.deepEqual(h.authorization, ['owner-A']);
  assert.deepEqual(h.calls, [{ ownerId: 'owner-A', commands: [{ type: 'recalculate' }] }]);
  assert.deepEqual(h.slots.get('owner-protected').latest, before);
  assert.equal(h.slots.get('owner-protected').commits.length, 0);
  return {
    receipt,
    protectedBefore: before,
    protectedAfter: h.slots.get('owner-protected').latest,
    noProtectedAuthorizationOrApply: true,
  };
});
await check('A late stop after the only trusted owner accepted reports completed, not cancelled', async () => {
  const h = host([owner('A')]),
    stop = new AbortController(),
    receipt = await runCalculationGroup({
      ...h.args,
      signal: stop.signal,
      apply: async (target, commands) => {
        const reply = await h.apply(target, commands);
        stop.abort();
        return reply;
      },
    });
  assert.equal(receipt.outcome, 'completed');
  assert.deepEqual(
    receipt.results.map((row) => row.status),
    ['accepted']
  );
  assert.equal(receipt.results[0].afterState, h.slots.get('owner-A').latest.state);
  assert.equal(h.slots.get('owner-A').commits.length, 1);
  assert.equal(h.progress.at(-1).phase, 'completed');
  return { receipt, lateStop: true };
});
await check('Stop requested with an unrelated current failure keeps failed and prevents later owner calls', async () => {
  const h = host([owner('A'), owner('B')]),
    stop = new AbortController(),
    receipt = await runCalculationGroup({
      ...h.args,
      signal: stop.signal,
      apply: async () => {
        stop.abort();
        throw Error('Independent native calculation failure while stopping');
      },
    });
  assert.equal(receipt.outcome, 'cancelled');
  assert.deepEqual(
    receipt.results.map((row) => row.status),
    ['failed', 'notRun']
  );
  assert.equal(receipt.results[0].reason, 'Independent native calculation failure while stopping');
  assert.deepEqual(h.calls, []);
  for (const slot of h.slots.values()) assert.equal(slot.commits.length, 0);
  assert.equal(h.progress.at(-1).phase, 'cancelled');
  return { receipt, noAuthoritativeCurrentCancellation: true };
});
