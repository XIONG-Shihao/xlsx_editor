/** Real grant/snapshot/inspection/staging services with a controlled owner.
 * Native history/Save and actual app qualification remain separate. */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { LocalQueryRefreshAll, REFRESH_ALL_LIMITS } from '../adapters/local-query-refresh-all.mjs';
import { createControlledQueryOwner } from './support/local-query-controlled-owner.mjs';
const textQuery = { columns: [], filters: [], keepHeader: false };
async function fixture(t, { count = 3, now } = {}) {
  const directory = await fs.realpath(await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-refresh-all-')));
  const owner = await createControlledQueryOwner(t, directory, {
      text: '01\n=1+2\n',
      query: textQuery,
      target: { sheet: 'S', start: 'A1' },
    }),
    files = [owner.filename],
    records = [owner.record];
  for (let n = 1; n < count; n++) {
    const file = path.join(owner.root, 'source-' + n + '.csv');
    await fs.writeFile(file, String(10 + n) + '\n');
    files.push(file);
    records.push(
      await owner.store.create(
        'doc',
        {
          name: 'Connection ' + (n + 1),
          target: { sheet: 'S', start: String.fromCharCode(65 + n * 3) + '1' },
          query: textQuery,
        },
        file
      )
    );
  }
  const host = new LocalQueryRefreshAll({
    store: owner.store,
    document: 'doc',
    directory: owner.staging,
    context: owner.context,
    ...(now ? { now } : {}),
  });
  t.after(() => host.dispose());
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  return { ...owner, files, records, all: host };
}
test('all connections preview without a write; three accepted units chain only their own heads and preserve literal types', async (t) => {
  const f = await fixture(t),
    { refreshAllPreview: p } = await f.all.preview();
  assert.equal(p.entries.length, 3);
  assert(p.entries.every((e) => e.status === 'ready'));
  assert.equal(f.accepted(), 0);
  assert.equal(f.scope.state, 'head');
  assert(!JSON.stringify(p).includes(f.root));
  // Public mutations cannot change the private stored target or source proof.
  p.entries[0].target.start = 'Z99';
  p.entries[1].revision = 999;
  const { refreshAllResult: r } = await f.all.refresh(p.id);
  assert.equal(r.accepted, 3);
  assert.equal(r.undoUnits, 3);
  assert.equal(f.accepted(), 3);
  assert.deepEqual(
    r.outcomes.map((e) => [e.parentState, e.state]),
    [
      ['head', 'accepted-1'],
      ['accepted-1', 'accepted-2'],
      ['accepted-2', 'accepted-3'],
    ]
  );
  assert.equal(f.cells.A1.value, '01');
  assert.equal(f.cells.A2.value, '=1+2');
  assert.equal(f.cells.D1.value, '11');
  assert.equal(f.cells.G1.value, '12');
  assert.equal(f.cells.Z99, undefined);
  await assert.rejects(f.all.refresh(p.id), { code: 'query-stale' });
});
test('changed source/settings, manual managed edits and a missing source refuse visibly while independent units continue', async (t) => {
  const f = await fixture(t);
  let { refreshAllPreview: p } = await f.all.preview();
  await fs.writeFile(f.files[1], 'changed\n');
  const { refreshAllResult: r } = await f.all.refresh(p.id);
  assert.equal(r.accepted, 2);
  assert.equal(r.outcomes[1].status, 'refused');
  assert.match(r.outcomes[1].reason, /changed|Preview/);
  assert.equal(f.cells.D1, undefined);
  assert.equal(f.cells.G1.value, '12');
  f.cells.A1.value = 'manual';
  await fs.unlink(f.files[2]);
  p = (await f.all.preview()).refreshAllPreview;
  assert.equal(p.entries[0].status, 'refused');
  assert.match(p.entries[0].reason, /output changed/);
  assert.equal(p.entries[2].status, 'refused');
  assert.match(p.entries[2].reason, /ENOENT/);
  assert(!p.entries[2].reason.includes(f.root));
  const fresh = await f.store.granted('doc', f.records[1].id);
  p = (await f.all.preview()).refreshAllPreview;
  await f.store.update('doc', fresh.id, fresh.revision, {
    name: 'Retargeted',
    target: { sheet: 'S', start: 'J1' },
    query: textQuery,
  });
  const changed = (await f.all.refresh(p.id)).refreshAllResult;
  assert.equal(changed.accepted, 0);
  assert.equal(changed.outcomes[1].status, 'refused');
  assert.equal(f.cells.J1, undefined);
});
test('cancellation after the first accepted unit reports that prefix and runs no later unit', async (t) => {
  const f = await fixture(t),
    controller = new AbortController(),
    p = (await f.all.preview()).refreshAllPreview;
  const r = (await f.all.refresh(p.id, { signal: controller.signal, afterApply: () => controller.abort() }))
    .refreshAllResult;
  assert.equal(r.accepted, 1);
  assert.equal(r.undoUnits, 1);
  assert(r.cancelled);
  assert.equal(r.outcomes[0].status, 'refreshed');
  assert.equal(r.outcomes[1].status, 'cancelled');
  assert.equal(r.outcomes[2].status, 'not-run');
  assert.equal(f.accepted(), 1);
  assert.equal((await f.store.list('doc')).entries[0].pending, undefined);
  assert.equal(f.cells.D1, undefined);
});
test('foreign owner/head retirement prevents execution and does not consume a valid plan for another scope', async (t) => {
  for (const field of ['state', 'ownerId', 'documentId']) {
    const f = await fixture(t, { count: 2 }),
      p = (await f.all.preview()).refreshAllPreview,
      old = f.scope[field];
    f.scope[field] = 'replacement';
    await assert.rejects(f.all.refresh(p.id), /owner|head/);
    assert.equal(f.accepted(), 0);
    f.scope[field] = old;
    const r = (await f.all.refresh(p.id)).refreshAllResult;
    assert.equal(r.accepted, 2);
  }
});
test('overlapping outputs, protected/occupied new targets and rejected conversions are present as refused units', async (t) => {
  const f = await fixture(t),
    third = f.records[2];
  await f.store.update('doc', third.id, third.revision, {
    name: 'Overlap',
    target: { sheet: 'S', start: 'A2' },
    query: textQuery,
  });
  f.cells.D1 = { value: 'Existing' };
  const p = (await f.all.preview()).refreshAllPreview;
  assert(p.entries.every((e) => e.status === 'refused'));
  assert.match(p.entries[0].reason, /overlap/);
  assert.match(p.entries[1].reason, /occupied/);
  assert.equal((await f.all.refresh(p.id)).refreshAllResult.accepted, 0);
  const typed = await fixture(t, { count: 2 });
  await typed.store.update('doc', typed.records[0].id, 1, {
    name: 'Typed reject',
    target: { sheet: 'S', start: 'A1' },
    query: { ...textQuery, policies: ['number'] },
  });
  const q = (await typed.all.preview()).refreshAllPreview;
  assert.equal(q.entries[0].status, 'refused');
  assert(q.entries[0].conversions.rejected > 0);
  assert.equal(q.entries[1].status, 'ready');
  assert.equal((await typed.all.refresh(q.id)).refreshAllResult.accepted, 1);
});
test('32-connection metadata plans survive the independent four-receipt spool limit; preview samples remain bounded', async (t) => {
  // Use spaced columns within the actual worksheet limits, rather than ASCII after Z.
  const f = await fixture(t, { count: 1 });
  for (let n = 1; n < 32; n++) {
    const file = path.join(f.root, 'many-' + n + '.csv');
    await fs.writeFile(file, String(n) + '\n');
    await f.store.create(
      'doc',
      { name: 'Many ' + n, target: { sheet: 'S', start: 'A' + (n * 3 + 1) }, query: textQuery },
      file
    );
  }
  const p = (await f.all.preview()).refreshAllPreview;
  assert.equal(p.entries.length, 32);
  assert(p.entries.every((e) => e.status === 'ready'));
  assert(JSON.stringify(p).length < REFRESH_ALL_LIMITS.responseBytes);
  const r = (await f.all.refresh(p.id)).refreshAllResult;
  assert.equal(r.accepted, 32);
  assert.equal(r.undoUnits, 32);
  assert.equal(f.cells.A94.value, '31');
});
test('one physical operation prevents duplicate calls and lost acceptance remains uncertain with its original intent', async (t) => {
  const f = await fixture(t, { count: 2 }),
    p = (await f.all.preview()).refreshAllPreview;
  let release;
  const gate = new Promise((r) => (release = r)),
    original = f.context.apply;
  f.context.apply = async (...args) => {
    await gate;
    return original(...args);
  };
  const first = f.all.refresh(p.id);
  await assert.rejects(f.all.refresh(p.id), { code: 'query-bad-input' });
  await new Promise((r) => setTimeout(r, 10));
  await assert.rejects(f.all.refresh(p.id), { code: 'query-bad-input' });
  await assert.rejects(f.all.preview(), { code: 'query-bad-input' });
  release();
  assert.equal((await first).refreshAllResult.accepted, 2);
  assert.equal(f.accepted(), 2);
  const lost = await fixture(t, { count: 2 }),
    q = (await lost.all.preview()).refreshAllPreview,
    apply = lost.context.apply;
  lost.context.apply = async (...args) => {
    await apply(...args);
    throw Error('Lost after acceptance');
  };
  const result = (await lost.all.refresh(q.id)).refreshAllResult;
  assert.equal(result.outcomes[0].status, 'uncertain');
  assert.equal(result.outcomes[0].undoUnits, null);
  assert.equal(result.undoUnits, null);
  assert.equal(result.uncertain, 1);
  assert.equal(result.outcomes[1].status, 'not-run');
  assert.equal(lost.accepted(), 1);
  assert((await lost.store.list('doc')).entries[0].pending);
  await assert.rejects(lost.all.refresh(q.id), { code: 'query-stale' });
});
test('expired/disposed plans and a cancelled preview never grant execution', async (t) => {
  let now = 0;
  const f = await fixture(t, { count: 2, now: () => now }),
    p = (await f.all.preview()).refreshAllPreview;
  now = REFRESH_ALL_LIMITS.lifetimeMs + 1;
  await assert.rejects(f.all.refresh(p.id), { code: /^(object-stale|query-stale|session-stale)$/ });
  const controller = new AbortController();
  controller.abort();
  await assert.rejects(f.all.preview({ signal: controller.signal }));
  assert.equal(f.accepted(), 0);
  f.all.dispose();
  await assert.rejects(f.all.preview(), /retired/);
});

test('new or removed connection membership retires the complete plan before a write', async (t) => {
  const f = await fixture(t, { count: 2 }),
    p = (await f.all.preview()).refreshAllPreview;
  const added = await f.store.create(
    'doc',
    { name: 'New connection', target: { sheet: 'S', start: 'J1' }, query: textQuery },
    f.filename
  );
  await assert.rejects(f.all.refresh(p.id), { code: 'query-stale' });
  assert.equal(f.accepted(), 0);
  await f.store.remove('doc', added.id, added.revision);
  const q = (await f.all.preview()).refreshAllPreview;
  await f.store.remove('doc', f.records[1].id, 1);
  await assert.rejects(f.all.refresh(q.id), { code: 'query-stale' });
  assert.equal(f.accepted(), 0);
});
test('a verified native acceptance stays acknowledged if later finalization fails, and stops later units', async (t) => {
  const f = await fixture(t, { count: 2 }),
    p = (await f.all.preview()).refreshAllPreview;
  const r = (
    await f.all.refresh(p.id, {
      afterApply: () => {
        throw Error('Finalization callback failed');
      },
    })
  ).refreshAllResult;
  assert.equal(r.accepted, 1);
  assert.equal(r.uncertain, 0);
  assert.equal(r.undoUnits, 1);
  assert.equal(r.outcomes[0].status, 'refreshed');
  assert.match(r.outcomes[0].warning, /acknowledged|reconciliation/);
  assert.equal(r.outcomes[1].status, 'not-run');
  assert.equal(f.accepted(), 1);
});
