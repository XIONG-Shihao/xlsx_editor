import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { LocalActions } from '../adapters/local-actions.mjs';
const recalc = (name) => ({ name, descriptor: { schema: 1, command: 'recalculate', parameters: {} } });
async function fixture(t) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'workbook-local-actions-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const filename = path.join(directory, 'actions.json');
  return { directory, filename };
}
test('actual private create/list/edit/delete/cold persistence never touches an unrelated workbook', async (t) => {
  const { directory, filename } = await fixture(t),
    workbook = path.join(directory, 'source.xlsx'),
    original = Buffer.from('independent opaque workbook bytes');
  await fs.writeFile(workbook, original);
  const store = await LocalActions.open(filename);
  assert.deepEqual(await store.list(), { schema: 1, scope: 'private-profile', revision: 0, entries: [] });
  assert.equal(await fs.stat(filename).catch(() => null), null);
  let result = await store.create(recalc('Recalculate'), 0),
    saved = result.entries[0];
  assert.equal(result.revision, 1);
  assert.equal(saved.revision, 1);
  assert.equal(saved.author, 'local-profile');
  assert.equal(typeof saved.createdAt, 'number');
  result.entries[0].name = 'Caller-mutated';
  assert.equal((await store.list()).entries[0].name, 'Recalculate');
  const cold = await LocalActions.open(filename);
  assert.deepEqual(await cold.list(), await store.list());
  result = await cold.update(
    saved.id,
    {
      name: 'Sort rows',
      descriptor: {
        schema: 1,
        command: 'sortRows',
        parameters: { sheet: 'Data', range: '$a$1:$b$3', keys: [{ column: 'b', direction: 'desc' }] },
      },
    },
    1
  );
  assert.equal(result.revision, 2);
  assert.equal(result.entries[0].revision, 2);
  assert.equal(result.entries[0].createdAt, saved.createdAt);
  assert.deepEqual(
    (await (await LocalActions.open(filename)).list()).entries[0].descriptor,
    result.entries[0].descriptor
  );
  result = await cold.remove(saved.id, 2);
  assert.deepEqual(result, { schema: 1, scope: 'private-profile', revision: 3, entries: [] });
  assert.deepEqual(await (await LocalActions.open(filename)).list(), result);
  assert.deepEqual(await fs.readFile(workbook), original);
  assert.deepEqual((await fs.readdir(directory)).sort(), ['actions.json', 'source.xlsx']);
});
test('same-instance queued stale revisions and two independent stores cannot overwrite current entries', async (t) => {
  const { filename } = await fixture(t),
    a = await LocalActions.open(filename),
    b = await LocalActions.open(filename);
  const requests = await Promise.allSettled([a.create(recalc('First'), 0), b.create(recalc('Second'), 0)]);
  assert.equal(requests.filter((r) => r.status === 'fulfilled').length, 1);
  assert.equal(requests.filter((r) => r.status === 'rejected').length, 1);
  const before = await fs.readFile(filename),
    current = await LocalActions.open(filename);
  await assert.rejects(current.create(recalc('Stale'), 0), /Stale/);
  const attempts = await Promise.allSettled([
    current.create(recalc('Next'), 1),
    current.create(recalc('Wrong same head'), 1),
  ]);
  assert.equal(attempts[0].status, 'fulfilled');
  assert.equal(attempts[1].status, 'rejected');
  const result = await current.list();
  assert.equal(result.revision, 2);
  assert.equal(result.entries.length, 2);
  assert(!Buffer.from(JSON.stringify(result)).equals(before));
});
test('external valid rewrite, replacement and deletion lock old snapshots without overwriting disk', async (t) => {
  const { filename } = await fixture(t);
  for (const mutation of ['write', 'replace', 'delete']) {
    const store = await LocalActions.open(filename),
      before = await store.list();
    if (before.warning) {
      await fs.rm(filename, { force: true });
      continue;
    }
    if (!before.entries.length) await store.create(recalc('Original'), before.revision);
    const old = await LocalActions.open(filename),
      saved = JSON.parse(await fs.readFile(filename));
    saved.revision++;
    saved.entries[0].name = 'External';
    const raw = Buffer.from(JSON.stringify(saved));
    if (mutation === 'write') await fs.writeFile(filename, raw);
    else if (mutation === 'replace') {
      const other = filename + '.external';
      await fs.writeFile(other, raw);
      await fs.rename(other, filename);
    } else await fs.unlink(filename);
    await assert.rejects(old.create(recalc('Must not overwrite'), 1), /changed on disk/);
    assert((await old.list()).warning);
    assert.deepEqual(await fs.readFile(filename).catch(() => null), mutation === 'delete' ? null : raw);
    await fs.rm(filename, { force: true });
  }
});
test('unsupported schema, actor/path injection, duplicate keys, invalid UTF8 and byte excess remain exact', async (t) => {
  const { filename } = await fixture(t),
    valid = { schema: 1, scope: 'private-profile', revision: 0, entries: [] };
  const cases = [
    Buffer.from(JSON.stringify({ ...valid, schema: 2 })),
    Buffer.from(JSON.stringify({ ...valid, actor: 'admin' })),
    Buffer.from(JSON.stringify({ ...valid, filename: '/tmp/other.json' })),
    Buffer.from('{"schema":2,"schema":1,"scope":"private-profile","revision":0,"entries":[]}'),
    Buffer.from([0xff, 0xfe]),
    Buffer.alloc(1024 * 1024 + 1, 32),
  ];
  for (const raw of cases) {
    await fs.writeFile(filename, raw);
    const store = await LocalActions.open(filename);
    assert((await store.list()).warning);
    await assert.rejects(store.create(recalc('No write'), 0), /preserved/);
    assert.deepEqual(await fs.readFile(filename), raw);
  }
});
test('final symlink and hardlinked stores refuse; exclusive writer lock is not removed by another instance', async (t) => {
  const { directory, filename } = await fixture(t),
    target = path.join(directory, 'other.json'),
    raw = Buffer.from('{"schema":1,"scope":"private-profile","revision":0,"entries":[]}');
  await fs.writeFile(target, raw);
  await fs.symlink(target, filename);
  let store = await LocalActions.open(filename);
  assert((await store.list()).warning);
  await assert.rejects(store.create(recalc('No write'), 0));
  assert.deepEqual(await fs.readFile(target), raw);
  await fs.unlink(filename);
  await fs.link(target, filename);
  store = await LocalActions.open(filename);
  assert((await store.list()).warning);
  await assert.rejects(store.create(recalc('No write'), 0));
  assert.deepEqual(await fs.readFile(target), raw);
  await fs.unlink(filename);
  store = await LocalActions.open(filename);
  await fs.writeFile(filename + '.lock', 'Independent writer');
  await assert.rejects(store.create(recalc('No write'), 0), { code: 'session-busy' });
  assert.equal(await fs.readFile(filename + '.lock', 'utf8'), 'Independent writer');
  assert.equal(await fs.stat(filename).catch(() => null), null);
});
test('128 entries are retained without eviction and malformed mutation arguments do not publish', async (t) => {
  const { filename } = await fixture(t),
    store = await LocalActions.open(filename);
  for (let i = 0; i < 128; i++) await store.create(recalc('Action ' + i), i);
  const before = await fs.readFile(filename),
    result = await store.list();
  assert.equal(result.entries.length, 128);
  assert.equal(result.revision, 128);
  await assert.rejects(store.create(recalc('Overflow'), 128), { code: 'query-limit' });
  assert.throws(() => store.create({ ...recalc('Forged'), actor: 'local-profile' }, 128));
  assert.throws(() => store.update('/tmp/id', recalc('Forged'), 128));
  assert.throws(() => store.remove(result.entries[0].id, '128'));
  assert.deepEqual(await fs.readFile(filename), before);
  assert.deepEqual(await (await LocalActions.open(filename)).list(), result);
});
