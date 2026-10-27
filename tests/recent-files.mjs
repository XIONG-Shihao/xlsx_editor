import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, readFile, realpath, rm, symlink, writeFile } from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { randomUUID } from 'node:crypto';
import { RecentFiles } from '../standalone/recent-files.mjs';

async function fixture(t) {
  const root = await realpath(await mkdtemp(path.join(os.tmpdir(), 'xlsx-recent-files-')));
  t.after(() => rm(root, { recursive: true, force: true }));
  const filename = path.join(root, 'profile', 'recent-files.json');
  async function file(name) {
    const target = path.join(root, name);
    await writeFile(target, 'fixture');
    return target;
  }
  return { root, filename, file, open: () => RecentFiles.open(filename) };
}

test('successful opens persist canonical paths, deduplicate aliases and retain opaque IDs', async (t) => {
  const f = await fixture(t),
    a = await f.file('a.xlsx'),
    b = await f.file('b.csv'),
    alias = path.join(f.root, 'alias.xlsx'),
    store = await f.open();
  await symlink(a, alias);
  await store.remember(a);
  const first = (await store.list()).entries[0];
  assert.notEqual(first.id, a);
  await store.remember(b);
  await store.remember(alias);
  const entries = (await store.list()).entries;
  assert.deepEqual(
    entries.map((entry) => entry.displayPath),
    [a, b]
  );
  assert.equal(entries[0].id, first.id);
  assert(entries.every((entry) => entry.available));
  const restarted = await f.open();
  assert.deepEqual((await restarted.list()).entries, entries);
  assert.equal(await restarted.resolve(first.id), a);
  await assert.rejects(restarted.resolve(a), /not granted/);
  await assert.rejects(restarted.resolve(randomUUID()), /not granted/);
});

test('concurrent remembers serialize without lost records and bound durable history to 20', async (t) => {
  const f = await fixture(t),
    store = await f.open(),
    files = await Promise.all(Array.from({ length: 23 }, (_, index) => f.file(index + '.tsv')));
  await store.remember(files[0]);
  const evicted = (await store.list()).entries[0].id;
  await Promise.all(files.slice(1).map((file) => store.remember(file)));
  const entries = (await store.list()).entries;
  assert.equal(entries.length, 20);
  assert.deepEqual(
    entries.map((entry) => entry.displayPath),
    files.slice(3).reverse()
  );
  const persisted = JSON.parse(await readFile(f.filename, 'utf8'));
  assert.equal(persisted.entries.length, 20);
  assert.equal(persisted.schema, 1);
  assert.deepEqual((await (await f.open()).list()).entries, entries);
  await assert.rejects(store.resolve(evicted), /not granted/);
});

test('missing entries stay listed until explicit removal; stale IDs never reopen or remove another record', async (t) => {
  const f = await fixture(t),
    a = await f.file('missing.xlsx'),
    b = await f.file('kept.xlsx'),
    store = await f.open();
  await store.remember(a);
  await store.remember(b);
  const entries = (await store.list()).entries,
    id = entries[1].id;
  const recovery = path.join(f.root, 'recovery.json');
  await writeFile(recovery, 'accepted recovery');
  await rm(a);
  const missing = (await store.list()).entries.find((entry) => entry.id === id);
  assert.equal(missing.available, false);
  assert.match(missing.reason, /missing/);
  await assert.rejects(store.resolve(id), /current spreadsheet was kept open/);
  assert.equal((await store.list()).entries.length, 2);
  await store.remove(id);
  assert.equal((await store.list()).entries.length, 1);
  await assert.rejects(store.resolve(id), /not granted/);
  await assert.rejects(store.remove(id), /not granted/);
  assert.equal(await readFile(b, 'utf8'), 'fixture');
  assert.equal(await readFile(recovery, 'utf8'), 'accepted recovery');
  assert.equal((await (await f.open()).list()).entries[0].displayPath, b);
});

test('a remembered canonical path replaced by a symlink cannot grant its new target', async (t) => {
  const f = await fixture(t),
    a = await f.file('original.xlsx'),
    other = await f.file('other.xlsx'),
    store = await f.open();
  await store.remember(a);
  const id = (await store.list()).entries[0].id;
  await rm(a);
  await symlink(other, a);
  assert.equal((await store.list()).entries[0].available, false);
  await assert.rejects(store.resolve(id), /another file/);
  assert.equal(await readFile(other, 'utf8'), 'fixture');
});

for (const content of [
  '{broken',
  JSON.stringify({ schema: 2, entries: [] }),
  JSON.stringify({ schema: 1, entries: [{ id: randomUUID(), filename: '/etc/passwd', lastOpened: 1 }] }),
])
  test('corrupt/unsupported history is reported and preserved: ' + content.slice(0, 35), async (t) => {
    const f = await fixture(t),
      a = await f.file('valid.xlsx');
    await mkdir(path.dirname(f.filename));
    await writeFile(f.filename, content);
    const store = await f.open(),
      snapshot = await store.list();
    assert.equal(snapshot.entries.length, 0);
    assert.match(snapshot.warning, /existing file is preserved/);
    await assert.rejects(store.remember(a), /preserved/);
    assert.equal(await readFile(f.filename, 'utf8'), content);
  });

test('external registry changes are not silently overwritten by an open or removal', async (t) => {
  const f = await fixture(t),
    a = await f.file('a.xlsx'),
    b = await f.file('b.xlsx'),
    store = await f.open();
  await store.remember(a);
  const id = (await store.list()).entries[0].id,
    external = '{external change';
  await writeFile(f.filename, external);
  await assert.rejects(store.remember(b), /changed on disk/);
  await assert.rejects(store.remove(id), /changed on disk/);
  assert.equal(await readFile(f.filename, 'utf8'), external);
  assert.equal((await store.list()).entries[0].id, id);
  assert.match((await store.list()).warning, /retained/);
});

test('oversized registry cannot consume an unbounded load or be overwritten', async (t) => {
  const f = await fixture(t),
    a = await f.file('valid.xlsx');
  await mkdir(path.dirname(f.filename));
  const content = 'x'.repeat(256 * 1024 + 1);
  await writeFile(f.filename, content);
  const store = await f.open();
  assert.match((await store.list()).warning, /size/);
  await assert.rejects(store.remember(a), /preserved/);
  assert.equal(await readFile(f.filename, 'utf8'), content);
});
