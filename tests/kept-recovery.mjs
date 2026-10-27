import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createHash } from 'node:crypto';
import { KeptRecoveryBindings } from '../adapters/kept-recovery.mjs';
import { RecentFiles } from '../standalone/recent-files.mjs';
const hash = (filename) => createHash('sha256').update(filename).digest('hex');
async function fixture(t) {
  const root = await fs.realpath(await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-kept-bindings-'))),
    file = path.join(root, 'kept-recovery.json');
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  return { root, file, bindings: await KeptRecoveryBindings.open(file) };
}

test('missing-source binding survives restart and idempotent pin preserves creation time and exact bytes', async (t) => {
  const f = await fixture(t),
    filename = path.join(f.root, 'missing.xlsx'),
    id = hash(filename);
  const first = await f.bindings.pin(filename, id),
    bytes = await fs.readFile(f.file);
  assert.equal(first.id, id);
  assert(Number.isSafeInteger(first.createdAt));
  await assert.rejects(fs.stat(filename), { code: 'ENOENT' });
  assert.deepEqual(await f.bindings.pin(filename, id), first);
  assert.deepEqual(await fs.readFile(f.file), bytes);
  const reopened = await KeptRecoveryBindings.open(f.file);
  assert.deepEqual(await reopened.list(), { entries: [first] });
  const returned = await reopened.list();
  returned.entries[0].filename = 'mutated';
  assert.equal((await reopened.list()).entries[0].filename, filename);
});

test('ordinary recent removal and21-file eviction never remove a separately kept legacy grant', async (t) => {
  const f = await fixture(t),
    recent = await RecentFiles.open(path.join(f.root, 'recent.json')),
    source = path.join(f.root, 'source.xlsx');
  await fs.writeFile(source, 'placeholder');
  await recent.remember(source);
  await f.bindings.pin(source, hash(source));
  const old = (await recent.list()).entries[0];
  await fs.unlink(source);
  await recent.remove(old.id);
  for (let i = 0; i < 21; i++) {
    const name = path.join(f.root, `other-${i}.xlsx`);
    await fs.writeFile(name, 'placeholder');
    await recent.remember(name);
  }
  assert.equal((await recent.list()).entries.length, 20);
  assert(!(await recent.list()).entries.some((value) => value.displayPath === source));
  assert.equal((await (await KeptRecoveryBindings.open(f.file)).list()).entries[0].filename, source);
  await assert.rejects(fs.stat(source), { code: 'ENOENT' });
});

test('128-entry capacity refuses overflow without eviction and explicit removal frees only its own binding', async (t) => {
  const f = await fixture(t),
    files = Array.from({ length: 129 }, (_, i) => path.join(f.root, `missing-${i}.csv`));
  for (const file of files.slice(0, 128)) await f.bindings.pin(file, hash(file));
  const before = await fs.readFile(f.file);
  await assert.rejects(f.bindings.pin(files[128], hash(files[128])), { code: 'session-limit' });
  assert.deepEqual(await fs.readFile(f.file), before);
  assert.equal((await f.bindings.list()).entries.length, 128);
  assert.equal(await f.bindings.remove(hash(files[5])), true);
  assert.equal(await f.bindings.remove(hash(files[5])), false);
  await f.bindings.pin(files[128], hash(files[128]));
  const entries = (await f.bindings.list()).entries;
  assert.equal(entries.length, 128);
  assert(!entries.some((value) => value.filename === files[5]));
  assert(entries.some((value) => value.filename === files[0]));
});

test('encoded256KiB budget refuses large bound paths without losing already kept bindings', async (t) => {
  const f = await fixture(t);
  let accepted = 0,
    before;
  for (let i = 0; i < 128; i++) {
    const filename = path.join(f.root, String(i) + '界'.repeat(1500) + '.xlsx');
    before = await fs.readFile(f.file).catch((error) => {
      if (error.code === 'ENOENT') return null;
      throw error;
    });
    try {
      await f.bindings.pin(filename, hash(filename));
      accepted++;
    } catch (error) {
      assert.match(String(error), /256 KiB bound/);
      assert.deepEqual(await fs.readFile(f.file), before);
      break;
    }
  }
  assert(accepted > 0 && accepted < 128);
  assert.equal((await f.bindings.list()).entries.length, accepted);
  assert((await fs.stat(f.file)).size <= 256 * 1024);
  await f.bindings.pin(path.join(f.root, 'small.xlsx'), hash(path.join(f.root, 'small.xlsx')));
});

test('closed schema, exact filename hash and byte limits lock corrupt registries and preserve every byte', async (t) => {
  const f = await fixture(t),
    filename = path.join(f.root, 'old.tsv'),
    good = { id: hash(filename), filename, createdAt: 1 },
    values = [
      '{broken',
      JSON.stringify({ schema: 2, entries: [] }),
      JSON.stringify({ schema: 1, entries: [], extra: true }),
      JSON.stringify({ schema: 1, entries: [{ ...good, extra: true }] }),
      JSON.stringify({ schema: 1, entries: [{ ...good, id: '0'.repeat(64) }] }),
      JSON.stringify({ schema: 1, entries: [good, good] }),
      JSON.stringify({
        schema: 1,
        entries: Array.from({ length: 129 }, (_, i) => {
          const name = path.join(f.root, `${i}.xlsx`);
          return { id: hash(name), filename: name, createdAt: 1 };
        }),
      }),
      Buffer.alloc(256 * 1024 + 1, 32),
      Buffer.from([0xff]),
    ];
  for (const raw of values) {
    await fs.writeFile(f.file, raw);
    const original = await fs.readFile(f.file),
      bindings = await KeptRecoveryBindings.open(f.file);
    assert((await bindings.list()).warning);
    await assert.rejects(bindings.pin(filename, hash(filename)), /preserved|read|restart/);
    await assert.rejects(bindings.remove(hash(filename)), /preserved|read|restart/);
    assert.deepEqual(await fs.readFile(f.file), original);
  }
});

test('invalid caller bindings never write, and a registry symlink never writes through to an external file', async (t) => {
  const f = await fixture(t);
  for (const filename of [
    'relative.xlsx',
    path.join(f.root, 'wrong.txt'),
    path.join(f.root, 'nul\0.xlsx'),
    '/' + 'x'.repeat(4097) + '.xlsx',
  ])
    await assert.rejects(f.bindings.pin(filename, hash(filename)), /Invalid/);
  await assert.rejects(f.bindings.pin(path.join(f.root, 'valid.xlsx'), '0'.repeat(64)), /Invalid/);
  await assert.rejects(fs.stat(f.file), { code: 'ENOENT' });
  const target = path.join(f.root, 'external.json');
  await fs.writeFile(target, 'external sentinel');
  await fs.symlink(target, f.file);
  const blocked = await KeptRecoveryBindings.open(f.file);
  assert((await blocked.list()).warning);
  await assert.rejects(
    blocked.pin(path.join(f.root, 'ok.xlsx'), hash(path.join(f.root, 'ok.xlsx'))),
    /preserved|restart/
  );
  assert.equal(await fs.readFile(target, 'utf8'), 'external sentinel');
  assert((await fs.lstat(f.file)).isSymbolicLink());
});

test('stale raw JSON locks writes even on idempotent pin, without normalizing or replacing new disk bytes', async (t) => {
  const f = await fixture(t),
    filename = path.join(f.root, 'kept.xlsx');
  await f.bindings.pin(filename, hash(filename));
  const before = await fs.readFile(f.file, 'utf8'),
    changed = before + '\n';
  await fs.writeFile(f.file, changed);
  await assert.rejects(f.bindings.pin(filename, hash(filename)), /changed|restart/);
  assert((await f.bindings.list()).warning);
  await fs.writeFile(f.file, before);
  await assert.rejects(f.bindings.remove(hash(filename)), /restart/);
  assert.equal(await fs.readFile(f.file, 'utf8'), before);
});

test('competing registry instances serialize before exact prior-byte comparison rather than overwrite a winning pin', async (t) => {
  const f = await fixture(t),
    other = await KeptRecoveryBindings.open(f.file),
    a = path.join(f.root, 'a.xlsx'),
    b = path.join(f.root, 'b.xlsx'),
    results = await Promise.allSettled([f.bindings.pin(a, hash(a)), other.pin(b, hash(b))]);
  assert.equal(results.filter((value) => value.status === 'fulfilled').length, 1);
  assert.equal(results.filter((value) => value.status === 'rejected').length, 1);
  const disk = await (await KeptRecoveryBindings.open(f.file)).list();
  assert.equal(disk.entries.length, 1);
  assert.equal(disk.entries[0].filename, a);
  assert((await other.list()).warning);
});
