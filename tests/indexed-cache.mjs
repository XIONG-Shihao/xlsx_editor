import assert from 'node:assert/strict';
import { test, after } from 'node:test';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { execFileSync } from 'node:child_process';
import { indexedCache, indexedCacheUsage, fileDigest } from '../core/indexed-cache.mjs';

const roots = [];
after(async () => {
  for (const root of roots) await fs.rm(root, { recursive: true, force: true });
});
async function fixture(extra = {}) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-cache-test-'));
  roots.push(root);
  const source = path.join(root, 'book.xlsx');
  await fs.writeFile(source, 'immutable workbook input');
  const index = path.join(root, 'index');
  await fs.mkdir(index);
  const revision = await fileDigest(source);
  for (const [name, bytes] of Object.entries({
    'source.xlsx': await fs.readFile(source),
    'index.json': JSON.stringify({ revision }),
    'cells.sqlite': 'closed database',
    'metadata.xlsx': 'bounded metadata',
    'sheet-0.xml': 'original worksheet bytes',
    ...extra,
  }))
    await fs.writeFile(path.join(index, name), bytes);
  const options = { directory: path.join(root, 'cache') };
  const target = async (name) => {
    const out = path.join(root, name);
    await fs.mkdir(out);
    return out;
  };
  return { root, index, source, revision, options, target };
}
async function entry(options) {
  const names = (await fs.readdir(options.directory)).filter((n) => /^[a-f0-9]{64}$/.test(n));
  assert.equal(names.length, 1);
  return path.join(options.directory, names[0]);
}

test('Linux user caches do not reuse shared legacy temp paths', { skip: process.platform !== 'linux' }, async () => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-shared-tmp-'));
  roots.push(root);
  for (const name of ['workbook-indexes-v1', 'workbook-cache-v1', 'spreadsheet-workers-v1'])
    await fs.writeFile(path.join(root, name), 'belongs to another account');
  const storage = new URL('../core/indexed-storage.mjs', import.meta.url).href;
  const cache = new URL('../core/indexed-cache.mjs', import.meta.url).href;
  const workers = new URL('../adapters/worker-storage.mjs', import.meta.url).href;
  execFileSync(
    process.execPath,
    [
      '--input-type=module',
      '-e',
      `
  const {createIndexedStorage}=await import(${JSON.stringify(storage)});
  const {indexedCacheUsage}=await import(${JSON.stringify(cache)});
  const {allocateWorkerStorage}=await import(${JSON.stringify(workers)});
  await createIndexedStorage('linux-user-isolation');
  await indexedCacheUsage();
  allocateWorkerStorage();
 `,
    ],
    { env: { ...process.env, TMPDIR: root }, timeout: 10000 }
  );
  for (const name of ['workbook-indexes-v1', 'workbook-cache-v1', 'spreadsheet-workers-v1']) {
    assert.equal(await fs.readFile(path.join(root, name), 'utf8'), 'belongs to another account');
    const stat = await fs.stat(path.join(root, `${name}-${process.getuid()}`));
    assert(stat.isDirectory());
    assert.equal(stat.uid, process.getuid());
    assert.equal(stat.mode & 0o077, 0);
  }
});

test('index snapshots verify every copied artifact and preserve independent mutable clones', async () => {
  const f = await fixture({
    'export.xlsx': 'old output',
    '.quota.lock': '',
    cancelled: 'stale marker',
    'sort-order.jsonl': 'transient',
  });
  const cache = await indexedCache(f.source, 'native-schema-policy-A', f.options);
  assert((await cache.publish(f.index, 0)).published);
  const target = await f.target('restored'),
    manifest = await cache.restore(target);
  assert.equal(manifest.revision, f.revision);
  assert.equal(manifest.reuseCalculation, true);
  assert.deepEqual((await fs.readdir(target)).sort(), [
    'cells.sqlite',
    'index.json',
    'metadata.xlsx',
    'sheet-0.xml',
    'source.xlsx',
  ]);
  await fs.writeFile(path.join(target, 'cells.sqlite'), 'private mutation');
  const other = await f.target('other');
  await cache.restore(other);
  assert.equal(await fs.readFile(path.join(other, 'cells.sqlite'), 'utf8'), 'closed database');
  assert.equal(await fileDigest(f.source), f.revision);
});

test('same-size artifact corruption and unexpected extra artifacts become safe cache misses', async () => {
  for (const mode of ['bytes', 'extra', 'manifest', 'missing', 'index-revision']) {
    const f = await fixture(),
      cache = await indexedCache(f.source, 'policy', f.options);
    await cache.publish(f.index, 0);
    const root = await entry(f.options);
    if (mode === 'bytes') await fs.writeFile(path.join(root, 'cells.sqlite'), 'broken database');
    if (mode === 'extra') await fs.writeFile(path.join(root, 'unlisted'), 'unexpected');
    if (mode === 'manifest') await fs.writeFile(path.join(root, 'cache.json'), '{incomplete');
    if (mode === 'missing') await fs.rm(path.join(root, 'sheet-0.xml'));
    if (mode === 'index-revision') {
      const data = JSON.parse(await fs.readFile(path.join(root, 'cache.json'), 'utf8'));
      const file = data.files.find((file) => file.name === 'index.json');
      await fs.writeFile(path.join(root, 'index.json'), JSON.stringify({ revision: '0'.repeat(64) }));
      file.sha256 = await fileDigest(path.join(root, 'index.json'));
      await fs.writeFile(path.join(root, 'cache.json'), JSON.stringify(data));
    }
    const target = await f.target('restored');
    assert.equal(await cache.restore(target), null, mode);
    assert.deepEqual(await fs.readdir(target), [], mode);
    assert.equal(await fileDigest(f.source), f.revision);
  }
});

test('output bytes and schema/calculation identities fence reuse; original source stays unchanged', async () => {
  const f = await fixture(),
    cache = await indexedCache(f.source, 'policy-A', f.options);
  await cache.publish(f.index, 0);
  const otherPolicy = await indexedCache(f.source, 'policy-B', f.options);
  assert.equal(await otherPolicy.restore(await f.target('new-policy')), null);
  await fs.writeFile(f.source, 'external workbook update');
  const external = await indexedCache(f.source, 'policy-A', f.options);
  assert.equal(await external.restore(await f.target('external')), null);
  assert.equal(await fs.readFile(path.join(f.index, 'source.xlsx'), 'utf8'), 'immutable workbook input');
});

test('symlink corruption misses without following or changing its target; nonempty destinations reject', async () => {
  const f = await fixture(),
    cache = await indexedCache(f.source, 'policy', f.options);
  await cache.publish(f.index, 0);
  const root = await entry(f.options);
  await fs.rm(path.join(root, 'metadata.xlsx'));
  await fs.symlink(f.source, path.join(root, 'metadata.xlsx'));
  assert.equal(await cache.restore(await f.target('restored')), null);
  assert.equal(await fileDigest(f.source), f.revision);
  const target = await f.target('nonempty');
  await fs.writeFile(path.join(target, 'user.txt'), 'keep');
  await assert.rejects(cache.restore(target), { code: 'session-unsupported' });
  assert.equal(await fs.readFile(path.join(target, 'user.txt'), 'utf8'), 'keep');
});

test('cache and session quotas include staging, manifests and compaction reserve', async () => {
  const f = await fixture({ 'sheet-1.xml': Buffer.alloc(160 * 1024) });
  f.options.maxBytes = 1500 * 1024;
  f.options.maxEntries = 2;
  for (let i = 0; i < 5; i++)
    assert((await (await indexedCache(f.source, `policy-${i}`, f.options)).publish(f.index, 0)).published);
  const usage = await indexedCacheUsage(f.options);
  assert.equal(usage.entries, 2);
  assert.equal(usage.staging, 0);
  assert(usage.bytes <= usage.maxBytes);
  const cache = await indexedCache(f.source, 'too-big', f.options);
  assert.equal(await cache.publish(f.index, 0, { workingBytes: 2 * 1024 * 1024 }), null);
  assert.equal((await indexedCacheUsage(f.options)).entries, 2);
  const existing = await indexedCache(f.source, 'policy-4', f.options),
    target = await f.target('quota');
  await assert.rejects(existing.restore(target, 128 * 1024), /disk budget/);
  assert.deepEqual(await fs.readdir(target), []);
  assert.equal((await indexedCacheUsage(f.options)).entries, 2);
});

test('interrupted checkpoint preparation never publishes an incomplete entry', async () => {
  const f = await fixture(),
    controller = new AbortController(),
    cache = await indexedCache(f.source, 'abort', { ...f.options, signal: controller.signal });
  await assert.rejects(
    cache.publish(f.index, 0, {
      prepareSnapshot: async () => {
        controller.abort(Error('cancelled cache publication'));
      },
    }),
    /cancelled/
  );
  const usage = await indexedCacheUsage(f.options);
  assert.equal(usage.entries, 0);
  assert.equal(usage.staging, 0);
  assert.equal(await fileDigest(f.source), f.revision);
});

test('formula cache reuse requires explicit proof, not the static-references hint alone', async () => {
  const f = await fixture({ 'dependency-plan.json': JSON.stringify({ scope: { mode: 'static-references' } }) });
  const unproved = await indexedCache(f.source, 'unproved', f.options);
  await unproved.publish(f.index, 7);
  assert.equal((await unproved.restore(await f.target('unproved'))).reuseCalculation, false);
  const proved = await indexedCache(f.source, 'proved', f.options);
  await proved.publish(f.index, 7, { reuseCalculation: true });
  assert.equal((await proved.restore(await f.target('proved'))).reuseCalculation, true);
});

test('explicit invalidation removes only its keyed entry', async () => {
  const f = await fixture(),
    a = await indexedCache(f.source, 'A', f.options),
    b = await indexedCache(f.source, 'B', f.options);
  await a.publish(f.index, 0);
  await b.publish(f.index, 0);
  await a.invalidate();
  assert.equal(await a.restore(await f.target('A')), null);
  assert(await b.restore(await f.target('B')));
  assert.equal((await indexedCacheUsage(f.options)).entries, 1);
});

test('a file changed after cache lookup is never opened from an earlier revision', async () => {
  const f = await fixture(),
    cache = await indexedCache(f.source, 'policy', f.options);
  await cache.publish(f.index, 0);
  await fs.writeFile(f.source, 'external mutation after lookup');
  const target = await f.target('changed');
  await assert.rejects(cache.restore(target), /changed while restoring/);
  assert.deepEqual(await fs.readdir(target), []);
  // The old immutable entry remains valid for other byte-identical copies.
  await fs.writeFile(f.source, 'immutable workbook input');
  assert(await cache.restore(target));
});

test('abandoned staging from a live owner still counts against the cache byte cap', async () => {
  const f = await fixture();
  f.options.maxBytes = 1500 * 1024;
  const cache = await indexedCache(f.source, 'policy', f.options);
  const staging = path.join(f.options.directory, `pending-${process.pid}-00000000-0000-0000-0000-000000000000`);
  await fs.mkdir(staging);
  await fs.writeFile(path.join(staging, 'cells.sqlite'), Buffer.alloc(600 * 1024));
  assert.equal(await cache.publish(f.index, 0), null);
  const usage = await indexedCacheUsage(f.options);
  assert.equal(usage.entries, 0);
  assert.equal(usage.staging, 1);
  assert(usage.bytes <= usage.maxBytes);
});

test('reused digest buffers remain isolated across concurrent readers, tails and cancellation', async () => {
  const { createHash } = await import('node:crypto'),
    f = await fixture(),
    files = [];
  for (let i = 0; i < 6; i++) {
    const data = Buffer.alloc(300000 + i * 193, i + 17),
      file = path.join(f.root, `hash-${i}`);
    await fs.writeFile(file, data);
    files.push({ file, expected: createHash('sha256').update(data).digest('hex') });
  }
  for (let round = 0; round < 3; round++)
    assert.deepEqual(
      await Promise.all(files.map(({ file }) => fileDigest(file))),
      files.map(({ expected }) => expected)
    );
  const controller = new AbortController();
  controller.abort(Error('cancel digest'));
  await assert.rejects(fileDigest(files[0].file, { signal: controller.signal }), /cancel digest/);
  assert.equal(await fileDigest(files[1].file), files[1].expected);
});

test('lowering a persistent cache limit evicts oversized earlier entries before reuse', async () => {
  const f = await fixture({ 'sheet-1.xml': Buffer.alloc(16 * 1024) }),
    original = await indexedCache(f.source, 'policy', f.options);
  await original.publish(f.index, 0);
  assert.equal((await indexedCacheUsage(f.options)).entries, 1);
  const reduced = { ...f.options, maxBytes: 4096 },
    cache = await indexedCache(f.source, 'policy', reduced);
  assert.equal(await cache.restore(await f.target('small-budget')), null);
  const usage = await indexedCacheUsage(reduced);
  assert.equal(usage.entries, 0);
  assert(usage.bytes <= 4096);
  assert.equal(await fileDigest(f.source), f.revision);
});
