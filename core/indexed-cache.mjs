import { formatHashSalt, FORMAT, formatTemplate } from './format.mjs';
/** Immutable, content-addressed native index snapshots. Cached data is disposable:
 * every artifact is verified before a private copy is accepted by a session. */
import fs from 'node:fs/promises';
import { constants } from 'node:fs';
import { createHash, randomUUID } from 'node:crypto';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { refusal } from './refusal-codes.mjs';

// Keep each Linux account's private cache separate within the shared /tmp.
const base = path.join(tmpdir(), `workbook-cache-v1${process.platform === 'linux' ? `-${process.getuid()}` : ''}`);
const SCHEMA = 2,
  MAX_BYTES = 2 * 1024 ** 3,
  MAX_ENTRIES = 8,
  MAX_MANIFEST = 1024 * 1024;
const hashPattern = /^[a-f0-9]{64}$/;
const durable = (name) =>
  [
    'source.xlsx',
    'cells.sqlite',
    'metadata.xlsx',
    'index.json',
    'display-formats.json',
    'dependency-plan.json',
  ].includes(name) || /^sheet-\d+\.xml$/.test(name);
const artifact = (name) => durable(name) || ['saved.xlsx', 'checkpoint.json'].includes(name);
const cancelled = (signal) => signal?.throwIfAborted();
// Consecutive hashes reuse storage rather than leaving a new stream/buffer
// allocation for GC each time. Concurrent hashes never share a leased buffer.
const digestBuffers = [];
function settings(options = {}) {
  const directory = options.directory === undefined ? base : path.resolve(options.directory);
  const maxBytes = options.maxBytes ?? MAX_BYTES,
    maxEntries = options.maxEntries ?? MAX_ENTRIES;
  if (
    !Number.isSafeInteger(maxBytes) ||
    maxBytes < 4096 ||
    maxBytes > MAX_BYTES ||
    !Number.isSafeInteger(maxEntries) ||
    maxEntries < 1 ||
    maxEntries > MAX_ENTRIES
  )
    throw refusal('internal-invalid-field', 'Invalid workbook cache limits');
  return { directory, maxBytes, maxEntries, signal: options.signal };
}
export async function fileDigest(file, { signal } = {}) {
  const hash = createHash('sha256'),
    handle = await fs.open(file, 'r'),
    buffer = digestBuffers.pop() ?? Buffer.allocUnsafe(256 * 1024);
  try {
    for (;;) {
      cancelled(signal);
      const { bytesRead } = await handle.read(buffer, 0, buffer.length, null);
      if (!bytesRead) break;
      hash.update(buffer.subarray(0, bytesRead));
    }
    cancelled(signal);
    return hash.digest('hex');
  } finally {
    try {
      await handle.close();
    } finally {
      if (digestBuffers.length < 2) digestBuffers.push(buffer);
    }
  }
}
/** Hash the actual executing native program, not a build label or stale sidecar.
 * Callers using a custom transport must supply the corresponding implementation
 * identity themselves. Resolving manifests does not load the formula runtime. */
export async function indexedImplementationIdentity(binary) {
  const hash = createHash('sha256').update(
    formatTemplate([formatHashSalt('XLSX_INDEX_POLICY'), FORMAT.FORMAT], [SCHEMA])
  );
  hash.update(await sourceIdentity(binary));
  hash.update(
    JSON.stringify({
      runtime: process.versions.bun
        ? { kind: 'bun', version: process.versions.bun }
        : { kind: 'node', version: process.versions.node, v8: process.versions.v8 },
      platform: process.platform,
      arch: process.arch,
    })
  );
  const core = fileURLToPath(new URL('./', import.meta.url));
  for (const name of (await fs.readdir(core)).filter((name) => name.endsWith('.mjs')).sort())
    hash.update(name).update(await fs.readFile(path.join(core, name)));
  const require = createRequire(import.meta.url);
  for (const name of [
    '@univerjs/core',
    '@univerjs/engine-formula',
    '@univerjs/sheets',
    '@univerjs/sheets-formula',
    '@univerjs/sheets-table',
    '@xmldom/xmldom',
    'fflate',
    'rxjs',
    'decimal.js',
  ]) {
    const resolver = name === 'decimal.js' ? createRequire(require.resolve('@univerjs/engine-formula')) : require;
    let directory = path.dirname(resolver.resolve(name)),
      manifest;
    for (;;) {
      try {
        const candidate = JSON.parse(await fs.readFile(path.join(directory, 'package.json'), 'utf8'));
        if (candidate.name === name) {
          manifest = candidate;
          break;
        }
      } catch (error) {
        if (error.code !== 'ENOENT') throw error;
      }
      const parent = path.dirname(directory);
      if (parent === directory)
        throw refusal('session-unsupported', `Workbook cache dependency identity unavailable: ${name}`);
      directory = parent;
    }
    hash.update(name).update(JSON.stringify(manifest));
  }
  return hash.digest('hex');
}
async function regular(file) {
  const stat = await fs.lstat(file);
  if (!stat.isFile() || stat.isSymbolicLink()) throw refusal('session-unsupported', 'Unsafe cached index artifact');
  return stat;
}
async function readJson(file, maxBytes = MAX_MANIFEST) {
  if ((await regular(file)).size > maxBytes) throw refusal('session-limit', 'Oversized workbook cache metadata');
  return JSON.parse(await fs.readFile(file, 'utf8'));
}
async function directory(parent) {
  await fs.mkdir(parent, { recursive: true, mode: 0o700 });
  const stat = await fs.lstat(parent);
  if (
    !stat.isDirectory() ||
    stat.isSymbolicLink() ||
    (process.getuid && stat.uid !== process.getuid()) ||
    stat.mode & 0o077
  )
    throw refusal('session-unsupported', 'Unsafe workbook cache directory');
  for (const item of await fs.readdir(parent, { withFileTypes: true })) {
    const match = /^pending-(\d+)-[a-f0-9-]+$/.exec(item.name);
    if (!match || !item.isDirectory()) continue;
    try {
      process.kill(Number(match[1]), 0);
    } catch (error) {
      if (error.code === 'ESRCH') await fs.rm(path.join(parent, item.name), { recursive: true, force: true });
    }
  }
  return parent;
}
async function lock(parent, signal, run) {
  const file = path.join(parent, '.publish-lock');
  let held = false;
  for (let i = 0; i < 1500; i++) {
    cancelled(signal);
    try {
      await fs.mkdir(file, { mode: 0o700 });
      held = true;
      await fs.writeFile(path.join(file, 'owner.json'), JSON.stringify({ pid: process.pid }));
      break;
    } catch (error) {
      if (error.code !== 'EEXIST') throw error;
      try {
        const owner = await readJson(path.join(file, 'owner.json'), 4096);
        if (!Number.isSafeInteger(owner.pid) || owner.pid <= 0)
          throw refusal('internal-invalid-field', 'Invalid workbook cache lock');
        try {
          process.kill(owner.pid, 0);
        } catch (error) {
          if (error.code === 'ESRCH') await fs.rm(file, { recursive: true, force: true });
        }
      } catch (error) {
        if (error.code === 'ENOENT') {
          const stat = await fs.lstat(file).catch(() => null);
          if (stat && Date.now() - stat.mtimeMs > 30000 && !stat.isSymbolicLink())
            await fs.rm(file, { recursive: true, force: true });
        } else throw error;
      }
      await new Promise((resolve) => setTimeout(resolve, 20));
    }
  }
  if (!held) throw refusal('session-busy', 'Workbook cache is busy');
  try {
    cancelled(signal);
    return await run();
  } finally {
    await fs.rm(file, { recursive: true, force: true });
  }
}
async function filesIn(root, filter = () => true) {
  const files = [];
  for (const entry of await fs.readdir(root, { withFileTypes: true })) {
    if (!filter(entry.name)) continue;
    const stat = await regular(path.join(root, entry.name));
    files.push({ name: entry.name, bytes: stat.size });
  }
  return files.sort((a, b) => a.name.localeCompare(b.name));
}
async function artifactBytes(root) {
  return (await filesIn(root)).reduce((n, file) => n + file.bytes, 0);
}
async function trimToLimits(parent, maxBytes, maxEntries) {
  const entries = [];
  let occupied = 0;
  for (const item of await fs.readdir(parent, { withFileTypes: true }))
    if (item.isDirectory()) {
      const file = path.join(parent, item.name);
      if (/^pending-\d+-[a-f0-9-]+$/.test(item.name)) {
        occupied += await artifactBytes(file);
        continue;
      }
      if (!hashPattern.test(item.name)) continue;
      try {
        const info = await fs.lstat(file),
          bytes = await artifactBytes(file);
        occupied += bytes;
        entries.push({ file, bytes, time: info.mtimeMs });
      } catch {
        await fs.rm(file, { recursive: true, force: true });
      }
    }
  entries.sort((a, b) => a.time - b.time);
  while (entries.length && (occupied > maxBytes || entries.length > maxEntries)) {
    const entry = entries.shift();
    await fs.rm(entry.file, { recursive: true, force: true });
    occupied -= entry.bytes;
  }
}
async function sourceIdentity(source, signal) {
  const before = await regular(source),
    revision = await fileDigest(source, { signal }),
    after = await regular(source);
  if (
    before.dev !== after.dev ||
    before.ino !== after.ino ||
    before.size !== after.size ||
    before.mtimeMs !== after.mtimeMs ||
    before.ctimeMs !== after.ctimeMs
  )
    throw refusal('session-stale', 'Workbook changed while preparing cache');
  return revision;
}
function manifestValid(manifest, identity, revision) {
  if (
    manifest.schema !== SCHEMA ||
    manifest.identity !== identity ||
    manifest.revision !== revision ||
    !Number.isSafeInteger(manifest.version) ||
    manifest.version < 0 ||
    typeof manifest.reuseCalculation !== 'boolean' ||
    !Array.isArray(manifest.files) ||
    manifest.files.length > 10000
  )
    throw refusal('session-malformed', 'Invalid cache identity');
  const names = new Set();
  let bytes = 0;
  for (const file of manifest.files) {
    if (
      !artifact(file.name) ||
      names.has(file.name) ||
      !Number.isSafeInteger(file.bytes) ||
      file.bytes < 0 ||
      !hashPattern.test(file.sha256)
    )
      throw refusal('internal-invalid-field', 'Invalid cache artifact inventory');
    names.add(file.name);
    bytes += file.bytes;
  }
  if (
    !['source.xlsx', 'cells.sqlite', 'metadata.xlsx', 'index.json'].every((name) => names.has(name)) ||
    bytes !== manifest.bytes
  )
    throw refusal('session-unsupported', 'Incomplete cached index');
  if (manifest.checkpoint && !['saved.xlsx', 'checkpoint.json'].every((name) => names.has(name)))
    throw refusal('session-unsupported', 'Incomplete cached checkpoint');
  const source = manifest.files.find((file) => file.name === (manifest.checkpoint ? 'saved.xlsx' : 'source.xlsx'));
  if (source.sha256 !== revision) throw refusal('internal-mismatch', 'Cached workbook revision mismatch');
  return manifest;
}
async function verifyManifest(root, identity, revision) {
  const manifest = manifestValid(await readJson(path.join(root, 'cache.json')), identity, revision);
  const names = (await fs.readdir(root)).sort(),
    expected = [...manifest.files.map((file) => file.name), 'cache.json'].sort();
  if (JSON.stringify(names) !== JSON.stringify(expected))
    throw refusal('session-malformed', 'Unexpected cached index artifacts');
  for (const file of manifest.files)
    if ((await regular(path.join(root, file.name))).size !== file.bytes)
      throw refusal('internal-mismatch', 'Cached index size mismatch');
  return manifest;
}
/** options.directory/maxBytes/maxEntries permit isolated component hosts/tests.
 * Default quotas include staging. A host identity must cover native/schema and
 * calculation implementation/policy, not just the application version string. */
export async function indexedCache(source, identity, options = {}) {
  if (typeof identity !== 'string' || !identity.length || identity.length > 4096)
    throw refusal('session-malformed', 'Invalid workbook cache identity');
  const config = settings(options),
    { signal, maxBytes, maxEntries } = config,
    parent = await directory(config.directory);
  // A host may lower a persistent cache's allowance between launches. Existing
  // entries must obey the new allowance before restore, not only on next publish.
  if (maxBytes < MAX_BYTES || maxEntries < MAX_ENTRIES)
    await lock(parent, signal, () => trimToLimits(parent, maxBytes, maxEntries));
  const revision = await sourceIdentity(source, signal);
  const key = createHash('sha256')
      .update(`index-cache:${SCHEMA}\0`)
      .update(identity)
      .update('\0')
      .update(revision)
      .digest('hex'),
    root = path.join(parent, key);
  return {
    revision,
    async invalidate() {
      return lock(parent, signal, () => fs.rm(root, { recursive: true, force: true }));
    },
    async restore(target, sessionBytes = 4 * 1024 ** 3) {
      if (!Number.isSafeInteger(sessionBytes) || sessionBytes < 0)
        throw refusal('internal-invalid-field', 'Invalid workbook session disk budget');
      const targetStat = await fs.lstat(target);
      if (!targetStat.isDirectory() || targetStat.isSymbolicLink() || (await fs.readdir(target)).length)
        throw refusal('session-unsupported', 'Cache restore requires an empty private directory');
      return lock(parent, signal, async () => {
        const written = [];
        try {
          const stat = await fs.lstat(root);
          if (!stat.isDirectory() || stat.isSymbolicLink()) throw refusal('session-unsupported', 'Unsafe cached index');
          const manifest = await verifyManifest(root, identity, revision);
          if (manifest.bytes + 4096 > sessionBytes) {
            const error = refusal('session-limit', 'Workbook session disk budget exhausted while restoring index');
            error.code = 'CACHE_SESSION_BUDGET';
            throw error;
          }
          for (const file of manifest.files) {
            cancelled(signal);
            const from = path.join(root, file.name),
              to = path.join(target, file.name);
            if ((await regular(from)).size !== file.bytes)
              throw refusal('internal-mismatch', 'Cached index size mismatch');
            await fs.copyFile(from, to, constants.COPYFILE_EXCL | constants.COPYFILE_FICLONE);
            written.push(to);
            if ((await regular(to)).size !== file.bytes || (await fileDigest(to, { signal })) !== file.sha256)
              throw refusal('internal-mismatch', 'Cached index digest mismatch');
          }
          const index = await readJson(path.join(target, 'index.json'), 16 * 1024 * 1024);
          if (index.revision !== manifest.files.find((file) => file.name === 'source.xlsx').sha256)
            throw refusal('internal-mismatch', 'Cached index base revision mismatch');
          if ((await sourceIdentity(source, signal)) !== revision) {
            const error = refusal('session-stale', 'Workbook changed while restoring cached index');
            error.code = 'CACHE_SOURCE_CHANGED';
            throw error;
          }
          cancelled(signal);
          await fs.utimes(root, new Date(), new Date());
          return manifest;
        } catch (error) {
          for (const file of written) await fs.rm(file, { force: true });
          if (signal?.aborted || error.code === 'CACHE_SESSION_BUDGET' || error.code === 'CACHE_SOURCE_CHANGED')
            throw error;
          // Disposable corruption or eviction is a miss. Never modify the workbook.
          if (error.code !== 'ENOENT') await fs.rm(root, { recursive: true, force: true });
          return null;
        }
      });
    },
    async publish(sourceRoot, version, publication = {}) {
      if (!Number.isSafeInteger(version) || version < 0)
        throw refusal('internal-invalid-field', 'Invalid cached version');
      return lock(parent, signal, async () => {
        // Check the complete private directory first, including transient files. A
        // huge failed operation must not bypass quota by hiding its temporary files.
        if ((await artifactBytes(sourceRoot)) > maxBytes) return null;
        const selected = await filesIn(sourceRoot, durable);
        if (
          !['source.xlsx', 'cells.sqlite', 'metadata.xlsx', 'index.json'].every((name) =>
            selected.some((file) => file.name === name)
          )
        )
          return null;
        const checkpoint = publication.checkpoint;
        const checkpointBytes = checkpoint ? Buffer.from(JSON.stringify(checkpoint)) : null;
        if (checkpointBytes?.length > 16 * 1024 * 1024) return null;
        const savedBytes = checkpoint ? (await regular(source)).size : 0;
        const workingBytes = publication.workingBytes ?? 0;
        if (!Number.isSafeInteger(workingBytes) || workingBytes < 0)
          throw refusal('session-malformed', 'Invalid cache staging budget');
        const needed =
          selected.reduce((n, file) => n + file.bytes, 0) + (checkpointBytes?.length ?? 0) + savedBytes + MAX_MANIFEST;
        if (needed + workingBytes > maxBytes) return null;
        const entries = [];
        let occupied = 0;
        for (const item of await fs.readdir(parent, { withFileTypes: true }))
          if (item.isDirectory()) {
            const file = path.join(parent, item.name);
            if (/^pending-\d+-[a-f0-9-]+$/.test(item.name)) {
              // A worker can disappear while its owning process remains live. Its
              // disposable staging still consumes space until safe owner cleanup.
              occupied += await artifactBytes(file);
              continue;
            }
            if (!hashPattern.test(item.name)) continue;
            try {
              const stat = await fs.lstat(file),
                bytes = await artifactBytes(file);
              occupied += bytes;
              entries.push({ file, time: stat.mtimeMs, bytes });
            } catch {
              await fs.rm(file, { recursive: true, force: true });
            }
          }
        const existing = entries.find((entry) => entry.file === root);
        if (existing) {
          try {
            await verifyManifest(root, identity, revision);
            return { published: false, existing: true };
          } catch {
            await fs.rm(root, { recursive: true, force: true });
            occupied -= existing.bytes;
            entries.splice(entries.indexOf(existing), 1);
          }
        }
        entries.sort((a, b) => a.time - b.time);
        while (entries.length && (occupied + needed + workingBytes > maxBytes || entries.length >= maxEntries)) {
          const entry = entries.shift();
          await fs.rm(entry.file, { recursive: true, force: true });
          occupied -= entry.bytes;
        }
        if (occupied + needed + workingBytes > maxBytes) return null;
        const staging = path.join(parent, `pending-${process.pid}-${randomUUID()}`);
        await fs.mkdir(staging, { mode: 0o700 });
        try {
          for (const file of selected) {
            cancelled(signal);
            await fs.copyFile(
              path.join(sourceRoot, file.name),
              path.join(staging, file.name),
              constants.COPYFILE_EXCL | constants.COPYFILE_FICLONE
            );
          }
          if (checkpoint) {
            await fs.copyFile(
              source,
              path.join(staging, 'saved.xlsx'),
              constants.COPYFILE_EXCL | constants.COPYFILE_FICLONE
            );
            await fs.writeFile(path.join(staging, 'checkpoint.json'), checkpointBytes, { mode: 0o600, flag: 'wx' });
          }
          if (publication.prepareSnapshot) {
            cancelled(signal);
            await publication.prepareSnapshot(staging);
          }
          // Compaction may create lock files, but only this closed allowlist becomes
          // persistent. No earlier export, candidate, cancellation or journal leaks.
          for (const file of await fs.readdir(staging))
            if (!artifact(file)) await fs.rm(path.join(staging, file), { recursive: true, force: true });
          const files = await filesIn(staging);
          let bytes = 0;
          for (const file of files) {
            cancelled(signal);
            bytes += file.bytes;
            if (occupied + bytes + MAX_MANIFEST > maxBytes)
              throw refusal('session-limit', 'Workbook cache staging exceeded its reserved disk budget');
            file.sha256 = await fileDigest(path.join(staging, file.name), { signal });
          }
          // The caller's native cache-status proof covers the exact version and
          // metadata. A tiny dependency hint is not proof: dynamic formulas may
          // also have a complete runtime dependency graph.
          const reuseCalculation = !!checkpoint || (publication.reuseCalculation ?? version === 0);
          const manifest = manifestValid(
            { schema: SCHEMA, revision, identity, version, bytes, reuseCalculation, checkpoint: !!checkpoint, files },
            identity,
            revision
          );
          const index = await readJson(path.join(staging, 'index.json'), 16 * 1024 * 1024);
          if (index.revision !== files.find((file) => file.name === 'source.xlsx').sha256)
            throw refusal('internal-mismatch', 'Cached index base revision mismatch');
          const data = JSON.stringify(manifest);
          if (Buffer.byteLength(data) > MAX_MANIFEST)
            throw refusal('session-limit', 'Workbook cache inventory exceeds budget');
          await fs.writeFile(path.join(staging, 'cache.json'), data, { mode: 0o600, flag: 'wx' });
          cancelled(signal);
          await fs.rename(staging, root);
          return { published: true, bytes: bytes + Buffer.byteLength(data) };
        } finally {
          await fs.rm(staging, { recursive: true, force: true });
        }
      });
    },
  };
}
export async function indexedCacheUsage(options = {}) {
  const config = settings(options),
    parent = await directory(config.directory);
  return lock(parent, config.signal, async () => {
    let bytes = 0,
      entries = 0,
      staging = 0;
    for (const item of await fs.readdir(parent, { withFileTypes: true }))
      if (item.isDirectory() && (hashPattern.test(item.name) || item.name.startsWith('pending-'))) {
        bytes += await artifactBytes(path.join(parent, item.name));
        if (item.name.startsWith('pending-')) staging++;
        else entries++;
      }
    return { bytes, entries, staging, maxBytes: config.maxBytes, maxEntries: config.maxEntries };
  });
}
