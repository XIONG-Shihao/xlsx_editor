/** Main-only bounded conversion/replacement review. Exact source/destination
 * identities stay private; only bounded typed display pages reach the UI. */
import fs from 'node:fs/promises';
import { constants } from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { convertTextField } from '../core/text-conversion.mjs';
import { localQueryConversionIssue } from '../core/local-query.mjs';
import { refusal } from '../core/refusal-codes.mjs';
export const QUERY_INSPECTION_LIMITS = Object.freeze({
  receipts: 4,
  lifetimeMs: 10 * 60 * 1000,
  rawBytes: 64 * 1024 * 1024,
  indexBytes: 8 * 1024 * 1024,
  pageIssues: 100,
  pageBytes: 1024 * 1024,
});
const stamp = (s) => [s.dev, s.ino, s.size, s.mtimeNs, s.ctimeNs].join(':'),
  owned = (s) => typeof process.getuid !== 'function' || Number(s.uid) === process.getuid(),
  entries = new Map(),
  canonical = JSON.stringify;
async function write(fd, bytes) {
  let at = 0;
  while (at < bytes.length) {
    const r = await fd.write(bytes, at, bytes.length - at);
    if (!r.bytesWritten) throw refusal('query-bad-input', 'Query inspection write did not advance');
    at += r.bytesWritten;
  }
}
async function read(fd, length, position) {
  const bytes = Buffer.alloc(length);
  let at = 0;
  while (at < length) {
    const r = await fd.read(bytes, at, length - at, position + at);
    if (!r.bytesRead) throw refusal('query-malformed', 'Query inspection ended early');
    at += r.bytesRead;
  }
  return bytes;
}
export async function querySourceStamp(filename) {
  const info = await fs.lstat(filename, { bigint: true });
  if (!info.isFile() || info.isSymbolicLink() || (await fs.realpath(filename)) !== filename)
    throw refusal('query-stale', 'Saved query source changed canonical identity');
  return stamp(info);
}
async function retire(entry) {
  if (entries.get(entry.id) === entry) entries.delete(entry.id);
  clearTimeout(entry.timer);
  if (entry.users) {
    entry.retired = true;
    return;
  }
  await entry.cleanup();
}
export async function releaseQueryInspections(document, directory, options) {
  if (
    typeof document !== 'string' ||
    !document ||
    !path.isAbsolute(directory) ||
    !options ||
    typeof options !== 'object' ||
    Array.isArray(options) ||
    Object.keys(options).some((key) => !['ownerId', 'connectionId'].includes(key)) ||
    !Object.hasOwn(options, 'ownerId') ||
    (options.ownerId !== null && (typeof options.ownerId !== 'string' || !options.ownerId)) ||
    (options.connectionId !== undefined && typeof options.connectionId !== 'string')
  )
    throw refusal('query-bad-input', 'Query inspection disposal requires its captured owner');
  for (const entry of [...entries.values()])
    if (
      entry.binding.document === document &&
      entry.directory === directory &&
      entry.binding.scope.ownerId === options.ownerId &&
      (!options.connectionId || entry.binding.connectionId === options.connectionId)
    )
      await retire(entry);
}
export async function createQueryInspectionSpool({ directory, binding, signal, check = () => {} }) {
  await fs.mkdir(directory, { recursive: true, mode: 0o700 });
  const base = await fs.lstat(directory);
  if (!base.isDirectory() || base.isSymbolicLink() || !owned(base) || (base.mode & 0o077) !== 0)
    throw refusal('query-unsupported', 'Query inspection directory must be privately owned');
  directory = await fs.realpath(directory);
  for (const entry of [...entries.values()])
    if (
      entry.expires <= Date.now() ||
      (entry.binding.document === binding.document && entry.binding.connectionId === binding.connectionId)
    )
      await retire(entry);
  while (entries.size >= QUERY_INSPECTION_LIMITS.receipts) await retire(entries.values().next().value);
  const id = randomUUID(),
    root = await fs.mkdtemp(path.join(directory, 'query-review-'));
  await fs.chmod(root, 0o700);
  const rootInfo = await fs.lstat(root, { bigint: true }),
    streams = {};
  let bytes = 0,
    count = 0,
    closed = false;
  const names = ['issues.ndjson', 'offsets.bin', 'replacements.ndjson', 'replacement-offsets.bin'];
  const cleanup = async () => {
    if (closed) return;
    closed = true;
    for (const stream of Object.values(streams)) {
      await stream.data?.close();
      await stream.index?.close();
      stream.data = stream.index = null;
    }
    const info = await fs.lstat(root, { bigint: true });
    if (
      !info.isDirectory() ||
      info.isSymbolicLink() ||
      info.dev !== rootInfo.dev ||
      info.ino !== rootInfo.ino ||
      !owned(info)
    )
      throw refusal('query-stale', 'Query inspection directory changed; bytes retained');
    const actual = await fs.readdir(root);
    if (actual.some((name) => !names.includes(name)))
      throw refusal('query-malformed', 'Unexpected query inspection files retained');
    for (const name of actual) {
      const file = path.join(root, name),
        s = await fs.lstat(file);
      if (!s.isFile() || s.isSymbolicLink() || s.nlink !== 1 || !owned(s))
        throw refusal('query-stale', 'Query inspection file changed; bytes retained');
      await fs.unlink(file);
    }
    await fs.rmdir(root);
  };
  try {
    for (const [key, dataName, indexName] of [
      ['issues', names[0], names[1]],
      ['replacements', names[2], names[3]],
    ]) {
      const stream = {
        dataPath: path.join(root, dataName),
        indexPath: path.join(root, indexName),
        bytes: 0,
        count: 0,
        buffers: [],
        indices: [],
        bufferBytes: 0,
      };
      streams[key] = stream;
      stream.data = await fs.open(stream.dataPath, 'wx', 0o600);
      stream.index = await fs.open(stream.indexPath, 'wx', 0o600);
    }
  } catch (error) {
    await cleanup().catch(() => {});
    throw error;
  }
  const flush = async (stream) => {
    if (!stream.buffers.length) return;
    signal?.throwIfAborted();
    await check();
    await write(stream.index, Buffer.concat(stream.indices));
    await write(stream.data, Buffer.concat(stream.buffers));
    stream.buffers = [];
    stream.indices = [];
    stream.bufferBytes = 0;
  };
  const append = async (key, value) => {
    signal?.throwIfAborted();
    const stream = streams[key],
      raw = Buffer.from(JSON.stringify(value) + '\n');
    if (bytes + raw.length > QUERY_INSPECTION_LIMITS.rawBytes || (count + 1) * 8 > QUERY_INSPECTION_LIMITS.indexBytes)
      throw refusal('query-limit', 'Query review exceeds its combined 64 MiB raw / 8 MiB index limit');
    if (stream.bufferBytes + raw.length > 1024 * 1024 || stream.buffers.length >= 1000) await flush(stream);
    const offset = Buffer.alloc(8);
    offset.writeBigUInt64LE(BigInt(stream.bytes));
    stream.indices.push(offset);
    stream.buffers.push(raw);
    stream.bufferBytes += raw.length;
    stream.bytes += raw.length;
    stream.count++;
    bytes += raw.length;
    count++;
  };
  return {
    id,
    async add(cell) {
      if (cell.result.status === 'converted' && !cell.result.notices.length) return;
      await append('issues', [
        cell.index,
        cell.sourceRow,
        cell.sourceColumn,
        cell.outputRow,
        cell.outputColumn,
        cell.raw,
        cell.policy,
        cell.result.padded,
      ]);
    },
    async addReplacement(row) {
      await append('replacements', row);
    },
    async seal(inspection, replacementSummary) {
      signal?.throwIfAborted();
      await check();
      for (const stream of Object.values(streams)) {
        await flush(stream);
        await stream.data.sync();
        await stream.index.sync();
        await stream.data.close();
        await stream.index.close();
        stream.data = stream.index = null;
        await fs.chmod(stream.dataPath, 0o400);
        await fs.chmod(stream.indexPath, 0o400);
        stream.stamps = await Promise.all(
          [stream.dataPath, stream.indexPath].map(async (file) => stamp(await fs.lstat(file, { bigint: true })))
        );
      }
      while (entries.size >= QUERY_INSPECTION_LIMITS.receipts) await retire(entries.values().next().value);
      const entry = {
        id,
        directory,
        binding: structuredClone(binding),
        inspection,
        replacementSummary,
        root,
        rootInfo,
        streams,
        count: streams.issues.count,
        bytes,
        expires: Date.now() + QUERY_INSPECTION_LIMITS.lifetimeMs,
        cleanup,
        users: 0,
        retired: false,
      };
      entries.set(id, entry);
      entry.timer = setTimeout(
        () => void retire(entry).catch((e) => console.warn('Query review cleanup deferred', e)),
        QUERY_INSPECTION_LIMITS.lifetimeMs
      );
      entry.timer.unref?.();
      return entry;
    },
    dispose: cleanup,
  };
}
async function page(entry, key, offset, { signal, check, budget }) {
  const stream = entry.streams[key];
  if (!Number.isInteger(offset) || offset < 0 || offset > stream.count)
    throw refusal('internal-invalid-field', 'Invalid query review page');
  const fds = [];
  try {
    for (let n = 0; n < 2; n++) {
      const file = [stream.dataPath, stream.indexPath][n],
        fd = await fs.open(file, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0));
      fds.push(fd);
      const info = await fd.stat({ bigint: true }),
        named = await fs.lstat(file, { bigint: true });
      if (
        !info.isFile() ||
        !owned(info) ||
        info.nlink !== 1n ||
        stamp(info) !== stream.stamps[n] ||
        stamp(named) !== stream.stamps[n] ||
        (info.mode & 0o222n) !== 0n
      )
        throw refusal('query-stale', 'Query inspection file changed');
    }
    const result = [];
    let next = offset,
      bytes = 2;
    for (; next < stream.count && result.length < QUERY_INSPECTION_LIMITS.pageIssues; next++) {
      signal?.throwIfAborted();
      const start = Number((await read(fds[1], 8, next * 8)).readBigUInt64LE()),
        end =
          next + 1 < stream.count ? Number((await read(fds[1], 8, (next + 1) * 8)).readBigUInt64LE()) : stream.bytes;
      if (end <= start || end > stream.bytes || end - start > 1024 * 1024)
        throw refusal('query-malformed', 'Invalid query review index');
      let row = JSON.parse((await read(fds[0], end - start, start)).toString('utf8'));
      if (key === 'issues') {
        const [index, sourceRow, sourceColumn, outputRow, outputColumn, text, policy, padded] = row,
          result = convertTextField(text, policy, { date1904: entry.inspection.conversion?.date1904, padded });
        row = localQueryConversionIssue({
          index,
          raw: text,
          sourceRow,
          sourceColumn,
          outputRow,
          outputColumn,
          target: entry.binding.target,
          policy,
          result,
        });
      }
      const size = Buffer.byteLength(JSON.stringify(row));
      const added = size + (result.length ? 1 : 0);
      if (bytes + added > budget) {
        if (!result.length && budget >= QUERY_INSPECTION_LIMITS.pageBytes - 256 * 1024)
          throw refusal('query-limit', 'Query review row exceeds its response limit');
        break;
      }
      bytes += added;
      result.push(row);
    }
    for (let n = 0; n < 2; n++) {
      const file = [stream.dataPath, stream.indexPath][n];
      if (
        stamp(await fds[n].stat({ bigint: true })) !== stream.stamps[n] ||
        stamp(await fs.lstat(file, { bigint: true })) !== stream.stamps[n]
      )
        throw refusal('query-stale', 'Query inspection file changed during read');
    }
    await check();
    return { rows: result, offset, nextOffset: next < stream.count ? next : null, bytes };
  } finally {
    await Promise.all(fds.map((fd) => fd.close()));
  }
}
export async function readQueryInspection(
  id,
  binding,
  offset = 0,
  { signal, check = () => {}, replacementOffset = 0, replacementFirst = false } = {}
) {
  const entry = entries.get(id);
  if (entry?.expires <= Date.now()) await retire(entry);
  if (!entry || entry.expires <= Date.now() || canonical(entry.binding) !== canonical(binding))
    throw refusal(
      'query-stale',
      'Query review expired or its source, settings, owner or target changed. Preview again.'
    );
  entry.users++;
  try {
    signal?.throwIfAborted();
    await check();
    const parent = await fs.lstat(entry.root, { bigint: true });
    if (
      !parent.isDirectory() ||
      parent.isSymbolicLink() ||
      parent.dev !== entry.rootInfo.dev ||
      parent.ino !== entry.rootInfo.ino ||
      !owned(parent)
    )
      throw refusal('query-stale', 'Query inspection directory changed');
    const budget = QUERY_INSPECTION_LIMITS.pageBytes - 256 * 1024;
    let issues, replacement;
    if (replacementFirst) {
      replacement = await page(entry, 'replacements', replacementOffset, { signal, check, budget });
      issues = await page(entry, 'issues', offset, { signal, check, budget: Math.max(2, budget - replacement.bytes) });
    } else {
      issues = await page(entry, 'issues', offset, { signal, check, budget });
      replacement = await page(entry, 'replacements', replacementOffset, {
        signal,
        check,
        budget: Math.max(2, budget - issues.bytes),
      });
    }
    await check();
    signal?.throwIfAborted();
    if (entries.get(id) !== entry || entry.retired) throw refusal('query-stale', 'Query review retired');
    return {
      entry,
      issues: issues.rows,
      offset: issues.offset,
      nextOffset: issues.nextOffset,
      replacements: replacement.rows,
      replacementOffset: replacement.offset,
      nextReplacementOffset: replacement.nextOffset,
    };
  } finally {
    entry.users--;
    if (entry.retired && !entry.users) await entry.cleanup();
  }
}
