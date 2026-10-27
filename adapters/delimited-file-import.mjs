/** Trusted Node file-chooser adapter. Absolute input paths never form a public
 * renderer API; callers hold the chosen-file grant and captured workbook scope. */
import fs from 'node:fs/promises';
import { constants } from 'node:fs';
import path from 'node:path';
import { createHash, randomUUID } from 'node:crypto';
import { DELIMITED_READ_BYTES, inspectDelimitedSource, stageDelimitedImport } from '../core/delimited-import.mjs';
import { PASTE_BYTE_LIMIT } from '../core/paste-protocol.mjs';
import { refusal } from '../core/refusal-codes.mjs';

const stamp = (info) => [info.dev, info.ino, info.size, info.mtimeNs, info.ctimeNs].join(':');
const uid = () => (typeof process.getuid === 'function' ? process.getuid() : undefined);
const owned = (info) => uid() === undefined || BigInt(info.uid) === BigInt(uid());
async function writeAll(fd, bytes) {
  for (let at = 0; at < bytes.length; ) {
    const result = await fd.write(bytes, at, bytes.length - at);
    if (!result.bytesWritten) throw refusal('csv-bad-input', 'Delimited snapshot write did not advance');
    at += result.bytesWritten;
  }
}

/** Exclusive private snapshot, three matching raw-byte hashes and descriptor/
 * path identity checks. Retain the original descriptor until accept/dispose so
 * a source replacement or edit before acceptance invalidates the one-shot grant. */
export async function snapshotDelimitedFile({
  source: chosen,
  directory,
  signal,
  checkCurrent = () => {},
  onProgress = () => {},
}) {
  if (
    typeof chosen !== 'string' ||
    !path.isAbsolute(chosen) ||
    chosen.length > 4096 ||
    chosen.includes('\0') ||
    !['.csv', '.tsv'].includes(path.extname(chosen).toLowerCase()) ||
    typeof directory !== 'string' ||
    !path.isAbsolute(directory)
  )
    throw refusal('csv-unsupported', 'A granted CSV/TSV file and private import directory are required');
  const name = path.basename(chosen);
  if (name.length > 255) throw refusal('csv-limit', 'Delimited source filename exceeds its bound');
  const guard = async () => {
    signal?.throwIfAborted();
    await checkCurrent();
    signal?.throwIfAborted();
  };
  await guard();
  await fs.mkdir(directory, { recursive: true, mode: 0o700 });
  const baseInfo = await fs.lstat(directory);
  if (!baseInfo.isDirectory() || baseInfo.isSymbolicLink() || !owned(baseInfo) || (baseInfo.mode & 0o077) !== 0)
    throw refusal('csv-unsupported', 'Delimited import directory is not privately owned');
  const canonical = await fs.realpath(chosen);
  directory = await fs.realpath(directory);
  const root = await fs.mkdtemp(path.join(directory, 'file-'));
  await fs.chmod(root, 0o700);
  const rootInfo = await fs.lstat(root, { bigint: true }),
    filename = path.join(root, 'source.bytes'),
    id = randomUUID();
  let original,
    destination,
    closed = false;
  const cleanup = async () => {
    const actual = await fs.lstat(root, { bigint: true });
    if (
      !actual.isDirectory() ||
      actual.isSymbolicLink() ||
      actual.dev !== rootInfo.dev ||
      actual.ino !== rootInfo.ino ||
      !owned(actual)
    )
      throw refusal('csv-stale', 'Delimited private directory identity changed; existing entries retained');
    const names = await fs.readdir(root);
    if (names.some((value) => value !== 'source.bytes'))
      throw refusal('csv-malformed', 'Unexpected delimited snapshot entries retained');
    if (names.length) {
      const info = await fs.lstat(filename);
      if (!info.isFile() || info.isSymbolicLink() || !owned(info) || info.nlink !== 1)
        throw refusal('csv-malformed', 'Unexpected delimited snapshot entry retained');
      await fs.unlink(filename);
    }
    await fs.rmdir(root);
  };
  try {
    original = await fs.open(canonical, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0));
    const originalInfo = await original.stat({ bigint: true }),
      originalStamp = stamp(originalInfo),
      size = Number(originalInfo.size);
    if (!originalInfo.isFile() || !Number.isSafeInteger(size) || size > PASTE_BYTE_LIMIT)
      throw refusal('csv-limit', 'Delimited source must be a regular file of at most 64 MiB');
    const originalCheck = async () => {
      await guard();
      if (
        stamp(await original.stat({ bigint: true })) !== originalStamp ||
        (await fs.realpath(chosen)) !== canonical ||
        stamp(await fs.stat(canonical, { bigint: true })) !== originalStamp
      )
        throw refusal('csv-stale', 'Chosen delimited source changed before acceptance');
    };
    const readOriginal = async (copy = false) => {
      const hash = createHash('sha256'),
        buffer = Buffer.allocUnsafe(DELIMITED_READ_BYTES);
      let at = 0;
      while (at < size) {
        await originalCheck();
        const result = await original.read(buffer, 0, Math.min(buffer.length, size - at), at);
        if (!result.bytesRead) throw refusal('csv-malformed', 'Delimited source ended during snapshot');
        const chunk = buffer.subarray(0, result.bytesRead);
        hash.update(chunk);
        if (copy) await writeAll(destination, chunk);
        at += result.bytesRead;
        onProgress({ phase: copy ? 'copying-file' : 'checking-file', completed: at, total: size });
      }
      await originalCheck();
      return hash.digest('hex');
    };
    const before = await readOriginal();
    destination = await fs.open(filename, 'wx', 0o600);
    const copied = await readOriginal(true);
    await destination.sync();
    await destination.close();
    destination = undefined;
    const after = await readOriginal();
    if (before !== copied || before !== after)
      throw refusal('csv-stale', 'Chosen delimited source changed while copying');
    await fs.chmod(filename, 0o400);
    const snapshotStamp = stamp(await fs.lstat(filename, { bigint: true }));
    const check = async () => {
      if (closed) throw refusal('session-closed', 'Delimited snapshot owner is closed');
      await originalCheck();
      const info = await fs.lstat(filename, { bigint: true }),
        parent = await fs.lstat(root, { bigint: true });
      if (
        !info.isFile() ||
        info.isSymbolicLink() ||
        info.nlink !== 1n ||
        !owned(info) ||
        stamp(info) !== snapshotStamp ||
        (info.mode & 0o222n) !== 0n ||
        !parent.isDirectory() ||
        parent.dev !== rootInfo.dev ||
        parent.ino !== rootInfo.ino ||
        (await fs.realpath(root)) !== root
      )
        throw refusal('csv-stale', 'Private delimited snapshot changed');
    };
    return {
      id,
      name,
      size,
      sha256: before,
      check,
      async *chunks() {
        await check();
        const fd = await fs.open(filename, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0));
        try {
          if (stamp(await fd.stat({ bigint: true })) !== snapshotStamp)
            throw refusal('csv-stale', 'Delimited snapshot changed while opening');
          const buffer = Buffer.allocUnsafe(DELIMITED_READ_BYTES);
          let at = 0;
          while (at < size) {
            await check();
            const result = await fd.read(buffer, 0, Math.min(buffer.length, size - at), at);
            if (!result.bytesRead) throw refusal('csv-malformed', 'Delimited snapshot ended early');
            at += result.bytesRead;
            yield buffer.subarray(0, result.bytesRead);
          }
          await check();
        } finally {
          await fd.close();
        }
      },
      async dispose() {
        if (closed) return;
        closed = true;
        await original.close();
        original = undefined;
        await cleanup();
      },
    };
  } catch (error) {
    await destination?.close().catch(() => {});
    await original?.close().catch(() => {});
    await cleanup().catch((cleanupError) => {
      error.cleanupFailure = String(cleanupError);
    });
    throw error;
  }
}

/** Host-side preparation only. Hosts commit the returned compact descriptor
 * through their existing owner apply/cancellation fence and finally dispose. */
export async function prepareDelimitedFileImport({
  source,
  directory,
  sheet,
  start,
  delimiter,
  signal,
  checkCurrent,
  onProgress,
  sink,
  state,
}) {
  const snapshot = await snapshotDelimitedFile({ source, directory, signal, checkCurrent, onProgress });
  try {
    const inspection = await inspectDelimitedSource(snapshot, {
      sheet,
      start,
      delimiter: delimiter ?? (snapshot.name.toLowerCase().endsWith('.tsv') ? '\t' : 'auto'),
      signal,
      onProgress,
    });
    const ready = await stageDelimitedImport(snapshot, inspection, { state, signal, checkCurrent, sink, onProgress });
    return { ...ready, inspection, check: () => snapshot.check(), dispose: () => snapshot.dispose() };
  } catch (error) {
    await snapshot.dispose().catch((cleanupError) => {
      error.cleanupFailure = String(cleanupError);
    });
    throw error;
  }
}
