import { createHash, randomUUID } from 'node:crypto';
import { constants } from 'node:fs';
import { mkdir, lstat, realpath, open, rename, rm, chmod, readdir } from 'node:fs/promises';
import path from 'node:path';
import { syncDirectory } from './storage.mjs';
import {
  PASTE_BYTE_LIMIT,
  PASTE_PAGE_BYTES,
  PASTE_PAGE_CELLS,
  PASTE_SCHEMA,
  TYPED_PASTE_SCHEMA,
  TYPED_PASTE_POLICY,
  pasteHeader,
  pasteId,
  validatePasteState,
  validatePastePlan,
  validatePasteDescriptor,
  validatePasteCells,
} from '../core/paste-protocol.mjs';
import { refusal } from '../core/refusal-codes.mjs';

const sealedSuffix = '.paste.jsonl',
  partialSuffix = '.paste.partial';
const fileId = (name) =>
  name.endsWith(sealedSuffix)
    ? name.slice(0, -sealedSuffix.length)
    : name.endsWith(partialSuffix)
      ? name.slice(0, -partialSuffix.length)
      : null;
const equal = (a, b) => JSON.stringify(a) === JSON.stringify(b);
async function fileHash(fd, check) {
  const hash = createHash('sha256'),
    buffer = Buffer.allocUnsafe(64 * 1024);
  let position = 0;
  for (;;) {
    check();
    const { bytesRead } = await fd.read(buffer, 0, buffer.length, position);
    if (!bytesRead) break;
    hash.update(buffer.subarray(0, bytesRead));
    position += bytesRead;
  }
  check();
  return { sha256: hash.digest('hex'), bytes: position };
}
/** Private owner storage only. Public requests and durable commands carry opaque
 * IDs; this object never accepts a renderer path or resolved cell commands. */
export function createPasteStager({
  directory,
  documentId,
  maxBytes = PASTE_BYTE_LIMIT,
  checkCancelled = () => {},
  accountBytes = async () => {},
}) {
  if (
    typeof directory !== 'string' ||
    !path.isAbsolute(directory) ||
    !(pasteId(documentId) || (typeof documentId === 'string' && /^[a-f0-9]{64}$/.test(documentId))) ||
    !Number.isSafeInteger(maxBytes) ||
    maxBytes < 1 ||
    maxBytes > PASTE_BYTE_LIMIT
  )
    throw refusal('paste-malformed', 'Invalid private paste staging owner');
  directory = path.resolve(directory);
  const active = new Map();
  let ready,
    used = 0,
    fileCount = 0,
    work = Promise.resolve(),
    closed = false;
  const check = () => {
    if (closed) throw refusal('session-closed', 'Paste staging owner is closed');
    checkCancelled();
  };
  const confined = async () => {
    if ((await realpath(directory)) !== directory || (await lstat(directory)).isSymbolicLink())
      throw refusal('paste-bad-input', 'Paste staging directory is not confined');
  };
  const initialize = async () => {
    await mkdir(directory, { recursive: true, mode: 0o700 });
    const info = await lstat(directory);
    if (!info.isDirectory() || info.isSymbolicLink())
      throw refusal('paste-bad-input', 'Paste staging directory is not confined');
    // macOS /var and /tmp aliases may occur in a trusted owner root. Resolve
    // those once; attachment names remain confined beneath the canonical root.
    directory = await realpath(directory);
    await confined();
    const names = await readdir(directory);
    if (names.length > 4096) throw refusal('paste-limit', 'Paste staging directory has too many entries');
    fileCount = names.length;
    for (const name of names) {
      const id = fileId(name);
      if (!pasteId(id)) continue;
      const filename = path.join(directory, name),
        info = await lstat(filename);
      if (!info.isFile()) throw refusal('paste-bad-input', 'Paste attachment must be a regular confined file');
      if (name.endsWith(partialSuffix)) {
        await rm(filename);
        fileCount--;
      } else {
        used += info.size;
        if (used > maxBytes) throw refusal('paste-limit', 'Paste attachments exceed journal byte budget');
      }
    }
    await syncDirectory(directory);
  };
  const serial = (operation, cancelling = false) => {
    const checkOwner = () => {
      if (closed) throw refusal('session-closed', 'Paste staging owner is closed');
      if (!cancelling) checkCancelled();
    };
    const result = work.then(async () => {
      checkOwner();
      ready ??= initialize();
      await ready;
      checkOwner();
      await confined();
      return operation();
    });
    work = result.catch(() => {});
    return result;
  };
  const uploadFor = (request) => {
    if (!request || !pasteId(request.id)) throw refusal('internal-invalid-field', 'Invalid paste upload handle');
    const upload = active.get(request.id);
    if (!upload || validatePasteState(request.state) !== upload.state)
      throw refusal('paste-stale', 'Paste upload belongs to another state or is already sealed');
    return upload;
  };
  const reserve = async (id, bytes) => {
    if (used + bytes > maxBytes) throw refusal('paste-limit', 'Paste attachments exceed journal byte budget');
    await accountBytes({ id, delta: bytes, kind: 'reserve' });
    used += bytes;
  };
  const release = async (id, bytes) => {
    used -= bytes;
    await accountBytes({ id, delta: -bytes, kind: 'release' });
  };
  const cleanup = async (upload) => {
    active.delete(upload.id);
    await upload.fd.close();
    await rm(upload.filename, { force: true });
    fileCount--;
    await syncDirectory(directory);
    await release(upload.id, upload.bytes);
  };
  return {
    begin(request) {
      return serial(async () => {
        if (active.size) throw refusal('paste-refused', 'Another paste upload is active');
        if (fileCount >= 4096) throw refusal('paste-limit', 'Paste staging has reached its 4,096 attachment limit');
        const state = validatePasteState(request?.state),
          plan = validatePastePlan({
            shape: request?.shape,
            count: request?.count,
            policy: request?.policy,
            ...(request?.conversion !== undefined ? { conversion: request.conversion } : {}),
          });
        const id = randomUUID(),
          filename = path.join(directory, id + partialSuffix),
          header = pasteHeader(plan, documentId, state);
        const bytes = Buffer.from(JSON.stringify(header) + '\n');
        await reserve(id, bytes.length);
        let fd;
        try {
          fd = await open(filename, 'wx+', 0o600);
          await fd.writeFile(bytes);
          check();
        } catch (error) {
          await fd?.close();
          await rm(filename, { force: true });
          await release(id, bytes.length);
          throw error;
        }
        fileCount++;
        active.set(id, {
          id,
          state,
          plan,
          filename,
          fd,
          bytes: bytes.length,
          count: 0,
          sequence: 0,
          lastOffset: -1,
          maxColumn: -1,
        });
        return { id, sequence: 0 };
      });
    },
    upload(request) {
      return serial(async () => {
        const upload = uploadFor(request);
        if (request.sequence !== upload.sequence)
          throw refusal('paste-bad-input', 'Paste upload sequence is not advancing');
        const page = validatePasteCells(request.cells, upload.plan.shape, upload.lastOffset);
        const cells = request.cells.map(([offset, text]) => [offset, text]);
        if (upload.count + cells.length > upload.plan.count)
          throw refusal('paste-limit', 'Paste upload exceeds declared cell count');
        const bytes = Buffer.from(cells.map((entry) => JSON.stringify(entry) + '\n').join(''));
        await reserve(upload.id, bytes.length);
        upload.bytes += bytes.length;
        try {
          check();
          await upload.fd.writeFile(bytes);
          check();
        } catch (error) {
          await cleanup(upload);
          throw error;
        }
        upload.count += cells.length;
        upload.sequence++;
        upload.lastOffset = page.lastOffset;
        for (const [offset] of cells) upload.maxColumn = Math.max(upload.maxColumn, offset % upload.plan.shape.columns);
        return { id: upload.id, sequence: upload.sequence, count: upload.count, bytes: upload.bytes };
      });
    },
    seal(request) {
      return serial(async () => {
        const upload = uploadFor(request);
        if (upload.count !== upload.plan.count) throw refusal('paste-unsupported', 'Paste upload is incomplete');
        if (
          Math.floor(upload.lastOffset / upload.plan.shape.columns) + 1 !== upload.plan.shape.rows ||
          upload.maxColumn + 1 !== upload.plan.shape.columns
        )
          throw refusal('paste-bad-input', 'Paste upload does not match its declared row and column shape');
        await upload.fd.sync();
        const proof = await fileHash(upload.fd, check);
        if (proof.bytes !== upload.bytes) throw refusal('internal-mismatch', 'Paste upload byte accounting disagrees');
        // UUID publication refuses an existing sealed attachment.
        const destination = path.join(directory, upload.id + sealedSuffix);
        try {
          await lstat(destination);
          throw refusal('paste-duplicate', 'Paste attachment already exists');
        } catch (error) {
          if (error.code !== 'ENOENT') throw error;
        }
        await upload.fd.close();
        await rename(upload.filename, destination);
        upload.filename = destination;
        await chmod(destination, 0o400);
        await syncDirectory(directory);
        active.delete(upload.id);
        return validatePasteDescriptor({
          type: 'importCells',
          id: upload.id,
          ...proof,
          ...upload.plan,
          ...(upload.plan.policy === TYPED_PASTE_POLICY ? { schema: TYPED_PASTE_SCHEMA } : {}),
        });
      });
    },
    cancel(request) {
      return serial(async () => {
        const upload = uploadFor(request);
        await cleanup(upload);
        return { cancelled: true };
      }, true);
    },
    async *pages(input, { state } = {}) {
      const descriptor = validatePasteDescriptor(input);
      validatePasteState(state);
      await work;
      check();
      ready ??= initialize();
      await ready;
      await confined();
      const filename = path.join(directory, descriptor.id + sealedSuffix),
        before = await lstat(filename);
      if (!before.isFile() || before.size !== descriptor.bytes)
        throw refusal('internal-missing-field', 'Staged paste attachment is missing, changed or exceeds its size');
      const fd = await open(filename, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0));
      try {
        const info = await fd.stat();
        if (!info.isFile() || info.ino !== before.ino || info.dev !== before.dev)
          throw refusal('paste-stale', 'Staged paste attachment changed while opening');
        const proof = await fileHash(fd, check);
        if (proof.bytes !== descriptor.bytes || proof.sha256 !== descriptor.sha256)
          throw refusal('paste-malformed', 'Staged paste attachment hash is damaged');
        const buffer = Buffer.allocUnsafe(64 * 1024),
          decoder = new TextDecoder('utf-8', { fatal: true }),
          hash = createHash('sha256');
        let position = 0,
          pending = '',
          header = false,
          cells = [],
          pageBytes = 0,
          count = 0,
          sequence = 0,
          previous = -1,
          maxColumn = -1;
        const consume = (line) => {
          if (!line || Buffer.byteLength(line) > PASTE_PAGE_BYTES)
            throw refusal('paste-malformed', 'Invalid staged paste line');
          let value;
          try {
            value = JSON.parse(line);
          } catch {
            throw refusal('paste-malformed', 'Invalid staged paste JSON');
          }
          if (!header) {
            const expected = pasteHeader(
              {
                shape: descriptor.shape,
                count: descriptor.count,
                policy: descriptor.policy,
                ...(descriptor.policy === TYPED_PASTE_POLICY ? { conversion: descriptor.conversion } : {}),
              },
              documentId,
              state
            );
            if (!equal(value, expected))
              throw refusal('internal-mismatch', 'Staged paste document, parent state or plan disagrees');
            header = true;
            return null;
          }
          const validation = validatePasteCells([value], descriptor.shape, previous);
          previous = validation.lastOffset;
          count++;
          maxColumn = Math.max(maxColumn, value[0] % descriptor.shape.columns);
          if (count > descriptor.count) throw refusal('paste-limit', 'Staged paste contains too many cells');
          return { value, bytes: validation.bytes };
        };
        for (;;) {
          check();
          const { bytesRead } = await fd.read(buffer, 0, buffer.length, position);
          if (!bytesRead) break;
          position += bytesRead;
          hash.update(buffer.subarray(0, bytesRead));
          pending += decoder.decode(buffer.subarray(0, bytesRead), { stream: true });
          let end;
          while ((end = pending.indexOf('\n')) >= 0) {
            const entry = consume(pending.slice(0, end));
            pending = pending.slice(end + 1);
            if (!entry) continue;
            if (cells.length && (cells.length >= PASTE_PAGE_CELLS || pageBytes + entry.bytes > PASTE_PAGE_BYTES)) {
              check();
              yield { sequence: sequence++, cells };
              cells = [];
              pageBytes = 0;
            }
            cells.push(entry.value);
            pageBytes += entry.bytes;
          }
          if (Buffer.byteLength(pending) > PASTE_PAGE_BYTES)
            throw refusal('paste-limit', 'Staged paste line exceeds its limit');
        }
        pending += decoder.decode();
        if (
          pending ||
          !header ||
          count !== descriptor.count ||
          Math.floor(previous / descriptor.shape.columns) + 1 !== descriptor.shape.rows ||
          maxColumn + 1 !== descriptor.shape.columns ||
          position !== descriptor.bytes ||
          hash.digest('hex') !== descriptor.sha256
        )
          throw refusal('paste-stale', 'Staged paste changed or ended before its declared cells');
        const after = await lstat(filename);
        if (
          !after.isFile() ||
          after.ino !== before.ino ||
          after.dev !== before.dev ||
          after.size !== before.size ||
          after.mtimeMs !== before.mtimeMs ||
          after.ctimeMs !== before.ctimeMs
        )
          throw refusal('paste-stale', 'Staged paste attachment changed during replay');
        check();
        if (cells.length) yield { sequence, cells };
      } finally {
        await fd.close();
      }
    },
    prune(referencedIds) {
      return serial(async () => {
        if (!Array.isArray(referencedIds) || referencedIds.some((id) => !pasteId(id)))
          throw refusal('internal-invalid-field', 'Invalid retained paste attachment list');
        const keep = new Set(referencedIds);
        let removed = 0;
        for (const name of await readdir(directory))
          if (name.endsWith(sealedSuffix) && pasteId(fileId(name)) && !keep.has(fileId(name))) {
            const filename = path.join(directory, name),
              info = await lstat(filename);
            if (!info.isFile()) throw refusal('paste-bad-input', 'Paste attachment must be a regular file');
            await rm(filename);
            fileCount--;
            await release(fileId(name), info.size);
            removed++;
          }
        if (removed) await syncDirectory(directory);
        return { removed };
      });
    },
    async close() {
      await work;
      for (const upload of active.values()) await cleanup(upload);
      closed = true;
    },
    stats() {
      return { activeUploads: active.size, attachmentBytes: used };
    },
  };
}
