import { createHash, randomUUID } from 'node:crypto';
import { constants, createReadStream, createWriteStream } from 'node:fs';
import {
  copyFile,
  mkdir,
  open,
  opendir,
  readFile,
  readdir,
  realpath,
  rename,
  rm,
  stat,
  lstat,
  link,
} from 'node:fs/promises';
import path from 'node:path';
import { pipeline } from 'node:stream/promises';
import { refusal } from '../core/refusal-codes.mjs';

export const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex');
export const validHash = (value) => typeof value === 'string' && /^[a-f0-9]{64}$/.test(value);
export const validId = (value) => typeof value === 'string' && /^[a-f0-9-]{36}$/.test(value);
export async function revision(filename) {
  const hash = createHash('sha256');
  for await (const chunk of createReadStream(filename)) hash.update(chunk);
  return hash.digest('hex');
}
export async function syncDirectory(directory) {
  if (process.platform === 'win32') return;
  const fd = await open(directory, 'r');
  try {
    await fd.sync();
  } finally {
    await fd.close();
  }
}
export async function syncFile(filename) {
  const fd = await open(filename, constants.O_RDONLY | constants.O_NOFOLLOW);
  try {
    await fd.sync();
  } finally {
    await fd.close();
  }
}
export async function atomicJson(filename, data) {
  await mkdir(path.dirname(filename), { recursive: true, mode: 0o700 });
  const temporary = filename + '.' + randomUUID() + '.tmp';
  try {
    const fd = await open(temporary, 'wx', 0o600);
    try {
      await fd.writeFile(JSON.stringify(data));
      await fd.sync();
    } finally {
      await fd.close();
    }
    await rename(temporary, filename);
    await syncDirectory(path.dirname(filename));
  } finally {
    await rm(temporary, { force: true });
  }
}
/** Exclusive private immutable base; never publishes a user-visible destination. */
export async function writePrivateSnapshot(filename, bytes) {
  const fd = await open(filename, 'wx', 0o600);
  try {
    await fd.writeFile(bytes);
    await fd.sync();
  } finally {
    await fd.close();
  }
  await syncDirectory(path.dirname(filename));
}
export async function readJson(filename, maxBytes = 2 * 1024 * 1024) {
  const info = await lstat(filename);
  if (!info.isFile() || info.size > maxBytes) throw refusal('session-limit', 'Recovery metadata exceeds its limit.');
  return JSON.parse(await readFile(filename, 'utf8'));
}
export async function copySnapshot(source, destination) {
  const before = await revision(source);
  await copyFile(source, destination, constants.COPYFILE_EXCL);
  await syncFile(destination);
  const copied = await revision(destination);
  if (before !== copied || before !== (await revision(source))) {
    await rm(destination, { force: true });
    throw refusal('session-stale', 'Source changed while opening. Try again.');
  }
  return copied;
}
export async function stageCopy(source, destination) {
  const temporary = path.join(
    path.dirname(destination),
    '.' + path.basename(destination) + '.' + randomUUID() + '.tmp'
  );
  try {
    await copyFile(source, temporary, constants.COPYFILE_EXCL);
    await syncFile(temporary);
    return temporary;
  } catch (error) {
    await rm(temporary, { force: true });
    throw error;
  }
}
export async function assertRevision(filename, expected) {
  const conflict = () => {
    const error = refusal('session-stale', 'The file changed on disk. Save a copy to preserve your edits.');
    error.code = 'REVISION_CONFLICT';
    return error;
  };
  try {
    if (!validHash(expected) || (await realpath(filename)) !== filename || (await revision(filename)) !== expected)
      throw conflict();
  } catch (error) {
    if (['ENOENT', 'ENOTDIR'].includes(error.code)) throw conflict();
    throw error;
  }
}
export function validateDestinationIdentity(value) {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.keys(value).sort().join(',') !== 'dev,ino' ||
    Object.values(value).some((item) => typeof item !== 'string' || !/^(?:0|[1-9]\d*)$/.test(item) || item.length > 32)
  )
    throw refusal('session-malformed', 'Invalid staged publication identity.');
  return value;
}
export async function destinationIdentity(filename) {
  const info = await lstat(filename, { bigint: true });
  if (!info.isFile() || info.isSymbolicLink())
    throw refusal('session-bad-input', 'Publication destination is not a regular file.');
  return validateDestinationIdentity({ dev: String(info.dev), ino: String(info.ino) });
}
/** The optional callback sees staged-file identity/revision before mutation. */
export async function publishCopy(source, destination, expected, beforePublish) {
  const filename = path.join(await realpath(path.dirname(destination)), path.basename(destination));
  const temporary = await stageCopy(source, filename);
  try {
    if (beforePublish !== undefined) {
      if (typeof beforePublish !== 'function') throw refusal('internal-invalid-field', 'Invalid publication callback.');
      await beforePublish(await destinationIdentity(temporary), await revision(temporary));
    }
    if (expected === null)
      await link(temporary, filename); // exclusive creation
    else {
      await assertRevision(filename, expected);
      await rename(temporary, filename);
    }
    await syncDirectory(path.dirname(filename));
  } finally {
    await rm(temporary, { force: true });
  }
}
export async function directoryBytes(directory) {
  let size = 0;
  for (const item of await readdir(directory, { withFileTypes: true })) {
    if (item.isFile()) size += (await stat(path.join(directory, item.name))).size;
    else if (item.isDirectory()) size += await directoryBytes(path.join(directory, item.name));
    else throw refusal('session-malformed', 'Unexpected recovery filesystem entry.');
  }
  return size;
}
/** Stream source byte spans and replacement bytes without assembling the file. */
export async function writePatchedFile(source, destination, patches) {
  let position = 0;
  const size = (await stat(source)).size;
  async function* parts() {
    for (const patch of patches) {
      if (patch.start < position || patch.end < patch.start || patch.end > size)
        throw refusal('internal-invalid-field', 'Invalid CSV byte patch.');
      if (patch.start > position) yield* createReadStream(source, { start: position, end: patch.start - 1 });
      yield patch.bytes;
      position = patch.end;
    }
    if (position < size) yield* createReadStream(source, { start: position });
  }
  try {
    await pipeline(parts(), createWriteStream(destination, { flags: 'wx', mode: 0o600 }));
    await syncFile(destination);
  } catch (error) {
    await rm(destination, { force: true });
    throw error;
  }
}

export const RECOVERY_DRAFT_LIMIT = 128,
  RECOVERY_MANIFEST_LIMIT = 64 * 1024;
// A macro-enabled workbook is the same package as an ordinary one with its macro part
// carried along, so both are edited by the indexed workbook and keep their own extension.
export const isOoxmlExtension = (extension) => extension === '.xlsx' || extension === '.xlsm';
const filePath = (value) =>
  typeof value === 'string' && value.length <= 4096 && !value.includes('\0') && path.isAbsolute(value);
export function validateSavedSource(saved) {
  if (
    !saved ||
    typeof saved !== 'object' ||
    Array.isArray(saved) ||
    Object.keys(saved)
      .filter((key) => key !== 'calculation')
      .sort()
      .join(',') !== 'bytes,head,headHash,revision,schema,snapshot' ||
    saved.schema !== 1 ||
    !validId(saved.snapshot) ||
    !validHash(saved.revision) ||
    !Number.isSafeInteger(saved.bytes) ||
    saved.bytes < 1 ||
    saved.bytes > 512 * 1024 * 1024 ||
    (saved.head === null ? saved.headHash !== null : !validId(saved.head) || !validHash(saved.headHash))
  )
    throw refusal('internal-invalid-field', 'Invalid saved-source recovery checkpoint.');
  if (Object.hasOwn(saved, 'calculation')) {
    const item = saved.calculation;
    if (
      !item ||
      typeof item !== 'object' ||
      Array.isArray(item) ||
      Object.keys(item).sort().join(',') !== 'bytes,revision,schema,snapshot' ||
      item.schema !== 1 ||
      !validId(item.snapshot) ||
      !validHash(item.revision) ||
      !Number.isSafeInteger(item.bytes) ||
      item.bytes < 1 ||
      item.bytes > 256 * 1024 * 1024
    )
      throw refusal('internal-invalid-field', 'Invalid saved-source calculation checkpoint.');
  }
  return saved;
}
export function validateRecoveryManifest(saved, documentId) {
  if (saved?.initialCalculationPolicy !== undefined) {
    const p = saved.initialCalculationPolicy;
    if (
      !p ||
      typeof p !== 'object' ||
      Array.isArray(p) ||
      Object.keys(p).sort().join(',') !== 'mode,recalculateBeforeSave' ||
      !['auto', 'manual'].includes(p.mode) ||
      typeof p.recalculateBeforeSave !== 'boolean'
    )
      throw refusal('internal-invalid-field', 'Invalid initial calculation policy in recovery manifest');
  }
  if (
    !saved ||
    ![1, 2].includes(saved.schema) ||
    !['.xlsx', '.xlsm', '.xls', '.xlsb', '.csv', '.tsv'].includes(saved.extension) ||
    !validId(saved.snapshot) ||
    !validHash(saved.baseRevision) ||
    (saved.head !== null && !validId(saved.head)) ||
    (saved.savedHead !== null && !validId(saved.savedHead)) ||
    (saved.head === null ? saved.headHash !== null : !validHash(saved.headHash))
  )
    throw refusal('internal-invalid-field', 'Invalid spreadsheet recovery manifest.');
  if (
    ['.xls', '.xlsb'].includes(saved.extension) &&
    (saved.head !== null || saved.savedHead !== null || saved.pendingPublication)
  )
    throw refusal('session-unsupported', 'Passive workbook recovery must contain an unchanged source snapshot.');
  if (saved.schema === 1) {
    if (!validHash(saved.sourceRevision))
      throw refusal('internal-invalid-field', 'Invalid spreadsheet recovery manifest.');
  } else if (
    !validId(saved.documentId) ||
    (documentId && saved.documentId !== documentId) ||
    typeof saved.name !== 'string' ||
    saved.name.length < 1 ||
    saved.name.length > 80 ||
    !Number.isSafeInteger(saved.createdAt) ||
    saved.createdAt < 0 ||
    (saved.filename !== null &&
      (!filePath(saved.filename) || path.extname(saved.filename).toLowerCase() !== saved.extension)) ||
    (saved.filename === null
      ? saved.sourceRevision !== null || saved.extension !== '.xlsx'
      : !validHash(saved.sourceRevision)) ||
    typeof saved.prepared !== 'boolean'
  )
    throw refusal('session-malformed', 'Invalid unsaved workbook recovery identity.');
  if (saved.pendingPublication) {
    const pending = saved.pendingPublication;
    if (!validHash(pending.revision) || pending.head !== saved.head)
      throw refusal('internal-invalid-field', 'Invalid pending save receipt.');
    if (
      (pending.kind !== undefined && pending.kind !== 'bind') ||
      (saved.schema === 2 && saved.filename === null && pending.kind !== 'bind')
    )
      throw refusal('internal-invalid-field', 'Invalid pending save receipt kind.');
    if (
      pending.kind === 'bind' &&
      (saved.schema !== 2 ||
        saved.filename !== null ||
        !filePath(pending.destination) ||
        path.extname(pending.destination).toLowerCase() !== '.xlsx' ||
        (pending.expectedRevision !== null && !validHash(pending.expectedRevision)))
    )
      throw refusal('internal-invalid-field', 'Invalid first-save binding receipt.');
    if (pending.savedSource) {
      const source = validateSavedSource(pending.savedSource);
      if (
        !isOoxmlExtension(saved.extension) ||
        source.head !== pending.head ||
        source.headHash !== saved.headHash ||
        source.revision !== pending.revision
      )
        throw refusal('session-stale', 'Pending saved source belongs to another accepted head.');
      if (pending.destinationIdentity === undefined)
        throw refusal('session-bad-input', 'Saved-source intent requires staged publication identity.');
    }
    if (pending.destinationIdentity !== undefined) validateDestinationIdentity(pending.destinationIdentity);
  }
  if (saved.savedSources !== undefined) {
    if (!isOoxmlExtension(saved.extension) || !Array.isArray(saved.savedSources) || saved.savedSources.length > 102)
      throw refusal('internal-invalid-field', 'Invalid saved-source recovery inventory.');
    const heads = new Set();
    for (const source of saved.savedSources) {
      validateSavedSource(source);
      const head = source.head ?? 'base';
      if (heads.has(head)) throw refusal('session-duplicate', 'Duplicate saved-source accepted head.');
      heads.add(head);
    }
  }
  if (saved.history) {
    const history = saved.history;
    if (
      !Array.isArray(history.frames) ||
      history.frames.length < 1 ||
      history.frames.length > 101 ||
      !Number.isInteger(history.cursor) ||
      history.cursor < 0 ||
      history.cursor >= history.frames.length
    )
      throw refusal('internal-invalid-field', 'Invalid recovery undo history.');
    const heads = new Set();
    for (const frame of history.frames) {
      if (
        !frame ||
        (frame.head === null ? frame.hash !== null : !validId(frame.head) || !validHash(frame.hash)) ||
        heads.has(frame.head)
      )
        throw refusal('internal-invalid-field', 'Invalid recovery undo frame.');
      heads.add(frame.head);
    }
    if (history.frames[history.cursor].head !== saved.head || history.frames[history.cursor].hash !== saved.headHash)
      throw refusal('session-bad-input', 'Recovery cursor does not match the accepted state.');
  }
  return saved;
}
export async function discardSpreadsheetRecovery(recoveryDirectory, id, { preparedOnly = false } = {}) {
  if (!validId(id)) throw refusal('internal-invalid-field', 'Invalid workbook recovery handle.');
  const directory = path.join(path.resolve(recoveryDirectory), 'draft-' + id),
    info = await lstat(directory);
  if (!info.isDirectory()) throw refusal('internal-invalid-field', 'Invalid workbook recovery directory.');
  let manifest;
  try {
    manifest = validateRecoveryManifest(
      await readJson(path.join(directory, 'current.json'), RECOVERY_MANIFEST_LIMIT),
      id
    );
  } catch (error) {
    if (!preparedOnly || error.code !== 'ENOENT') throw error;
  }
  const nonemptyHistory =
    manifest?.history &&
    (manifest.history.cursor !== 0 || manifest.history.frames.length !== 1 || manifest.history.frames[0].head !== null);
  if (
    manifest &&
    (manifest.schema !== 2 ||
      (preparedOnly &&
        (!manifest.prepared ||
          manifest.head !== null ||
          manifest.savedHead !== null ||
          manifest.pendingPublication ||
          nonemptyHistory)))
  )
    throw refusal('session-bad-input', 'This recovery record is not an untouched prepared candidate.');
  if (manifest && !preparedOnly) {
    const snapshot = path.join(directory, manifest.snapshot + manifest.extension),
      snapshotInfo = await lstat(snapshot);
    if (
      !snapshotInfo.isFile() ||
      snapshotInfo.size > 512 * 1024 * 1024 ||
      (await revision(snapshot)) !== manifest.baseRevision
    )
      throw refusal('session-malformed', 'Recovery snapshot is damaged. Existing data was preserved.');
    const tip = manifest.history?.frames.at(-1) ?? { head: manifest.head, hash: manifest.headHash },
      frames = [],
      seen = new Set();
    let head = tip.head,
      hash = tip.hash,
      bytes = 0;
    while (head !== null) {
      if (!validId(head) || seen.has(head)) throw refusal('internal-invalid-field', 'Invalid recovery command chain.');
      seen.add(head);
      const file = path.join(directory, head + '.json'),
        nodeInfo = await lstat(file);
      if (!nodeInfo.isFile() || nodeInfo.size > 4 * 1024 * 1024 + 4096 || (bytes += nodeInfo.size) > 64 * 1024 * 1024)
        throw refusal('session-limit', 'Recovery command exceeds its limit.');
      const data = await readFile(file);
      if (sha256(data) !== hash)
        throw refusal('session-malformed', 'Recovery commands are damaged. Existing data was preserved.');
      const node = JSON.parse(data);
      if (
        node.schema !== 1 ||
        typeof node.commands !== 'string' ||
        (node.parent !== null && (!validId(node.parent) || !validHash(node.parentHash)))
      )
        throw refusal('internal-invalid-field', 'Invalid recovery command record.');
      frames.push({ head, hash });
      head = node.parent;
      hash = node.parentHash;
    }
    frames.push({ head: null, hash: null });
    if (manifest.history && JSON.stringify(frames.reverse().slice(-101)) !== JSON.stringify(manifest.history.frames))
      throw refusal('session-bad-input', 'Recovery undo frames do not match the verified command chain.');
  }
  await rm(directory, { recursive: true, force: true });
  await syncDirectory(path.resolve(recoveryDirectory));
}
export async function listDraftRecords(recoveryDirectory) {
  const root = path.resolve(recoveryDirectory),
    records = [],
    warnings = [];
  let directory,
    count = 0;
  // Some non-Node hosts defer opendir errors until iteration. Detect a missing
  // root here; later disappearance must still fail rather than return partial data.
  try {
    await lstat(root);
    directory = await opendir(root);
  } catch (error) {
    if (error.code === 'ENOENT') return { records, warnings, count };
    throw error;
  }
  for await (const entry of directory) {
    if (!entry.name.startsWith('draft-')) continue;
    if (++count > RECOVERY_DRAFT_LIMIT)
      throw refusal(
        'session-limit',
        'Workbook recovery discovery exceeds its 128-document limit. Existing data was preserved.'
      );
    let hasManifest = false;
    try {
      const id = entry.name.slice(6);
      if (!entry.isDirectory() || !validId(id)) throw refusal('internal-invalid-field', 'Invalid draft directory.');
      const owned = path.join(root, entry.name),
        manifest = validateRecoveryManifest(
          await readJson(path.join(owned, 'current.json'), RECOVERY_MANIFEST_LIMIT),
          id
        );
      hasManifest = true;
      if (manifest.schema !== 2) throw refusal('internal-invalid-field', 'Invalid draft schema.');
      const snapshot = path.join(owned, manifest.snapshot + manifest.extension),
        info = await lstat(snapshot);
      if (!info.isFile() || info.size > 512 * 1024 * 1024)
        throw refusal('internal-invalid-field', 'Invalid draft source snapshot.');
      records.push({ id, directory: owned, manifest });
    } catch (error) {
      if (hasManifest || error.code !== 'ENOENT')
        warnings.push(
          'A workbook recovery record could not be read. Existing recovery data was preserved. ' +
            String(error.message ?? error)
        );
    }
  }
  return { records, warnings, count };
}
/** Bounded metadata discovery. Snapshot checksum and journal hashes are verified
 * by recover() before an owner can be activated or edited. */
export async function listSpreadsheetRecovery(recoveryDirectory) {
  const { records, warnings } = await listDraftRecords(recoveryDirectory),
    bindings = new Map();
  for (const { manifest: m } of records)
    if (!m.prepared && m.filename) bindings.set(m.filename, (bindings.get(m.filename) ?? 0) + 1);
  if ([...bindings.values()].some((count) => count > 1))
    warnings.push(
      'Multiple workbook recovery records refer to the same file. Recover a specific record; their data was preserved.'
    );
  const unsavedTip = (m) => (m.history ? m.history.frames.at(-1).head : m.head) !== m.savedHead;
  return {
    entries: records
      .filter(
        ({ manifest: m }) =>
          !m.prepared &&
          (m.filename === null ||
            m.head !== m.savedHead ||
            unsavedTip(m) ||
            m.pendingPublication ||
            bindings.get(m.filename) > 1)
      )
      .map(({ id, manifest: m }) => ({
        id,
        name: m.filename ? path.basename(m.filename) : m.name,
        unsaved: m.filename === null,
        dirty: m.head !== m.savedHead,
        createdAt: m.createdAt,
        ...(m.filename ? { displayPath: m.filename } : {}),
      })),
    warnings,
  };
}
