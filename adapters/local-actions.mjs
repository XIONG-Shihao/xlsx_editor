/** Main-only private-profile persistence. Descriptors never execute or open a workbook. */
import fs from 'node:fs/promises';
import { constants } from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import {
  LOCAL_ACTION_LIMITS,
  validateLocalActionDraft,
  validateLocalActionList,
  localActionId,
  localActionRevision,
} from '../core/local-actions.mjs';
import { refusal } from '../core/refusal-codes.mjs';
const queues = new Map(),
  stamp = (value) => [value.dev, value.ino, value.size, value.mtimeNs, value.ctimeNs, value.nlink].join(':');
function serialize(filename, work) {
  const result = (queues.get(filename) ?? Promise.resolve()).then(work),
    tail = result.then(
      () => {},
      () => {}
    );
  queues.set(filename, tail);
  void tail.then(() => {
    if (queues.get(filename) === tail) queues.delete(filename);
  });
  return result;
}
const unavailable =
  'Private local actions could not be read. Existing bytes were preserved; reopen the registry before editing.';
const conflict =
  'Private local actions changed on disk. Existing bytes were preserved; reopen the registry before editing.';
export class LocalActions {
  #filename;
  #directory;
  #directoryIdentity;
  #snapshot = null;
  #value = { schema: 1, scope: 'private-profile', revision: 0, entries: [] };
  #warning;
  #locked = false;
  #ready = false;
  #queue = Promise.resolve();
  constructor(filename) {
    if (
      typeof filename !== 'string' ||
      !path.isAbsolute(filename) ||
      filename.length > 4096 ||
      filename.includes('\0') ||
      filename.endsWith(path.sep)
    )
      throw refusal('query-malformed', 'Invalid trusted private action registry path');
    this.#filename = filename;
    this.#directory = path.dirname(filename);
  }
  static async open(filename) {
    const store = new LocalActions(filename);
    try {
      await fs.mkdir(store.#directory, { recursive: true, mode: 0o700 });
      store.#directory = await fs.realpath(store.#directory);
      store.#filename = path.join(store.#directory, path.basename(filename));
      const parent = await fs.lstat(store.#directory, { bigint: true });
      if (
        !parent.isDirectory() ||
        parent.isSymbolicLink() ||
        (typeof process.getuid === 'function' && parent.uid !== BigInt(process.getuid()))
      )
        throw refusal('query-unsupported', 'Unsafe private registry directory');
      store.#directoryIdentity = [parent.dev, parent.ino].join(':');
      store.#snapshot = await store.#read();
      if (store.#snapshot) store.#value = store.#decode(store.#snapshot.raw);
    } catch {
      store.#warning = unavailable;
      store.#locked = true;
    }
    store.#ready = true;
    return store;
  }
  #decode(raw) {
    const parsed = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(raw));
    const value = validateLocalActionList(parsed);
    if (Object.hasOwn(parsed, 'warning') || !Buffer.from(JSON.stringify(value)).equals(raw))
      throw refusal('query-unsupported', 'Private action registry is not a canonical stored record');
    return value;
  }
  async #parent() {
    const parent = await fs.lstat(this.#directory, { bigint: true });
    if (
      !parent.isDirectory() ||
      parent.isSymbolicLink() ||
      [parent.dev, parent.ino].join(':') !== this.#directoryIdentity
    )
      throw refusal('query-stale', 'Private registry directory changed');
  }
  async #read() {
    await this.#parent();
    let handle;
    try {
      handle = await fs.open(this.#filename, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0));
    } catch (error) {
      if (error.code === 'ENOENT') return null;
      throw error;
    }
    try {
      const before = await handle.stat({ bigint: true });
      if (
        !before.isFile() ||
        before.nlink !== 1n ||
        (typeof process.getuid === 'function' && before.uid !== BigInt(process.getuid())) ||
        before.size > BigInt(LOCAL_ACTION_LIMITS.registryBytes)
      )
        throw refusal('query-limit', 'Unsafe or oversized action registry');
      const raw = Buffer.alloc(Number(before.size) + 1);
      let count = 0;
      while (count < raw.length) {
        const read = await handle.read(raw, count, raw.length - count, count);
        if (!read.bytesRead) break;
        count += read.bytesRead;
      }
      const after = await handle.stat({ bigint: true }),
        named = await fs.lstat(this.#filename, { bigint: true });
      if (
        named.isSymbolicLink() ||
        stamp(before) !== stamp(after) ||
        stamp(after) !== stamp(named) ||
        count !== Number(before.size)
      )
        throw refusal('query-stale', 'Action registry changed during bounded read');
      return { raw: Buffer.from(raw.subarray(0, count)), stamp: stamp(after) };
    } finally {
      await handle.close();
    }
  }
  async #unchanged() {
    try {
      const current = await this.#read();
      if (
        current === null
          ? this.#snapshot !== null
          : this.#snapshot === null || current.stamp !== this.#snapshot.stamp || !current.raw.equals(this.#snapshot.raw)
      )
        throw refusal('query-stale', 'Action registry changed');
    } catch {
      this.#locked = true;
      this.#warning = conflict;
      throw refusal('query-refused', conflict);
    }
  }
  #serial(work) {
    if (!this.#ready)
      return Promise.reject(refusal('query-unsupported', 'Open the private action registry before using it'));
    const pending = this.#queue.then(work);
    this.#queue = pending.catch(() => {});
    return pending;
  }
  #public() {
    return validateLocalActionList({ ...this.#value, ...(this.#warning ? { warning: this.#warning } : {}) });
  }
  async list() {
    if (!this.#ready) throw refusal('query-unsupported', 'Open the private action registry before using it');
    await this.#queue;
    if (!this.#locked) await this.#unchanged().catch(() => {});
    return this.#public();
  }
  async #write(next) {
    const value = validateLocalActionList(next),
      raw = Buffer.from(JSON.stringify(value));
    if (raw.length > LOCAL_ACTION_LIMITS.registryBytes)
      throw refusal('query-limit', 'Action registry exceeds its byte budget');
    return serialize(this.#filename, async () => {
      if (this.#locked) throw Error(this.#warning);
      await this.#unchanged();
      const lock = this.#filename + '.lock';
      let lockHandle;
      try {
        lockHandle = await fs.open(
          lock,
          constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL | (constants.O_NOFOLLOW ?? 0),
          0o600
        );
      } catch (error) {
        if (error.code === 'EEXIST')
          throw refusal(
            'session-busy',
            'Private action registry is being updated; retry with a current registry revision'
          );
        throw error;
      }
      const lockInfo = await lockHandle.stat({ bigint: true }),
        temporary = this.#filename + '.' + randomUUID() + '.tmp';
      let published = false;
      try {
        await this.#unchanged();
        const handle = await fs.open(
          temporary,
          constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL | (constants.O_NOFOLLOW ?? 0),
          0o600
        );
        try {
          await handle.writeFile(raw);
          await handle.sync();
        } finally {
          await handle.close();
        }
        await this.#unchanged();
        await fs.rename(temporary, this.#filename);
        published = true;
        this.#value = value;
        this.#snapshot = await this.#read();
        if (!this.#snapshot.raw.equals(raw)) throw refusal('query-stale', 'Action registry changed after publication');
        const directory = await fs.open(this.#directory, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0));
        try {
          if (process.platform !== 'win32') await directory.sync();
        } finally {
          await directory.close();
        }
        return this.#public();
      } catch (error) {
        if (published) {
          this.#locked = true;
          this.#warning =
            'Private action publication could not be confirmed. Existing bytes were preserved; reopen before retrying.';
        }
        throw error;
      } finally {
        await fs.rm(temporary, { force: true });
        await lockHandle.close();
        const named = await fs.lstat(lock, { bigint: true }).catch(() => null);
        if (named && named.dev === lockInfo.dev && named.ino === lockInfo.ino && !named.isSymbolicLink())
          await fs.unlink(lock);
      }
    });
  }
  #expected(value) {
    const revision = localActionRevision(value);
    if (this.#locked) throw Error(this.#warning);
    if (revision !== this.#value.revision) throw refusal('query-stale', 'Stale local action registry revision');
    if (revision === Number.MAX_SAFE_INTEGER)
      throw refusal('query-limit', 'Action registry revision reached its bound');
    return revision;
  }
  create(draft, expectedRegistryRevision) {
    const captured = validateLocalActionDraft(draft),
      expected = localActionRevision(expectedRegistryRevision);
    return this.#serial(async () => {
      const revision = this.#expected(expected);
      if (this.#value.entries.length >= 128)
        throw refusal('query-limit', 'Private local actions reached their 128-entry limit');
      const now = Date.now(),
        entry = { id: randomUUID(), revision: 1, ...captured, author: 'local-profile', createdAt: now, updatedAt: now };
      return this.#write({ ...this.#value, revision: revision + 1, entries: [...this.#value.entries, entry] });
    });
  }
  update(id, draft, expectedRegistryRevision) {
    const key = localActionId(id),
      captured = validateLocalActionDraft(draft),
      expected = localActionRevision(expectedRegistryRevision);
    return this.#serial(async () => {
      const revision = this.#expected(expected),
        prior = this.#value.entries.find((entry) => entry.id === key);
      if (!prior) throw refusal('query-not-found', 'Local action does not exist');
      if (prior.revision === Number.MAX_SAFE_INTEGER)
        throw refusal('query-limit', 'Local action entry revision reached its bound');
      const entry = {
        ...prior,
        ...captured,
        revision: prior.revision + 1,
        updatedAt: Math.max(Date.now(), prior.updatedAt),
      };
      return this.#write({
        ...this.#value,
        revision: revision + 1,
        entries: this.#value.entries.map((old) => (old === prior ? entry : old)),
      });
    });
  }
  remove(id, expectedRegistryRevision) {
    const key = localActionId(id),
      expected = localActionRevision(expectedRegistryRevision);
    return this.#serial(async () => {
      const revision = this.#expected(expected);
      if (!this.#value.entries.some((entry) => entry.id === key))
        throw refusal('query-not-found', 'Local action does not exist');
      return this.#write({
        ...this.#value,
        revision: revision + 1,
        entries: this.#value.entries.filter((entry) => entry.id !== key),
      });
    });
  }
}
