/** Trusted main-owned saved-file recovery bindings. Ordinary recent history
 * may be removed/evicted without removing these explicitly kept grants. */
import { open, lstat, realpath } from 'node:fs/promises';
import { constants } from 'node:fs';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { atomicJson } from './storage.mjs';
import { refusal } from '../core/refusal-codes.mjs';

const LIMIT = 128,
  MAX_BYTES = 256 * 1024,
  HASH = /^[a-f0-9]{64}$/;
const hash = (value) => createHash('sha256').update(value).digest('hex');
const supported = (filename) =>
  typeof filename === 'string' &&
  path.isAbsolute(filename) &&
  filename.length <= 4096 &&
  !filename.includes('\0') &&
  /\.(xlsx|csv|tsv)$/i.test(filename);
function entry(value) {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.keys(value).length !== 3 ||
    Object.keys(value).some((key) => !['id', 'filename', 'createdAt'].includes(key)) ||
    !supported(value.filename) ||
    !HASH.test(value.id ?? '') ||
    value.id !== hash(value.filename) ||
    !Number.isSafeInteger(value.createdAt) ||
    value.createdAt < 0
  )
    throw refusal('internal-invalid-field', 'Invalid kept recovery binding.');
  return { id: value.id, filename: value.filename, createdAt: value.createdAt };
}
function decode(raw) {
  const value = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(raw));
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.keys(value).length !== 2 ||
    value.schema !== 1 ||
    !Object.hasOwn(value, 'entries') ||
    !Array.isArray(value.entries) ||
    value.entries.length > LIMIT
  )
    throw refusal('internal-invalid-field', 'Invalid kept recovery registry.');
  const ids = new Set(),
    files = new Set();
  return value.entries.map((item) => {
    const checked = entry(item);
    if (ids.has(checked.id) || files.has(checked.filename))
      throw refusal('session-duplicate', 'Duplicate kept recovery binding.');
    ids.add(checked.id);
    files.add(checked.filename);
    return checked;
  });
}

// Serialize mutations across multiple readers of the same registry in this
// process, then compare exact prior bytes immediately before atomicJson.
const writes = new Map();
function serializeFile(filename, work) {
  const next = (writes.get(filename) ?? Promise.resolve()).then(work),
    tail = next.then(
      () => {},
      () => {}
    );
  writes.set(filename, tail);
  void tail.then(() => {
    if (writes.get(filename) === tail) writes.delete(filename);
  });
  return next;
}
export class KeptRecoveryBindings {
  #filename;
  #entries = [];
  #raw = null;
  #warning;
  #locked = false;
  #queue = Promise.resolve();
  constructor(filename) {
    if (typeof filename !== 'string' || !path.isAbsolute(filename) || filename.includes('\0'))
      throw refusal('session-malformed', 'Invalid kept recovery registry path.');
    this.#filename = path.resolve(filename);
  }
  static async open(filename) {
    const bindings = new KeptRecoveryBindings(filename);
    await bindings.#load();
    return bindings;
  }
  async #read() {
    let file;
    try {
      file = await open(this.#filename, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0));
    } catch (error) {
      if (error.code === 'ENOENT') return null;
      throw error;
    }
    try {
      const before = await file.stat({ bigint: true });
      if (!before.isFile() || before.size > BigInt(MAX_BYTES))
        throw refusal('session-limit', 'Kept recovery registry exceeds its size bound.');
      const buffer = Buffer.alloc(MAX_BYTES + 1);
      let count = 0;
      while (count < buffer.length) {
        const { bytesRead } = await file.read(buffer, count, buffer.length - count, count);
        if (!bytesRead) break;
        count += bytesRead;
      }
      if (count > MAX_BYTES) throw refusal('session-limit', 'Kept recovery registry exceeds its size bound.');
      const after = await file.stat({ bigint: true }),
        named = await lstat(this.#filename, { bigint: true });
      const stamp = (info) => [info.dev, info.ino, info.size, info.mtimeNs, info.ctimeNs].join(':');
      if (
        named.isSymbolicLink() ||
        named.dev !== before.dev ||
        named.ino !== before.ino ||
        stamp(before) !== stamp(after) ||
        BigInt(count) !== before.size
      )
        throw refusal('session-stale', 'Kept recovery registry changed during read.');
      return Buffer.from(buffer.subarray(0, count));
    } finally {
      await file.close();
    }
  }
  async #load() {
    try {
      // Normalize trusted system directory aliases without following a registry
      // file symlink. A missing parent can be created by atomicJson on first pin.
      try {
        this.#filename = path.join(await realpath(path.dirname(this.#filename)), path.basename(this.#filename));
      } catch (error) {
        if (error.code !== 'ENOENT') throw error;
      }
      this.#raw = await this.#read();
      if (this.#raw !== null) this.#entries = decode(this.#raw);
    } catch (error) {
      this.#locked = true;
      this.#warning =
        'Kept recovery bindings could not be read. Existing bytes were preserved; repair the registry and restart before updating it. ' +
        String(error.message ?? error);
    }
  }
  #serial(work) {
    const next = this.#queue.then(work);
    this.#queue = next.catch(() => {});
    return next;
  }
  async #write(entries) {
    return serializeFile(this.#filename, async () => {
      if (this.#locked) throw Error(this.#warning);
      try {
        const current = await this.#read();
        if (current === null ? this.#raw !== null : this.#raw === null || !current.equals(this.#raw)) {
          this.#locked = true;
          throw refusal('session-stale', 'Kept recovery registry changed on disk; restart before updating it.');
        }
        const raw = Buffer.from(JSON.stringify({ schema: 1, entries }));
        if (raw.length > MAX_BYTES)
          throw Object.assign(
            Error('Kept recovery bindings exceed their 256 KiB bound. Save or discard recovery before keeping more.'),
            { code: 'kept-recovery-full' }
          );
        if (JSON.stringify(entries) === JSON.stringify(this.#entries)) return;
        await atomicJson(this.#filename, { schema: 1, entries });
        this.#raw = raw;
        this.#entries = entries;
        this.#warning = undefined;
      } catch (error) {
        if (this.#locked || error?.code !== 'kept-recovery-full') {
          this.#locked = true;
          this.#warning =
            'Kept recovery bindings could not be updated. Existing bytes were preserved; restart before updating them. ' +
            String(error.message ?? error);
        }
        throw Error(this.#warning ?? String(error.message ?? error));
      }
    });
  }
  async list() {
    await this.#queue;
    return {
      entries: this.#entries.map((value) => ({ ...value })),
      ...(this.#warning ? { warning: this.#warning } : {}),
    };
  }
  pin(filename, id) {
    return this.#serial(async () => {
      const checked = entry({ id, filename, createdAt: Date.now() }),
        existing = this.#entries.find((value) => value.id === id);
      if (!existing && this.#entries.length >= LIMIT)
        throw refusal(
          'session-limit',
          'Kept recovery bindings reached their 128-document limit. Save or discard recovery before keeping more.'
        );
      await this.#write(existing ? this.#entries : [...this.#entries, checked]);
      return { ...(existing ?? checked) };
    });
  }
  remove(id) {
    return this.#serial(async () => {
      if (typeof id !== 'string' || !HASH.test(id))
        throw refusal('internal-invalid-field', 'Invalid kept recovery handle.');
      const entries = this.#entries.filter((value) => value.id !== id),
        removed = entries.length !== this.#entries.length;
      await this.#write(entries);
      return removed;
    });
  }
}
