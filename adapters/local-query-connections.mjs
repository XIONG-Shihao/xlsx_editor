/** Main-only bounded connection grants. Source paths never form renderer DTOs. */
import { open, lstat, realpath } from 'node:fs/promises';
import { constants } from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { atomicJson } from './storage.mjs';
import { validateLocalQuery } from '../core/local-query.mjs';
import {
  pasteId,
  validatePasteShape,
  validatePasteDescriptor,
  validatePasteState,
  pasteDescriptorIdentity,
  validateTypedPasteConversion,
} from '../core/paste-protocol.mjs';
import { isRefusal, refusal } from '../core/refusal-codes.mjs';
const MAX = 128,
  PER_DOCUMENT = 32,
  BYTES = 2 * 1024 * 1024,
  HASH = /^[a-f0-9]{64}$/;
const exact = (v, keys) =>
  v && typeof v === 'object' && !Array.isArray(v) && Object.keys(v).every((key) => keys.includes(key));
const key = (value) => {
  if (typeof value !== 'string' || value.length > 512 || !value.length || value.includes('\0'))
    throw refusal('internal-invalid-field', 'Invalid query document scope');
  return value;
};
const shape = (value) => validatePasteShape(value);
const target = (value) => {
  if (!exact(value, ['sheet', 'start'])) throw refusal('query-malformed', 'Invalid query target');
  const checked = shape({ ...value, rows: 1, columns: 1 });
  return { sheet: checked.sheet, start: checked.start };
};
const settings = (value) => {
  if (
    !exact(value, ['name', 'target', 'query']) ||
    typeof value.name !== 'string' ||
    !value.name.trim() ||
    value.name.length > 128 ||
    /[\r\n\0]/.test(value.name)
  )
    throw refusal('internal-invalid-field', 'Invalid saved query settings');
  return { name: value.name, target: target(value.target), query: validateLocalQuery(value.query) };
};
const source = (value) => {
  if (
    !exact(value, ['id', 'name', 'filename']) ||
    !pasteId(value.id) ||
    typeof value.filename !== 'string' ||
    !path.isAbsolute(value.filename) ||
    value.filename.length > 4096 ||
    value.filename.includes('\0') ||
    !['.csv', '.tsv'].includes(path.extname(value.filename).toLowerCase()) ||
    value.name !== path.basename(value.filename) ||
    value.name.length > 255
  )
    throw refusal('internal-invalid-field', 'Invalid saved source grant');
  return { ...value };
};
const output = (value) => {
  if (
    !exact(value, [
      'shape',
      'fingerprint',
      'sourceSha256',
      'sourceBytes',
      'outputRows',
      'outputColumns',
      'acceptedHead',
      'updatedAt',
      'fingerprintProfile',
      'conversion',
    ]) ||
    !HASH.test(value.fingerprint ?? '') ||
    !HASH.test(value.sourceSha256 ?? '') ||
    !Number.isSafeInteger(value.sourceBytes) ||
    value.sourceBytes < 0 ||
    value.sourceBytes > 64 * 1024 * 1024 ||
    !Number.isSafeInteger(value.updatedAt) ||
    value.updatedAt < 0 ||
    !Number.isInteger(value.outputRows) ||
    value.outputRows < 0 ||
    !Number.isInteger(value.outputColumns) ||
    value.outputColumns < 1 ||
    value.outputColumns > 256
  )
    throw refusal('internal-invalid-field', 'Invalid accepted query output');
  validatePasteState(value.acceptedHead);
  const checked = shape(value.shape);
  if (value.outputRows > checked.rows || value.outputColumns > checked.columns)
    throw refusal('query-limit', 'Query output dimensions exceed their managed shape');
  if (value.fingerprintProfile !== undefined || value.conversion !== undefined) {
    if (value.fingerprintProfile !== 'typed-semantic-v1' || value.conversion === undefined)
      throw refusal('internal-invalid-field', 'Invalid query fingerprint profile');
    const conversion = validateTypedPasteConversion(value.conversion, checked);
    if (conversion.outputRows !== value.outputRows || conversion.outputColumns !== value.outputColumns)
      throw refusal('query-limit', 'Query conversion output bounds changed');
    return { ...value, shape: checked, conversion };
  }
  return { ...value, shape: checked };
};
const pending = (value) => {
  if (!exact(value, ['id', 'descriptor', 'parentState', 'output']) || !pasteId(value.id))
    throw refusal('internal-invalid-field', 'Invalid pending query receipt');
  validatePasteState(value.parentState);
  const descriptor = validatePasteDescriptor(value.descriptor),
    result = output(value.output);
  if (
    !['text', 'typed-text-v1'].includes(descriptor.policy) ||
    (descriptor.policy === 'typed-text-v1' &&
      (result.fingerprintProfile !== 'typed-semantic-v1' ||
        JSON.stringify(descriptor.conversion) !== JSON.stringify(result.conversion))) ||
    (descriptor.policy === 'text' && result.fingerprintProfile !== undefined) ||
    descriptor.count !== descriptor.shape.rows * descriptor.shape.columns ||
    JSON.stringify(descriptor.shape) !== JSON.stringify(result.shape)
  )
    throw refusal('query-stale', 'Pending query shape changed');
  return { ...value, descriptor, output: result };
};
function record(value) {
  if (
    !exact(value, ['id', 'document', 'revision', 'name', 'target', 'query', 'source', 'last', 'pending']) ||
    !pasteId(value.id) ||
    !Number.isSafeInteger(value.revision) ||
    value.revision < 1
  )
    throw refusal('internal-invalid-field', 'Invalid query connection');
  const checked = {
    id: value.id,
    document: key(value.document),
    revision: value.revision,
    ...settings({ name: value.name, target: value.target, query: value.query }),
    source: source(value.source),
  };
  if (value.last) checked.last = output(value.last);
  if (value.pending) checked.pending = pending(value.pending);
  return checked;
}
function decode(raw) {
  const value = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(raw));
  if (
    !exact(value, ['schema', 'entries']) ||
    value.schema !== 1 ||
    !Array.isArray(value.entries) ||
    value.entries.length > MAX
  )
    throw refusal('internal-invalid-field', 'Invalid query registry');
  const seen = new Set(),
    counts = new Map();
  return value.entries.map((v) => {
    const r = record(v);
    if (
      seen.has(r.id) ||
      (counts.set(r.document, (counts.get(r.document) ?? 0) + 1), counts.get(r.document) > PER_DOCUMENT)
    )
      throw refusal('query-duplicate', 'Duplicate or excessive query connections');
    seen.add(r.id);
    return r;
  });
}
const descriptorKey = pasteDescriptorIdentity;
const queues = new Map();
function serialize(filename, work) {
  const next = (queues.get(filename) ?? Promise.resolve()).then(work),
    tail = next.then(
      () => {},
      () => {}
    );
  queues.set(filename, tail);
  void tail.then(() => {
    if (queues.get(filename) === tail) queues.delete(filename);
  });
  return next;
}
export class LocalQueryConnections {
  #filename;
  #raw = null;
  #entries = [];
  #warning;
  #locked = false;
  #queue = Promise.resolve();
  constructor(filename) {
    if (!path.isAbsolute(filename) || filename.includes('\0'))
      throw refusal('query-malformed', 'Invalid query registry path');
    this.#filename = filename;
  }
  static async open(filename) {
    const store = new LocalQueryConnections(filename);
    try {
      try {
        store.#filename = path.join(await realpath(path.dirname(filename)), path.basename(filename));
      } catch (e) {
        if (e.code !== 'ENOENT') throw e;
      }
      store.#raw = await store.#read();
      if (store.#raw) store.#entries = decode(store.#raw);
    } catch (e) {
      store.#locked = true;
      store.#warning =
        'Saved query connections could not be read. Existing bytes were preserved. ' +
        (e.code && !isRefusal(e) ? 'Registry read failed (' + e.code + ').' : String(e.message ?? e));
    }
    return store;
  }
  async #read() {
    let fd;
    try {
      fd = await open(this.#filename, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0));
    } catch (e) {
      if (e.code === 'ENOENT') return null;
      throw e;
    }
    try {
      const before = await fd.stat({ bigint: true });
      if (!before.isFile() || before.size > BigInt(BYTES)) throw refusal('query-limit', 'Query registry exceeds 2 MiB');
      const buffer = Buffer.alloc(BYTES + 1);
      let count = 0;
      while (count < buffer.length) {
        const read = await fd.read(buffer, count, buffer.length - count, count);
        if (!read.bytesRead) break;
        count += read.bytesRead;
      }
      if (count > BYTES) throw refusal('query-limit', 'Query registry exceeds 2 MiB');
      const raw = Buffer.from(buffer.subarray(0, count));
      const after = await fd.stat({ bigint: true }),
        named = await lstat(this.#filename, { bigint: true }),
        stamp = (v) => [v.dev, v.ino, v.size, v.mtimeNs, v.ctimeNs].join(':');
      if (
        named.isSymbolicLink() ||
        stamp(before) !== stamp(after) ||
        named.ino !== before.ino ||
        named.dev !== before.dev ||
        raw.length !== Number(before.size)
      )
        throw refusal('query-stale', 'Query registry changed during read');
      return raw;
    } finally {
      await fd.close();
    }
  }
  #serial(work) {
    const next = this.#queue.then(work);
    this.#queue = next.catch(() => {});
    return next;
  }
  async #write(entries) {
    return serialize(this.#filename, async () => {
      if (this.#locked) throw Error(this.#warning);
      const current = await this.#read();
      if (current === null ? this.#raw !== null : this.#raw === null || !current.equals(this.#raw)) {
        this.#locked = true;
        this.#warning =
          'Saved query connections changed on disk. Restart before updating them; existing bytes were preserved.';
        throw Error(this.#warning);
      }
      const value = { schema: 1, entries },
        raw = Buffer.from(JSON.stringify(value));
      if (raw.length > BYTES) throw refusal('query-limit', 'Query registry exceeds 2 MiB');
      await atomicJson(this.#filename, value);
      this.#raw = raw;
      this.#entries = entries;
    });
  }
  #find(document, id, revision) {
    const r = this.#entries.find((r) => r.id === id && r.document === key(document));
    if (!r || (revision !== undefined && r.revision !== revision))
      throw refusal('query-bad-input', 'Query access or exact prior definition is not granted');
    return r;
  }
  async list(document) {
    await this.#queue;
    key(document);
    return {
      entries: this.#entries.filter((r) => r.document === document).map((r) => this.public(r)),
      ...(this.#warning ? { warning: this.#warning } : {}),
    };
  }
  public(r) {
    return {
      id: r.id,
      revision: r.revision,
      name: r.name,
      target: { ...r.target },
      query: structuredClone(r.query),
      source: { id: r.source.id, name: r.source.name },
      ...(r.last ? { last: structuredClone(r.last) } : {}),
      ...(r.pending ? { pending: true } : {}),
    };
  }
  async granted(document, id, revision) {
    await this.#queue;
    return structuredClone(this.#find(document, id, revision));
  }
  create(document, draft, filename) {
    return this.#serial(async () => {
      key(document);
      if (this.#entries.length >= MAX || this.#entries.filter((r) => r.document === document).length >= PER_DOCUMENT)
        throw refusal('query-limit', 'Saved queries reached their 32 per workbook / 128 profile limit');
      const canonical = await realpath(filename),
        r = record({
          id: randomUUID(),
          document,
          revision: 1,
          ...settings(draft),
          source: { id: randomUUID(), name: path.basename(canonical), filename: canonical },
        });
      await this.#write([...this.#entries, r]);
      return this.public(r);
    });
  }
  update(document, id, revision, draft) {
    return this.#serial(async () => {
      const old = this.#find(document, id, revision);
      if (old.pending)
        throw refusal('query-bad-input', 'A pending refresh receipt must be reconciled before updating its connection');
      if (old.revision >= Number.MAX_SAFE_INTEGER)
        throw refusal('query-limit', 'Query definition revision reached its bound');
      const next = { ...old, ...settings(draft), revision: old.revision + 1 };
      if (JSON.stringify(old.target) !== JSON.stringify(next.target)) delete next.last;
      await this.#write(this.#entries.map((r) => (r === old ? next : r)));
      return this.public(next);
    });
  }
  remove(document, id, revision) {
    return this.#serial(async () => {
      const old = this.#find(document, id, revision);
      if (old.pending)
        throw refusal(
          'query-unsupported',
          'A pending refresh receipt is retained; it cannot be removed as a rejected import'
        );
      await this.#write(this.#entries.filter((r) => r !== old));
      return { removed: true };
    });
  }
  begin(document, id, revision, intent) {
    return this.#serial(async () => {
      const old = this.#find(document, id, revision);
      if (old.pending) throw refusal('query-refused', 'An uncertain refresh is awaiting its accepted history receipt');
      const next = { ...old, pending: pending(intent) };
      await this.#write(this.#entries.map((r) => (r === old ? next : r)));
      return this.public(next);
    });
  }
  finalize(document, id, intentId, receipt) {
    return this.#serial(async () => {
      const old = this.#find(document, id),
        p = old.pending;
      if (old.revision >= Number.MAX_SAFE_INTEGER)
        throw refusal('query-limit', 'Query definition revision reached its bound');
      if (
        !p ||
        p.id !== intentId ||
        !receipt?.found ||
        receipt.parentState !== p.parentState ||
        descriptorKey(validatePasteDescriptor(receipt.descriptor)) !== descriptorKey(p.descriptor)
      )
        throw refusal('query-bad-input', 'Verified accepted refresh receipt does not match the pending intent');
      const next = { ...old, revision: old.revision + 1, last: { ...p.output, acceptedHead: receipt.head } };
      delete next.pending;
      await this.#write(this.#entries.map((r) => (r === old ? next : r)));
      return this.public(next);
    });
  }
  // Only before apply was sent may a pending intent be safely rolled back.
  abandonBeforeApply(document, id, intentId) {
    return this.#serial(async () => {
      const old = this.#find(document, id);
      if (old.pending?.id !== intentId) throw refusal('query-stale', 'Pending refresh changed');
      const next = { ...old };
      delete next.pending;
      await this.#write(this.#entries.map((r) => (r === old ? next : r)));
    });
  }
}
