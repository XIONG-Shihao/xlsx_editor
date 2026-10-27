import { randomUUID } from 'node:crypto';
import { readFile, realpath, stat } from 'node:fs/promises';
import path from 'node:path';
import { atomicJson } from '../adapters/storage.mjs';

const LIMIT = 20,
  MAX_BYTES = 256 * 1024;
const validId = (value) => typeof value === 'string' && /^[a-f0-9]{8}(?:-[a-f0-9]{4}){3}-[a-f0-9]{12}$/.test(value);
const supported = (filename) =>
  typeof filename === 'string' &&
  path.isAbsolute(filename) &&
  filename.length <= 4096 &&
  !filename.includes('\0') &&
  /\.(xlsx|xlsm|xls|xlsb|csv|tsv)$/i.test(filename);
const failure = (error) =>
  error?.code === 'ENOENT' || error?.code === 'ENOTDIR' ? 'File is missing.' : 'File is unavailable.';

/** The main process grants paths by recording successful opens. Renderer callers
 * receive opaque IDs; accepting an ID never accepts a renderer-supplied path. */
export class RecentFiles {
  #filename;
  #entries = [];
  #raw = null;
  #warning;
  #locked = false;
  #queue = Promise.resolve();
  constructor(filename) {
    this.#filename = filename;
  }
  static async open(filename) {
    const recent = new RecentFiles(filename);
    await recent.#load();
    return recent;
  }
  async #read() {
    try {
      const info = await stat(this.#filename);
      if (!info.isFile() || info.size > MAX_BYTES) throw Error('Recent-file history exceeds its supported size.');
      return await readFile(this.#filename, 'utf8');
    } catch (error) {
      if (error.code === 'ENOENT') return null;
      throw error;
    }
  }
  async #load() {
    try {
      this.#raw = await this.#read();
      if (this.#raw === null) return;
      const value = JSON.parse(this.#raw);
      if (
        value?.schema !== 1 ||
        Object.keys(value).some((key) => !['schema', 'entries'].includes(key)) ||
        !Array.isArray(value.entries) ||
        value.entries.length > LIMIT
      )
        throw Error('Invalid recent-file history.');
      const ids = new Set(),
        files = new Set();
      for (const entry of value.entries) {
        if (
          !entry ||
          Object.keys(entry).some((key) => !['id', 'filename', 'lastOpened'].includes(key)) ||
          !validId(entry.id) ||
          !supported(entry.filename) ||
          !Number.isSafeInteger(entry.lastOpened) ||
          entry.lastOpened < 0 ||
          ids.has(entry.id) ||
          files.has(entry.filename)
        )
          throw Error('Invalid recent-file record.');
        ids.add(entry.id);
        files.add(entry.filename);
      }
      this.#entries = value.entries.map(({ id, filename, lastOpened }) => ({ id, filename, lastOpened }));
    } catch (error) {
      this.#locked = true;
      this.#warning =
        'Recent-file history could not be read. Its existing file is preserved; repair it and restart to use history. ' +
        String(error.message ?? error);
    }
  }
  #serial(work) {
    const next = this.#queue.then(work);
    this.#queue = next.catch(() => {});
    return next;
  }
  #entry(id) {
    const entry = validId(id) && this.#entries.find((entry) => entry.id === id);
    if (!entry) throw Error('Recent-file access is not granted. Refresh the recent files list.');
    return entry;
  }
  async #available(entry) {
    try {
      const canonical = await realpath(entry.filename);
      if (canonical !== entry.filename)
        return { available: false, reason: 'The path now points to another file. Open it using Open….' };
      if (!(await stat(canonical)).isFile()) return { available: false, reason: 'The path is not a file.' };
      return { available: true };
    } catch (error) {
      return { available: false, reason: failure(error) };
    }
  }
  async list() {
    await this.#queue;
    return {
      entries: await Promise.all(
        this.#entries.map(async (entry) => ({
          id: entry.id,
          name: path.basename(entry.filename),
          displayPath: entry.filename,
          lastOpened: entry.lastOpened,
          ...(await this.#available(entry)),
        }))
      ),
      ...(this.#warning ? { warning: this.#warning } : {}),
    };
  }
  async resolve(id) {
    await this.#queue;
    const entry = this.#entry(id),
      available = await this.#available(entry);
    if (!available.available)
      throw Error(
        available.reason +
          ' The current spreadsheet was kept open. Remove this recent entry or use Open… to locate the file.'
      );
    return entry.filename;
  }
  async #write(entries) {
    if (this.#locked) throw Error(this.#warning);
    try {
      if ((await this.#read()) !== this.#raw) {
        this.#locked = true;
        throw Error('The recent-file history changed on disk; restart before updating it.');
      }
      const value = { schema: 1, entries };
      await atomicJson(this.#filename, value);
      this.#raw = JSON.stringify(value);
      this.#entries = entries;
      this.#warning = undefined;
    } catch (error) {
      this.#warning =
        'Recent-file history could not be updated. Existing records were retained. ' + String(error.message ?? error);
      throw Error(this.#warning);
    }
  }
  /** Call only after the owner has opened and the shell has activated it. */
  remember(filename) {
    return this.#serial(async () => {
      const canonical = await realpath(filename);
      if (!supported(canonical) || !(await stat(canonical)).isFile()) throw Error('Unsupported recent-file path.');
      const previous = this.#entries.find((entry) => entry.filename === canonical);
      await this.#write(
        [
          { id: previous?.id ?? randomUUID(), filename: canonical, lastOpened: Date.now() },
          ...this.#entries.filter((entry) => entry.filename !== canonical),
        ].slice(0, LIMIT)
      );
    });
  }
  remove(id) {
    return this.#serial(async () => {
      this.#entry(id);
      await this.#write(this.#entries.filter((entry) => entry.id !== id));
    });
  }
}
