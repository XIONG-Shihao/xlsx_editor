import { workerRetirement } from './worker-retirement.mjs';
/** Main/owner-only passive file reader. Original snapshots never reach renderer paths. */
import fs from 'node:fs/promises';
import { constants } from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { randomUUID, createHash } from 'node:crypto';
import { Worker } from 'node:worker_threads';
import { PASSIVE_FORMAT_LIMITS as LIMITS } from '../core/passive-container.mjs';
import { refusal } from '../core/refusal-codes.mjs';
const fail = (message) => {
    throw refusal('package-refused', 'Passive workbook: ' + message);
  },
  stamp = (stat) => [stat.dev, stat.ino, stat.size, stat.mtimeNs, stat.ctimeNs].join(':');
const fileHash = async (filename) => {
  const handle = await fs.open(filename, constants.O_RDONLY | constants.O_NOFOLLOW),
    hash = createHash('sha256');
  try {
    const before = await handle.stat({ bigint: true });
    if (!before.isFile() || before.uid !== BigInt(process.getuid()) || before.size > BigInt(LIMITS.sourceBytes))
      fail('owned snapshot identity differs');
    for await (const bytes of handle.createReadStream({ autoClose: false })) hash.update(bytes);
    if (
      stamp(await handle.stat({ bigint: true })) !== stamp(before) ||
      stamp(await fs.lstat(filename, { bigint: true })) !== stamp(before)
    )
      fail('owned snapshot changed while checking');
    return hash.digest('hex');
  } finally {
    await handle.close();
  }
};
async function snapshot(source, directory, format) {
  const input = await fs.open(source, constants.O_RDONLY | constants.O_NOFOLLOW),
    target = path.join(directory, 'source.' + format);
  let output;
  try {
    const before = await input.stat({ bigint: true });
    if (
      !before.isFile() ||
      before.size > BigInt(LIMITS.sourceBytes) ||
      before.size === 0n ||
      before.uid !== BigInt(process.getuid())
    )
      fail('source file exceeds its regular owned-file budget');
    output = await fs.open(
      target,
      constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL | constants.O_NOFOLLOW,
      0o600
    );
    const buffer = Buffer.alloc(65536),
      hash = createHash('sha256');
    let count = 0;
    for (;;) {
      const { bytesRead } = await input.read(buffer, 0, buffer.length, null);
      if (!bytesRead) break;
      count += bytesRead;
      if (count > LIMITS.sourceBytes) fail('source grew beyond its byte budget');
      const bytes = buffer.subarray(0, bytesRead);
      hash.update(bytes);
      let at = 0;
      while (at < bytesRead) {
        const written = await output.write(bytes, at, bytesRead - at, null);
        if (!written.bytesWritten) fail('source snapshot write stopped');
        at += written.bytesWritten;
      }
    }
    if (
      count !== Number(before.size) ||
      stamp(await input.stat({ bigint: true })) !== stamp(before) ||
      stamp(await fs.lstat(source, { bigint: true })) !== stamp(before)
    )
      fail('source changed while snapshotting');
    await output.sync();
    await output.chmod(0o400);
    return { path: target, revision: hash.digest('hex'), bytes: count };
  } finally {
    await output?.close();
    await input.close();
  }
}
export class PassiveWorkbook {
  id = randomUUID();
  #directory;
  #snapshot;
  #worker;
  #terminate;
  #view;
  #sequence = 0;
  #pending = new Map();
  #closed = false;
  #started = false;
  #work = Promise.resolve();
  #exportNames = new Set();
  #exports = new Map();
  diagnosticStats() {
    return {
      retainedStates: this.#view && !this.#closed ? 1 : 0,
      storageRoots: this.#directory && !this.#closed ? 1 : 0,
      closed: Number(this.#closed),
      calculationCalls: 0,
      pendingCalls: this.#pending.size,
    };
  }
  has(state) {
    return !this.#closed && !!this.#view && state === '0';
  }
  retainedStates() {
    return this.has('0') ? ['0'] : [];
  }
  storageRoots() {
    return this.#directory && !this.#closed ? [this.#directory] : [];
  }
  #state(state) {
    if (!this.has(state)) fail('unknown, stale or closed passive state');
  }
  open(source, options = {}) {
    if (this.#started || this.#closed)
      return Promise.reject(refusal('session-closed', 'Passive workbook already opened or closed'));
    this.#started = true;
    const work = this.#open(source, options);
    this.#work = work.catch(() => {});
    return work;
  }
  async #open(source, options) {
    if (
      typeof source !== 'string' ||
      !path.isAbsolute(source) ||
      source.includes('\0') ||
      !options ||
      Object.keys(options).some((key) => !['format', 'storageDirectory'].includes(key)) ||
      !['xlsm', 'xls', 'xlsb'].includes(options.format)
    )
      fail('invalid trusted source or format');
    try {
      if (options.storageDirectory) {
        if (typeof options.storageDirectory !== 'string' || !path.isAbsolute(options.storageDirectory))
          fail('invalid private storage root');
        await fs.mkdir(options.storageDirectory, { recursive: true, mode: 0o700 });
        const stat = await fs.lstat(options.storageDirectory);
        if (!stat.isDirectory() || stat.isSymbolicLink() || stat.uid !== process.getuid() || stat.mode & 0o077)
          fail('storage root is not privately owned');
        this.#directory = await fs.mkdtemp(path.join(options.storageDirectory, 'passive-'));
      } else this.#directory = await fs.mkdtemp(path.join(os.tmpdir(), 'passive-workbook-'));
      await fs.chmod(this.#directory, 0o700);
      await fs.writeFile(
        path.join(this.#directory, 'owner.json'),
        JSON.stringify({ schema: 1, id: this.id, format: options.format }),
        { flag: 'wx', mode: 0o600 }
      );
      this.#snapshot = await snapshot(source, this.#directory, options.format);
      if (this.#closed) fail('owner closed while preparing source');
      this.#worker = new Worker(new URL('./passive-workbook-worker.mjs', import.meta.url), {
        workerData: {
          id: this.id,
          source: this.#snapshot.path,
          revision: this.#snapshot.revision,
          format: options.format,
        },
        resourceLimits: { maxOldGenerationSizeMb: LIMITS.heapMiB, maxYoungGenerationSizeMb: 32 },
        execArgv: [],
      });
      this.#terminate = workerRetirement(this.#worker);
      this.#worker.on('message', (message) => {
        const pending = this.#pending.get(message.id);
        if (!pending) return;
        this.#pending.delete(message.id);
        clearTimeout(pending.timer);
        message.error ? pending.reject(Error(message.error)) : pending.resolve(message.result);
      });
      const stopped = (error) => {
        for (const item of this.#pending.values()) {
          clearTimeout(item.timer);
          item.reject(
            refusal('session-closed', 'Passive parser stopped: ' + String(error?.message ?? error).slice(0, 1024))
          );
        }
        this.#pending.clear();
        this.#view = undefined;
      };
      this.#worker.on('error', stopped);
      this.#worker.on('exit', (code) => stopped('exit' + code));
      this.#view = await this.#rpc('info', {});
      return this.view('0');
    } catch (error) {
      await this.#stop();
      await this.#cleanup().catch(() => {});
      this.#closed = true;
      throw refusal(
        'package-refused',
        'Passive workbook opening refused: ' +
          String(error.message ?? error)
            .replaceAll(source, '[source]')
            .replaceAll(this.#directory ?? '\0', '[private]')
            .slice(0, 8192)
      );
    }
  }
  #rpc(operation, args) {
    if (!this.#worker || this.#closed) return Promise.reject(refusal('session-closed', 'Passive parser is closed'));
    if (this.#pending.size >= 16)
      return Promise.reject(refusal('package-limit', 'Passive request queue exceeds its budget'));
    const id = ++this.#sequence,
      message = { id, operation, args };
    if (Buffer.byteLength(JSON.stringify(message)) > 65536)
      return Promise.reject(refusal('package-limit', 'Passive request exceeds its byte budget'));
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.#pending.delete(id);
        reject(refusal('operation-timed-out', 'Passive parser operation timed out'));
        void this.#stop();
      }, LIMITS.openTimeoutMs);
      this.#pending.set(id, { resolve, reject, timer });
      this.#worker.postMessage(message);
    });
  }
  async #request(state, operation, args) {
    this.#state(state);
    const result = await this.#rpc(operation, args);
    this.#state(state);
    return result;
  }
  view(state) {
    this.#state(state);
    return structuredClone(this.#view);
  }
  read(state, sheet, range) {
    return this.#request(state, 'read', { sheet, range });
  }
  scan(state, sheet, after) {
    return this.#request(state, 'scan', { sheet, after });
  }
  query(state, sheet, query, mode, after, options = {}) {
    if (Object.keys(options).some((key) => !['range', 'matchLimit'].includes(key)))
      return Promise.reject(refusal('internal-invalid-field', 'Invalid passive search options'));
    return this.#request(state, 'query', { sheet, query, mode, after, ...options });
  }
  selectionSummary(state, sheet, range) {
    return this.#request(state, 'summary', { sheet, range });
  }
  calculationSummary(state) {
    return this.#request(state, 'calculation-summary', {});
  }
  objects(state, sheet, options = {}) {
    return this.#request(state, 'objects', { sheet, ...options });
  }
  vbaModules(state, options = {}) {
    return this.#request(state, 'vba-list', options);
  }
  vbaModule(state, moduleId, options = {}) {
    return this.#request(state, 'vba-code', { moduleId, ...options });
  }
  vbaFunctions(state, moduleId) {
    return this.#request(state, 'vba-functions', { moduleId });
  }
  evaluateVba(state, moduleId, functionName, args) {
    return this.#request(state, 'vba-evaluate', { moduleId, functionName, args });
  }
  dataOnlyExport(state, options) {
    this.#state(state);
    if (!options || Object.keys(options).length !== 1 || !['xls', 'xlsb', 'xlsx'].includes(options.format))
      return Promise.reject(refusal('package-malformed', 'Invalid data-only target format'));
    if (this.#exports.has(options.format))
      return this.#exports.get(options.format).then(async (result) => {
        if ((await fileHash(result.path)) !== result.revision) fail('fresh export changed before publication');
        this.#state(state);
        return { ...result };
      });
    if (this.#exportNames.size >= 8)
      return Promise.reject(refusal('package-limit', 'Fresh export resource budget exceeded'));
    const exportId = randomUUID(),
      filename = 'data-' + exportId + '.' + options.format;
    this.#exportNames.add(filename);
    const pending = this.#request(state, 'data-export', { format: options.format, exportId }).then(async (result) => {
      if (
        result.filename !== filename ||
        result.sourceRevision !== this.#snapshot.revision ||
        !/^[a-f0-9]{64}$/.test(result.revision) ||
        !Number.isSafeInteger(result.bytes) ||
        result.bytes < 1 ||
        result.bytes > LIMITS.sourceBytes
      )
        fail('fresh export identity disagrees');
      const output = path.join(this.#directory, filename);
      if ((await fileHash(output)) !== result.revision) fail('fresh export hash disagrees');
      this.#state(state);
      const { filename: privateName, ...metadata } = result;
      return { ...metadata, path: output };
    });
    this.#exports.set(options.format, pending);
    pending.catch(() => {
      this.#exports.delete(options.format);
    });
    return pending;
  }
  object(state) {
    this.#state(state);
    return Promise.reject(refusal('package-protected', 'Object graphs remain protected in this passive view'));
  }
  apply(state) {
    this.#state(state);
    return Promise.reject(refusal('package-protected', 'This original workbook format is read-only'));
  }
  conditionalRules(state) {
    this.#state(state);
    return Promise.reject(refusal('package-unsupported', 'Passive conditional-rule inspection is not qualified'));
  }
  validationRules(state) {
    this.#state(state);
    return Promise.reject(refusal('package-unsupported', 'Passive validation inspection is not qualified'));
  }
  printDefinition(state) {
    this.#state(state);
    return Promise.reject(refusal('package-unsupported', 'Passive-format print/PDF is not qualified'));
  }
  printSnapshot(state) {
    return this.printDefinition(state);
  }
  filterValues(state) {
    this.#state(state);
    return Promise.reject(refusal('package-unsupported', 'Passive-format filter values are not qualified'));
  }
  keys(state) {
    this.#state(state);
    return Promise.reject(refusal('package-protected', 'Passive-format transformations are read-only'));
  }
  async compact(keep = ['0']) {
    if (!Array.isArray(keep) || keep.length !== 1 || keep[0] !== '0')
      fail('compact must retain the only immutable state');
    this.#state('0');
    return { retainedStates: 1, storageRoots: 1, unchanged: true };
  }
  async export(state) {
    this.#state(state);
    const stat = await fs.lstat(this.#snapshot.path);
    if (
      !stat.isFile() ||
      stat.isSymbolicLink() ||
      stat.uid !== process.getuid() ||
      stat.size !== this.#snapshot.bytes ||
      (await fileHash(this.#snapshot.path)) !== this.#snapshot.revision
    )
      fail('original snapshot changed before Copy');
    this.#state(state);
    return { path: this.#snapshot.path, revision: this.#snapshot.revision };
  }
  async #stop() {
    const worker = this.#worker;
    this.#worker = undefined;
    this.#view = undefined;
    if (worker) await this.#terminate();
    for (const item of this.#pending.values()) {
      clearTimeout(item.timer);
      item.reject(refusal('package-refused', 'Passive parser closed'));
    }
    this.#pending.clear();
  }
  async #cleanup() {
    if (!this.#directory) return;
    const stat = await fs.lstat(this.#directory),
      names = await fs.readdir(this.#directory);
    if (!stat.isDirectory() || stat.isSymbolicLink() || stat.uid !== process.getuid() || stat.mode & 0o077)
      fail('private cleanup root changed');
    if (!names.length) {
      await fs.rmdir(this.#directory);
      this.#directory = undefined;
      return;
    }
    const marker = path.join(this.#directory, 'owner.json'),
      markerStat = await fs.lstat(marker);
    if (
      !markerStat.isFile() ||
      markerStat.isSymbolicLink() ||
      markerStat.uid !== process.getuid() ||
      markerStat.size > 4096
    )
      fail('private cleanup marker changed');
    const handle = await fs.open(marker, constants.O_RDONLY | constants.O_NOFOLLOW);
    let owner;
    try {
      owner = JSON.parse(await handle.readFile('utf8'));
    } finally {
      await handle.close();
    }
    if (
      owner.schema !== 1 ||
      owner.id !== this.id ||
      !['xlsm', 'xls', 'xlsb'].includes(owner.format) ||
      Object.keys(owner).length !== 3 ||
      names.some((name) => !['owner.json', 'source.' + owner.format].includes(name) && !this.#exportNames.has(name))
    )
      fail('private cleanup retains unexpected resources');
    for (const name of names) {
      const info = await fs.lstat(path.join(this.#directory, name));
      if (!info.isFile() || info.isSymbolicLink() || info.uid !== process.getuid())
        fail('private cleanup retains an unexpected file type');
    }
    await fs.rm(this.#directory, { recursive: true, force: false });
    this.#directory = undefined;
  }
  async close() {
    if (this.#closed && !this.#directory) return;
    this.#closed = true;
    await this.#work;
    await this.#stop();
    this.#view = undefined;
    await this.#cleanup();
  }
}
