import { workerRetirement } from './worker-retirement.mjs';
import { Worker } from 'node:worker_threads';
import { randomUUID } from 'node:crypto';
import { allocateWorkerStorage, releaseWorkerStorage, collectDeadWorkerStorage } from './worker-storage.mjs';
import { refusal } from '../core/refusal-codes.mjs';

/** Persistent literal CSV owner. Only metadata and bounded results cross this boundary. */
export class CsvFileWorker {
  id = randomUUID();
  #worker;
  #storage;
  #pending = new Map();
  #next = 0;
  #closed = false;
  #opened = false;
  #closing;
  #exited;
  #terminate;
  #stats = {};
  constructor() {
    this.#storage = allocateWorkerStorage();
    try {
      this.#worker = new Worker(new URL('./csv-owner-worker.mjs', import.meta.url), {
        execArgv: [],
        workerData: { diagnostics: !!process.env.XLSX_EDITOR_TRACE_PATH },
      });
    } catch (error) {
      void releaseWorkerStorage(this.#storage).catch(() => {});
      throw error;
    }
    this.#terminate = workerRetirement(this.#worker);
    this.#exited = new Promise((resolve, reject) => {
      this.#worker.once('exit', (code) => {
        this.#fail(refusal('session-closed', `CSV owner stopped (${code}); reopen to recover accepted edits.`));
        releaseWorkerStorage(this.#storage).then(() => {
          this.#stats = Object.fromEntries(Object.keys(this.#stats).map((key) => [key, 0]));
          resolve();
        }, reject);
      });
    });
    this.#exited.catch(() => {});
    this.#worker.on('error', (error) => this.#fail(error));
    this.#worker.on('message', (message) => {
      if (this.#closed) return;
      if (
        !message ||
        !Number.isSafeInteger(message.id) ||
        !this.#pending.has(message.id) ||
        (message.error !== undefined && typeof message.error !== 'string')
      ) {
        this.#fail(refusal('internal-invalid-field', 'Invalid CSV owner reply; reopen to recover accepted edits.'));
        void this.#terminate();
        return;
      }
      let { id, result, error, stats } = message;
      const request = this.#pending.get(id);
      this.#pending.delete(id);
      if (stats && typeof stats === 'object')
        this.#stats = Object.fromEntries(
          Object.entries(stats)
            .filter(([key, value]) => /^[a-zA-Z][a-zA-Z0-9_]{0,63}$/.test(key) && Number.isFinite(value))
            .slice(0, 32)
        );
      if (result?.storage) result = { ...result, storage: { ...result.storage, sessionId: this.id } };
      if (result?.view?.storage)
        result = { ...result, view: { ...result.view, storage: { ...result.view.storage, sessionId: this.id } } };
      error ? request.reject(Error(error)) : request.resolve(result);
    });
  }
  #fail(error) {
    this.#closed = true;
    for (const task of this.#pending.values()) task.reject(error);
    this.#pending.clear();
  }
  #call(method, args = []) {
    if (this.#closed || (this.#closing && method !== 'close'))
      return Promise.reject(refusal('session-closed', 'CSV owner is closed'));
    // The bounded shutdown message must remain admissible when all data slots
    // are occupied; it drains already accepted work and fences later requests.
    if (this.#pending.size >= 16 && method !== 'close')
      return Promise.reject(refusal('csv-limit', 'CSV owner request queue is full'));
    const id = ++this.#next;
    return new Promise((resolve, reject) => {
      try {
        if (Buffer.byteLength(JSON.stringify(args)) > 2 * 1024 * 1024)
          throw refusal('csv-limit', 'CSV request exceeds the 2 MiB transport budget');
        this.#pending.set(id, { resolve, reject });
        this.#worker.postMessage({ id, method, args });
      } catch (error) {
        this.#pending.delete(id);
        reject(error);
      }
    });
  }
  async open(source, options = {}) {
    if (this.#opened) throw refusal('csv-refused', 'CSV owner is already open');
    this.#opened = true;
    try {
      await collectDeadWorkerStorage();
      return await this.#call('open', [source, { ...options, storageDirectory: this.#storage.root }]);
    } catch (error) {
      await this.close().catch(() => {});
      throw error;
    }
  }
  read(...args) {
    return this.#call('read', args);
  }
  scan(...args) {
    return this.#call('scan', args);
  }
  query(...args) {
    return this.#call('query', args);
  }
  selectionSummary(...args) {
    return this.#call('selectionSummary', args);
  }
  apply(...args) {
    return this.#call('apply', args);
  }
  compact(...args) {
    return this.#call('compact', args);
  }
  export(...args) {
    return this.#call('export', args);
  }
  diagnosticStats() {
    return {
      ...this.#stats,
      workerId: this.#worker.threadId,
      pendingRequests: this.#pending.size,
      closed: Number(this.#closed),
    };
  }
  close() {
    if (this.#closing) return this.#closing;
    return (this.#closing = (async () => {
      let timer;
      try {
        if (!this.#closed)
          await Promise.race([
            this.#call('close'),
            new Promise((_, reject) => {
              timer = setTimeout(
                () => reject(refusal('operation-timed-out', 'CSV owner close deadline exceeded')),
                5000
              );
            }),
          ]);
      } finally {
        clearTimeout(timer);
        this.#fail(refusal('session-closed', 'CSV owner is closed'));
        await this.#terminate();
        await this.#exited;
      }
    })());
  }
}
