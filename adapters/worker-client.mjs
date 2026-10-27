import { workerRetirement } from './worker-retirement.mjs';
import { Worker, MessageChannel } from 'node:worker_threads';
import { createNativeRelay } from './native-relay.mjs';
import { allocateWorkerStorage, releaseWorkerStorage, collectDeadWorkerStorage } from './worker-storage.mjs';
import { nativeWorkbookBinary } from '../core/native-runtime.mjs';
import { refusal } from '../core/refusal-codes.mjs';

// This compiled file worker needs no application preloads, eval-only flags or
// test-runner/V8 CLI arguments. Explicitly keep its entry environment independent.
/** Bounded cell requests, no workbook graph in the host application's main thread. */
export class SpreadsheetWorker {
  #worker;
  #next = 0;
  #pending = new Map();
  #closed = false;
  #ready;
  #storage;
  #exited;
  #terminate;
  #closing;
  #native;
  constructor({
    filename,
    newWorkbook,
    recoveryId,
    resume = false,
    fresh = false,
    prepared = false,
    recoveryDirectory,
    maxRecoveryBytes = 512 * 1024 * 1024,
    maxIndexCacheBytes = 2 * 1024 ** 3,
    initialCalculationPolicy,
    onProgress,
    onStart,
  }) {
    if (!Number.isSafeInteger(maxRecoveryBytes) || maxRecoveryBytes < 1024 * 1024)
      throw refusal('internal-invalid-field', 'Invalid recovery budget.');
    if (!Number.isSafeInteger(maxIndexCacheBytes) || maxIndexCacheBytes < 4096 || maxIndexCacheBytes > 2 * 1024 ** 3)
      throw refusal('session-malformed', 'Invalid workbook index cache budget.');
    if (
      [typeof filename === 'string', newWorkbook !== undefined, recoveryId !== undefined].filter(Boolean).length !== 1
    )
      throw refusal('session-bad-input', 'Choose one workbook owner source.');
    const binary =
      newWorkbook !== undefined || recoveryId !== undefined || /\.xlsx$/i.test(filename)
        ? nativeWorkbookBinary()
        : undefined;
    this.#storage = allocateWorkerStorage();
    const { port1, port2 } = new MessageChannel();
    this.#native = createNativeRelay(port1, binary, { storageRoot: this.#storage.root });
    try {
      this.#worker = new Worker(new URL('./owner-worker.mjs', import.meta.url), {
        execArgv: [],
        workerData: {
          filename,
          newWorkbook,
          recoveryId,
          resume,
          fresh,
          prepared,
          recoveryDirectory,
          maxRecoveryBytes,
          maxIndexCacheBytes,
          initialCalculationPolicy,
          indexDirectory: this.#storage.root,
          nativeBinary: binary,
          nativePort: port2,
        },
        transferList: [port2],
      });
    } catch (error) {
      port2.close();
      void this.#native
        .close()
        .then(() => releaseWorkerStorage(this.#storage))
        .catch(() => {});
      throw error;
    }
    this.#terminate = workerRetirement(this.#worker);
    this.#exited = new Promise((resolve, reject) => {
      this.#worker.once('exit', (code) => {
        this.#fail(
          refusal('session-closed', `Spreadsheet owner stopped (${code}); reopen to recover its accepted edits.`)
        );
        this.#native
          .close()
          .then((proof) =>
            releaseWorkerStorage(this.#storage, {
              ...(Number.isSafeInteger(proof?.reapedNativePid) && proof.reapedNativePid > 0
                ? { reapedNativeProof: proof }
                : {}),
            })
          )
          .then(resolve, reject);
      });
    });
    this.#exited.catch(() => {});
    this.#worker.on('message', (message) => {
      if (message.event === 'progress') {
        const pending = this.#pending.get(message.requestId);
        if (
          pending &&
          typeof message.phase === 'string' &&
          message.phase.length <= 64 &&
          typeof message.cancellable === 'boolean'
        )
          try {
            pending.onProgress?.(message);
          } catch {}
        return;
      }
      const { id, result, error, errorCode } = message,
        pending = this.#pending.get(id);
      if (!pending) return;
      this.#pending.delete(id);
      if (error) {
        const failure = Error(error);
        if (typeof errorCode === 'string' && errorCode.length <= 64) failure.code = errorCode;
        pending.reject(failure);
      } else pending.resolve(result);
    });
    this.#worker.on('error', (error) => this.#fail(error));
    this.#ready = (async () => {
      await collectDeadWorkerStorage();
      return this.#call('open', undefined, { onProgress, onStart });
    })();
    this.#ready.catch(() => {});
  }
  #fail(error) {
    this.#closed = true;
    for (const pending of this.#pending.values()) pending.reject(error);
    this.#pending.clear();
  }
  #call(operation, request, { onProgress, onStart } = {}) {
    if (this.#closed) return Promise.reject(refusal('session-closed', 'Spreadsheet owner is closed.'));
    if (this.#pending.size >= 16) return Promise.reject(refusal('session-limit', 'Spreadsheet request queue is full.'));
    const id = ++this.#next;
    return new Promise((resolve, reject) => {
      try {
        if (request && Buffer.byteLength(JSON.stringify(request)) > 2 * 1024 * 1024)
          throw refusal('session-limit', 'Spreadsheet request exceeds the 2 MiB transport budget.');
        this.#pending.set(id, { resolve, reject, onProgress });
        onStart?.(id);
        this.#worker.postMessage({ id, operation, request });
      } catch (error) {
        this.#pending.delete(id);
        reject(error);
      }
    });
  }
  open() {
    return this.#ready;
  }
  async request(request, options) {
    await this.#ready;
    return this.#call('request', request, options);
  }
  authorizeCancel(requestId) {
    if (!Number.isSafeInteger(requestId) || requestId < 1)
      return Promise.reject(refusal('internal-invalid-field', 'Invalid cancellation request.'));
    return this.#call('cancel-request', { requestId });
  }
  get closed() {
    return this.#closed;
  }
  /** Stop this document owner. Accepted journal entries survive; reopen to continue. */
  cancel() {
    if (this.#closing) return this.#closing;
    this.#fail(refusal('operation-cancelled', 'Spreadsheet operation cancelled; reopen to recover accepted edits.'));
    return (this.#closing = (async () => {
      await this.#terminate();
      await this.#exited;
    })());
  }
  close() {
    if (this.#closing) return this.#closing;
    return (this.#closing = (async () => {
      let timer;
      try {
        await Promise.race([
          (async () => {
            await this.#ready.catch(() => {});
            if (!this.#closed) await this.#call('close');
          })(),
          new Promise((_, reject) => {
            timer = setTimeout(
              () =>
                reject(refusal('operation-timed-out', 'Spreadsheet owner close deadline exceeded; recovery retained.')),
              5000
            );
          }),
        ]);
      } finally {
        clearTimeout(timer);
        this.#fail(refusal('session-closed', 'Spreadsheet owner is closed.'));
        await this.#terminate();
        await this.#exited;
      }
    })());
  }
}
