import { workerRetirement } from './worker-retirement.mjs';
/** Main-only offline spelling. A cancelled/expired worker never adopts late data. */
import { Worker } from 'node:worker_threads';
import {
  validateSpellingRequest,
  unavailableSpellingAvailability,
  OFFLINE_SPELLING_LIMITS,
  OFFLINE_SPELLING_ENGINE,
  OFFLINE_SPELLING_PROVIDER_ID,
} from './offline-spelling-assets.mjs';
import { refusal } from '../core/refusal-codes.mjs';
export {
  OFFLINE_SPELLING_LIMITS,
  OFFLINE_SPELLING_ENGINE,
  OFFLINE_SPELLING_DICTIONARIES,
  OFFLINE_SPELLING_PROVIDER_ID,
} from './offline-spelling-assets.mjs';

export function createOfflineSpellingProvider() {
  let closed = false,
    worker = null,
    terminateWorker = null,
    epoch = 0,
    boot = null,
    available = null,
    startTimer = null,
    active = null,
    pumping = false,
    retiring = null;
  const queued = [],
    used = new Set();
  const result = (request, status, reason) => ({
    schema: 1,
    requestId: request.requestId,
    locale: request.locale,
    provider: 'nspell',
    providerVersion: OFFLINE_SPELLING_ENGINE.version,
    providerId: OFFLINE_SPELLING_PROVIDER_ID,
    status,
    results: [],
    reason,
  });
  function settle(item, reply) {
    if (!item || item.done) return;
    item.done = true;
    clearTimeout(item.timer);
    item.resolve(reply);
  }
  function retire(reason) {
    if (retiring) return retiring;
    const previous = terminateWorker;
    worker = null;
    terminateWorker = null;
    epoch++;
    available = null;
    clearTimeout(startTimer);
    startTimer = null;
    if (boot) {
      boot.resolve(unavailableSpellingAvailability(reason));
      boot = null;
    }
    retiring = (previous ? previous() : Promise.resolve())
      .then(() => {})
      .finally(() => {
        retiring = null;
      });
    return retiring;
  }
  async function failed(reason) {
    const item = active;
    active = null;
    settle(item, result(item?.request ?? { requestId: '', locale: '' }, 'unavailable', reason));
    await retire(reason);
    void pump();
  }
  function start() {
    if (closed) return Promise.resolve(unavailableSpellingAvailability('The offline spelling provider is closed.'));
    if (retiring) return retiring.then(start);
    if (available) return Promise.resolve(structuredClone(available));
    if (boot) return boot.promise;
    let resolve;
    const promise = new Promise((done) => {
      resolve = done;
    });
    boot = { promise, resolve };
    const generation = ++epoch;
    try {
      worker = new Worker(new URL('./offline-spelling-worker.mjs', import.meta.url), {
        execArgv: [],
        resourceLimits: {
          maxOldGenerationSizeMb: OFFLINE_SPELLING_LIMITS.workerOldMb,
          maxYoungGenerationSizeMb: OFFLINE_SPELLING_LIMITS.workerYoungMb,
          stackSizeMb: 4,
        },
      });
    } catch {
      const pending = boot;
      boot = null;
      pending.resolve(unavailableSpellingAvailability('The offline spelling worker could not start.'));
      return promise;
    }
    terminateWorker = workerRetirement(worker);
    worker.on('message', (message) => {
      if (generation !== epoch || closed) return;
      if (message?.kind === 'ready' && boot) {
        clearTimeout(startTimer);
        startTimer = null;
        available = message.availability;
        const pending = boot;
        boot = null;
        pending.resolve(structuredClone(available));
        return;
      }
      if (message?.kind === 'result' && active && message.reply?.requestId === active.request.requestId) {
        const item = active;
        active = null;
        settle(item, message.reply);
        void pump();
      } else if (message?.kind === 'failure' && active && message.requestId === active.request.requestId)
        void failed('Offline spelling could not complete this bounded request.');
    });
    worker.on('error', () => {
      if (generation === epoch && !closed)
        void failed('The offline spelling worker stopped before completing its request.');
    });
    worker.on('exit', () => {
      if (generation === epoch && !closed)
        void failed('The offline spelling worker stopped before completing its request.');
    });
    startTimer = setTimeout(() => {
      if (generation === epoch) void failed('Offline dictionaries could not be loaded within the worker time limit.');
    }, OFFLINE_SPELLING_LIMITS.startupMs);
    return promise;
  }
  async function pump() {
    if (pumping || closed || active || !queued.length) return;
    pumping = true;
    try {
      await start();
      if (closed || active) return;
      const item = queued.shift();
      if (!item) return;
      if (item.done) return;
      const locale = available?.locales.find((value) => value.locale === item.request.locale);
      if (!worker || !locale || locale.status !== 'available') {
        settle(
          item,
          result(item.request, 'unavailable', locale?.reason ?? 'This offline spelling locale is not installed.')
        );
        return;
      }
      active = item;
      item.timer = setTimeout(() => {
        if (active === item) void failed('Offline spelling exceeded its worker time limit.');
      }, OFFLINE_SPELLING_LIMITS.workMs);
      worker.postMessage(item.request);
    } catch {
      await failed('The offline spelling worker could not complete the request.');
    } finally {
      pumping = false;
      if (!active && queued.length && !closed) void pump();
    }
  }
  const api = {
    availability: async () => structuredClone(await start()),
    check(input) {
      if (closed) throw refusal('session-closed', 'The offline spelling provider is closed.');
      const request = validateSpellingRequest(input);
      if (used.has(request.requestId))
        throw refusal('spelling-refused', 'Offline spelling request ID has already been used.');
      if (used.size >= OFFLINE_SPELLING_LIMITS.requestIds)
        throw refusal('spelling-limit', 'Offline spelling request lifetime limit reached. Reopen the provider.');
      if (queued.length + (active ? 1 : 0) >= OFFLINE_SPELLING_LIMITS.queuedRequests)
        throw refusal('spelling-limit', 'Offline spelling request queue is full.');
      used.add(request.requestId);
      return new Promise((resolve) => {
        queued.push({ request, resolve, done: false });
        void pump();
      });
    },
    async cancel(requestId) {
      if (
        typeof requestId !== 'string' ||
        !/^[a-f0-9]{8}-[a-f0-9]{4}-[1-8][a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}$/.test(requestId)
      )
        throw refusal('internal-invalid-field', 'Invalid offline spelling request ID.');
      const index = queued.findIndex((item) => item.request.requestId === requestId);
      if (index !== -1) {
        const [item] = queued.splice(index, 1);
        settle(item, result(item.request, 'cancelled', 'Offline spelling request cancelled.'));
        return;
      }
      if (active?.request.requestId === requestId) {
        const item = active;
        active = null;
        settle(item, result(item.request, 'cancelled', 'Offline spelling request cancelled.'));
        await retire('Offline spelling request cancelled.');
        void pump();
      }
    },
    async dispose() {
      if (closed) return;
      closed = true;
      const pending = [...queued.splice(0), ...(active ? [active] : [])];
      active = null;
      for (const item of pending)
        settle(item, result(item.request, 'cancelled', 'The offline spelling provider is closed.'));
      await retire('The offline spelling provider is closed.');
    },
  };
  return Object.freeze(api);
}
