/** Main-process dictionary work bound to one accepted owner and read grant.
 * This module never receives workbook paths or mutates workbook history. */
import {
  createOfflineSpellingProvider,
  OFFLINE_SPELLING_LIMITS,
  OFFLINE_SPELLING_PROVIDER_ID,
} from './offline-spelling.mjs';
import { refusal } from '../core/refusal-codes.mjs';

const fields = ['documentId', 'ownerId', 'state', 'revision'];
const uuid = /^[a-f0-9]{8}-[a-f0-9]{4}-[1-8][a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}$/;
const assert = (value, message) => {
  if (!value) throw refusal('package-malformed', message);
};
const exact = (value, keys) =>
  value &&
  typeof value === 'object' &&
  !Array.isArray(value) &&
  Object.keys(value).sort().join(',') === [...keys].sort().join(',');
const freeze = (value) => {
  if (value && typeof value === 'object') {
    for (const item of Object.values(value)) freeze(item);
    Object.freeze(value);
  }
  return value;
};
const receipt = (value) => {
  assert(
    exact(value, fields) &&
      fields.every((key) => typeof value[key] === 'string' && value[key].length > 0 && value[key].length <= 4096),
    'Invalid spelling owner receipt.'
  );
  return freeze(Object.fromEntries(fields.map((key) => [key, value[key]])));
};
const same = (one, two) => fields.every((key) => one[key] === two[key]);
export function spellingRequest(input) {
  const operation = input?.operation;
  assert(
    ['spellingAvailability', 'spellingCheck', 'spellingCancel'].includes(operation),
    'Invalid spelling operation.'
  );
  const transport = Object.hasOwn(input, 'kind') || Object.hasOwn(input, 'sessionId');
  const keys = [
    'operation',
    ...fields,
    'requestId',
    ...(transport ? ['kind', 'sessionId'] : []),
    ...(operation === 'spellingCheck' ? ['locale', 'words'] : []),
  ];
  assert(
    exact(input, keys) &&
      (!transport ||
        (input.kind === 'indexed' && typeof input.sessionId === 'string' && input.sessionId === input.ownerId)) &&
      typeof input.requestId === 'string' &&
      uuid.test(input.requestId),
    'Invalid closed spelling request.'
  );
  receipt(Object.fromEntries(fields.map((key) => [key, input[key]])));
  if (operation === 'spellingCheck') {
    assert(
      typeof input.locale === 'string' &&
        input.locale.length > 0 &&
        input.locale.length <= 64 &&
        Array.isArray(input.words) &&
        input.words.length > 0 &&
        input.words.length <= OFFLINE_SPELLING_LIMITS.words,
      'Invalid spelling locale or word batch.'
    );
    for (const word of input.words)
      assert(
        typeof word === 'string' &&
          word.length > 0 &&
          Array.from(word).length <= OFFLINE_SPELLING_LIMITS.wordCodePoints &&
          !/[\s\p{C}]/u.test(word),
        'Spelling requires bounded individual words.'
      );
    assert(
      Buffer.byteLength(JSON.stringify({ requestId: input.requestId, locale: input.locale, words: input.words })) <=
        OFFLINE_SPELLING_LIMITS.requestBytes,
      'Spelling word batch exceeds 32 KiB.'
    );
  }
  assert(Buffer.byteLength(JSON.stringify(input)) <= 65536, 'Spelling request exceeds its transport bound.');
  return freeze({ ...input, ...(operation === 'spellingCheck' ? { words: [...input.words] } : {}) });
}
const cancelled = (job) => ({
  spellingCheck: {
    schema: 1,
    receipt: job.receipt,
    result: {
      schema: 1,
      requestId: job.request.requestId,
      locale: job.request.locale,
      provider: 'nspell',
      providerVersion: '2.1.5',
      providerId: OFFLINE_SPELLING_PROVIDER_ID,
      status: 'cancelled',
      results: [],
      reason: 'Spelling request cancelled.',
    },
  },
});

export class WorkbookSpellingJobs {
  #defaults;
  #factory;
  #provider;
  #ids = new Set();
  #jobs = new Map();
  #stops = new Set();
  #active;
  #tail = Promise.resolve();
  #retired = false;
  #closing;
  constructor({ capture, verify, authorize, providerFactory = createOfflineSpellingProvider } = {}) {
    this.#defaults = { capture, verify, authorize };
    this.#factory = providerFactory;
  }
  #access(access) {
    const value = access ?? this.#defaults;
    assert(
      ['capture', 'verify', 'authorize'].every((key) => typeof value[key] === 'function'),
      'Spelling requires captured owner verification and read authorization.'
    );
    return value;
  }
  #check() {
    assert(!this.#retired, 'Spelling owner retired.');
  }
  #stop(work) {
    const pending = Promise.resolve()
      .then(work)
      .finally(() => this.#stops.delete(pending));
    this.#stops.add(pending);
    return pending;
  }
  async #verify(job) {
    this.#check();
    await job.access.authorize(job.request, job.receipt);
    this.#check();
    await job.access.verify(job.request, job.receipt);
    this.#check();
  }
  run(incoming, access) {
    const request = spellingRequest(incoming);
    assert(request.operation !== 'spellingCancel', 'Use scoped spelling cancellation.');
    this.#check();
    const guard = this.#access(access),
      captured = receipt(guard.capture(request));
    assert(same(captured, request), 'Spelling request belongs to another owner or accepted state.');
    assert(!this.#ids.has(request.requestId), 'Spelling request ID cannot be reused.');
    assert(this.#ids.size < OFFLINE_SPELLING_LIMITS.requestIds, 'Spelling request lifetime limit reached.');
    assert(this.#jobs.size < OFFLINE_SPELLING_LIMITS.queuedRequests, 'Too many queued spelling requests.');
    this.#ids.add(request.requestId);
    const job = { request, receipt: captured, access: guard, cancelled: false, done: undefined };
    this.#jobs.set(request.requestId, job);
    const work = this.#tail.then(async () => {
      await Promise.all([...this.#stops]);
      await this.#verify(job);
      if (job.cancelled) {
        if (request.operation === 'spellingCheck') return cancelled(job);
        throw refusal('operation-cancelled', 'Spelling availability cancelled.');
      }
      this.#active = job;
      const provider = (this.#provider ??= this.#factory());
      try {
        const result =
          request.operation === 'spellingAvailability'
            ? await provider.availability()
            : await provider.check({ requestId: request.requestId, locale: request.locale, words: request.words });
        await this.#verify(job);
        if (job.cancelled) {
          if (request.operation === 'spellingCheck') return cancelled(job);
          throw refusal('operation-cancelled', 'Spelling availability cancelled.');
        }
        assert(
          result?.schema === 1 &&
            result.provider === 'nspell' &&
            result.providerVersion === '2.1.5' &&
            result.providerId === OFFLINE_SPELLING_PROVIDER_ID &&
            Buffer.byteLength(JSON.stringify(result)) <= OFFLINE_SPELLING_LIMITS.replyBytes,
          'Invalid bounded offline spelling reply.'
        );
        if (request.operation === 'spellingAvailability')
          return {
            spellingAvailability: { schema: 1, requestId: request.requestId, receipt: captured, availability: result },
          };
        assert(
          result.requestId === request.requestId && result.locale === request.locale && Array.isArray(result.results),
          'Spelling reply belongs to another request.'
        );
        if (result.status === 'completed')
          assert(
            result.results.length === request.words.length &&
              result.results.every((item, index) => item.word === request.words[index]),
            'Spelling reply changed its word batch.'
          );
        return { spellingCheck: { schema: 1, receipt: captured, result } };
      } finally {
        if (this.#active === job) this.#active = undefined;
      }
    });
    job.done = work.finally(() => this.#jobs.delete(request.requestId));
    this.#tail = job.done.catch(() => {});
    return job.done;
  }
  async cancel(incoming, access) {
    const request = spellingRequest(incoming);
    assert(request.operation === 'spellingCancel', 'Invalid spelling cancellation.');
    this.#check();
    const guard = this.#access(access),
      captured = receipt(guard.capture(request));
    assert(same(captured, request), 'Spelling cancellation belongs to another owner or accepted state.');
    await guard.authorize(request, captured);
    this.#check();
    await guard.verify(request, captured);
    this.#check();
    const job = this.#jobs.get(request.requestId);
    if (!job) return { spellingCancellation: { schema: 1, requestId: request.requestId, cancelled: false } };
    assert(same(job.receipt, captured), 'Spelling cancellation belongs to another job owner.');
    job.cancelled = true;
    if (this.#active === job && this.#provider) {
      const provider = this.#provider;
      if (job.request.operation === 'spellingCheck') await this.#stop(() => provider.cancel(request.requestId));
      else {
        this.#provider = undefined;
        await this.#stop(() => provider.dispose());
      }
    }
    this.#check();
    await guard.authorize(request, captured);
    this.#check();
    await guard.verify(request, captured);
    this.#check();
    return { spellingCancellation: { schema: 1, requestId: request.requestId, cancelled: true } };
  }
  retire() {
    this.#retired = true;
    for (const job of this.#jobs.values()) job.cancelled = true;
  }
  close() {
    if (this.#closing) return this.#closing;
    this.retire();
    const provider = this.#provider;
    this.#provider = undefined;
    this.#closing = (async () => {
      const results = await Promise.allSettled([
        Promise.resolve().then(() => provider?.dispose()),
        ...this.#stops,
        this.#tail,
      ]);
      const failure = results.find((result) => result.status === 'rejected');
      if (failure) throw failure.reason;
    })();
    return this.#closing;
  }
}
