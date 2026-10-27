/** Dictionary parsing and suggestion CPU remain outside the host event loop. */
import { parentPort } from 'node:worker_threads';
import {
  loadSpellingAssets,
  validateSpellingRequest,
  OFFLINE_SPELLING_LIMITS,
  OFFLINE_SPELLING_PROVIDER_ID,
  OFFLINE_SPELLING_ENGINE,
} from './offline-spelling-assets.mjs';
import { refusal } from '../core/refusal-codes.mjs';

const { availability, engines } = await loadSpellingAssets();
parentPort.postMessage({ kind: 'ready', availability });
parentPort.on('message', (input) => {
  let request;
  try {
    request = validateSpellingRequest(input);
    const locale = availability.locales.find((item) => item.locale === request.locale),
      engine = engines.get(request.locale);
    const base = {
      schema: 1,
      requestId: request.requestId,
      locale: request.locale,
      provider: 'nspell',
      providerVersion: OFFLINE_SPELLING_ENGINE.version,
      providerId: OFFLINE_SPELLING_PROVIDER_ID,
    };
    if (!engine || !locale?.verified) {
      parentPort.postMessage({
        kind: 'result',
        reply: {
          ...base,
          status: 'unavailable',
          results: [],
          reason: locale?.reason ?? 'This offline spelling locale is not installed.',
        },
      });
      return;
    }
    const results = request.words.map((word) => {
      const correct = engine.correct(word),
        suggestions = correct
          ? []
          : [...new Set(engine.suggest(word))]
              .filter(
                (value) =>
                  typeof value === 'string' &&
                  value.length &&
                  Array.from(value).length <= OFFLINE_SPELLING_LIMITS.wordCodePoints &&
                  !/[\p{C}]/u.test(value)
              )
              .slice(0, OFFLINE_SPELLING_LIMITS.suggestions);
      return { word, correct, suggestions };
    });
    const reply = {
      ...base,
      status: 'completed',
      dictionaryId: locale.verified.dictionaryId,
      dictionary: locale.verified,
      results,
    };
    if (Buffer.byteLength(JSON.stringify(reply)) > OFFLINE_SPELLING_LIMITS.replyBytes)
      throw refusal('spelling-limit', 'Reply bound');
    parentPort.postMessage({ kind: 'result', reply });
  } catch {
    parentPort.postMessage({
      kind: 'failure',
      requestId: request?.requestId,
      reason: 'Offline spelling could not complete this bounded request.',
    });
  }
});
