import { createRequire } from 'node:module';
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { randomUUID, createHash } from 'node:crypto';
import {
  createOfflineSpellingProvider,
  OFFLINE_SPELLING_DICTIONARIES,
  OFFLINE_SPELLING_ENGINE,
  OFFLINE_SPELLING_PROVIDER_ID,
  OFFLINE_SPELLING_LIMITS,
} from '../adapters/offline-spelling.mjs';

const component = fileURLToPath(new URL('../', import.meta.url));
const require = createRequire(import.meta.url),
  installed = (name) =>
    name.startsWith('dictionary-')
      ? path.dirname(require.resolve(name))
      : path.dirname(require.resolve(name + '/package.json'));
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const request = (locale, words) => ({ requestId: randomUUID(), locale, words });
const adapterFiles = [
  'offline-spelling.mjs',
  'offline-spelling-assets.mjs',
  'offline-spelling-worker.mjs',
  'offline-spelling.d.mts',
];
async function withProvider(fn, create = createOfflineSpellingProvider) {
  const provider = create();
  try {
    return await fn(provider);
  } finally {
    await provider.dispose();
  }
}
async function copiedProvider(change) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'offline-spelling-assets-'));
  await fs.mkdir(path.join(root, 'adapters'));
  await fs.mkdir(path.join(root, 'node_modules'));
  try {
    for (const name of adapterFiles)
      await fs.copyFile(path.join(component, 'adapters', name), path.join(root, 'adapters', name));
    for (const name of ['nspell', 'is-buffer', 'dictionary-en', 'dictionary-en-gb'])
      await fs.cp(installed(name), path.join(root, 'node_modules', name), { recursive: true });
    await change(root);
    const module = await import(pathToFileURL(path.join(root, 'adapters/offline-spelling.mjs')).href);
    return {
      root,
      create: module.createOfflineSpellingProvider,
      dispose: () => fs.rm(root, { recursive: true, force: true }),
    };
  } catch (error) {
    await fs.rm(root, { recursive: true, force: true });
    throw error;
  }
}
async function check(name, fn) {
  await test(name, async () => {
    await fn();
  });
}

await check(
  'Actual pinned en-US/en-GB dictionaries: identity, dialect and ranked suggestions are deterministic',
  async () =>
    withProvider(async (provider) => {
      const availability = await provider.availability();
      assert.equal(availability.providerId, OFFLINE_SPELLING_PROVIDER_ID);
      assert.deepEqual(availability.expectedProvider, OFFLINE_SPELLING_ENGINE);
      assert.deepEqual(availability.verifiedProvider, availability.expectedProvider);
      for (const [index, locale] of availability.locales.entries()) {
        assert.equal(locale.status, 'available');
        assert.deepEqual(locale.expected, OFFLINE_SPELLING_DICTIONARIES[index]);
        assert.deepEqual(locale.verified, locale.expected);
        for (const [name, sha] of [
          ['index.aff', locale.verified.affSha256],
          ['index.dic', locale.verified.dicSha256],
          ['license', locale.verified.licenseSha256],
        ])
          assert.equal(hash(await fs.readFile(path.join(installed(locale.verified.package), name))), sha);
      }
      const words = ['hello', 'world', 'helo', 'wrld', 'color', 'colour', 'café', 'naïve'],
        us = await provider.check(request('en-US', words)),
        gb = await provider.check(request('en-GB', words)),
        repeat = await provider.check(request('en-US', words));
      assert.equal(us.status, 'completed');
      assert.equal(gb.status, 'completed');
      assert.equal(us.dictionaryId, availability.locales[0].verified.dictionaryId);
      assert.equal(gb.dictionaryId, availability.locales[1].verified.dictionaryId);
      assert.notEqual(us.dictionaryId, gb.dictionaryId);
      assert.deepEqual(us.results, repeat.results);
      assert.deepEqual(us.results.slice(0, 2), [
        { word: 'hello', correct: true, suggestions: [] },
        { word: 'world', correct: true, suggestions: [] },
      ]);
      assert.equal(us.results[2].correct, false);
      assert(us.results[2].suggestions.includes('hello'));
      assert(us.results[3].suggestions.includes('world'));
      assert.equal(us.results[4].correct, true);
      assert.equal(us.results[5].correct, false);
      assert.deepEqual(us.results[5].suggestions, ['color']);
      assert.equal(gb.results[4].correct, false);
      assert.equal(gb.results[5].correct, true);
      assert(gb.results[4].suggestions.includes('colour'));
      assert.deepEqual(
        us.results.slice(6).map(({ word, correct }) => ({ word, correct })),
        [
          { word: 'café', correct: false },
          { word: 'naïve', correct: false },
        ]
      );
      return { profile: 'actual installed dictionary assets; no workbook/native/UI claim', availability, us, gb };
    })
);

await check(
  'Closed request and whole JSON-byte bounds refuse without truncation; unsupported locale is unavailable',
  async () =>
    withProvider(async (provider) => {
      const unsupported = await provider.check(request('fr-FR', ['hello']));
      assert.equal(unsupported.status, 'unavailable');
      assert.deepEqual(unsupported.results, []);
      assert(!Object.hasOwn(unsupported, 'dictionaryId'));
      for (const bad of [
        { ...request('en-US', ['hello']), filename: '/private/a.dic' },
        request('en-US', []),
        request('en-US', Array(101).fill('hello')),
        request('en-US', ['a'.repeat(257)]),
        request('en-US', [' hello']),
        request('en-US', ['a\ud800']),
        { ...request('en-US', ['hello']), requestId: 'not-a-token' },
        request('en-US', Array(100).fill('😀'.repeat(256))),
      ])
        assert.throws(() => provider.check(bad), /Invalid|bounded|32 KiB/);
      const id = request('en-US', ['hello']);
      await provider.check(id);
      assert.throws(() => provider.check(id), /already been used/);
      const admitted = await provider.check(request('en-US', Array(100).fill('hello')));
      assert.equal(admitted.status, 'completed');
      assert.equal(admitted.results.length, 100);
      assert(admitted.results.every((value) => value.correct));
      return { unsupported, limits: OFFLINE_SPELLING_LIMITS };
    })
);

await check(
  'A missing copied US dictionary is explicitly unavailable while the intact copied GB dictionary remains usable',
  async () => {
    const copy = await copiedProvider((root) => fs.rm(path.join(root, 'node_modules/dictionary-en/index.dic')));
    try {
      return await withProvider(async (provider) => {
        const availability = await provider.availability(),
          us = availability.locales[0];
        assert.equal(us.status, 'unavailable');
        assert(!us.verified);
        assert(us.expected.dicSha256);
        assert.equal(availability.locales[1].status, 'available');
        const result = await provider.check(request('en-US', ['hello']));
        assert.equal(result.status, 'unavailable');
        assert.deepEqual(result.results, []);
        assert(!JSON.stringify(result).includes(copy.root));
        assert.equal((await provider.check(request('en-GB', ['colour']))).results[0].correct, true);
        return { profile: 'missing asset fault in isolated real-package copy', availability };
      }, copy.create);
    } finally {
      await copy.dispose();
    }
  }
);

await check(
  'Copied affix/dictionary/license tampering and package-version drift never become verified availability',
  async () => {
    const outcomes = [];
    for (const name of ['index.aff', 'index.dic', 'license', 'package.json']) {
      const copy = await copiedProvider(async (root) => {
        const file = path.join(root, 'node_modules/dictionary-en', name);
        if (name === 'package.json') {
          const value = JSON.parse(await fs.readFile(file, 'utf8'));
          value.version = '0.0.0';
          await fs.writeFile(file, JSON.stringify(value));
        } else {
          const value = await fs.readFile(file);
          value[0] ^= 1;
          await fs.writeFile(file, value);
        }
      });
      try {
        await withProvider(async (provider) => {
          const availability = await provider.availability();
          assert.equal(availability.locales[0].status, 'unavailable');
          assert(!availability.locales[0].verified);
          assert.equal(availability.locales[1].status, 'available');
          const reply = await provider.check(request('en-US', ['hello']));
          assert.equal(reply.status, 'unavailable');
          assert(!JSON.stringify(reply).includes(copy.root));
          outcomes.push({ tampered: name, reply });
        }, copy.create);
      } finally {
        await copy.dispose();
      }
    }
    return { profile: 'same-byte-length tampering/metadata faults in isolated real-package copies', outcomes };
  }
);

await check(
  'Tampered copied engine code is refused before executing it; engine absence carries no verified provider',
  async () => {
    const copy = await copiedProvider(async (root) => {
      await fs.appendFile(path.join(root, 'node_modules/nspell/lib/index.js'), '\nthrow Error("must not execute");');
    });
    try {
      return await withProvider(async (provider) => {
        const availability = await provider.availability();
        assert(!availability.verifiedProvider);
        assert(availability.locales.every((item) => item.status === 'unavailable' && !item.verified));
        const reply = await provider.check(request('en-US', ['hello']));
        assert.equal(reply.status, 'unavailable');
        assert(!JSON.stringify(reply).includes('must not execute'));
        assert(!JSON.stringify(reply).includes(copy.root));
        return { profile: 'isolated actual engine source tamper', availability };
      }, copy.create);
    } finally {
      await copy.dispose();
    }
  }
);

await check(
  'Physical worker cancellation retires only its token, rejects stale ID reuse and resumes genuine checks',
  async () =>
    withProvider(async (provider) => {
      await provider.availability();
      const first = request('en-US', Array(100).fill('q'.repeat(120))),
        second = request('en-GB', ['colour']);
      const one = provider.check(first),
        two = provider.check(second);
      await provider.cancel(second.requestId);
      assert.equal((await two).status, 'cancelled');
      await new Promise((resolve) => setImmediate(resolve));
      await provider.cancel(first.requestId);
      const cancelled = await one;
      assert.equal(cancelled.status, 'cancelled');
      assert.deepEqual(cancelled.results, []);
      assert.throws(() => provider.check(first), /already been used/);
      const resumed = await provider.check(request('en-US', ['hello']));
      assert.equal(resumed.status, 'completed');
      assert.equal(resumed.results[0].correct, true);
      const fresh = request('en-US', Array(100).fill('q'.repeat(120))),
        pending = provider.check(fresh);
      await provider.dispose();
      assert.equal((await pending).status, 'cancelled');
      assert.throws(() => provider.check(request('en-US', ['hello'])), /closed/);
      assert((await provider.availability()).locales.every((item) => item.status === 'unavailable'));
      return { profile: 'actual worker/token cancellation; restarted verified dictionary', cancelled, resumed };
    })
);

await check('Host event loop stays responsive during real suggestion CPU; every returned word is bounded', async () =>
  withProvider(async (provider) => {
    await provider.availability();
    const pending = provider.check(request('en-US', Array(100).fill('helo')));
    let tick = false;
    await new Promise((resolve) =>
      setTimeout(() => {
        tick = true;
        resolve();
      }, 0)
    );
    assert(tick);
    const reply = await pending;
    assert.equal(reply.status, 'completed');
    assert.equal(reply.results.length, 100);
    assert(
      reply.results.every(
        (value) => !value.correct && value.suggestions.includes('hello') && value.suggestions.length <= 10
      )
    );
    assert(Buffer.byteLength(JSON.stringify(reply)) <= OFFLINE_SPELLING_LIMITS.replyBytes);
    return { profile: 'actual real-dictionary suggestion batch + main-loop timer', words: reply.results.length };
  })
);

await check(
  'Injected worker CPU stall is terminated at the hard budget and returns unavailable, never fabricated corrections',
  async () => {
    const copy = await copiedProvider(async (root) => {
      const file = path.join(root, 'adapters/offline-spelling-worker.mjs');
      const source = await fs.readFile(file, 'utf8');
      assert(source.includes('request=validateSpellingRequest(input);'));
      await fs.writeFile(
        file,
        source.replace(
          'request=validateSpellingRequest(input);',
          'request=validateSpellingRequest(input); while(true){}'
        )
      );
    });
    try {
      return await withProvider(async (provider) => {
        await provider.availability();
        const started = performance.now(),
          reply = await provider.check(request('en-US', ['hello']));
        assert.equal(reply.status, 'unavailable');
        assert.deepEqual(reply.results, []);
        assert.match(reply.reason, /time limit/);
        const elapsedMs = performance.now() - started;
        assert(elapsedMs >= OFFLINE_SPELLING_LIMITS.workMs - 100 && elapsedMs < OFFLINE_SPELLING_LIMITS.workMs + 4000);
        return {
          profile: 'explicit copied-worker CPU fault injection; real dictionaries loaded beforehand',
          elapsedMs,
          reply,
        };
      }, copy.create);
    } finally {
      await copy.dispose();
    }
  }
);
