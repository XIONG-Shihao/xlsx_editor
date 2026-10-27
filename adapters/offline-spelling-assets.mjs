import { formatHashSalt, FORMAT } from '../core/format.mjs';
/** Trusted installed assets only. No renderer paths or runtime dictionary download. */
import fs from 'node:fs/promises';
import { constants } from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { createHash } from 'node:crypto';
import { refusal } from '../core/refusal-codes.mjs';

const require = createRequire(import.meta.url);
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const freeze = (value) => {
  if (value && typeof value === 'object') {
    for (const child of Object.values(value)) freeze(child);
    Object.freeze(value);
  }
  return value;
};
export const OFFLINE_SPELLING_LIMITS = freeze({
  words: 100,
  wordCodePoints: 256,
  requestBytes: 32768,
  suggestions: 10,
  replyBytes: 1048576,
  queuedRequests: 4,
  requestIds: 4096,
  startupMs: 15000,
  workMs: 5000,
  workerOldMb: 128,
  workerYoungMb: 16,
});
const engineFiles = [
  'lib/add.js',
  'lib/correct.js',
  'lib/dictionary.js',
  'lib/index.js',
  'lib/personal.js',
  'lib/remove.js',
  'lib/spell.js',
  'lib/suggest.js',
  'lib/util/add.js',
  'lib/util/affix.js',
  'lib/util/apply.js',
  'lib/util/casing.js',
  'lib/util/dictionary.js',
  'lib/util/exact.js',
  'lib/util/flag.js',
  'lib/util/form.js',
  'lib/util/normalize.js',
  'lib/util/rule-codes.js',
  'lib/word-characters.js',
];
export const OFFLINE_SPELLING_ENGINE = freeze({
  package: 'nspell',
  version: '2.1.5',
  codeSha256: '7c646dc1cdded8fcb1bf4e892f98185b4b5f6dda4ea3d36be13bc2b9ae32c807',
  license: 'MIT',
  licenseSha256: 'ca4662cb5d1b738fbe5350c0d5485ba11773b4b7208974082ae6e129a52d631d',
  dependency: {
    package: 'is-buffer',
    version: '2.0.5',
    codeSha256: '49e2297954c33cdbdc9f81e9241af4baa703c7311981a3f2e12ae769ecb2f750',
    license: 'MIT',
    licenseSha256: 'c7cc929b57080f4b9d0c6cf57669f0463fc5b39906344dfc8d3bc43426b30eac',
  },
});
export const OFFLINE_SPELLING_PROVIDER_ID = hash(
  JSON.stringify([formatHashSalt('SPELLING_PROVIDER_V1'), OFFLINE_SPELLING_ENGINE])
);
const dictionary = (locale, packageName, version, dicSha256, dicBytes, licenseSha256, licenseBytes) => {
  const value = {
    locale,
    package: packageName,
    version,
    affSha256: '8ae1f19d4840d957728ad90555d5a8dff6cc5c046279c95ff0c00fc0a0136c7b',
    affBytes: 3086,
    dicSha256,
    dicBytes,
    license: '(MIT AND BSD)',
    licenseSha256,
    licenseBytes,
  };
  return freeze({
    ...value,
    dictionaryId: hash(JSON.stringify([formatHashSalt('SPELLING_DICTIONARY_V1'), OFFLINE_SPELLING_PROVIDER_ID, value])),
  });
};
export const OFFLINE_SPELLING_DICTIONARIES = freeze([
  dictionary(
    'en-US',
    'dictionary-en',
    '4.0.0',
    'f0b1a234bd178bdd01875b2a392a9647f888b8fe879f79c52aae62c2759b3647',
    551762,
    '2a7e8d8ae9e8facc84818546ae2a8d83aec5e9c80a675ff789acd1c338b53b3d',
    15731
  ),
  dictionary(
    'en-GB',
    'dictionary-en-gb',
    '3.0.0',
    '869fe17ba4ee4b5401c60a666ee2d6a3dcc237f460b7294435df1cc6a799aa57',
    552106,
    '81b644347b40804c25811267efbea024a30b2bb1642fefa9c9c50f7b0bbcb67f',
    15739
  ),
]);
export function unavailableSpellingAvailability(reason) {
  return {
    schema: 1,
    provider: 'nspell',
    providerVersion: OFFLINE_SPELLING_ENGINE.version,
    providerId: OFFLINE_SPELLING_PROVIDER_ID,
    expectedProvider: OFFLINE_SPELLING_ENGINE,
    locales: OFFLINE_SPELLING_DICTIONARIES.map((expected) => ({
      locale: expected.locale,
      status: 'unavailable',
      expected,
      reason,
    })),
  };
}
export function validateSpellingRequest(input) {
  if (
    !input ||
    typeof input !== 'object' ||
    Array.isArray(input) ||
    Object.keys(input).sort().join(',') !== 'locale,requestId,words' ||
    typeof input.requestId !== 'string' ||
    !/^[a-f0-9]{8}-[a-f0-9]{4}-[1-8][a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}$/.test(input.requestId) ||
    typeof input.locale !== 'string' ||
    !input.locale.length ||
    input.locale.length > 64 ||
    !Array.isArray(input.words) ||
    input.words.length < 1 ||
    input.words.length > OFFLINE_SPELLING_LIMITS.words
  )
    throw refusal('internal-invalid-field', 'Invalid offline spelling request.');
  for (const word of input.words) {
    if (
      typeof word !== 'string' ||
      !word.length ||
      Array.from(word).length > OFFLINE_SPELLING_LIMITS.wordCodePoints ||
      /[\s\p{C}]/u.test(word)
    )
      throw refusal('spelling-bad-input', 'Offline spelling requires bounded individual words.');
  }
  if (Buffer.byteLength(JSON.stringify(input)) > OFFLINE_SPELLING_LIMITS.requestBytes)
    throw refusal('spelling-limit', 'Offline spelling request exceeds 32 KiB. Split the words into smaller batches.');
  return freeze({ requestId: input.requestId, locale: input.locale, words: [...input.words] });
}
async function bytes(file, limit) {
  const handle = await fs.open(file, constants.O_RDONLY | constants.O_NOFOLLOW);
  try {
    const stat = await handle.stat();
    if (!stat.isFile() || stat.size > limit) throw refusal('spelling-limit', 'Asset bound differs');
    return await handle.readFile();
  } finally {
    await handle.close();
  }
}
async function metadata(root, expected) {
  const value = JSON.parse(await bytes(path.join(root, 'package.json'), 32768));
  if (value.name !== expected.package || value.version !== expected.version || value.license !== expected.license)
    throw refusal('internal-mismatch', 'Asset package identity differs');
}
async function code(root, names, expected) {
  const records = [];
  for (const relative of names) {
    const value = await bytes(path.join(root, relative), 65536);
    records.push({ relative, sha256: hash(value), bytes: value.length });
  }
  if (hash(JSON.stringify(records)) !== expected.codeSha256)
    throw refusal('internal-mismatch', 'Asset code identity differs');
  const licenseFile = expected.package === 'nspell' ? 'license' : expected.package === 'is-buffer' ? 'LICENSE' : null;
  if (!licenseFile) throw refusal('spelling-bad-input', 'Asset package is not trusted');
  if (hash(await bytes(path.join(root, licenseFile), 32768)) !== expected.licenseSha256)
    throw refusal('internal-mismatch', 'Asset license identity differs');
}
export async function loadSpellingAssets() {
  const availability = unavailableSpellingAvailability(
    'The offline spelling engine is missing or its installed identity differs.'
  );
  let nspell;
  try {
    const entry = require.resolve('nspell'),
      root = path.dirname(path.dirname(entry));
    await metadata(root, OFFLINE_SPELLING_ENGINE);
    await code(root, engineFiles, OFFLINE_SPELLING_ENGINE);
    const engineRequire = createRequire(entry),
      dependencyRoot = path.dirname(engineRequire.resolve('is-buffer'));
    await metadata(dependencyRoot, OFFLINE_SPELLING_ENGINE.dependency);
    await code(dependencyRoot, ['index.js'], OFFLINE_SPELLING_ENGINE.dependency);
    nspell = require(entry);
    availability.verifiedProvider = OFFLINE_SPELLING_ENGINE;
  } catch {
    return { availability, engines: new Map() };
  }
  const engines = new Map();
  for (const item of availability.locales) {
    try {
      const expected = item.expected,
        root = path.dirname(require.resolve(expected.package));
      await metadata(root, expected);
      const aff = await bytes(path.join(root, 'index.aff'), expected.affBytes),
        dic = await bytes(path.join(root, 'index.dic'), expected.dicBytes),
        license = await bytes(path.join(root, 'license'), expected.licenseBytes);
      if (
        aff.length !== expected.affBytes ||
        dic.length !== expected.dicBytes ||
        license.length !== expected.licenseBytes ||
        hash(aff) !== expected.affSha256 ||
        hash(dic) !== expected.dicSha256 ||
        hash(license) !== expected.licenseSha256
      )
        throw refusal('internal-mismatch', 'Dictionary identity differs');
      engines.set(item.locale, nspell({ aff, dic }));
      item.status = 'available';
      item.verified = expected;
      delete item.reason;
    } catch {
      item.reason = 'The ' + item.locale + ' offline dictionary is missing or its installed identity differs.';
    }
  }
  return { availability, engines };
}
