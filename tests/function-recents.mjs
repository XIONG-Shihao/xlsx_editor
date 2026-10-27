import './support/scratch-dir.mjs';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
const component = fileURLToPath(new URL('../', import.meta.url)),
  run = await fs.mkdtemp(path.join(component, '.build/tests/function-recents-'));
const bundle = path.join(run, 'recents.mjs'),
  source = path.join(component, 'ui/functionRecents.ts');
await build({
  entryPoints: [source],
  bundle: true,
  format: 'esm',
  platform: 'node',
  outfile: bundle,
  logLevel: 'silent',
});
const { createFunctionRecentStore, FUNCTION_RECENTS_KEY } = await import(pathToFileURL(bundle));
const memory = () => {
  const items = new Map();
  return { items, getItem: (key) => items.get(key) ?? null, setItem: (key, value) => items.set(key, value) };
};
test('new preferences save only IDs in accepted order and survive another store', () => {
  const storage = memory(),
    store = createFunctionRecentStore(storage);
  assert.deepEqual(store.read(), []);
  store.recordAccepted('SUM');
  store.recordAccepted('IF');
  store.recordAccepted('SUM');
  assert.deepEqual(store.read(), ['SUM', 'IF']);
  assert.equal(storage.items.size, 1);
  assert.equal(storage.getItem(FUNCTION_RECENTS_KEY), '{' + '"version":1,"ids":["SUM","IF"]}');
  assert.deepEqual(createFunctionRecentStore(storage).read(), ['SUM', 'IF']);
  assert(Object.isFrozen(store.read()));
});
test('history is capped, duplicate-free and invalid catalog IDs cannot be recorded', () => {
  const store = createFunctionRecentStore(null),
    ids = [
      'SUM',
      'AVERAGE',
      'COUNT',
      'MIN',
      'MAX',
      'IF',
      'IFERROR',
      'LEN',
      'LEFT',
      'TRIM',
      'DATE',
      'YEAR',
      'MONTH',
      'DAY',
      'ABS',
      'ROUND',
      'DEC2HEX',
      'ROWS',
      'COLUMNS',
    ];
  ids.forEach((id) => store.recordAccepted(id));
  assert.deepEqual(store.read(), ids.slice(-12).reverse());
  assert.throws(() => store.recordAccepted('NPV'));
  assert.throws(() => store.recordAccepted({ id: 'SUM' }));
  assert.equal(store.read().length, 12);
});
test('corrupt, unknown, foreign-version or oversized profile values cannot introduce formulas or functions', () => {
  for (const raw of [
    'bad',
    'null',
    '[]',
    '{"version":2,"ids":["SUM"]}',
    '{"version":1,"ids":"SUM"}',
    ' '.repeat(4097),
    JSON.stringify({ version: 1, ids: Array(101).fill('SUM') }),
  ]) {
    const storage = memory();
    storage.setItem(FUNCTION_RECENTS_KEY, raw);
    assert.deepEqual(createFunctionRecentStore(storage).read(), []);
  }
  const storage = memory();
  storage.setItem(
    FUNCTION_RECENTS_KEY,
    JSON.stringify({ version: 1, ids: ['SUM', 'NPV', 'SUM', null, 'IF', 'SUM(A1:A9)', { formula: 'SUM(A1)' }] })
  );
  assert.deepEqual(createFunctionRecentStore(storage).read(), ['SUM', 'IF']);
});
test('denied reads or failed persistence keep accepted in-memory order and permit later recovery', () => {
  const storage = memory();
  let denyRead = true,
    denyWrite = true;
  const adapter = {
    getItem(key) {
      if (denyRead) throw Error('denied');
      return storage.getItem(key);
    },
    setItem(key, value) {
      if (denyWrite) throw Error('quota');
      storage.setItem(key, value);
    },
  };
  const store = createFunctionRecentStore(adapter);
  store.recordAccepted('SUM');
  store.recordAccepted('IF');
  assert.deepEqual(store.read(), ['IF', 'SUM']);
  denyRead = false;
  denyWrite = false;
  store.recordAccepted('LEN');
  assert.deepEqual(store.read(), ['LEN', 'IF', 'SUM']);
  assert.deepEqual(createFunctionRecentStore(storage).read(), ['LEN', 'IF', 'SUM']);
});
test('two workbook panes see accepted changes in a shared private profile and profiles stay isolated', () => {
  const shared = memory(),
    one = createFunctionRecentStore(shared),
    two = createFunctionRecentStore(shared),
    other = createFunctionRecentStore(memory());
  one.recordAccepted('SUM');
  assert.deepEqual(two.read(), ['SUM']);
  two.recordAccepted('DATE');
  assert.deepEqual(one.read(), ['DATE', 'SUM']);
  one.recordAccepted('LEN');
  assert.deepEqual(two.read(), ['LEN', 'DATE', 'SUM']);
  assert.deepEqual(other.read(), []);
});
