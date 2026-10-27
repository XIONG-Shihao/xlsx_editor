/** Actual JS Document/Indexed owners with a deliberately mocked native boundary.
 * This qualifies retirement/envelope behavior, never native UsedRange semantics. */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createHash } from 'node:crypto';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { indexedCache } from '../core/indexed-cache.mjs';
import { createWorkbook } from '../core/workbook-package.mjs';
import { validateUsedRange } from '../core/used-range.mjs';
import { engineQuestion } from './support/engine-questions.mjs';

const digest = (bytes) => createHash('sha256').update(bytes).digest('hex');
const zero = { minRow: 0, minColumn: 0, maxRow: 0, maxColumn: 0 };
const provenance = {
  factsSchema: 'workbook.xlsx.used-range-facts.v1',
  declaration: { status: 'absent', ref: null, bounds: null },
  sourceCells: null,
  sourceMerges: null,
  effectiveCells: null,
  sourceAxisFormatting: null,
  acceptedHistory: null,
  sourceDefaultFormatting: false,
};
async function fixture(t) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-used-range-envelope-')),
    source = path.join(directory, 'source.xlsx');
  const bytes = createWorkbook(),
    sourceRevision = digest(bytes);
  await fs.writeFile(source, bytes);
  const calls = [];
  let hold,
    replyTransform = (reply) => reply;
  const native = async (input) => {
    const answer = engineQuestion(input, ['used-range']);
    if (answer) return answer;
    calls.push(structuredClone(input));
    if (input.operation === 'copy-source') {
      await fs.copyFile(input.source, path.join(input.root, 'source.xlsx'));
      return {};
    }
    if (input.operation === 'index') {
      const index = {
        revision: sourceRevision,
        sheets: [
          {
            name: 'Sheet1',
            path: 'xl/worksheets/sheet1.xml',
            rows: 20,
            columns: 8,
            hiddenRows: [],
            visibleRows: [],
            rowHeights: {},
            storedCells: 0,
            formulaCount: 0,
          },
        ],
      };
      await fs.writeFile(path.join(input.root, 'metadata.xlsx'), bytes);
      await fs.writeFile(path.join(input.root, 'index.json'), JSON.stringify(index));
      return index;
    }
    if (input.operation === 'artifact-write') {
      await fs.writeFile(path.join(input.root, input.name), input.data);
      return {};
    }
    if (input.operation === 'compact') return { retainedVersions: input.keep.length };
    if (input.operation === 'calculation-policy' && input.method === 'initialize')
      return {
        policy: {
          schema: 1,
          mode: 'auto',
          recalculateBeforeSave: true,
          pendingFormulaCells: 0,
          canCalculateSheet: true,
          canChangeOptions: true,
        },
      };
    if (input.operation === 'used-range') {
      const reply = replyTransform({
        schema: 'workbook.xlsx.used-range.v1',
        policy: 'xlsx-used-range-v1',
        status: 'complete',
        sheet: 'Sheet1',
        sheetPath: input.sheet,
        version: input.version,
        sourceRevision,
        scopeMinMax: zero,
        range: 'A1',
        lastCell: 'A1',
        empty: true,
        provenance,
        reasons: [],
      });
      if (hold) {
        const pending = hold;
        hold = undefined;
        pending.started();
        await pending.wait;
      }
      return reply;
    }
    assert.fail('Unexpected mocked native operation: ' + input.operation);
  };
  const owner = new SpreadsheetDocument({
    native,
    recoveryDirectory: path.join(directory, 'recovery'),
    indexDirectory: path.join(directory, 'indices'),
    cacheIdentity: null,
  });
  t.after(async () => {
    await owner.close();
    await fs.rm(directory, { recursive: true, force: true });
  });
  await owner.open(source);
  return {
    directory,
    source,
    sourceRevision,
    owner,
    native,
    calls,
    setReply: (transform) => (replyTransform = transform),
    hold: () => {
      let started, release;
      const ready = new Promise((r) => (started = r)),
        wait = new Promise((r) => (release = r));
      hold = { started, wait };
      return { ready, release };
    },
  };
}
const request = (f) => ({
  operation: 'usedRange',
  documentId: f.owner.status().documentId,
  state: f.owner.status().state,
  sheet: 'Sheet1',
});
const receipt = (f) => {
  const status = f.owner.status();
  return {
    documentId: status.documentId,
    ownerId: status.view.storage.sessionId,
    state: status.state,
    revision: status.view.revision,
  };
};
async function footprint(f) {
  const files = [];
  async function walk(directory) {
    for (const entry of await fs.readdir(directory, { withFileTypes: true })) {
      const file = path.join(directory, entry.name);
      if (entry.isDirectory()) await walk(file);
      else files.push({ name: path.relative(f.directory, file), sha256: digest(await fs.readFile(file)) });
    }
  }
  await walk(f.directory);
  return { status: f.owner.status(), files: files.sort((a, b) => a.name.localeCompare(b.name)) };
}

test('JS owner binds original public head and leaves source/recovery/index state unchanged', async (t) => {
  const f = await fixture(t),
    before = await footprint(f),
    expected = receipt(f),
    reply = (await f.owner.request(request(f))).usedRange;
  validateUsedRange(reply, { sheet: 'Sheet1', receipt: expected });
  assert.equal(reply.native.version, 0);
  assert.equal(reply.native.sourceRevision, f.sourceRevision);
  assert.equal(reply.range, 'A1');
  assert.notEqual(reply.range, 'A1:H20', 'mock index display padding is never used');
  assert.deepEqual(await footprint(f), before);
  assert.equal(f.calls.filter((c) => c.operation === 'used-range').length, 1);
});

test('foreign/stale public frames and unsupported request authority refuse before native query', async (t) => {
  const f = await fixture(t),
    query = request(f),
    before = await footprint(f);
  for (const extra of [{ documentId: 'other' }, { state: 'unaccepted-head' }, { commands: [] }, { range: 'A1:H20' }])
    await assert.rejects(() => f.owner.request({ ...query, ...extra }), /another|stale|fields/);
  await assert.rejects(() => f.owner.request({ ...query, sheet: 'Other' }), { code: 'internal-unknown-reference' });
  assert.equal(f.calls.filter((c) => c.operation === 'used-range').length, 0);
  assert.deepEqual(await footprint(f), before);
});

test('the asynchronous read retains its original request instead of adopting caller retargeting', async (t) => {
  const f = await fixture(t),
    original = request(f),
    before = await footprint(f),
    expected = receipt(f),
    timing = f.hold();
  const pending = f.owner.request(original);
  original.documentId = 'other';
  original.state = 'other';
  original.sheet = 'Other';
  await timing.ready;
  timing.release();
  const reply = (await pending).usedRange;
  validateUsedRange(reply, { sheet: 'Sheet1', receipt: expected });
  assert.deepEqual(await footprint(f), before);
});

test('native selected-version mismatch and incomplete geometry survive the real owner path', async (t) => {
  const f = await fixture(t),
    before = await footprint(f);
  f.setReply((reply) => ({ ...reply, version: 1 }));
  await assert.rejects(() => f.owner.request(request(f)), { code: /^(layout-stale|note-stale)$/ });
  f.setReply((reply) => ({
    ...reply,
    status: 'incomplete',
    scopeMinMax: null,
    range: null,
    lastCell: null,
    empty: null,
    reasons: [{ code: 'source-facts-incomplete', message: 'Source facts unavailable.' }],
  }));
  const reply = (await f.owner.request(request(f))).usedRange;
  validateUsedRange(reply, { receipt: receipt(f), sheet: 'Sheet1' });
  assert.equal(reply.status, 'incomplete');
  assert.equal(reply.range, null);
  assert.deepEqual(await footprint(f), before);
});

test('cancellation or actual owner retirement rejects a completed held native reply', async (t) => {
  const f = await fixture(t),
    before = await footprint(f),
    timing = f.hold();
  let cancelled = false;
  const pending = f.owner.request(request(f), {
    checkCancelled: () => {
      if (cancelled) throw Error('UsedRange test cancellation');
    },
  });
  await timing.ready;
  cancelled = true;
  timing.release();
  await assert.rejects(() => pending, /cancellation/);
  assert.deepEqual(await footprint(f), before);
  const retired = await fixture(t),
    retiring = retired.hold(),
    late = retired.owner.request(request(retired));
  await retiring.ready;
  const closing = retired.owner.close();
  retiring.release();
  await assert.rejects(() => late, /cancelled|closed|retired/);
  await closing;
});

test('a hash-valid old-schema cache is missed before opening the new UsedRange owner', async (t) => {
  const f = await fixture(t),
    oldRoot = path.join(f.directory, 'old-index'),
    cacheDirectory = path.join(f.directory, 'derived-cache');
  await fs.mkdir(oldRoot);
  await fs.copyFile(f.source, path.join(oldRoot, 'source.xlsx'));
  await fs.copyFile(f.source, path.join(oldRoot, 'metadata.xlsx'));
  // Only the cache's filesystem/hash policy is exercised; this placeholder is
  // never opened as SQLite or offered to a native process.
  await fs.writeFile(path.join(oldRoot, 'cells.sqlite'), 'mock-boundary old index');
  await fs.writeFile(
    path.join(oldRoot, 'index.json'),
    JSON.stringify({ revision: f.sourceRevision, sheets: [{ path: 'xl/worksheets/sheet1.xml', rows: 20, columns: 8 }] })
  );
  const oldCache = await indexedCache(f.source, 'xlsx-b9-calculation-policy:controlled-implementation:null', {
    directory: cacheDirectory,
  });
  assert.equal((await oldCache.publish(oldRoot, 0, { reuseCalculation: false })).published, true);
  const retained = await fs.readdir(cacheDirectory),
    start = f.calls.length;
  const owner = new IndexedWorkbook(
    f.native,
    () => assert.fail('No fallback writer'),
    () => assert.fail('No formulas in boundary fixture')
  );
  t.after(() => owner.close());
  await owner.open(f.source, {
    cacheIdentity: 'controlled-implementation',
    cacheMode: 'saved-only',
    cacheOptions: { directory: cacheDirectory },
    storageDirectory: path.join(f.directory, 'new-indices'),
  });
  assert.equal(owner.cacheHit, false);
  assert.deepEqual(
    f.calls
      .slice(start)
      .filter((call) => ['copy-source', 'index'].includes(call.operation))
      .map((call) => call.operation),
    ['copy-source', 'index']
  );
  assert.deepEqual(
    await fs.readdir(cacheDirectory),
    retained,
    'old disposable evidence is left intact by a namespace miss'
  );
  const reply = await owner.usedRange('0', 'Sheet1');
  assert.equal(reply.empty, true);
  assert.equal(reply.range, 'A1');
  assert.equal(reply.native.sourceRevision, f.sourceRevision);
});
