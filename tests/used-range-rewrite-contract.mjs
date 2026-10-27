/** Controlled JavaScript protocol/FS tests only. The fake replies below test
 * routing, refusal and lifetime boundaries, never native geometry or SQLite
 * authority. Actual native writer/transfer qualification is a separate suite. */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createHash, randomUUID } from 'node:crypto';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook } from '../core/index.mjs';
import { indexedPackageEdit, readWorkbookMetadata } from '../core/workbook-package.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { resetIndexedSnapshotExtents } from '../core/indexed-checkpoint.mjs';
import { engineQuestion } from './support/engine-questions.mjs';

const V2 = 'workbook.xlsx.used-range-facts.v2',
  V1 = 'workbook.xlsx.used-range-facts.v1';
const sha = (bytes) => createHash('sha256').update(bytes).digest('hex');
const resize = { operation: 'resizeRow', sheet: 'Sheet1', index: 0, height: 24 };
const deferred = () => {
  let resolve;
  const promise = new Promise((r) => {
    resolve = r;
  });
  return { promise, resolve };
};
async function bounded(promise, label) {
  let timer;
  try {
    return await Promise.race([
      promise,
      new Promise((_, reject) => {
        timer = setTimeout(() => reject(Error('Controlled protocol timeout: ' + label)), 5000);
      }),
    ]);
  } finally {
    clearTimeout(timer);
  }
}
async function fixture(configuration = {}) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-rewrite-protocol-'));
  const source = path.join(directory, 'source.xlsx');
  await fs.writeFile(source, createWorkbook());
  const calls = [],
    inputId = randomUUID(),
    rewriteId = randomUUID(),
    books = [];
  const policy = {
    mode: configuration.manual ? 'manual' : 'auto',
    recalculateBeforeSave: true,
    pendingFormulaCells: 0,
    pendingCacheObjects: 0,
    protectedSource: false,
  };
  const native = async (request) => {
    const answer = engineQuestion(request);
    if (answer) return answer;
    calls.push(structuredClone(request));
    switch (request.operation) {
      case 'copy-source':
        await fs.copyFile(request.source, path.join(request.root, 'source.xlsx'));
        return { copied: true };
      case 'index': {
        const bytes = await fs.readFile(path.join(request.root, 'source.xlsx'));
        const index = {
          revision: sha(bytes),
          sourceFactSchema: configuration.facts ?? V2,
          sheets: [
            {
              path: 'xl/worksheets/sheet1.xml',
              rows: 20,
              columns: 8,
              storedCells: 0,
              formulaCount: 0,
              specialFormulaCount: 0,
              extended: false,
              dataStart: 1,
              hiddenRows: [],
              visibleRows: [],
              rowHeights: [],
            },
          ],
        };
        await fs.writeFile(path.join(request.root, 'index.json'), JSON.stringify(index));
        await fs.writeFile(path.join(request.root, 'metadata.xlsx'), bytes);
        await fs.writeFile(path.join(request.root, 'cells.sqlite'), 'controlled protocol database');
        return index;
      }
      case 'artifact-write':
        await fs.writeFile(path.join(request.root, request.name), request.data);
        return { written: true };
      case 'calculation-policy':
        assert.equal(request.method, 'initialize');
        return { policy: { ...policy } };
      case 'spill-stage':
        return { version: 3, groups: 0 };
      case 'export-plan':
        return { styleCount: 0, special: false, incomplete: false };
      case 'used-range-export-needed':
        return {
          schema: 'workbook.xlsx.used-range-export-needed.v1',
          version: request.version,
          sourceRevision: request.sourceRevision,
          required: false,
        };
      case 'export': {
        if (request.internalRewrite) {
          assert.equal(path.dirname(request.rewriteTargetRoot), path.dirname(request.root));
          assert.deepEqual(await fs.readdir(request.rewriteTargetRoot), []);
          assert.equal(request.publication, undefined);
        }
        const bytes = await fs.readFile(path.join(request.root, 'source.xlsx'));
        await fs.writeFile(path.join(request.root, 'export.xlsx'), bytes);
        const reply = {
          file: 'export.xlsx',
          revision: sha(bytes),
          ...(request.internalRewrite ? { rewriteInputId: inputId } : {}),
        };
        if (configuration.exportHold) await configuration.exportHold(request);
        return configuration.exportReply ? configuration.exportReply(reply, request) : reply;
      }
      case 'layout': {
        if (configuration.writerFailure) throw Error('Controlled writer refusal');
        await fs.copyFile(path.join(request.root, request.source), path.join(request.root, 'layout.xlsx'));
        const reply = { file: 'layout.xlsx', ...(request.rewriteInputId ? { rewriteId } : {}) };
        return configuration.writerReply ? configuration.writerReply(reply) : reply;
      }
      case 'used-range-transfer': {
        if (configuration.transferFailure) throw Error('Controlled transfer refusal');
        const reply = {
          schema: 'workbook.xlsx.used-range-transfer.v1',
          version: request.version + 17,
          sourceRevision: request.sourceRevision,
          previousVersion: request.previousVersion,
          previousSourceRevision: request.previousSourceRevision,
          rewriteId: request.rewriteId,
          transferredSheets: 1,
          incompleteSheets: 0,
        };
        return configuration.transferReply ? configuration.transferReply(reply, request) : reply;
      }
      case 'dependencies':
        assert.equal(request.method, 'cache-status');
        return { qualified: false, reason: 'Controlled fixture excludes calculation certification' };
      case 'used-range-source-snapshot': {
        if (configuration.snapshotFailure) throw Error(configuration.snapshotFailure);
        const reply = {
          schema: 'workbook.xlsx.used-range-source-snapshot.v1',
          version: request.version,
          sourceRevision: request.sourceRevision,
        };
        // Changing only the copy demonstrates the clone boundary, not a SQL reset.
        await fs.writeFile(path.join(request.root, 'cells.sqlite'), 'sanitized controlled protocol copy');
        return configuration.snapshotReply ? configuration.snapshotReply(reply) : reply;
      }
      case 'compact':
        return { compacted: true };
      default:
        assert.fail('Unexpected controlled protocol operation: ' + request.operation);
    }
  };
  const make = () => {
    const book = new IndexedWorkbook(
      native,
      () => assert.fail('No legacy writer'),
      () => assert.fail('No evaluator')
    );
    books.push(book);
    return book;
  };
  const options = { storageDirectory: path.join(directory, 'owners') };
  const cacheOptions = {
    ...options,
    cacheIdentity: 'controlled-rewrite-protocol',
    cacheOptions: { directory: path.join(directory, 'cache') },
  };
  return {
    directory,
    source,
    calls,
    inputId,
    rewriteId,
    make,
    options,
    cacheOptions,
    async close() {
      try {
        await Promise.all(books.map((book) => book.close()));
      } finally {
        await fs.rm(directory, { recursive: true, force: true });
      }
    },
  };
}
async function owner(f, options = f.options) {
  const book = f.make();
  await book.open(f.source, options);
  const root = f.calls.findLast((call) => call.operation === 'index').root;
  return { book, root, view: book.view('0') };
}
async function unchanged({ book, root, view }) {
  assert.deepEqual(book.view('0'), view);
  assert.deepEqual(
    // The engine keeps its disk-budget lock beside the owner record.
    (await fs.readdir(path.dirname(root))).filter((name) => !['owner.json', '.quota.lock'].includes(name)),
    ['base']
  );
  assert.equal(book.diagnosticStats().retainedStates, 1);
}

test('source snapshot sanitizer accepts only an exact identity-bound native receipt', async () => {
  const revision = 'a'.repeat(64),
    calls = [];
  const native = async (request) => {
    calls.push(request);
    return {
      schema: 'workbook.xlsx.used-range-source-snapshot.v1',
      version: request.version,
      sourceRevision: request.sourceRevision,
    };
  };
  await resetIndexedSnapshotExtents('/private/cache-copy', 7, revision, native);
  assert.deepEqual(calls, [
    {
      action: 'workbook',
      operation: 'used-range-source-snapshot',
      root: '/private/cache-copy',
      version: 7,
      sourceRevision: revision,
    },
  ]);
  for (const patch of [
    { schema: 'other' },
    { version: 8 },
    { sourceRevision: 'b'.repeat(64) },
    { callerBounds: {} },
    { diagnostics: {} },
  ])
    await assert.rejects(
      resetIndexedSnapshotExtents('/private/cache-copy', 7, revision, async () => ({
        schema: 'workbook.xlsx.used-range-source-snapshot.v1',
        version: 7,
        sourceRevision: revision,
        ...patch,
      })),
      /receipt/
    );
  for (const version of [-1, 1.5, Infinity])
    await assert.rejects(resetIndexedSnapshotExtents('/private/cache-copy', version, revision, native), /identity/);
  // A real pure metadata planner, not a fake native success: only the generic
  // simulator's scheduling delta may disappear from a createTable plan.
  for (const mode of ['absent', 'auto', 'manual']) {
    const parts = unzipSync(createWorkbook());
    let workbook = strFromU8(parts['xl/workbook.xml']).replace(/<calcPr\b[^>]*(?:\/>|>[\s\S]*?<\/calcPr>)/g, '');
    if (mode !== 'absent')
      workbook = workbook.replace(
        '</workbook>',
        `<calcPr calcId="162194" calcMode="${mode}" calcOnSave="${mode === 'manual' ? 0 : 1}" fullCalcOnLoad="0" forceFullCalc="0"/></workbook>`
      );
    parts['xl/workbook.xml'] = strToU8(workbook);
    const bytes = zipSync(parts),
      view = readWorkbookMetadata(bytes).view,
      command = { operation: 'createTable', sheet: 'Sheet1', range: 'A1:B4', tableName: 'DataTable' },
      headers = [
        { sheet: 'Sheet1', cell: 'A1', value: 'Category' },
        { sheet: 'Sheet1', cell: 'B1', value: 'Amount' },
      ];
    const plan = indexedPackageEdit(bytes, command, view, headers);
    assert.deepEqual(Object.keys(plan.replacements).sort(), [
      '[Content_Types].xml',
      'xl/tables/table1.xml',
      'xl/worksheets/_rels/sheet1.xml.rels',
    ]);
    assert(!Object.hasOwn(plan.replacements, 'xl/workbook.xml'));
    assert.equal(strFromU8(unzipSync(bytes)['xl/workbook.xml']), workbook);
    assert.throws(
      () => indexedPackageEdit(bytes, command, view, [{ operation: 'addSheet', sheet: 'Unrelated' }, ...headers]),
      /workbook.*metadata|scheduling/i,
      'An unrelated simulator workbook change must not be discarded'
    );
  }
});

test('v2 layout allocates its empty target before export and transfers before publishing', async () => {
  const f = await fixture();
  try {
    const o = await owner(f),
      result = await o.book.apply('0', [resize]);
    const exported = f.calls.find((call) => call.operation === 'export'),
      writer = f.calls.find((call) => call.operation === 'layout'),
      transfer = f.calls.find((call) => call.operation === 'used-range-transfer');
    assert.equal(exported.internalRewrite, true);
    assert.equal(writer.source, 'export.xlsx');
    assert.equal(writer.rewriteInputId, f.inputId);
    assert.equal(writer.version, 0);
    assert.deepEqual(writer.command, resize);
    assert.equal(transfer.root, exported.rewriteTargetRoot);
    assert.equal(transfer.previousRoot, o.root);
    assert.equal(transfer.previousVersion, 0);
    assert.equal(transfer.rewriteId, f.rewriteId);
    assert.equal(transfer.version, 0);
    assert(
      f.calls.findIndex((call) => call.operation === 'index' && call.root === transfer.root) < f.calls.indexOf(transfer)
    );
    assert.equal(o.book.diagnosticStats().retainedStates, 2);
    assert.equal(result.view.storage.state, '1');
    assert.deepEqual(o.book.view('0'), o.view);
    assert.equal(
      f.calls.filter((call) => call.operation === 'used-range-source-snapshot').length,
      0,
      'An accepted owner never uses fresh-cache reset'
    );
  } finally {
    await f.close();
  }
});

test('Manual reindex transfer uses the actual spill-stage version and retains policy', async () => {
  const f = await fixture({ manual: true });
  try {
    const o = await owner(f),
      result = await o.book.apply('0', [resize]),
      transfer = f.calls.find((call) => call.operation === 'used-range-transfer');
    assert.equal(transfer.previousVersion, 3);
    assert.equal(transfer.version, 3);
    assert.equal(result.view.calculationPolicy.mode, 'manual');
    assert.equal(result.view.calculationPolicy.recalculateBeforeSave, true);
  } finally {
    await f.close();
  }
});

test('legacy v1 remains usable without new export, writer, transfer or cache opcodes', async () => {
  const f = await fixture({ facts: V1 });
  try {
    const o = await owner(f, f.cacheOptions);
    await o.book.apply('0', [resize]);
    const writer = f.calls.find((call) => call.operation === 'layout');
    assert.equal(writer.rewriteInputId, undefined);
    assert.equal(writer.command, undefined);
    assert.equal(writer.version, undefined);
    assert(!f.calls.some((call) => ['used-range-transfer', 'used-range-source-snapshot'].includes(call.operation)));
    assert(!f.calls.some((call) => call.operation === 'export' && call.internalRewrite));
  } finally {
    await f.close();
  }
});

test('malformed internal export receipts refuse before the writer and preserve the owner', async () => {
  for (const mutate of [
    (reply) => ({ file: reply.file, revision: reply.revision }),
    (reply) => ({ ...reply, rewriteInputId: 'caller-token' }),
    (reply) => ({ ...reply, revision: 'changed' }),
    (reply) => ({ ...reply, file: '../export.xlsx' }),
    (reply) => ({ ...reply, callerBounds: {} }),
  ]) {
    const f = await fixture({ exportReply: mutate });
    try {
      const o = await owner(f);
      await assert.rejects(o.book.apply('0', [resize]), { code: 'internal-invalid-field' });
      assert(!f.calls.some((call) => call.operation === 'layout'));
      await unchanged(o);
    } finally {
      await f.close();
    }
  }
});

test('missing writer tokens and unsafe output paths never become indexed candidates', async () => {
  for (const mutate of [
    (reply) => ({ file: reply.file }),
    (reply) => ({ ...reply, rewriteId: 'invalid' }),
    (reply) => ({ ...reply, file: '../layout.xlsx' }),
  ]) {
    const f = await fixture({ writerReply: mutate });
    try {
      const o = await owner(f);
      await assert.rejects(o.book.apply('0', [resize]), { code: /^(internal-invalid-field|internal-missing-field)$/ });
      assert.equal(f.calls.filter((call) => call.operation === 'index').length, 1);
      await unchanged(o);
    } finally {
      await f.close();
    }
  }
});

test('transfer identities, new version, counts and closed fields are mandatory', async () => {
  for (const patch of [
    { schema: 'other' },
    { sourceRevision: 'b'.repeat(64) },
    { previousSourceRevision: 'b'.repeat(64) },
    { previousVersion: 999 },
    { rewriteId: randomUUID() },
    { version: 0 },
    { version: -1 },
    { transferredSheets: 2 },
    { incompleteSheets: -1 },
    { callerBounds: {} },
  ]) {
    const f = await fixture({ transferReply: (reply) => ({ ...reply, ...patch }) });
    try {
      const o = await owner(f);
      await assert.rejects(o.book.apply('0', [resize]), /extent transfer/);
      await unchanged(o);
    } finally {
      await f.close();
    }
  }
});

test('native writer and transfer failures clean only their unaccepted target', async () => {
  for (const configuration of [{ writerFailure: true }, { transferFailure: true }]) {
    const f = await fixture(configuration);
    try {
      const o = await owner(f),
        source = await fs.readFile(path.join(o.root, 'source.xlsx')),
        index = await fs.readFile(path.join(o.root, 'index.json'));
      await assert.rejects(o.book.apply('0', [resize]), /Controlled.*refusal/);
      await unchanged(o);
      assert.deepEqual(await fs.readFile(path.join(o.root, 'source.xlsx')), source);
      assert.deepEqual(await fs.readFile(path.join(o.root, 'index.json')), index);
    } finally {
      await f.close();
    }
  }
});

test('ordinary Copy and publication export never request internal rewrite authority', async () => {
  const f = await fixture();
  try {
    const o = await owner(f);
    await o.book.export('0');
    await o.book.export('0', { publication: true });
    assert(
      !f.calls.some((call) => call.internalRewrite || call.rewriteInputId || call.operation === 'used-range-transfer')
    );
    assert.equal(f.calls.findLast((call) => call.operation === 'export').publication, true);
    await unchanged(o);
  } finally {
    await f.close();
  }
});

test('closing during a held completed export refuses publication and removes private roots', async () => {
  const entered = deferred(),
    release = deferred(),
    f = await fixture({
      exportHold: async () => {
        entered.resolve();
        await release.promise;
      },
    });
  try {
    const o = await owner(f),
      pending = o.book.apply('0', [resize]);
    const rejected = assert.rejects(pending, /cancelled|closed/i);
    void rejected.catch(() => {});
    await bounded(
      Promise.race([
        entered.promise,
        pending.then(
          () => {
            throw Error('Export hold was bypassed');
          },
          (error) => {
            throw error;
          }
        ),
      ]),
      'export hold'
    );
    const closed = o.book.close();
    release.resolve();
    await bounded(rejected, 'cancel refusal');
    await bounded(closed, 'owner close');
    assert(!f.calls.some((call) => call.operation === 'layout'));
    assert.equal(o.book.diagnosticStats().retainedStates, 0);
    await assert.rejects(fs.stat(path.dirname(o.root)), { code: 'ENOENT' });
  } finally {
    release.resolve();
    await f.close();
  }
});

test('initial cache clones and restored v2 copies reset once before accepting a cached owner', async () => {
  const configuration = {},
    f = await fixture(configuration);
  try {
    const first = await owner(f, f.cacheOptions),
      sourceDB = await fs.readFile(path.join(first.root, 'cells.sqlite'));
    const reset = f.calls.filter((call) => call.operation === 'used-range-source-snapshot');
    assert.equal(reset.length, 1);
    assert.notEqual(reset[0].root, first.root);
    assert(reset[0].root.startsWith(f.cacheOptions.cacheOptions.directory + path.sep));
    assert.deepEqual(await fs.readFile(path.join(first.root, 'cells.sqlite')), sourceDB);
    await first.book.close();
    f.calls.length = 0;
    const second = f.make();
    await second.open(f.source, f.cacheOptions);
    assert.equal(second.cacheHit, true);
    assert.equal(f.calls.filter((call) => call.operation === 'used-range-source-snapshot').length, 1);
    assert(!f.calls.some((call) => call.operation === 'index'));
    const resetAt = f.calls.findIndex((call) => call.operation === 'used-range-source-snapshot'),
      initializeAt = f.calls.findIndex((call) => call.operation === 'calculation-policy');
    assert(resetAt < initializeAt);
    await second.close();
    configuration.snapshotReply = (reply) => ({ ...reply, sourceRevision: 'f'.repeat(64) });
    f.calls.length = 0;
    const fallback = f.make();
    await fallback.open(f.source, f.cacheOptions);
    assert.equal(fallback.cacheHit, false);
    assert.equal(
      f.calls.filter((call) => call.operation === 'index').length,
      1,
      'A refused restored receipt must invalidate and cold import'
    );
    await fallback.close();
    delete configuration.snapshotReply;
    const healthy = f.make();
    await healthy.open(f.source, f.cacheOptions);
    await healthy.close();
    configuration.snapshotFailure = 'Workbook session disk budget exhausted; controlled quota failure';
    await assert.rejects(
      f.make().open(f.source, f.cacheOptions),
      /disk budget/,
      'Quota failures propagate rather than silently importing another owner'
    );
  } finally {
    await f.close();
  }
});
