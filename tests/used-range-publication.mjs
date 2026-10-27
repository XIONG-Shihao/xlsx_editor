/** Real JS owners + filesystem publication, deliberately mocked native geometry. */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import { syncBuiltinESMExports } from 'node:module';
import path from 'node:path';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { validateRecoveryManifest, validateSavedSource } from '../adapters/storage.mjs';
import { validateUsedRange } from '../core/used-range.mjs';
import { publicationFixture, farClear, edit, used, hash } from './support/used-range-publication-fixture.mjs';

test('Copy preserves the sticky source frame; successful Save adopts fresh bytes at the same public head', async (t) => {
  const f = await publicationFixture(t),
    far = await farClear(f.owner),
    before = f.owner.status(),
    sticky = await used(f.owner);
  assert.equal(sticky.lastCell, 'Z100');
  const copy = path.join(f.directory, 'copy.xlsx');
  await f.owner.request({ operation: 'copy', destination: copy, expectedRevision: null });
  assert.deepEqual(f.owner.status(), before);
  assert.deepEqual(await used(f.owner), sticky);
  const saved = await f.owner.request({ operation: 'save' }),
    fresh = await used(f.owner);
  assert.equal(saved.state, before.state);
  assert.equal(saved.dirty, false);
  assert.equal(saved.canUndo, true);
  assert.notEqual(saved.view.storage.state, before.view.storage.state);
  assert.notEqual(saved.view.revision, before.view.revision);
  assert.equal(fresh.lastCell, 'C6');
  assert.throws(() => validateUsedRange(sticky, { sheet: 'Sheet1', receipt: fresh.receipt }), {
    code: /^(formula-stale|layout-stale|note-stale)$/,
  });
  const manifest = validateRecoveryManifest(JSON.parse(await fs.readFile(f.manifestPath(), 'utf8')));
  assert.equal(manifest.head, before.state);
  assert.equal(manifest.savedHead, before.state);
  assert.equal(manifest.savedSources.length, 1);
  validateSavedSource(manifest.savedSources[0]);
  assert.equal(manifest.savedSources[0].revision, hash(await fs.readFile(f.source)));
  await f.owner.request({ operation: 'undo' });
  assert.equal(f.owner.status().state, far.state);
  assert.deepEqual(f.owner.status().view, far.view);
  assert.equal((await used(f.owner)).lastCell, 'Z100');
  await f.owner.request({ operation: 'redo' });
  assert.equal(f.owner.status().state, saved.state);
  assert.deepEqual(f.owner.status().view, saved.view);
  assert.equal((await used(f.owner)).lastCell, 'C6');
});

test('definite first-Save rejection leaves the original owner, source and private snapshot inventory intact', async (t) => {
  const f = await publicationFixture(t, { untitled: true });
  await edit(f.owner, 'A100', 9);
  const before = f.owner.status(),
    scope = await used(f.owner),
    directory = path.dirname(f.manifestPath()),
    files = await fs.readdir(directory);
  const destination = path.join(f.directory, 'exists.xlsx');
  await fs.copyFile(f.source, destination);
  const original = hash(await fs.readFile(destination));
  await assert.rejects(() => f.owner.request({ operation: 'saveAs', destination, expectedRevision: null }), {
    code: 'EEXIST',
  });
  assert.deepEqual(f.owner.status(), before);
  assert.deepEqual(await used(f.owner), scope);
  assert.equal(hash(await fs.readFile(destination)), original);
  assert.deepEqual(await fs.readdir(directory), files);
});

test('first Save refuses changed staged bytes before writing an intent or destination', async (t) => {
  const f = await publicationFixture(t, { untitled: true });
  await edit(f.owner, 'A1', 8);
  const before = f.owner.status(),
    directory = path.dirname(f.manifestPath()),
    inventory = await fs.readdir(directory),
    destination = path.join(f.directory, 'first.xlsx');
  let altered = false;
  f.setCall(async (input) => {
    if (input.operation === 'index' && path.basename(input.root).startsWith('publication-')) {
      const exported = f.calls.findLast((call) => call.operation === 'export');
      assert(exported, 'Actual export must precede preparation');
      await fs.writeFile(
        path.join(exported.root, 'saved.xlsx'),
        'changed output after immutable checkpoint preparation'
      );
      altered = true;
    }
  });
  await assert.rejects(() => f.owner.request({ operation: 'saveAs', destination, expectedRevision: null }), {
    code: 'session-stale',
  });
  f.setCall(undefined);
  assert.equal(altered, true);
  assert.deepEqual(f.owner.status(), before);
  await assert.rejects(() => fs.lstat(destination), { code: 'ENOENT' });
  assert.deepEqual(await fs.readdir(directory), inventory);
  assert.equal(JSON.parse(await fs.readFile(f.manifestPath(), 'utf8')).pendingPublication, undefined);
});

test('new saved-source intents require staged identity while historical hash-only intents remain readable', async (t) => {
  const f = await publicationFixture(t);
  await farClear(f.owner);
  await f.owner.request({ operation: 'save' });
  const manifest = JSON.parse(await fs.readFile(f.manifestPath(), 'utf8')),
    savedSource = manifest.savedSources[0],
    historical = { ...manifest, pendingPublication: { head: manifest.head, revision: savedSource.revision } };
  assert.equal(validateRecoveryManifest(historical), historical);
  assert.throws(
    () =>
      validateRecoveryManifest({ ...manifest, pendingPublication: { ...historical.pendingPublication, savedSource } }),
    { code: /^(session-bad-input|session-malformed)$/ }
  );
});

test('Save at a retained non-tip preserves warm Undo/Redo frames and cold exact saved floor/redo data', async (t) => {
  const f = await publicationFixture(t),
    previous = await farClear(f.owner),
    savedHead = f.owner.status().state;
  await edit(f.owner, 'B3', 42);
  const tip = f.owner.status().state,
    tipView = f.owner.status().view,
    tipScope = await used(f.owner);
  await f.owner.request({ operation: 'undo' });
  const history = JSON.parse(await fs.readFile(f.manifestPath(), 'utf8')).history;
  await f.owner.request({ operation: 'save' });
  const savedView = f.owner.status().view;
  assert.equal(f.owner.status().canRedo, true);
  assert.deepEqual(JSON.parse(await fs.readFile(f.manifestPath(), 'utf8')).history, history);
  await f.owner.request({ operation: 'undo' });
  assert.equal(f.owner.status().state, previous.state);
  assert.deepEqual(f.owner.status().view, previous.view);
  await f.owner.request({ operation: 'redo' });
  assert.equal(f.owner.status().state, savedHead);
  assert.deepEqual(f.owner.status().view, savedView);
  await f.owner.request({ operation: 'redo' });
  assert.equal(f.owner.status().state, tip);
  assert.deepEqual(f.owner.status().view, tipView);
  assert.deepEqual(await used(f.owner), tipScope);
  await f.owner.request({ operation: 'undo' });
  await f.owner.close();
  const cold = f.make();
  await cold.open(f.source);
  assert.equal(cold.status().state, savedHead);
  assert.equal((await used(cold)).lastCell, 'C6');
  assert.equal(cold.status().canRedo, true);
  await cold.request({ operation: 'redo' });
  assert.equal(cold.status().state, tip);
  const read = await cold.request({ operation: 'read', state: tip, sheet: 'Sheet1', range: 'B3' });
  assert.equal(read.cells.B3.value, 42);
});

test('a saved boundary is replayed before further edits after close/recovery', async (t) => {
  const f = await publicationFixture(t);
  await farClear(f.owner);
  await f.owner.request({ operation: 'save' });
  const savedHead = f.owner.status().state;
  await edit(f.owner, 'D8', 21);
  const edited = f.owner.status().state;
  assert.equal((await used(f.owner)).lastCell, 'D8');
  await f.owner.close();
  const cold = f.make();
  await cold.open(f.source);
  assert.equal(cold.status().state, edited);
  assert.equal((await used(cold)).lastCell, 'D8');
  await cold.request({ operation: 'undo' });
  assert.equal(cold.status().state, savedHead);
  assert.equal((await used(cold)).lastCell, 'C6');
});

test('lost final receipt write keeps a durable intent while warm adoption and cold reconciliation use the published hash', async (t) => {
  const f = await publicationFixture(t);
  await farClear(f.owner);
  const head = f.owner.status().state,
    originalRename = fs.rename;
  let published = false;
  fs.rename = async (source, destination) => {
    if (destination === f.source) {
      const result = await originalRename(source, destination);
      published = true;
      return result;
    }
    if (published && destination === f.manifestPath()) {
      const error = Error('Injected final receipt ENOSPC');
      error.code = 'ENOSPC';
      throw error;
    }
    return originalRename(source, destination);
  };
  syncBuiltinESMExports();
  let saved;
  try {
    saved = await f.owner.request({ operation: 'save' });
  } finally {
    fs.rename = originalRename;
    syncBuiltinESMExports();
  }
  assert.equal(saved.state, head);
  assert.equal(saved.dirty, false);
  assert.match(saved.warning, /receipt could not be finalized/);
  assert.equal((await used(f.owner)).lastCell, 'C6');
  const pending = validateRecoveryManifest(JSON.parse(await fs.readFile(f.manifestPath(), 'utf8')));
  assert.equal(pending.pendingPublication.head, head);
  assert.equal(pending.pendingPublication.revision, hash(await fs.readFile(f.source)));
  assert(pending.pendingPublication.savedSource);
  const calculation = pending.pendingPublication.savedSource.calculation;
  assert(calculation);
  const calculationFile = path.join(path.dirname(f.manifestPath()), calculation.snapshot + '.calculation.sqlite');
  assert.equal(hash(await fs.readFile(calculationFile)), calculation.revision);
  assert.equal((await fs.stat(calculationFile)).size, calculation.bytes);
  assert.equal(pending.savedHead, null);
  await f.owner.close();
  const cold = f.make();
  await cold.resume(f.source);
  assert.equal(cold.status().dirty, false);
  assert.equal(cold.status().state, head);
  assert.equal((await used(cold)).lastCell, 'C6');
  assert.equal(JSON.parse(await fs.readFile(f.manifestPath(), 'utf8')).pendingPublication, undefined);
  await cold.request({ operation: 'undo' });
  assert.equal((await used(cold)).lastCell, 'Z100');
  await cold.request({ operation: 'redo' });
  assert.equal((await used(cold)).lastCell, 'C6');
  await cold.close();
  const ordinary = f.make();
  await ordinary.open(f.source);
  assert.equal(ordinary.status().state, 'base');
  assert.equal(ordinary.status().canUndo, false);
  assert.equal((await used(ordinary)).lastCell, 'C6');
});

test('two-phase Indexed tokens cannot adopt after a later write, compaction, close or a forged source hash', async (t) => {
  const f = await publicationFixture(t),
    owner = new IndexedWorkbook(
      f.native,
      () => assert.fail('No fallback writer'),
      () => assert.fail('No formulas')
    );
  t.after(() => owner.close());
  await owner.open(f.source, { storageDirectory: path.join(f.directory, 'token-indices') });
  const before = owner.view('0');
  await assert.rejects(() => owner.preparePublicationRebase('0', f.source, 'f'.repeat(64)), { code: 'internal-mismatch' });
  assert.deepEqual(owner.view('0'), before);
  await owner.export('0', { publication: true });
  const token = await owner.preparePublicationRebase('0', f.source, hash(await fs.readFile(f.source)));
  assert.deepEqual(owner.view('0'), before);
  await owner.apply('0', [{ sheet: 'Sheet1', cell: 'B3', value: 8 }]);
  assert.throws(() => owner.adoptPublicationRebase(token), /unavailable|retired/);
  const compacted = await owner.preparePublicationRebase('0', f.source, hash(await fs.readFile(f.source)));
  await owner.compact(['0']);
  assert.throws(() => owner.adoptPublicationRebase(compacted), /unavailable|retired/);
  const closing = await owner.preparePublicationRebase('0', f.source, hash(await fs.readFile(f.source)));
  await owner.close();
  assert.throws(() => owner.adoptPublicationRebase(closing), /cancelled|closed/);
});

test('cancelled preparation never publishes or changes the captured owner frame', async (t) => {
  const f = await publicationFixture(t);
  await farClear(f.owner);
  const before = f.owner.status(),
    original = hash(await fs.readFile(f.source));
  let cancelled = false;
  f.setCall((input) => {
    if (input.operation === 'index' && path.basename(input.root).startsWith('publication-')) cancelled = true;
  });
  await assert.rejects(
    () =>
      f.owner.request(
        { operation: 'save' },
        {
          checkCancelled: () => {
            if (cancelled) throw Error('Actual cancellation boundary');
          },
        }
      ),
    /cancellation/
  );
  f.setCall(undefined);
  assert.deepEqual(f.owner.status(), before);
  assert.equal(hash(await fs.readFile(f.source)), original);
  assert.equal((await used(f.owner)).lastCell, 'Z100');
});

test('cancellation after a durable intent cannot mistake an unchanged identical file for a completed Save', async (t) => {
  const f = await publicationFixture(t);
  await farClear(f.owner);
  const before = f.owner.status(),
    original = hash(await fs.readFile(f.source)),
    originalRename = fs.rename;
  let cancelled = false;
  fs.rename = async (source, destination) => {
    const result = await originalRename(source, destination);
    if (destination === f.manifestPath() && JSON.parse(await fs.readFile(destination, 'utf8')).pendingPublication)
      cancelled = true;
    return result;
  };
  syncBuiltinESMExports();
  try {
    await assert.rejects(
      () =>
        f.owner.request(
          { operation: 'save' },
          {
            checkCancelled: () => {
              if (cancelled) throw Error('Cancel before publishing identical bytes');
            },
          }
        ),
      /Cancel before publishing/
    );
  } finally {
    fs.rename = originalRename;
    syncBuiltinESMExports();
  }
  assert.deepEqual(f.owner.status(), before);
  assert.equal(hash(await fs.readFile(f.source)), original);
  assert.equal((await used(f.owner)).lastCell, 'Z100');
  assert.equal(JSON.parse(await fs.readFile(f.manifestPath(), 'utf8')).pendingPublication, undefined);
});

test('an unpublished identical-output intent is rejected by its staged identity during cold reconciliation', async (t) => {
  const f = await publicationFixture(t);
  await farClear(f.owner);
  const head = f.owner.status().state,
    original = hash(await fs.readFile(f.source)),
    originalRename = fs.rename;
  let pendingWritten = false,
    cancelled = false;
  fs.rename = async (source, destination) => {
    if (destination === f.manifestPath() && pendingWritten) {
      const error = Error('Injected lost rollback receipt');
      error.code = 'ENOSPC';
      throw error;
    }
    const result = await originalRename(source, destination);
    if (destination === f.manifestPath() && JSON.parse(await fs.readFile(destination, 'utf8')).pendingPublication) {
      pendingWritten = true;
      cancelled = true;
    }
    return result;
  };
  syncBuiltinESMExports();
  try {
    await assert.rejects(
      () =>
        f.owner.request(
          { operation: 'save' },
          {
            checkCancelled: () => {
              if (cancelled) throw Error('Cancelled staged identical Save');
            },
          }
        ),
      /Cancelled staged/
    );
  } finally {
    fs.rename = originalRename;
    syncBuiltinESMExports();
  }
  const pending = validateRecoveryManifest(JSON.parse(await fs.readFile(f.manifestPath(), 'utf8')));
  assert.equal(pending.pendingPublication.revision, original);
  assert(pending.pendingPublication.destinationIdentity);
  assert.equal((await used(f.owner)).lastCell, 'Z100');
  await f.owner.close();
  const cold = f.make();
  await cold.resume(f.source);
  assert.equal(cold.status().state, head);
  assert.equal(cold.status().dirty, true);
  assert.equal((await used(cold)).lastCell, 'Z100');
  assert.equal(hash(await fs.readFile(f.source)), original);
  assert.equal(JSON.parse(await fs.readFile(f.manifestPath(), 'utf8')).pendingPublication, undefined);
});

test('a lost rename reply adopts the actually published staged identity even when bytes equal the prior source', async (t) => {
  const f = await publicationFixture(t);
  await farClear(f.owner);
  const original = hash(await fs.readFile(f.source)),
    originalRename = fs.rename;
  let injected = false;
  fs.rename = async (source, destination) => {
    const result = await originalRename(source, destination);
    if (destination === f.source && !injected) {
      injected = true;
      const error = Error('Lost rename reply after real publication');
      error.code = 'EIO';
      throw error;
    }
    return result;
  };
  syncBuiltinESMExports();
  let saved;
  try {
    saved = await f.owner.request({ operation: 'save' });
  } finally {
    fs.rename = originalRename;
    syncBuiltinESMExports();
  }
  assert.equal(hash(await fs.readFile(f.source)), original);
  assert.equal(saved.dirty, false);
  assert.match(saved.warning, /saved.*publication reported an error/);
  assert.equal((await used(f.owner)).lastCell, 'C6');
  assert.equal(JSON.parse(await fs.readFile(f.manifestPath(), 'utf8')).pendingPublication, undefined);
});

test('first Save preserves its UUID and adopts a confirmed lost-link reply before later cold recovery', async (t) => {
  const f = await publicationFixture(t, { untitled: true });
  await edit(f.owner, 'A1', 8);
  const before = f.owner.status(),
    destination = path.join(f.directory, 'first.xlsx'),
    originalLink = fs.link;
  let injected = false;
  fs.link = async (source, target) => {
    const result = await originalLink(source, target);
    if (target === destination && !injected) {
      injected = true;
      const error = Error('Lost link reply after real first Save');
      error.code = 'EIO';
      throw error;
    }
    return result;
  };
  syncBuiltinESMExports();
  let saved;
  try {
    saved = await f.owner.request({ operation: 'saveAs', destination, expectedRevision: null });
  } finally {
    fs.link = originalLink;
    syncBuiltinESMExports();
  }
  assert.equal(saved.documentId, before.documentId);
  assert.equal(saved.state, before.state);
  assert.equal(saved.unsaved, false);
  assert.equal(saved.dirty, false);
  assert.notEqual(saved.view.storage.state, before.view.storage.state);
  assert.match(saved.warning, /saved.*publication reported an error/);
  assert.equal((await used(f.owner)).lastCell, 'A1');
  assert.equal((await f.owner.request({ operation: 'binding' })).filename, destination);
  await edit(f.owner, 'B3', 21);
  const tip = f.owner.status().state;
  await f.owner.close();
  const cold = f.make();
  await cold.open(destination);
  assert.equal(cold.status().documentId, before.documentId);
  assert.equal(cold.status().state, tip);
  assert.equal((await used(cold)).lastCell, 'B3');
  assert.equal(
    (await cold.request({ operation: 'read', state: tip, sheet: 'Sheet1', range: 'A1:B3' })).cells.A1.value,
    8
  );
});

test('repeated Saves at the same public head reuse snapshots and reclaim superseded internal roots', async (t) => {
  const f = await publicationFixture(t);
  await edit(f.owner, 'A1', 8);
  await farClear(f.owner);
  await f.owner.request({ operation: 'save' });
  const head = f.owner.status().state,
    initialManifest = JSON.parse(await fs.readFile(f.manifestPath(), 'utf8'));
  async function inventory() {
    const [owner] = await fs.readdir(f.indices),
      entries = await fs.readdir(path.join(f.indices, owner), { withFileTypes: true });
    return {
      roots: entries.filter((entry) => entry.isDirectory()).map((entry) => entry.name),
      snapshots: (await fs.readdir(path.dirname(f.manifestPath()))).filter((name) => name.endsWith('.xlsx')),
    };
  }
  const before = await inventory();
  for (let i = 0; i < 6; i++) {
    const revision = f.owner.status().view.revision;
    await f.owner.request({ operation: 'save' });
    assert.equal(f.owner.status().state, head);
    assert.notEqual(f.owner.status().view.revision, revision);
    assert.equal((await used(f.owner)).lastCell, 'C6');
    const current = await inventory();
    assert.equal(current.roots.length, before.roots.length);
    assert.deepEqual(current.snapshots.sort(), before.snapshots.sort());
  }
  const final = validateRecoveryManifest(JSON.parse(await fs.readFile(f.manifestPath(), 'utf8')));
  assert.deepEqual(final.history, initialManifest.history);
  assert.deepEqual(final.savedSources, initialManifest.savedSources);
});

test('the nearest saved ancestor survives history-window pruning and restores the source floor', async (t) => {
  const f = await publicationFixture(t);
  await farClear(f.owner);
  await f.owner.request({ operation: 'save' });
  const savedHead = f.owner.status().state;
  for (let i = 0; i < 105; i++) await edit(f.owner, 'D8', i);
  const tip = f.owner.status().state;
  const manifest = validateRecoveryManifest(JSON.parse(await fs.readFile(f.manifestPath(), 'utf8')));
  assert.equal(manifest.history.frames.length, 101);
  assert.equal(
    manifest.history.frames.some((frame) => frame.head === savedHead),
    false
  );
  assert.equal(manifest.savedSources.length, 1);
  assert.equal(manifest.savedSources[0].head, savedHead);
  const calculation = manifest.savedSources[0].calculation;
  assert(calculation);
  const calculationFile = path.join(path.dirname(f.manifestPath()), calculation.snapshot + '.calculation.sqlite');
  assert.equal(hash(await fs.readFile(calculationFile)), calculation.revision);
  assert.equal((await fs.lstat(calculationFile)).nlink, 1);
  await f.owner.close();
  const callStart = f.calls.length,
    cold = f.make();
  await cold.open(f.source);
  assert.equal(cold.status().state, tip);
  assert.equal((await used(cold)).lastCell, 'D8');
  const read = await cold.request({ operation: 'read', state: tip, sheet: 'Sheet1', range: 'D8' });
  assert.equal(read.cells.D8.value, 104);
  assert(
    f.calls
      .slice(callStart)
      .some(
        (call) =>
          call.operation === 'calculation-transfer' &&
          call.purpose === 'publication-restore' &&
          call.witness === calculationFile
      )
  );
  assert.equal(hash(await fs.readFile(calculationFile)), calculation.revision);
});

test('checkpoint budget and hash violations reject before publication or recovery adoption', async (t) => {
  const limited = await publicationFixture(t, { maxRecoveryBytes: 1024 * 1024 });
  await edit(limited.owner, 'A1', 8);
  await fs.writeFile(
    path.join(path.dirname(limited.manifestPath()), 'retained-test.filler'),
    Buffer.alloc(1024 * 1024)
  );
  const before = limited.owner.status(),
    original = hash(await fs.readFile(limited.source));
  await assert.rejects(() => limited.owner.request({ operation: 'save' }), { code: 'session-limit' });
  assert.deepEqual(limited.owner.status(), before);
  assert.equal(hash(await fs.readFile(limited.source)), original);
  const f = await publicationFixture(t);
  await edit(f.owner, 'A1', 8);
  await farClear(f.owner);
  await f.owner.request({ operation: 'save' });
  await edit(f.owner, 'D8', 21);
  const manifest = JSON.parse(await fs.readFile(f.manifestPath(), 'utf8')),
    checkpoint = manifest.savedSources[0],
    snapshot = path.join(path.dirname(f.manifestPath()), checkpoint.snapshot + '.xlsx');
  assert.notEqual(checkpoint.snapshot, manifest.snapshot);
  await f.owner.close();
  await fs.writeFile(snapshot, 'damaged saved-source payload');
  const saved = hash(await fs.readFile(f.source)),
    cold = f.make();
  await assert.rejects(() => cold.open(f.source), { code: 'session-malformed' });
  assert.equal(hash(await fs.readFile(f.source)), saved);
  assert.equal(await fs.readFile(snapshot, 'utf8'), 'damaged saved-source payload');
});
