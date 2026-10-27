/** Actual JS owner and filesystem checks with deliberately mocked native witness bytes.
 * These tests do not execute SQLite, a native helper, a formula evaluator or an app. */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import {
  validateSavedSource,
  validateRecoveryManifest,
  directoryBytes,
  RECOVERY_MANIFEST_LIMIT,
} from '../adapters/storage.mjs';
import { publicationFixture, farClear, edit, used, hash } from './support/used-range-publication-fixture.mjs';

const manifest = async (f) => validateRecoveryManifest(JSON.parse(await fs.readFile(f.manifestPath(), 'utf8')));
const witnessPath = (f, checkpoint) =>
  path.join(path.dirname(f.manifestPath()), checkpoint.calculation.snapshot + '.calculation.sqlite');
const witnesses = (f) =>
  fs
    .readdir(path.dirname(f.manifestPath()))
    .then((names) => names.filter((name) => name.endsWith('.calculation.sqlite')).sort());
async function witness(f, checkpoint) {
  const file = witnessPath(f, checkpoint),
    info = await fs.lstat(file),
    bytes = await fs.readFile(file);
  assert.equal(info.isFile(), true);
  assert.equal(info.isSymbolicLink(), false);
  assert.equal(info.nlink, 1);
  assert.equal(info.mode & 0o777, 0o600);
  assert.equal(info.size, checkpoint.calculation.bytes);
  assert.equal(hash(bytes), checkpoint.calculation.revision);
  return { file, bytes };
}

test('saved-source calculation metadata is optional for legacy records and closed/bounded when present', () => {
  const source = { schema: 1, head: null, headHash: null, snapshot: randomUUID(), revision: 'a'.repeat(64), bytes: 1 },
    calculation = { schema: 1, snapshot: randomUUID(), revision: 'b'.repeat(64), bytes: 1 };
  assert.equal(validateSavedSource(source), source);
  assert.deepEqual(validateSavedSource({ ...source, calculation }).calculation, calculation);
  assert.equal(
    validateSavedSource({ ...source, calculation: { ...calculation, bytes: 256 * 1024 * 1024 } }).calculation.bytes,
    256 * 1024 * 1024
  );
  for (const item of [
    undefined,
    null,
    false,
    [],
    {},
    'witness',
    ...Object.keys(calculation).map((key) => {
      const item = { ...calculation };
      delete item[key];
      return item;
    }),
    { ...calculation, path: '/must-not-be-stored' },
    { ...calculation, schema: 2 },
    { ...calculation, snapshot: 'not-a-uuid' },
    { ...calculation, revision: 'bad' },
    ...[0, -1, 1.5, Number.MAX_SAFE_INTEGER, 256 * 1024 * 1024 + 1].map((bytes) => ({ ...calculation, bytes })),
  ])
    assert.throws(() => validateSavedSource({ ...source, calculation: item }), /calculation checkpoint/);
});

test('new Save stores a private hash/size witness and repeated same-head saves reuse it while Copy stays pure', async (t) => {
  const f = await publicationFixture(t);
  await edit(f.owner, 'A1', 8);
  await farClear(f.owner);
  const before = f.owner.status();
  await f.owner.request({ operation: 'save' });
  const first = await manifest(f),
    checkpoint = first.savedSources[0];
  assert(checkpoint.calculation);
  const original = await witness(f, checkpoint);
  assert.equal(checkpoint.head, before.state);
  assert.equal(checkpoint.headHash, first.headHash);
  assert.equal(checkpoint.revision, hash(await fs.readFile(f.source)));
  assert.equal((await used(f.owner)).lastCell, 'C6');
  const copyState = f.owner.status(),
    beforeCopy = await fs.readFile(f.manifestPath()),
    copy = path.join(f.directory, 'copy.xlsx'),
    callStart = f.calls.length;
  await f.owner.request({ operation: 'copy', destination: copy, expectedRevision: null });
  assert.deepEqual(f.owner.status(), copyState);
  assert.deepEqual(await fs.readFile(f.manifestPath()), beforeCopy);
  assert.equal(
    f.calls.slice(callStart).some((call) => call.operation === 'calculation-transfer' || call.publication === true),
    false,
    'Copy does not authorize publication transfer'
  );
  for (let i = 0; i < 6; i++) {
    await f.owner.request({ operation: 'save' });
    assert.deepEqual((await manifest(f)).savedSources, first.savedSources);
    assert.deepEqual(await witnesses(f), [path.basename(original.file)]);
    assert.deepEqual((await witness(f, checkpoint)).bytes, original.bytes);
  }
});

test('cold recovery restores the exact saved witness before replaying later edits', async (t) => {
  const f = await publicationFixture(t);
  await farClear(f.owner);
  await f.owner.request({ operation: 'save' });
  const saved = await manifest(f),
    checkpoint = saved.savedSources[0],
    savedState = f.owner.status().state,
    pin = await witness(f, checkpoint);
  await edit(f.owner, 'D8', 21);
  const tip = f.owner.status().state;
  await f.owner.close();
  const callStart = f.calls.length,
    cold = f.make();
  await cold.open(f.source);
  assert.equal(cold.status().state, tip);
  assert.equal((await used(cold)).lastCell, 'D8');
  const restores = f.calls
    .slice(callStart)
    .filter((call) => call.operation === 'calculation-transfer' && call.purpose === 'publication-restore');
  assert.equal(restores.length, 1);
  assert.equal(restores[0].witness, pin.file);
  assert.equal(restores[0].witnessRevision, checkpoint.calculation.revision);
  assert.equal(restores[0].sourceRevision, checkpoint.revision);
  await cold.request({ operation: 'undo' });
  assert.equal(cold.status().state, savedState);
  assert.equal((await used(cold)).lastCell, 'C6');
  assert.deepEqual((await witness(f, checkpoint)).bytes, pin.bytes);
});

test('damaged, missing, symbolic or multiply linked witnesses refuse cold adoption and preserve source/recovery', async (t) => {
  for (const kind of ['hash', 'size', 'missing', 'symlink', 'hardlink']) {
    const f = await publicationFixture(t);
    await edit(f.owner, 'A1', 8);
    await farClear(f.owner);
    await f.owner.request({ operation: 'save' });
    await edit(f.owner, 'D8', 21);
    const checkpoint = (await manifest(f)).savedSources[0],
      pin = await witness(f, checkpoint);
    await f.owner.close();
    const sourceHash = hash(await fs.readFile(f.source)),
      receipt = await fs.readFile(f.manifestPath()),
      external = path.join(f.directory, 'external-witness.sqlite');
    if (kind === 'hash') {
      const changed = Buffer.from(pin.bytes);
      changed[0] ^= 1;
      await fs.writeFile(pin.file, changed);
    } else if (kind === 'size') await fs.truncate(pin.file, pin.bytes.length - 1);
    else if (kind === 'missing') await fs.unlink(pin.file);
    else if (kind === 'symlink') {
      await fs.writeFile(external, pin.bytes);
      await fs.unlink(pin.file);
      await fs.symlink(external, pin.file);
    } else await fs.link(pin.file, external);
    const beforeInfo = await fs.lstat(pin.file).catch((error) => {
        if (error.code === 'ENOENT') return null;
        throw error;
      }),
      beforeBytes = await fs.readFile(pin.file).catch((error) => {
        if (error.code === 'ENOENT') return null;
        throw error;
      }),
      callStart = f.calls.length,
      cold = f.make();
    await assert.rejects(
      () => cold.open(f.source),
      (error) => error.code === 'ENOENT' || /calculation checkpoint is damaged/.test(error.message)
    );
    assert.equal(hash(await fs.readFile(f.source)), sourceHash);
    assert.deepEqual(await fs.readFile(f.manifestPath()), receipt);
    assert.deepEqual(
      await fs.readFile(pin.file).catch((error) => {
        if (error.code === 'ENOENT') return null;
        throw error;
      }),
      beforeBytes
    );
    const afterInfo = await fs.lstat(pin.file).catch((error) => {
      if (error.code === 'ENOENT') return null;
      throw error;
    });
    assert.equal(afterInfo?.isSymbolicLink(), beforeInfo?.isSymbolicLink());
    assert.equal(afterInfo?.nlink, beforeInfo?.nlink);
    assert.equal(
      f.calls.slice(callStart).some((call) => call.operation === 'calculation-transfer'),
      false,
      'Invalid witness refuses before native restoration'
    );
  }
});

test('pruning a discarded redo branch removes its witness and retains the saved ancestor needed for recovery', async (t) => {
  const f = await publicationFixture(t);
  await edit(f.owner, 'A1', 8);
  await f.owner.request({ operation: 'save' });
  const ancestorState = f.owner.status().state,
    ancestor = (await manifest(f)).savedSources[0],
    kept = await witness(f, ancestor);
  await edit(f.owner, 'D8', 21);
  await f.owner.request({ operation: 'save' });
  const second = (await manifest(f)).savedSources.find((source) => source.head === f.owner.status().state),
    retired = await witness(f, second);
  assert.notEqual(kept.file, retired.file);
  assert.equal((await witnesses(f)).length, 2);
  await f.owner.request({ operation: 'undo' });
  assert.equal(f.owner.status().state, ancestorState);
  await edit(f.owner, 'E9', 22);
  const branch = f.owner.status().state,
    pruned = await manifest(f);
  assert.deepEqual(pruned.savedSources, [ancestor]);
  const beforeCleanup = f.owner.status();
  await f.owner.request({ operation: 'importPrune', documentId: beforeCleanup.documentId, state: branch });
  assert.deepEqual(f.owner.status(), beforeCleanup);
  assert.deepEqual(await witnesses(f), [path.basename(kept.file)]);
  await assert.rejects(() => fs.lstat(retired.file), { code: 'ENOENT' });
  assert.deepEqual((await witness(f, ancestor)).bytes, kept.bytes);
  await f.owner.close();
  const cold = f.make();
  await cold.open(f.source);
  assert.equal(cold.status().state, branch);
  assert.equal((await used(cold)).lastCell, 'E9');
  const read = await cold.request({ operation: 'read', state: branch, sheet: 'Sheet1', range: 'A1:E9' });
  assert.equal(read.cells.E9.value, 22);
  assert.equal(read.cells.D8, undefined);

  // Save must reclaim an already retired redo checkpoint before charging its
  // replacement. This uses the public Save path, with no explicit prune call.
  const budget = 1024 * 1024,
    witnessBytes = 300 * 1024,
    g = await publicationFixture(t, { maxRecoveryBytes: budget, witnessBytes });
  await edit(g.owner, 'A1', 8);
  await g.owner.request({ operation: 'save' });
  const savedAncestorState = g.owner.status().state,
    savedAncestor = (await manifest(g)).savedSources[0],
    ancestorPin = await witness(g, savedAncestor);
  await edit(g.owner, 'D8', 21);
  await g.owner.request({ operation: 'save' });
  const discarded = (await manifest(g)).savedSources.find((source) => source.head === g.owner.status().state),
    discardedPin = await witness(g, discarded),
    directory = path.dirname(g.manifestPath()),
    discardedWorkbook = path.join(directory, discarded.snapshot + '.xlsx');
  await g.owner.request({ operation: 'undo' });
  assert.equal(g.owner.status().state, savedAncestorState);
  await edit(g.owner, 'E9', 22);
  const savedBranchState = g.owner.status().state;
  assert.deepEqual((await manifest(g)).savedSources, [savedAncestor]);
  assert.equal(
    (await fs.stat(discardedWorkbook)).isFile(),
    true,
    'Retired redo payload remains until verified reclamation'
  );
  const copy = path.join(g.directory, 'branched-size-only-copy.xlsx');
  await g.owner.request({ operation: 'copy', destination: copy, expectedRevision: null });
  const copyBytes = await fs.readFile(copy),
    candidateBytes = copyBytes.length,
    currentBytes = await directoryBytes(directory),
    fillerBytes = budget - currentBytes - candidateBytes - witnessBytes - RECOVERY_MANIFEST_LIMIT + 4096;
  assert(fillerBytes > 0);
  const filler = path.join(directory, 'retained-near-budget.filler');
  await fs.writeFile(filler, Buffer.alloc(fillerBytes));
  const beforeBytes = await directoryBytes(directory),
    discardedBytes = (await fs.stat(discardedWorkbook)).size + discarded.calculation.bytes;
  assert(beforeBytes + candidateBytes + RECOVERY_MANIFEST_LIMIT < budget, 'Workbook fits before orphan reclamation');
  assert(
    beforeBytes + candidateBytes + witnessBytes + RECOVERY_MANIFEST_LIMIT > budget,
    'Retired witness would falsely exhaust Save quota'
  );
  assert(
    beforeBytes - discardedBytes + candidateBytes + witnessBytes + RECOVERY_MANIFEST_LIMIT < budget,
    'Referenced data and replacement fit after orphan reclamation'
  );
  await g.owner.request({ operation: 'save' });
  assert.equal(g.owner.status().state, savedBranchState, 'Reclamation preserves the accepted branch head');
  const published = await manifest(g),
    branchCheckpoint = published.savedSources.find((source) => source.head === savedBranchState);
  assert(branchCheckpoint);
  assert.deepEqual(
    published.savedSources.find((source) => source.head === savedAncestorState),
    savedAncestor
  );
  assert.equal(
    published.savedSources.some((source) => source.head === discarded.head),
    false
  );
  await assert.rejects(() => fs.lstat(discardedWorkbook), { code: 'ENOENT' });
  await assert.rejects(() => fs.lstat(discardedPin.file), { code: 'ENOENT' });
  assert.deepEqual((await witness(g, savedAncestor)).bytes, ancestorPin.bytes);
  await witness(g, branchCheckpoint);
  assert.equal(hash(await fs.readFile(g.source)), hash(copyBytes));
  assert.equal((await fs.stat(filler)).size, fillerBytes);
  assert((await directoryBytes(directory)) + RECOVERY_MANIFEST_LIMIT < budget);
  await g.owner.close();
  const resumed = g.make();
  await resumed.resume(g.source);
  assert.equal(resumed.status().state, savedBranchState);
  assert.equal(resumed.status().recovered, true);
  await resumed.request({ operation: 'undo' });
  assert.equal(resumed.status().state, savedAncestorState);
  const ancestorRead = await resumed.request({
    operation: 'read',
    state: savedAncestorState,
    sheet: 'Sheet1',
    range: 'A1:E9',
  });
  assert.equal(ancestorRead.cells.A1.value, 8);
  assert.equal(ancestorRead.cells.D8, undefined);
  assert.equal(ancestorRead.cells.E9, undefined);
});

test('combined workbook/witness recovery quota refuses before publication and cleans only rejected candidates', async (t) => {
  const budget = 1024 * 1024,
    witnessBytes = 768 * 1024,
    f = await publicationFixture(t, { maxRecoveryBytes: budget, witnessBytes });
  await edit(f.owner, 'A1', 8);
  const copy = path.join(f.directory, 'size-only-copy.xlsx');
  await f.owner.request({ operation: 'copy', destination: copy, expectedRevision: null });
  const directory = path.dirname(f.manifestPath()),
    sourceBytes = (await fs.stat(copy)).size,
    currentBytes = await directoryBytes(directory),
    fillerBytes = budget - currentBytes - sourceBytes - witnessBytes - RECOVERY_MANIFEST_LIMIT + 1;
  assert(fillerBytes > 0);
  await fs.writeFile(path.join(directory, 'retained-test.filler'), Buffer.alloc(fillerBytes));
  assert(
    (await directoryBytes(directory)) + sourceBytes + RECOVERY_MANIFEST_LIMIT < budget,
    'Workbook alone fits; witness pushes combined quota over budget'
  );
  const before = f.owner.status(),
    scope = await used(f.owner),
    sourceHash = hash(await fs.readFile(f.source)),
    receipt = await fs.readFile(f.manifestPath()),
    inventory = await fs.readdir(directory);
  await assert.rejects(() => f.owner.request({ operation: 'save' }), { code: 'session-limit' });
  assert.deepEqual(f.owner.status(), before);
  assert.deepEqual(await used(f.owner), scope);
  assert.equal(hash(await fs.readFile(f.source)), sourceHash);
  assert.deepEqual(await fs.readFile(f.manifestPath()), receipt);
  assert.deepEqual(await fs.readdir(directory), inventory);
  assert.deepEqual(await witnesses(f), []);
});

test('malformed native witness descriptors and hash/size mismatches cannot publish or adopt a new source frame', async (t) => {
  for (const change of [
    (reply) => ({ ...reply, extra: 'refuse' }),
    (reply) => ({ ...reply, witness: { ...reply.witness, file: '../source.xlsx' } }),
    (reply) => ({ ...reply, witness: { ...reply.witness, schema: 2 } }),
    (reply) => ({ ...reply, witness: { ...reply.witness, revision: 'f'.repeat(64) } }),
    (reply) => ({ ...reply, witness: { ...reply.witness, bytes: reply.witness.bytes + 1 } }),
    (reply) => ({ ...reply, witness: { ...reply.witness, bytes: 256 * 1024 * 1024 + 1 } }),
  ]) {
    const f = await publicationFixture(t);
    await farClear(f.owner);
    const before = f.owner.status(),
      scope = await used(f.owner),
      sourceHash = hash(await fs.readFile(f.source)),
      receipt = await fs.readFile(f.manifestPath()),
      inventory = await fs.readdir(path.dirname(f.manifestPath()));
    f.setPublicationReply(change);
    await assert.rejects(
      () => f.owner.request({ operation: 'save' }),
      { code: /^(internal-invalid-field|session-malformed)$/ }
    );
    f.setPublicationReply();
    assert.deepEqual(f.owner.status(), before);
    assert.deepEqual(await used(f.owner), scope);
    assert.equal(hash(await fs.readFile(f.source)), sourceHash);
    assert.deepEqual(await fs.readFile(f.manifestPath()), receipt);
    assert.deepEqual(await fs.readdir(path.dirname(f.manifestPath())), inventory);
    assert.deepEqual(await witnesses(f), []);
  }
});
