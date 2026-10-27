import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, readFile, realpath, writeFile, rm, readdir, symlink } from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { randomUUID } from 'node:crypto';
import { SpreadsheetDocument, listSpreadsheetRecovery } from '../adapters/document-session.mjs';
import { SpreadsheetWorker } from '../adapters/worker-client.mjs';
import { revision } from '../adapters/storage.mjs';
import { workbook } from '../fixtures/standalone.mjs';
import { readWorkbook } from '../core/workbook-package.mjs';
import { closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());

async function fixture(t) {
  const root = await realpath(await mkdtemp(path.join(os.tmpdir(), 'xlsx-new-workbook-'))),
    recoveryDirectory = path.join(root, 'recovery'),
    owners = [];
  const document = () => {
    const owner = new SpreadsheetDocument({ recoveryDirectory });
    owners.push(owner);
    return owner;
  };
  t.after(async () => {
    for (const owner of owners) await owner.close();
    await rm(root, { recursive: true, force: true });
  });
  return { root, recoveryDirectory, document, owners };
}
const read = async (owner, range = 'A1') =>
  (
    await owner.request({
      operation: 'read',
      state: owner.status().state,
      sheet: owner.status().kind === 'csv' ? 'CSV' : 'Sheet1',
      range,
    })
  ).cells;
const edit = (owner, value) =>
  owner.request({
    operation: 'apply',
    state: owner.status().state,
    commands: [{ sheet: owner.status().kind === 'csv' ? 'CSV' : 'Sheet1', cell: 'A1', value }],
  });
const manifestPath = (f, id) => path.join(f.recoveryDirectory, 'draft-' + id, 'current.json');

test('New has no user filename; first Save binds the same owner and normal Save retains undo', async (t) => {
  const f = await fixture(t),
    owner = f.document(),
    blank = await owner.create('Untitled 1'),
    destination = path.join(f.root, 'first.xlsx'),
    copy = path.join(f.root, 'copy.xlsx');
  assert.equal(blank.unsaved, true);
  assert.equal(blank.dirty, false);
  assert.equal(blank.name, 'Untitled 1');
  assert.equal((await owner.request({ operation: 'binding' })).filename, null);
  assert(!JSON.stringify(blank).includes(f.root));
  assert.equal(Object.keys(await read(owner)).length, 0);
  const session = blank.view.storage.sessionId,
    id = blank.documentId;
  await edit(owner, 10);
  const accepted = owner.status().state;
  await owner.request({ operation: 'copy', destination: copy, expectedRevision: null });
  assert.equal(owner.status().unsaved, true);
  assert.equal(owner.status().state, accepted);
  assert.equal(owner.status().dirty, true);
  assert.equal(readWorkbook(await readFile(copy)).sheets[0].cells.A1.value, 10);
  await owner.request({ operation: 'saveAs', destination, expectedRevision: null });
  assert.equal(owner.status().unsaved, false);
  assert.equal(owner.status().name, 'first.xlsx');
  assert.equal(owner.status().documentId, id);
  assert.equal(owner.status().state, accepted);
  assert.equal(owner.status().view.storage.sessionId, session);
  assert.equal(owner.status().dirty, false);
  assert.equal(owner.status().canUndo, true);
  await owner.request({ operation: 'undo' });
  assert.equal((await read(owner)).A1?.value ?? null, null);
  assert.equal(owner.status().dirty, true);
  await owner.request({ operation: 'redo' });
  assert.equal((await read(owner)).A1.value, 10);
  assert.equal(owner.status().dirty, false);
  await edit(owner, 20);
  await owner.request({ operation: 'save' });
  assert.equal(readWorkbook(await readFile(destination)).sheets[0].cells.A1.value, 20);
  assert.equal(owner.status().view.storage.sessionId, session);
  await owner.request({ operation: 'undo' });
  assert.equal((await read(owner)).A1.value, 10);
  assert.equal(owner.status().dirty, true);
  await owner.request({ operation: 'redo' });
  assert.equal(owner.status().dirty, false);
});

test('prepared blank owners cannot mutate and cancelling them leaves no discoverable draft', async (t) => {
  const f = await fixture(t),
    owner = f.document(),
    status = await owner.create('Untitled 2', { prepared: true });
  assert.deepEqual((await listSpreadsheetRecovery(f.recoveryDirectory)).entries, []);
  await assert.rejects(edit(owner, 9), { code: 'session-bad-input' });
  await assert.rejects(
    owner.request({ operation: 'saveAs', destination: path.join(f.root, 'no.xlsx'), expectedRevision: null }),
    { code: 'session-bad-input' }
  );
  await owner.close();
  await assert.rejects(readFile(manifestPath(f, status.documentId)), { code: 'ENOENT' });
  const active = f.document();
  await active.create('Untitled 3', { prepared: true });
  await active.request({ operation: 'activate' });
  await edit(active, 9);
  const id = active.status().documentId;
  await active.close();
  assert.equal((await listSpreadsheetRecovery(f.recoveryDirectory)).entries[0].id, id);
});

test('accepted unsaved work is discovered and recovered by durable UUID without any public file', async (t) => {
  const f = await fixture(t),
    owner = f.document();
  await owner.create('Untitled recovery');
  await edit(owner, 42);
  const id = owner.status().documentId;
  await owner.close();
  assert.deepEqual((await readdir(f.root)).sort(), ['recovery']);
  const list = await listSpreadsheetRecovery(f.recoveryDirectory);
  assert.equal(list.entries.length, 1);
  assert.equal(list.entries[0].id, id);
  assert.equal(list.entries[0].unsaved, true);
  assert.equal(list.entries[0].displayPath, undefined);
  const recovered = f.document();
  await recovered.recover(id);
  assert.equal(recovered.status().recovered, true);
  assert.equal(recovered.status().unsaved, true);
  assert.equal((await read(recovered)).A1.value, 42);
  assert.equal(recovered.status().canUndo, true);
  await recovered.request({ operation: 'discard' });
  await recovered.close();
  assert.deepEqual((await listSpreadsheetRecovery(f.recoveryDirectory)).entries, []);
});

test('failed first Save refuses stale/existing targets and rolls back filename binding only', async (t) => {
  const f = await fixture(t),
    owner = f.document(),
    target = path.join(f.root, 'target.xlsx');
  await owner.create();
  await edit(owner, 7);
  const state = owner.status().state;
  await writeFile(target, workbook());
  const original = await readFile(target),
    expected = await revision(target);
  await assert.rejects(owner.request({ operation: 'saveAs', destination: target, expectedRevision: null }), /exist/i);
  assert.equal(owner.status().unsaved, true);
  assert.equal(owner.status().state, state);
  assert.equal((await read(owner)).A1.value, 7);
  assert.deepEqual(await readFile(target), original);
  await writeFile(target, 'changed externally');
  await assert.rejects(
    owner.request({ operation: 'saveAs', destination: target, expectedRevision: expected }),
    /changed on disk/
  );
  assert.equal((await owner.request({ operation: 'binding' })).filename, null);
  assert.equal((await read(owner)).A1.value, 7);
  const stored = JSON.parse(await readFile(manifestPath(f, owner.status().documentId)));
  assert.equal(stored.filename, null);
  assert.equal(stored.pendingPublication, undefined);
  await assert.rejects(
    owner.request({
      operation: 'saveAs',
      destination: path.join(f.recoveryDirectory, 'private.xlsx'),
      expectedRevision: null,
    }),
    { code: 'session-unsupported' }
  );
});

for (const published of [false, true])
  test(
    'first-save crash receipt ' +
      (published ? 'binds an exact published candidate' : 'keeps an unpublished candidate unsaved'),
    async (t) => {
      const f = await fixture(t),
        owner = f.document(),
        target = path.join(f.root, 'bound.xlsx'),
        candidate = path.join(f.root, 'candidate.xlsx');
      await owner.create('Crash draft');
      await edit(owner, 66);
      const id = owner.status().documentId,
        file = manifestPath(f, id),
        before = JSON.parse(await readFile(file));
      await owner.request({ operation: 'copy', destination: published ? target : candidate, expectedRevision: null });
      const hash = await revision(published ? target : candidate);
      await writeFile(
        file,
        JSON.stringify({
          ...before,
          pendingPublication: {
            kind: 'bind',
            destination: target,
            expectedRevision: null,
            head: before.head,
            revision: hash,
          },
        })
      );
      await owner.close();
      const recovered = f.document();
      await recovered.recover(id);
      assert.equal((await read(recovered)).A1.value, 66);
      assert.equal(recovered.status().unsaved, !published);
      assert.equal(recovered.status().dirty, !published);
      assert.equal(recovered.status().canUndo, true);
      assert.equal(JSON.parse(await readFile(file)).pendingPublication, undefined);
      assert.equal((await recovered.request({ operation: 'binding' })).filename, published ? target : null);
    }
  );

test('same-file private reload candidates do not share the dirty owner manifest, including CSV', async (t) => {
  const f = await fixture(t);
  for (const extension of ['xlsx', 'csv']) {
    const target = path.join(f.root, 'same.' + extension);
    await writeFile(target, extension === 'xlsx' ? workbook() : 'original\n');
    const previous = f.document();
    await previous.open(target);
    await edit(previous, extension === 'xlsx' ? 99 : 'discarded');
    const accepted = previous.status().state;
    const cancelled = f.document();
    await cancelled.open(target, { fresh: true, prepared: true });
    assert.notEqual(cancelled.status().documentId, previous.status().documentId);
    assert.equal((await read(cancelled)).A1.value, extension === 'xlsx' ? 10 : 'original');
    await cancelled.close();
    assert.equal(previous.status().state, accepted);
    assert.equal((await read(previous)).A1.value, extension === 'xlsx' ? 99 : 'discarded');
    await previous.request({ operation: 'discard' });
    const replacement = f.document();
    await replacement.open(target, { fresh: true, prepared: true });
    await replacement.request({ operation: 'activate' });
    await previous.close();
    assert.equal((await read(replacement)).A1.value, extension === 'xlsx' ? 10 : 'original');
    await edit(replacement, extension === 'xlsx' ? 33 : 'accepted');
    const id = replacement.status().documentId;
    await replacement.close();
    const restarted = f.document();
    await restarted.open(target);
    assert.equal(restarted.status().documentId, id);
    assert.equal((await read(restarted)).A1.value, extension === 'xlsx' ? 33 : 'accepted');
    await restarted.request({ operation: 'discard' });
    await restarted.close();
  }
});

test('discovery rejects mismatched identities and missing/symlink snapshots without deleting records', async (t) => {
  const f = await fixture(t),
    owner = f.document();
  await owner.create();
  await edit(owner, 5);
  const id = owner.status().documentId,
    file = manifestPath(f, id),
    before = await readFile(file);
  await owner.close();
  const manifest = JSON.parse(before);
  await writeFile(file, JSON.stringify({ ...manifest, documentId: randomUUID() }));
  let list = await listSpreadsheetRecovery(f.recoveryDirectory);
  assert.equal(list.entries.length, 0);
  assert.match(list.warnings[0], /preserved/);
  await assert.rejects(f.document().recover(id), /identity/);
  await writeFile(file, before);
  const snapshot = path.join(path.dirname(file), manifest.snapshot + '.xlsx');
  await rm(snapshot);
  list = await listSpreadsheetRecovery(f.recoveryDirectory);
  assert.equal(list.entries.length, 0);
  assert.match(list.warnings[0], /preserved/);
  const outside = path.join(f.root, 'outside.xlsx');
  await writeFile(outside, workbook());
  await symlink(outside, snapshot);
  list = await listSpreadsheetRecovery(f.recoveryDirectory);
  assert.equal(list.entries.length, 0);
  assert.match(list.warnings[0], /snapshot/);
  assert.deepEqual(await readFile(file), before);
});

test('new/recovered worker protocol keeps status path-free and normal cancellation retains accepted draft', async (t) => {
  const f = await fixture(t),
    worker = new SpreadsheetWorker({ newWorkbook: { name: 'Worker draft' }, recoveryDirectory: f.recoveryDirectory });
  f.owners.push(worker);
  const initial = await worker.open();
  assert.equal(initial.unsaved, true);
  assert(!JSON.stringify(initial).includes(f.root));
  const changed = await worker.request({
    operation: 'apply',
    state: initial.state,
    commands: [{ sheet: 'Sheet1', cell: 'A1', value: 123 }],
  });
  await worker.cancel();
  const restored = new SpreadsheetWorker({ recoveryId: changed.documentId, recoveryDirectory: f.recoveryDirectory });
  f.owners.push(restored);
  const status = await restored.open();
  assert.equal(status.recovered, true);
  assert.equal(
    (await restored.request({ operation: 'read', state: status.state, sheet: 'Sheet1', range: 'A1' })).cells.A1.value,
    123
  );
});

test('conflicting dirty filename maps refuse ordinary Open and preserve both independent journals', async (t) => {
  const f = await fixture(t),
    target = path.join(f.root, 'conflict.xlsx');
  await writeFile(target, workbook());
  const before = await readFile(target),
    a = f.document(),
    b = f.document();
  await a.open(target, { fresh: true });
  await edit(a, 11);
  await b.open(target, { fresh: true });
  await edit(b, 22);
  const first = a.status().documentId,
    second = b.status().documentId;
  await a.close();
  await b.close();
  const manifests = await Promise.all([first, second].map((id) => readFile(manifestPath(f, id))));
  const list = await listSpreadsheetRecovery(f.recoveryDirectory);
  assert.equal(list.entries.length, 2);
  assert.match(list.warnings[0], /same file/);
  await assert.rejects(f.document().open(target), { code: /^(package-refused|session-unsupported)$/ });
  assert.deepEqual(await readFile(target), before);
  assert.deepEqual(await Promise.all([first, second].map((id) => readFile(manifestPath(f, id)))), manifests);
  const selected = f.document();
  await selected.recover(first);
  assert.equal((await read(selected)).A1.value, 11);
  await selected.request({ operation: 'discard' });
  await selected.close();
  const remaining = f.document();
  await remaining.open(target);
  assert.equal((await read(remaining)).A1.value, 22);
});

test('first Save cannot replace a destination with an accepted legacy recovery journal', async (t) => {
  const f = await fixture(t),
    target = path.join(f.root, 'pending.xlsx');
  await writeFile(target, workbook());
  const existing = f.document();
  await existing.open(target);
  await edit(existing, 99);
  const next = f.document();
  await next.create();
  await edit(next, 7);
  const expected = await revision(target),
    before = await readFile(target);
  await assert.rejects(next.request({ operation: 'saveAs', destination: target, expectedRevision: expected }), {
    code: 'session-refused',
  });
  assert.equal(next.status().unsaved, true);
  assert.equal((await read(next)).A1.value, 7);
  assert.deepEqual(await readFile(target), before);
  assert.equal((await read(existing)).A1.value, 99);
});

test('a failed clean recovery does not delete its borrowed record or reconcile a damaged journal', async (t) => {
  const f = await fixture(t),
    worker = new SpreadsheetWorker({ newWorkbook: { name: 'Receipt record' }, recoveryDirectory: f.recoveryDirectory });
  f.owners.push(worker);
  const initial = await worker.open(),
    target = path.join(f.root, 'published.xlsx');
  const changed = await worker.request({
    operation: 'apply',
    state: initial.state,
    commands: [{ sheet: 'Sheet1', cell: 'A1', value: 7 }],
  });
  await worker.request({ operation: 'saveAs', destination: target, expectedRevision: null });
  await worker.cancel();
  const file = manifestPath(f, changed.documentId),
    manifest = JSON.parse(await readFile(file)),
    node = path.join(path.dirname(file), manifest.head + '.json'),
    originalNode = await readFile(node);
  const receipt = JSON.stringify({
    ...manifest,
    pendingPublication: { head: manifest.head, revision: await revision(target) },
  });
  await writeFile(file, receipt);
  await writeFile(node, '{}');
  const broken = f.document();
  await assert.rejects(broken.recover(changed.documentId), { code: /^(paste-malformed|session-malformed)$/ });
  await broken.close();
  assert.equal(await readFile(file, 'utf8'), receipt);
  assert.equal(await readFile(node, 'utf8'), '{}');
  assert.equal(readWorkbook(await readFile(target)).sheets[0].cells.A1.value, 7);
  await writeFile(node, originalNode);
  const repaired = f.document();
  await repaired.recover(changed.documentId);
  assert.equal(repaired.status().dirty, false);
  assert.equal((await read(repaired)).A1.value, 7);
  assert.equal(JSON.parse(await readFile(file)).pendingPublication, undefined);
});

test('discovery stops at its directory admission bound without deleting overflow data', async (t) => {
  const f = await fixture(t);
  await mkdir(f.recoveryDirectory);
  await Promise.all(Array.from({ length: 129 }, () => mkdir(path.join(f.recoveryDirectory, 'draft-' + randomUUID()))));
  await assert.rejects(listSpreadsheetRecovery(f.recoveryDirectory), { code: 'session-limit' });
  assert.equal((await readdir(f.recoveryDirectory)).length, 129);
});
