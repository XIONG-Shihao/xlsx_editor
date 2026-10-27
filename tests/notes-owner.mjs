/** Actual native/Document owned note CRUD, history/recovery, saved XML and membership. No UI completion claim. */
import childProcess from 'node:child_process';
import { syncBuiltinESMExports } from 'node:module';
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { unzipSync, strFromU8, zipSync, strToU8 } from 'fflate';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { closeNativeWorkbookWorker, nativeWorkbookBinary } from '../core/native-runtime.mjs';
import { pageThemePrintFixture } from './support/page-theme-print-fixture.mjs';
import { noteMembershipFixture } from './support/note-membership-fixture.mjs';
import { validateNotes } from '../core/notes.mjs';
const roots = [],
  helper = nativeWorkbookBinary(),
  hash = (b) => createHash('sha256').update(b).digest('hex');
const helperHash = hash(await fs.readFile(helper));
const children = [],
  realSpawn = childProcess.spawn;
childProcess.spawn = (binary, ...args) => {
  const child = realSpawn(binary, ...args);
  if (binary === helper)
    children.push({
      pid: child.pid,
      exit: new Promise((resolve) => child.once('exit', (code, signal) => resolve({ code, signal }))),
    });
  return child;
};
syncBuiltinESMExports();
after(async () => {
  await closeNativeWorkbookWorker();
  childProcess.spawn = realSpawn;
  syncBuiltinESMExports();
  const lifecycle = [];
  for (const child of children) {
    let timer;
    const ended = await Promise.race([
      child.exit,
      new Promise((_, reject) => (timer = setTimeout(() => reject(Error('Native note helper did not exit')), 5000))),
    ]).finally(() => clearTimeout(timer));
    assert.throws(
      () => process.kill(child.pid, 0),
      (error) => error.code === 'ESRCH'
    );
    lifecycle.push({ pid: child.pid, ...ended, confirmedAbsent: true });
  }
  assert.equal(children.length, 1, 'One owned framed helper for this backend qualification');
  for (const root of roots) await fs.rm(root, { recursive: true, force: true });
});
async function fixture(t, { foreign = false } = {}) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'notes-owner-'));
  roots.push(root);
  const bytes = foreign ? noteMembershipFixture().bytes : pageThemePrintFixture({ theme: null }),
    filename = path.join(root, 'source.xlsx'),
    recoveryDirectory = path.join(root, 'recovery');
  await fs.writeFile(filename, bytes);
  const owners = [];
  const open = async (file = filename) => {
    const o = new SpreadsheetDocument({ recoveryDirectory, cacheIdentity: null });
    owners.push(o);
    await o.open(file, { fresh: true });
    return o;
  };
  const recover = async (id) => {
    const o = new SpreadsheetDocument({ recoveryDirectory, cacheIdentity: null });
    owners.push(o);
    await o.recover(id);
    return o;
  };
  t.after(async () => {
    for (const o of owners) await o.close();
  });
  return { root, filename, bytes, open, recover };
}
const request = (o, operation, rest = {}) => o.request({ operation, state: o.status().state, ...rest }),
  apply = (o, c) => request(o, 'apply', { commands: Array.isArray(c) ? c : [c] }),
  receipt = (o) => {
    const s = o.status();
    return { documentId: s.documentId, ownerId: s.view.storage.sessionId, state: s.state, revision: s.view.revision };
  };
async function inventory(o) {
  const r = (await request(o, 'notes', { documentId: o.status().documentId, sheet: 'Sheet1' })).notes;
  validateNotes(r, receipt(o), { sheet: 'Sheet1' });
  return r;
}
async function command(o, cell, body, author = 'Alice') {
  const v = await inventory(o);
  return {
    type: body === null ? 'removeNote' : 'setNote',
    sheet: 'Sheet1',
    cell,
    expectedNoteRevision: v.noteRevision,
    ...(body === null ? {} : { note: { author, text: body } }),
  };
}
const facts = async (o) => {
  const r = (await request(o, 'read', { sheet: 'Sheet1', range: 'A1:F1' })).cells;
  for (const v of Object.values(r)) {
    delete v.revision;
    if (v.calculation) delete v.calculation.revision;
    if (v.formula) {
      if (v.calculation?.status === 'calculated' && !v.calculation.pending) v.value = v.calculation.value;
      delete v.rawValue;
      delete v.cachedValue;
    }
  }
  return r;
};
test('exact older owned notes remain readable and a captured edit migrates to portable comments without losing text or Undo', async (t) => {
  const f = await fixture(t),
    o = await f.open();
  await apply(o, await command(o, 'A1', 'Legacy body', 'Alice'));
  const modern = path.join(f.root, 'modern.xlsx');
  await request(o, 'copy', { destination: modern, expectedRevision: null });
  await o.close();
  const parts = unzipSync(await fs.readFile(modern)),
    part = Object.keys(parts).find((name) => name.startsWith('xl/comments/') && name.endsWith('.xml'));
  assert(part);
  const xml = strFromU8(parts[part]);
  assert(!xml.includes('shapeId='));
  parts[part] = strToU8(xml.replace('authorId="0"', 'authorId="0" shapeId="1025"'));
  const legacy = path.join(f.root, 'legacy.xlsx'),
    bytes = Buffer.from(zipSync(parts));
  await fs.writeFile(legacy, bytes);
  const old = await f.open(legacy);
  let v = await inventory(old);
  assert(v.canEdit);
  assert.deepEqual(v.notes, [{ cell: 'A1', author: 'Alice', text: 'Legacy body' }]);
  await apply(old, await command(old, 'A1', 'Updated body', 'Alice'));
  await request(old, 'undo');
  assert.equal((await inventory(old)).notes[0].text, 'Legacy body');
  await request(old, 'redo');
  assert.equal((await inventory(old)).notes[0].text, 'Updated body');
  const output = path.join(f.root, 'migrated.xlsx');
  await request(old, 'copy', { destination: output, expectedRevision: null });
  assert(!strFromU8(unzipSync(await fs.readFile(output))[part]).includes('shapeId='));
  assert.deepEqual(await fs.readFile(legacy), bytes);
});
async function save(f, o, name, expected) {
  const saved = path.join(f.root, name + '.xlsx');
  await request(o, 'copy', { destination: saved, expectedRevision: null });
  const expectedFile = path.join(f.root, name + '-expected.json');
  await fs.writeFile(expectedFile, JSON.stringify(expected, null, 2));
  const proof = JSON.parse(
    execFileSync(
      process.env.XLSX_ORACLE_PYTHON ?? 'python3',
      [path.resolve('tests/support/check-notes.py'), f.filename, saved, expectedFile],
      { encoding: 'utf8' }
    )
  );
  assert(proof.passed);
  return { saved, proof };
}
test('actual notes attach to text/formula/off-extent cells, leave cells and source intact and survive Save/cold reopen', async (t) => {
  const f = await fixture(t),
    o = await f.open(),
    before = await facts(o),
    expected = [];
  for (const [cell, author, text] of [
    ['A1', 'Alice', '  Plan & <build>\r\n😀  '],
    ['B1', 'Bob', 'Computed total'],
    ['Z80', 'Carol', 'Off extent note'],
  ]) {
    const state = o.status().state;
    await apply(o, await command(o, cell, text, author));
    assert.notEqual(o.status().state, state);
    expected.push({ cell, author, text });
    assert.deepEqual(await facts(o), before);
  }
  const v = await inventory(o);
  assert.deepEqual(v.notes, expected);
  const membership = (await request(o, 'noteMembership', { documentId: o.status().documentId, sheet: 'Sheet1' }))
    .noteMembership;
  assert.equal(membership.status, 'complete', JSON.stringify(membership.reasons));
  assert.deepEqual(membership.notes, ['A1', 'B1', 'Z80']);
  assert.deepEqual(membership.threadedPlaceholders, []);
  const saved = await save(f, o, 'three-notes', expected);
  assert.equal(hash(await fs.readFile(f.filename)), hash(f.bytes));
  await o.close();
  const cold = await f.open(saved.saved);
  assert.deepEqual((await inventory(cold)).notes, expected);
  assert.equal(cold.status().canUndo, false);
});
test('edit author/body, Undo/Redo and durable recovery preserve siblings and exact note model', async (t) => {
  const f = await fixture(t),
    o = await f.open();
  await apply(o, await command(o, 'A1', 'Initial'));
  await apply(o, await command(o, 'C9', 'Sibling', 'Bob'));
  const old = (await inventory(o)).notes,
    before = o.status().state;
  await apply(o, await command(o, 'A1', 'Edited\nBody', 'Carol'));
  const accepted = o.status().state,
    current = (await inventory(o)).notes;
  assert.equal(current[0].author, 'Carol');
  assert.equal(current[1].text, 'Sibling');
  await request(o, 'undo');
  assert.equal(o.status().state, before);
  assert.deepEqual((await inventory(o)).notes, old);
  await request(o, 'redo');
  assert.deepEqual((await inventory(o)).notes, current);
  const id = o.status().documentId;
  await o.close();
  const recovered = await f.recover(id);
  assert.deepEqual((await inventory(recovered)).notes, current);
  await request(recovered, 'undo');
  assert.deepEqual((await inventory(recovered)).notes, old);
  await request(recovered, 'redo');
  assert.equal(recovered.status().state, accepted);
  const saved = await save(f, recovered, 'edited-notes', current);
});
test('no-change set/remove preserves head, Redo and journal; final deletion removes the owned notes/VML', async (t) => {
  const f = await fixture(t),
    o = await f.open();
  const empty = structuredClone(o.status());
  await apply(o, await command(o, 'A1', null));
  assert.deepEqual(o.status(), empty);
  await apply(o, await command(o, 'A1', 'First'));
  const first = o.status().state;
  await apply(o, await command(o, 'A1', 'Second'));
  await request(o, 'undo');
  const before = structuredClone(o.status()),
    manifest = path.join(f.root, 'recovery', 'draft-' + before.documentId, 'current.json'),
    bytes = await fs.readFile(manifest);
  await apply(o, await command(o, 'A1', 'First'));
  assert.deepEqual(o.status(), before);
  assert.equal(o.status().state, first);
  assert.deepEqual(await fs.readFile(manifest), bytes);
  await request(o, 'redo');
  await apply(o, await command(o, 'A1', null));
  assert.deepEqual((await inventory(o)).notes, []);
  const saved = await save(f, o, 'removed', []);
  await o.close();
  const cold = await f.open(saved.saved);
  assert.deepEqual((await inventory(cold)).notes, []);
});
test('Auto current and Manual pending formula truth survive note edits and subsequent ordinary cell writes', async (t) => {
  const f = await fixture(t),
    o = await f.open();
  const first = await facts(o);
  assert.equal(first.B1.calculation.value, 6);
  await apply(o, await command(o, 'B1', 'Auto result'));
  assert.deepEqual(await facts(o), first);
  await apply(o, { type: 'setCalculationOptions', mode: 'manual', recalculateBeforeSave: false });
  await apply(o, { sheet: 'Sheet1', cell: 'D1', value: 9 });
  const pending = await facts(o);
  assert(pending.B1.calculation.pending);
  await apply(o, await command(o, 'B1', 'Manual result'));
  assert.deepEqual(await facts(o), pending);
  await apply(o, { sheet: 'Sheet1', cell: 'A1', value: 'Ordinary cell edit' });
  assert.equal((await facts(o)).A1.value, 'Ordinary cell edit');
  const saved = path.join(f.root, 'manual.xlsx');
  await request(o, 'copy', { destination: saved, expectedRevision: null });
  await o.close();
  const cold = await f.open(saved);
  assert((await facts(cold)).B1.calculation.pending);
  assert.equal((await inventory(cold)).notes[0].text, 'Manual result');
});
test('stale prior/owner/target and late mixed-batch failure refuse without changing accepted state; incoming data is detached', async (t) => {
  const f = await fixture(t),
    o = await f.open();
  const captured = await command(o, 'A1', 'Original');
  const running = apply(o, captured);
  captured.note.text = 'Caller mutation';
  await running;
  assert.equal((await inventory(o)).notes[0].text, 'Original');
  const before = structuredClone(o.status()),
    cells = await facts(o),
    cmd = await command(o, 'A1', 'Wrong');
  for (const c of [
    { ...cmd, expectedNoteRevision: null },
    { ...cmd, xml: 'caller' },
    { ...cmd, cell: 'a1' },
    { ...cmd, note: { author: 'tc=thread', text: 'body' } },
  ])
    await assert.rejects(apply(o, c));
  await assert.rejects(
    apply(o, [await command(o, 'A1', 'First batch edit'), { sheet: 'Missing', cell: 'A1', value: 'late' }])
  );
  assert.deepEqual(o.status(), before);
  assert.deepEqual(await facts(o), cells);
  const good = { operation: 'notes', documentId: before.documentId, state: before.state, sheet: 'Sheet1' };
  for (const patch of [{ documentId: 'other' }, { state: 'other' }, { sheet: 'Missing' }, { cell: 'A1' }])
    await assert.rejects(o.request({ ...good, ...patch }));
});
test('imported rich/threaded note graphs remain preserved without a writable-model grant', async (t) => {
  const f = await fixture(t, { foreign: true }),
    o = await f.open(),
    v = await inventory(o);
  assert.equal(v.canEdit, false);
  assert.equal(v.notes, null);
  assert.match(v.protectedReason, /thread|Imported/i);
  const before = structuredClone(o.status());
  await assert.rejects(
    apply(o, {
      type: 'setNote',
      sheet: 'Sheet1',
      cell: 'A1',
      expectedNoteRevision: null,
      note: { author: 'Author', text: 'New' },
    })
  );
  assert.deepEqual(o.status(), before);
  const saved = path.join(f.root, 'foreign-copy.xlsx');
  await request(o, 'copy', { destination: saved, expectedRevision: null });
  const old = unzipSync(f.bytes),
    copy = unzipSync(await fs.readFile(saved));
  for (const name of Object.keys(old).filter(
    (n) => n.includes('comments') || n.includes('Comments') || n.includes('/persons/') || n.endsWith('.vml')
  ))
    assert.deepEqual(copy[name], old[name]);
});

test('protected owned notes have readable captured text and location while Document writes refuse', async (t) => {
  const f = await fixture(t),
    o = await f.open();
  await apply(o, await command(o, 'A1', 'Readonly note', 'Author'));
  const saved = await save(f, o, 'before-protection', [{ cell: 'A1', author: 'Author', text: 'Readonly note' }]);
  await o.close();
  const parts = unzipSync(await fs.readFile(saved.saved));
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    strFromU8(parts['xl/worksheets/sheet1.xml']).replace(
      '<legacyDrawing',
      '<sheetProtection sheet="1" objects="1"/><legacyDrawing'
    )
  );
  const filename = path.join(f.root, 'protected.xlsx');
  await fs.writeFile(filename, zipSync(parts));
  const readonly = await f.open(filename);
  assert.equal(readonly.status().view.protectedSheets.length, 1);
  const raw = await inventory(readonly);
  assert.equal(raw.canEdit, false);
  assert.equal(raw.notes[0].text, 'Readonly note');
  assert.equal(raw.notes[0].cell, 'A1');
  const before = structuredClone(readonly.status());
  await assert.rejects(
    apply(readonly, { type: 'removeNote', sheet: 'Sheet1', cell: 'A1', expectedNoteRevision: raw.noteRevision })
  );
  assert.deepEqual(readonly.status(), before);
});
