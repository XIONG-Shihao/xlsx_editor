import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, realpath, readFile, writeFile, rm } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import path from 'node:path';
import os from 'node:os';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { createWorkbook } from '../core/workbook-package.mjs';
import { closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
async function fixture(t) {
  const root = await realpath(await mkdtemp(path.join(os.tmpdir(), 'xlsx-view-identity-'))),
    owners = [];
  const owner = () => {
    const document = new SpreadsheetDocument({ recoveryDirectory: path.join(root, 'recovery') });
    owners.push(document);
    return document;
  };
  t.after(async () => {
    for (const document of owners) await document.close();
    await rm(root, { recursive: true, force: true });
  });
  return { root, owner };
}
test('first Save changes only the private view binding; grant identity stays intact and clean reopen retains the file key', async (t) => {
  const f = await fixture(t),
    document = f.owner(),
    before = await document.create('View draft'),
    destination = path.join(f.root, 'saved.xlsx');
  assert.equal(before.viewPreferenceId, before.documentId);
  await document.request({ operation: 'saveAs', destination, expectedRevision: null });
  const saved = document.status();
  assert.equal(saved.documentId, before.documentId);
  assert.equal(saved.state, before.state);
  assert.equal(saved.viewPreferenceId, hash(destination));
  const bytes = await readFile(destination);
  await document.close();
  const reopened = await f.owner().open(destination);
  assert.notEqual(reopened.documentId, before.documentId);
  assert.equal(reopened.viewPreferenceId, saved.viewPreferenceId);
  assert.deepEqual(await readFile(destination), bytes);
  assert(!JSON.stringify(reopened).includes(destination));
});
test('fresh reload and two byte-identical files have canonical private keys independent of temporary UUID grants', async (t) => {
  const f = await fixture(t),
    a = path.join(f.root, 'a.xlsx'),
    b = path.join(f.root, 'b.xlsx'),
    bytes = createWorkbook();
  await writeFile(a, bytes);
  await writeFile(b, bytes);
  const first = await f.owner().open(a),
    other = await f.owner().open(b),
    fresh = await f.owner().open(a, { fresh: true, prepared: true });
  assert.notEqual(first.viewPreferenceId, other.viewPreferenceId);
  assert.notEqual(first.documentId, fresh.documentId);
  assert.equal(first.viewPreferenceId, fresh.viewPreferenceId);
  assert.equal(first.state, fresh.state);
  assert.equal(hash(await readFile(a)), hash(bytes));
  assert.equal(hash(await readFile(b)), hash(bytes));
});
