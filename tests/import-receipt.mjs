import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { randomUUID } from 'node:crypto';
import { findImportReceipt } from '../adapters/import-receipt.mjs';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
const descriptor = () => ({
  type: 'importCells',
  id: randomUUID(),
  sha256: 'a'.repeat(64),
  count: 2,
  bytes: 20,
  shape: { sheet: 'Sheet1', start: 'A1', rows: 1, columns: 2 },
  policy: 'text',
});
test('receipt proves the whole descriptor identity and exact parent; absent history is not rejection', () => {
  const expected = descriptor(),
    request = { operation: 'importReceipt', descriptor: expected, parentState: 'base' },
    context = { baseState: 'base', currentState: 'base', tipState: 'accepted' },
    nodes = [{ head: 'accepted', parent: null, commands: [expected] }];
  assert.equal(findImportReceipt(nodes, request, context).found, true);
  assert.equal(
    findImportReceipt(
      nodes,
      { ...request, descriptor: { ...expected, shape: { columns: 2, rows: 1, start: 'A1', sheet: 'Sheet1' } } },
      context
    ).found,
    true
  );
  for (const patch of [{ sha256: 'b'.repeat(64) }, { count: 1 }, { shape: { ...expected.shape, start: 'B1' } }])
    assert.throws(
      () => findImportReceipt(nodes, { ...request, descriptor: { ...expected, ...patch } }, context),
      /identity|hash|parent/
    );
  assert.throws(() => findImportReceipt(nodes, { ...request, parentState: 'changed' }, context), /parent/);
  assert.throws(() => findImportReceipt(nodes, { ...request, destination: '/tmp/private' }, context), /Invalid/);
  assert.deepEqual(findImportReceipt([], request, context), {
    found: false,
    reason: 'not-retained',
    currentState: 'base',
    tipState: 'accepted',
  });
});
test('real accepted import receipt survives Undo/Redo/recovery and corrupt node throws rather than rejecting acceptance', async (t) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'import-receipt-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const recoveryDirectory = path.join(directory, 'recovery');
  let owner = new SpreadsheetDocument({ recoveryDirectory, cacheIdentity: null });
  t.after(() => owner.close());
  let status = await owner.create('Receipt'),
    base = status.state,
    documentId = status.documentId;
  const bound = () => ({ documentId, state: status.state });
  let upload = await owner.request({
    operation: 'importBegin',
    ...bound(),
    shape: { sheet: 'Sheet1', start: 'A1', rows: 1, columns: 2 },
    count: 2,
    policy: 'text',
  });
  await owner.request({
    operation: 'importUpload',
    ...bound(),
    uploadId: upload.id,
    sequence: 0,
    cells: [
      [0, '01'],
      [1, '=SUM(A1)'],
    ],
  });
  const descriptor = await owner.request({ operation: 'importSeal', ...bound(), uploadId: upload.id });
  status = await owner.request({ operation: 'apply', state: status.state, commands: [descriptor] });
  const accepted = status.state,
    receipt = () => owner.request({ operation: 'importReceipt', ...bound(), descriptor, parentState: base });
  assert.equal((await receipt()).head, accepted);
  status = await owner.request({ operation: 'undo', state: status.state });
  assert.equal(status.state, base);
  assert.equal((await receipt()).found, true);
  assert.equal((await receipt()).tipState, accepted);
  await owner.close();
  owner = new SpreadsheetDocument({ recoveryDirectory, cacheIdentity: null });
  status = await owner.recover(documentId);
  assert.equal((await receipt()).head, accepted);
  status = await owner.request({ operation: 'redo', state: status.state });
  assert.equal(status.state, accepted);
  const filename = path.join(recoveryDirectory, 'draft-' + documentId, accepted + '.json'),
    before = await fs.readFile(filename),
    changed = Buffer.from(before);
  changed[changed.length - 2] = changed[changed.length - 2] === 32 ? 33 : 32;
  await fs.writeFile(filename, changed);
  await assert.rejects(receipt(), /damaged|Invalid|changed/);
  assert.deepEqual(await fs.readFile(filename), changed);
});
