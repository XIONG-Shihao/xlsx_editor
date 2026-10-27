/** Current editor identities, plus explicit retirement and foreign-object protection. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { unzipSync, strFromU8 } from 'fflate';
import { FORMAT, matchesFormat } from '../core/format.mjs';
import { createWorkbook, editWorkbook, readWorkbookMetadata } from '../core/workbook-package.mjs';
import { indexedObjectList } from '../core/indexed-object-authoring.mjs';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { smartartSampleModel } from '../core/smartart-model.mjs';
import { validatePasteDescriptor } from '../core/paste-protocol.mjs';
after(() => closeNativeWorkbookWorker());
const request = (owner, operation, args = {}) =>
  owner.request({ operation, documentId: owner.status().documentId, state: owner.status().state, ...args });
test('only current descriptor identities are admitted', () => {
  for (const [key, value] of Object.entries(FORMAT)) {
    assert.doesNotMatch(value, /rpf|rfp/i);
    if (key.startsWith('SCHEMA_')) {
      assert(matchesFormat(value, key));
      assert(!matchesFormat(value.replace('workbook', 'rpf'), key));
    }
  }
  const descriptor = {
    type: 'importCells',
    id: '33333333-3333-4333-8333-333333333333',
    schema: FORMAT.SCHEMA_STAGED_PASTE_V2,
    sha256: 'a'.repeat(64),
    bytes: 120,
    count: 1,
    shape: { sheet: 'Sheet1', start: 'A1', rows: 1, columns: 1 },
    policy: 'typed-text-v1',
    conversion: { schema: 1, policies: ['text'], date1904: false, keepHeader: false, outputRows: 1, outputColumns: 1 },
  };
  assert.equal(validatePasteDescriptor(descriptor).schema, descriptor.schema);
  assert.throws(() => validatePasteDescriptor({ ...descriptor, schema: 'rpf.xlsx.staged-paste.v2' }));
});
for (const kind of ['smartart', 'model3d'])
  test('retired ' + kind + ' metadata is protected as foreign content', async () => {
    const file = new URL('./fixtures/retired-namespace-' + kind + '.xlsx', import.meta.url),
      bytes = await fs.readFile(file);
    const view = readWorkbookMetadata(bytes).view;
    const result = indexedObjectList(bytes, view, 'Sheet1');
    assert(result.objects.length > 0);
    assert(result.objects.every((object) => object.canEdit === false));
    assert.deepEqual(await fs.readFile(file), bytes);
  });
test('new object metadata is neutral and ordinary user text survives save and reopen', async (t) => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'workbook-current-format-'));
  const owner = new SpreadsheetDocument({ recoveryDirectory: path.join(root, 'recovery'), cacheIdentity: null });
  t.after(async () => {
    await owner.close();
    await fs.rm(root, { recursive: true, force: true });
  });
  const file = path.join(root, 'source.xlsx'),
    text = 'User RPF text stays user text';
  await fs.writeFile(file, editWorkbook(createWorkbook(), [{ sheet: 'Sheet1', cell: 'A1', value: text }]));
  await owner.open(file, { fresh: true });
  await request(owner, 'apply', {
    commands: [
      {
        operation: 'createSmartArt',
        sheet: 'Sheet1',
        objectId: '11111111-1111-4111-8111-111111111111',
        expectedObjectRevision: null,
        smartart: smartartSampleModel('process'),
        anchor: { cell: 'E2' },
      },
    ],
  });
  const saved = path.join(root, 'saved.xlsx');
  await request(owner, 'copy', { destination: saved, expectedRevision: null });
  const parts = unzipSync(await fs.readFile(saved));
  assert(Object.keys(parts).some((name) => name.includes('workbook-smartart-')));
  for (const [name, bytes] of Object.entries(parts))
    if (name.endsWith('.xml') || name.endsWith('.rels')) {
      assert.doesNotMatch(name, /rpf|rfp/i);
      // User content is allowed; reserved identifiers must use the current owner.
      assert.doesNotMatch(strFromU8(bytes), /urn:rpf:|name="rpf:|rpfObject_|rpfDiagram/);
    }
  assert.equal(readWorkbookMetadata(await fs.readFile(saved)).view.sheets[0].cells.A1.value, text);
});
