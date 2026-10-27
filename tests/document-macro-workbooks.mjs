import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { unzipSync, strFromU8 } from 'fflate';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { SpreadsheetWorker } from '../adapters/worker-client.mjs';
import { nativeWorkbookOperationAsync, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';

after(() => closeNativeWorkbookWorker());
const sha = (bytes) => createHash('sha256').update(bytes).digest('hex');
const MACRO = 'xl/vbaProject.bin';
// A workbook written by Excel with one module (Module1, MyFunc = 42).
const source = path.resolve('fixtures/remaining-formats/vba-user-function.xlsm');

async function fixture(t, { worker = false } = {}) {
  const root = await fs.realpath(await fs.mkdtemp(path.join(os.tmpdir(), 'document-macro-'))),
    filename = path.join(root, 'book.xlsm'),
    recoveryDirectory = path.join(root, 'recovery'),
    owners = [];
  await fs.copyFile(source, filename);
  const owner = worker
    ? new SpreadsheetWorker({ filename, recoveryDirectory })
    : new SpreadsheetDocument({
        recoveryDirectory,
        native: (input) => nativeWorkbookOperationAsync(input),
        cacheIdentity: null,
      });
  owners.push(owner);
  t.after(async () => {
    try {
      for (const item of owners) await item.close();
    } finally {
      await fs.rm(root, { recursive: true, force: true });
    }
  });
  return { root, filename, owner, bytes: await fs.readFile(source) };
}

for (const worker of [false, true])
  test(`${worker ? 'worker' : 'Document'}: an .xlsm file opens for editing, saves as .xlsm with the macro part byte for byte the same and still shows its macros`, async (t) => {
    const f = await fixture(t, { worker }),
      opened = worker ? await f.owner.open() : await f.owner.open(f.filename);
    assert.equal(opened.kind, 'xlsx');
    assert.equal(opened.view.readOnly, false);
    assert.equal(opened.view.macroProject, true);
    const edited = await f.owner.request({
      operation: 'apply',
      state: opened.state,
      commands: [{ sheet: 'Sheet1', cell: 'E5', value: 'edited' }],
    });
    assert.equal(edited.dirty, true);
    await f.owner.request({ operation: 'save' });
    const before = unzipSync(f.bytes),
      after = unzipSync(await fs.readFile(f.filename)),
      types = strFromU8(after['[Content_Types].xml']);
    assert.equal(sha(after[MACRO]), sha(before[MACRO]));
    assert(types.includes('macroEnabled.main+xml'), 'the saved file lost its macro-enabled type');
    assert(strFromU8(after['xl/_rels/workbook.xml.rels']).includes('relationships/vbaProject'));
    assert.equal(path.extname(f.filename), '.xlsm');
    // The macros can be read, never run, while the workbook is being edited.
    // The worker client answers status as a request; the document directly.
    const status = worker ? await f.owner.request({ operation: 'status' }) : f.owner.status(),
      bound = { state: status.state, documentId: status.documentId },
      modules = await f.owner.request({ operation: 'vbaModules', ...bound }),
      module = modules.modules.find((value) => value.name === 'Module1');
    assert((await f.owner.request({ operation: 'vbaModule', ...bound, moduleId: module.id })).code.includes('MyFunc = 42'));
    assert.equal(
      (
        await f.owner.request({
          operation: 'evaluateVba',
          ...bound,
          moduleId: module.id,
          functionName: 'MyFunc',
          args: [],
        })
      ).value,
      42
    );
    assert.equal(
      (await f.owner.request({ operation: 'vbaFunctions', ...bound, moduleId: module.id })).functions[0].name,
      'MyFunc'
    );
    for (const fields of [{ state: 'old' }, { documentId: 'other' }])
      await assert.rejects(f.owner.request({ operation: 'vbaModules', ...bound, ...fields }), /stale/);
    await assert.rejects(
      f.owner.request({ operation: 'vbaModule', ...bound, moduleId: module.id, path: f.filename }),
      /fields/
    );
  });

test('Document: a macro workbook cannot be saved as a copy with another extension', async (t) => {
  const f = await fixture(t),
    opened = await f.owner.open(f.filename);
  await assert.rejects(
    f.owner.request({ operation: 'copy', destination: path.join(f.root, 'copy.xlsx'), expectedRevision: null }),
    { code: 'session-bad-input' }
  );
  const copy = path.join(f.root, 'copy.xlsm');
  await f.owner.request({ operation: 'copy', destination: copy, expectedRevision: null });
  assert.equal(sha(unzipSync(await fs.readFile(copy))[MACRO]), sha(unzipSync(f.bytes)[MACRO]));
  assert.equal(opened.view.readOnly, false);
});
