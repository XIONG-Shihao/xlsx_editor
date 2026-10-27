// Two readers say what a workbook contains: the engine (`workbook-view`, used by
// indexed workbooks) and core/workbook-package.mjs (the in-memory path). Until
// the second is retired they must agree, on the cell-free copy the editor keeps.
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {
  readWorkbookMetadata,
  viewFromEngine,
  structuralEditReasons,
  structuralReasonsFromEngine,
} from '../core/workbook-package.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';

const root = await fs.mkdtemp(path.join(os.tmpdir(), 'workbook-view-parity-'));
after(async () => {
  closeNativeWorkbookWorker();
  await fs.rm(root, { recursive: true, force: true });
});

async function workbooks(directory) {
  const out = [];
  for (const entry of await fs.readdir(directory, { withFileTypes: true })) {
    const file = path.join(directory, entry.name);
    if (entry.isDirectory()) out.push(...(await workbooks(file)));
    else if (/\.xls[xm]$/i.test(entry.name)) out.push(file);
  }
  return out.sort();
}
const options = { objects: true, scoped: true };
// Compared as the screen receives them: a field set to undefined and a missing one are the same.
const sent = (value) => JSON.parse(JSON.stringify(value));

test('the engine and the in-memory reader give the same view of every fixture', async () => {
  const files = await workbooks(path.resolve('fixtures'));
  assert(files.length > 20, 'the fixtures were found');
  for (const file of files) {
    const dir = await fs.mkdtemp(path.join(root, 'book-'));
    await native({ action: 'workbook', operation: 'copy-source', root: dir, source: file });
    await native({ action: 'workbook', operation: 'index', root: dir });
    const bytes = await fs.readFile(path.join(dir, 'metadata.xlsx'));
    const answer = await native({ action: 'workbook', operation: 'workbook-view', root: dir });
    const js = readWorkbookMetadata(bytes),
      engine = viewFromEngine(answer, 0);
    js.view.revision = 0;
    const name = path.relative('fixtures', file);
    assert.deepEqual(sent(engine.view), sent(js.view), name);
    assert.deepEqual(sent(engine.styles), sent(js.styles), name);
    assert.deepEqual(
      Object.fromEntries(structuralReasonsFromEngine(answer, options)),
      Object.fromEntries(structuralEditReasons(bytes, options)),
      name
    );
  }
});
