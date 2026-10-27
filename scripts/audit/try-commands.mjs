import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { IndexedWorkbook } from '../../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../../core/native-runtime.mjs';
import { editCalculatedWorkbookResult } from '@office-editors/xlsx-editor/engine';
const [file, json, read] = process.argv.slice(2);
process.on('unhandledRejection', (e) => console.log('UNH', e));
const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'try-'));
const copy = path.join(dir, 's.xlsx');
await fs.copyFile(file, copy);
const book = new IndexedWorkbook(
  native,
  async (f, e, o) => {
    const r = await editCalculatedWorkbookResult(await fs.readFile(f), e);
    await fs.writeFile(o, r.bytes);
    return { revision: r.view.revision, edits: r.edits };
  },
  (i) => calculateIndexedWorkbook(i, native)
);
const view = await book.open(copy, { storageDirectory: path.join(dir, 'i') });
console.log(
  'sheet',
  JSON.stringify({
    name: view.sheets[0].name,
    rows: view.sheets[0].rows,
    columns: view.sheets[0].columns,
    filter: view.sheets[0].filter,
  })
);
try {
  let st = '0';
  for (const c of JSON.parse(json)) {
    const r = await book.apply(st, [c]);
    st = r.view.storage.state;
  }
  if (read) {
    const cells = (await book.read(st, view.sheets[0].name, read)).cells;
    for (const [k, v] of Object.entries(cells))
      console.log(k, JSON.stringify(v.formula ?? v.value), '=>', v.calculation?.value ?? v.calculation?.status ?? '');
  }
  const out = await book.export(st);
  await fs.copyFile(out.path, process.env.OUT ?? path.join(dir, 'out.xlsx'));
  console.log('saved');
} catch (e) {
  console.log('ERROR', e.message);
}
await book.close();
closeNativeWorkbookWorker();
