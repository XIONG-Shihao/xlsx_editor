// Apply editor commands to a workbook and save the result.
//
//   node --import ./scripts/ts-resolve-register.mjs scripts/audit/save-edited.mjs SOURCE.xlsx OUTPUT.xlsx 'JSON commands'
//
// The source is only read (a copy is edited). Each command is applied as its own step.
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { IndexedWorkbook } from '../../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../../core/native-runtime.mjs';
import { editCalculatedWorkbookResult } from '@office-editors/xlsx-editor/engine';

export async function saveEdited(source, output, commands, scratch) {
  const own = scratch ?? (await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-save-edited-'))),
    copy = path.join(own, 'source' + path.extname(source));
  await fs.mkdir(path.join(own, 'index'), { recursive: true });
  await fs.copyFile(source, copy);
  const book = new IndexedWorkbook(
    native,
    async (file, edits, out) => {
      const result = await editCalculatedWorkbookResult(await fs.readFile(file), edits);
      await fs.writeFile(out, result.bytes);
      return { revision: result.view.revision, edits: result.edits };
    },
    (input) => calculateIndexedWorkbook(input, native)
  );
  try {
    const opened = await book.open(copy, { storageDirectory: path.join(own, 'index') });
    let state = '0';
    for (const command of commands) state = (await book.apply(state, [command])).view.storage.state;
    await fs.copyFile((await book.export(state)).path, output);
    return opened;
  } finally {
    await book.close();
    if (!scratch) await fs.rm(own, { recursive: true, force: true });
  }
}

if (process.argv[1] === new URL(import.meta.url).pathname) {
  const [source, output, commands] = process.argv.slice(2);
  if (!source || !output || !commands) throw Error('Usage: save-edited.mjs SOURCE.xlsx OUTPUT.xlsx JSON_COMMANDS');
  await saveEdited(path.resolve(source), path.resolve(output), JSON.parse(commands));
  closeNativeWorkbookWorker();
}
