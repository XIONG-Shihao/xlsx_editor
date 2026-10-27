// Refusal codes: the list in core/refusal-codes.mjs and the code that uses it
// stay in step, and a code survives the trip from the engine to the caller.
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { zipSync, unzipSync, strToU8 } from 'fflate';
import { REFUSAL_CODES, refusal, reworded } from '../core/refusal-codes.mjs';
import { pendingCodes } from '../scripts/refusal-codes/check.mjs';
import { createWorkbook } from '../core/workbook-package.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';

after(() => closeNativeWorkbookWorker());
const component = new URL('..', import.meta.url).pathname;
const tracked = execFileSync('git', ['ls-files', 'core', 'ui', 'adapters', 'standalone', 'native/src'], {
  cwd: component,
  encoding: 'utf8',
})
  .split('\n')
  .filter((name) => /\.(mjs|ts|tsx|rs)$/.test(name));

test('every code in use is listed, and every listed code is in use', async () => {
  // Codes that the places still waiting for a code (scripts/refusal-codes/allow-list.json) will get count as in use.
  const used = new Set(pendingCodes());
  for (const name of tracked) {
    if (name === 'core/refusal-codes.mjs') continue;
    const text = await fs.readFile(path.join(component, name), 'utf8');
    for (const match of text.matchAll(/\b(?:refusal|coded|rephrase)\(\s*['"]([a-z-]+)['"]/g)) used.add(match[1]);
    for (const match of text.matchAll(/code:\s*'([a-z]+(?:-[a-z]+)+)'/g))
      if (Object.hasOwn(REFUSAL_CODES, match[1])) used.add(match[1]);
  }
  assert.deepEqual([...used].sort(), Object.keys(REFUSAL_CODES).sort());
});

test('an unknown code cannot be created, and rewording keeps the code', () => {
  assert.throws(() => refusal('made-up', 'x'), /Unknown refusal code/);
  const original = refusal('sheet-protected', 'first wording');
  assert.equal(reworded(original, 'second wording').code, 'sheet-protected');
  assert.equal(reworded(Error('plain'), 'still plain').code, undefined);
});

test('the engine sends the code apart from the sentence', async (t) => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-codes-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    parts = unzipSync(createWorkbook());
  // A merged block inside the sorted rows is something the sort refuses.
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    [
      `<worksheet xmlns="${ns}"><sheetData>`,
      '<row r="1"><c r="A1"><v>2</v></c></row><row r="2"><c r="A2"><v>1</v></c></row>',
      '</sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells></worksheet>',
    ].join('')
  );
  const source = path.join(root, 'source.xlsx');
  await fs.writeFile(source, zipSync(parts));
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('no fallback writer'),
    (input) => calculateIndexedWorkbook(input, native)
  );
  t.after(() => book.close());
  await book.open(source, { storageDirectory: path.join(root, 'index') });
  await assert.rejects(
    book.apply('0', [{ sheet: 'Sheet1', operation: 'sortRows', range: 'A1:A2', column: 'A', direction: 'asc' }]),
    (error) => {
      assert.equal(error.code, 'needs-structural-writer');
      assert.doesNotMatch(error.message, /^\[/, 'the sentence does not carry the code');
      return true;
    }
  );
});
