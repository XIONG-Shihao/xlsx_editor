import { createPasteStager } from '../adapters/large-paste.mjs';
import { inspectPaste, pastePages } from '../ui/largePaste.ts';
import { randomUUID } from 'node:crypto';
import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { spawnSync } from 'node:child_process';
import { unzipSync, strFromU8 } from 'fflate';
import { preservedFixture } from './support/preserved-calculation-fixture.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { createWorkbookNativeClient } from '../core/native-client.mjs';
import { readWorkbook } from '../core/workbook-package.mjs';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
const client = createWorkbookNativeClient(new URL('../native/bin/darwin-arm64/workbook-engine', import.meta.url).pathname);
after(() => {
  client.close();
  closeNativeWorkbookWorker();
});
async function fixture(t, kind) {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-preserved-everyday-')),
    file = path.join(dir, 'source.xlsx'),
    bytes = preservedFixture(kind);
  await fs.writeFile(file, bytes);
  const book = new IndexedWorkbook(
    client.call,
    () => assert.fail('No fallback'),
    (i) => calculateIndexedWorkbook(i, client.call)
  );
  await book.open(file);
  t.after(async () => {
    await book.close();
    await fs.rm(dir, { recursive: true, force: true });
  });
  return { dir, file, bytes, book };
}
for (const kind of ['shared', 'cache', 'iterative'])
  test(kind + ' supports new literals, clear, rich/formula formatting and literal pasted rectangles', async (t) => {
    const f = await fixture(t, kind),
      { book } = f;
    let next = await book.apply('0', [
      {
        type: 'setCells',
        sheet: 'Sheet1',
        cells: [
          { cell: 'B1', value: 7 },
          { cell: 'E1', value: '00123' },
          { cell: 'F4', value: '=SUM(A1)' },
          { cell: 'B2', value: false },
        ],
      },
    ]);
    next = await book.apply(next.view.storage.state, [
      { type: 'formatRange', sheet: 'Sheet1', range: 'A1:C1', format: { bold: true, fontColor: '123456' } },
    ]);
    next = await book.apply(next.view.storage.state, [
      { type: 'pasteValues', sheet: 'Sheet1', source: 'C1', range: 'G4' },
    ]);
    next = await book.apply(next.view.storage.state, [{ type: 'clearRange', sheet: 'Sheet1', range: 'D1' }]);
    const state = next.view.storage.state,
      cells = (await book.read(state, 'Sheet1', 'A1:G4')).cells;
    assert.equal(cells.B1.value, 7);
    assert.equal(cells.E1.value, '00123');
    assert.equal(cells.F4.value, '=SUM(A1)');
    assert.equal(cells.B2.value, false);
    assert.equal(cells.G4.value, 'Keep rich');
    assert.equal(cells.A1.calculation.pending, true);
    await assert.rejects(
      book.apply(state, [{ type: 'clearRange', sheet: 'Sheet1', range: 'A1:B1' }]),
      /protect|preserv/
    );
    await assert.rejects(
      book.apply(state, [
        {
          type: 'setCells',
          sheet: 'Sheet1',
          cells: [
            { cell: 'B1', value: 99 },
            { cell: 'A1', value: 99 },
          ],
        },
      ]),
      /protect|preserv/
    );
    assert.equal((await book.read(state, 'Sheet1', 'B1')).cells.B1.value, 7);
    const output = await book.export(state),
      saved = path.join(f.dir, 'saved.xlsx');
    await fs.copyFile(output.path, saved);
    const before = unzipSync(f.bytes),
      afterParts = unzipSync(await fs.readFile(saved));
    for (const part of Object.keys(before))
      if (!['xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/styles.xml'].includes(part))
        assert.deepEqual(afterParts[part], before[part], part);
    assert(strFromU8(afterParts['xl/worksheets/sheet1.xml']).includes('<dimension ref="A1:G4"'));
    const sourceFormula = strFromU8(before['xl/worksheets/sheet1.xml']).match(/<c r="A1">(.*?)<\/c>/)[1];
    assert(strFromU8(afterParts['xl/worksheets/sheet1.xml']).includes(sourceFormula));
    assert(strFromU8(afterParts['xl/worksheets/sheet1.xml']).includes('<r><rPr><b/></rPr><t>Keep rich</t></r>'));
    const oracle = spawnSync(
      process.env.XLSX_EDITOR_PYTHON ?? 'python3',
      [new URL('../scripts/test/xlsx-package-check.py', import.meta.url).pathname],
      {
        input: JSON.stringify({
          source: f.file,
          saved,
          strictPointEdits: true,
          targets: { Sheet1: ['A1', 'B1', 'B2', 'C1', 'D1', 'E1', 'F4', 'G4'] },
          styleTargets: { Sheet1: ['A1', 'B1', 'C1'] },
          addedLiteralCells: { Sheet1: ['E1', 'F4', 'G4'] },
          clearedLiteralCells: { Sheet1: ['D1'] },
          appendStyles: true,
          expected: {
            'Sheet1!B1': { value: 7 },
            'Sheet1!E1': { value: '00123' },
            'Sheet1!F4': { value: '=SUM(A1)' },
            'Sheet1!G4': { value: 'Keep rich' },
          },
        }),
        encoding: 'utf8',
      }
    );
    assert.equal(oracle.status, 0, oracle.stdout + oracle.stderr);
    assert.equal(readWorkbook(await fs.readFile(saved)).sheets[0].cells.B1.value, 7);
    assert.deepEqual(await fs.readFile(f.file), f.bytes);
  });
test('Document everyday edits remain usable after Save and recovery replay', async (t) => {
  const f = await fixture(t, 'cache');
  await f.book.close();
  const owner = new SpreadsheetDocument({ cacheIdentity: null, recoveryDirectory: path.join(f.dir, 'recovery') });
  t.after(() => owner.close());
  await owner.open(f.file);
  await owner.request({
    operation: 'apply',
    state: owner.status().state,
    commands: [
      {
        type: 'setCells',
        sheet: 'Sheet1',
        cells: [
          { cell: 'E2', value: 'New data' },
          { cell: 'B1', value: 8 },
        ],
      },
    ],
  });
  await owner.request({
    operation: 'apply',
    state: owner.status().state,
    commands: [{ type: 'formatRange', sheet: 'Sheet1', range: 'A1:B1', format: { italic: true } }],
  });
  await owner.request({ operation: 'save' });
  await owner.request({
    operation: 'apply',
    state: owner.status().state,
    commands: [{ type: 'clearRange', sheet: 'Sheet1', range: 'E2' }],
  });
  await owner.close();
  const recovered = new SpreadsheetDocument({ cacheIdentity: null, recoveryDirectory: path.join(f.dir, 'recovery') });
  t.after(() => recovered.close());
  const info = await recovered.open(f.file);
  assert.equal(info.recovered, true);
  const read = await recovered.request({ operation: 'read', state: info.state, sheet: 'Sheet1', range: 'A1:E2' });
  assert.equal(read.cells.B1.value, 8);
  assert.equal(read.cells.A1.cachedResult.value, 2);
  assert.equal(read.cells.E2?.value ?? null, null);
});
test('staged literal paste preserves unsupported formulas; a late formula page cannot publish an earlier partial paste', async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-preserved-staged-')),
    file = path.join(dir, 'source.xlsx');
  const bytes = preservedFixture('shared');
  await fs.writeFile(file, bytes);
  let attachmentState = '0';
  const stager = createPasteStager({ directory: path.join(dir, 'imports'), documentId: randomUUID() }),
    book = new IndexedWorkbook(
      client.call,
      () => assert.fail('No fallback'),
      (i) => calculateIndexedWorkbook(i, client.call),
      { importPages: (descriptor) => stager.pages(descriptor, { state: attachmentState }) }
    );
  t.after(async () => {
    await stager.close();
    await book.close();
    await fs.rm(dir, { recursive: true, force: true });
  });
  await book.open(file);
  async function prepare(text, start, policy) {
    const plan = inspectPaste(text, { sheet: 'Sheet1', start, policy });
    const { id } = await stager.begin({ ...plan, state: attachmentState });
    for (const page of pastePages(text, plan)) await stager.upload({ id, state: attachmentState, ...page });
    return stager.seal({ id, state: attachmentState });
  }
  const descriptor = await prepare('001\t=Literal\n東京\tFALSE', 'B2', 'text');
  const next = await book.apply('0', [descriptor]),
    state = next.view.storage.state;
  const cells = (await book.read(state, 'Sheet1', 'A1:D3')).cells;
  assert.equal(cells.B2.value, '001');
  assert.equal(cells.C2.value, '=Literal');
  assert.equal(cells.B3.value, '東京');
  assert.equal(cells.C3.value, 'FALSE');
  assert.equal(cells.A1.calculation.pending, true);
  await book.export(state);
  attachmentState = state;
  const bad = await prepare('safe\n'.repeat(256) + '=A1', 'B10', 'grid');
  await assert.rejects(book.apply(state, [bad]), { code: 'formula-protected' });
  assert.equal((await book.read(state, 'Sheet1', 'B10:B10')).cells.B10?.value ?? null, null);
  assert.equal((await book.read(state, 'Sheet1', 'B2')).cells.B2.value, '001');
  assert.deepEqual(await fs.readFile(file), bytes);
});
