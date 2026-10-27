import { spawnSync } from 'node:child_process';
import test, { after } from 'node:test';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import assert from 'node:assert/strict';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
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
for (const kind of ['shared', 'cache', 'iterative'])
  test('append a blank worksheet to ' + kind + ' without rewriting existing cells or opaque parts', async (t) => {
    const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-preserved-append-')),
      file = path.join(dir, 'source.xlsx'),
      bytes = preservedFixture(kind);
    await fs.writeFile(file, bytes);
    const book = new IndexedWorkbook(
      client.call,
      () => assert.fail('No fallback'),
      (i) => calculateIndexedWorkbook(i, client.call)
    );
    t.after(async () => {
      await book.close();
      await fs.rm(dir, { recursive: true, force: true });
    });
    await book.open(file);
    const next = await book.apply('0', [{ operation: 'addSheet', sheet: 'Working data' }]);
    assert.deepEqual(
      next.view.sheets.map((s) => s.name),
      ['Sheet1', 'Working data']
    );
    const saved = await fs.readFile((await book.export(next.view.storage.state)).path),
      a = unzipSync(bytes),
      b = unzipSync(saved);
    for (const part of Object.keys(a))
      if (!['xl/workbook.xml', 'xl/_rels/workbook.xml.rels', '[Content_Types].xml'].includes(part))
        assert.deepEqual(b[part], a[part], part);
    const exportedBlank = path.join(dir, 'blank-append.xlsx');
    await fs.writeFile(exportedBlank, saved);
    const check = spawnSync(
      process.env.XLSX_EDITOR_PYTHON ?? 'python3',
      [
        new URL('../scripts/test/xlsx-stream-preservation.py', import.meta.url).pathname,
        '--append-sheet',
        file,
        exportedBlank,
        'Working data',
      ],
      { encoding: 'utf8' }
    );
    assert.equal(check.status, 0, check.stdout + check.stderr);
    assert.equal(Object.keys(b).length, Object.keys(a).length + 1);
    assert(strFromU8(b['xl/worksheets/sheet2.xml']).includes('<sheetData/>'));
    assert.equal((await book.read(next.view.storage.state, 'Sheet1', 'A1')).cells.A1.calculation.pending, true);
    const edited = await book.apply(next.view.storage.state, [
      { sheet: 'Working data', cell: 'A1', value: 'New worksheet data' },
    ]);
    const exported = await book.export(edited.view.storage.state);
    assert.equal(
      readWorkbook(await fs.readFile(exported.path)).sheets.find((s) => s.name === 'Working data').cells.A1.value,
      'New worksheet data'
    );
    await assert.rejects(
      book.apply(next.view.storage.state, [{ operation: 'renameSheet', sheet: 'Sheet1', newName: 'Changed' }]),
      /protect/
    );
    await assert.rejects(
      book.apply(next.view.storage.state, [{ operation: 'insertRow', sheet: 'Sheet1', index: 0 }]),
      /protect/
    );
    assert.deepEqual(await fs.readFile(file), bytes);
  });
test('appended worksheet, values and history survive Document Save and cold reopening', async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-append-document-')),
    file = path.join(dir, 'source.xlsx');
  await fs.writeFile(file, preservedFixture('shared'));
  const owners = [];
  t.after(async () => {
    for (const owner of owners) await owner.close();
    await fs.rm(dir, { recursive: true, force: true });
  });
  const owner = new SpreadsheetDocument({ cacheIdentity: null, recoveryDirectory: path.join(dir, 'recovery') });
  owners.push(owner);
  await owner.open(file);
  await owner.request({
    operation: 'apply',
    state: owner.status().state,
    commands: [{ operation: 'addSheet', sheet: 'Working data' }],
  });
  await owner.request({ operation: 'undo' });
  assert.equal(owner.status().view.sheets.length, 1);
  await owner.request({ operation: 'redo' });
  await owner.request({
    operation: 'apply',
    state: owner.status().state,
    commands: [{ sheet: 'Working data', cell: 'A1', value: 7 }],
  });
  await owner.request({ operation: 'save' });
  await owner.close();
  const cold = new SpreadsheetDocument({ cacheIdentity: null, recoveryDirectory: path.join(dir, 'cold') });
  owners.push(cold);
  const opened = await cold.open(file);
  assert.equal(opened.view.sheets.length, 2);
  assert.equal(
    (await cold.request({ operation: 'read', state: opened.state, sheet: 'Working data', range: 'A1' })).cells.A1.value,
    7
  );
});

test('new worksheet never inherits foreign template attributes; duplicate names refuse before accepted state changes', async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-sheet-template-')),
    file = path.join(dir, 'source.xlsx'),
    parts = unzipSync(preservedFixture('cache'));
  parts['xl/workbook.xml'] = strToU8(
    strFromU8(parts['xl/workbook.xml']).replace(
      'name="Sheet1"',
      'xmlns:u="urn:keep" u:marker="original-only" name="Sheet1"'
    )
  );
  await fs.writeFile(file, zipSync(parts));
  const book = new IndexedWorkbook(
    client.call,
    () => assert.fail('No fallback'),
    (i) => calculateIndexedWorkbook(i, client.call)
  );
  t.after(async () => {
    await book.close();
    await fs.rm(dir, { recursive: true, force: true });
  });
  await book.open(file);
  await assert.rejects(book.apply('0', [{ operation: 'addSheet', sheet: 'sheet1' }]), /already exists/);
  assert.equal(book.view('0').sheets.length, 1);
  const added = await book.apply('0', [{ operation: 'addSheet', sheet: 'New' }]),
    saved = unzipSync(await fs.readFile((await book.export(added.view.storage.state)).path));
  const raw = strFromU8(saved['xl/workbook.xml']);
  assert.equal((raw.match(/original-only/g) || []).length, 1);
  assert(!raw.match(/<sheet[^>]*name="New"[^>]*u:marker/));
});
