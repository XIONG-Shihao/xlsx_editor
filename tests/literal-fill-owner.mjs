/** Actual Document owner: literal-error fill/internal Paste/Cut and durable typed source. */
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import assert from 'node:assert/strict';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { createWorkbookNativeClient } from '../core/native-client.mjs';
import { nativeWorkbookBinary } from '../core/native-runtime.mjs';
import { readWorkbook } from '../core/workbook-package.mjs';
import { valuePasteFixture, valueSheet } from './support/value-paste-fixture.mjs';
import { record } from './support/file-record.mjs';
const output = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-literal-fill-owner-'));
const helper = nativeWorkbookBinary();
const client = createWorkbookNativeClient(helper),
  owners = [];
function fixture(options) {
  const parts = unzipSync(valuePasteFixture(options)),
    part = 'xl/worksheets/sheet1.xml',
    text = strFromU8(parts[part]);
  assert(/<c r="C6" s="1"[^>]*><\/c>/.test(text));
  parts[part] = strToU8(text.replace(/<c r="C6" s="1"[^>]*><\/c>/, '<c r="C6" s="1" t="e"><v>#REF!</v></c>'));
  return Buffer.from(zipSync(parts));
}
async function open(name, options = {}) {
  const directory = path.join(output, name);
  await fs.mkdir(directory);
  const source = path.join(directory, 'source.xlsx'),
    bytes = fixture(options);
  await fs.writeFile(source, bytes);
  const recoveryDirectory = path.join(directory, 'recovery');
  let root, version;
  const native = async (input) => {
    root = input.root ?? root;
    version = input.version ?? version;
    try {
      const result = await client.call(input);
      version = result.version ?? version;
      return result;
    } catch (error) {
      throw error;
    }
  };
  async function create() {
    const owner = new SpreadsheetDocument({ recoveryDirectory, native, cacheIdentity: 'literal-fill-owner' });
    owners.push(owner);
    await owner.open(source);
    return owner;
  }
  const owner = await create();
  return { owner, source, bytes, create, native, root: () => root, version: () => version };
}
const read = (owner, range) =>
  owner.request({ operation: 'read', state: owner.status().state, sheet: valueSheet, range });
const apply = (owner, commands) => owner.request({ operation: 'apply', state: owner.status().state, commands });
async function check(name, body) {
  await body();
}
function error(cell, code) {
  assert.equal(cell.value, code);
  assert.equal(cell.dataType, 'error');
  assert.equal(cell.error, code);
  assert.equal(cell.formula, undefined);
  assert.equal(cell.calculation, undefined);
}
function style(a, b) {
  for (const key of [
    'styleId',
    'format',
    'fontFamily',
    'fontSize',
    'bold',
    'italic',
    'underline',
    'fontColor',
    'fillColor',
    'wrapText',
    'border',
  ])
    assert.deepEqual(a[key], b[key], key);
}
async function footprint(root) {
  const files = [];
  async function walk(dir) {
    for (const entry of await fs.readdir(dir, { withFileTypes: true })) {
      assert(!entry.isSymbolicLink());
      const filename = path.join(dir, entry.name);
      if (entry.isDirectory()) await walk(filename);
      else files.push(await record(filename));
    }
  }
  await walk(root);
  return files.sort((a, b) => a.path.localeCompare(b.path));
}
try {
  await check('literal-error-down-right-paste-source-styles-one-history-copy-save-clean-cold', async () => {
    const f = await open('ordinary'),
      owner = f.owner,
      before = owner.status(),
      initial = (await read(owner, 'A1:H20')).cells;
    const accepted = await apply(owner, [
        { type: 'fillRange', sheet: valueSheet, source: 'A5:A6', range: 'E5:E6' },
        { type: 'fillRange', sheet: valueSheet, source: 'A6', range: 'A7:A9' },
        { type: 'fillRange', sheet: valueSheet, source: 'C6', range: 'E12:G12' },
      ]),
      head = owner.status().state,
      cells = (await read(owner, 'A1:H20')).cells;
    assert.equal(cells.E5.dataType, 'text');
    assert.equal(cells.E5.value, '#DIV/0!');
    for (const ref of ['E6', 'A7', 'A8', 'A9']) {
      error(cells[ref], '#N/A');
      style(cells[ref], initial.A6);
    }
    for (const ref of ['E12', 'F12', 'G12']) {
      error(cells[ref], '#REF!');
      style(cells[ref], initial.C6);
    }
    assert.deepEqual(await fs.readFile(f.source), f.bytes);
    await owner.request({ operation: 'undo' });
    assert.equal(owner.status().state, before.state);
    assert.equal(owner.status().canUndo, false);
    assert.deepEqual((await read(owner, 'A1:H20')).cells, initial);
    await owner.request({ operation: 'redo' });
    assert.equal(owner.status().state, head);
    assert.deepEqual((await read(owner, 'A1:H20')).cells, cells);
    const copyFile = path.join(output, 'ordinary-copy.xlsx'),
      copy = await owner.request({ operation: 'copy', destination: copyFile, expectedRevision: null }),
      copied = readWorkbook(await fs.readFile(copyFile)).sheets[0].cells;
    error(copied.E6, '#N/A');
    error(copied.E12, '#REF!');
    assert.equal(copied.E5.dataType, 'text');
    assert.equal(copied.E12.fontFamily, initial.C6.fontFamily);
    assert.deepEqual(await fs.readFile(f.source), f.bytes);
    const saved = await owner.request({ operation: 'save' }),
      savedBytes = await fs.readFile(f.source),
      serialized = readWorkbook(savedBytes).sheets[0].cells;
    error(serialized.E6, '#N/A');
    error(serialized.E12, '#REF!');
    await owner.close();
    const cold = await f.create(),
      coldCells = (await read(cold, 'A1:H20')).cells;
    for (const ref of ['E6', 'A7', 'A8', 'A9']) error(coldCells[ref], '#N/A');
    for (const ref of ['E12', 'F12', 'G12']) {
      error(coldCells[ref], '#REF!');
      style(coldCells[ref], serialized[ref]);
    }
    assert.equal(coldCells.E5.dataType, 'text');
    assert.equal(cold.status().canUndo, false);
    assert.equal(cold.status().canRedo, false);
    assert.equal(cold.status().dirty, false);
    assert.deepEqual(await fs.readFile(f.source), savedBytes);
    return { before, initial, accepted, head, cells, copy, saved, cold: cold.status(), coldCells };
  });
  await check('ordinary-same-owner-cut-command-batch-one-undo-redo-manual-cache-and-cold-type', async () => {
    const f = await open('cut', { mode: 'manual' }),
      owner = f.owner,
      before = owner.status(),
      initial = (await read(owner, 'A1:H20')).cells;
    const accepted = await apply(owner, [
        { type: 'fillRange', sheet: valueSheet, source: 'C6', range: 'H12' },
        { type: 'clearRange', sheet: valueSheet, range: 'C6' },
      ]),
      head = owner.status().state,
      cells = (await read(owner, 'A1:H20')).cells;
    error(cells.H12, '#REF!');
    style(cells.H12, initial.C6);
    assert.equal(cells.C6.value, null);
    assert.equal(cells.C6.error, undefined);
    style(cells.C6, initial.C6);
    assert.equal(cells.B1.calculation.pending, true);
    assert.equal(cells.B1.calculation.value, undefined);
    await owner.request({ operation: 'undo' });
    assert.equal(owner.status().state, before.state);
    assert.equal(owner.status().canUndo, false);
    assert.deepEqual((await read(owner, 'A1:H20')).cells, initial);
    await owner.request({ operation: 'redo' });
    assert.equal(owner.status().state, head);
    error((await read(owner, 'H12')).cells.H12, '#REF!');
    assert.deepEqual(await fs.readFile(f.source), f.bytes);
    await owner.request({ operation: 'save' });
    const saved = await fs.readFile(f.source);
    await owner.close();
    const cold = await f.create(),
      coldCells = (await read(cold, 'A1:H20')).cells;
    error(coldCells.H12, '#REF!');
    assert.equal(coldCells.C6.value, null);
    assert.equal(coldCells.B1.cachedResult.value, 103);
    assert.equal(coldCells.B1.calculation.pending, true);
    assert.equal(coldCells.B1.calculation.value, undefined);
    assert.deepEqual(await fs.readFile(f.source), saved);
    return { before, initial, accepted, head, cells, cold: cold.status(), coldCells };
  });
  await check('late-protected-source-output-refuses-native-prefix-and-public-cut-history', async () => {
    const f = await open('protected', { boundary: 'unknown-array' }),
      owner = f.owner,
      before = owner.status(),
      cells = (await read(owner, 'A1:H20')).cells,
      outputCell = (await read(owner, 'R1101')).cells.R1101,
      root = f.root(),
      version = f.version(),
      sheet = before.view.sheets.find((sheet) => sheet.name === valueSheet).path;
    assert.deepEqual(outputCell.formulaOutput, { schema: 1, master: 'R1100', range: 'R1100:R1101' });
    const commands = [
        { type: 'fillRange', sheet, source: 'A6', range: 'E6' },
        { type: 'fillRange', sheet, source: 'R1101', range: 'E7' },
      ],
      pins = await footprint(root);
    await assert.rejects(
      f.native({
        action: 'workbook',
        operation: 'bulk',
        root,
        version,
        commands,
        view: {
          sheets: [
            { path: sheet, cells: { R1101: { value: 914, calculation: { status: 'calculated', value: 914 } } } },
          ],
        },
      }),
      /Source formula output/
    );
    assert.deepEqual(await footprint(root), pins);
    await assert.rejects(
      apply(owner, [
        { type: 'fillRange', sheet: valueSheet, source: 'A6', range: 'E6' },
        { type: 'clearRange', sheet: valueSheet, range: 'A6' },
        { type: 'fillRange', sheet: valueSheet, source: 'R1101', range: 'E7' },
      ]),
      /Source formula output/
    );
    assert.equal(owner.status().state, before.state);
    assert.equal(owner.status().canUndo, false);
    assert.deepEqual((await read(owner, 'A1:H20')).cells, cells);
    assert.deepEqual(await fs.readFile(f.source), f.bytes);
    return { before, cells, outputCell, commands };
  });
} finally {
  const closing = await Promise.allSettled(owners.map((owner) => owner.close()));
  await client.close();
  await fs.rm(output, { recursive: true, force: true });
  for (const result of closing) if (result.status === 'rejected') throw result.reason;
}
