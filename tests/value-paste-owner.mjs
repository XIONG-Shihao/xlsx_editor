/** Actual Indexed/native Values: typed snapshots, immutable source and atomic refusal. */
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import assert from 'node:assert/strict';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { createWorkbookNativeClient } from '../core/native-client.mjs';
import { nativeWorkbookBinary } from '../core/native-runtime.mjs';
import { readWorkbook } from '../core/workbook-package.mjs';
import { valuePasteFixture, valueSheet, literalResults, formulaResults } from './support/value-paste-fixture.mjs';
import { record } from './support/file-record.mjs';
const output = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-value-paste-owner-'));
const helper = nativeWorkbookBinary();
const client = createWorkbookNativeClient(helper),
  books = [];
async function open(name, options = {}) {
  const directory = path.join(output, name);
  await fs.mkdir(directory);
  const source = path.join(directory, 'source.xlsx'),
    bytes = valuePasteFixture(options);
  await fs.writeFile(source, bytes);
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
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('No legacy writer'),
    (input) => calculateIndexedWorkbook(input, native)
  );
  books.push(book);
  await book.open(source, { storageDirectory: path.join(directory, 'indexes') });
  return { book, native, source, bytes, root: () => root, version: () => version };
}
async function check(name, body) {
  await body();
}
async function footprint(root) {
  const files = [];
  async function walk(directory) {
    for (const entry of await fs.readdir(directory, { withFileTypes: true })) {
      assert(!entry.isSymbolicLink());
      const filename = path.join(directory, entry.name);
      if (entry.isDirectory()) await walk(filename);
      else files.push(await record(filename));
    }
  }
  await walk(root);
  return files.sort((a, b) => a.path.localeCompare(b.path));
}
function typed(cell, expected) {
  assert.equal(cell.dataType, expected.kind);
  assert(Object.is(cell.value, expected.value), JSON.stringify({ cell, expected }));
  assert.equal(cell.formula, undefined);
  assert.equal(cell.calculation, undefined);
  assert.equal(cell.cachedValue, null);
}
try {
  await check('actual-literals-and-current-formula-results-preserve-real-error-vs-text-and-target-style', async () => {
    const f = await open('auto'),
      before = (await f.book.read('0', valueSheet, 'A1:H20')).cells;
    assert.equal(before.B4.calculation.status, 'error');
    assert.equal(before.B4.calculation.error, '#DIV/0!');
    const result = await f.book.apply('0', [
        { type: 'pasteValues', sheet: valueSheet, source: 'A1:A8', range: 'C1:C8' },
        { type: 'pasteValues', sheet: valueSheet, source: 'B1:B4', range: 'D1:D4' },
      ]),
      state = result.view.storage.state,
      cells = (await f.book.read(state, valueSheet, 'A1:H20')).cells;
    for (const [index, expected] of literalResults.entries()) typed(cells['C' + (index + 1)], expected);
    for (const [index, expected] of formulaResults.entries()) typed(cells['D' + (index + 1)], expected);
    for (const ref of ['C1', 'C6', 'C8', 'D1', 'D4'])
      for (const key of ['styleId', 'format', 'fontFamily', 'fontSize', 'bold', 'italic', 'fontColor', 'fillColor'])
        assert.deepEqual(cells[ref][key], before[ref][key]);
    assert.deepEqual((await f.book.read('0', valueSheet, 'A1:H20')).cells, before);
    const exported = await f.book.export(state),
      bytes = await fs.readFile(exported.path);
    await fs.writeFile(path.join(output, 'auto-values-copy.xlsx'), bytes);
    const readback = readWorkbook(bytes).sheets.find((s) => s.name === valueSheet).cells;
    assert.equal(readback.C6.dataType, 'error');
    assert.equal(readback.C5.dataType, 'text');
    assert.equal(readback.D4.dataType, 'error');
    assert.equal(readback.D4.value, '#DIV/0!');
    const styled = await f.book.apply(state, [
        { type: 'formatRange', sheet: valueSheet, range: 'C6', format: { bold: false } },
      ]),
      styledState = styled.view.storage.state,
      styledCells = (await f.book.read(styledState, valueSheet, 'C6')).cells;
    assert.equal(styledCells.C6.dataType, 'error');
    assert.equal(styledCells.C6.error, '#N/A');
    assert.equal(styledCells.C6.bold, false);
    const styledExport = await f.book.export(styledState),
      styledBytes = await fs.readFile(styledExport.path);
    await fs.writeFile(path.join(output, 'auto-error-style-copy.xlsx'), styledBytes);
    const styledReadback = readWorkbook(styledBytes).sheets.find((s) => s.name === valueSheet).cells;
    assert.equal(styledReadback.C6.dataType, 'error');
    assert.equal(styledReadback.C6.bold, false);
    const replaced = await f.book.apply(styledState, [
        {
          type: 'setCells',
          sheet: valueSheet,
          cells: [
            { cell: 'C6', value: '#N/A' },
            { cell: 'D4', value: 12 },
          ],
        },
      ]),
      replacementState = replaced.view.storage.state,
      replacementCells = (await f.book.read(replacementState, valueSheet, 'C1:D8')).cells;
    typed(replacementCells.C6, { kind: 'text', value: '#N/A' });
    typed(replacementCells.D4, { kind: 'number', value: 12 });
    assert.equal(replacementCells.C6.error, undefined);
    assert.equal(replacementCells.D4.error, undefined);
    assert.equal(replacementCells.C6.bold, false);
    const replacementExport = await f.book.export(replacementState),
      replacementBytes = await fs.readFile(replacementExport.path);
    await fs.writeFile(path.join(output, 'auto-error-replaced-copy.xlsx'), replacementBytes);
    const replacementReadback = readWorkbook(replacementBytes).sheets.find((s) => s.name === valueSheet).cells;
    assert.equal(replacementReadback.C6.dataType, 'text');
    assert.equal(replacementReadback.D4.dataType, 'number');
    assert.equal(replacementReadback.D4.value, 12);
    assert.equal(replacementReadback.C6.bold, false);
    assert.deepEqual((await f.book.read(state, valueSheet, 'A1:H20')).cells, cells);
    assert.deepEqual(await fs.readFile(f.source), f.bytes);
    return {
      before,
      result,
      cells,
      exported,
      readback,
      styled,
      styledCells,
      styledReadback,
      replaced,
      replacementCells,
      replacementReadback,
    };
  });
  await check('manual-stored-current-policy-refusal-and-late-private-candidate-leaves-retained-head', async () => {
    const f = await open('manual', { mode: 'manual' }),
      view = f.book.view('0'),
      before = (await f.book.read('0', valueSheet, 'A1:H20')).cells;
    assert.equal(before.B1.calculation.pending, true);
    assert.equal(before.B1.calculation.value, undefined);
    await assert.rejects(
      f.book.apply('0', [
        { type: 'pasteValues', sheet: valueSheet, source: 'A1:A8', range: 'C1:C8' },
        { type: 'pasteValues', sheet: valueSheet, source: 'B1:B4', range: 'D1:D4' },
      ]),
      /current calculated|pending/
    );
    assert.deepEqual(f.book.view('0'), view);
    assert.deepEqual((await f.book.read('0', valueSheet, 'A1:H20')).cells, before);
    const result = await f.book.apply('0', [
        { type: 'pasteValues', sheet: valueSheet, source: 'A1:A8', range: 'C1:C8' },
      ]),
      cells = (await f.book.read(result.view.storage.state, valueSheet, 'C1:C8')).cells;
    for (const [index, expected] of literalResults.entries()) typed(cells['C' + (index + 1)], expected);
    assert.deepEqual(await fs.readFile(f.source), f.bytes);
    return { before, view, result, cells };
  });
  await check(
    'actual-source-output-master-outside-selection-and-caller-current-forgery-refuse-before-prefix',
    async () => {
      const f = await open('unknown', { boundary: 'unknown-array' }),
        view = f.book.view('0'),
        before = (await f.book.read('0', valueSheet, 'R1101')).cells,
        root = f.root(),
        version = f.version(),
        sheet = view.sheets.find((s) => s.name === valueSheet).path;
      assert.deepEqual(before.R1101.formulaOutput, { schema: 1, master: 'R1100', range: 'R1100:R1101' });
      const attempts = [
          [{ type: 'pasteValues', sheet, source: 'R1101', range: 'D11' }],
          [
            { type: 'pasteValues', sheet, source: 'A1', range: 'C1' },
            { type: 'pasteValues', sheet, source: 'R1101', range: 'D11' },
          ],
        ],
        refusals = [];
      for (const commands of attempts) {
        const pins = await footprint(root);
        let error;
        try {
          await f.native({
            action: 'workbook',
            operation: 'bulk',
            root,
            version,
            commands,
            view: {
              sheets: [
                {
                  path: sheet,
                  cells: {
                    R1101: { value: 914, calculation: { status: 'calculated', value: 914 }, formulaOutput: null },
                  },
                },
              ],
              tables: [],
            },
          });
        } catch (e) {
          error = String(e);
        }
        assert.match(error, /current calculated|source output/);
        assert.deepEqual(await footprint(root), pins);
        refusals.push({ commands, error });
      }
      await assert.rejects(
        f.book.apply('0', [{ type: 'pasteValues', sheet: valueSheet, source: 'R1101', range: 'D11' }]),
        /current calculated|source output/
      );
      assert.deepEqual(f.book.view('0'), view);
      assert.deepEqual(await fs.readFile(f.source), f.bytes);
      return { before, refusals };
    }
  );
  await check(
    'actual-current-owned-sequence-follower-is-a-typed-snapshot-and-target-owner-stays-protected',
    async () => {
      const f = await open('owned', { boundary: 'owned-spill' }),
        before = (await f.book.read('0', valueSheet, 'A10:D11')).cells;
      assert.deepEqual(before.B10.spill, { schema: 1, master: 'A10', range: 'A10:B11', role: 'follower' });
      assert.equal(before.B10.calculation.status, 'calculated');
      assert.equal(before.B10.calculation.value, 2);
      assert.notEqual(before.B10.calculation.pending, true);
      const result = await f.book.apply('0', [{ type: 'pasteValues', sheet: valueSheet, source: 'B10', range: 'D10' }]),
        state = result.view.storage.state,
        cells = (await f.book.read(state, valueSheet, 'A10:D11')).cells;
      typed(cells.D10, { kind: 'number', value: 2 });
      assert.equal(cells.D10.spill, undefined);
      assert.equal(cells.D10.styleId, before.D10.styleId);
      assert.deepEqual((await f.book.read('0', valueSheet, 'A10:D11')).cells, before);
      const view = f.book.view(state);
      await assert.rejects(
        f.book.apply(state, [{ type: 'pasteValues', sheet: valueSheet, source: 'A4', range: 'B10' }]),
        /spill follower/
      );
      assert.deepEqual(f.book.view(state), view);
      assert.deepEqual(await fs.readFile(f.source), f.bytes);
      return { before, result, cells };
    }
  );
} finally {
  const closing = await Promise.allSettled(books.map((book) => book.close()));
  await client.close();
  await fs.rm(output, { recursive: true, force: true });
  for (const result of closing) if (result.status === 'rejected') throw result.reason;
}
