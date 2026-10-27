/** Native formula-output markers on point reads and bounded scans, for Auto and Manual sources. */
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import assert from 'node:assert/strict';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook } from '../core/workbook-package.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { createWorkbookNativeClient } from '../core/native-client.mjs';
import { nativeWorkbookBinary } from '../core/native-runtime.mjs';
import { record } from './support/file-record.mjs';

const output = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-formula-output-owner-'));
const helper = nativeWorkbookBinary();
const client = createWorkbookNativeClient(helper),
  owners = [];
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const marker = { schema: 1, master: 'A1', range: 'A1:A2' };
function fixture(mode, owned = false) {
  const parts = unzipSync(createWorkbook());
  parts['xl/workbook.xml'] = strToU8(
    strFromU8(parts['xl/workbook.xml'])
      .replace(/<calcPr\b[^>]*(?:\/>|>[\s\S]*?<\/calcPr>)/g, '')
      .replace('</workbook>', `<calcPr calcMode="${mode}" calcOnSave="0"/></workbook>`)
  );
  const rows = owned
    ? '<row r="1"><c r="A1"><f>SEQUENCE($J$1,2)</f><v>777</v></c><c r="D1"><f>SUM(_xlfn.ANCHORARRAY(A1))</f><v>888</v></c><c r="J1"><v>3</v></c></row>'
    : '<row r="1"><c r="A1" t="str"><f t="array" ref="A1:A2">IF(1,&quot;helo&quot;,&quot;world&quot;)</f><v>helo</v></c><c r="C1" t="str"><f t="shared" si="0" ref="C1:C2">&quot;helo&quot;</f><v>helo</v></c><c r="D1"><v>7</v></c></row><row r="2"><c r="A2" t="str"><v>wrld</v></c><c r="C2" t="str"><f t="shared" si="0"/><v>wrld</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>outside</t></is></c></row>';
  parts['xl/worksheets/sheet1.xml'] = strToU8(`<worksheet xmlns="${ns}"><sheetData>${rows}</sheetData></worksheet>`);
  parts['custom/formula-output-opaque.bin'] = new Uint8Array([0, 255, 17, 39]);
  return Buffer.from(zipSync(parts));
}
async function open(name, bytes) {
  const directory = path.join(output, name);
  await fs.mkdir(directory);
  const source = path.join(directory, 'source.xlsx');
  await fs.writeFile(source, bytes);
  let root;
  const native = async (input) => {
    const result = await client.call(input);
    if (input.root) root = input.root;
    return result;
  };
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('No legacy writer expected'),
    (input) => calculateIndexedWorkbook(input, native)
  );
  owners.push(book);
  await book.open(source, { storageDirectory: path.join(directory, 'indexes') });
  return { book, source, bytes, root: () => root };
}
async function footprint(root) {
  const files = [];
  async function visit(directory) {
    for (const entry of await fs.readdir(directory, { withFileTypes: true })) {
      const file = path.join(directory, entry.name);
      assert(!entry.isSymbolicLink());
      if (entry.isDirectory()) await visit(file);
      else if (entry.isFile()) files.push(await record(file));
    }
  }
  await visit(root);
  return files.sort((a, b) => a.path.localeCompare(b.path));
}
async function scan(f, state) {
  const cells = {},
    cursors = [];
  let after;
  do {
    const page = await f.book.scan(state, 'Sheet1', after, { limit: 2 });
    for (const [ref, cell] of Object.entries(page.cells)) {
      assert(!Object.hasOwn(cells, ref), 'No duplicate cursor records');
      cells[ref] = cell;
    }
    if (page.nextCursor !== null) {
      assert(!cursors.includes(page.nextCursor), 'Cursor advances');
      cursors.push(page.nextCursor);
    }
    after = page.nextCursor;
  } while (after !== null);
  return { cells, cursors };
}
async function pureRead(f, state, range) {
  const before = await footprint(f.root()),
    viewBefore = f.book.view(state),
    read = await f.book.read(state, 'Sheet1', range),
    scanned = await scan(f, state),
    after = await footprint(f.root());
  assert.deepEqual(after, before, 'Point and scan must not mutate persistent index files');
  assert.deepEqual(f.book.view(state), viewBefore);
  for (const [ref, cell] of Object.entries(read.cells))
    assert.deepEqual(scanned.cells[ref], cell, 'Point and scan decorate identically: ' + ref);
  return { view: viewBefore, read, scanned, before, after };
}
async function test(name, body) {
  await body();
}
try {
  for (const mode of ['auto', 'manual'])
    await test('source-' + mode + '-four-formula-records-point-scan-purity', async () => {
      const f = await open('source-' + mode, fixture(mode)),
        result = await pureRead(f, '0', 'A1:D3'),
        cells = result.read.cells;
      assert.deepEqual(cells.A2.formulaOutput, marker);
      assert.equal(cells.A2.value, 'wrld');
      assert.equal(cells.A2.rawValue, 'wrld');
      assert.equal(cells.A2.dataType, 'text');
      assert.equal(cells.A2.formula, undefined);
      assert.equal(cells.A2.formulaType, undefined);
      assert.equal(cells.A2.spill, undefined);
      for (const ref of ['A1', 'C1', 'C2', 'D1', 'A3'])
        assert.equal(cells[ref].formulaOutput, undefined, ref + ' is not a source array follower');
      assert.equal(cells.A1.formulaType, 'array');
      assert.equal(cells.C1.formulaType, 'shared');
      assert.equal(cells.C2.formulaType, 'shared');
      if (mode === 'manual') {
        assert.equal(cells.A2.calculation.pending, true);
        assert.equal(cells.A2.calculation.status, 'stored');
        assert.equal(cells.A2.calculation.value, undefined);
        assert.equal(cells.A2.cachedResult.authority, 'stored');
        assert.equal(cells.A2.cachedResult.value, 'wrld');
      } else {
        assert.equal(
          cells.A2.calculation,
          undefined,
          'An unqualified source follower is not made current or pending by decoration'
        );
        assert.equal(cells.A2.cachedResult, undefined);
      }
      assert.deepEqual(await fs.readFile(f.source), f.bytes);
      return { ...result, source: await record(f.source) };
    });
  await test('owned-spill-current-pending-shrink-and-retained-version-authority', async () => {
    const f = await open('owned-spill', fixture('auto', true)),
      initial = await pureRead(f, '0', 'A1:J3');
    assert.equal(initial.read.cells.B3.value, 6);
    assert.equal(initial.read.cells.B3.spill.role, 'follower');
    assert.equal(initial.read.cells.B3.formulaOutput, undefined);
    assert.equal(initial.read.cells.D1.calculation.value, 21);
    const manual = await f.book.apply('0', [
        { type: 'setCalculationOptions', mode: 'manual', recalculateBeforeSave: false },
      ]),
      manualState = manual.view.storage.state;
    const edit = await f.book.apply(manualState, [{ sheet: 'Sheet1', cell: 'J1', value: 2 }]),
      pendingState = edit.view.storage.state,
      pending = await pureRead(f, pendingState, 'A1:J3');
    assert.equal(pending.read.cells.B3.value, 6);
    assert.equal(pending.read.cells.B3.spill.role, 'follower');
    assert.equal(pending.read.cells.B3.formulaOutput, undefined);
    assert.equal(pending.read.cells.B3.calculation.pending, true);
    assert.equal(pending.read.cells.B3.calculation.value, undefined);
    assert.equal(pending.read.cells.B3.cachedResult.value, 6);
    assert.equal(pending.read.cells.B3.cachedResult.authority, 'stored');
    const calculated = await f.book.apply(pendingState, [{ type: 'recalculate' }]),
      currentState = calculated.view.storage.state,
      current = await pureRead(f, currentState, 'A1:J3');
    assert.equal(current.read.cells.B2.value, 4);
    assert.equal(current.read.cells.B2.spill.role, 'follower');
    assert.equal(current.read.cells.B2.formulaOutput, undefined);
    assert.equal(current.read.cells.B3.value, null);
    assert.equal(current.read.cells.B3.spill, undefined);
    assert.equal(current.read.cells.B3.formulaOutput, undefined);
    assert.equal(current.read.cells.D1.calculation.value, 10);
    const retainedPending = await pureRead(f, pendingState, 'A1:J3'),
      retainedInitial = await pureRead(f, '0', 'A1:J3');
    assert.deepEqual(retainedPending.read, pending.read);
    assert.deepEqual(retainedInitial.read, initial.read);
    assert.deepEqual(await fs.readFile(f.source), f.bytes);
    return {
      initial,
      pending,
      current,
      retainedPending,
      retainedInitial,
      manualState,
      pendingState,
      currentState,
      source: await record(f.source),
    };
  });
} finally {
  const closing = await Promise.allSettled(owners.map((owner) => owner.close()));
  await client.close();
  await fs.rm(output, { recursive: true, force: true });
  for (const result of closing) if (result.status === 'rejected') throw result.reason;
  assert(owners.every((owner) => owner.diagnosticStats().closed === 1));
}
