/** Unsupported spill geometry must protect calculation without refusing Open. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { zipSync, unzipSync, strToU8 } from 'fflate';
import { createWorkbook } from '../core/workbook-package.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
const sha = (bytes) => createHash('sha256').update(bytes).digest('hex');
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
async function fixture(t, masters, literals = {}) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-spill-boundary-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  const parts = unzipSync(createWorkbook()),
    rows = new Map();
  const add = (ref, xml) => {
    const row = Number(ref.match(/\d+$/)[0]);
    if (!rows.has(row)) rows.set(row, []);
    rows.get(row).push(xml);
  };
  for (const [ref, formula] of masters) add(ref, `<c r="${ref}"><f>${formula}</f><v>999</v></c>`);
  for (const [ref, value] of Object.entries(literals)) add(ref, `<c r="${ref}"><v>${value}</v></c>`);
  add('F1', '<c r="F1"><v>5</v></c>');
  add('G1', '<c r="G1"><f>F1+1</f><v>999</v></c>');
  add('G2', '<c r="G2"><f>SUM(A1:C600)</f><v>999</v></c>');
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${ns}"><sheetData>${[...rows]
      .sort((a, b) => a[0] - b[0])
      .map(([r, cells]) => `<row r="${r}">${cells.join('')}</row>`)
      .join('')}</sheetData></worksheet>`
  );
  parts['customXml/spill-boundary.xml'] = strToU8('<opaque> unchanged </opaque>');
  const source = path.join(root, 'source.xlsx'),
    bytes = Buffer.from(zipSync(parts));
  await fs.writeFile(source, bytes);
  const calls = [],
    transport = async (input) => {
      const result = await native(input);
      if (input.operation === 'spill-stage') calls.push({ input, result });
      return result;
    },
    make = () =>
      new IndexedWorkbook(
        transport,
        () => assert.fail('No legacy export fallback'),
        (input) => calculateIndexedWorkbook(input, transport)
      );
  return { root, source, bytes, hash: sha(bytes), calls, make };
}
function protectedMaster(cell) {
  assert.equal(cell.spill, undefined);
  assert.equal(cell.calculation.status, 'unresolved');
  assert.match(cell.calculation.error, /Array\/spill/);
  assert.notEqual(cell.calculation.value, 999);
}
async function unchanged(f) {
  assert.equal(sha(await fs.readFile(f.source)), f.hash);
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
}
async function checkSaved(f, output, masters) {
  const result = spawnSync(
    'python3',
    [
      '-c',
      `import json,sys,zipfile,xml.etree.ElementTree as E
ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with zipfile.ZipFile(sys.argv[1]) as a,zipfile.ZipFile(sys.argv[2]) as b:
 assert a.read('customXml/spill-boundary.xml')==b.read('customXml/spill-boundary.xml')
 cells={c.attrib['r']:c for c in E.fromstring(b.read('xl/worksheets/sheet1.xml')).findall('.//s:c',ns)}
 pr=E.fromstring(b.read('xl/workbook.xml')).find('s:calcPr',ns);forced=pr is not None and pr.attrib.get('fullCalcOnLoad')=='1' and pr.attrib.get('forceFullCalc')=='1'
 for ref,formula in json.loads(sys.argv[3]):
  c=cells[ref];f=c.find('s:f',ns);assert f is not None and f.text==formula
  assert not any(k in c.attrib for k in ['cm','vm']) and f.attrib.get('t')!='array'
  v=c.find('s:v',ns);assert v is None or v.text!='999' or forced
 assert 'xl/metadata.xml' not in b.namelist()
`,
      f.source,
      output.path,
      JSON.stringify(masters),
    ],
    { encoding: 'utf8' }
  );
  assert.equal(result.status, 0, result.stderr);
}
test('129 retained-vector masters open protected; 128 can qualify without altering old revisions or replay', async (t) => {
  const masters = Array.from({ length: 129 }, (_, i) => [`A${i * 4 + 1}`, 'SEQUENCE(3,2)']),
    f = await fixture(t, masters),
    book = f.make();
  t.after(() => book.close());
  await book.open(f.source);
  assert.equal(book.diagnosticStats().retainedStates, 1);
  const initial = await book.read('0', 'Sheet1', 'A1:G600');
  for (const [ref] of masters) protectedMaster(initial.cells[ref]);
  assert.equal(
    initial.cells.G1.calculation.status,
    'unresolved',
    'Unowned geometry protects reads from possible follower positions'
  );
  assert.equal(initial.cells.G2.calculation.status, 'unresolved');
  assert.equal(initial.cells.B1, undefined);
  assert.equal(f.calls[0].result.groups, 0);
  assert.equal(f.calls[0].result.refusedConstructors, 129);
  assert.match(f.calls[0].result.reason, /128-master/);
  // An unchanged export keeps every formula and source cache and tells Excel
  // to recalculate on open; the source file itself is never touched. Force a real saved candidate with
  // one unrelated literal edit before asserting unqualified-cache removal.
  await checkSaved(f, await book.export('0'), masters);
  await unchanged(f);
  const edited = await book.apply('0', [{ sheet: 'Sheet1', cell: 'F1', value: 8 }]);
  await checkSaved(f, await book.export(edited.view.storage.state), masters);
  const changed = await book.apply(edited.view.storage.state, [
      { sheet: 'Sheet1', cell: 'A513', value: null, expectedFormula: 'SEQUENCE(3,2)' },
    ]),
    state = changed.view.storage.state;
  const qualified = (await book.read(state, 'Sheet1', 'A1:B3')).cells;
  assert.equal(qualified.A1.spill.role, 'master');
  assert.equal(qualified.B3.value, 6);
  assert.equal(f.calls.at(-1).result.groups, 128);
  protectedMaster((await book.read('0', 'Sheet1', 'A1')).cells.A1);
  const before = book.diagnosticStats().retainedStates;
  await assert.rejects(book.apply(state, [{ sheet: 'Sheet1', cell: 'A513', formula: 'SEQUENCE(3,2)' }]), /128-master/);
  assert.equal(book.diagnosticStats().retainedStates, before);
  assert.equal((await book.read(state, 'Sheet1', 'B3')).cells.B3.spill.master, 'A1');
  assert.equal((await book.read(state, 'Sheet1', 'A513')).cells.A513.value, null);
  const replay = f.make();
  await replay.open(f.source);
  try {
    const copy = await replay.apply('0', changed.edits);
    assert.equal((await replay.read(copy.view.storage.state, 'Sheet1', 'B3')).cells.B3.value, 6);
  } finally {
    await replay.close();
  }
  await unchanged(f);
});
test('growth into an accepted owner refuses atomically without stale followers or history publication', async (t) => {
  const f = await fixture(
      t,
      [
        ['B1', 'SEQUENCE($J$1,2)'],
        ['A3', 'SEQUENCE(2,2)'],
      ],
      { J1: 2 }
    ),
    book = f.make();
  t.after(() => book.close());
  await book.open(f.source);
  const before = (await book.read('0', 'Sheet1', 'A1:J5')).cells;
  assert.equal(before.B1.spill.role, 'master');
  assert.equal(before.A3.spill.role, 'master');
  assert.equal(before.C2.value, 4);
  assert.equal(before.C3, undefined);
  await assert.rejects(book.apply('0', [{ sheet: 'Sheet1', cell: 'J1', value: 3 }]), { code: 'spill-unsupported' });
  assert.equal(book.diagnosticStats().retainedStates, 1);
  assert.deepEqual((await book.read('0', 'Sheet1', 'A1:J5')).cells, before);
  await unchanged(f);
  const changed = await book.apply('0', [{ sheet: 'Sheet1', cell: 'F1', value: 8 }]);
  assert.equal((await book.read(changed.view.storage.state, 'Sheet1', 'G1')).cells.G1.calculation.value, 9);
  assert.equal((await book.read(changed.view.storage.state, 'Sheet1', 'C2')).cells.C2.spill.master, 'B1');
});
test('intersecting retained-vector intended areas open unresolved and preserve source/export', async (t) => {
  const masters = [
      ['B1', 'SEQUENCE(3,2)'],
      ['A2', 'SEQUENCE(2,2)'],
    ],
    f = await fixture(t, masters),
    book = f.make();
  t.after(() => book.close());
  await book.open(f.source);
  const initial = (await book.read('0', 'Sheet1', 'A1:G5')).cells;
  for (const [ref] of masters) protectedMaster(initial[ref]);
  assert.equal(initial.G2.calculation.status, 'unresolved');
  assert.equal(initial.G1.calculation.status, 'unresolved');
  assert.equal(initial.B2, undefined);
  assert.equal(f.calls[0].result.groups, 0);
  assert.equal(f.calls[0].result.refusedConstructors, 2);
  assert.match(f.calls[0].result.reason, /intersecting/);
  const changed = await book.apply('0', [{ sheet: 'Sheet1', cell: 'F1', value: 8 }]);
  assert.equal((await book.read(changed.view.storage.state, 'Sheet1', 'F1')).cells.F1.value, 8);
  for (const [ref] of masters) protectedMaster((await book.read(changed.view.storage.state, 'Sheet1', ref)).cells[ref]);
  assert.equal((await book.read('0', 'Sheet1', 'F1')).cells.F1.value, 5);
  await checkSaved(f, await book.export(changed.view.storage.state), masters);
  await unchanged(f);
  const reopened = f.make();
  await reopened.open((await book.export(changed.view.storage.state)).path);
  try {
    for (const [ref] of masters) protectedMaster((await reopened.read('0', 'Sheet1', ref)).cells[ref]);
  } finally {
    await reopened.close();
  }
});
