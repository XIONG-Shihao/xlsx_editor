/** N2: exact retained scalar spill consumers, immutable owners and cache fencing. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { zipSync, unzipSync, strToU8 } from 'fflate';
import { createWorkbook, readWorkbook } from '../core/workbook-package.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  sheet = 'xl/worksheets/sheet1.xml';
const sha = (bytes) => createHash('sha256').update(bytes).digest('hex');
async function owner(t, bytes) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-spill-consumer-')),
    source = path.join(directory, 'source.xlsx');
  await fs.writeFile(source, bytes);
  let root, version;
  const calls = [],
    stats = [];
  const bridge = async (input) => {
    const result = await native(input);
    calls.push({ input, result });
    if (input.operation === 'index') root = input.root;
    if (input.operation === 'range') version = input.version;
    return result;
  };
  const make = () =>
    new IndexedWorkbook(
      bridge,
      () => assert.fail('No legacy writer'),
      async (input) => {
        const result = await calculateIndexedWorkbook(input, bridge);
        stats.push(result);
        return result;
      }
    );
  const book = make();
  t.after(async () => {
    await book.close();
    await fs.rm(directory, { recursive: true, force: true });
  });
  await book.open(source);
  return { book, make, source, directory, bytes, root: () => root, version: () => version, calls, stats };
}
function authored(formula = 'SUM(A1#)') {
  const parts = unzipSync(createWorkbook()),
    xml = (value) => String(value).replaceAll('&', '&amp;').replaceAll('<', '&lt;');
  const f = (ref, formula) => `<c r="${ref}"><f>${xml(formula)}</f><v>999</v></c>`;
  parts[sheet] = strToU8(
    `<worksheet xmlns="${ns}"><sheetData><row r="1">${f('A1', 'SEQUENCE($J$1,2)')}${f('D1', formula)}${f('E1', 'D1+1')}${f('F1', 'DEC2HEX(D1,4)')}<c r="J1"><v>3</v></c></row><row r="2">${f('D2', 'INDEX(_xlfn.ANCHORARRAY($A$1),3,2)')}${f('E2', 'IFERROR(D2,99)')}</row><row r="3">${f('D3', 'SUM(ANCHORARRAY($A$1))')}</row><row r="4">${f('D4', 'INDEX(A1#,2,2)')}</row></sheetData></worksheet>`
  );
  parts['customXml/n2-opaque.xml'] = strToU8('<opaque> preserved </opaque>');
  return Buffer.from(zipSync(parts));
}
function scalar(cell, value) {
  assert.equal(cell.calculation.status, 'calculated');
  assert.equal(cell.calculation.value, value);
  assert.equal(cell.calculation.spillConsumer.schema, 1);
  assert.equal(cell.spill, undefined);
}
async function preserved(f) {
  assert.equal(sha(await fs.readFile(f.source)), sha(f.bytes));
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
}
function rawAudit(source, saved, expected) {
  const result = spawnSync(
    'python3',
    [
      '-c',
      `import json,sys,zipfile,xml.etree.ElementTree as E
s='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
with zipfile.ZipFile(sys.argv[1]) as a,zipfile.ZipFile(sys.argv[2]) as b:
 before={c.attrib['r']:c for c in E.fromstring(a.read('xl/worksheets/sheet1.xml')).iter(s+'c')}
 after={c.attrib['r']:c for c in E.fromstring(b.read('xl/worksheets/sheet1.xml')).iter(s+'c')}
 for ref,item in json.loads(sys.argv[3]).items():
  c=after[ref];f=c.find(s+'f');v=c.find(s+'v');assert f is not None and v is not None and float(v.text)==item['value'],ref
  assert c.attrib.get('t','n')=='n' and f.text==item['formula'],ref
  old=before.get(ref)
  if item.get('singleton'):
   assert f.attrib==old.find(s+'f').attrib=={'t':'array','ref':ref} and c.attrib.get('cm')==old.attrib['cm'] and 'vm' not in c.attrib,ref
  else: assert f.attrib=={} and 'cm' not in c.attrib and 'vm' not in c.attrib,ref
 if 'customXml/n2-opaque.xml' in a.namelist(): assert a.read('customXml/n2-opaque.xml')==b.read('customXml/n2-opaque.xml')
`,
      source,
      saved,
      JSON.stringify(expected),
    ],
    { encoding: 'utf8' }
  );
  assert.equal(result.status, 0, result.stderr);
}
for (const name of ['excel-dynamic-source.xlsx', 'excel-dynamic-reference.xlsx'])
  test(`actual retained Excel21/6/52 singleton consumers qualify and preserve source XML: ${name}`, async (t) => {
    const oracle = JSON.parse(
        await fs.readFile(new URL('../fixtures/excel-dynamic-reference.json', import.meta.url), 'utf8')
      ),
      bytes = await fs.readFile(new URL('../fixtures/' + name, import.meta.url));
    assert.equal(sha(bytes), name === oracle.source ? oracle.sourceSha256 : oracle.oracleWorkbookSha256);
    const f = await owner(t, bytes),
      cells = (await f.book.read('0', 'Cases', 'A1:G80')).cells,
      expected = {};
    for (const ref of ['D1', 'D2', 'D6']) {
      scalar(cells[ref], oracle.cells[ref].expected);
      assert.equal(cells[ref].formulaType, 'array');
      assert.equal(cells[ref].formulaReference, ref);
      expected[ref] = { value: oracle.cells[ref].expected, formula: oracle.authoredFormulas[ref], singleton: true };
    }
    assert.equal(cells.D18.calculation.status, 'unresolved');
    assert.equal(cells.D24.calculation.status, 'unresolved');
    assert.equal(f.stats[0].spillConsumerCells, 3);
    assert.equal(f.stats[0].nativePartialCells, 0);
    const output = await f.book.export('0');
    rawAudit(f.source, output.path, expected);
    await preserved(f);
    const replay = f.make();
    await replay.open(output.path);
    try {
      const again = (await replay.read('0', 'Cases', 'D1:D6')).cells;
      for (const ref of ['D1', 'D2', 'D6']) scalar(again[ref], expected[ref].value);
    } finally {
      await replay.close();
    }
  });
test('authored #/ANCHORARRAY aliases feed scalar and fallback dependents through resize/release/replay', async (t) => {
  const f = await owner(t, authored()),
    initial = (await f.book.read('0', 'Sheet1', 'A1:J5')).cells;
  for (const ref of ['D1', 'D3']) scalar(initial[ref], 21);
  scalar(initial.D2, 6);
  scalar(initial.D4, 4);
  assert.equal(initial.E1.calculation.value, 22);
  assert.equal(initial.E2.calculation.value, 6);
  assert.equal(initial.F1.calculation.value, '0015');
  assert.equal(f.stats[0].nativeEvaluator, false);
  const shrink = await f.book.apply('0', [{ sheet: 'Sheet1', cell: 'J1', value: 2 }]),
    state = shrink.view.storage.state,
    small = (await f.book.read(state, 'Sheet1', 'A1:J5')).cells;
  scalar(small.D1, 10);
  scalar(small.D3, 10);
  scalar(small.D4, 4);
  assert.equal(small.D2.calculation.status, 'unresolved');
  assert.equal(small.E2.calculation.status, 'unresolved');
  assert.equal(small.E1.calculation.value, 11);
  assert.equal(small.F1.calculation.value, '000A');
  assert.equal(small.B3.value, null);
  const grow = await f.book.apply(state, [{ sheet: 'Sheet1', cell: 'J1', value: 3 }]),
    grown = (await f.book.read(grow.view.storage.state, 'Sheet1', 'D1:F4')).cells;
  scalar(grown.D1, 21);
  scalar(grown.D2, 6);
  assert.equal(grown.E2.calculation.value, 6);
  scalar((await f.book.read('0', 'Sheet1', 'D2')).cells.D2, 6);
  const output = await f.book.export(grow.view.storage.state);
  rawAudit(f.source, output.path, {
    D1: { value: 21, formula: 'SUM(_xlfn.ANCHORARRAY(A1))' },
    D2: { value: 6, formula: 'INDEX(_xlfn.ANCHORARRAY($A$1),3,2)' },
    D3: { value: 21, formula: 'SUM(_xlfn.ANCHORARRAY($A$1))' },
    D4: { value: 4, formula: 'INDEX(_xlfn.ANCHORARRAY(A1),2,2)' },
  });
  const replay = f.make();
  await replay.open(f.source);
  try {
    const restored = await replay.apply('0', [...shrink.edits, ...grow.edits]);
    scalar((await replay.read(restored.view.storage.state, 'Sheet1', 'D1')).cells.D1, 21);
  } finally {
    await replay.close();
  }
  await f.book.compact(['0', grow.view.storage.state]);
  scalar((await f.book.read(grow.view.storage.state, 'Sheet1', 'D1')).cells.D1, 21);
  await f.book.export(grow.view.storage.state);
  const cleared = await f.book.apply(grow.view.storage.state, [{ sheet: 'Sheet1', cell: 'A1', value: null }]),
    empty = (await f.book.read(cleared.view.storage.state, 'Sheet1', 'A1:F5')).cells;
  for (const ref of ['D1', 'D2', 'D3', 'D4', 'E1', 'E2', 'F1'])
    assert.equal(empty[ref].calculation.status, 'unresolved', ref);
  assert.equal(empty.B2.value, null);
  scalar((await f.book.read(grow.view.storage.state, 'Sheet1', 'D1')).cells.D1, 21);
  await preserved(f);
});
test('unowned/follower/blocked anchors and broader syntax stay explicitly protected', async (t) => {
  for (const formula of [
    'SUM(B1#)',
    'SUM(C1#)',
    'SUM(Sheet1!A1#)',
    'SUM(Name#)',
    'SUM(A1#,2)',
    'ABS(A1#)',
    'IF(TRUE,SUM(A1#),0)',
    'SUM(A1#)+1',
    'INDEX(A1#,0,1)',
    'INDEX(A1#,1)',
    'INDEX(A1#,4,2)',
    'INDEX(A1#,J1,2)',
    'INDEX(A1#,2.0,2)',
  ])
    await t.test(formula, async (sub) => {
      const f = await owner(sub, authored(formula)),
        cell = (await f.book.read('0', 'Sheet1', 'D1')).cells.D1;
      assert.equal(cell.calculation.status, 'unresolved');
      assert.equal(cell.calculation.spillConsumer, undefined);
      await preserved(f);
    });
  const parts = unzipSync(authored());
  parts[sheet] = strToU8(
    new TextDecoder().decode(parts[sheet]).replace('<row r="2">', '<row r="2"><c r="B2"><v>0</v></c>')
  );
  const f = await owner(t, Buffer.from(zipSync(parts))),
    cells = (await f.book.read('0', 'Sheet1', 'A1:E2')).cells;
  assert.equal(cells.A1.calculation.error, '#SPILL!');
  assert.equal(cells.D1.calculation.status, 'unresolved');
  assert.equal(cells.D2.calculation.status, 'unresolved');
  await preserved(f);
});
test('independent export refuses missing/forged/stale proofs while retaining accepted output and source', async (t) => {
  const f = await owner(t, authored()),
    cells = (await f.book.read('0', 'Sheet1', 'A1:J5')).cells,
    parent = f.version(),
    view = f.book.view('0'),
    output = await f.book.export('0'),
    accepted = await fs.readFile(output.path);
  for (const mutate of [
    (cell) => {
      cell.calculation.value = 999;
    },
    (cell) => {
      delete cell.calculation.spillConsumer;
    },
    (cell) => {
      cell.calculation.spillConsumer.range = 'A1:C3';
    },
  ]) {
    const cell = structuredClone(cells.D1);
    mutate(cell);
    const changed = await native({
      action: 'workbook',
      operation: 'patch',
      root: f.root(),
      version: parent,
      changes: [{ sheet, cell: 'D1', value: cell }],
    });
    await assert.rejects(
      native({ action: 'workbook', operation: 'export-plan', root: f.root(), version: changed.version, view }),
      /spill|Spill|ownership|proof/
    );
    assert.deepEqual(await fs.readFile(output.path), accepted);
    await preserved(f);
  }
  const changed = await native({
    action: 'workbook',
    operation: 'patch',
    root: f.root(),
    version: parent,
    changes: [{ sheet, cell: 'J1', value: { value: 2, dataType: 'number' } }],
  });
  await assert.rejects(
    native({ action: 'workbook', operation: 'export-plan', root: f.root(), version: changed.version, view }),
    /stale/
  );
  assert.deepEqual(await fs.readFile(output.path), accepted);
  scalar((await f.book.read('0', 'Sheet1', 'D1')).cells.D1, 21);
  await preserved(f);
});
test('stream export revalidates mutable prepared scratch and requires complete spill metadata', async (t) => {
  const f = await owner(t, authored());
  await f.book.read('0', 'Sheet1', 'D1');
  const version = f.version(),
    revision = f.book.view('0').revision,
    output = await f.book.export('0'),
    accepted = await fs.readFile(output.path),
    metadataPath = path.join(f.root(), 'export-metadata.json'),
    metadata = await fs.readFile(metadataPath, 'utf8');
  const altered = spawnSync(
    'python3',
    [
      '-c',
      `import sqlite3,json,sys
db=sqlite3.connect(sys.argv[1]);part='xl/worksheets/sheet1.xml';v=json.loads(db.execute('SELECT data FROM export_changes WHERE sheet=? AND r=0 AND c=3',(part,)).fetchone()[0]);v['calculation']['value']=999;db.execute('UPDATE export_changes SET data=? WHERE sheet=? AND r=0 AND c=3',(json.dumps(v),part));db.commit()
`,
      path.join(f.root(), 'cells.sqlite'),
    ],
    { encoding: 'utf8' }
  );
  assert.equal(altered.status, 0, altered.stderr);
  await assert.rejects(
    native({ action: 'workbook', operation: 'export', root: f.root(), version, revision, calculationMetadata: true }),
    { code: /^(internal-mismatch|spill-bad-input)$/ }
  );
  assert.deepEqual(await fs.readFile(output.path), accepted);
  await f.book.export('0');
  const removed = JSON.parse(metadata);
  delete removed.spillMetadata;
  await native({
    action: 'workbook',
    operation: 'artifact-write',
    root: f.root(),
    name: 'export-metadata.json',
    data: JSON.stringify(removed),
  });
  await assert.rejects(
    native({ action: 'workbook', operation: 'export', root: f.root(), version, revision, calculationMetadata: true }),
    { code: /^(internal-missing-field|spill-unsupported)$/ }
  );
  assert.deepEqual(await fs.readFile(output.path), accepted);
  await preserved(f);
});
test('imported singleton metadata refuses content clear while ordinary authored consumers clear coherently', async (t) => {
  const bytes = await fs.readFile(new URL('../fixtures/excel-dynamic-reference.xlsx', import.meta.url)),
    f = await owner(t, bytes),
    before = (await f.book.read('0', 'Cases', 'D1:D6')).cells;
  for (const range of ['D1', 'D1:D2', 'C1:D1']) {
    await assert.rejects(f.book.apply('0', [{ type: 'clearRange', sheet: 'Cases', range }]), {
      code: 'fill-unsupported',
    });
    assert.equal(f.book.diagnosticStats().retainedStates, 1);
    assert.deepEqual((await f.book.read('0', 'Cases', 'D1:D6')).cells, before);
    await preserved(f);
  }
  const output = await f.book.export('0');
  rawAudit(f.source, output.path, {
    D1: { value: 21, formula: 'SUM(_xlfn.ANCHORARRAY(A1))', singleton: true },
    D2: { value: 6, formula: 'INDEX(_xlfn.ANCHORARRAY(A1),3,2)', singleton: true },
    D6: { value: 52, formula: 'SUM(_xlfn.ANCHORARRAY(A6))', singleton: true },
  });
  const ordinary = await owner(t, authored()),
    cleared = await ordinary.book.apply('0', [{ type: 'clearRange', sheet: 'Sheet1', range: 'D1' }]),
    cells = (await ordinary.book.read(cleared.view.storage.state, 'Sheet1', 'D1:E1')).cells;
  assert.equal(cells.D1.value, null);
  assert.equal(cells.D1.formula, undefined);
  assert.equal(cells.E1.calculation.value, 1);
  scalar((await ordinary.book.read('0', 'Sheet1', 'D1')).cells.D1, 21);
  await preserved(ordinary);
});
test('consumer proof pages remain bounded at64 without imposing a workbook consumer-count limit', async (t) => {
  const parts = unzipSync(createWorkbook()),
    rows = Array.from(
      { length: 65 },
      (_, i) =>
        `<row r="${i + 1}">${i === 0 ? '<c r="A1"><f>SEQUENCE(3,2)</f></c>' : ''}<c r="D${i + 1}"><f>SUM($A$1#)</f><v>999</v></c></row>`
    ).join('');
  parts[sheet] = strToU8(`<worksheet xmlns="${ns}"><sheetData>${rows}</sheetData></worksheet>`);
  const f = await owner(t, Buffer.from(zipSync(parts))),
    cells = (await f.book.read('0', 'Sheet1', 'D1:D65')).cells;
  for (let i = 1; i <= 65; i++) scalar(cells['D' + i], 21);
  const pages = f.calls.filter((call) => call.input.operation === 'spill-consumers');
  assert.deepEqual(
    pages.map((call) => call.input.items.length),
    [64, 1]
  );
  assert.equal(f.stats[0].spillConsumerCells, 65);
  const version = f.version(),
    view = f.book.view('0');
  await assert.rejects(
    native({
      action: 'workbook',
      operation: 'spill-consumers',
      root: f.root(),
      version,
      view,
      items: Array.from({ length: 65 }, (_, i) => ({ sheet, ref: 'D' + (i + 1) })),
    }),
    { code: 'spill-limit' }
  );
  await assert.rejects(
    native({
      action: 'workbook',
      operation: 'spill-consumers',
      root: f.root(),
      version,
      view,
      items: [
        { sheet, ref: 'D1' },
        { sheet, ref: 'D1' },
      ],
    }),
    /duplicate/
  );
  scalar((await f.book.read('0', 'Sheet1', 'D65')).cells.D65, 21);
  await preserved(f);
});
test('missing/noncanonical cm and rich metadata cannot borrow singleton dynamic ownership', async (t) => {
  const bytes = await fs.readFile(new URL('../fixtures/excel-dynamic-source.xlsx', import.meta.url));
  for (const change of [
    (raw) => raw.replace(' cm="1"', ''),
    (raw) => raw.replace(' cm="1"', ' cm="0"'),
    (raw) => raw.replace(' cm="1"', ' cm="1" vm="1"'),
  ])
    await t.test(String(change), async (sub) => {
      const parts = unzipSync(bytes);
      parts[sheet] = strToU8(new TextDecoder().decode(parts[sheet]).replace(/<c r="D1"[^>]*>/, change));
      const f = await owner(sub, Buffer.from(zipSync(parts))),
        cells = (await f.book.read('0', 'Cases', 'D1:D6')).cells;
      assert.equal(cells.D1.calculation.status, 'unresolved');
      assert.equal(cells.D1.calculation.spillConsumer, undefined);
      scalar(cells.D2, 6);
      scalar(cells.D6, 52);
      await preserved(f);
    });
});
