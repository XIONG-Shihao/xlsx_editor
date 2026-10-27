/** Actual table-local insert/delete owner, source facts and recovery qualification. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { unzipSync, strFromU8 } from 'fflate';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { tableColumnStructureFixture } from './support/table-column-structure-fixture.mjs';
after(() => closeNativeWorkbookWorker());
const insert = (index, name = 'Tax', ref = 'A1:C4') => ({
    operation: 'insertTableColumn',
    sheet: 'Sheet1',
    tableName: 'Prices',
    expectedTableRef: ref,
    index,
    newName: name,
  }),
  remove = (columnName, ref = 'A1:C4') => ({
    operation: 'deleteTableColumn',
    sheet: 'Sheet1',
    tableName: 'Prices',
    expectedTableRef: ref,
    columnName,
  });
async function owner(t, options = {}, fault = null) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-table-columns-')),
    source = path.join(root, 'source.xlsx'),
    bytes = tableColumnStructureFixture(options);
  await fs.writeFile(source, bytes);
  let activeFault = fault;
  const bridge = async (input) => {
    if (activeFault && input.operation === 'artifact-write' && input.name === 'package-metadata.json') {
      const p = JSON.parse(input.data);
      activeFault(p);
      input = { ...input, data: JSON.stringify(p) };
    }
    return native(input);
  };
  const make = () =>
    new IndexedWorkbook(
      bridge,
      () => assert.fail('No shared structural writer'),
      (input) => calculateIndexedWorkbook(input, bridge)
    );
  const book = make(),
    owners = [book];
  t.after(async () => {
    for (const o of owners) await o.close();
    await fs.rm(root, { recursive: true, force: true });
  });
  await book.open(source);
  return {
    root,
    source,
    bytes,
    book,
    make: () => {
      const o = make();
      owners.push(o);
      return o;
    },
    fault: (f) => (activeFault = f),
  };
}
const read = (book, state, sheet = 'Sheet1', range = 'A1:H8') =>
  book.read(state, sheet, range).then((reply) => reply.cells);
test('insert first/middle/end shifts only table cells, preserves styles/formulas and updates direct range references', async (t) => {
  for (const index of [0, 1, 3]) {
    const f = await owner(t),
      before = await read(f.book, '0'),
      command = insert(index),
      result = await f.book.apply('0', [command]),
      state = result.view.storage.state,
      c = await read(f.book, state);
    const labels = ['Value', 'Cost', 'Total'];
    labels.splice(index, 0, 'Tax');
    assert.deepEqual(
      result.view.tables[0].columns.map((c) => c.name),
      labels
    );
    assert.equal(result.view.tables[0].ref, 'A1:D4');
    assert.equal(c.A8.value, 77);
    assert.equal(c.H1.calculation.value, 9);
    assert.equal(c.H2.calculation.value, 60);
    assert.equal(c.H3.calculation.value, 18);
    assert.equal(c.H4.calculation.value, 61);
    assert.equal(c.H5.calculation.value, 69);
    assert.equal(c.H5.formula, ['SUM(B2:C4)', 'SUM(A2:C4)', 'SUM(A2:B4)'][[0, 1, 3].indexOf(index)]);
    assert.equal(c.H6.formula, '"Prices[Cost]"');
    assert.equal((await read(f.book, state, 'Other', 'A1')).A1.calculation.value, 60);
    for (let col = 0; col < 3; col++)
      for (let row = 1; row <= 4; row++) {
        const a = 'ABC'[col] + row,
          b = 'ABCD'[col + (col >= index ? 1 : 0)] + row;
        for (const key of [
          'formula',
          'styleId',
          'fontColor',
          'bold',
          'format',
          ...(before[a]?.formula ? [] : ['value']),
        ])
          assert.deepEqual(c[b]?.[key], before[a]?.[key], a + ' -> ' + b + '.' + key);
        if (before[a]?.formula) assert.equal(c[b].calculation.value, before[a].calculation.value);
      }
    assert.equal(c['ABCD'[index] + '1'].value, 'Tax');
    assert.equal(c['ABCD'[index] + '2']?.value ?? null, null);
    assert.deepEqual(result.commands, [command]);
    assert(JSON.stringify(result.commands).length < 256);
    assert.deepEqual(await read(f.book, '0'), before);
    assert.deepEqual(await fs.readFile(f.source), f.bytes);
    const output = await f.book.export(state),
      cold = f.make();
    await cold.open(output.path);
    assert.equal((await read(cold, '0')).H5.calculation.value, 69);
    const parts = unzipSync(await fs.readFile(output.path));
    assert.equal(strFromU8(parts['customXml/table-opaque.xml']), '<opaque> untouched </opaque>');
  }
});
test('delete first/middle/end removes selected data, shifts survivors and updates ordinary/deleted structured references', async (t) => {
  for (const columnName of ['Value', 'Cost', 'Total']) {
    const f = await owner(t),
      before = await read(f.book, '0'),
      result = await f.book.apply('0', [remove(columnName)]),
      state = result.view.storage.state,
      c = await read(f.book, state),
      index = ['Value', 'Cost', 'Total'].indexOf(columnName);
    assert.equal(result.view.tables[0].ref, 'A1:B4');
    assert.deepEqual(
      result.view.tables[0].columns.map((c) => c.name),
      ['Value', 'Cost', 'Total'].filter((n) => n !== columnName)
    );
    assert.equal(c.A8.value, 77);
    assert.equal(c.C1, undefined);
    assert.equal(c.C2, undefined);
    assert.equal(c.H6.formula, '"Prices[Cost]"');
    assert.equal(c['H' + (index + 1)].formula, 'SUM(#REF!)');
    assert.equal(c['H' + (index + 1)].calculation.error, '#REF!');
    assert.equal(c.H5.calculation.value, [60, 9, 69][index]);
    for (let col = 0; col < 3; col++)
      if (col !== index)
        for (let row = 1; row <= 4; row++) {
          const a = 'ABC'[col] + row,
            b = 'AB'[col - (col > index ? 1 : 0)] + row;
          assert.deepEqual(c[b]?.styleId, before[a]?.styleId);
          if (before[a]?.formula) {
            if (columnName === 'Value') {
              assert.equal(c[b].calculation.status, 'error');
              assert.equal(c[b].calculation.error, '#REF!');
            } else {
              assert.equal(c[b].formula, before[a].formula);
              assert.equal(c[b].calculation.value, before[a].calculation.value);
            }
          } else assert.deepEqual(c[b]?.value, before[a]?.value);
        }
    const output = await f.book.export(state),
      cold = f.make();
    await cold.open(output.path);
    assert.equal((await read(cold, '0'))['H' + (index + 1)].calculation.error, '#REF!');
    assert.deepEqual(await read(f.book, '0'), before);
    assert.deepEqual(await fs.readFile(f.source), f.bytes);
  }
});
test('sparse rows retain their row metadata without ghost headers during append', async (t) => {
  const f = await owner(t, { sparse: true }),
    result = await f.book.apply('0', [insert(3)]),
    state = result.view.storage.state,
    c = await read(f.book, state);
  assert.equal(c.D1.value, 'Tax');
  assert.equal(c.D2, undefined);
  assert.equal(c.D4, undefined);
  assert.equal(c.B4.styleId, 0);
  const output = await f.book.export(state),
    check = spawnSync(
      'python3',
      [
        '-c',
        `import zipfile,sys,xml.etree.ElementTree as E
n='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
with zipfile.ZipFile(sys.argv[1]) as z:
 s=E.fromstring(z.read('xl/worksheets/sheet1.xml'));rows={r.get('r'):r for r in s.iter(n+'row')};assert rows['2'].attrib['ht']=='20' and len(list(rows['2']))==0
 assert [c.get('r') for c in rows['1']]==['A1','B1','C1','D1']
`,
        output.path,
      ],
      { encoding: 'utf8' }
    );
  assert.equal(check.status, 0, check.stderr);
});
test('occupied/crossing/label/sole-column/stale and forged definition changes refuse atomically', async (t) => {
  for (const [options, command, reason] of [
    [{ occupied: true }, insert(1), /absorb/],
    [{ crossing: true }, insert(1), /boundary/],
    [{}, insert(1, 'cost'), /collision/],
    [{}, insert(4), /bounds/],
    [{}, remove('Missing'), /changed|bounds/],
  ]) {
    const f = await owner(t, options),
      before = f.book.view('0'),
      cells = await read(f.book, '0');
    await assert.rejects(f.book.apply('0', [command]), reason);
    assert.deepEqual(f.book.view('0'), before);
    assert.deepEqual(await read(f.book, '0'), cells);
    assert.deepEqual(await fs.readFile(f.source), f.bytes);
  }
  const f = await owner(t);
  for (const fault of [
    (p) => (p.columnChange.index = 0),
    (p) =>
      (p.replacements['xl/tables/table1.xml'] = p.replacements['xl/tables/table1.xml'].replace(
        'TableStyleMedium2',
        'TableStyleMedium9'
      )),
    (p) => (p.sheets = []),
  ]) {
    f.fault(fault);
    await assert.rejects(f.book.apply('0', [insert(1)]));
    assert.equal(f.book.view('0').tables[0].ref, 'A1:C4');
  }
  f.fault(null);
  const one = await f.book.apply('0', [remove('Cost'), remove('Total', 'A1:B4')]);
  await assert.rejects(f.book.apply(one.view.storage.state, [remove('Value', 'A1:A4')]), /at least one/);
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
});
test('one Document action Undo/Redo, caller capture, durable replay and Manual pending policy survive column structure', async (t) => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-columns-document-')),
    source = path.join(root, 'source.xlsx'),
    bytes = tableColumnStructureFixture();
  await fs.writeFile(source, bytes);
  const recovery = path.join(root, 'recovery'),
    owners = [];
  t.after(async () => {
    for (const o of owners) await o.close();
    await fs.rm(root, { recursive: true, force: true });
  });
  const o = new SpreadsheetDocument({ recoveryDirectory: recovery });
  owners.push(o);
  await o.open(source, { fresh: true });
  const req = (o, operation, extra = {}) =>
    o.request({ documentId: o.status().documentId, state: o.status().state, operation, ...extra });
  const cells = async (o) => (await req(o, 'read', { sheet: 'Sheet1', range: 'A1:H8' })).cells;
  const before = await cells(o),
    command = insert(1),
    work = req(o, 'apply', { commands: [command] });
  command.index = 0;
  command.newName = 'Late caller';
  await work;
  assert.equal((await cells(o)).B1.value, 'Tax');
  await req(o, 'undo');
  assert.deepEqual(await cells(o), before);
  await req(o, 'redo');
  assert.equal((await cells(o)).C2.value, 10);
  const id = o.status().documentId;
  await o.close();
  const recovered = new SpreadsheetDocument({ recoveryDirectory: recovery });
  owners.push(recovered);
  await recovered.recover(id);
  assert.equal((await cells(recovered)).B1.value, 'Tax');
  await req(recovered, 'undo');
  assert.deepEqual(await cells(recovered), before);
  await req(recovered, 'redo');
  await req(recovered, 'apply', {
    commands: [{ type: 'setCalculationOptions', mode: 'manual', recalculateBeforeSave: false }],
  });
  await req(recovered, 'apply', { commands: [{ sheet: 'Sheet1', cell: 'A2', value: 20 }] });
  await req(recovered, 'apply', { commands: [remove('Cost', 'A1:D4')] });
  const current = await cells(recovered);
  assert.equal(current.B1.value, 'Tax');
  assert.equal(current.H2.formula, 'SUM(#REF!)');
  assert(current.H2.calculation.pending);
  assert.notEqual(current.H2.calculation.status, 'calculated');
  const saved = path.join(root, 'saved.xlsx');
  await req(recovered, 'copy', { destination: saved, expectedRevision: null });
  await recovered.close();
  const cold = new SpreadsheetDocument({ recoveryDirectory: path.join(root, 'cold') });
  owners.push(cold);
  await cold.open(saved, { fresh: true });
  assert.equal(cold.status().canUndo, false);
  assert.equal(cold.status().view.calculationPolicy.mode, 'manual');
  assert((await cells(cold)).H2.calculation.pending);
  assert.deepEqual(await fs.readFile(source), bytes);
});
