import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook } from '../core/workbook-package.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { defaultPrintSettings } from '../core/print-model.mjs';
after(() => closeNativeWorkbookWorker());
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
async function fixture(t, { globalUse = true, dynamic = false } = {}) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'connected-names-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const parts = unzipSync(createWorkbook());
  parts['xl/workbook.xml'] = strToU8(
    strFromU8(parts['xl/workbook.xml'])
      .replace('</sheets>', '<sheet name="Other" sheetId="2" r:id="rId3"/></sheets>')
      .replace(
        '</workbook>',
        '<definedNames><definedName name="Inputs">\'Sheet1\'!$A$1:$A$1</definedName><definedName name="Inputs" localSheetId="0">\'Sheet1\'!$B$1:$B$1</definedName><definedName name="Unused">\'Sheet1\'!$A$2:$A$2</definedName></definedNames></workbook>'
      )
  );
  parts['xl/_rels/workbook.xml.rels'] = strToU8(
    strFromU8(parts['xl/_rels/workbook.xml.rels']).replace(
      '</Relationships>',
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>'
    )
  );
  parts['[Content_Types].xml'] = strToU8(
    strFromU8(parts['[Content_Types].xml']).replace(
      '</Types>',
      '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
    )
  );
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    '<worksheet xmlns="' +
      ns +
      '"><dimension ref="A1:D3"/><sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1"><v>3</v></c><c r="C1"><f>SUM(Inputs)</f><v>999</v></c><c r="D1"><f>' +
      (dynamic ? 'INDIRECT("Inputs")' : 'LEN("Inputs")') +
      '</f><v>999</v></c></row><row r="2"><c r="A2"><v>7</v></c><c r="B2"><v>5</v></c></row></sheetData></worksheet>'
  );
  parts['xl/worksheets/sheet2.xml'] = strToU8(
    '<worksheet xmlns="' +
      ns +
      '"><dimension ref="A1:C2"/><sheetData><row r="1"><c r="A1"><v>11</v></c><c r="B1"><f>' +
      (globalUse ? 'SUM(Inputs)' : 'Sheet1!A1+1') +
      '</f><v>999</v></c><c r="C1"><f>Sheet1!Inputs</f><v>999</v></c></row></sheetData></worksheet>'
  );
  parts['custom/unchanged.bin'] = Uint8Array.from([0, 17, 255]);
  const source = path.join(directory, 'source.xlsx');
  await fs.writeFile(source, zipSync(parts));
  const original = await fs.readFile(source);
  let sequence = 0;
  const bridge = async (input) => {
    return native(input);
  };
  const book = new IndexedWorkbook(
    bridge,
    () => assert.fail('No compatibility writer'),
    (input) => calculateIndexedWorkbook(input, bridge)
  );
  await book.open(source);
  t.after(() => book.close());
  return { book, source, original, directory, parts };
}
const update = (scope, old, range) => ({
  operation: 'updateName',
  sheet: 'Sheet1',
  name: 'Inputs',
  nameScope: scope,
  expectedNameFormula: old,
  range,
});
const remove = (name, scope, old) => ({
  operation: 'removeName',
  sheet: 'Sheet1',
  name,
  nameScope: scope,
  expectedNameFormula: old,
});
async function value(book, state, sheet, ref) {
  const c = (await book.read(state, sheet, ref)).cells[ref];
  assert.equal(c.calculation.status, 'calculated', JSON.stringify(c));
  return c.calculation.value;
}
test('ordinary named formula consumers follow global/local range edits without rewriting formulas', async (t) => {
  const f = await fixture(t),
    b = f.book;
  assert.equal(await value(b, '0', 'Sheet1', 'C1'), 3);
  assert.equal(await value(b, '0', 'Other', 'B1'), 2);
  let made = await b.apply('0', [update('workbook', "'Sheet1'!$A$1:$A$1", 'A2')]),
    state = made.view.storage.state;
  assert.equal(await value(b, state, 'Other', 'B1'), 7);
  assert.equal(await value(b, state, 'Sheet1', 'C1'), 3);
  assert.equal(await value(b, '0', 'Other', 'B1'), 2);
  made = await b.apply(state, [update('worksheet', "'Sheet1'!$B$1:$B$1", 'B2')]);
  state = made.view.storage.state;
  assert.equal(await value(b, state, 'Sheet1', 'C1'), 5);
  assert.equal(await value(b, state, 'Other', 'C1'), 5);
  assert.equal(await value(b, state, 'Other', 'B1'), 7);
  assert.equal((await b.read(state, 'Sheet1', 'C1')).cells.C1.formula, 'SUM(Inputs)');
  assert.deepEqual(await fs.readFile(f.source), f.original);
});
test('unused names can be removed from formula workbooks, used global/local names cannot', async (t) => {
  const f = await fixture(t),
    b = f.book,
    removed = await b.apply('0', [remove('Unused', 'workbook', "'Sheet1'!$A$2:$A$2")]),
    state = removed.view.storage.state;
  assert.equal(removed.view.namedRanges.length, 2);
  assert.equal(await value(b, state, 'Sheet1', 'C1'), 3);
  assert.equal(await value(b, state, 'Other', 'B1'), 2);
  const count = b.diagnosticStats().retainedStates;
  await assert.rejects(b.apply(state, [remove('Inputs', 'workbook', "'Sheet1'!$A$1:$A$1")]), {
    code: 'name-unsupported',
  });
  await assert.rejects(b.apply(state, [remove('Inputs', 'worksheet', "'Sheet1'!$B$1:$B$1")]), {
    code: 'name-unsupported',
  });
  assert.equal(b.diagnosticStats().retainedStates, count);
});
test('local shadow and explicit local references do not make an unused workbook name used', async (t) => {
  const f = await fixture(t, { globalUse: false }),
    b = f.book,
    removed = await b.apply('0', [remove('Inputs', 'workbook', "'Sheet1'!$A$1:$A$1")]),
    state = removed.view.storage.state;
  assert.equal(removed.view.namedRanges.filter((n) => n.name === 'Inputs').length, 1);
  assert.equal(await value(b, state, 'Sheet1', 'C1'), 3);
  assert.equal(await value(b, state, 'Other', 'C1'), 3);
  assert.equal(await value(b, state, 'Other', 'B1'), 3);
});
test('print area/titles remain exact neighbors of edited or removed ordinary names', async (t) => {
  const f = await fixture(t),
    b = f.book,
    settings = {
      ...defaultPrintSettings(),
      area: 'A1:D3',
      repeatRows: { first: 1, last: 1 },
      repeatColumns: { first: 'A', last: 'A' },
    };
  const printed = await b.apply('0', [
      { operation: 'setPrintSettings', sheet: 'Sheet1', expectedPrintRevision: null, settings },
    ]),
    before = printed.view.storage.state,
    definition = await b.printDefinition(before, 'Sheet1');
  const changed = await b.apply(before, [update('workbook', "'Sheet1'!$A$1:$A$1", 'A2')]),
    removed = await b.apply(changed.view.storage.state, [remove('Unused', 'workbook', "'Sheet1'!$A$2:$A$2")]),
    state = removed.view.storage.state;
  assert.deepEqual(await b.printDefinition(state, 'Sheet1'), definition);
  assert.equal(await value(b, state, 'Other', 'B1'), 7);
  const saved = await b.export(state),
    parts = unzipSync(await fs.readFile(saved.path));
  assert.deepEqual(parts['custom/unchanged.bin'], f.parts['custom/unchanged.bin']);
  assert(strFromU8(parts['xl/workbook.xml']).includes('_xlnm.Print_Titles'));
});
test('Manual name edits keep the prior numeric cache stored until explicit calculation', async (t) => {
  const f = await fixture(t),
    b = f.book,
    manual = await b.apply('0', [{ type: 'setCalculationOptions', mode: 'manual', recalculateBeforeSave: false }]),
    changed = await b.apply(manual.view.storage.state, [update('workbook', "'Sheet1'!$A$1:$A$1", 'A2')]),
    state = changed.view.storage.state;
  let c = (await b.read(state, 'Other', 'B1')).cells.B1;
  assert.notEqual(c.calculation.status, 'calculated');
  assert.equal(c.calculation.value, undefined);
  assert.equal(c.cachedValue, 2);
  assert.equal(changed.view.calculationPolicy.mode, 'manual');
  const calculated = await b.apply(state, [{ type: 'recalculate' }]);
  assert.equal(await value(b, calculated.view.storage.state, 'Other', 'B1'), 7);
  c = (await b.read(state, 'Other', 'B1')).cells.B1;
  assert.notEqual(c.calculation.status, 'calculated');
  const saved = await b.export(calculated.view.storage.state);
  assert(strFromU8(unzipSync(await fs.readFile(saved.path))['xl/workbook.xml']).includes('calcMode="manual"'));
});
test('dynamic string-based consumers stay protected and text literals are not references', async (t) => {
  const f = await fixture(t, { dynamic: true });
  await assert.rejects(
    f.book.apply('0', [remove('Unused', 'workbook', "'Sheet1'!$A$2:$A$2")]),
    /Dynamic|qualification/
  );
  const ordinary = await fixture(t, { globalUse: false });
  const removed = await ordinary.book.apply('0', [remove('Unused', 'workbook', "'Sheet1'!$A$2:$A$2")]);
  assert.equal(await value(ordinary.book, removed.view.storage.state, 'Sheet1', 'D1'), 6);
});
for (const [label, formula, accepted] of [
  ['normal namespace declarations', '<f xmlns="' + ns + '" t="normal">SUM(Inputs)</f>', true],
  ['literal name text', '<f>LEN("Unused")</f>', true],
  ['shared formula', '<f t="shared" si="0">SUM(Inputs)</f>', false],
  ['array formula', '<f t="array" ref="C1:C2">SUM(Inputs)</f>', false],
  ['foreign formula namespace', '<f xmlns="urn:opaque">SUM(Inputs)</f>', false],
  ['unknown function', '<f>VendorName(Inputs)</f>', false],
  ['dynamic indirect', '<f>INDIRECT("Unused")</f>', false],
])
  test('independent native admission: ' + label, async (t) => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'native-name-consumers-'));
    t.after(() => fs.rm(directory, { recursive: true, force: true }));
    const parts = unzipSync(createWorkbook()),
      unused = '<definedName name="Unused">\'Sheet1\'!$A$2:$A$2</definedName>';
    const xml = strFromU8(parts['xl/workbook.xml']).replace(
      '</workbook>',
      '<definedNames><definedName name="Inputs">\'Sheet1\'!$A$1:$A$1</definedName>' +
        unused +
        '</definedNames></workbook>'
    );
    parts['xl/workbook.xml'] = strToU8(xml);
    parts['xl/worksheets/sheet1.xml'] = strToU8(
      '<worksheet xmlns="' +
        ns +
        '"><dimension ref="A1:C2"/><sheetData><row r="1"><c r="A1"><v>2</v></c><c r="C1">' +
        formula +
        '<v>999</v></c></row><row r="2"><c r="A2"><v>7</v></c></row></sheetData></worksheet>'
    );
    const original = Buffer.from(zipSync(parts)),
      source = path.join(directory, 'source.xlsx');
    await fs.writeFile(source, original);
    await native({ action: 'workbook', operation: 'index', root: directory });
    const plan = {
      kind: 'removeName',
      name: 'Unused',
      scope: null,
      priorFormula: "'Sheet1'!$A$2:$A$2",
      replacements: { 'xl/workbook.xml': xml.replace(unused, '') },
      removed: [],
      copies: {},
      fragments: {},
    };
    await fs.writeFile(path.join(directory, 'package-metadata.json'), JSON.stringify(plan));
    const output = path.join(directory, 'package.xlsx'),
      sentinel = Buffer.from('prior complete output');
    await fs.writeFile(output, sentinel);
    const operation = () =>
      native({ action: 'workbook', operation: 'package-edit', root: directory, version: 0, source: 'source.xlsx' });
    if (accepted) {
      await operation();
      assert.equal(
        strFromU8(unzipSync(await fs.readFile(output))['xl/workbook.xml']),
        plan.replacements['xl/workbook.xml']
      );
      assert.deepEqual(
        unzipSync(await fs.readFile(output))['xl/worksheets/sheet1.xml'],
        parts['xl/worksheets/sheet1.xml']
      );
    } else {
      await assert.rejects(operation(), /qualification|qualified|Dynamic|unknown|Shared|array|extended/);
      assert.deepEqual(await fs.readFile(output), sentinel);
    }
    assert.deepEqual(await fs.readFile(source), original);
  });
test('unused worksheet-local names remove safely; workbook names can retarget another worksheet', async (t) => {
  const f = await fixture(t),
    b = f.book,
    created = await b.apply('0', [
      {
        operation: 'defineName',
        sheet: 'Sheet1',
        name: 'UnusedLocal',
        nameScope: 'worksheet',
        expectedNameFormula: null,
        range: 'A2',
      },
    ]);
  const removed = await b.apply(created.view.storage.state, [remove('UnusedLocal', 'worksheet', "'Sheet1'!$A$2:$A$2")]);
  assert(!removed.view.namedRanges.some((n) => n.name === 'UnusedLocal'));
  assert.equal(await value(b, removed.view.storage.state, 'Sheet1', 'C1'), 3);
  const changed = await b.apply(removed.view.storage.state, [
    { ...update('workbook', "'Sheet1'!$A$1:$A$1", 'A1'), sheet: 'Other' },
  ]);
  assert.equal(changed.view.namedRanges.find((n) => n.name === 'Inputs' && !n.sheet).formula, "'Other'!$A$1:$A$1");
  assert.equal(await value(b, changed.view.storage.state, 'Other', 'B1'), 11);
  assert.equal(await value(b, changed.view.storage.state, 'Sheet1', 'C1'), 3);
});
