/** Real Document/native owned-cache boundaries; no Excel/GUI parity claim. */
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import assert from 'node:assert/strict';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook, editWorkbook } from '../core/workbook-package.mjs';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { createWorkbookNativeClient } from '../core/native-client.mjs';
import { nativeWorkbookBinary } from '../core/native-runtime.mjs';
import { record } from './support/file-record.mjs';
const output = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-calculation-derived-owner-'));
const helper = nativeWorkbookBinary();
const client = createWorkbookNativeClient(helper),
  owners = [];
let nativeRoot, nativeVersion;
const native = async (input) => {
  const result = await client.call(input);
  if (input.root) {
    nativeRoot = input.root;
    nativeVersion = result.version ?? input.version ?? 0;
  }
  return result;
};
function fixture(edits, { mode = 'auto', extraCalc = '', twoSheets = false } = {}) {
  let bytes = editWorkbook(createWorkbook(), edits),
    parts = unzipSync(bytes);
  let workbook = strFromU8(parts['xl/workbook.xml']).replace(/<calcPr\b[^>]*(?:\/>|>[\s\S]*?<\/calcPr>)/g, '');
  if (twoSheets) {
    workbook = workbook.replace('</sheets>', '<sheet name="Other" sheetId="2" r:id="rId3"/></sheets>');
    parts['xl/worksheets/sheet2.xml'] = parts['xl/worksheets/sheet1.xml'].slice();
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
  }
  parts['xl/workbook.xml'] = strToU8(
    workbook.replace('</workbook>', `<calcPr calcMode="${mode}" calcOnSave="0" ${extraCalc}/></workbook>`)
  );
  parts['customXml/opaque.bin'] = new Uint8Array([0, 255, 29, 81]);
  return Buffer.from(zipSync(parts));
}
async function open(name, bytes, options = {}) {
  const directory = path.join(output, name);
  await fs.mkdir(directory);
  const source = path.join(directory, 'source.xlsx');
  await fs.writeFile(source, bytes);
  const original = path.join(directory, 'original.xlsx');
  await fs.copyFile(source, original);
  const owner = new SpreadsheetDocument({
    native,
    indexDirectory: path.join(directory, 'index'),
    recoveryDirectory: path.join(directory, 'recovery'),
    cacheIdentity: null,
    ...options,
  });
  owners.push(owner);
  await owner.open(source);
  return { directory, source, original, owner, originalHash: (await record(source)).sha256 };
}
const status = (f) => f.owner.status();
const request = (f, operation, fields = {}) =>
  f.owner.request({ operation, documentId: status(f).documentId, state: status(f).state, ...fields });
const apply = (f, commands) => request(f, 'apply', { commands });
const read = (f, sheet, range) => request(f, 'read', { sheet, range });
async function cells(f, sheet, range) {
  return (await read(f, sheet, range)).cells;
}
async function cold(f) {
  await f.owner.close();
  f.owner = new SpreadsheetDocument({
    native,
    indexDirectory: path.join(f.directory, 'index'),
    recoveryDirectory: path.join(f.directory, 'recovery'),
    cacheIdentity: null,
  });
  owners.push(f.owner);
  await f.owner.open(f.source);
}
function stored(cell, value) {
  assert.equal(cell.calculation?.pending, true);
  assert.equal(cell.calculation?.status, 'stored');
  assert.equal(cell.calculation?.value, undefined);
  assert.equal(cell.cachedResult?.authority, 'stored');
  assert(Object.is(cell.cachedResult?.value, value), JSON.stringify(cell));
}
function current(cell, value) {
  assert.equal(cell.calculation?.pending, undefined);
  assert.equal(cell.calculation?.status, 'calculated');
  assert(Object.is(cell.calculation.value, value), JSON.stringify(cell));
}
const chartId = '11111111-1111-4111-8111-111111111111',
  otherId = '22222222-2222-4222-8222-222222222222';
const chart = (sheet, id, range) => ({
  operation: 'createChart',
  sheet,
  objectId: id,
  expectedObjectRevision: null,
  chart: { kind: 'column', range, title: 'Accepted values' },
});
const chartPart = (id) => 'xl/charts/workbook-chart-' + id + '.xml';
function chartValues(parts, id) {
  const xml = strFromU8(parts[chartPart(id)]);
  return [...xml.matchAll(/<c:numCache>([\s\S]*?)<\/c:numCache>/g)].map((match) =>
    [...match[1].matchAll(/<c:pt idx="\d+"><c:v>(.*?)<\/c:v><\/c:pt>/g)].map((value) => Number(value[1]))
  );
}
async function sourceExact(f) {
  assert.equal((await record(f.source)).sha256, f.originalHash);
}
// Every case runs; failures are collected and asserted at the end.
const failures = [];
async function test(name, body) {
  try {
    await body();
  } catch (error) {
    failures.push(name + ': ' + (error.stack ?? error));
  }
}
try {
  await test('ordinary-automatic-sheet-and-protected-source', async () => {
    const base = fixture([
      { sheet: 'Sheet1', cell: 'A1', value: 2 },
      { sheet: 'Sheet1', cell: 'B1', formula: 'A1+1' },
    ]);
    const f = await open('auto-sheet', base);
    assert.equal(status(f).view.calculationPolicy.canCalculateSheet, true);
    const before = status(f);
    await apply(f, [{ type: 'calculateSheet', sheet: 'Sheet1' }]);
    assert.notEqual(status(f).state, before.state);
    current((await cells(f, 'Sheet1', 'B1')).B1, 3);
    await sourceExact(f);
    const refused = [];
    for (const [label, mode, extraCalc] of [
      ['mode', 'autoNoTable', ''],
      ['iterate', 'auto', 'iterate="1"'],
      ['opaque', 'auto', 'opaque="1"'],
    ]) {
      const g = await open(
        'protected-' + label,
        fixture(
          [
            { sheet: 'Sheet1', cell: 'A1', value: 2 },
            { sheet: 'Sheet1', cell: 'B1', formula: 'A1+1' },
          ],
          { mode, extraCalc }
        ),
        { initialCalculationPolicy: { mode: 'auto', recalculateBeforeSave: true } }
      );
      assert.equal(status(g).view.readOnly, false);
      assert.equal(status(g).view.calculationPolicy.preserveSource, true);
      assert.equal(status(g).view.calculationPolicy.protectedSource, true);
      assert.equal(status(g).view.calculationPolicy.canRecalculate, false);
      const before = status(g);
      await assert.rejects(apply(g, [{ type: 'recalculate' }]), /read.only|protected/);
      assert.equal(status(g).state, before.state);
      await read(g, 'Sheet1', 'B1');
      const view = { ...status(g).view, readOnly: false };
      await assert.rejects(
        native({
          action: 'workbook',
          operation: 'calculation-policy',
          method: 'options',
          root: nativeRoot,
          version: nativeVersion,
          view,
          mode: 'auto',
        }),
        /protected|source/
      );
      await sourceExact(g);
      refused.push({ label, status: status(g) });
    }
    return { before, after: status(f), protected: refused };
  });
  await test('pending-signed-zero-style-false-save-and-native-consumer-refusals', async () => {
    let parts = unzipSync(
      fixture(
        [
          { sheet: 'Sheet1', cell: 'A1', value: 2 },
          { sheet: 'Sheet1', cell: 'B1', formula: 'A1-A1' },
        ],
        { mode: 'manual' }
      )
    );
    let xml = strFromU8(parts['xl/worksheets/sheet1.xml']);
    xml = xml.replace(
      /(<c[^>]*r="B1"[^>]*>[\s\S]*?<f[^>]*>[\s\S]*?<\/f>)(?:<v[^>]*>[\s\S]*?<\/v>)?(<\/c>)/,
      '$1<v>-0</v>$2'
    );
    assert.match(xml, /<v>-0<\/v>/);
    parts['xl/worksheets/sheet1.xml'] = strToU8(xml);
    const f = await open('signed-zero', Buffer.from(zipSync(parts)));
    stored((await cells(f, 'Sheet1', 'B1')).B1, -0);
    const before = status(f);
    await assert.rejects(
      apply(f, [{ type: 'pasteValues', sheet: 'Sheet1', source: 'B1', range: 'C1' }]),
      /pending|stored|calculated/
    );
    assert.equal(status(f).state, before.state);
    await read(f, 'Sheet1', 'B1');
    await assert.rejects(
      native({
        action: 'workbook',
        operation: 'print-check',
        root: nativeRoot,
        version: nativeVersion,
        sheet: 'xl/worksheets/sheet1.xml',
        range: 'B1',
        view: status(f).view,
      }),
      /PRINT_PENDING_CALCULATION/
    );
    await apply(f, [{ sheet: 'Sheet1', cell: 'B1', bold: true }]);
    stored((await cells(f, 'Sheet1', 'B1')).B1, -0);
    await request(f, 'save');
    const saved = unzipSync(await fs.readFile(f.source));
    assert.match(strFromU8(saved['xl/worksheets/sheet1.xml']), /<v>-0<\/v>/);
    await cold(f);
    stored((await cells(f, 'Sheet1', 'B1')).B1, -0);
    return { before, after: status(f), signedZeroPreserved: true, saved: await record(f.source) };
  });
  await test('manual-owned-spill-pending-follower-cache-false-save-cold-and-full-now', async () => {
    const f = await open(
      'spill',
      fixture([
        { sheet: 'Sheet1', cell: 'J1', value: 3 },
        { sheet: 'Sheet1', cell: 'A1', formula: 'SEQUENCE($J$1,2)' },
        { sheet: 'Sheet1', cell: 'D1', formula: 'SUM(_xlfn.ANCHORARRAY(A1))' },
      ])
    );
    let page = await cells(f, 'Sheet1', 'A1:J3');
    current(page.D1, 21);
    assert.equal(page.B3.value, 6);
    await apply(f, [{ type: 'setCalculationOptions', mode: 'manual', recalculateBeforeSave: false }]);
    await apply(f, [{ sheet: 'Sheet1', cell: 'J1', value: 2 }]);
    const before = status(f);
    page = await cells(f, 'Sheet1', 'A1:J3');
    stored(page.B3, 6);
    stored(page.D1, 21);
    await assert.rejects(apply(f, [{ sheet: 'Sheet1', cell: 'B2', value: 99 }]), /follower/);
    await request(f, 'save');
    assert.equal(status(f).state, before.state);
    const savedParts = unzipSync(await fs.readFile(f.source));
    assert.match(strFromU8(savedParts['xl/worksheets/sheet1.xml']), /ref="A1:B3"/);
    assert.match(
      strFromU8(savedParts['xl/worksheets/sheet1.xml']),
      /<c\b[^>]*\br="B3"[^>]*><v\b[^>]*>6(?:\.0)?<\/v><\/c>/
    );
    await cold(f);
    page = await cells(f, 'Sheet1', 'A1:J3');
    stored(page.B3, 6);
    stored(page.D1, 21);
    const retained = status(f);
    const accepted = await apply(f, [{ type: 'recalculate' }]);
    page = await cells(f, 'Sheet1', 'A1:J3');
    current(page.D1, 10);
    assert.equal(page.B2.value, 4);
    assert.equal(page.B3.value, null);
    await request(f, 'undo');
    assert.equal(status(f).state, retained.state);
    stored((await cells(f, 'Sheet1', 'B3')).B3, 6);
    await request(f, 'redo');
    assert.equal(status(f).state, accepted.state);
    current((await cells(f, 'Sheet1', 'D1')).D1, 10);
    return { before, retained, accepted, saved: await record(f.source), pendingFollower: 6, refreshedSum: 10 };
  });
  await test('manual-owned-pivot-and-chart-full-now-one-undo', async () => {
    const data = [
        ['Category', 'Amount'],
        ['A', 2],
        ['B', 5],
        ['A', -1],
      ],
      edits = data.flatMap((row, r) =>
        row.map((value, c) => ({ sheet: 'Sheet1', cell: String.fromCharCode(65 + c) + (r + 1), value }))
      );
    const f = await open('pivot-chart', fixture(edits));
    await apply(f, [
      {
        operation: 'createPivot',
        sheet: 'Sheet1',
        name: 'SalesPivot',
        range: 'A1:B4',
        rowField: 0,
        valueField: 1,
        aggregate: 'sum',
        pivotSheet: 'Summary',
        expectedPivotRevision: null,
      },
    ]);
    await apply(f, [chart('Sheet1', chartId, 'A1:B4')]);
    await apply(f, [{ type: 'setCalculationOptions', mode: 'manual', recalculateBeforeSave: false }]);
    await apply(f, [{ sheet: 'Sheet1', cell: 'B2', value: 7 }]);
    const before = status(f);
    stored((await cells(f, 'Summary', 'B4')).B4, 6);
    await apply(f, [{ type: 'recalculate' }]);
    const accepted = status(f);
    assert.equal(accepted.view.calculationPolicy.pendingCacheObjects, 0);
    assert.equal((await cells(f, 'Summary', 'B4')).B4.value, 11);
    await request(f, 'undo');
    assert.equal(status(f).state, before.state);
    stored((await cells(f, 'Summary', 'B4')).B4, 6);
    await request(f, 'redo');
    assert.equal(status(f).state, accepted.state);
    await request(f, 'save');
    const parts = unzipSync(await fs.readFile(f.source));
    assert.deepEqual(chartValues(parts, chartId), [[7, 5, -1]]);
    assert.match(strFromU8(parts['xl/worksheets/sheet2.xml']), /<c\b[^>]*\br="B4"[^>]*><v\b[^>]*>11(?:\.0)?<\/v><\/c>/);
    assert.deepEqual(parts['customXml/opaque.bin'], unzipSync(await fs.readFile(f.original))['customXml/opaque.bin']);
    return { before, accepted, saved: await record(f.source), chartValues: chartValues(parts, chartId) };
  });
  await test('selected-sheet-chart-prerequisites-and-offbranch-pending-export', async () => {
    const edits = [
        { sheet: 'Sheet1', cell: 'A1', value: 'Month' },
        { sheet: 'Sheet1', cell: 'B1', value: 'Amount' },
        { sheet: 'Sheet1', cell: 'A2', value: 'Jan' },
        { sheet: 'Sheet1', cell: 'B2', formula: 'C2*2' },
        { sheet: 'Sheet1', cell: 'C2', value: 2 },
        { sheet: 'Sheet1', cell: 'A3', value: 'Feb' },
        { sheet: 'Sheet1', cell: 'B3', formula: 'C3*2' },
        { sheet: 'Sheet1', cell: 'C3', value: 3 },
      ],
      f = await open('selected-charts', fixture(edits, { twoSheets: true }));
    await apply(f, [chart('Sheet1', chartId, 'A1:B3')]);
    await apply(f, [chart('Other', otherId, 'A1:B3')]);
    await apply(f, [{ type: 'setCalculationOptions', mode: 'manual', recalculateBeforeSave: false }]);
    await apply(f, [
      { sheet: 'Sheet1', cell: 'C2', value: 5 },
      { sheet: 'Other', cell: 'C2', value: 7 },
    ]);
    const before = status(f);
    const accepted = await apply(f, [{ type: 'calculateSheet', sheet: 'Sheet1' }]);
    assert.equal(status(f).view.calculationPolicy.pendingCacheObjects, 1);
    current((await cells(f, 'Sheet1', 'B2')).B2, 10);
    stored((await cells(f, 'Other', 'B2')).B2, 4);
    await request(f, 'save');
    const parts = unzipSync(await fs.readFile(f.source));
    assert.deepEqual(chartValues(parts, chartId), [[10, 6]]);
    assert.deepEqual(chartValues(parts, otherId), [[4, 6]]);
    assert.match(strFromU8(parts['xl/workbook.xml']), /fullCalcOnLoad="1"/);
    await request(f, 'undo');
    assert.equal(status(f).state, before.state);
    stored((await cells(f, 'Sheet1', 'B2')).B2, 4);
    await request(f, 'redo');
    assert.equal(status(f).state, accepted.state);
    return {
      before,
      accepted,
      saved: await record(f.source),
      selected: chartValues(parts, chartId),
      offbranch: chartValues(parts, otherId),
      requiredPivotTransfer: 'pending',
    };
  });
  assert.deepEqual(failures, []);
} finally {
  const closing = await Promise.allSettled(owners.map((owner) => owner.close()));
  await client.close();
  await fs.rm(output, { recursive: true, force: true });
  for (const result of closing) if (result.status === 'rejected') throw result.reason;
}
