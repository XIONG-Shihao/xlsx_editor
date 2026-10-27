import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { createWorkbook, editWorkbook } from '../core/workbook-package.mjs';
import { indexedPivotPlan } from '../core/indexed-pivots.mjs';
import {
  nativeWorkbookOperationAsync as native,
  closeNativeWorkbookWorker,
  nativeWorkbookBinary,
} from '../core/native-runtime.mjs';
import { decodedCommands } from '../dist/core/transport.js';
after(() => closeNativeWorkbookWorker());
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex'),
  create = {
    operation: 'createPivot',
    sheet: 'Sheet1',
    name: 'SalesPivot',
    range: 'A1:C5',
    rowField: 0,
    valueField: 1,
    aggregate: 'sum',
    pivotSheet: 'Summary',
    expectedPivotRevision: null,
  };
function sourceBytes() {
  const data = [
      ['Category', 'Amount', 'Flag'],
      ['A', 2, true],
      ['B', 5, false],
      ['A', -1, true],
      ['A', null, false],
    ],
    edits = data.flatMap((row, r) =>
      row.map((value, c) => ({ sheet: 'Sheet1', cell: String.fromCharCode(65 + c) + (r + 1), value }))
    ),
    parts = unzipSync(editWorkbook(createWorkbook(), edits));
  parts['customXml/opaque.bin'] = new Uint8Array([0, 255, 17]);
  return Buffer.from(zipSync(parts));
}
async function fixture(t, bytes = sourceBytes()) {
  const root = await fs.realpath(await fs.mkdtemp(path.join(os.tmpdir(), 'document-pivot-'))),
    filename = path.join(root, 'source.xlsx'),
    recoveryDirectory = path.join(root, 'recovery'),
    owners = [],
    calls = [];
  let lastExport;
  await fs.writeFile(filename, bytes);
  const bridge = async (input) => {
      calls.push(input);
      const result = await native(input);
      if (input.operation === 'export') lastExport = { input, result };
      return result;
    },
    make = () => {
      const owner = new SpreadsheetDocument({ recoveryDirectory, native: bridge, cacheIdentity: null });
      owners.push(owner);
      return owner;
    };
  t.after(async () => {
    try {
      for (const owner of owners) await owner.close();
    } finally {
      await fs.rm(root, { recursive: true, force: true });
    }
  });
  return { root, filename, recoveryDirectory, bytes, make, calls, lastExport: () => lastExport };
}
const apply = (owner, commands) => owner.request({ operation: 'apply', state: owner.status().state, commands });
const read = (owner, range = 'A1:B4', state = owner.status().state, sheet = 'Summary') =>
  owner.request({ operation: 'read', state, sheet, range });
const values = async (owner, expected) => {
  const cells = (await read(owner)).cells;
  assert.deepEqual(
    [
      cells.A1.value,
      cells.B1.value,
      cells.A2.value,
      cells.B2.value,
      cells.A3.value,
      cells.B3.value,
      cells.A4.value,
      cells.B4.value,
    ],
    expected
  );
};
const initial = ['Category', 'Sum of Amount', 'A', 1, 'B', 5, 'Grand Total', 6],
  refreshed = ['Category', 'Sum of Amount', 'A', 6, 'B', 5, 'Grand Total', 11];
async function copy(owner, filename) {
  await owner.request({ operation: 'copy', destination: filename, expectedRevision: null });
  return filename;
}
async function oracle(source, saved, amount = 2, aggregate = 'sum', output) {
  const args = [
    'tests/support/pivot-package-oracle.py',
    source,
    saved,
    '--source-amount',
    String(amount),
    '--aggregate',
    aggregate,
  ];
  if (output) args.push('--output', output);
  const result = spawnSync(process.env.XLSX_EDITOR_PYTHON ?? 'python3', args, {
    encoding: 'utf8',
    maxBuffer: 1024 * 1024,
  });
  assert.equal(result.status, 0, result.stderr);
  return JSON.parse(result.stdout);
}
test('durable pivot create/edit/refresh has one frame per action, Undo/Redo replay and independently saved native graph', async (t) => {
  const f = await fixture(t),
    owner = f.make(),
    base = await owner.open(f.filename),
    created = await apply(owner, [create]);
  assert.equal(created.view.pivots.length, 1);
  assert.equal(created.view.pivots[0].canEdit, true);
  assert(created.canUndo);
  assert(!JSON.stringify(created).includes(f.root));
  await values(owner, initial);
  await owner.request({ operation: 'undo' });
  assert.equal(owner.status().state, base.state);
  assert.equal(owner.status().view.sheets.length, 1);
  await owner.request({ operation: 'redo' });
  assert.equal(owner.status().state, created.state);
  await values(owner, initial);
  const copied = await copy(owner, path.join(f.root, 'created.xlsx')),
    createdOracle = await oracle(f.filename, copied),
    directory = path.join(f.recoveryDirectory, created.documentId),
    node = JSON.parse(await fs.readFile(path.join(directory, created.state + '.json'), 'utf8'));
  assert.deepEqual(decodedCommands({ encodedCommands: node.commands }), [create]);
  assert(Buffer.byteLength(JSON.stringify(node)) < 2048);
  assert.deepEqual(await fs.readFile(f.filename), f.bytes);
  const changed = await apply(owner, [{ sheet: 'Sheet1', cell: 'B2', value: 7 }]);
  await values(owner, initial);
  const refresh = {
      operation: 'refreshPivot',
      sheet: 'Summary',
      name: 'SalesPivot',
      expectedPivotRevision: created.view.pivots[0].revision,
    },
    accepted = await apply(owner, [refresh]);
  await values(owner, refreshed);
  await assert.rejects(owner.request({ operation: 'apply', state: changed.state, commands: [refresh] }), /stale/);
  await assert.rejects(apply(owner, [refresh]), /Conflict/);
  assert.equal(owner.status().state, accepted.state);
  await owner.request({ operation: 'undo' });
  await values(owner, initial);
  await owner.request({ operation: 'redo' });
  await values(owner, refreshed);
  const final = await copy(owner, path.join(f.root, 'refreshed.xlsx')),
    finalOracle = await oracle(f.filename, final, 7);
  await owner.request({ operation: 'save' });
  assert.equal(owner.status().dirty, false);
  await owner.close();
  const cold = f.make(),
    opened = await cold.open(f.filename);
  assert(opened.view.pivots[0].canEdit);
  assert.equal(opened.canUndo, false);
  await values(cold, refreshed);
  assert.deepEqual(await fs.readFile(f.filename), await fs.readFile(final));
});
test('saved pivot clean cursor retains an unsaved refresh Redo tip across close and cold owner recovery', async (t) => {
  const f = await fixture(t),
    owner = f.make();
  await owner.open(f.filename);
  const created = await apply(owner, [create]);
  await owner.request({ operation: 'save' });
  const saved = hash(await fs.readFile(f.filename)),
    edited = await apply(owner, [{ sheet: 'Sheet1', cell: 'B2', value: 7 }]),
    refresh = {
      operation: 'refreshPivot',
      sheet: 'Summary',
      name: 'SalesPivot',
      expectedPivotRevision: created.view.pivots[0].revision,
    },
    accepted = await apply(owner, [refresh]);
  await owner.request({ operation: 'undo' });
  await owner.request({ operation: 'undo' });
  assert.equal(owner.status().dirty, false);
  assert.equal(owner.status().canRedo, true);
  await owner.close();
  const restored = f.make(),
    recovered = await restored.open(f.filename);
  assert(recovered.recovered);
  assert.equal(recovered.dirty, false);
  assert.equal(recovered.canRedo, true);
  await values(restored, initial);
  await restored.request({ operation: 'redo' });
  assert.equal(restored.status().state, edited.state);
  await values(restored, initial);
  await restored.request({ operation: 'redo' });
  assert.equal(restored.status().state, accepted.state);
  await values(restored, refreshed);
  assert.equal(hash(await fs.readFile(f.filename)), saved);
});
test('imported/opaque pivot graphs, readonly sources and modified output cells remain protected and byte-preserved', async (t) => {
  const base = await fixture(t),
    writer = base.make();
  await writer.open(base.filename);
  await apply(writer, [create]);
  const saved = await copy(writer, path.join(base.root, 'managed.xlsx')),
    original = unzipSync(await fs.readFile(saved));
  const imported = structuredClone(original),
    old = 'xl/pivotTables/workbookPivot1.xml',
    renamed = 'xl/pivotTables/imported1.xml';
  imported[renamed] = imported[old];
  delete imported[old];
  imported['xl/pivotTables/_rels/imported1.xml.rels'] = imported['xl/pivotTables/_rels/workbookPivot1.xml.rels'];
  delete imported['xl/pivotTables/_rels/workbookPivot1.xml.rels'];
  for (const name of ['[Content_Types].xml', 'xl/worksheets/_rels/sheet2.xml.rels'])
    imported[name] = strToU8(strFromU8(imported[name]).replaceAll('workbookPivot1.xml', 'imported1.xml'));
  const f = await fixture(t, Buffer.from(zipSync(imported))),
    owner = f.make(),
    opened = await owner.open(f.filename);
  assert.equal(opened.view.pivots[0].canEdit, false);
  await assert.rejects(apply(owner, [create]), /existing/);
  await assert.rejects(
    apply(owner, [
      { operation: 'refreshPivot', sheet: 'Summary', name: 'SalesPivot', expectedPivotRevision: '0'.repeat(64) },
    ]),
    /protected|Conflict/
  );
  const copied = await copy(owner, path.join(f.root, 'imported-copy.xlsx'));
  assert.deepEqual(await fs.readFile(copied), f.bytes);
  const readonly = structuredClone(original);
  // ActiveX keeps a workbook read-only (a macro project no longer does).
  readonly['customXml/activeX1.xml'] = strToU8('<opaque/>');
  const r = await fixture(t, Buffer.from(zipSync(readonly))),
    ro = r.make(),
    roStatus = await ro.open(r.filename);
  assert(roStatus.view.readOnly);
  await assert.rejects(
    apply(ro, [
      {
        operation: 'refreshPivot',
        sheet: 'Summary',
        name: 'SalesPivot',
        expectedPivotRevision: roStatus.view.pivots[0].revision,
      },
    ]),
    /read-only|readonly/
  );
  const exact = await copy(ro, path.join(r.root, 'readonly-copy.xlsx'));
  assert.deepEqual(await fs.readFile(exact), r.bytes);
  const changed = structuredClone(original);
  changed['xl/worksheets/sheet2.xml'] = strToU8(
    strFromU8(changed['xl/worksheets/sheet2.xml']).replace('<v>1</v>', '<v>99</v>')
  );
  const c = await fixture(t, Buffer.from(zipSync(changed))),
    cold = c.make(),
    status = await cold.open(c.filename);
  await assert.rejects(
    apply(cold, [
      {
        operation: 'refreshPivot',
        sheet: 'Summary',
        name: 'SalesPivot',
        expectedPivotRevision: status.view.pivots[0].revision,
      },
    ]),
    /changed|protected|Conflict/
  );
  assert.deepEqual(await fs.readFile(c.filename), c.bytes);
});
test('cold managed output metadata mutants cannot be overwritten by refresh', async (t) => {
  const f = await fixture(t),
    writer = f.make();
  await writer.open(f.filename);
  const created = await apply(writer, [create]),
    managed = await copy(writer, path.join(f.root, 'managed.xlsx')),
    original = unzipSync(await fs.readFile(managed));
  const mutations = [
    (raw) => raw.replace('<row r="2">', '<row r="2" hidden="1">'),
    (raw) => raw.replace('width="24"', 'width="37"'),
    (raw) => raw.replace('workbookViewId="0"', 'workbookViewId="0" showGridLines="0"'),
    (raw) => raw.replace('<worksheet ', '<worksheet custom="opaque" '),
  ];
  for (const [index, mutate] of mutations.entries()) {
    const parts = structuredClone(original);
    parts['xl/worksheets/sheet2.xml'] = strToU8(mutate(strFromU8(parts['xl/worksheets/sheet2.xml'])));
    const c = await fixture(t, Buffer.from(zipSync(parts))),
      owner = c.make(),
      opened = await owner.open(c.filename),
      revision = opened.view.pivots[0].revision;
    let error;
    try {
      await apply(owner, [
        { operation: 'refreshPivot', sheet: 'Summary', name: 'SalesPivot', expectedPivotRevision: revision },
      ]);
    } catch (failure) {
      error = failure;
    }
    assert(error, 'Metadata mutant ' + index + ' was accepted');
    assert.match(error.message, /protected|Conflict|metadata/);
    assert.deepEqual(await fs.readFile(c.filename), c.bytes);
  }
});
test('independent native admission rejects forged before-target metadata even with a valid plan/current accepted cells', async (t) => {
  const f = await fixture(t),
    writer = f.make();
  await writer.open(f.filename);
  const created = await apply(writer, [create]);
  await apply(writer, [{ sheet: 'Sheet1', cell: 'B2', value: 7 }]);
  const command = {
      operation: 'refreshPivot',
      sheet: 'Summary',
      name: 'SalesPivot',
      expectedPivotRevision: created.view.pivots[0].revision,
    },
    current = writer.status();
  await read(writer);
  const rangeCall = f.calls.filter((input) => input.operation === 'range').at(-1),
    root = rangeCall.root,
    metadata = await fs.readFile(path.join(root, 'metadata.xlsx')),
    prepared = await indexedPivotPlan({
      metadataBytes: metadata,
      view: current.view,
      command,
      readRange: (sheet, range) => writer.request({ operation: 'read', state: current.state, sheet, range }),
      readWorksheet: async (part) => {
        await copy(writer, path.join(f.root, 'actual-target-source.xlsx'));
        return unzipSync(await fs.readFile(path.join(f.root, 'actual-target-source.xlsx')))[part];
      },
    });
  await copy(writer, path.join(f.root, 'before-forgery.xlsx'));
  const exported = f.lastExport();
  assert(exported?.result.file);
  const exportedFile = path.join(exported.input.root, exported.result.file),
    forged = unzipSync(await fs.readFile(exportedFile));
  forged[prepared.plan.targetPath] = strToU8(
    strFromU8(forged[prepared.plan.targetPath]).replace('<row r="2">', '<row r="2" hidden="1">')
  );
  await fs.writeFile(exportedFile, zipSync(forged));
  await native({
    action: 'workbook',
    operation: 'artifact-write',
    root: exported.input.root,
    name: 'package-metadata.json',
    data: JSON.stringify(prepared.plan),
  });
  let failure, result;
  try {
    result = await native({
      action: 'workbook',
      operation: 'package-edit',
      root: exported.input.root,
      version: exported.input.version,
      source: exported.result.file,
    });
  } catch (error) {
    failure = error;
  }
  assert(failure, 'Native accepted forged prior target metadata');
  assert.match(failure.message, /metadata|profile|target|preserve|Extended/i);
  assert.equal(writer.status().state, current.state);
  assert.deepEqual(await fs.readFile(f.filename), f.bytes);
});
