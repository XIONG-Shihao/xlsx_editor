/** Actual owner creation; protected geometry is read metadata, never edit authority. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { createWorkbook, editWorkbook } from '../core/workbook-package.mjs';
import { indexedPivotInventory } from '../core/indexed-pivots.mjs';
import { closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
test('owner exposes actual field names; protected location inspection requires one valid relationship-bound rectangle', async (t) => {
  const dir = await fs.realpath(await fs.mkdtemp(path.join(os.tmpdir(), 'pivot-context-inventory-'))),
    file = path.join(dir, 'source.xlsx'),
    owner = new SpreadsheetDocument({ recoveryDirectory: path.join(dir, 'recovery'), cacheIdentity: null });
  t.after(async () => {
    await owner.close();
    await fs.rm(dir, { recursive: true, force: true });
  });
  const values = [
    ['Category', 'Amount'],
    ['A', 2],
    ['B', 3],
  ];
  await fs.writeFile(
    file,
    editWorkbook(
      createWorkbook(),
      values.flatMap((row, r) =>
        row.map((value, c) => ({ sheet: 'Sheet1', cell: String.fromCharCode(65 + c) + (r + 1), value }))
      )
    )
  );
  await owner.open(file);
  const created = await owner.request({
    operation: 'apply',
    state: owner.status().state,
    commands: [
      {
        operation: 'createPivot',
        sheet: 'Sheet1',
        name: 'SalesPivot',
        range: 'A1:B3',
        rowField: 0,
        valueField: 1,
        aggregate: 'sum',
        pivotSheet: 'Summary',
        expectedPivotRevision: null,
      },
    ],
  });
  const p = created.view.pivots[0];
  assert.deepEqual(p.fieldNames, ['Category', 'Amount']);
  assert.equal(p.outputRange, 'A1:B4');
  assert(!('model' in p));
  assert(!('parts' in p));
  const saved = path.join(dir, 'owned.xlsx');
  await owner.request({ operation: 'copy', destination: saved, expectedRevision: null });
  const parts = unzipSync(await fs.readFile(saved)),
    table = 'xl/pivotTables/workbookPivot1.xml',
    raw = strFromU8(parts[table]);
  function inventory(change, view = created.view) {
    const altered = structuredClone(parts);
    altered[table] = strToU8(change(raw));
    return indexedPivotInventory(zipSync(altered), view)[0];
  }
  const protectedPivot = inventory((x) =>
    x.replace('<pivotTableDefinition ', '<pivotTableDefinition custom="opaque" ')
  );
  assert.equal(protectedPivot.name, 'SalesPivot');
  assert.equal(protectedPivot.canEdit, false);
  assert.equal(protectedPivot.revision, null);
  assert.equal(protectedPivot.sheet, 'Summary');
  assert.equal(protectedPivot.outputRange, 'A1:B4');
  assert(protectedPivot.reason);
  assert.equal(protectedPivot.sourceRange, undefined);
  for (const mutate of [
    (x) => x.replace('ref="A1:B4"', 'ref="XFE1:XFE4"'),
    (x) => x.replace('ref="A1:B4"', 'ref="B4:A1"'),
    (x) => x.replace('ref="A1:B4"', 'ref="A1:B4:C9"'),
    (x) => x.replace('<location ', '<unknownLocation '),
    (x) => x.replace('</pivotTableDefinition>', '<location ref="C1:D4"/></pivotTableDefinition>'),
  ]) {
    const bad = inventory(mutate);
    assert.equal(bad.canEdit, false);
    assert.equal(bad.outputRange, undefined);
  }
  const unbound = inventory((x) => x.replace('<pivotTableDefinition ', '<pivotTableDefinition custom="opaque" '), {
    ...created.view,
    sheets: created.view.sheets.filter((s) => s.name !== 'Summary'),
  });
  assert.equal(unbound.sheet, null);
  assert.equal(unbound.outputRange, undefined);
  const target = created.view.sheets.find((s) => s.name === 'Summary'),
    ambiguous = inventory((x) => x.replace('<pivotTableDefinition ', '<pivotTableDefinition custom="opaque" '), {
      ...created.view,
      sheets: [...created.view.sheets, { ...target, name: 'Ambiguous', path: target.path }],
    });
  assert.equal(ambiguous.sheet, null);
  assert.equal(ambiguous.outputRange, undefined);
});
