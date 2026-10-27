import { test } from 'node:test';
import assert from 'node:assert/strict';
import { localExecutionPreview } from '../core/local-action-execution.mjs';
import { readWorkbook, applyProjectedOperation } from '../core/workbook-package.mjs';
import { projectWorkbookCommands } from '../core/workbook-projection.mjs';
import { localExecutionFixture } from './support/local-execution-fixture.mjs';
test('two built-ins capture explicit target fields and use the first selected cell for bounded Freeze', () => {
  assert.deepEqual(localExecutionPreview('unhide-all', { sheet: 'Sheet1', selection: 'B3:C5' }).commands, [
    { operation: 'unhideAll', sheet: 'Sheet1' },
  ]);
  assert.deepEqual(localExecutionPreview('freeze-selection', { sheet: 'Sheet1', selection: 'B3:C5' }).commands, [
    { operation: 'freeze', sheet: 'Sheet1', rows: 2, columns: 1 },
  ]);
  assert.deepEqual(localExecutionPreview('freeze-selection', { sheet: 'Sheet1', selection: 'AE101' }).commands[0], {
    operation: 'freeze',
    sheet: 'Sheet1',
    rows: 100,
    columns: 30,
  });
});
test('invalid scope, arbitrary code/private descriptors, A1 and out-of-profile Freeze reject before any command', () => {
  for (const args of [
    ['unknown', { sheet: 'Sheet1', selection: 'B2' }],
    ['freeze-selection', { sheet: 'Sheet1', selection: 'A1' }],
    ['freeze-selection', { sheet: 'Sheet1', selection: 'AF102' }],
    ['unhide-all', { sheet: 'Sheet1', selection: 'B2', script: 'bad' }],
    ['unhide-all', { sheet: 'Sheet1', selection: 'no range' }],
    ['unhide-all', { sheet: 'Sheet1', selection: ['B2'] }],
  ])
    assert.throws(() => localExecutionPreview(...args));
});
test('portable projection and package writing preserve cells/custom geometry while unhiding all sparse/default rows and columns', async () => {
  const bytes = localExecutionFixture({ defaultHidden: true }),
    view = readWorkbook(bytes),
    before = structuredClone(view),
    cmd = { operation: 'unhideAll', sheet: 'Sheet1' },
    out = await projectWorkbookCommands(view, [cmd], {}),
    written = applyProjectedOperation(bytes, cmd, out.view),
    actual = written.view.sheets[0];
  assert.deepEqual(view, before);
  assert.equal(actual.sheetFormat.rowsHiddenByDefault, false);
  assert.deepEqual(actual.hiddenRows, []);
  assert.equal(actual.rowHeights.find((r) => r.row === 1).height, 37.5);
  assert.equal(actual.rowHeights.find((r) => r.row === 2).height, 22);
  assert(actual.widths.every((w) => !w.hidden));
  assert.equal(actual.widths.find((w) => w.min === 4).width, 12);
  assert.deepEqual(actual.cells, view.sheets[0].cells);
});
test('portable Unhide refuses linked filter/table visibility and malformed command fields atomically', async () => {
  for (const mode of ['filter', 'table']) {
    const b = localExecutionFixture({ [mode]: true }),
      v = readWorkbook(b);
    await assert.rejects(async () => {
      const p = await projectWorkbookCommands(v, [{ operation: 'unhideAll', sheet: 'Sheet1' }], {});
      applyProjectedOperation(b, { operation: 'unhideAll', sheet: 'Sheet1' }, p.view);
    }, /filter|table|protected/);
  }
  const v = readWorkbook(localExecutionFixture());
  await assert.rejects(projectWorkbookCommands(v, [{ operation: 'unhideAll', sheet: 'Sheet1', range: 'A1:B2' }], {}), {
    code: 'workbook-bad-input',
  });
});
