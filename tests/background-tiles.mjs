import test from 'node:test';
import assert from 'node:assert/strict';
import { backgroundTile, backgroundCellStyle } from '../ui/backgroundTiles.ts';
const sheet = {
  path: 'xl/worksheets/main.xml',
  background: {
    schema: 'workbook.xlsx.sheet-background.v1',
    sheet: 'xl/worksheets/main.xml',
    sample: 'graph-paper',
    revision: 'a'.repeat(64),
    canEdit: true,
    protectedReason: null,
  },
};
test('tile coordinates remain source-relative across zoom, neighboring cells and merged/frozen placement', () => {
  const normal = backgroundTile(sheet, 1),
    zoom = backgroundTile(sheet, 1.5);
  assert.equal(normal.size, 32);
  assert.equal(zoom.size, 48);
  assert.equal(backgroundCellStyle(normal, 64, 96, false).backgroundPosition, '-64px -96px');
  assert.equal(backgroundCellStyle(zoom, 96, 144, false).backgroundPosition, '-96px -144px');
  const a = backgroundCellStyle(normal, 71, 93, false),
    b = backgroundCellStyle(normal, 142, 93, false);
  assert.equal(a.backgroundPosition, '-71px -93px');
  assert.equal(b.backgroundPosition, '-142px -93px');
  assert.equal(a.backgroundImage, b.backgroundImage);
  assert.deepEqual(backgroundCellStyle(normal, 0, 0, true), {});
  assert.equal(backgroundCellStyle(normal, 0, 0, false).backgroundRepeat, 'repeat');
});
test('foreign/retired/protected metadata never substitutes a sample or blocks ordinary cell display', () => {
  for (const patch of [
    { sheet: 'other' },
    { schema: 'other' },
    { sample: 'imported' },
    { canEdit: false },
    { protectedReason: 'shared' },
    { revision: null },
  ])
    assert.equal(backgroundTile({ ...sheet, background: { ...sheet.background, ...patch } }, 1), undefined);
  for (const scale of [0, NaN, Infinity, 5]) assert.equal(backgroundTile(sheet, scale), undefined);
  assert.equal(backgroundTile(undefined, 1), undefined);
  assert.deepEqual(backgroundCellStyle(undefined, 0, 0, false), {});
});
