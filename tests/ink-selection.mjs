import { test } from 'node:test';
import assert from 'node:assert/strict';
import { pickOwnedInk } from '../ui/inkSelection.ts';
import { inkFoundationCases } from './support/ink-foundation-fixture.mjs';
import { createObjectAxes, objectPlacementCommand, objectGestureBox } from '../ui/worksheetObjectGeometry.ts';
const axes = createObjectAxes({ sheetFormat: { defaultColumnWidth: 10, defaultRowHeight: 15 } }),
  objects = inkFoundationCases().map((c) => ({
    id: c.id,
    name: c.label,
    kind: 'ink',
    revision: 'r1',
    managed: true,
    canEdit: true,
    canDelete: true,
    protectedReason: null,
    anchor: c.anchor,
    stroke: c.stroke,
  }));
test('owned picking follows actual open paths, not a rectangular silhouette; nearest painted owned stroke wins', () => {
  assert.equal(pickOwnedInk(objects, [90, 50], axes, 1).id, objects[0].id);
  assert.equal(pickOwnedInk(objects, [95, 78], axes, 1), null);
  assert.equal(pickOwnedInk(objects, [120, 120], axes, 1).id, objects[2].id);
  const top = { ...objects[0], id: 'top' };
  assert.equal(pickOwnedInk([objects[0], top], [90, 50], axes, 1).id, 'top');
});
test('pick tolerance is six display pixels at ordinary zoom and protected/unmanaged/other types never establish authority', () => {
  assert.equal(pickOwnedInk([objects[0]], [85, 50], axes, 1)?.id, objects[0].id);
  assert.equal(pickOwnedInk([objects[0]], [85, 50], axes, 2), null);
  for (const patch of [
    { managed: false },
    { canEdit: false },
    { canDelete: false },
    { protectedReason: 'Imported' },
    { kind: 'shape' },
    { stroke: undefined },
  ])
    assert.equal(pickOwnedInk([{ ...objects[0], ...patch }], [90, 50], axes, 1), null);
  assert.throws(() => pickOwnedInk(objects, [90, 50], axes, 0), /zoom/);
});
test('movement emits only captured placement and source revision; immutable vector geometry cannot be resized', () => {
  const o = objects[0],
    box = axes.box(o.anchor),
    anchor = axes.anchor({ ...box, left: box.left + 20, top: box.top + 10 }),
    c = objectPlacementCommand('Sheet1', o, anchor);
  assert.deepEqual(c, {
    operation: 'updateInkStroke',
    sheet: 'Sheet1',
    objectId: o.id,
    expectedObjectRevision: 'r1',
    anchor,
  });
  assert.throws(() => objectGestureBox(box, 'resize', 10, 10, o), /resizing/);
  assert.throws(() => objectPlacementCommand('Sheet1', o, { ...anchor, widthEmu: anchor.widthEmu + 1 }), /geometry/);
});
