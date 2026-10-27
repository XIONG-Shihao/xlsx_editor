import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  createObjectAxes,
  objectScreenBox,
  objectGestureBox,
  objectPlacementCommand,
  OBJECT_EMU_PER_PIXEL as E,
} from '../ui/worksheetObjectGeometry.ts';
import { createWorksheetObjectController } from '../ui/worksheetObjectController.ts';
const sheet = {
    sheetFormat: { defaultRowHeight: 15, rowsHiddenByDefault: false },
    widths: [{ min: 1, max: 8, width: 8.43, hidden: false }],
  },
  axes = createObjectAxes(sheet),
  anchor = { kind: 'oneCell', cell: 'B2', dxEmu: 10 * E, dyEmu: 5 * E, widthEmu: 100 * E, heightEmu: 50 * E };
const object = {
  id: '11111111-1111-4111-8111-111111111111',
  name: 'Sales',
  kind: 'chart',
  revision: 'r1',
  anchor,
  managed: true,
  canEdit: true,
  canDelete: true,
  protectedReason: null,
  chart: { kind: 'column', range: 'A1:B3' },
};
const preview = { ...object, data: { categories: ['A'], series: [{ name: 'Amount', values: [2] }], hidden: [false] } };
const deferred = () => {
    let resolve, reject;
    const promise = new Promise((yes, no) => {
      resolve = yes;
      reject = no;
    });
    return { promise, resolve, reject };
  },
  tick = () => new Promise((r) => setTimeout(r, 0));
const frame = (patch = {}) => ({
  key: 'owner/head/epoch1',
  family: 'owner/sheet',
  sheet: 'Data',
  active: true,
  controls: {
    objects: async () => ({ objects: [object], total: 1, nextCursor: null }),
    object: async () => preview,
    onBatch: async () => {},
  },
  ...patch,
});
async function ready(controller) {
  for (let n = 0; n < 20 && controller.getSnapshot().loading; n++) await tick();
  assert.equal(controller.getSnapshot().loading, false);
  return controller.getSnapshot();
}
test('source axes round-trip canonical cell/offsets and project ordinary zoom/scroll/header visibility', () => {
  assert.deepEqual(axes.box(anchor), { left: 69, top: 25, width: 100, height: 50 });
  assert.deepEqual(axes.anchor(axes.box(anchor)), anchor);
  assert.deepEqual(
    objectScreenBox(axes.box(anchor), {
      scale: 1.5,
      rowHeaderWidth: 72,
      columnHeaderHeight: 48,
      scrollLeft: 30,
      scrollTop: 10,
    }),
    { left: 145.5, top: 75.5, width: 150, height: 75 }
  );
  assert.deepEqual(
    objectScreenBox(axes.box(anchor), {
      scale: 2,
      rowHeaderWidth: 0,
      columnHeaderHeight: 0,
      scrollLeft: 80,
      scrollTop: 30,
    }),
    { left: 58, top: 20, width: 200, height: 100 }
  );
  const far = axes.box({ ...anchor, cell: 'B300' });
  assert(axes.extent(far).rows >= 302);
  assert.equal(far.top, 5985);
});
test('hidden/nonuniform source axes map to visible canonical markers and enforce drawing bounds', () => {
  const a = createObjectAxes({
    ...sheet,
    hiddenRows: [1],
    rowHeights: [{ row: 2, height: 30 }],
    widths: [
      { min: 1, max: 1, width: 8.43, hidden: false },
      { min: 2, max: 2, width: 8.43, hidden: true },
      { min: 3, max: 3, width: 12, hidden: false },
    ],
  });
  const moved = a.anchor({ left: 64, top: 25, width: 100, height: 50 });
  assert.equal(moved.cell, 'C3');
  assert.equal(moved.dxEmu, 5 * E);
  assert.equal(moved.dyEmu, 5 * E);
  assert.deepEqual(a.box(moved), { left: 64, top: 25, width: 100, height: 50 });
  for (const box of [
    { left: -1, top: 0, width: 1, height: 1 },
    { left: 0, top: 0, width: 0, height: 1 },
    { left: 0, top: 0, width: 6000, height: 1 },
    { left: Infinity, top: 0, width: 1, height: 1 },
  ])
    assert.throws(() => axes.anchor(box));
});
test('chart resize has independent dimensions; PNG expands and shrinks proportionally from either dominant axis', () => {
  const start = { left: 10, top: 20, width: 200, height: 100 };
  assert.deepEqual(objectGestureBox(start, 'move', -30, 5, preview), { ...start, left: 0, top: 25 });
  assert.deepEqual(objectGestureBox(start, 'resize', 30, -20, preview), { ...start, width: 230, height: 80 });
  const image = { ...preview, kind: 'image', image: { mime: 'image/png', base64: 'x', widthPx: 2, heightPx: 1 } };
  assert.deepEqual(objectGestureBox(start, 'resize', -100, 0, image), { ...start, width: 100, height: 50 });
  assert.deepEqual(objectGestureBox(start, 'resize', 0, -50, image), { ...start, width: 100, height: 50 });
  assert.deepEqual(objectGestureBox(start, 'resize', 100, 0, image), { ...start, width: 300, height: 150 });
});
test('placement commands target exactly the captured object without cell commands or replacing image bytes', () => {
  assert.deepEqual(objectPlacementCommand('Data', preview, anchor), {
    operation: 'updateChart',
    sheet: 'Data',
    objectId: preview.id,
    expectedObjectRevision: 'r1',
    patch: {},
    anchor,
  });
  assert.deepEqual(objectPlacementCommand('Data', { ...preview, kind: 'image' }, anchor), {
    operation: 'updateImage',
    sheet: 'Data',
    objectId: preview.id,
    expectedObjectRevision: 'r1',
    anchor,
  });
  assert.throws(() => objectPlacementCommand('Data', { ...preview, canEdit: false }, anchor));
});
test('one physical read joins retirement before the latest owner; late results never populate another frame', async () => {
  const c = createWorksheetObjectController(),
    old = deferred(),
    calls = [];
  c.setFrame(
    frame({
      controls: {
        ...frame().controls,
        objects: () => {
          calls.push('old');
          return old.promise;
        },
      },
    })
  );
  c.setFrame(
    frame({
      key: 'new',
      controls: {
        ...frame().controls,
        objects: async () => {
          calls.push('new');
          return { objects: [], total: 0, nextCursor: null };
        },
      },
    })
  );
  assert.deepEqual(calls, ['old']);
  old.resolve({ objects: [object], total: 1, nextCursor: null });
  await ready(c);
  assert.deepEqual(calls, ['old', 'new']);
  assert.equal(c.getSnapshot().key, 'new');
  assert.deepEqual(c.getSnapshot().objects, []);
  c.dispose();
});
test('one captured write is awaited; late acceptance retains original truth but cannot become current', async () => {
  const c = createWorksheetObjectController(),
    gate = deferred(),
    writes = [],
    f = frame({
      controls: {
        ...frame().controls,
        onBatch: (commands) => {
          writes.push(commands);
          return gate.promise;
        },
      },
    });
  c.setFrame(f);
  await ready(c);
  const command = objectPlacementCommand('Data', preview, anchor),
    work = c.accept(f, command);
  assert.equal(c.getSnapshot().pending, true);
  await assert.rejects(c.accept(f, command), /Wait/);
  c.setFrame(frame({ key: 'new' }));
  gate.resolve();
  assert.deepEqual(await work, { accepted: true, current: false, key: f.key });
  assert.deepEqual(writes, [[command]]);
  c.dispose();
});
test('foreign/stale/protected/busy commands refuse before an owner callback', async () => {
  const c = createWorksheetObjectController(),
    writes = [],
    f = frame({ controls: { ...frame().controls, onBatch: async (commands) => writes.push(commands) } });
  c.setFrame(f);
  await ready(c);
  const command = objectPlacementCommand('Data', preview, anchor);
  for (const patch of [{ sheet: 'Other' }, { objectId: 'foreign' }, { expectedObjectRevision: 'old' }])
    await assert.rejects(c.accept(f, { ...command, ...patch }));
  c.setFrame({ ...f, reason: 'Read only' });
  await assert.rejects(c.accept(f, command), /Read only/);
  c.setFrame(null);
  await assert.rejects(c.accept(f, command), /changed/);
  assert.deepEqual(writes, []);
  c.dispose();
});
test('malformed pages/previews cannot establish a drawing or editing grant', async () => {
  for (const controls of [
    { ...frame().controls, objects: async () => ({ objects: [object, object], total: 2, nextCursor: null }) },
    { ...frame().controls, object: async () => ({ ...preview, revision: 'changed' }) },
    { ...frame().controls, objects: async () => ({ objects: [], total: 300, nextCursor: null }) },
  ]) {
    const c = createWorksheetObjectController();
    c.setFrame(frame({ controls }));
    const snapshot = await ready(c);
    assert(snapshot.error);
    assert.deepEqual(snapshot.objects, []);
    c.dispose();
  }
});
