import './support/scratch-dir.mjs';
/** Actual source layer/controller rendering and scoped React pointer events. Native/app proof is separate. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
const component = fileURLToPath(new URL('../', import.meta.url)),
  out = await fs.mkdtemp(path.join(component, '.build/tests/worksheet-object-ui-')),
  bundle = path.join(out, 'ui.mjs');
await build({
  stdin: {
    contents:
      'export {WorksheetObjectLayer} from "./ui/WorksheetObjectLayer"; export {createObjectAxes} from "./ui/worksheetObjectGeometry";',
    resolveDir: component,
    loader: 'tsx',
  },
  bundle: true,
  write: true,
  outfile: bundle,
  format: 'esm',
  platform: 'node',
  packages: 'external',
  jsx: 'automatic',
  logLevel: 'silent',
});
const dom = new JSDOM('<div id="root"></div>', { url: 'http://localhost' });
Object.assign(globalThis, {
  window: dom.window,
  document: dom.window.document,
  HTMLElement: dom.window.HTMLElement,
  Element: dom.window.Element,
  Node: dom.window.Node,
  IS_REACT_ACT_ENVIRONMENT: true,
});
const { createRoot } = await import('react-dom/client'),
  { WorksheetObjectLayer, createObjectAxes } = await import(pathToFileURL(bundle));
after(async () => {
  dom.window.close();
});
const E = 9525,
  anchor = { kind: 'oneCell', cell: 'B2', dxEmu: 0, dyEmu: 0, widthEmu: 100 * E, heightEmu: 50 * E },
  chart = {
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
    data: { categories: ['A'], series: [{ name: 'Amount', values: [2] }], hidden: [false] },
  },
  png = {
    ...chart,
    id: '22222222-2222-4222-8222-222222222222',
    name: 'Logo',
    kind: 'image',
    image: { mime: 'image/png', base64: 'fixture', widthPx: 2, heightPx: 1 },
  };
const axes = createObjectAxes({
    sheetFormat: { defaultRowHeight: 15, rowsHiddenByDefault: false },
    widths: [{ min: 1, max: 8, width: 8.43, hidden: false }],
  }),
  view = { scale: 1, rowHeaderWidth: 48, columnHeaderHeight: 32, scrollLeft: 0, scrollTop: 0, width: 800, height: 500 };
const tick = () => new Promise((r) => setTimeout(r, 0)),
  flush = async () => {
    for (let n = 0; n < 3; n++) await act(async () => tick());
  },
  props = (n) => Object.entries(n).find(([key]) => key.startsWith('__reactProps$'))[1];
const node = (id) => document.querySelector('[data-worksheet-object="' + id + '"]'),
  button = (label) => [...document.querySelectorAll('button')].find((n) => n.getAttribute('aria-label') === label);
async function pointer(target, type, x = 0, y = 0) {
  assert(target);
  await act(async () => {
    const e = new window.MouseEvent(type, { bubbles: true, cancelable: true, button: 0, clientX: x, clientY: y });
    Object.defineProperty(e, 'pointerId', { value: 1 });
    target.dispatchEvent(e);
  });
  await flush();
}
async function key(target, name, patch = {}) {
  await act(async () =>
    target.dispatchEvent(new window.KeyboardEvent('keydown', { key: name, bubbles: true, cancelable: true, ...patch }))
  );
  await flush();
}
async function mount(t, patch = {}) {
  const root = createRoot(document.getElementById('root')),
    writes = [],
    extents = [];
  let args = {
    keyValue: 'owner/head/epoch1',
    family: 'owner/sheet',
    gestureKey: 'geometry/view',
    sheet: 'Data',
    active: true,
    controls: {
      objects: async () => ({ objects: [chart, png], total: 2, nextCursor: null }),
      object: async (_s, id) => (id === png.id ? png : chart),
      onBatch: async (commands) => writes.push(commands),
    },
    axes,
    view,
    onExtent: (family, e) => extents.push({ family, ...e }),
    ...patch,
  };
  const render = async (p) => {
    args = { ...args, ...p };
    await act(async () => root.render(React.createElement(WorksheetObjectLayer, args)));
    await flush();
  };
  await render({});
  t.after(async () => act(async () => root.unmount()));
  return { writes, extents, render };
}
test('chart/PNG render at source coordinates; embedded preview has no gallery controls or data-table overlay', async (t) => {
  const ui = await mount(t);
  assert.equal(node(chart.id).style.left, '107px');
  assert.equal(node(chart.id).style.top, '52px');
  assert(node(chart.id).querySelector('svg'));
  assert(node(png.id).querySelector('img').src.startsWith('data:image/png;base64,'));
  assert.equal(document.querySelector('details'), null);
  assert.equal(document.querySelector('table'), null);
  assert.deepEqual(ui.writes, []);
  assert(ui.extents.some((e) => e.rows >= 30));
});
test('pointer drafts and Escape cancel without acceptance; one finished move uses the captured original revision', async (t) => {
  const ui = await mount(t),
    target = node(chart.id);
  await pointer(target, 'pointerdown', 120, 65);
  await pointer(target, 'pointermove', 145, 75);
  assert.equal(target.style.left, '132px');
  assert.deepEqual(ui.writes, []);
  await key(target, 'Escape');
  assert.equal(target.style.left, '107px');
  await pointer(target, 'pointerdown', 120, 65);
  await pointer(target, 'pointerup', 145, 75);
  assert.deepEqual(ui.writes, [
    [
      {
        operation: 'updateChart',
        sheet: 'Data',
        objectId: chart.id,
        expectedObjectRevision: 'r1',
        patch: {},
        anchor: { ...anchor, dxEmu: 25 * E, dyEmu: 10 * E },
      },
    ],
  ]);
});
test('chart resizes both dimensions while PNG retains intrinsic proportions', async (t) => {
  const ui = await mount(t);
  await pointer(node(chart.id), 'pointerdown');
  await pointer(node(chart.id), 'pointerup');
  await pointer(button('Resize object: Sales'), 'pointerdown', 200, 100);
  await pointer(button('Resize object: Sales'), 'pointerup', 230, 110);
  assert.equal(ui.writes[0][0].anchor.widthEmu, 130 * E);
  assert.equal(ui.writes[0][0].anchor.heightEmu, 60 * E);
  await pointer(node(png.id), 'pointerdown');
  await pointer(node(png.id), 'pointerup');
  await pointer(button('Resize object: Logo'), 'pointerdown', 200, 100);
  await pointer(button('Resize object: Logo'), 'pointerup', 150, 100);
  assert.equal(ui.writes[1][0].anchor.widthEmu, 50 * E);
  assert.equal(ui.writes[1][0].anchor.heightEmu, 25 * E);
  assert(!('asset' in ui.writes[1][0]));
});
test('zoom/scroll affect presentation and pointer deltas once, not stored anchor sizes', async (t) => {
  const ui = await mount(t, { view: { ...view, scale: 2, scrollLeft: 30, scrollTop: 10 } }),
    target = node(chart.id);
  assert.equal(target.style.left, '136px');
  assert.equal(target.style.top, '62px');
  assert.equal(target.style.width, '200px');
  await pointer(target, 'pointerdown', 150, 80);
  await pointer(target, 'pointerup', 190, 100);
  assert.equal(ui.writes[0][0].anchor.dxEmu, 20 * E);
  assert.equal(ui.writes[0][0].anchor.dyEmu, 10 * E);
  assert.equal(ui.writes[0][0].anchor.widthEmu, 100 * E);
});
test('source/head/geometry retirement cancels a draft and stale pointer callback cannot retarget', async (t) => {
  const ui = await mount(t),
    target = node(chart.id),
    old = props(target).onPointerUp;
  await pointer(target, 'pointerdown', 120, 60);
  await pointer(target, 'pointermove', 140, 80);
  await ui.render({ keyValue: 'owner/new-head/epoch2' });
  await act(async () => old({ pointerId: 1, clientX: 180, clientY: 80, preventDefault() {}, stopPropagation() {} }));
  assert.deepEqual(ui.writes, []);
  await pointer(node(chart.id), 'pointerdown', 120, 60);
  await ui.render({ gestureKey: 'new-zoom' });
  await pointer(node(chart.id), 'pointerup', 160, 80);
  assert.deepEqual(ui.writes, []);
});
test('readonly or protected controls cannot accept a pointer or delete command', async (t) => {
  const ui = await mount(t, { reason: 'This workbook is read-only.' });
  await pointer(node(chart.id), 'pointerdown', 120, 60);
  await pointer(node(chart.id), 'pointerup', 160, 80);
  await key(node(chart.id), 'Delete');
  assert.equal(button('Delete object: Sales').disabled, true);
  assert.deepEqual(ui.writes, []);
});
test('keyboard move/delete is object-specific; object clipboard is refused without cell commands', async (t) => {
  const ui = await mount(t);
  await pointer(node(chart.id), 'pointerdown');
  await pointer(node(chart.id), 'pointerup');
  await key(node(chart.id), 'ArrowRight', { shiftKey: true });
  assert.equal(ui.writes[0][0].anchor.dxEmu, 10 * E);
  await key(node(chart.id), 'c', { ctrlKey: true });
  assert.equal(ui.writes.length, 1);
  assert.match(document.body.textContent, /clipboard/);
  await key(node(chart.id), 'Delete');
  assert.deepEqual(ui.writes[1], [
    { operation: 'deleteObject', sheet: 'Data', objectId: chart.id, expectedObjectRevision: 'r1' },
  ]);
});
test('inactive/replaced owner immediately masks old objects and joined reads never populate the next workbook', async (t) => {
  const ui = await mount(t);
  await ui.render({ active: false });
  assert.equal(document.querySelectorAll('[data-worksheet-object]').length, 0);
  await ui.render({
    keyValue: 'other/head/epoch2',
    family: 'other/sheet',
    active: true,
    controls: {
      objects: async () => ({ objects: [], total: 0, nextCursor: null }),
      object: async () => chart,
      onBatch: async () => {
        throw Error('No writes');
      },
    },
  });
  assert.equal(document.querySelectorAll('[data-worksheet-object]').length, 0);
  assert.deepEqual(ui.writes, []);
});
