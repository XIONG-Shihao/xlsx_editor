import './support/scratch-dir.mjs';
/** Actual React source pointer handlers with controlled DOM geometry; native/app proof is separate. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
import { inkFoundationCases } from './support/ink-foundation-fixture.mjs';
const out = await fs.mkdtemp(path.resolve('.build/tests/dw-m2-ui-')),
  bundle = path.join(out, 'ui.mjs');
await build({
  stdin: {
    contents:
      'export {WorksheetObjectLayer} from "./ui/WorksheetObjectLayer";export {createObjectAxes} from "./ui/worksheetObjectGeometry";',
    resolveDir: process.cwd(),
    loader: 'tsx',
  },
  bundle: true,
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
  })),
  view = { scale: 1, rowHeaderWidth: 48, columnHeaderHeight: 32, scrollLeft: 0, scrollTop: 0, width: 800, height: 400 };
const tick = () => new Promise((r) => setTimeout(r, 0)),
  flush = async () => {
    for (let n = 0; n < 3; n++) await act(async () => tick());
  },
  surface = () => document.querySelector('.sheet-ink-edit-surface'),
  props = (n) => Object.entries(n).find(([k]) => k.startsWith('__reactProps$'))[1];
async function pointer(type, x = 138, y = 82) {
  const node = surface();
  assert(node);
  await act(async () => {
    const e = new window.MouseEvent(type, { bubbles: true, cancelable: true, button: 0, clientX: x, clientY: y });
    Object.defineProperties(e, { pointerId: { value: 1 }, isPrimary: { value: true } });
    node.dispatchEvent(e);
  });
  await flush();
}
async function mount(t, patch = {}) {
  const container = document.createElement('div');
  document.getElementById('root').append(container);
  const root = createRoot(container),
    writes = [];
  let args = {
    keyValue: 'owner/head/activation',
    family: 'owner/sheet',
    gestureKey: 'geometry/mode',
    sheet: 'Sheet1',
    active: true,
    inkMode: 'select',
    axes,
    view,
    onExtent: () => {},
    controls: {
      objects: async () => ({ objects, total: 3, nextCursor: null }),
      object: async (_s, id) => objects.find((o) => o.id === id),
      onBatch: async (commands) => writes.push(commands),
    },
    ...patch,
  };
  const render = async (p) => {
    args = { ...args, ...p };
    await act(async () => root.render(React.createElement(WorksheetObjectLayer, args)));
    await flush();
  };
  await render({});
  t.after(async () => {
    await act(async () => root.unmount());
    container.remove();
  });
  return { writes, render };
}
after(async () => {
  dom.window.close();
});
test('clicking a path selects it without a write; rectangle whitespace selects nothing', async (t) => {
  const h = await mount(t);
  await pointer('pointerdown');
  await pointer('pointerup');
  assert(document.querySelector('[data-object-selected="true"]'));
  assert(document.querySelector('[aria-label="Delete stroke"]'));
  assert(document.querySelector('[aria-label="Resize object: Black zigzag"]').disabled);
  assert.equal(h.writes.length, 0);
  await pointer('pointerdown', 143, 110);
  await pointer('pointerup', 143, 110);
  assert.equal(document.querySelector('[data-object-selected="true"]'), null);
  assert.equal(h.writes.length, 0);
});
test('a move drafts source placement then emits one immutable-stroke command on release', async (t) => {
  const h = await mount(t);
  await pointer('pointerdown');
  await pointer('pointermove', 158, 92);
  assert.equal(h.writes.length, 0);
  await pointer('pointerup', 158, 92);
  assert.equal(h.writes.length, 1);
  const c = h.writes[0][0];
  assert.equal(c.operation, 'updateInkStroke');
  assert.equal(c.objectId, objects[0].id);
  assert.equal(c.expectedObjectRevision, 'r1');
  assert.equal(c.anchor.widthEmu, objects[0].stroke.widthEmu);
  assert(!('stroke' in c));
});
test('selected Delete button and keyboard Delete use only the selected stroke', async (t) => {
  const h = await mount(t);
  await pointer('pointerdown');
  await pointer('pointerup');
  await act(async () => document.querySelector('[aria-label="Delete stroke"]').click());
  await flush();
  assert.equal(h.writes[0][0].operation, 'deleteInkStroke');
  assert.equal(h.writes[0][0].objectId, objects[0].id);
  await pointer('pointerdown');
  await pointer('pointerup');
  await act(async () => surface().dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Delete', bubbles: true })));
  await flush();
  assert.equal(h.writes[1][0].operation, 'deleteObject');
});
test('whole eraser waits for release; Escape and pointercancel retire its target', async (t) => {
  const h = await mount(t, { inkMode: 'erase' });
  await pointer('pointerdown');
  assert.equal(h.writes.length, 0);
  await act(async () => document.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true })));
  await pointer('pointerup');
  assert.equal(h.writes.length, 0);
  await pointer('pointerdown');
  await pointer('pointercancel');
  await pointer('pointerup');
  assert.equal(h.writes.length, 0);
  await pointer('pointerdown');
  await pointer('pointerup');
  assert.equal(h.writes.length, 1);
  assert.equal(h.writes[0][0].operation, 'deleteInkStroke');
});
test('owner/head, geometry, mode and activation changes discard captured movement/erase without retargeting', async (t) => {
  for (const patch of [
    { keyValue: 'replacement/head' },
    { gestureKey: 'new-scroll' },
    { inkMode: null },
    { active: false },
  ]) {
    const h = await mount(t);
    await pointer('pointerdown');
    const finish = props(surface()).onPointerUp;
    await h.render(patch);
    await act(async () =>
      finish({ pointerId: 1, clientX: 180, clientY: 90, preventDefault() {}, stopPropagation() {} })
    );
    assert.equal(h.writes.length, 0);
    await h.render({ active: false });
  }
});
test('zoom/scroll moves scale once and readonly/busy reason refuses before callback', async (t) => {
  const h = await mount(t, { view: { ...view, scale: 2, scrollLeft: 30, scrollTop: 10 } });
  await pointer('pointerdown', 198, 122);
  await pointer('pointerup', 238, 142);
  assert.equal(h.writes.length, 1);
  const b = axes.box(h.writes[0][0].anchor),
    prior = axes.box(objects[0].anchor);
  assert.equal(b.left - prior.left, 20);
  assert.equal(b.top - prior.top, 10);
  await h.render({ reason: 'Read-only' });
  await pointer('pointerdown', 198, 122);
  await pointer('pointerup', 238, 142);
  assert.equal(h.writes.length, 1);
});
test('one physical write lane remains reserved while accepted work retires; late errors stay in the original owner', async (t) => {
  let done;
  const pending = new Promise((r) => (done = r)),
    writes = [];
  const h = await mount(t, {
    controls: {
      objects: async () => ({ objects, total: 3, nextCursor: null }),
      object: async (_s, id) => objects.find((o) => o.id === id),
      onBatch: async (c) => {
        writes.push(c);
        await pending;
      },
    },
  });
  await pointer('pointerdown');
  await pointer('pointerup', 158, 92);
  await pointer('pointerdown');
  await pointer('pointerup', 168, 102);
  assert.equal(writes.length, 1);
  await h.render({ keyValue: 'replacement' });
  done();
  await flush();
  assert.equal(writes.length, 1);
});

test('Draw editing keys never act on a previously selected chart', async (t) => {
  const chart = {
    ...objects[0],
    id: 'chart',
    name: 'Other chart',
    kind: 'chart',
    stroke: undefined,
    chart: { kind: 'column', range: 'A1:B2' },
    data: { categories: ['A'], series: [{ name: 'Sales', values: [1] }], hidden: [false] },
  };
  const h = await mount(t, {
    inkMode: null,
    controls: {
      objects: async () => ({ objects: [chart], total: 1, nextCursor: null }),
      object: async () => chart,
      onBatch: async (c) => h.writes.push(c),
    },
  });
  await act(async () => document.querySelector('[data-worksheet-object="chart"]').focus());
  await h.render({ inkMode: 'select', gestureKey: 'ink-mode' });
  await act(async () => surface().dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Delete', bubbles: true })));
  await flush();
  assert.equal(h.writes.length, 0);
});

test('changing the selected stroke retires the original movement draft without accepting or retargeting', async (t) => {
  const h = await mount(t);
  await pointer('pointerdown');
  await pointer('pointermove', 158, 92);
  await act(async () => props(document.querySelector('[data-worksheet-object="' + objects[1].id + '"]')).onFocus());
  await pointer('pointerup', 178, 102);
  assert.equal(h.writes.length, 0);
});
