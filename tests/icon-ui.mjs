import './support/scratch-dir.mjs';
/** Real React controls and shared controller/worksheet rendering; native proof is separate. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { JSDOM } from 'jsdom';
import { ICON_SAMPLES } from '../core/icon-samples.mjs';
import { normalizeShape } from '../core/shape-model.mjs';
const out = await fs.mkdtemp(path.resolve('.build/tests/text-shape-ui-')),
  bundle = path.join(out, 'ui.mjs');
await build({
  stdin: {
    contents:
      'export {WorkbookShapes} from "./ui/WorkbookShapes"; export {WorkbookShapePreview} from "./ui/WorkbookShapePreview"; export {WorksheetObjectLayer} from "./ui/WorksheetObjectLayer"; export {createObjectAxes} from "./ui/worksheetObjectGeometry";',
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
const { WorkbookShapes, WorkbookShapePreview, WorksheetObjectLayer, createObjectAxes } = await import(
    pathToFileURL(bundle)
  ),
  dom = new JSDOM('<div id="root"></div>', { url: 'http://localhost' });
Object.assign(globalThis, {
  window: dom.window,
  document: dom.window.document,
  HTMLElement: dom.window.HTMLElement,
  Element: dom.window.Element,
  Node: dom.window.Node,
  IS_REACT_ACT_ENVIRONMENT: true,
});
const { createRoot } = await import('react-dom/client');
const E = 9525,
  id = '11111111-1111-4111-8111-111111111111',
  anchor = { kind: 'oneCell', cell: 'B2', dxEmu: 3 * E, dyEmu: 4 * E, widthEmu: 240 * E, heightEmu: 120 * E },
  shape = {
    id,
    name: 'Editable text',
    kind: 'shape',
    revision: 'r1',
    anchor,
    managed: true,
    canEdit: true,
    canDelete: true,
    protectedReason: null,
    altText: 'Original alt',
    shape: { sample: 'shape-rectangle', text: 'Editable text' },
  };
const props = (n) => Object.entries(n).find(([k]) => k.startsWith('__reactProps$'))[1],
  button = (s) =>
    [...document.querySelectorAll('button')].find((n) => n.textContent === s || n.getAttribute('aria-label') === s),
  flush = async () => {
    for (let n = 0; n < 4; n++) await act(async () => new Promise((r) => setTimeout(r, 0)));
  },
  click = async (s) => {
    assert(button(s), s);
    await act(async () => button(s).click());
    await flush();
  },
  input = async (label, value) => {
    const n = document.querySelector('[aria-label="' + label + '"]');
    assert(n);
    await act(async () => props(n).onChange({ target: { value } }));
  },
  submit = async () => {
    await act(async () => props(document.querySelector('form')).onSubmit({ preventDefault() {} }));
    await flush();
  },
  deferred = () => {
    let resolve;
    const promise = new Promise((r) => (resolve = r));
    return { promise, resolve };
  };
const frame = (patch = {}) => ({
  key: 'head1/selection1/pane1',
  family: 'owner/sheet',
  sheet: 'Data',
  active: true,
  controls: {
    objects: async () => ({ objects: [], total: 0, nextCursor: null }),
    object: async () => shape,
    onBatch: async () => {},
  },
  ...patch,
});
async function mount(t, patch = {}) {
  const root = createRoot(document.getElementById('root')),
    writes = [];
  let args = {
    frame: frame({ controls: { ...frame().controls, onBatch: async (c) => writes.push(c) } }),
    mode: 'shape',
    selection: 'B2',
    ...patch,
  };
  const render = async (p) => {
    args = { ...args, ...p };
    await act(async () => root.render(React.createElement(WorkbookShapes, args)));
    await flush();
  };
  await render({});
  t.after(async () => act(async () => root.unmount()));
  return { writes, render, args };
}
after(async () => {
  dom.window.close();
});
test('three real icon paths render solid editable color without raster or font/text substitute', () => {
  for (const s of ICON_SAMPLES) {
    const html = renderToStaticMarkup(
        React.createElement(WorkbookShapePreview, {
          shape: { sample: s.id, text: '', color: '2563EB' },
          width: 96,
          height: 96,
          label: 'Accessible ' + s.label,
        })
      ),
      doc = new JSDOM(html).window.document;
    assert.equal(doc.querySelector('svg').getAttribute('aria-label'), 'Accessible ' + s.label);
    assert.equal(doc.querySelector('polygon').getAttribute('points'), s.points.map((p) => p.join(',')).join(' '));
    assert.equal(doc.querySelector('polygon').getAttribute('fill'), '#2563EB');
    assert.equal(doc.querySelector('polygon').getAttribute('transform'), 'scale(0.96 0.96)');
    assert.equal(doc.querySelector('img,image,foreignObject,text'), null);
    assert(!/(NaN|Infinity)/.test(html));
  }
});
test('Icons gallery creates all three captured color/placement/alt models at square default size', async (t) => {
  const ui = await mount(t, { mode: 'icon' });
  assert.equal(document.querySelector('textarea'), null);
  for (const s of ICON_SAMPLES) {
    await ui.render({ mode: 'icon', frame: { ...ui.args.frame, key: s.id } });
    await click(s.label);
    await input('Icon color', '#dc2626');
    await input('Object alternative text', 'My ' + s.label);
    await submit();
    const c = ui.writes.at(-1)[0];
    assert.equal(c.operation, 'createShape');
    assert.deepEqual(c.shape, { sample: s.id, text: '', color: 'DC2626' });
    assert.deepEqual(c.anchor, { cell: 'B2', dxEmu: 0, dyEmu: 0, widthEmu: 96 * E, heightEmu: 96 * E });
    assert.equal(c.altText, 'My ' + s.label);
  }
  assert.equal(ui.writes.length, 3);
});
test('existing icon recolor/sample replacement/delete retains captured identity and offsets', async (t) => {
  const icon = {
      ...shape,
      name: 'Check',
      shape: { sample: 'icon-check', text: '', color: '217346' },
      anchor: { ...anchor, widthEmu: 96 * E, heightEmu: 96 * E },
    },
    writes = [],
    f = frame({
      controls: {
        objects: async () => ({ objects: [icon], total: 1, nextCursor: null }),
        object: async () => icon,
        onBatch: async (c) => writes.push(c),
      },
    });
  await mount(t, { frame: f, mode: 'icon' });
  await click('Check — B2');
  await click('Star');
  await input('Icon color', '#7c3aed');
  await submit();
  assert.deepEqual(writes[0][0].patch, { sample: 'icon-star', text: '', color: '7C3AED' });
  assert.equal(writes[0][0].expectedObjectRevision, 'r1');
  assert.equal(writes[0][0].objectId, id);
  assert.equal(writes[0][0].anchor.dxEmu, 3 * E);
  await click('Check — B2');
  await click('Delete object');
  assert.equal(writes[1][0].operation, 'deleteShape');
});
test('color drafts reset and retire without writes; readonly and later galleries remain clear', async (t) => {
  const ui = await mount(t, { mode: 'icon' });
  await input('Icon color', '#dc2626');
  await click('Reset form');
  assert.equal(document.querySelector('[aria-label="Icon color"]').value, '#217346');
  const old = props(document.querySelector('form')).onSubmit;
  await ui.render({ frame: { ...ui.args.frame, key: 'other-cell' }, selection: 'D4' });
  await act(async () => old({ preventDefault() {} }));
  assert.equal(ui.writes.length, 0);
  await ui.render({ frame: { ...ui.args.frame, key: 'readonly', reason: 'This workbook is read-only.' } });
  await submit();
  assert(document.querySelector('fieldset').disabled);
  assert.equal(ui.writes.length, 0);
  for (const label of ['3D Models']) assert.equal(button(label), undefined);
});
test('closed icon color/text profiles cannot quietly become a different model', () => {
  for (const value of [
    { sample: 'icon-check', text: 'text', color: '217346' },
    { sample: 'icon-star', text: '', color: 'abc123' },
    { sample: 'icon-plus', text: '', color: '#217346' },
    { sample: 'icon-plus', text: '', path: [1, 2] },
    { sample: 'shape-rectangle', text: 'Normal', color: '217346' },
  ])
    assert.throws(() => normalizeShape(value));
  assert.deepEqual(normalizeShape({ sample: 'icon-star' }), { sample: 'icon-star', text: '', color: '217346' });
});
