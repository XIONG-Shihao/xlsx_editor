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
import { WORDART_SAMPLES } from '../core/wordart-samples.mjs';
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
test('three WordArt samples render real bold 32 pt text with distinct saved-profile effects', () => {
  for (const s of WORDART_SAMPLES) {
    const html = renderToStaticMarkup(
        React.createElement(WorkbookShapePreview, {
          shape: { sample: s.id, text: 'Title &é\nNext' },
          width: 320,
          height: 120,
        })
      ),
      doc = new JSDOM(html).window.document,
      n = doc.querySelector('foreignObject div div');
    assert.equal(n.textContent, 'Title &é\nNext');
    assert.equal(n.style.fontSize, '32pt');
    assert.equal(n.style.fontWeight, '700');
    assert.equal(doc.querySelector('img,image'), null);
    if (s.outline) assert.match(html, /-webkit-text-stroke:1.333333px #C65911/);
    if (s.shadow) assert.equal(n.style.textShadow, '4px 4px 2px rgba(0,0,0,0.35)');
    assert(!/(NaN|Infinity)/.test(html));
  }
});
test('WordArt gallery keeps the text draft while switching styles and creates all three captured profiles', async (t) => {
  const ui = await mount(t, { mode: 'wordart' });
  for (const s of WORDART_SAMPLES) {
    await ui.render({ frame: { ...ui.args.frame, key: s.id }, mode: 'wordart' });
    await input('Object text', 'Title "&é\nNext');
    await click(s.label);
    assert.equal(document.querySelector('textarea').value, 'Title "&é\nNext');
    await submit();
    const c = ui.writes.at(-1)[0];
    assert.deepEqual(c.shape, { sample: s.id, text: 'Title "&é\nNext' });
    assert.equal(c.anchor.widthEmu, 320 * E);
    assert.equal(c.anchor.heightEmu, 120 * E);
    assert.equal(c.expectedObjectRevision, null);
  }
  assert.equal(ui.writes.length, 3);
});
test('editing WordArt text/style/alt keeps original object revision, dimensions and offsets', async (t) => {
  const art = { ...shape, name: 'Title', shape: { sample: 'wordart-blue', text: 'Title' } },
    writes = [],
    f = frame({
      controls: {
        objects: async () => ({ objects: [art], total: 1, nextCursor: null }),
        object: async () => art,
        onBatch: async (c) => writes.push(c),
      },
    });
  await mount(t, { frame: f, mode: 'wordart' });
  await click('Blue Fill — B2');
  await click('Purple Shadow');
  await input('Object text', 'Edited');
  await input('Object alternative text', 'Accessible title');
  await submit();
  assert.deepEqual(writes[0][0].patch, { sample: 'wordart-shadow', text: 'Edited' });
  assert.equal(writes[0][0].expectedObjectRevision, 'r1');
  assert.equal(writes[0][0].anchor.dxEmu, 3 * E);
  assert.equal(writes[0][0].altText, 'Accessible title');
  await click('Blue Fill — B2');
  await click('Delete object');
  assert.equal(writes[1][0].operation, 'deleteShape');
});
test('WordArt bounds reject long/many-paragraph/unknown effects without an owner write', async (t) => {
  const ui = await mount(t, { mode: 'wordart' });
  assert.equal(document.querySelector('textarea').maxLength, 256);
  await input('Object text', 'x'.repeat(257));
  await submit();
  assert.match(document.body.textContent, /256 characters/);
  await input('Object text', 'a\nb\nc\nd\ne');
  await submit();
  assert.equal(ui.writes.length, 0);
  for (const v of [
    { sample: 'wordart-blue', text: 'Title', shadow: 99 },
    { sample: 'wordart-shadow', text: 'Title', color: 'FFFFFF' },
  ])
    assert.throws(() => normalizeShape(v));
});
test('retired WordArt drafts and readonly panes cannot accept a stale callback', async (t) => {
  const ui = await mount(t, { mode: 'wordart' });
  await input('Object text', 'Draft');
  const old = props(document.querySelector('form')).onSubmit;
  await ui.render({ frame: { ...ui.args.frame, key: 'other' }, selection: 'C4' });
  assert.equal(document.querySelector('textarea').value, 'WordArt');
  await act(async () => old({ preventDefault() {} }));
  assert.equal(ui.writes.length, 0);
  await ui.render({ frame: { ...ui.args.frame, key: 'readonly', reason: 'This workbook is read-only.' } });
  await submit();
  assert.equal(ui.writes.length, 0);
  assert(document.querySelector('fieldset').disabled);
  for (const label of ['3D Models']) assert.equal(button(label), undefined);
});
