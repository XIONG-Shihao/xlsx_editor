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
import { SHAPE_SAMPLES } from '../core/shape-model.mjs';
import { createWorksheetObjectController } from '../ui/worksheetObjectController.ts';
import { objectPlacementCommand } from '../ui/worksheetObjectGeometry.ts';
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
test('all five samples render real geometry and wrapping text with constant point font and no bitmap', () => {
  for (const s of SHAPE_SAMPLES) {
    const html = renderToStaticMarkup(
        React.createElement(WorkbookShapePreview, {
          shape: { sample: s.id, text: 'First & <second>\nThird' },
          width: 360,
          height: 180,
        })
      ),
      doc = new JSDOM(html).window.document;
    assert.equal(doc.querySelector('svg').getAttribute('data-shape-sample'), s.id);
    assert.equal(doc.querySelector('img,image'), null);
    assert(doc.querySelector(s.geometry === 'ellipse' ? 'ellipse' : s.geometry === 'rightArrow' ? 'path' : 'rect'));
    assert.equal(doc.querySelector('foreignObject div div').style.fontSize, '14pt');
    assert.equal(doc.querySelector('foreignObject').textContent, 'First & <second>\nThird');
    assert(!/(NaN|Infinity)/.test(html));
    if (s.geometry === 'rightArrow') {
      const region = doc.querySelector('foreignObject');
      assert.equal(Number(region.getAttribute('y')), 49.8);
      assert.equal(Number(region.getAttribute('height')), 80.4);
      assert(Number(region.getAttribute('x')) + Number(region.getAttribute('width')) < 360 * 0.65);
    }
  }
});
test('gallery insertion for all five samples captures text, dimensions, selection and no bitmap data', async (t) => {
  const ui = await mount(t);
  for (const s of SHAPE_SAMPLES) {
    await ui.render({ mode: s.family, frame: { ...ui.args.frame, key: s.id } });
    await click(s.label);
    await input('Object text', 'Edited "&é\nNext');
    await input('Object width', '160');
    await input('Object height', '80');
    await input('Object alternative text', 'My ' + s.label);
    await submit();
    const c = ui.writes.at(-1)[0];
    assert.equal(c.operation, 'createShape');
    assert.equal(c.expectedObjectRevision, null);
    assert.deepEqual(c.shape, { sample: s.id, text: 'Edited "&é\nNext' });
    assert.deepEqual(c.anchor, { cell: 'B2', dxEmu: 0, dyEmu: 0, widthEmu: 160 * E, heightEmu: 80 * E });
    assert.equal(c.altText, 'My ' + s.label);
    assert(!('asset' in c));
  }
  assert.equal(ui.writes.length, 5);
});
test('edit keeps object revision and untouched offsets; remove uses same owned identity', async (t) => {
  const writes = [],
    f = frame({
      controls: {
        objects: async () => ({ objects: [shape], total: 1, nextCursor: null }),
        object: async () => shape,
        onBatch: async (c) => writes.push(c),
      },
    });
  await mount(t, { frame: f });
  await click('Rectangle — B2');
  await input('Object text', 'Changed text');
  await input('Object width', '300');
  await submit();
  assert.equal(writes[0][0].operation, 'updateShape');
  assert.equal(writes[0][0].expectedObjectRevision, 'r1');
  assert.deepEqual(writes[0][0].anchor, {
    cell: 'B2',
    dxEmu: 3 * E,
    dyEmu: 4 * E,
    widthEmu: 300 * E,
    heightEmu: 120 * E,
  });
  await click('Rectangle — B2');
  await click('Delete object');
  assert.deepEqual(writes[1], [
    { operation: 'deleteShape', sheet: 'Data', objectId: id, expectedObjectRevision: 'r1' },
  ]);
});
test('draft reset, invalid paragraphs and future galleries leave owners untouched', async (t) => {
  const ui = await mount(t);
  await input('Object text', 'Draft');
  await click('Reset form');
  assert.equal(document.querySelector('textarea').value, 'Rectangle');
  await input('Object text', Array(9).fill('line').join('\n'));
  await submit();
  assert.match(document.body.textContent, /eight paragraphs/);
  assert.equal(ui.writes.length, 0);
  for (const s of ['3D Models']) assert.equal(button(s), undefined);
});
test('closing/switching selection retires draft and stale callbacks; readonly blocks new changes', async (t) => {
  const ui = await mount(t);
  await input('Object text', 'Draft');
  const old = props(document.querySelector('form')).onSubmit;
  await ui.render({ frame: { ...ui.args.frame, key: 'head1/selection2/pane2' }, selection: 'C3' });
  assert.equal(document.querySelector('textarea').value, 'Rectangle');
  assert.equal(document.querySelector('[aria-label="Object anchor"]').value, 'C3');
  await act(async () => old({ preventDefault() {} }));
  assert.equal(ui.writes.length, 0);
  await ui.render({ frame: { ...ui.args.frame, key: 'closed', active: false } });
  await submit();
  assert.equal(ui.writes.length, 0);
  await ui.render({ frame: { ...ui.args.frame, key: 'readonly', reason: 'This workbook is read-only.' } });
  assert(document.querySelector('fieldset').disabled);
  await submit();
  assert.equal(ui.writes.length, 0);
});
test('physical controller write lane survives retirement; commands cloned before queue callbacks', async () => {
  const gate = deferred(),
    writes = [],
    c = createWorksheetObjectController(),
    f = frame({
      controls: {
        ...frame().controls,
        onBatch: async (batch) => {
          writes.push(batch);
          await gate.promise;
        },
      },
    });
  c.setFrame(f);
  for (let n = 0; n < 8 && c.getSnapshot().loading; n++) await new Promise((r) => setTimeout(r, 0));
  const cmd = {
    operation: 'createShape',
    sheet: 'Data',
    objectId: id,
    expectedObjectRevision: null,
    shape: { sample: 'text-plain', text: 'Captured' },
  };
  const run = c.accept(f, cmd);
  cmd.shape.text = 'Caller mutation';
  assert.equal(writes[0][0].shape.text, 'Captured');
  c.setFrame({ ...f, key: 'next' });
  await assert.rejects(c.accept({ ...f, key: 'next' }, cmd), /finish|inventory/);
  gate.resolve();
  assert.equal((await run).current, false);
  c.dispose();
});
test('shape placement command cannot impersonate chart and rejects wrong-kind controller edit', async () => {
  assert.equal(objectPlacementCommand('Data', shape, anchor).operation, 'updateShape');
  const c = createWorksheetObjectController(),
    f = frame({
      controls: {
        objects: async () => ({ objects: [shape], total: 1, nextCursor: null }),
        object: async () => shape,
        onBatch: async () => {
          throw Error('unexpected');
        },
      },
    });
  c.setFrame(f);
  for (let n = 0; n < 8 && c.getSnapshot().loading; n++) await new Promise((r) => setTimeout(r, 0));
  await assert.rejects(
    c.accept(f, {
      operation: 'updateChart',
      sheet: 'Data',
      objectId: id,
      expectedObjectRevision: 'r1',
      patch: {},
      anchor,
    }),
    /protected/
  );
  c.dispose();
});
test('worksheet pointer resize/move and keyboard deletion use shape commands without cell writes', async (t) => {
  const writes = [],
    root = createRoot(document.getElementById('root')),
    axes = createObjectAxes({ sheetFormat: { defaultRowHeight: 15 }, widths: [] }),
    args = {
      keyValue: 'owner/head',
      family: 'owner/sheet',
      gestureKey: 'view',
      sheet: 'Data',
      active: true,
      controls: {
        objects: async () => ({ objects: [shape], total: 1, nextCursor: null }),
        object: async () => shape,
        onBatch: async (c) => writes.push(c),
      },
      axes,
      view: {
        scale: 1.5,
        rowHeaderWidth: 48,
        columnHeaderHeight: 32,
        scrollLeft: 0,
        scrollTop: 0,
        width: 1200,
        height: 800,
      },
      onExtent: () => {},
    };
  await act(async () => root.render(React.createElement(WorksheetObjectLayer, args)));
  await flush();
  t.after(async () => act(async () => root.unmount()));
  let node = document.querySelector('[data-worksheet-object]');
  assert(node.querySelector('[data-shape-sample]'));
  const event = (x, y) => ({
    button: 0,
    pointerId: 1,
    clientX: x,
    clientY: y,
    preventDefault() {},
    stopPropagation() {},
    currentTarget: node,
  });
  await act(async () => props(node).onPointerDown(event(100, 100)));
  await act(async () => props(node).onPointerUp(event(130, 115)));
  await flush();
  assert.equal(writes[0][0].operation, 'updateShape');
  assert.equal(writes[0][0].anchor.dxEmu - anchor.dxEmu, 20 * E);
  node = document.querySelector('[data-worksheet-object]');
  await act(async () => props(node).onKeyDown({ key: 'Delete', preventDefault() {}, stopPropagation() {} }));
  await flush();
  assert.equal(writes[1][0].operation, 'deleteObject');
});
