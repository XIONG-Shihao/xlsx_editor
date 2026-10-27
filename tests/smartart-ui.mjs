import './support/scratch-dir.mjs';
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
import { smartartSampleModel } from '../core/smartart-model.mjs';
import { createWorksheetObjectController } from '../ui/worksheetObjectController.ts';
import { objectPlacementCommand } from '../ui/worksheetObjectGeometry.ts';
const out = await fs.mkdtemp(path.resolve('.build/tests/smartart-ui-')),
  bundle = path.join(out, 'ui.mjs');
await build({
  stdin: {
    contents:
      'export {WorkbookSmartArt} from "./ui/WorkbookSmartArt"; export {WorkbookIllustrations} from "./ui/WorkbookIllustrations"; export {WorkbookSmartArtPreview} from "./ui/WorkbookSmartArtPreview"; export {WorksheetObjectLayer} from "./ui/WorksheetObjectLayer"; export {createObjectAxes} from "./ui/worksheetObjectGeometry";',
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
const { WorkbookSmartArt, WorkbookIllustrations, WorkbookSmartArtPreview, WorksheetObjectLayer, createObjectAxes } =
    await import(pathToFileURL(bundle)),
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
  anchor = { kind: 'oneCell', cell: 'B2', dxEmu: 3 * E, dyEmu: 4 * E, widthEmu: 600 * E, heightEmu: 360 * E },
  diagram = {
    id,
    name: 'Plan',
    kind: 'smartart',
    revision: 'r1',
    anchor,
    managed: true,
    canEdit: true,
    canDelete: true,
    protectedReason: null,
    altText: 'Original diagram',
    smartart: smartartSampleModel('process'),
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
    assert(n, label);
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
  key: 'head1/selection1/smartart',
  family: 'owner/sheet',
  sheet: 'Data',
  active: true,
  controls: {
    objects: async () => ({ objects: [], total: 0, nextCursor: null }),
    object: async () => diagram,
    onBatch: async () => {},
  },
  ...patch,
});
async function mount(t, patch = {}, Component = WorkbookSmartArt) {
  const root = createRoot(document.getElementById('root')),
    writes = [];
  let args = {
    frame: frame({ controls: { ...frame().controls, onBatch: async (c) => writes.push(c) } }),
    mode: 'smartart',
    selection: 'B2',
    ...patch,
  };
  const render = async (p) => {
    args = { ...args, ...p };
    await act(async () => root.render(React.createElement(Component, args)));
    await flush();
  };
  await render({});
  t.after(async () => act(async () => root.unmount()));
  return {
    writes,
    render,
    get args() {
      return args;
    },
  };
}
after(async () => {
  dom.window.close();
});
test('three layout buttons display actual semantic nodes and preserve edited labels through layout browsing', async (t) => {
  await mount(t);
  await click('Hierarchy');
  assert.equal(document.querySelector('[aria-label="SmartArt node 1"]').value, 'Lead');
  await input('SmartArt node 1', 'Edited <&é');
  await click('Cycle');
  assert.equal(document.querySelector('[aria-label="SmartArt node 1"]').value, 'Edited <&é');
  assert.equal(document.querySelectorAll('[aria-label="SmartArt layouts"] [data-smartart-node]').length, 9);
  assert.equal(document.querySelector('img,image'), null);
});
test('each gallery creates a captured literal node/layout/anchor/alt model with the native command', async (t) => {
  const ui = await mount(t);
  for (const [layout, label] of [
    ['process', 'Process'],
    ['hierarchy', 'Hierarchy'],
    ['cycle', 'Cycle'],
  ]) {
    await ui.render({ frame: { ...ui.args.frame, key: layout } });
    await click(label);
    await input('SmartArt node 1', 'Literal "&é');
    await input('Object alternative text', 'Accessible ' + layout);
    await submit();
    const c = ui.writes.at(-1)[0];
    assert.equal(c.operation, 'createSmartArt');
    assert.equal(c.expectedObjectRevision, null);
    assert.equal(c.smartart.layout, layout);
    assert.equal(c.smartart.nodes[0].text, 'Literal "&é');
    assert.equal(c.anchor.cell, 'B2');
    assert.equal(c.anchor.widthEmu, 600 * E);
    assert.equal(c.anchor.heightEmu, 360 * E);
    assert.equal(c.altText, 'Accessible ' + layout);
  }
  assert.equal(ui.writes.length, 3);
});
test('editing nodes/count/layout/alt preserves source identity and offsets; Delete uses the scoped owner command', async (t) => {
  const writes = [],
    f = frame({
      controls: {
        objects: async () => ({ objects: [diagram], total: 1, nextCursor: null }),
        object: async () => diagram,
        onBatch: async (c) => writes.push(c),
      },
    });
  await mount(t, { frame: f });
  await click('Process — B2');
  await input('SmartArt node count', '4');
  await input('SmartArt node 4', 'Added node');
  await click('Cycle');
  await input('Object alternative text', 'Edited diagram');
  await submit();
  const c = writes[0][0];
  assert.equal(c.operation, 'updateSmartArt');
  assert.equal(c.expectedObjectRevision, 'r1');
  assert.equal(c.patch.nodes.length, 4);
  assert.equal(c.patch.nodes[3].text, 'Added node');
  assert.equal(c.patch.layout, 'cycle');
  assert.equal(c.patch.connections.length, 4);
  assert.equal(c.anchor.dxEmu, 3 * E);
  assert.equal(c.anchor.dyEmu, 4 * E);
  await click('Process — B2');
  await click('Delete object');
  assert.equal(writes[1][0].operation, 'deleteSmartArt');
});
test('invalid labels/extent and retired or readonly drafts cannot write or crash the preview', async (t) => {
  const ui = await mount(t);
  await input('SmartArt node 1', ' ');
  await submit();
  assert.equal(ui.writes.length, 0);
  assert.match(document.body.textContent, /nonblank/);
  await input('SmartArt node 1', 'x'.repeat(129));
  await submit();
  assert.equal(ui.writes.length, 0);
  await input('SmartArt node 1', 'Valid');
  await input('Object width', 'NaN');
  await submit();
  assert.equal(ui.writes.length, 0);
  assert.doesNotMatch(document.querySelector('svg').outerHTML, /(NaN|Infinity)/);
  await click('Reset form');
  const old = props(document.querySelector('form')).onSubmit;
  await input('SmartArt node 1', 'Discard draft');
  await ui.render({ frame: { ...ui.args.frame, key: 'other' }, selection: 'C4' });
  assert.equal(document.querySelector('[aria-label="SmartArt node 1"]').value, 'Plan');
  await act(async () => old({ preventDefault() {} }));
  assert.equal(ui.writes.length, 0);
  await ui.render({ frame: { ...ui.args.frame, key: 'readonly', reason: 'This workbook is read-only.' } });
  await submit();
  assert.equal(ui.writes.length, 0);
  assert(document.querySelector('fieldset').disabled);
});
test('shared illustration controller retains one physical write lane across a SmartArt to Shapes switch', async (t) => {
  const gate = deferred(),
    writes = [],
    f = frame({
      controls: {
        ...frame().controls,
        onBatch: async (c) => {
          writes.push(c);
          await gate.promise;
        },
      },
    }),
    ui = await mount(t, { frame: f }, WorkbookIllustrations);
  await act(async () => props(document.querySelector('form')).onSubmit({ preventDefault() {} }));
  await flush();
  assert.equal(writes.length, 1);
  assert(document.querySelector('fieldset').disabled);
  await ui.render({ mode: 'shape', frame: { ...f, key: 'head1/selection1/shapes' } });
  assert(document.querySelector('fieldset').disabled);
  await act(async () => props(document.querySelector('form')).onSubmit({ preventDefault() {} }));
  assert.equal(writes.length, 1);
  await act(async () => gate.resolve());
  await flush();
  assert(!document.querySelector('fieldset').disabled);
  assert.equal(writes[0][0].operation, 'createSmartArt');
});
test('controller and placement reject wrong-kind or protected edits while cloning accepted SmartArt commands', async () => {
  const controller = createWorksheetObjectController(),
    gate = deferred(),
    writes = [],
    f = frame({
      controls: {
        objects: async () => ({ objects: [diagram], total: 1, nextCursor: null }),
        object: async () => diagram,
        onBatch: async (c) => {
          writes.push(c);
          await gate.promise;
        },
      },
    });
  controller.setFrame(f);
  await flush();
  const c = objectPlacementCommand('Data', diagram, { ...anchor });
  assert.equal(c.operation, 'updateSmartArt');
  assert.throws(() => objectPlacementCommand('Data', { ...diagram, canEdit: false }, anchor));
  await assert.rejects(controller.accept(f, { ...c, operation: 'updateChart' }));
  const run = controller.accept(f, c);
  c.anchor.cell = 'A100';
  assert.equal(writes[0][0].anchor.cell, 'B2');
  gate.resolve();
  await run;
  controller.dispose();
});
test('worksheet body/resize/keyboard/delete route SmartArt through captured object commands without cell writes', async (t) => {
  const axes = createObjectAxes({
      sheetFormat: { defaultRowHeight: 15, rowsHiddenByDefault: false },
      widths: [{ min: 1, max: 8, width: 8.43, hidden: false }],
    }),
    writes = [],
    root = createRoot(document.getElementById('root'));
  await act(async () =>
    root.render(
      React.createElement(WorksheetObjectLayer, {
        keyValue: 'head1',
        family: 'owner/sheet',
        gestureKey: 'view1',
        sheet: 'Data',
        active: true,
        controls: {
          objects: async () => ({ objects: [diagram], total: 1, nextCursor: null }),
          object: async () => diagram,
          onBatch: async (c) => writes.push(c),
        },
        axes,
        view: {
          scale: 1,
          rowHeaderWidth: 48,
          columnHeaderHeight: 32,
          scrollLeft: 0,
          scrollTop: 0,
          width: 1000,
          height: 600,
        },
        onExtent: () => {},
      })
    )
  );
  await flush();
  t.after(async () => act(async () => root.unmount()));
  const node = document.querySelector('[data-worksheet-object]');
  assert(node);
  const pointer = async (type, x, y, target = node) => {
    await act(async () => {
      const e = new window.MouseEvent(type, { bubbles: true, cancelable: true, button: 0, clientX: x, clientY: y });
      Object.defineProperty(e, 'pointerId', { value: 1 });
      target.dispatchEvent(e);
    });
    await flush();
  };
  await pointer('pointerdown', 150, 80);
  await pointer('pointerup', 170, 90);
  assert.equal(writes[0][0].operation, 'updateSmartArt');
  const resize = document.querySelector('[aria-label^="Resize object:"]');
  await pointer('pointerdown', 700, 400, resize);
  await pointer('pointerup', 730, 420, resize);
  assert.equal(writes[1][0].operation, 'updateSmartArt');
  assert.equal(writes[1][0].anchor.widthEmu, 630 * E);
  await act(async () =>
    document
      .querySelector('[data-worksheet-object]')
      .dispatchEvent(new window.KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }))
  );
  await flush();
  assert.equal(writes[2][0].operation, 'updateSmartArt');
  await act(async () =>
    document
      .querySelector('[data-worksheet-object]')
      .dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Delete', bubbles: true }))
  );
  await flush();
  assert.equal(writes[3][0].operation, 'deleteObject');
  assert(writes.every((c) => c.length === 1 && c[0].expectedObjectRevision === 'r1'));
});
