import './support/scratch-dir.mjs';
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
import { model3DSampleAsset, normalizeModel3D } from '../core/model3d-model.mjs';
import { createWorksheetObjectController } from '../ui/worksheetObjectController.ts';
import { objectPlacementCommand } from '../ui/worksheetObjectGeometry.ts';
const out = await fs.mkdtemp(path.resolve('.build/tests/model3d-ui-')),
  bundle = path.join(out, 'ui.mjs');
await build({
  stdin: {
    contents:
      'export {WorkbookModel3D} from "./ui/WorkbookModel3D"; export {WorkbookIllustrations} from "./ui/WorkbookIllustrations"; export {WorkbookModel3DPreview} from "./ui/WorkbookModel3DPreview"; export {WorksheetObjectLayer} from "./ui/WorksheetObjectLayer"; export {createObjectAxes} from "./ui/worksheetObjectGeometry";',
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
const { WorkbookModel3D, WorkbookIllustrations, WorkbookModel3DPreview, WorksheetObjectLayer, createObjectAxes } =
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
    name: 'Cube',
    kind: 'model3d',
    revision: 'r1',
    anchor,
    managed: true,
    canEdit: true,
    canDelete: true,
    protectedReason: null,
    altText: 'Original model',
    model3d: model3DSampleAsset('cube').model,
    asset3d: model3DSampleAsset('cube'),
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
  key: 'head1/selection1/model3d',
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
async function mount(t, patch = {}, Component = WorkbookModel3D) {
  const root = createRoot(document.getElementById('root')),
    writes = [];
  let args = {
    frame: frame({ controls: { ...frame().controls, onBatch: async (c) => writes.push(c) } }),
    mode: 'model3d',
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
test('three real solid samples render decoded faces and sample browsing keeps color/rotation drafts', async (t) => {
  await mount(t);
  await input('3D model color', '#cc3333');
  await input('3D rotation Y', '75');
  await click('Pyramid');
  assert.equal(document.querySelector('[aria-label="3D model color"]').value, '#cc3333');
  assert.equal(document.querySelector('[aria-label="3D rotation Y"]').value, '75');
  await click('Octahedron');
  assert.equal(document.querySelector('[aria-label="3D rotation Y"]').value, '75');
  assert.equal(document.querySelectorAll('[aria-label="3D model samples"] [data-model3d-sample]').length, 3);
  assert(document.querySelectorAll('[data-model3d-face]').length > 6);
  assert.equal(document.querySelector('img,image'), null);
});
test('each gallery creates a captured color/rotation/placement/alt command', async (t) => {
  const ui = await mount(t);
  for (const [sample, label] of [
    ['cube', 'Cube'],
    ['pyramid', 'Pyramid'],
    ['octahedron', 'Octahedron'],
  ]) {
    await ui.render({ frame: { ...ui.args.frame, key: sample } });
    await click(label);
    await input('3D rotation X', '-35');
    await input('3D model color', '#3366cc');
    await input('Object alternative text', 'Accessible ' + sample);
    await submit();
    const c = ui.writes.at(-1)[0];
    assert.equal(c.operation, 'createModel3D');
    assert.equal(c.expectedObjectRevision, null);
    assert.deepEqual(c.model3d, normalizeModel3D({ sample, color: '3366CC', rotation: { x: -35, y: -30, z: 0 } }));
    assert.equal(c.anchor.cell, 'B2');
    assert.equal(c.anchor.widthEmu, 320 * E);
    assert.equal(c.anchor.heightEmu, 240 * E);
    assert.equal(c.altText, 'Accessible ' + sample);
  }
  assert.equal(ui.writes.length, 3);
});
test('saved-model editing keeps identity/offsets and changes sample, color, rotation and alt; scoped Delete', async (t) => {
  const writes = [],
    f = frame({
      controls: {
        objects: async () => ({ objects: [diagram], total: 1, nextCursor: null }),
        object: async () => diagram,
        onBatch: async (c) => writes.push(c),
      },
    });
  await mount(t, { frame: f });
  await click('Cube — B2');
  await click('Pyramid');
  await input('3D rotation Z', '-180');
  await input('3D model color', '#000000');
  await input('Object alternative text', 'Edited model');
  await submit();
  const c = writes[0][0];
  assert.equal(c.operation, 'updateModel3D');
  assert.equal(c.expectedObjectRevision, 'r1');
  assert.equal(c.patch.sample, 'pyramid');
  assert.equal(c.patch.rotation.z, -180);
  assert.equal(c.patch.color, '000000');
  assert.equal(c.anchor.dxEmu, 3 * E);
  assert.equal(c.anchor.dyEmu, 4 * E);
  await click('Cube — B2');
  await click('Delete object');
  assert.equal(writes[1][0].operation, 'deleteModel3D');
});
test('invalid fractional/out-of-range rotations and dimensions refuse without crashing; retired/readonly drafts cannot write', async (t) => {
  const ui = await mount(t);
  for (const value of ['181', '0.5', 'NaN']) {
    await input('3D rotation X', value);
    await submit();
    assert.equal(ui.writes.length, 0);
  }
  await click('Reset form');
  await input('Object width', '99999');
  await submit();
  assert.equal(ui.writes.length, 0);
  assert.doesNotMatch(document.querySelector('svg').outerHTML, /(NaN|Infinity)/);
  await click('Reset form');
  const old = props(document.querySelector('form')).onSubmit;
  await input('3D model color', '#ffffff');
  await ui.render({ frame: { ...ui.args.frame, key: 'other' }, selection: 'C4' });
  assert.equal(document.querySelector('[aria-label="3D model color"]').value, '#217346');
  await act(async () => old({ preventDefault() {} }));
  assert.equal(ui.writes.length, 0);
  await ui.render({ frame: { ...ui.args.frame, key: 'readonly', reason: 'This workbook is read-only.' } });
  await submit();
  assert.equal(ui.writes.length, 0);
  assert(document.querySelector('fieldset').disabled);
});
test('one physical write lane survives a 3D-to-Shapes gallery switch and drops the stale completion', async (t) => {
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
  await ui.render({ mode: 'shape', frame: { ...f, key: 'head1/selection1/shapes' } });
  assert(document.querySelector('fieldset').disabled);
  await act(async () => props(document.querySelector('form')).onSubmit({ preventDefault() {} }));
  assert.equal(writes.length, 1);
  await act(async () => gate.resolve());
  await flush();
  assert(!document.querySelector('fieldset').disabled);
  assert.equal(writes[0][0].operation, 'createModel3D');
});
test('controller/placement require decoded actual model and capture accepted nested edits', async () => {
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
  assert.equal(c.operation, 'updateModel3D');
  assert.throws(() => objectPlacementCommand('Data', { ...diagram, asset3d: undefined }, anchor));
  await assert.rejects(controller.accept(f, { ...c, operation: 'updateSmartArt' }));
  const command = { ...c, patch: { rotation: { x: 10, y: 20, z: 30 } } },
    run = controller.accept(f, command);
  command.patch.rotation.y = 99;
  command.anchor.cell = 'A100';
  assert.equal(writes[0][0].patch.rotation.y, 20);
  assert.equal(writes[0][0].anchor.cell, 'B2');
  gate.resolve();
  await run;
  controller.dispose();
});
test('worksheet body/resize/keyboard/delete route 3D through captured object commands without cell writes', async (t) => {
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
  assert.equal(writes[0][0].operation, 'updateModel3D');
  const resize = document.querySelector('[aria-label^="Resize object:"]');
  await pointer('pointerdown', 700, 400, resize);
  await pointer('pointerup', 730, 420, resize);
  assert.equal(writes[1][0].operation, 'updateModel3D');
  assert.equal(writes[1][0].anchor.widthEmu, 630 * E);
  await act(async () =>
    document
      .querySelector('[data-worksheet-object]')
      .dispatchEvent(new window.KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }))
  );
  await flush();
  assert.equal(writes[2][0].operation, 'updateModel3D');
  await act(async () =>
    document
      .querySelector('[data-worksheet-object]')
      .dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Delete', bubbles: true }))
  );
  await flush();
  assert.equal(writes[3][0].operation, 'deleteObject');
  assert(writes.every((c) => c.length === 1 && c[0].expectedObjectRevision === 'r1'));
});

test('a retired inventory reply and preview error do not grant 3D editing authority', async () => {
  const gate = deferred(),
    controller = createWorksheetObjectController(),
    old = frame({ controls: { ...frame().controls, objects: () => gate.promise } });
  controller.setFrame(old);
  const next = frame({
    key: 'other',
    controls: {
      ...frame().controls,
      objects: async () => ({ objects: [diagram], total: 1, nextCursor: null }),
      object: async () => {
        throw Error('Corrupt GLB');
      },
    },
  });
  controller.setFrame(next);
  gate.resolve({ objects: [diagram], total: 1, nextCursor: null });
  await flush();
  assert.equal(controller.getSnapshot().key, 'other');
  assert.equal(controller.getSnapshot().objects.length, 0);
  assert.match(controller.getSnapshot().error, /Corrupt/);
  await assert.rejects(
    controller.accept(next, {
      operation: 'updateModel3D',
      sheet: 'Data',
      objectId: id,
      expectedObjectRevision: 'r1',
      patch: {},
    })
  );
  controller.dispose();
});
