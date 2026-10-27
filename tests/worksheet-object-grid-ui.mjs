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
  out = await fs.mkdtemp(path.join(component, '.build/tests/worksheet-object-grid-ui-')),
  bundle = path.join(out, 'ui.mjs');
await build({
  stdin: {
    contents: 'export {SpreadsheetGrid} from "./ui/SpreadsheetGrid"; export {configureXlsxHost} from "./ui/host";',
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
  { SpreadsheetGrid, configureXlsxHost } = await import(pathToFileURL(bundle));
after(async () => {
  dom.window.close();
});
const E = 9525,
  object = {
    id: '11111111-1111-4111-8111-111111111111',
    name: 'Chart',
    kind: 'chart',
    revision: 'r1',
    anchor: { kind: 'oneCell', cell: 'E2', dxEmu: 0, dyEmu: 0, widthEmu: 100 * E, heightEmu: 50 * E },
    managed: true,
    canEdit: true,
    canDelete: true,
    protectedReason: null,
    chart: { kind: 'column', range: 'A1:B3' },
    data: { categories: ['A'], series: [{ name: 'Amount', values: [2] }], hidden: [false] },
  };
const sheet = (name) => ({
  name,
  path: 'xl/worksheets/' + name + '.xml',
  state: 'visible',
  rows: 30,
  columns: 8,
  cells: { A1: { value: 'Category' }, B1: { value: 'Amount' }, A2: { value: 'A' }, B2: { value: 2 } },
  merges: [],
  widths: [],
  hiddenRows: [],
  rowHeights: [],
  freeze: { rows: 0, columns: 0 },
});
Object.defineProperties(window.HTMLElement.prototype, {
  clientWidth: {
    get() {
      return 800;
    },
  },
  clientHeight: {
    get() {
      return 500;
    },
  },
});
globalThis.ResizeObserver = class {
  constructor(f) {
    this.f = f;
  }
  observe() {
    this.f([{ contentRect: { width: 800, height: 500 } }]);
  }
  disconnect() {}
};
globalThis.requestAnimationFrame = () => 1;
globalThis.cancelAnimationFrame = () => {};
configureXlsxHost({
  api: { copyText: async () => {}, readClipboardText: async () => '' },
  Icon: () => React.createElement('svg'),
  useI18n: () => ({ locale: 'en' }),
});
const flush = async () => {
    for (let n = 0; n < 3; n++) await act(async () => new Promise((r) => setTimeout(r, 0)));
  },
  node = () => document.querySelector('[data-worksheet-object]');
async function pointer(type, x, y) {
  await act(async () => {
    const e = new window.MouseEvent(type, { bubbles: true, cancelable: true, button: 0, clientX: x, clientY: y });
    Object.defineProperty(e, 'pointerId', { value: 1 });
    node().dispatchEvent(e);
  });
  await flush();
}
async function mount(t, patch = {}) {
  const root = createRoot(document.getElementById('root')),
    strict = [],
    legacy = [];
  let args = {
    sheets: [sheet('Data'), sheet('Other')],
    active: true,
    documentId: 'doc',
    ownerIdentity: 'owner',
    storageKey: 'head',
    onEdit: () => {},
    readRange: async () => ({}),
    onCommandBatch: async (c) => strict.push(c),
    objectControls: {
      documentId: 'doc',
      revision: 'head',
      objects: async () => ({ objects: [object], total: 1, nextCursor: null }),
      object: async () => object,
      onBatch: async (c) => legacy.push(c),
    },
    ...patch,
  };
  const render = async (p) => {
    args = { ...args, ...p };
    await act(async () => root.render(React.createElement(SpreadsheetGrid, args)));
    await flush();
  };
  await render({});
  t.after(async () => act(async () => root.unmount()));
  return { strict, legacy, render };
}
test('actual Grid routes a completed drawing gesture through strict captured commands rather than legacy object batches', async (t) => {
  const ui = await mount(t);
  assert(node());
  await pointer('pointerdown', 300, 70);
  await pointer('pointerup', 320, 80);
  assert.equal(ui.strict.length, 1);
  assert.equal(ui.strict[0][0].operation, 'updateChart');
  assert.equal(ui.strict[0][0].expectedObjectRevision, 'r1');
  assert.equal(ui.strict[0][0].sheet, 'Data');
  assert.deepEqual(ui.legacy, []);
});
test('drawing inspection is available but writes refuse in readonly and non-strict hosts', async (t) => {
  const ui = await mount(t, { readOnly: true });
  assert(node());
  await pointer('pointerdown', 300, 70);
  await pointer('pointerup', 320, 80);
  assert.deepEqual(ui.strict, []);
  await ui.render({ readOnly: false, onCommandBatch: undefined });
  await pointer('pointerdown', 300, 70);
  await pointer('pointerup', 320, 80);
  assert.deepEqual(ui.legacy, []);
});
test('source sheets keep cell values intact while drawing extents extend the local scroll surface', async (t) => {
  const far = { ...object, anchor: { ...object.anchor, cell: 'E200' } },
    ui = await mount(t, {
      objectControls: {
        documentId: 'doc',
        revision: 'head',
        objects: async () => ({ objects: [far], total: 1, nextCursor: null }),
        object: async () => far,
        onBatch: async () => {
          throw Error('Legacy forbidden');
        },
      },
    });
  assert.match(document.querySelector('.spreadsheet-editor footer').textContent, /of 20[0-9]/);
  assert.deepEqual(ui.strict, []);
  assert.deepEqual(ui.legacy, []);
});
