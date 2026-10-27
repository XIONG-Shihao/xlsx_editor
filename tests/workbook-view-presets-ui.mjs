import './support/scratch-dir.mjs';
/** Controlled actual Home/Grid routes; temporary test bundle, no native/app/product build. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
const component = fileURLToPath(new URL('../', import.meta.url)),
  output = await fs.mkdtemp(path.join(component, '.build/tests/view-presets-ui-'));
const entry = path.join(output, 'entry.ts'),
  bundle = path.join(output, 'grid.mjs');
await fs.writeFile(
  entry,
  `export {SpreadsheetGrid} from ${JSON.stringify(path.join(component, 'ui/SpreadsheetGrid.tsx'))};export {configureXlsxHost} from ${JSON.stringify(path.join(component, 'ui/host.tsx'))};`
);
await build({
  entryPoints: [entry],
  bundle: true,
  format: 'esm',
  platform: 'browser',
  jsx: 'automatic',
  outfile: bundle,
  external: ['react', 'react/jsx-runtime', 'react-dom', 'react-dom/client'],
  logLevel: 'silent',
});
const dom = new JSDOM('<div id="root"></div>', { url: 'http://localhost' });
Object.assign(globalThis, {
  window: dom.window,
  document: dom.window.document,
  HTMLElement: dom.window.HTMLElement,
  HTMLDetailsElement: dom.window.HTMLDetailsElement,
  Node: dom.window.Node,
  IS_REACT_ACT_ENVIRONMENT: true,
});
Object.defineProperties(window.HTMLElement.prototype, {
  clientWidth: {
    get() {
      return 1600;
    },
  },
  clientHeight: {
    get() {
      return 600;
    },
  },
});
globalThis.ResizeObserver = class {
  constructor(fn) {
    this.fn = fn;
  }
  observe() {
    this.fn([{ contentRect: { width: 1600, height: 600 } }]);
  }
  disconnect() {}
};
globalThis.requestAnimationFrame = (fn) => setTimeout(fn, 0);
globalThis.cancelAnimationFrame = clearTimeout;
const { createRoot } = await import('react-dom/client'),
  ui = await import(pathToFileURL(bundle));
ui.configureXlsxHost({
  api: { copyText: async () => {}, readClipboardText: async () => '' },
  Icon: () => React.createElement('svg'),
  useI18n: () => ({ locale: 'en' }),
});
const props = (node) => node[Object.keys(node).find((key) => key.startsWith('__reactProps'))],
  field = (name) =>
    [...document.querySelectorAll('[aria-label]')].find(
      (node) => node.getAttribute('aria-label') === name && !node.closest('[hidden]')
    ),
  node = (id) =>
    [...document.querySelectorAll(`[data-home-command="${id}"]`)].find((node) => !node.closest('[hidden]')),
  button = (name) =>
    [...document.querySelectorAll('button')].find(
      (node) => (node.getAttribute('aria-label') ?? node.textContent) === name && !node.closest('[hidden]')
    ),
  flush = async (ms = 6) => act(async () => new Promise((resolve) => setTimeout(resolve, ms)));
const defer = () => {
  let resolve, reject;
  const promise = new Promise((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
};
async function click(value) {
  const target = typeof value === 'string' ? (node(value) ?? button(value)) : value;
  assert(target);
  assert(!target.disabled);
  await act(async () => target.click());
  await flush();
}
async function change(target, value) {
  assert(target);
  await act(async () => props(target).onChange({ target: { value, checked: value } }));
  await flush();
}
async function go(range) {
  await change(field('Go to cell'), range);
  await act(async () =>
    field('Go to cell').dispatchEvent(
      new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })
    )
  );
  await flush();
}
const source = {
  A1: { value: 'Header', dataType: 'text' },
  B1: { value: 0, dataType: 'number' },
  C1: { value: false, dataType: 'boolean' },
  A2: { value: '#N/A', dataType: 'error', error: '#N/A' },
  B2: { value: 2, dataType: 'number' },
  C2: { value: '', dataType: 'text' },
  A3: { value: null, formula: 'SUM(B1:B2)', dataType: 'number' },
  C3: { value: 3, dataType: 'number' },
  E2: { value: 'separate', dataType: 'text' },
  A5: { value: 'separate', dataType: 'text' },
};
let ordinal = 0;
async function mount(t, patch = {}) {
  document.querySelector('#root').innerHTML = '';
  const root = createRoot(document.querySelector('#root')),
    accepted = [],
    legacy = [],
    reads = [],
    history = [],
    data = structuredClone(source);
  let input;
  const readRange = async (sheet, range) => {
    reads.push({ sheet, range });
    const cells = {},
      parts = range.split(':'),
      position = (ref) => {
        const m = /^([A-Z]+)(\d+)$/.exec(ref);
        return { row: Number(m[2]), col: [...m[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) };
      };
    const a = position(parts[0]),
      b = position(parts[1] ?? parts[0]);
    for (const [ref, value] of Object.entries(data)) {
      const p = position(ref);
      if (p.row >= a.row && p.row <= b.row && p.col >= a.col && p.col <= b.col) cells[ref] = structuredClone(value);
    }
    return cells;
  };
  input = {
    documentId: 'original-doc',
    viewPreferenceId: 'view-preset-test:' + ++ordinal,
    ownerIdentity: 'original-owner',
    acceptedCommandState: 'accepted-head',
    storageKey: 'native-head',
    active: true,
    readOnly: false,
    busy: false,
    readRange,
    sheets: [
      {
        name: 'Sheet1',
        merges: [],
        widths: [],
        rows: 30,
        columns: 4,
        cells: structuredClone(data),
        sourceView: { zoom: 175, gridlines: false, headings: false },
      },
    ],
    onCommandBatch: async (commands) => accepted.push(structuredClone(commands)),
    onBatch: (commands) => legacy.push(commands),
    onCommandUndo: async () => history.push('undo'),
    onCommandRedo: async () => history.push('redo'),
    words: (text) => text,
    ...patch,
  };
  const render = async (patch) => {
    input = { ...input, ...patch };
    await act(async () => root.render(React.createElement(ui.SpreadsheetGrid, input)));
    await flush();
  };
  await render({});
  await click([...document.querySelectorAll('[role=tab]')].find((node) => node.textContent === 'View'));
  t.after(async () => {
    await act(async () => root.unmount());
    await flush();
  });
  return { accepted, legacy, reads, history, data, render, input: () => input };
}
const check = (name, body) => {
  test(name, async (t) => {
    await body(t);
  });
};

const grid = () => document.querySelector('.sheet-grid');
async function zoom(value) {
  const input = document.querySelector('input[aria-label="Zoom"]');
  await change(input, String(value));
  await act(async () =>
    input.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
  );
  await flush();
}
function pure(h) {
  assert.deepEqual(h.accepted, []);
  assert.deepEqual(h.legacy, []);
  assert.deepEqual(h.history, []);
  assert.deepEqual(h.data, source);
}
check(
  'source defaults apply until individually overridden; reset and source replacement stay selection-only',
  async (t) => {
    const h = await mount(t);
    assert.equal(grid().dataset.zoom, '175');
    assert.equal(grid().dataset.gridlines, 'false');
    assert.equal(grid().dataset.headings, 'false');
    await go('B2');
    await zoom(200);
    assert.equal(grid().dataset.zoom, '200');
    assert.equal(grid().dataset.gridlines, 'false');
    const sheets = h
      .input()
      .sheets.map((sheet) => ({ ...sheet, sourceView: { zoom: 80, gridlines: true, headings: false } }));
    await h.render({ sheets });
    assert.equal(grid().dataset.zoom, '200');
    assert.equal(grid().dataset.gridlines, 'true');
    await click('Use workbook view');
    assert.equal(grid().dataset.zoom, '80');
    assert.equal(field('Go to cell').getAttribute('placeholder'), 'B2');
    pure(h);
  }
);
check(
  'three built-ins and named create/update/delete change only local display and retain current selection',
  async (t) => {
    const h = await mount(t);
    await go('B2');
    await change(field('Local view preset'), 'builtin:data-entry');
    assert.equal(grid().dataset.zoom, '125');
    assert.equal(grid().dataset.gridlines, 'true');
    await change(field('Local view name'), 'Entry');
    await click('Save local view');
    assert([...field('Local view preset').options].some((option) => option.value === 'named:Entry'));
    await zoom(200);
    await click('Update saved view');
    await change(field('Local view preset'), 'builtin:formula-audit');
    assert.equal(grid().dataset.showFormulas, 'true');
    await change(field('Local view preset'), 'named:Entry');
    assert.equal(grid().dataset.zoom, '200');
    assert.equal(grid().dataset.showFormulas, 'false');
    await click('Delete saved view');
    assert(![...field('Local view preset').options].some((option) => option.value === 'named:Entry'));
    await change(field('Local view preset'), 'builtin:default');
    assert.equal(grid().dataset.zoom, '100');
    assert.equal(field('Go to cell').getAttribute('placeholder'), 'B2');
    pure(h);
  }
);
check(
  'old preset handlers cannot save obsolete views or a replacement owner; readonly local controls stay usable',
  async (t) => {
    const h = await mount(t, { readOnly: true });
    await change(field('Local view name'), 'Original');
    const oldSave = props(button('Save local view')).onClick;
    await change(field('Local view preset'), 'builtin:data-entry');
    await act(async () => oldSave());
    await flush();
    assert(![...field('Local view preset').options].some((option) => option.value === 'named:Original'));
    const oldReset = props(button('Use workbook view')).onClick;
    await h.render({
      documentId: 'replacement-doc',
      ownerIdentity: 'replacement-owner',
      viewPreferenceId: 'replacement-file',
    });
    await act(async () => oldReset());
    await flush();
    assert.equal(grid().dataset.zoom, '175');
    await change(field('Local view preset'), 'builtin:formula-audit');
    assert.equal(grid().dataset.showFormulas, 'true');
    pure(h);
  }
);
check(
  'first Save moves named views to the stable file binding without changing source ownership/history',
  async (t) => {
    const h = await mount(t, { viewPreferenceId: 'new:presets' });
    await change(field('Local view preset'), 'builtin:data-entry');
    await change(field('Local view name'), 'First save');
    await click('Save local view');
    await h.render({ viewPreferenceId: 'file:presets' });
    assert.equal(grid().dataset.zoom, '125');
    assert([...field('Local view preset').options].some((option) => option.value === 'named:First save'));
    await h.render({ documentId: 'foreign-doc', ownerIdentity: 'foreign-owner', viewPreferenceId: 'foreign-file' });
    assert(![...field('Local view preset').options].some((option) => option.value === 'named:First save'));
    pure(h);
  }
);
