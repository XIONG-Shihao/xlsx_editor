import './support/scratch-dir.mjs';
/** Controlled actual Home/Grid routes; temporary test bundle, no native/app/product build. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createHash } from 'node:crypto';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
const component = fileURLToPath(new URL('../', import.meta.url)),
  output = await fs.mkdtemp(path.join(component, '.build/tests/current-region-ui-'));
const files = [
  'ui/currentRegion.ts',
  'ui/specialSelection.ts',
  'ui/WorkbookSpecialSelection.tsx',
  'ui/SpreadsheetGrid.tsx',
  'ui/homeCommands.ts',
  'ui/ribbonCommands.ts',
  'tests/workbook-current-region-ui.mjs',
];
const pins = () =>
    Promise.all(
      files.map(async (relative) => {
        const p = path.join(component, relative),
          bytes = await fs.readFile(p);
        return { relative, path: p, sha256: createHash('sha256').update(bytes).digest('hex'), bytes: bytes.length };
      })
    ),
  before = await pins();
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
async function mount(t) {
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
    ownerIdentity: 'original-owner',
    acceptedCommandState: 'accepted-head',
    storageKey: 'native-head',
    active: true,
    readOnly: false,
    busy: false,
    readRange,
    sheets: [{ name: 'Sheet1', merges: [], widths: [], rows: 30, columns: 4, cells: structuredClone(data) }],
    onCommandBatch: async (commands) => accepted.push(structuredClone(commands)),
    onBatch: (commands) => legacy.push(commands),
    onCommandUndo: async () => history.push('undo'),
    onCommandRedo: async () => history.push('redo'),
    words: (text) => text,
  };
  const render = async (patch) => {
    input = { ...input, ...patch };
    await act(async () => root.render(React.createElement(ui.SpreadsheetGrid, input)));
    await flush();
  };
  await render({});
  await click([...document.querySelectorAll('[role=tab]')].find((node) => node.textContent === 'Home'));
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

async function region() {
  await click('home.special-open');
  await change(field('Selection type'), 'current-region');
  await click('Preview matching cells');
  await flush(40);
  assert(field('Special selection summary'));
  assert(field('Special selection summary').textContent.includes('Sheet1!A1:C3'));
}
check(
  'Current Region previews complete source beyond the viewport and selects a normal rectangle with original active cell',
  async (t) => {
    const h = await mount(t);
    await go('B2');
    const before = structuredClone(h.data);
    await region();
    await click('Select matching cells');
    for (let i = 0; i < 100 && document.querySelector('.spreadsheet-editor').dataset.homePending === 'true'; i++)
      await flush(5);
    assert.equal(field('Go to cell').getAttribute('placeholder'), 'A1:C3');
    assert.equal(document.querySelectorAll('.sheet-special-status').length, 0);
    assert(document.body.textContent.includes('Matching cells selected'));
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
    assert.deepEqual(h.history, []);
    assert.deepEqual(h.data, before);
    assert(h.reads.some((row) => row.range !== 'B2'));
    await click('Close workbook tools');
    assert.equal(field('Go to cell').getAttribute('placeholder'), 'A1:C3');
    assert.equal(field('Formula or value').value, '2');
  }
);
check('Cancel and a retained old owner handler do not select or write a replacement workbook', async (t) => {
  const h = await mount(t);
  await go('B2');
  await region();
  const old = props(button('Select matching cells')).onClick;
  await click('Cancel selection');
  assert.equal(field('Go to cell').getAttribute('placeholder'), 'B2');
  assert.deepEqual(h.accepted, []);
  await region();
  await h.render({ ownerIdentity: 'replacement-owner', documentId: 'replacement-doc' });
  await act(async () => old());
  await flush();
  assert.equal(field('Go to cell').getAttribute('placeholder'), 'B2');
  assert.deepEqual(h.accepted, []);
  assert.deepEqual(h.legacy, []);
});
