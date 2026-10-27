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
  output = await fs.mkdtemp(path.join(component, '.build/tests/clear-hyperlinks-ui-'));
const files = [
  'core/clear-hyperlinks.mjs',
  'core/contracts.ts',
  'ui/clearHyperlinksTargets.ts',
  'ui/SpreadsheetGrid.tsx',
  'ui/homeCommands.ts',
  'ui/ribbonCommands.ts',
  'tests/workbook-clear-hyperlinks-ui.mjs',
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
  A1: {
    value: 'linked text',
    dataType: 'text',
    styleId: 7,
    underline: true,
    fontColor: 'FF0000',
    fillColor: 'FFFF00',
    format: '0.0000',
  },
  B1: { value: 2, dataType: 'number' },
  C1: { value: '#N/A', dataType: 'text' },
  A2: { value: true, dataType: 'boolean' },
  B2: { value: '#N/A', dataType: 'error', error: '#N/A' },
  A3: { value: '', dataType: 'text' },
  C3: { value: 9007199254740992, rawValue: '9007199254740993', dataType: 'number' },
  B4: {
    value: -7,
    formula: 'C4+1',
    cachedValue: -7,
    calculation: { status: 'stored', pending: true },
    cachedResult: { authority: 'stored', status: 'calculated', value: -7, rawValue: '-7' },
  },
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
check('Clear menu sends one full-sheet geometry command without reads or cell/style projection', async (t) => {
  const h = await mount(t);
  assert.equal(node('home.clear-hyperlinks').textContent, 'Clear Hyperlinks');
  await go('A1:XFD1048576');
  const count = h.reads.length,
    before = structuredClone(h.data);
  await click('home.clear-hyperlinks');
  assert.deepEqual(h.accepted, [[{ type: 'clearHyperlinks', sheet: 'Sheet1', ranges: ['A1:XFD1048576'] }]]);
  assert.equal(h.reads.length, count);
  assert.deepEqual(h.data, before);
  assert.deepEqual(h.legacy, []);
  assert.deepEqual(h.history, []);
  assert.equal(document.querySelector('.spreadsheet-editor').dataset.homePending, 'false');
});
check('Clear Hyperlinks consumes exact sparse Special targets with holes in one command', async (t) => {
  const h = await mount(t);
  await go('A1:C3');
  await click('home.special-open');
  await click('Preview matching cells');
  await flush(20);
  assert(field('Special selection summary'));
  await click('Select matching cells');
  await click('Close workbook tools');
  assert(document.querySelector('.sheet-special-status'));
  const count = h.reads.length,
    before = structuredClone(h.data);
  await click('home.clear-hyperlinks');
  assert.deepEqual(h.accepted, [
    [{ type: 'clearHyperlinks', sheet: 'Sheet1', ranges: ['A1:C1', 'A2:B2', 'A3', 'C3'] }],
  ]);
  assert.equal(h.reads.length, count);
  assert.deepEqual(h.data, before);
  assert.deepEqual(h.legacy, []);
  assert(document.querySelector('.sheet-special-status'));
});
check(
  'Readonly/passive/CSV/portable/missing callback/inactive/busy barriers refuse direct retained handlers',
  async (t) => {
    const h = await mount(t);
    for (const patch of [
      { readOnly: true },
      { readOnly: false, literal: true },
      { literal: false, passiveNotice: 'Passive source' },
      { passiveNotice: undefined, storageKey: undefined },
      { storageKey: 'native-head', ownerIdentity: undefined },
      { ownerIdentity: 'original-owner', onCommandBatch: undefined },
      { onCommandBatch: async (commands) => h.accepted.push(commands), active: false },
      { active: true, busy: true },
    ]) {
      await h.render(patch);
      const control = node('home.clear-hyperlinks');
      assert(control.disabled);
      assert(control.title);
      await act(async () => props(control).onClick());
      await flush();
      assert.deepEqual(h.accepted, []);
      assert.deepEqual(h.legacy, []);
    }
  }
);
check('Original selection/head/owner handlers retire before accepting a link command', async (t) => {
  const h = await mount(t);
  await go('B2:C3');
  const oldSelection = props(node('home.clear-hyperlinks')).onClick;
  await go('A1');
  await act(async () => oldSelection());
  await flush();
  assert.deepEqual(h.accepted, []);
  const oldHead = props(node('home.clear-hyperlinks')).onClick;
  await h.render({ acceptedCommandState: 'new-head', storageKey: 'new-native-head' });
  await act(async () => oldHead());
  await flush();
  assert.deepEqual(h.accepted, []);
  const oldOwner = props(node('home.clear-hyperlinks')).onClick;
  await h.render({ ownerIdentity: 'other-owner', documentId: 'other-doc' });
  await act(async () => oldOwner());
  await flush();
  assert.deepEqual(h.accepted, []);
  assert.deepEqual(h.legacy, []);
});
check('One captured acceptance awaits acknowledgement and suppresses retired owner errors', async (t) => {
  const h = await mount(t),
    held = defer();
  await go('B2:C3');
  await h.render({
    onCommandBatch: (commands) => {
      h.accepted.push(structuredClone(commands));
      return held.promise;
    },
  });
  await click('home.clear-hyperlinks');
  assert.deepEqual(h.accepted, [[{ type: 'clearHyperlinks', sheet: 'Sheet1', ranges: ['B2:C3'] }]]);
  assert.equal(document.querySelector('.spreadsheet-editor').dataset.homePending, 'true');
  assert(node('home.clear-contents').disabled);
  await h.render({ ownerIdentity: 'new-owner', documentId: 'new-doc' });
  await act(async () => held.resolve());
  await flush();
  assert.equal(h.accepted.length, 1);
  assert.equal(document.querySelector('.spreadsheet-editor').dataset.homePending, 'false');
  assert.equal(document.querySelector('.sheet-grid-error'), null);
  const refused = defer();
  await h.render({ onCommandBatch: () => refused.promise });
  await click('home.clear-hyperlinks');
  await h.render({ ownerIdentity: 'replacement-owner' });
  await act(async () => refused.reject(Error('Original hyperlink graph is incomplete')));
  await flush();
  assert(!document.body.textContent.includes('Original hyperlink graph is incomplete'));
  assert.equal(h.accepted.length, 1);
  assert.deepEqual(h.legacy, []);
});
