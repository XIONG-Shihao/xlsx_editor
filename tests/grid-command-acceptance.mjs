import './support/scratch-dir.mjs';
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
const component = fileURLToPath(new URL('../', import.meta.url)),
  prefix = path.join(component, '.build/tests/grid-command-acceptance');
await fs.mkdir(path.dirname(prefix), { recursive: true });
const directory = await fs.mkdtemp(prefix + '-'),
  entry = path.join(directory, 'entry.ts'),
  module = path.join(directory, 'grid.mjs');
await fs.writeFile(
  entry,
  `export {SpreadsheetGrid,acceptGridCommandBatch,normalizeGridCommandBatch} from ${JSON.stringify(path.join(component, 'ui/SpreadsheetGrid.tsx'))};export {configureXlsxHost} from ${JSON.stringify(path.join(component, 'ui/host.tsx'))};export {dispatchRibbonCommand} from ${JSON.stringify(path.join(component, 'ui/ribbonCommands.ts'))};`
);
await build({
  entryPoints: [entry],
  bundle: true,
  format: 'esm',
  platform: 'browser',
  jsx: 'automatic',
  outfile: module,
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
      return 500;
    },
  },
  clientHeight: {
    get() {
      return 300;
    },
  },
});
globalThis.ResizeObserver = class {
  constructor(fn) {
    this.fn = fn;
  }
  observe() {
    this.fn([{ contentRect: { width: 500, height: 300 } }]);
  }
  disconnect() {}
};
let frames = [];
globalThis.requestAnimationFrame = (fn) => {
  frames.push(fn);
  return frames.length;
};
globalThis.cancelAnimationFrame = (id) => {
  frames[id - 1] = null;
};
const { SpreadsheetGrid, acceptGridCommandBatch, normalizeGridCommandBatch, configureXlsxHost, dispatchRibbonCommand } =
  await import(pathToFileURL(module));
configureXlsxHost({
  api: { copyText: async () => {}, readClipboardText: async () => '' },
  Icon: () => React.createElement('svg'),
  useI18n: () => ({ locale: 'en' }),
});
const { createRoot } = await import('react-dom/client');
after(() => {
  dom.window.close();
  for (const key of [
    'window',
    'document',
    'HTMLElement',
    'HTMLDetailsElement',
    'Node',
    'ResizeObserver',
    'requestAnimationFrame',
    'cancelAnimationFrame',
    'IS_REACT_ACT_ENVIRONMENT',
  ])
    delete globalThis[key];
});
const deferred = () => {
  let resolve, reject;
  const promise = new Promise((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
};
const base = () => ({ active: true, readOnly: false, busy: false, specialSelection: false });
const sheet = () => ({
  name: 'Sheet1',
  state: 'visible',
  path: 'xl/worksheets/sheet1.xml',
  rows: 30,
  columns: 12,
  cells: { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 } },
  merges: [],
  widths: [],
  rowHeights: [],
  hiddenRows: [],
  freeze: { rows: 0, columns: 0 },
});
const visible = (node) => !node.closest('[hidden]'),
  button = (name) =>
    [...document.querySelectorAll('button')].find(
      (node) => visible(node) && (node.getAttribute('aria-label') ?? node.textContent) === name
    ),
  tab = (name) => [...document.querySelectorAll('[role="tab"]')].find((node) => node.textContent === name);
async function click(node) {
  assert(node);
  assert.equal(node.disabled, false);
  await act(async () => node.click());
}
async function mount(t, extra = {}) {
  frames = [];
  const root = createRoot(document.getElementById('root')),
    sheets = [sheet()],
    props = { sheets, documentId: 'A', storageKey: 'worker-A:base', readOnly: false, onEdit: () => {}, ...extra };
  await act(async () => root.render(React.createElement(SpreadsheetGrid, props)));
  t.after(async () => act(async () => root.unmount()));
  return {
    sheets,
    render: async (patch) => act(async () => root.render(React.createElement(SpreadsheetGrid, { ...props, ...patch }))),
  };
}
test('real strict batch and registry remain pending until owner accepts one normalized atomic batch', async () => {
  const owner = deferred(),
    commands = [
      {
        type: 'setCells',
        sheet: 'Sheet1',
        cells: Array.from({ length: 1201 }, (_, i) => ({ cell: 'A' + (i + 1), value: i })),
      },
      { sheet: 'Sheet1', operation: 'freeze', rows: 1, columns: 0 },
    ],
    before = structuredClone(commands);
  let current = { documentId: 'A', state: 'base', sheet: 'Sheet1', selection: 'A1:A1201', activationEpoch: 1 },
    calls = 0,
    received,
    settled = false;
  const scope = {
    captureBinding: () => current,
    isCurrent: (b) =>
      b.documentId === current.documentId && b.state === current.state && b.activationEpoch === current.activationEpoch,
    availability: () => ({ profile: 'xlsx', active: true, canEdit: true, busy: false }),
  };
  const pending = dispatchRibbonCommand('formulas.auto-sum', scope, () =>
    acceptGridCommandBatch(commands, {
      ...base(),
      onCommandBatch: (batch) => {
        calls++;
        received = batch;
        return owner.promise;
      },
    })
  );
  pending.then(() => {
    settled = true;
  });
  await Promise.resolve();
  assert.equal(calls, 1);
  assert.deepEqual(
    received.filter((c) => c.type === 'setCells').map((c) => c.cells.length),
    [1000, 201]
  );
  assert.equal(received.at(-1).operation, 'freeze');
  assert.equal(settled, false);
  assert.deepEqual(commands, before);
  assert.deepEqual(normalizeGridCommandBatch(commands), received);
  current = { ...current, state: 'accepted-1' };
  owner.resolve();
  const result = await pending;
  assert.equal(result.outcome, 'completed');
  assert.equal(result.current, false);
  assert.equal(result.binding.state, 'base');
  assert.equal(settled, true);
});
test('late strict rejection retains captured owner/current=false; preflight guards never call owner', async () => {
  const owner = deferred();
  let current = { documentId: 'A', state: 'base', sheet: 'Sheet1', selection: 'A1', activationEpoch: 1 };
  const scope = {
    captureBinding: () => current,
    isCurrent: (b) => b.documentId === current.documentId && b.activationEpoch === current.activationEpoch,
    availability: () => ({ profile: 'xlsx', active: true, canEdit: true, busy: false }),
  };
  const error = new Error('owner refused');
  const pending = dispatchRibbonCommand(
    'view.freeze',
    scope,
    () =>
      acceptGridCommandBatch([{ sheet: 'Sheet1', operation: 'freeze', rows: 1, columns: 0 }], {
        ...base(),
        onCommandBatch: () => owner.promise,
      }),
    { mode: 'row' }
  );
  current = { ...current, documentId: 'B', activationEpoch: 2 };
  owner.reject(error);
  const result = await pending;
  assert.equal(result.outcome, 'failed');
  assert.equal(result.error, error);
  assert.equal(result.current, false);
  assert.equal(result.binding.documentId, 'A');
  let calls = 0;
  for (const patch of [
    { active: false },
    { readOnly: true },
    { busy: true },
    { literal: true },
    { specialSelection: true },
  ])
    await assert.rejects(
      acceptGridCommandBatch([], {
        ...base(),
        onCommandBatch: async () => {
          calls++;
        },
        ...patch,
      })
    );
  await assert.rejects(acceptGridCommandBatch([], { ...base() }), /unavailable in this editor/);
  assert.equal(calls, 0);
  await acceptGridCommandBatch(
    [],
    {
      ...base(),
      specialSelection: true,
      onCommandBatch: async () => {
        calls++;
      },
    },
    true
  );
  assert.equal(calls, 1);
});
test('actual Grid Freeze awaits strict callback and suppresses its late rejection after document replacement', async (t) => {
  const owner = deferred(),
    received = [],
    legacy = [];
  const ui = await mount(t, {
    onBatch: (commands) => legacy.push(commands),
    onCommandBatch: (commands) => {
      received.push(commands);
      return owner.promise;
    },
  });
  await click(tab('View'));
  const menu = [...document.querySelectorAll('details')].find((node) => node.textContent.includes('Freeze Panes'));
  menu.open = true;
  await click(button('Freeze first row'));
  assert.equal(received.length, 1);
  assert.deepEqual(received[0], [{ sheet: 'Sheet1', operation: 'freeze', rows: 1, columns: 0 }]);
  assert.equal(legacy.length, 0);
  await ui.render({ documentId: 'B', storageKey: 'worker-B:base' });
  await act(async () => owner.reject(Error('A late failure must not appear in B')));
  assert(!document.body.textContent.includes('A late failure'));
  assert.equal(received.length, 1);
});
test('registered AutoSum/Sort/Filter use strict callback, while local display toggles create no owner history', async (t) => {
  const strict = [],
    legacy = [];
  const ui = await mount(t, {
    ownerIdentity: 'worker-A',
    onBatch: (commands) => legacy.push(commands),
    onCommandBatch: async (commands) => {
      strict.push(commands);
    },
  });
  const go = document.querySelector('input[aria-label="Go to cell"]');
  await act(async () => {
    Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(go, 'A1:A3');
    go.dispatchEvent(new window.Event('input', { bubbles: true }));
  });
  await act(async () =>
    go.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
  );
  await click(tab('Formulas'));
  await click(button('AutoSum'));
  assert.deepEqual(strict[0], [{ type: 'setCells', sheet: 'Sheet1', cells: [{ cell: 'A4', formula: 'SUM(A1:A3)' }] }]);
  await click(tab('Data'));
  const sort = [...document.querySelectorAll('details')].find((node) => node.textContent.includes('Primary column'));
  sort.open = true;
  await click(button('Sort ascending'));
  assert.equal(strict[1][0].operation, 'sortRows');
  sort.open = true;
  await click(button('Filter selected value'));
  assert.equal(strict[2][0].operation, 'filterRows');
  await ui.render({
    sheets: ui.sheets.map((s) => ({
      ...s,
      filter: {
        range: 'A1:A3',
        column: 'A',
        values: [1],
        owned: true,
        revision: 'a'.repeat(64),
        canReapply: true,
        canClear: true,
      },
    })),
  });
  sort.open = true;
  await click(button('Clear filter'));
  assert.deepEqual(strict[3][0], { operation: 'clearFilter', sheet: 'Sheet1', expectedFilterRevision: 'a'.repeat(64) });
  assert.equal(legacy.length, 0);
  await click(tab('View'));
  const count = strict.length;
  await click(document.querySelector('input[aria-label="Gridlines"]'));
  assert.equal(document.querySelector('.sheet-grid').dataset.gridlines, 'false');
  assert.equal(strict.length, count);
  assert.equal(legacy.length, 0);
});
test('generic host without strict acceptance callback explains disabled writes and leaves View/tools usable', async (t) => {
  await mount(t, {
    onBatch: () => {
      throw Error('legacy must not be invoked');
    },
  });
  await click(tab('Formulas'));
  assert.equal(button('AutoSum').disabled, true);
  assert.match(button('AutoSum').title, /unavailable in this editor/);
  assert.equal(button('Names and tables').disabled, false);
  await click(tab('View'));
  assert.equal(document.querySelector('input[aria-label="Gridlines"]').disabled, false);
  const menu = [...document.querySelectorAll('details')].find((node) => node.textContent.includes('Freeze Panes'));
  menu.open = true;
  assert.equal(button('Freeze here').disabled, true);
  assert.match(button('Freeze here').title, /unavailable in this editor/);
});
test('queued zoom reveal/fit frames cannot scroll after newer row selection navigation', async (t) => {
  await mount(t, { onCommandBatch: async () => {} });
  await click(tab('View'));
  const zoom = document.querySelector('input[aria-label="Zoom"]');
  await act(async () => {
    Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(zoom, '125');
    zoom.dispatchEvent(new window.Event('input', { bubbles: true }));
  });
  await act(async () =>
    zoom.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
  );
  assert(frames.some(Boolean));
  const row = [...document.querySelectorAll('tbody th')].find((node) => node.textContent === '3');
  await act(async () => row.click());
  const scroll = document.querySelector('.sheet-grid-scroll');
  scroll.scrollTop = 87;
  scroll.scrollLeft = 49;
  const held = frames;
  frames = [];
  await act(async () => held.forEach((fn) => fn?.()));
  assert.equal(scroll.scrollTop, 87);
  assert.equal(scroll.scrollLeft, 49);
  await click(button('Zoom to selection'));
  assert(frames.some(Boolean));
  const nextRow = [...document.querySelectorAll('tbody th')].find((node) => node.textContent === '4');
  await act(async () => nextRow.click());
  scroll.scrollTop = 93;
  scroll.scrollLeft = 61;
  const fit = frames;
  frames = [];
  await act(async () => fit.forEach((fn) => fn?.()));
  assert.equal(scroll.scrollTop, 93);
  assert.equal(scroll.scrollLeft, 61);
});
test('first Save migrates all sheet preferences before paint without workbook history or unrelated-document copying', async (t) => {
  const key = 'workbook.view-preferences.v1',
    full = { zoom: 125, gridlines: false, headings: true, formulaBar: true, showFormulas: false };
  window.localStorage.setItem(
    key,
    JSON.stringify({
      schema: 1,
      entries: [
        { documentId: 'untitled-save', sheet: 'Sheet1', preferences: full },
        { documentId: 'untitled-save', sheet: 'Sheet2', preferences: { ...full, zoom: 175, headings: false } },
      ],
    })
  );
  let ownerCalls = 0,
    writes = 0;
  const original = window.Storage.prototype.setItem;
  window.Storage.prototype.setItem = function (name, value) {
    if (name === key) writes++;
    return original.call(this, name, value);
  };
  t.after(() => {
    window.Storage.prototype.setItem = original;
  });
  const sheets = [sheet(), { ...sheet(), name: 'Sheet2', path: 'xl/worksheets/sheet2.xml' }],
    ui = await mount(t, {
      sheets,
      documentId: 'durable-save',
      viewPreferenceId: 'untitled-save',
      onCommandBatch: async () => {
        ownerCalls++;
      },
    });
  assert.equal(document.querySelector('.sheet-grid').dataset.zoom, '125');
  assert.equal(document.querySelector('.sheet-grid').dataset.gridlines, 'false');
  await ui.render({ viewPreferenceId: 'canonical-file', storageKey: 'worker-A:saved' });
  assert.equal(document.querySelector('.sheet-grid').dataset.zoom, '125');
  assert.equal(writes, 1);
  assert.equal(ownerCalls, 0);
  const saved = JSON.parse(window.localStorage.getItem(key));
  assert.equal(saved.entries.length, 2);
  assert(saved.entries.every((entry) => entry.documentId === 'canonical-file'));
  assert.deepEqual(saved.entries.find((entry) => entry.sheet === 'Sheet1').preferences, full);
  await click([...document.querySelectorAll('.sheet-tabs button')].find((node) => node.textContent === 'Sheet2'));
  assert.equal(document.querySelector('.sheet-grid').dataset.zoom, '175');
  assert.equal(document.querySelector('.sheet-grid').dataset.headings, 'false');
  await ui.render({ documentId: 'reopened-durable', viewPreferenceId: 'canonical-file', storageKey: 'reopen:base' });
  assert.equal(document.querySelector('.sheet-grid').dataset.zoom, '175');
  assert.equal(writes, 1);
  await ui.render({ documentId: 'other-durable', viewPreferenceId: 'other-file', storageKey: 'other:base' });
  assert.equal(document.querySelector('.sheet-grid').dataset.zoom, '100');
  assert.equal(document.querySelector('.sheet-grid').dataset.headings, 'true');
  assert.equal(writes, 1);
  assert.equal(ownerCalls, 0);
  assert(JSON.parse(window.localStorage.getItem(key)).entries.every((entry) => entry.documentId === 'canonical-file'));
});
