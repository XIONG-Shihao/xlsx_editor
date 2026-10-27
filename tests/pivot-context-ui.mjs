import './support/scratch-dir.mjs';
/** Actual source Grid/Ribbon/dock integration; mocked owner callbacks are not native proof. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
const component = fileURLToPath(new URL('../', import.meta.url)),
  out = await fs.mkdtemp(path.join(component, '.build/tests/pivot-context-ui-')),
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
  HTMLDetailsElement: dom.window.HTMLDetailsElement,
  Node: dom.window.Node,
  IS_REACT_ACT_ENVIRONMENT: true,
});
Object.defineProperties(window.HTMLElement.prototype, {
  clientWidth: {
    get() {
      return 1400;
    },
  },
  clientHeight: {
    get() {
      return 760;
    },
  },
});
globalThis.ResizeObserver = class {
  constructor(fn) {
    this.fn = fn;
  }
  observe() {
    this.fn([{ contentRect: { width: 1400, height: 760 } }]);
  }
  disconnect() {}
};
globalThis.requestAnimationFrame = () => 1;
globalThis.cancelAnimationFrame = () => {};
const { createRoot } = await import('react-dom/client'),
  { SpreadsheetGrid, configureXlsxHost } = await import(pathToFileURL(bundle));
after(async () => {
  dom.window.close();
});
configureXlsxHost({
  useI18n: () => ({ locale: 'en' }),
  Icon: () => React.createElement('span'),
  api: { copyText: async () => {}, readClipboardText: async () => '' },
});
const pivot = {
  name: 'SalesPivot',
  sheet: 'Summary',
  sourceSheet: 'Data',
  sourceRange: 'A1:B5',
  outputRange: 'A1:B4',
  rowField: 0,
  valueField: 1,
  fieldNames: ['Category', 'Amount'],
  aggregate: 'sum',
  revision: 'pivot-1',
  canEdit: true,
};
const sheets = ['Summary', 'Data'].map((name, i) => ({
  name,
  path: 'xl/worksheets/sheet' + (i + 1) + '.xml',
  rows: 1000,
  columns: 10,
  state: 'visible',
  cells: {},
  merges: [],
  widths: [],
  hiddenRows: [],
  rowHeights: [],
  freeze: { rows: 0, columns: 0 },
}));
const visible = (node) => !node.closest('[hidden]'),
  button = (name) =>
    [...document.querySelectorAll('button')].find(
      (n) => visible(n) && (n.getAttribute('aria-label') ?? n.textContent) === name
    ),
  tab = (name) => [...document.querySelectorAll('[role="tab"]')].find((n) => visible(n) && n.textContent === name),
  click = async (node) => {
    assert(node);
    await act(async () => node.click());
  };
async function go(value) {
  const input = document.querySelector('input[aria-label="Go to cell"]');
  await act(async () => {
    Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(input, value);
    input.dispatchEvent(new window.Event('input', { bubbles: true }));
    input.dispatchEvent(new window.Event('change', { bubbles: true }));
  });
  await act(async () =>
    input.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
  );
}
async function mount(t, patch = {}) {
  const root = createRoot(document.getElementById('root')),
    calls = [];
  let props = {
    sheets,
    documentId: 'doc',
    ownerIdentity: 'owner',
    storageKey: 'owner:head-1',
    pivots: [pivot],
    readOnly: false,
    onEdit: () => {},
    readRange: async () => ({}),
    onCommandBatch: async (commands) => calls.push(commands),
    ...patch,
  };
  const render = async (next) => {
    props = { ...props, ...next };
    await act(async () => root.render(React.createElement(SpreadsheetGrid, props)));
  };
  await render({});
  t.after(async () => act(async () => root.unmount()));
  return { calls, render };
}
test('known selection exposes Analyze and actual fields/source in one dock; unrelated selection retires both', async (t) => {
  await mount(t);
  assert(tab('PivotTable Analyze'));
  await click(tab('PivotTable Analyze'));
  await click(button('PivotTable Fields'));
  const panel = document.querySelector('section[aria-label="PivotTable Fields"]');
  assert(panel);
  assert.match(panel.textContent, /Data!A1:B5/);
  assert.match(panel.textContent, /Category/);
  assert.match(panel.textContent, /SUM.*Amount/);
  assert.equal(document.querySelectorAll('aside.sheet-task-pane:not([hidden])').length, 1);
  await go('C9');
  assert.equal(tab('PivotTable Analyze'), undefined);
  assert.equal(document.querySelector('section[aria-label="PivotTable Fields"]'), null);
  assert.equal(document.querySelectorAll('aside.sheet-task-pane:not([hidden])').length, 0);
});
test('readonly and protected definitions inspect without enabling refresh', async (t) => {
  const ui = await mount(t, { readOnly: true });
  await click(tab('PivotTable Analyze'));
  assert.equal(button('Refresh').disabled, true);
  await click(button('PivotTable Fields'));
  assert.match(document.querySelector('section[aria-label="PivotTable Fields"]').textContent, /read-only/);
  assert.equal(button('Refresh selected pivot').disabled, true);
  await ui.render({
    readOnly: false,
    pivots: [{ ...pivot, revision: null, canEdit: false, reason: 'Imported cache is protected.' }],
  });
  assert.match(
    document.querySelector('section[aria-label="PivotTable Fields"]').textContent,
    /Imported cache is protected/
  );
  assert.equal(button('Refresh selected pivot').disabled, true);
  assert.deepEqual(ui.calls, []);
});
test('contextual refresh uses the selected revision and source navigation adds no command', async (t) => {
  const ui = await mount(t);
  await click(tab('PivotTable Analyze'));
  await click(button('Refresh'));
  assert.deepEqual(ui.calls, [
    [{ operation: 'refreshPivot', sheet: 'Summary', name: 'SalesPivot', expectedPivotRevision: 'pivot-1' }],
  ]);
  await click(button('Go to Source'));
  assert.equal(tab('PivotTable Analyze'), undefined);
  assert.equal(ui.calls.length, 1);
  assert.equal(document.querySelector('.sheet-tabs button.active')?.textContent, 'Data');
  assert.equal(document.querySelector('[aria-label="Go to cell"]').placeholder, 'A1:B5');
});

const reactProps = (node) => Object.entries(node).find(([key]) => key.startsWith('__reactProps$'))?.[1];
const deferred = () => {
  let resolve, reject;
  const promise = new Promise((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
};
test('single dock toggle closes reliably and retiring Analyze restores the previous ordinary tab', async (t) => {
  await mount(t);
  await click(tab('Formulas'));
  await click(tab('PivotTable Analyze'));
  await click(button('PivotTable Fields'));
  assert(document.querySelector('section[aria-label="PivotTable Fields"]'));
  await click(button('PivotTable Fields'));
  assert.equal(document.querySelectorAll('aside.sheet-task-pane:not([hidden])').length, 0);
  await click(button('PivotTable Fields'));
  await click(button('Close workbook tools'));
  assert.equal(document.querySelectorAll('aside.sheet-task-pane:not([hidden])').length, 0);
  await go('C9');
  assert.equal(tab('Formulas').getAttribute('aria-selected'), 'true');
});
test('a captured pane callback cannot retarget after A to B to A navigation or an owner/head/revision change', async (t) => {
  const ui = await mount(t);
  await click(tab('PivotTable Analyze'));
  await click(button('PivotTable Fields'));
  const old = reactProps(button('Refresh selected pivot')).onClick;
  await go('C9');
  await go('A1');
  await act(async () => old());
  assert.deepEqual(ui.calls, []);
  for (const patch of [
    { ownerIdentity: 'other' },
    { storageKey: 'owner:head-2' },
    { pivots: [{ ...pivot, revision: 'pivot-2' }] },
  ]) {
    const captured = reactProps(button('Refresh selected pivot')).onClick;
    await ui.render(patch);
    await act(async () => captured());
    assert.deepEqual(ui.calls, []);
  }
});
test('one physical refresh is awaited and late failure cannot enter a newer owner pane', async (t) => {
  const job = deferred(),
    calls = [],
    ui = await mount(t, {
      onCommandBatch: (commands) => {
        calls.push(commands);
        return job.promise;
      },
    });
  await click(tab('PivotTable Analyze'));
  await click(button('PivotTable Fields'));
  const old = reactProps(button('Refresh selected pivot')).onClick;
  await act(async () => {
    old();
    old();
  });
  assert.equal(calls.length, 1);
  assert.equal(button('Refreshing pivot…').disabled, true);
  await ui.render({ ownerIdentity: 'new-owner', documentId: 'new-doc' });
  await act(async () => job.reject(Error('Old refresh failure')));
  assert(!document.body.textContent.includes('Old refresh failure'));
  assert.equal(button('Refresh selected pivot').disabled, false);
});
test('permission and activity changes retire or disable the context without accepting a command', async (t) => {
  const ui = await mount(t);
  await click(tab('PivotTable Analyze'));
  await click(button('PivotTable Fields'));
  await ui.render({ busy: true });
  assert.equal(button('Refresh selected pivot').disabled, true);
  await ui.render({ busy: false, active: false });
  assert.equal(tab('PivotTable Analyze'), undefined);
  assert.equal(document.querySelector('section[aria-label="PivotTable Fields"]'), null);
  await ui.render({ active: true });
  assert(tab('PivotTable Analyze'));
  assert.deepEqual(ui.calls, []);
});
test('an ordinary refresh refusal is visible only for the original selected pivot', async (t) => {
  await mount(t, {
    onCommandBatch: async () => {
      throw Error('Pivot source requires ordinary current cells');
    },
  });
  await click(tab('PivotTable Analyze'));
  await click(button('PivotTable Fields'));
  await click(button('Refresh selected pivot'));
  assert.match(document.querySelector('aside.sheet-task-pane').textContent, /requires ordinary current cells/);
  await go('C9');
  assert(!document.body.textContent.includes('requires ordinary current cells'));
});

test('source rectangle survives the actual scheduled browser focus callback', async (t) => {
  await mount(t);
  await click(tab('PivotTable Analyze'));
  await click(button('PivotTable Fields'));
  const original = requestAnimationFrame;
  let focus;
  globalThis.requestAnimationFrame = (callback) => {
    focus = callback;
    return 77;
  };
  try {
    await click(button('Go to pivot source'));
    assert(focus);
    await act(async () => focus());
    const input = document.querySelector('[aria-label="Go to cell"]');
    assert.equal(input.placeholder, 'A1:B5');
    assert.equal(document.activeElement, input);
  } finally {
    globalThis.requestAnimationFrame = original;
  }
});

async function changeField(label, value) {
  const node = document.querySelector('[aria-label="' + label + '"]');
  assert(node, label);
  const prototype = node.tagName === 'SELECT' ? window.HTMLSelectElement.prototype : window.HTMLInputElement.prototype;
  await act(async () => {
    Object.getOwnPropertyDescriptor(prototype, 'value').set.call(node, value);
    node.dispatchEvent(new window.Event('input', { bubbles: true }));
    node.dispatchEvent(new window.Event('change', { bubbles: true }));
  });
}
test('field controls apply the original selected revision and a complete captured source/field draft', async (t) => {
  const ui = await mount(t);
  await click(tab('PivotTable Analyze'));
  await click(button('PivotTable Fields'));
  assert.equal(button('Apply pivot fields').disabled, true);
  await changeField('Pivot source range', 'A1:B6');
  await changeField('Pivot value aggregate', 'count');
  await click(button('Apply pivot fields'));
  assert.deepEqual(ui.calls, [
    [
      {
        operation: 'updatePivot',
        sheet: 'Summary',
        name: 'SalesPivot',
        expectedPivotRevision: 'pivot-1',
        sourceSheet: 'Data',
        range: 'A1:B6',
        rowField: 0,
        valueField: 1,
        aggregate: 'count',
      },
    ],
  ]);
});
test('invalid draft and retired field callback cannot apply; readonly controls stay disabled', async (t) => {
  const ui = await mount(t);
  await click(tab('PivotTable Analyze'));
  await click(button('PivotTable Fields'));
  await changeField('Pivot source range', 'A1:A8');
  assert.equal(button('Apply pivot fields').disabled, true);
  assert.match(document.body.textContent, /2–8 source fields/);
  await changeField('Pivot source range', 'A1:B6');
  const old = reactProps(button('Apply pivot fields')).onClick;
  await ui.render({ ownerIdentity: 'replacement' });
  await act(async () => old());
  assert.deepEqual(ui.calls, []);
  await ui.render({ readOnly: true });
  assert(document.querySelector('[aria-label="Pivot source range"]').matches(':disabled'));
});
test('field writes share one physical flight with Refresh and suppress a late error in a newer owner', async (t) => {
  const held = deferred(),
    calls = [],
    ui = await mount(t, {
      onCommandBatch: (commands) => {
        calls.push(commands);
        return held.promise;
      },
    });
  await click(tab('PivotTable Analyze'));
  await click(button('PivotTable Fields'));
  await changeField('Pivot source range', 'A1:B6');
  const old = reactProps(button('Apply pivot fields')).onClick;
  await act(async () => {
    old();
    old();
  });
  assert.equal(calls.length, 1);
  assert.equal(button('Updating pivot…').matches(':disabled'), true);
  assert.equal(button('Refreshing pivot…').disabled, true);
  await ui.render({ ownerIdentity: 'new-owner', documentId: 'new-doc' });
  await act(async () => held.reject(Error('Old field update failed')));
  assert(!document.body.textContent.includes('Old field update failed'));
  assert.equal(calls.length, 1);
});
