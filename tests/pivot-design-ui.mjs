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
  out = await fs.mkdtemp(path.join(component, '.build/tests/pivot-design-ui-')),
  bundle = path.join(out, 'ui.mjs');
await build({
  stdin: {
    contents:
      'export {SpreadsheetGrid} from "./ui/SpreadsheetGrid"; export {configureXlsxHost} from "./ui/host"; export {setPivotLayoutContext,pivotLayoutDraftReason,pivotLayoutReason} from "./ui/pivotLayout"; export {pivotCellStyle} from "./ui/pivotStyles";',
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
  {
    SpreadsheetGrid,
    configureXlsxHost,
    setPivotLayoutContext,
    pivotLayoutDraftReason,
    pivotLayoutReason,
    pivotCellStyle,
  } = await import(pathToFileURL(bundle));
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
  sample: 'tabular-light',
  grandTotals: true,
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

const reactProps = (node) => Object.entries(node).find(([key]) => key.startsWith('__reactProps$'))?.[1];
const deferred = () => {
  let resolve, reject;
  const promise = new Promise((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
};
const panel = () => document.querySelector('section[aria-label="PivotTable Design"]'),
  input = (name) => document.querySelector('[aria-label="' + name + '"]');
const choose = async (sample) => {
  await act(async () => {
    input('Pivot layout sample').value = sample;
    input('Pivot layout sample').dispatchEvent(new window.Event('change', { bubbles: true }));
  });
};
async function design() {
  await click(tab('Design'));
  if (!panel() || !visible(panel())) await click(button('PivotTable Styles'));
  assert(panel() && visible(panel()));
}
test('Design has three real sample choices and one retained pane; draft preview does not paint accepted cells', async (t) => {
  await mount(t);
  assert(tab('Design'));
  await design();
  assert.deepEqual(
    [...input('Pivot layout sample').options].map((n) => n.textContent),
    ['Tabular Light', 'Compact Blue', 'Outline Dark']
  );
  assert.equal(button('Apply pivot layout').disabled, true);
  assert.equal(document.querySelectorAll('aside.sheet-task-pane:not([hidden])').length, 1);
  const accepted = document.querySelector('[data-pivot-cell="B1"]');
  assert.equal(accepted.dataset.pivotStyle, 'tabular-light');
  const color = accepted.style.backgroundColor;
  await choose('compact-blue');
  assert.equal(accepted.style.backgroundColor, color);
  assert.match(panel().textContent, /Row Labels/);
  assert.equal(button('Apply pivot layout').disabled, false);
});
test('Apply captures exact sample/totals and the selected revision in one command', async (t) => {
  const ui = await mount(t);
  await design();
  await choose('compact-blue');
  await click(input('Show pivot grand total'));
  await click(button('Apply pivot layout'));
  assert.deepEqual(ui.calls, [
    [
      {
        operation: 'setPivotLayout',
        sheet: 'Summary',
        name: 'SalesPivot',
        expectedPivotRevision: 'pivot-1',
        sample: 'compact-blue',
        grandTotals: false,
      },
    ],
  ]);
});
test('accepted metadata controls header/band/total styling; disabling totals does not style the last group as a total', async (t) => {
  const ui = await mount(t);
  const header = () => document.querySelector('[data-pivot-cell="B1"]');
  assert.equal(header().dataset.pivotRole, 'header');
  assert.equal(document.querySelector('[data-pivot-cell="B4"]').dataset.pivotRole, 'total');
  await ui.render({
    pivots: [{ ...pivot, sample: 'outline-dark', grandTotals: false, outputRange: 'A1:B3', revision: 'pivot-2' }],
  });
  assert.equal(header().dataset.pivotStyle, 'outline-dark');
  assert.equal(header().style.backgroundColor, 'rgb(51, 51, 51)');
  assert.equal(document.querySelector('[data-pivot-cell="B3"]').dataset.pivotRole, 'body');
  assert.equal(document.querySelector('[data-pivot-cell="B4"]'), null);
});
test('Close, reopen, Fields switch, navigation and owner changes retire old layout callbacks', async (t) => {
  const ui = await mount(t);
  await design();
  await choose('compact-blue');
  let old = reactProps(button('Apply pivot layout')).onClick;
  await click(button('Close workbook tools'));
  await click(button('PivotTable Styles'));
  await act(async () => old());
  assert.deepEqual(ui.calls, []);
  await choose('compact-blue');
  old = reactProps(button('Apply pivot layout')).onClick;
  await click(tab('PivotTable Analyze'));
  await click(button('PivotTable Fields'));
  await act(async () => old());
  assert.deepEqual(ui.calls, []);
  assert.equal(document.querySelectorAll('aside.sheet-task-pane:not([hidden])').length, 1);
  await design();
  await choose('outline-dark');
  old = reactProps(button('Apply pivot layout')).onClick;
  await go('C9');
  assert.equal(tab('Design'), undefined);
  await go('A1');
  await act(async () => old());
  assert.deepEqual(ui.calls, []);
  await design();
  await choose('compact-blue');
  old = reactProps(button('Apply pivot layout')).onClick;
  await ui.render({ ownerIdentity: 'other', storageKey: 'other:head-2' });
  await act(async () => old());
  assert.deepEqual(ui.calls, []);
});
test('layout and refresh share one physical lane; late failure cannot enter another owner pane', async (t) => {
  const job = deferred(),
    calls = [],
    ui = await mount(t, {
      onCommandBatch: (commands) => {
        calls.push(commands);
        return job.promise;
      },
    });
  await design();
  await choose('compact-blue');
  const old = reactProps(button('Apply pivot layout')).onClick;
  await act(async () => {
    old();
    old();
  });
  assert.equal(calls.length, 1);
  assert.equal(button('Applying pivot layout…').matches(':disabled'), true);
  await click(tab('PivotTable Analyze'));
  await click(button('PivotTable Fields'));
  assert.equal(button('Refreshing pivot…').disabled, true);
  await ui.render({ ownerIdentity: 'new-owner', documentId: 'new-doc' });
  await act(async () => job.reject(Error('Old layout failure')));
  assert(!document.body.textContent.includes('Old layout failure'));
  assert.equal(calls.length, 1);
});
test('readonly, protected and missing-profile pivots allow inspection but disable editing', async (t) => {
  const ui = await mount(t, { readOnly: true });
  await design();
  assert.equal(input('Pivot layout sample').closest('fieldset').disabled, true);
  assert.match(panel().textContent, /read-only/);
  assert.deepEqual(ui.calls, []);
  await ui.render({
    readOnly: false,
    pivots: [{ ...pivot, revision: null, canEdit: false, reason: 'Protected cache' }],
  });
  assert.equal(input('Pivot layout sample').closest('fieldset').disabled, true);
  assert.match(panel().textContent, /Protected cache/);
  await ui.render({ pivots: [{ ...pivot, sample: undefined, grandTotals: undefined }] });
  assert.match(panel().textContent, /no supported layout profile/);
  assert.equal(input('Pivot layout sample').closest('fieldset').disabled, true);
});
test('retiring Design restores the previous ordinary tab and repeated tool toggle closes reliably', async (t) => {
  await mount(t);
  await click(tab('Formulas'));
  await design();
  await click(button('PivotTable Styles'));
  assert.equal(document.querySelectorAll('aside.sheet-task-pane:not([hidden])').length, 0);
  await click(button('PivotTable Styles'));
  await go('C9');
  assert.equal(tab('Design'), undefined);
  assert.equal(tab('Formulas').getAttribute('aria-selected'), 'true');
  assert.equal(document.querySelectorAll('aside.sheet-task-pane:not([hidden])').length, 0);
});
test('controller clones a closed draft, fences stale/permission changes and preserves accepted truth after retirement', async () => {
  const calls = [],
    job = deferred(),
    captured = {
      key: 'one',
      pivot,
      active: true,
      normal: true,
      readOnly: false,
      busy: false,
      accept: (commands) => {
        calls.push(commands);
        return job.promise;
      },
    },
    draft = { sample: 'compact-blue', grandTotals: false };
  let live = captured;
  const accepted = setPivotLayoutContext(captured, () => live, draft);
  draft.sample = 'outline-dark';
  draft.grandTotals = true;
  assert.equal(calls[0][0].sample, 'compact-blue');
  assert.equal(calls[0][0].grandTotals, false);
  live = null;
  job.resolve();
  assert.deepEqual(await accepted, { accepted: true, key: 'one', current: false });
  await assert.rejects(
    setPivotLayoutContext(captured, () => null, { sample: 'compact-blue', grandTotals: true }),
    /changed/
  );
  await assert.rejects(
    setPivotLayoutContext(captured, () => ({ ...captured, readOnly: true }), {
      sample: 'compact-blue',
      grandTotals: true,
    }),
    /read-only/
  );
  for (const value of [
    null,
    [],
    { sample: 'missing', grandTotals: true },
    { sample: 'tabular-light', grandTotals: 'false' },
    { sample: 'tabular-light', grandTotals: true, range: 'A1' },
  ])
    assert(pivotLayoutDraftReason(value));
  assert(pivotLayoutReason({ ...captured, pivot: { ...pivot, sample: undefined } }));
});
test('style projection refuses unknown/ambiguous/protected/unbounded definitions and leaves unrelated cells plain', () => {
  for (const list of [
    [{ ...pivot, canEdit: false }],
    [{ ...pivot, revision: null }],
    [{ ...pivot, sample: 'unknown' }],
    [{ ...pivot, grandTotals: undefined }],
    [pivot, pivot],
    [{ ...pivot, outputRange: 'A1:B600' }],
  ])
    assert.equal(pivotCellStyle(list, 'Summary', 'A1'), undefined);
  assert.equal(pivotCellStyle([pivot], 'Data', 'A1'), undefined);
  assert.equal(pivotCellStyle([pivot], 'Summary', 'C1'), undefined);
  const compact = { ...pivot, sample: 'compact-blue' };
  assert.equal(pivotCellStyle([compact], 'Summary', 'A2').paddingLeft, 12);
  assert.equal(pivotCellStyle([compact], 'Summary', 'B4').role, 'total');
});
