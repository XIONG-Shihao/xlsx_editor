import './support/scratch-dir.mjs';
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';

const componentRoot = fileURLToPath(new URL('../', import.meta.url));
const prefix = path.join(componentRoot, '.build/tests/ribbon-component');
await fs.mkdir(path.dirname(prefix), { recursive: true });
const directory = await fs.mkdtemp(prefix + '-');
const entry = path.join(directory, 'entry.ts'),
  module = path.join(directory, 'ribbon.mjs');
await fs.writeFile(
  entry,
  `export {SpreadsheetRibbon,RibbonGroup} from ${JSON.stringify(path.join(componentRoot, 'ui/SpreadsheetRibbon.tsx'))};export {configureXlsxHost} from ${JSON.stringify(path.join(componentRoot, 'ui/host.tsx'))};`
);
await build({
  entryPoints: [entry],
  bundle: true,
  format: 'esm',
  platform: 'browser',
  jsx: 'automatic',
  outfile: module,
  external: ['react', 'react/jsx-runtime'],
  logLevel: 'silent',
});
const { SpreadsheetRibbon, RibbonGroup, configureXlsxHost } = await import(pathToFileURL(module));
const dom = new JSDOM('<!doctype html><div id="root"></div>', { url: 'http://localhost' });
Object.assign(globalThis, {
  window: dom.window,
  document: dom.window.document,
  HTMLElement: dom.window.HTMLElement,
  HTMLDetailsElement: dom.window.HTMLDetailsElement,
  Node: dom.window.Node,
  IS_REACT_ACT_ENVIRONMENT: true,
});
globalThis.ResizeObserver = class {
  constructor(fn) {
    this.fn = fn;
  }
  observe() {
    this.fn([{ contentRect: { width: 1400 } }]);
  }
  disconnect() {}
};
globalThis.requestAnimationFrame = (fn) => {
  fn();
  return 1;
};
configureXlsxHost({
  api: { copyText: async () => {}, readClipboardText: async () => '' },
  Icon: () => React.createElement('svg', { 'aria-hidden': true }),
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
    'IS_REACT_ACT_ENVIRONMENT',
  ])
    delete globalThis[key];
});
const visible = (node) => !node.closest('[hidden]');
const button = (name) =>
  [...document.querySelectorAll('button')].find(
    (node) => visible(node) && (node.getAttribute('aria-label') ?? node.textContent) === name
  );
const checkbox = (name) =>
  [...document.querySelectorAll('input[type="checkbox"]')].find(
    (node) => visible(node) && node.getAttribute('aria-label') === name
  );
const tab = (name) => [...document.querySelectorAll('[role="tab"]')].find((node) => node.textContent === name);
async function click(node) {
  assert(node);
  assert.equal(node.disabled, false);
  await act(async () => node.click());
}
async function key(node, key) {
  await act(async () =>
    node.dispatchEvent(new window.KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }))
  );
}
async function mount(t, extra = {}) {
  let available = { profile: 'xlsx', active: true, canEdit: true, busy: false },
    binding = { documentId: 'A', state: 'base', sheet: 'Sheet1', selection: 'A1:C3', activationEpoch: 1 };
  const result = [],
    changes = [],
    freezes = [],
    normal = [],
    zoomSelection = [],
    panes = [],
    mutations = [],
    root = createRoot(document.getElementById('root'));
  const scope = {
    captureBinding: () => binding,
    isCurrent: (b) => Object.keys(binding).every((key) => binding[key] === b[key]),
    availability: () => available,
  };
  let properties = extra;
  function App() {
    const [prefs, setPrefs] = React.useState({
        zoom: 100,
        gridlines: true,
        headings: true,
        formulaBar: true,
        showFormulas: false,
      }),
      [pane, setPane] = React.useState(null);
    return React.createElement(
      SpreadsheetRibbon,
      {
        canUndo: false,
        canRedo: false,
        tools: [{ id: 'names', label: 'Name Manager' }],
        activeTool: pane,
        taskPaneId: 'dock',
        commandScope: scope,
        onCommandResult: (id, reply) => result.push({ id, reply }),
        onToolToggle: (id) => {
          panes.push(id);
          setPane((current) => (current === id ? null : id));
        },
        view: {
          ...prefs,
          onChange: (patch) => {
            changes.push(patch);
            setPrefs((value) => ({ ...value, ...patch }));
          },
          onZoomToSelection: () => {
            zoomSelection.push(binding.selection);
            setPrefs((value) => ({ ...value, zoom: 150 }));
          },
          onNormal: () => normal.push('worksheet'),
          onFreeze: (mode) => freezes.push(mode),
          canFreeze: true,
        },
        commands: [{ id: 'formulas.auto-sum', handler: (b) => mutations.push(b) }],
        ...properties,
      },
      React.createElement(
        RibbonGroup,
        { name: 'Font', icon: 'font', collapseBelow: 0 },
        React.createElement('input', { 'aria-label': 'Home draft', defaultValue: '11' }),
        React.createElement('small', null, 'Font')
      )
    );
  }
  await act(async () => root.render(React.createElement(App)));
  t.after(async () => act(async () => root.unmount()));
  return {
    changes,
    freezes,
    normal,
    zoomSelection,
    panes,
    mutations,
    result,
    setAvailability: (value) => {
      available = { ...available, ...value };
    },
    setBinding: (value) => {
      binding = { ...binding, ...value };
    },
    render: async (value) => {
      properties = value;
      await act(async () => root.render(React.createElement(App)));
    },
  };
}
test('direct function categories open the shared dock with the captured owner and do not mutate cells', async (t) => {
  const ids = ['financial', 'logical', 'text', 'date-time', 'lookup-reference', 'math-trig', 'more-functions'].map(
    (name) => 'formulas.' + name
  );
  const opened = [],
    commands = ids.map((id) => ({
      id,
      pressed: id === 'formulas.financial',
      handler: (binding) => opened.push({ id, binding }),
    }));
  const ui = await mount(t, { commands });
  await click(tab('Formulas'));
  for (const id of ids) {
    const control = document.querySelector(`button[data-ribbon-command="${id}"]`);
    assert(visible(control));
    assert.equal(control.disabled, false);
    assert.equal(control.getAttribute('aria-controls'), 'dock');
    await click(control);
  }
  assert.deepEqual(
    opened.map((item) => item.id),
    ids
  );
  assert(
    opened.every(
      (item) => item.binding.documentId === 'A' && item.binding.state === 'base' && item.binding.selection === 'A1:C3'
    )
  );
  assert.equal(
    document.querySelector('[data-ribbon-command="formulas.financial"]').getAttribute('aria-pressed'),
    'true'
  );
  assert.equal(
    document.querySelector('[data-ribbon-command="formulas.logical"]').getAttribute('aria-pressed'),
    'false'
  );
  assert.equal(ui.mutations.length, 0);
  const groups = [
    ...document.querySelectorAll('.sheet-ribbon-tabpanel:not([hidden]) .sheet-ribbon-group-content[role="group"]'),
  ].map((node) => node.getAttribute('aria-label'));
  assert(groups.indexOf('Function library') < groups.indexOf('Defined names'));
  ui.setAvailability({ active: false });
  await ui.render({ commands });
  const inactive = document.querySelector('[data-ribbon-command="formulas.financial"]');
  assert.equal(inactive.disabled, true);
  await act(async () => inactive.dispatchEvent(new window.MouseEvent('click', { bubbles: true })));
  assert.equal(opened.length, 7);
});
test('local View controls execute their actual callbacks and share Show formulas checked state with Formulas', async (t) => {
  const ui = await mount(t);
  await click(tab('View'));
  await click(checkbox('Gridlines'));
  assert.equal(checkbox('Gridlines').checked, false);
  await click(checkbox('Headings'));
  await click(checkbox('Formula bar'));
  await click(checkbox('Show formulas'));
  assert.deepEqual(ui.changes, [
    { gridlines: false },
    { headings: false },
    { formulaBar: false },
    { showFormulas: true },
  ]);
  assert.equal(ui.mutations.length, 0);
  assert.equal(ui.freezes.length, 0);
  await click(tab('Formulas'));
  assert.equal(checkbox('Show formulas').checked, true);
  await click(checkbox('Show formulas'));
  await click(tab('View'));
  assert.equal(checkbox('Show formulas').checked, false);
  await click(button('Normal'));
  assert.deepEqual(ui.normal, ['worksheet']);
  assert(ui.result.every((item) => item.reply.outcome === 'completed' && item.reply.current));
});
test('zoom drafts validate integer bounds, commit Enter, and real 100%/selection callbacks update visible value', async (t) => {
  const ui = await mount(t);
  await click(tab('View'));
  const input = document.querySelector('input[data-ribbon-command="view.zoom"]');
  async function change(value) {
    await act(async () => {
      const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
      setter.call(input, value);
      input.dispatchEvent(new window.Event('input', { bubbles: true }));
    });
  }
  await change('125');
  await key(input, 'Enter');
  assert.deepEqual(ui.changes, [{ zoom: 125 }]);
  assert.equal(input.value, '125');
  await change('125.5');
  await key(input, 'Enter');
  assert.equal(ui.changes.length, 1);
  assert(document.querySelector('[role="alert"]'));
  await key(input, 'Escape');
  assert.equal(input.value, '125');
  await click(button('100%'));
  assert.equal(input.value, '100');
  await click(button('Zoom to selection'));
  assert.equal(input.value, '150');
  assert.deepEqual(ui.zoomSelection, ['A1:C3']);
});
test('all Freeze choices reach the real workbook callback with exact argument and read-only blocks mutation', async (t) => {
  const ui = await mount(t);
  await click(tab('View'));
  const menu = [...document.querySelectorAll('details')].find((node) => node.textContent.includes('Freeze Panes'));
  menu.open = true;
  for (const [mode, label] of [
    ['selection', 'Freeze here'],
    ['row', 'Freeze first row'],
    ['column', 'Freeze first column'],
    ['none', 'Unfreeze'],
  ]) {
    menu.open = true;
    await click(button(label));
    assert.equal(ui.freezes.at(-1), mode);
  }
  ui.setAvailability({ canEdit: false });
  await ui.render({});
  await click(tab('View'));
  menu.open = true;
  assert.equal(button('Freeze here').disabled, true);
  assert.match(button('Freeze here').title, /read-only/);
  assert.equal(checkbox('Gridlines').disabled, false);
});
test('inspection launchers dispatch scoped commands while passive unsupported commands/tabs are absent', async (t) => {
  const ui = await mount(t);
  await click(tab('Formulas'));
  await click(button('Name Manager'));
  assert.deepEqual(ui.panes, ['names']);
  assert.equal(ui.result.at(-1).id, 'names.manager');
  assert.equal(ui.result.at(-1).reply.binding.documentId, 'A');
  assert.equal(button('Name Manager').getAttribute('aria-controls'), 'dock');
  ui.setAvailability({ profile: 'passive', canEdit: false, busy: true });
  await ui.render({});
  assert.equal(button('AutoSum'), undefined);
  await click(button('Name Manager'));
  assert.deepEqual(ui.panes, ['names', 'names']);
  await click(tab('View'));
  assert.equal(
    [...document.querySelectorAll('details')].some((node) => node.textContent.includes('Freeze Panes')),
    false
  );
  assert.equal(checkbox('Gridlines').disabled, false);
  assert(
    ![...document.querySelectorAll('[role="tab"]')].some((node) =>
      ['Draw', 'Review', 'Automate'].includes(node.textContent)
    )
  );
});
test('inactive commands skip callbacks and Home state/roving keyboard behavior survive the expanded toolbar', async (t) => {
  const ui = await mount(t);
  const home = document.querySelector('input[aria-label="Home draft"]');
  home.value = '14';
  await click(tab('View'));
  ui.setAvailability({ active: false });
  await ui.render({});
  assert.equal(checkbox('Gridlines').disabled, true);
  checkbox('Gridlines').click();
  assert.equal(ui.changes.length, 0);
  await key(tab('View'), 'Home');
  assert.equal(document.activeElement, tab('Home'));
  assert.equal(document.querySelector('input[aria-label="Home draft"]'), home);
  assert.equal(home.value, '14');
  await key(tab('Home'), 'End');
  assert.equal(document.activeElement, tab('View'));
});
