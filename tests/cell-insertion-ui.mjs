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
  out = await fs.mkdtemp(path.join(component, '.build/tests/cell-insertion-ui-')),
  bundle = path.join(out, 'ui.mjs');
await build({
  stdin: {
    contents: 'export {WorkbookCellInsertions} from "./ui/WorkbookCellInsertions"',
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
  { WorkbookCellInsertions } = await import(pathToFileURL(bundle));
after(() => dom.window.close());
const tick = () => new Promise((r) => setTimeout(r, 0)),
  flush = async () => {
    for (let i = 0; i < 3; i++) await act(async () => tick());
  },
  props = (n) => Object.entries(n).find(([k]) => k.startsWith('__reactProps$'))[1],
  button = (s) => [...document.querySelectorAll('button')].find((n) => n.textContent === s);
async function mount(t, patch = {}, mode = 'checkbox') {
  const root = createRoot(document.getElementById('root')),
    writes = [];
  let frame = {
    key: 'A/1/epoch1',
    family: 'A/cell/epoch1',
    sheet: 'Main',
    cell: 'D1',
    active: true,
    writable: true,
    read: async () => null,
    apply: async (commands) => writes.push(commands),
    ...patch,
  };
  const render = async (value = {}, nextMode = mode) => {
    frame = { ...frame, ...value };
    mode = nextMode;
    await act(async () => root.render(React.createElement(WorkbookCellInsertions, { frame, mode })));
    await flush();
  };
  await render();
  t.after(async () => act(async () => root.unmount()));
  return { writes, render };
}
async function click(node) {
  assert(node);
  await act(async () => node.dispatchEvent(new window.MouseEvent('click', { bubbles: true, cancelable: true })));
  await flush();
}
test('blank-cell Insert emits one semantic command and discloses the Boolean profile', async (t) => {
  const ui = await mount(t);
  assert.match(document.body.textContent, /ordinary TRUE\/FALSE cells/);
  assert.deepEqual(ui.writes, []);
  await click(button('Insert checkbox'));
  assert.deepEqual(ui.writes, [
    [{ type: 'booleanCell', action: 'insert', sheet: 'Main', cell: 'D1', expectedValue: null }],
  ]);
});
test('existing Boolean checkbox toggles/removes the displayed value; readonly blocks both', async (t) => {
  const ui = await mount(t, { read: async () => ({ value: false, dataType: 'boolean' }) });
  const input = document.querySelector('input[type=checkbox]');
  assert.equal(input.checked, false);
  await act(async () => props(input).onChange());
  await flush();
  await click(button('Remove checkbox'));
  assert.deepEqual(
    ui.writes.map((c) => c[0].action),
    ['toggle', 'remove']
  );
  assert(ui.writes.every((c) => c[0].expectedValue === false));
  await ui.render({ writable: false, reason: 'Read-only' });
  assert.equal(document.querySelector('fieldset').disabled, true);
  await click(button('Remove checkbox'));
  assert.equal(ui.writes.length, 2);
});
test('all three Symbols buttons produce literal appends to the captured cell', async (t) => {
  const ui = await mount(t, { read: async () => ({ value: 'prefix', dataType: 'text' }) }, 'symbols');
  for (const s of ['Omega', 'Plus/minus', 'Degree'])
    await click(button([...document.querySelectorAll('button')].find((n) => n.textContent.includes(s)).textContent));
  assert.deepEqual(
    ui.writes.map((c) => c[0].value),
    ['prefixΩ', 'prefix±', 'prefix°']
  );
  assert(ui.writes.every((c) => c[0].expectedValue === 'prefix' && c[0].expectedFormula === null));
});
test('retired button handlers cannot write another cell or a closed/reopened pane', async (t) => {
  const ui = await mount(t);
  const old = props(button('Insert checkbox')).onClick;
  await ui.render({ key: 'B/1/epoch2', family: 'B/cell/epoch2', sheet: 'Dest', cell: 'A2' });
  await act(async () => old());
  await flush();
  assert.deepEqual(ui.writes, []);
  assert.match(document.body.textContent, /Dest!A2/);
});
test('formula targets explain refusal and disable Checkbox/Symbols rather than overwrite', async (t) => {
  const ui = await mount(t, { read: async () => ({ value: 1, formula: 'A1+1' }) });
  assert.equal(document.querySelector('fieldset').disabled, true);
  assert.match(document.body.textContent, /literal TRUE\/FALSE/);
  await ui.render({ key: 'symbols' }, 'symbols');
  assert.equal(document.querySelector('fieldset').disabled, true);
  assert.match(document.body.textContent, /blank or plain-text/);
  assert.deepEqual(ui.writes, []);
});
