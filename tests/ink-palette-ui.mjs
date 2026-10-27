import './support/scratch-dir.mjs';
/** Source palette React controls; scripted actual app/cold persistence is separate. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
const out = await fs.mkdtemp(path.resolve('.build/tests/dw-m3-ui-')),
  file = path.join(out, 'ui.mjs');
await build({
  stdin: {
    contents: 'export {WorkbookInkPalette} from "./ui/WorkbookInkPalette";',
    resolveDir: process.cwd(),
    loader: 'tsx',
  },
  bundle: true,
  outfile: file,
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
  Node: dom.window.Node,
  IS_REACT_ACT_ENVIRONMENT: true,
});
const { createRoot } = await import('react-dom/client'),
  { WorkbookInkPalette } = await import(pathToFileURL(file));
async function mount(t, patch = {}) {
  const root = createRoot(document.getElementById('root')),
    calls = [];
  let args = { selected: 'black', disabled: false, onChoose: (pen) => calls.push(pen), ...patch };
  const render = async (p) => {
    args = { ...args, ...p };
    await act(async () => root.render(React.createElement(WorkbookInkPalette, args)));
  };
  await render({});
  t.after(async () => act(async () => root.unmount()));
  return { calls, render };
}
after(async () => {
  dom.window.close();
});
test('three named SVG samples show widths/opacity and exactly one remembered selected choice', async (t) => {
  const h = await mount(t);
  const buttons = [...document.querySelectorAll('button')];
  assert.deepEqual(
    buttons.map((b) => b.getAttribute('aria-label')),
    ['Black pen', 'Red pen', 'Yellow highlighter']
  );
  assert.equal(document.querySelectorAll('[aria-pressed="true"]').length, 1);
  assert.equal(document.querySelector('[aria-pressed="true"]').dataset.inkPreset, 'black');
  assert.equal(document.querySelectorAll('svg').length, 3);
  assert.equal(document.querySelectorAll('img,canvas,dialog').length, 0);
  await act(async () => buttons[1].click());
  assert.deepEqual(h.calls, ['red']);
  await h.render({ selected: 'red' });
  assert.equal(document.querySelector('[aria-pressed="true"]').dataset.inkPreset, 'red');
  assert(buttons[2].textContent.includes('12 px'));
  assert(buttons[2].textContent.includes('30%'));
});
test('disabled controls cannot choose a preset; session-only warning remains visible and useful', async (t) => {
  const h = await mount(t, {
    selected: 'yellow-highlighter',
    disabled: true,
    reason: 'Read-only',
    warning: 'Pen choice is remembered for this session only.',
  });
  for (const b of document.querySelectorAll('button')) {
    assert(b.disabled);
    assert.equal(b.title, 'Read-only');
    await act(async () => b.click());
  }
  assert.equal(h.calls.length, 0);
  assert.equal(
    document.querySelector('[role="status"]').textContent,
    'Pen choice is remembered for this session only.'
  );
  assert.equal(document.querySelector('[aria-pressed="true"]').dataset.inkPreset, 'yellow-highlighter');
});
