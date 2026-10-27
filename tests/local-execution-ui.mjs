import './support/scratch-dir.mjs';
/** Actual source React preview/run controls under controlled owner callbacks. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
const out = await fs.mkdtemp(path.resolve('.build/tests/au2-run-ui-')),
  file = path.join(out, 'ui.mjs');
await build({
  stdin: {
    contents: 'export {WorkbookLocalExecution} from "./ui/WorkbookLocalExecution";',
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
  { WorkbookLocalExecution } = await import(pathToFileURL(file));
const tick = () => new Promise((r) => setTimeout(r, 0)),
  flush = async () => {
    for (let i = 0; i < 3; i++) await act(async () => tick());
  },
  button = (name) => [...document.querySelectorAll('button')].find((b) => b.textContent === name);
async function mount(t, patch = {}) {
  const root = createRoot(document.getElementById('root')),
    writes = [];
  let args = {
    frame: {
      key: 'owner/head',
      ownerKey: 'owner/selection',
      sheet: 'Sheet1',
      selection: 'B3:C7',
      active: true,
      apply: async (commands) => writes.push(commands),
    },
    intent: { serial: 1, action: 'freeze-selection' },
    ...patch,
  };
  const render = async (p) => {
    args = { ...args, ...p };
    await act(async () => root.render(React.createElement(WorkbookLocalExecution, args)));
    await flush();
  };
  await render({});
  t.after(async () => act(async () => root.unmount()));
  return {
    writes,
    get args() {
      return args;
    },
    render,
  };
}
after(async () => {
  dom.window.close();
});
test('opening previews captured worksheet/selection and boundary; explicit Run makes one acknowledged write', async (t) => {
  const h = await mount(t);
  assert.equal(h.writes.length, 0);
  assert(document.body.textContent.includes('B3:C7'));
  assert(document.body.textContent.includes('Freeze preceding rows: 2'));
  assert(!button('Run action').disabled);
  await act(async () => button('Run action').click());
  await flush();
  assert.deepEqual(h.writes, [[{ operation: 'freeze', sheet: 'Sheet1', rows: 2, columns: 1 }]]);
  assert(document.body.textContent.includes('completed for Sheet1'));
  assert(button('Run action').disabled);
});
test('Cancel drops preview; head/selection retirement requires another deliberate Preview', async (t) => {
  const h = await mount(t);
  await act(async () => button('Cancel action preview').click());
  await flush();
  assert(button('Run action').disabled);
  assert.equal(h.writes.length, 0);
  await act(async () => button('Preview action').click());
  await flush();
  await h.render({ frame: { ...h.args.frame, key: 'new-selection', ownerKey: 'owner/C4', selection: 'C4' } });
  assert(button('Run action').disabled);
  assert(document.body.textContent.includes('Preview the action again'));
  await act(async () => button('Preview action').click());
  await flush();
  assert(!button('Run action').disabled);
});
test('closing/reopening the same intent does not prepare or run automatically; readonly refuses controls', async (t) => {
  const h = await mount(t);
  await h.render({ frame: { ...h.args.frame, active: false } });
  await h.render({ frame: { ...h.args.frame, active: true } });
  assert(button('Run action').disabled);
  assert.equal(h.writes.length, 0);
  await h.render({ frame: { ...h.args.frame, reason: 'Read-only' } });
  assert(button('Preview action').disabled);
  assert(document.body.textContent.includes('Read-only'));
});
test('Run is disabled during the awaited write, cancellation cannot undo submission, and native refusal is shown', async (t) => {
  let reject;
  const h = await mount(t, {
    frame: {
      key: 'owner/head',
      ownerKey: 'owner/selection',
      sheet: 'Sheet1',
      selection: 'B3',
      active: true,
      apply: () => new Promise((_, no) => (reject = no)),
    },
  });
  await act(async () => button('Run action').click());
  assert(button('Run action').disabled);
  assert(button('Cancel action preview').disabled);
  await act(async () => reject(Error('Source has filters')));
  await flush();
  assert(document.body.textContent.includes('Source has filters'));
  assert(button('Run action').disabled);
});
