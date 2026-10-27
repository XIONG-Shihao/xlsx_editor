import './support/scratch-dir.mjs';
/** VBA pane: late replies from an earlier document scope never reach the current one. */
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
const out = await fs.mkdtemp(path.resolve('.build/tests/vba-scope-'));
const file = path.join(out, 'WorkbookVba.mjs');
await build({
  entryPoints: ['ui/WorkbookVba.tsx'],
  bundle: true,
  format: 'esm',
  platform: 'browser',
  jsx: 'automatic',
  outfile: file,
  external: ['react', 'react/jsx-runtime'],
  logLevel: 'silent',
});
const { WorkbookVba } = await import(pathToFileURL(file));
const dom = new JSDOM('<div id="root"></div>');
globalThis.window = dom.window;
globalThis.document = dom.window.document;
globalThis.IS_REACT_ACT_ENVIRONMENT = true;
const { createRoot } = await import('react-dom/client'),
  root = createRoot(document.getElementById('root')),
  button = (text) => [...document.querySelectorAll('button')].find((n) => n.textContent === text),
  page = (name) => ({ modules: [{ id: name, name, kind: 'standard' }], nextOffset: null, total: 1 }),
  deferred = () => {
    let resolve, reject;
    const promise = new Promise((yes, no) => {
      resolve = yes;
      reject = no;
    });
    return { promise, resolve, reject };
  },
  base = { busy: false, exportValues: async () => ({ cancelled: true }) },
  render = (props) => act(async () => root.render(React.createElement(WorkbookVba, { ...base, ...props }))),
  click = (text) =>
    act(async () => {
      assert(button(text) && !button(text).disabled);
      button(text).click();
    });
try {
  const a = deferred(),
    b = deferred();
  await render({ scope: 'A:base', request: () => a.promise });
  await click('Inspect VBA source');
  assert(button('Export new values-only file…').disabled);
  await render({ scope: 'B:base', request: () => b.promise });
  assert.equal(button('Export new values-only file…').disabled, false);
  await click('Inspect VBA source');
  assert(button('Export new values-only file…').disabled);
  await act(async () => a.resolve(page('OldModule')));
  assert(button('Export new values-only file…').disabled, 'old completion cannot clear B operation');
  assert(!document.body.textContent.includes('OldModule'));
  await act(async () => b.resolve(page('CurrentModule')));
  assert.equal(button('Export new values-only file…').disabled, false);
  assert(document.body.textContent.includes('CurrentModule'));
  const error = deferred();
  await render({ scope: 'C:base', request: () => error.promise });
  await click('Inspect VBA source');
  await render({ scope: 'D:base', request: async () => page('DModule') });
  assert.equal(button('Export new values-only file…').disabled, false);
  await act(async () => error.reject(Error('Old owner failure')));
  assert(!document.body.textContent.includes('Old owner failure'));
  assert.equal(button('Export new values-only file…').disabled, false);
  const evaluation = deferred();
  await render({
    scope: 'E:base',
    request: async (r) => {
      if (r.operation === 'vbaModules') return page('EModule');
      if (r.operation === 'vbaModule')
        return { name: 'EModule', sha256: 'a'.repeat(64), offset: 0, code: 'MyFunc = 42', total: 11, nextOffset: null };
      if (r.operation === 'vbaFunctions')
        return { functions: [{ name: 'MyFunc', parameters: [], returnType: 'numeric constant' }] };
      if (r.operation === 'evaluateVba') return evaluation.promise;
      throw Error('Unexpected operation');
    },
  });
  await click('Inspect VBA source');
  await act(async () => {
    const select = document.querySelector('[aria-label="VBA module"]');
    select.value = 'EModule';
    select.dispatchEvent(new window.Event('change', { bubbles: true }));
  });
  await click('Evaluate pure numeric function');
  assert(button('Export new values-only file…').disabled);
  await render({ scope: 'F:base', request: async () => page('FModule') });
  await act(async () => evaluation.resolve({ functionName: 'OldFunction', value: 42, steps: 1 }));
  assert(!document.body.textContent.includes('OldFunction'));
  assert.equal(button('Export new values-only file…').disabled, false);
} finally {
  await act(async () => root.unmount());
  dom.window.close();
  delete globalThis.window;
  delete globalThis.document;
  delete globalThis.IS_REACT_ACT_ENVIRONMENT;
  await fs.rm(out, { recursive: true, force: true });
}
