import './support/scratch-dir.mjs';
/** Source React controls under controlled captured callbacks; app proof is separate. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
const folder = await fs.mkdtemp(path.resolve('.build/tests/refresh-all-ui-')),
  file = path.join(folder, 'ui.mjs');
await build({
  stdin: {
    contents: 'export {WorkbookRefreshAll} from "./ui/WorkbookRefreshAll";',
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
  { WorkbookRefreshAll } = await import(pathToFileURL(file));
const flush = async () => {
  for (let n = 0; n < 3; n++) await act(async () => new Promise((r) => setTimeout(r, 0)));
};
after(() => dom.window.close());
const ids = ['11111111-1111-4111-8111-111111111111', '22222222-2222-4222-8222-222222222222'];
const scope = () => ({ documentId: 'a'.repeat(64), ownerId: ids[0], state: 'base' });
const preview = () => ({
  schema: 1,
  id: ids[1],
  scope: scope(),
  expiresAt: Date.now() + 60000,
  undoPolicy: 'One Undo step for each accepted connection',
  entries: [
    {
      connectionId: ids[0],
      revision: 1,
      name: 'Ready data',
      target: { sheet: 'S', start: 'A1' },
      status: 'ready',
      sample: [['00123', '=1+2']],
    },
    {
      connectionId: ids[1],
      revision: 1,
      name: 'Refused data',
      target: { sheet: 'S', start: 'D1' },
      status: 'refused',
      reason: 'Source unavailable.',
    },
  ],
});
const button = (name) => [...document.querySelectorAll('button')].find((b) => b.textContent === name),
  gate = () => {
    let resolve;
    const promise = new Promise((r) => (resolve = r));
    return { promise, resolve };
  };
async function mount(t, patch = {}) {
  const root = createRoot(document.getElementById('root')),
    calls = [];
  let args = {
    frame: {
      key: 'root/base',
      ownerKey: 'root/active',
      name: 'Budget.xlsx',
      active: true,
      controls: {
        scope: scope(),
        preview: async (id) => {
          calls.push(['preview', id]);
          return { refreshAllPreview: preview() };
        },
        refresh: async (id, requestId) => {
          calls.push(['refresh', id, requestId]);
          return {
            refreshAllResult: {
              schema: 1,
              previewId: id,
              scope: scope(),
              lastAcknowledgedState: ids[0],
              cancelled: false,
              stopped: false,
              accepted: 1,
              uncertain: 0,
              undoUnits: 1,
              outcomes: [
                {
                  connectionId: ids[0],
                  name: 'Ready data',
                  target: { sheet: 'S', start: 'A1' },
                  status: 'refreshed',
                  parentState: 'base',
                  state: ids[0],
                  undoUnits: 1,
                },
                {
                  connectionId: ids[1],
                  name: 'Refused data',
                  target: { sheet: 'S', start: 'D1' },
                  status: 'refused',
                  reason: 'Source unavailable.',
                  undoUnits: 0,
                },
              ],
            },
          };
        },
        cancel: async (id) => {
          calls.push(['cancel', id]);
          return { outcome: 'cancelled' };
        },
      },
    },
    intent: 1,
    ...patch,
  };
  const render = async (p) => {
    args = { ...args, ...p };
    await act(async () => root.render(React.createElement(WorkbookRefreshAll, args)));
    await flush();
  };
  await render({});
  t.after(async () => act(async () => root.unmount()));
  return {
    calls,
    get args() {
      return args;
    },
    render,
  };
}
test('opening lists every ready/refused unit, captured workbook and bounded literal samples; explicit Run shows per-unit outcomes', async (t) => {
  const h = await mount(t);
  assert(document.body.textContent.includes('Budget.xlsx'));
  assert(document.body.textContent.includes('2 saved local connections'));
  assert(document.body.textContent.includes('Source unavailable.'));
  assert(document.body.textContent.includes('00123'));
  assert.equal(h.calls.filter((c) => c[0] === 'refresh').length, 0);
  await act(async () => button('Refresh ready connections').click());
  await flush();
  assert(document.body.textContent.includes('1 accepted connections'));
  assert(document.body.textContent.includes('Ready data · Refreshed'));
  assert(document.body.textContent.includes('Refused data · Refused'));
  assert.equal(h.calls.filter((c) => c[0] === 'refresh').length, 1);
  assert(button('Refresh ready connections').disabled);
});
test('prepared head retirement requires a fresh Preview; cancellation discards a ready token without a write', async (t) => {
  const h = await mount(t);
  await h.render({
    frame: {
      ...h.args.frame,
      key: 'new-head',
      controls: { ...h.args.frame.controls, scope: { ...scope(), state: ids[0] } },
    },
  });
  assert(button('Refresh ready connections').disabled);
  assert(document.body.textContent.includes('Preview Refresh All again'));
  await h.render({
    frame: { ...h.args.frame, key: 'root/base', controls: { ...h.args.frame.controls, scope: scope() } },
  });
  await act(async () => button('Preview Refresh All').click());
  await flush();
  await act(async () => button('Cancel Refresh All preview').click());
  await flush();
  assert(button('Refresh ready connections').disabled);
  assert.equal(h.calls.filter((c) => c[0] === 'refresh').length, 0);
});
test('own busy rendering preserves preview while blocking Run; closing/replacement masks a late accepted result', async (t) => {
  const h = await mount(t);
  await h.render({ frame: { ...h.args.frame, busy: true } });
  assert(button('Refresh ready connections').disabled);
  assert(document.body.textContent.includes('2 saved local connections'));
  await h.render({ frame: { ...h.args.frame, busy: false } });
  const held = gate(),
    original = h.args.frame.controls.refresh;
  await h.render({ frame: { ...h.args.frame, controls: { ...h.args.frame.controls, refresh: () => held.promise } } });
  await act(async () => button('Refresh ready connections').click());
  await h.render({ frame: { ...h.args.frame, active: false, ownerKey: 'closed' } });
  await act(async () => held.resolve(await original(ids[1], ids[0])));
  await flush();
  assert.equal(document.querySelector('[aria-label="Refresh All outcomes"]'), null);
  assert(document.querySelector('section').hidden);
});
test('readonly unavailable scopes never call Preview/Run; cancellation targets a pending original request', async (t) => {
  const h = await mount(t, {
    frame: {
      key: 'readonly',
      ownerKey: 'readonly',
      name: 'Readonly',
      active: true,
      reason: 'Read-only',
      controls: {
        scope: scope(),
        preview: async () => {
          throw Error('must not call');
        },
        refresh: async () => {
          throw Error('must not call');
        },
        cancel: async () => ({ outcome: 'cancelled' }),
      },
    },
  });
  assert(button('Preview Refresh All').disabled);
  assert(button('Refresh ready connections').disabled);
  assert(document.body.textContent.includes('Read-only'));
  const pending = gate(),
    calls = [];
  await h.render({
    frame: {
      key: 'root/base',
      ownerKey: 'root/active',
      name: 'Budget.xlsx',
      active: true,
      controls: {
        scope: scope(),
        preview: async (id) => {
          calls.push(id);
          return pending.promise;
        },
        refresh: async () => {
          throw Error('must not call');
        },
        cancel: async (id) => {
          calls.push(id);
          return { outcome: 'cancelled' };
        },
      },
    },
    intent: 2,
  });
  await act(async () => button('Cancel remaining refreshes').click());
  assert.equal(calls[0], calls[1]);
  await act(async () => pending.resolve({ refreshAllPreview: preview() }));
  await flush();
  assert(button('Refresh ready connections').disabled);
});
