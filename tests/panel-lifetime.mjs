import './support/scratch-dir.mjs';
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { transform } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';

// Only compile the two actual child panels. No desktop, owner or native build.
const directory = path.resolve(
  '.build/tests/taskpane-lifetime-component'
);
await fs.mkdir(directory, { recursive: true });
const panels = {};
for (const name of ['WorkbookQueries', 'WorkbookReferences']) {
  const file = path.join(directory, name + '.mjs'),
    compiled = await transform(await fs.readFile('ui/' + name + '.tsx', 'utf8'), {
      loader: 'tsx',
      jsx: 'automatic',
      format: 'esm',
      target: 'es2022',
    });
  await fs.writeFile(file, compiled.code);
  panels[name] = (await import(pathToFileURL(file)))[name];
}
const dom = new JSDOM('<!doctype html><div id="root"></div>', { url: 'http://localhost/' });
globalThis.window = dom.window;
globalThis.document = dom.window.document;
globalThis.HTMLElement = dom.window.HTMLElement;
globalThis.IS_REACT_ACT_ENVIRONMENT = true;
const { createRoot } = await import('react-dom/client');
after(() => {
  dom.window.close();
  delete globalThis.window;
  delete globalThis.document;
  delete globalThis.HTMLElement;
  delete globalThis.IS_REACT_ACT_ENVIRONMENT;
});
function deferred() {
  let resolve, reject;
  const promise = new Promise((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
}
const connection = {
  id: 'query1',
  revision: 1,
  name: 'Saved source',
  source: { name: 'source.csv' },
  target: { sheet: 'Sheet1', start: 'A1' },
  query: { columns: [1], filters: [], keepHeader: true },
};
const base = () => ({
  documentId: 'doc1',
  revision: 'r1',
  sheet: 'Sheet1',
  selection: 'A1:B2',
  list: async () => ({ entries: [connection] }),
  create: async () => ({ cancelled: true }),
  update: async () => ({ connection }),
  remove: async () => ({ removed: true }),
  preview: async () => ({
    queryPreview: {
      connectionId: connection.id,
      rows: [['Old preview']],
      outputRows: 1,
      outputColumns: 1,
      source: connection.source,
      target: connection.target,
    },
  }),
  refresh: async () => ({ refreshed: true, outputRows: 1, outputColumns: 1 }),
  cancel: async () => ({ granted: true, outcome: 'cancelled' }),
});
const button = (text) => [...document.querySelectorAll('button')].find((node) => node.textContent === text);
async function click(text) {
  const node = button(text);
  assert.ok(node, 'Missing ' + text);
  assert.equal(node.disabled, false, text + ' disabled');
  await act(async () => node.click());
}
async function choose() {
  const select = document.querySelector('[aria-label="Saved query"]');
  await act(async () => {
    select.value = connection.id;
    select.dispatchEvent(new window.Event('change', { bubbles: true }));
  });
}
async function mount(t, Panel, props) {
  const root = createRoot(document.getElementById('root'));
  let unmounted = false;
  const render = async (next, hidden = false) =>
    act(async () => root.render(React.createElement('div', { hidden }, React.createElement(Panel, next))));
  await render(props);
  const unmount = async () => {
    if (!unmounted) {
      unmounted = true;
      await act(async () => root.unmount());
    }
  };
  t.after(unmount);
  return { render, unmount };
}

test('pending refresh keeps its cancellation request across an accepted revision and hidden dock', async (t) => {
  const result = deferred(),
    cancelled = [],
    requests = [],
    props = {
      ...base(),
      refresh: async (_id, _revision, requestId) => {
        requests.push(requestId);
        return result.promise;
      },
      cancel: async (requestId) => {
        cancelled.push(requestId);
        return { granted: false, reason: 'finishing' };
      },
    },
    ui = await mount(t, panels.WorkbookQueries, props);
  await choose();
  await click('Refresh query');
  assert.equal(requests.length, 1);
  const input = document.querySelector('[aria-label="Query name"]');
  await ui.render({ ...props, revision: 'r2' }, true);
  assert.equal(document.querySelector('[aria-label="Query name"]'), input);
  assert.equal(input.value, 'Saved source');
  assert.ok(button('Cancel query'), 'Accepted revision must not discard the pending request');
  await ui.render({ ...props, revision: 'r2' });
  await click('Cancel query');
  assert.deepEqual(cancelled, requests);
  assert.equal(button('Cancel query').disabled, true);
  assert.match(document.body.textContent, /finishing its accepted recovery write/);
  await act(async () => result.resolve({ refreshed: true, outputRows: 3, outputColumns: 2 }));
  assert.equal(button('Cancel query'), undefined);
  assert.match(document.body.textContent, /Refreshed 3 rows and 2 columns/);
  assert.doesNotMatch(document.body.textContent, /Refresh cancelled/);
});
test('document replacement resets saved query and draft; an old create response cannot list or select into the new owner', async (t) => {
  const created = deferred();
  let oldLists = 0,
    newLists = 0;
  const props = {
      ...base(),
      list: async () => {
        oldLists++;
        return { entries: [connection] };
      },
      create: async () => created.promise,
    },
    ui = await mount(t, panels.WorkbookQueries, props);
  await choose();
  await click('Choose CSV/TSV source…');
  const next = {
    ...base(),
    documentId: 'doc2',
    revision: 'r2',
    sheet: 'Other',
    selection: 'C7:D8',
    list: async () => {
      newLists++;
      return { entries: [] };
    },
  };
  await ui.render(next);
  assert.equal(document.querySelector('[aria-label="Saved query"]').value, '');
  assert.equal(document.querySelector('[aria-label="Query name"]').value, 'Local query');
  assert.equal(document.querySelector('[aria-label="Query worksheet"]').value, 'Other');
  assert.equal(document.querySelector('[aria-label="Query start cell"]').value, 'C7');
  await act(async () => created.resolve({ connection }));
  assert.equal(oldLists, 1);
  assert.equal(newLists, 1);
  assert.doesNotMatch(document.body.textContent, /source\.csv/);
});
test('unmounted query preview has no late list request or completion callback', async (t) => {
  const pending = deferred();
  let lists = 0;
  const props = {
      ...base(),
      list: async () => {
        lists++;
        return { entries: [connection] };
      },
      preview: async () => pending.promise,
    },
    ui = await mount(t, panels.WorkbookQueries, props);
  await choose();
  await click('Preview query');
  await ui.unmount();
  await act(async () =>
    pending.resolve({ queryPreview: { rows: [['Late']], outputRows: 1, outputColumns: 1, source: connection.source } })
  );
  assert.equal(lists, 1, 'Unmount must invalidate the continuation before it calls list');
  assert.equal(document.getElementById('root').textContent, '');
});
test('late cancellation response cannot report an old request into a replacement document', async (t) => {
  const refreshed = deferred(),
    cancellation = deferred(),
    cancelled = [],
    props = {
      ...base(),
      refresh: async () => refreshed.promise,
      cancel: async (id) => {
        cancelled.push(id);
        return cancellation.promise;
      },
    },
    ui = await mount(t, panels.WorkbookQueries, props);
  await choose();
  await click('Refresh query');
  await click('Cancel query');
  await ui.render({
    ...base(),
    documentId: 'doc2',
    revision: 'r2',
    list: async () => ({ entries: [] }),
    cancel: async () => {
      throw Error('New document cancel must not be used');
    },
  });
  await act(async () => cancellation.resolve({ granted: false, reason: 'finishing' }));
  assert.equal(cancelled.length, 1);
  assert.doesNotMatch(document.body.textContent, /Cancellation requested|finishing its accepted recovery write/);
  assert.equal(button('Cancel query'), undefined);
  await act(async () => refreshed.resolve({ refreshed: true, outputRows: 99, outputColumns: 9 }));
  assert.doesNotMatch(document.body.textContent, /Refreshed 99/);
});
test('denied completed cancellation stays pending and reports the actual cancelled refresh outcome', async (t) => {
  const refreshed = deferred(),
    props = {
      ...base(),
      refresh: async () => refreshed.promise,
      cancel: async () => ({ granted: false, reason: 'completed' }),
    };
  await mount(t, panels.WorkbookQueries, props);
  await choose();
  await click('Refresh query');
  await click('Cancel query');
  assert.ok(button('Cancel query'));
  assert.equal(button('Cancel query').disabled, true);
  assert.doesNotMatch(document.body.textContent, /Cancellation requested/);
  await act(async () => refreshed.resolve({ cancelled: true }));
  assert.match(document.body.textContent, /Refresh cancelled\. Accepted edits were kept\./);
  assert.equal(button('Cancel query'), undefined);
});
test('retained references keep their pending command and suppress old-sheet errors, then remain usable', async (t) => {
  const pending = deferred(),
    writes = [],
    props = {
      sheet: 'Sheet1',
      selection: 'A1:B3',
      revision: 'doc1:r1',
      tables: [{ name: 'Sales', sheet: 'Sheet1', ref: 'A1:B3', columns: [{ name: 'Region' }, { name: 'Amount' }] }],
      pivots: [],
      readOnly: false,
      busy: false,
      active: true,
      onBatch: async (commands) => {
        writes.push(commands);
        return pending.promise;
      },
    },
    ui = await mount(t, panels.WorkbookReferences, props);
  const table = document.querySelector('[aria-label="Table to edit"]');
  await act(async () => {
    table.value = 'Sales';
    table.dispatchEvent(new window.Event('change', { bubbles: true }));
  });
  await click('Use selected table range');
  // Create-pivot is a separate owner command; no synthetic text-input event is needed.
  await click('Create pivot');
  await ui.render(props, true);
  assert.equal(writes.length, 1);
  assert.match(document.body.textContent, /Updating workbook definitions/);
  await ui.render({ ...props, sheet: 'Other', selection: 'C1:D3', revision: 'doc1:r2' });
  await act(async () => pending.reject(Error('Old Sheet1 error')));
  assert.doesNotMatch(document.body.textContent, /Old Sheet1 error/);
  assert.equal(button('Create pivot').disabled, false);
  assert.equal(document.querySelector('[aria-label="Table to edit"]').value, '');
});

const referenceProps = (onBatch = async () => {}) => ({
  sheet: 'Sheet1',
  selection: 'A1:B3',
  revision: 'doc1:r1',
  tables: [],
  pivots: [],
  readOnly: false,
  busy: false,
  active: true,
  onBatch,
});
const handler = (node) => node[Object.keys(node).find((k) => k.startsWith('__reactProps'))].onClick;
test('reference commands use one physical write lane even before React renders disabled state', async (t) => {
  const pending = deferred(),
    writes = [],
    p = referenceProps(async (c) => {
      writes.push(c);
      return pending.promise;
    });
  await mount(t, panels.WorkbookReferences, p);
  const send = handler(button('Create pivot'));
  await act(async () => {
    send();
    send();
  });
  assert.equal(writes.length, 1);
  await act(async () => pending.resolve());
});
test('reference command handlers retire across selection, permission and visibility round trips', async (t) => {
  const writes = [],
    p = referenceProps(async (c) => writes.push(c)),
    ui = await mount(t, panels.WorkbookReferences, p);
  for (const change of [{ selection: 'C1:D3' }, { active: false }, { readOnly: true }, { busy: true }]) {
    const send = handler(button('Create pivot'));
    await ui.render({ ...p, ...change });
    await act(async () => send());
    assert.equal(writes.length, 0);
    await ui.render(p);
    await act(async () => send());
    assert.equal(writes.length, 0, 'A retired handler cannot revive after a round trip');
  }
  await click('Create pivot');
  assert.equal(writes.length, 1);
});

test('reference controls stay usable after the development StrictMode effect rehearsal', async (t) => {
  const writes = [],
    p = referenceProps(async (c) => writes.push(c)),
    StrictPanel = (props) =>
      React.createElement(React.StrictMode, null, React.createElement(panels.WorkbookReferences, props));
  await mount(t, StrictPanel, p);
  await click('Create pivot');
  assert.equal(writes.length, 1);
});
