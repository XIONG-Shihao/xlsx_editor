import './support/scratch-dir.mjs';
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { JSDOM } from 'jsdom';

// This test compiles only the new component. It does not bundle a product or
// launch a desktop/browser process; node's DOM exercises actual React handlers.
const directory = path.resolve('.build/tests/object-component-tests');
await fs.mkdir(directory, { recursive: true });
const modulePath = path.join(directory, 'WorkbookObjects.mjs');
await build({
  entryPoints: ['ui/WorkbookObjects.tsx'],
  bundle: true,
  write: true,
  outfile: modulePath,
  format: 'esm',
  platform: 'node',
  packages: 'external',
  jsx: 'automatic',
  logLevel: 'silent',
});
const { WorkbookObjects, WorkbookObjectPreviewView } = await import(pathToFileURL(modulePath));
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
const id = '11111111-1111-4111-8111-111111111111',
  entry = {
    id,
    kind: 'chart',
    name: 'Sales trend',
    revision: 'a'.repeat(64),
    managed: true,
    canEdit: true,
    canDelete: true,
    anchor: { cell: 'E2' },
    chart: { kind: 'column', range: 'A1:C4', title: 'Sales', legend: 'bottom' },
  };
const empty = { objects: [], total: 0, nextCursor: null },
  page = { objects: [entry], total: 1, nextCursor: null };
function deferred() {
  let resolve;
  const promise = new Promise((accept) => {
    resolve = accept;
  });
  return { promise, resolve };
}
const base = () => ({
  documentId: 'doc1',
  revision: 'rev1',
  sheet: 'Sales',
  selection: 'A1:C4',
  objects: async () => page,
  object: async () => entry,
  onBatch: async () => {},
  pickPng: async () => null,
});
const button = (text) => Array.from(document.querySelectorAll('button')).find((node) => node.textContent === text);
async function mount(t, props) {
  const root = createRoot(document.getElementById('root'));
  await act(async () => {
    root.render(React.createElement(WorkbookObjects, props));
  });
  t.after(async () => {
    await act(async () => root.unmount());
  });
  return { render: async (next) => act(async () => root.render(React.createElement(WorkbookObjects, next))) };
}
async function click(text) {
  assert.ok(button(text), 'missing button ' + text);
  await act(async () => {
    button(text).click();
  });
}

test('chart preview has accessible data, hidden-row policy, gaps and finite geometry across extreme numbers', () => {
  for (const kind of ['column', 'bar', 'line']) {
    const object = {
        ...entry,
        chart: { ...entry.chart, kind },
        data: {
          categories: ['Maximum', 'Minimum', 'Blank', 'Hidden'],
          series: [{ name: 'Values', values: [1e308, -1e308, null, 4] }],
          hidden: [false, false, false, true],
        },
      },
      html = renderToStaticMarkup(React.createElement(WorkbookObjectPreviewView, { object }));
    assert.match(html, /role="img" aria-label="Sales"/);
    assert.match(html, /<caption>Sales<\/caption>/);
    assert.match(html, /Hidden \(hidden\)/);
    assert.doesNotMatch(html, /(?:NaN|Infinity)/);
    assert.match(html, /<th scope="col">Values<\/th>/);
    const svg = new JSDOM(html).window.document.querySelector('svg');
    for (const node of svg.querySelectorAll('rect,line,circle,polyline'))
      for (const attribute of node.attributes)
        if (['x', 'y', 'y1', 'y2', 'cx', 'cy', 'width', 'height', 'points'].includes(attribute.name))
          assert.ok(
            attribute.value.split(/[ ,]+/).every((value) => Number.isFinite(Number(value))),
            attribute.name + '=' + attribute.value
          );
  }
});
test('replacement immediately clears old inventory and ignores delayed old list and preview results', async (t) => {
  const oldList = deferred(),
    oldPreview = deferred(),
    props = { ...base(), objects: async () => oldList.promise, object: async () => oldPreview.promise },
    ui = await mount(t, props);
  await ui.render({ ...props, documentId: 'doc2', revision: 'rev2', objects: async () => empty });
  await act(async () => oldList.resolve(page));
  assert.ok(!document.body.textContent.includes('Sales trend'));
  const current = { ...base(), documentId: 'doc2', revision: 'rev3', object: async () => oldPreview.promise };
  await ui.render(current);
  await click('Sales trend (chart)');
  await ui.render({ ...current, documentId: 'doc3', revision: 'rev4', objects: async () => empty });
  await act(async () =>
    oldPreview.resolve({
      ...entry,
      data: { categories: ['Old document'], series: [{ name: 'Values', values: [2] }], hidden: [false] },
    })
  );
  assert.ok(!document.body.textContent.includes('Old document'));
  assert.ok(!document.body.textContent.includes('Sales trend'));
});
test('delayed PNG picker cannot write after owner/revision replacement and normal insert uses explicit placement', async (t) => {
  const picked = deferred(),
    writes = [],
    asset = { mime: 'image/png', base64: 'fixture' },
    props = {
      ...base(),
      objects: async () => empty,
      pickPng: async () => picked.promise,
      onBatch: async (commands) => writes.push(commands),
    },
    ui = await mount(t, props);
  assert.equal(document.querySelector('[aria-label="Image anchor cell"]').value, 'E2');
  await click('Insert PNG');
  await ui.render({ ...props, documentId: 'doc2', revision: 'rev2' });
  await act(async () => picked.resolve(asset));
  assert.equal(writes.length, 0);
  await ui.render({ ...props, documentId: 'doc2', revision: 'rev3', pickPng: async () => asset });
  await click('Insert PNG');
  assert.equal(writes.length, 1);
  assert.equal(writes[0][0].sheet, 'Sales');
  assert.deepEqual(writes[0][0].anchor, { cell: 'E2' });
  assert.equal(writes[0][0].expectedObjectRevision, null);
});
test('chart create/edit/delete emits one scoped command with the accepted object guard', async (t) => {
  const writes = [],
    props = { ...base(), onBatch: async (commands) => writes.push(commands) };
  await mount(t, props);
  await click('Create chart');
  await act(async () =>
    document.querySelector('form').dispatchEvent(new window.Event('submit', { bubbles: true, cancelable: true }))
  );
  assert.equal(writes[0].length, 1);
  assert.equal(writes[0][0].operation, 'createChart');
  assert.equal(writes[0][0].chart.range, 'A1:C4');
  assert.equal(writes[0][0].expectedObjectRevision, null);
  await click('Sales trend (chart)');
  await click('Edit chart');
  await act(async () =>
    document.querySelector('form').dispatchEvent(new window.Event('submit', { bubbles: true, cancelable: true }))
  );
  assert.equal(writes[1][0].operation, 'updateChart');
  assert.equal(writes[1][0].expectedObjectRevision, entry.revision);
  assert.equal(writes[1][0].objectId, id);
  await click('Delete object');
  assert.deepEqual(writes[2], [
    { operation: 'deleteObject', sheet: 'Sales', objectId: id, expectedObjectRevision: entry.revision },
  ]);
});
test('image placement form retains prior guard, computes proportional height and excludes asset bytes', async (t) => {
  const image = {
      id: '22222222-2222-4222-8222-222222222222',
      kind: 'image',
      name: 'Logo',
      revision: 'b'.repeat(64),
      managed: true,
      canEdit: true,
      canDelete: true,
      anchor: { cell: 'F3', widthEmu: 3810000, heightEmu: 1905000 },
      altText: 'Logo',
    },
    writes = [],
    props = {
      ...base(),
      objects: async () => ({ objects: [image], total: 1, nextCursor: null }),
      object: async () => ({ ...image, image: { mime: 'image/png', base64: 'fixture', widthPx: 2, heightPx: 1 } }),
      onBatch: async (commands) => writes.push(commands),
    };
  await mount(t, props);
  await click('Logo (image)');
  assert.equal(button('Edit chart'), undefined);
  await click('Edit image placement');
  assert.equal(document.querySelector('[aria-label="Image width"]').value, '400');
  await act(async () =>
    document.querySelector('form').dispatchEvent(new window.Event('submit', { bubbles: true, cancelable: true }))
  );
  assert.deepEqual(writes, [
    [
      {
        operation: 'updateImage',
        sheet: 'Sales',
        objectId: image.id,
        expectedObjectRevision: image.revision,
        anchor: { cell: 'F3', widthEmu: 3810000, heightEmu: 1905000 },
        altText: 'Logo',
      },
    ],
  ]);
});
