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
  out = await fs.mkdtemp(path.join(component, '.build/tests/hyperlink-ui-')),
  bundle = path.join(out, 'ui.mjs');
await build({
  stdin: { contents: 'export {WorkbookHyperlink} from "./ui/WorkbookHyperlink"', resolveDir: component, loader: 'tsx' },
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
  { WorkbookHyperlink } = await import(pathToFileURL(bundle));
after(() => dom.window.close());
const receipt = { documentId: 'doc-A', ownerId: 'owner-A', state: 'head-1', revision: 'revision-1' },
  raw = {
    schema: 'workbook.xlsx.hyperlink-definition.v1',
    sheet: 'Main',
    cell: 'A1',
    hyperlinkRevision: null,
    link: null,
    tooltip: '',
    canEdit: true,
    protectedReason: null,
  };
const tick = () => new Promise((r) => setTimeout(r, 0)),
  flush = async () => {
    for (let i = 0; i < 3; i++) await act(async () => tick());
  },
  props = (n) => Object.entries(n).find(([k]) => k.startsWith('__reactProps$'))[1],
  label = (s) => document.querySelector('[aria-label="' + s + '"]'),
  button = (s) => [...document.querySelectorAll('button')].find((n) => n.textContent === s);
async function mount(t, patch = {}) {
  const root = createRoot(document.getElementById('root')),
    writes = [];
  let frame = {
    key: 'A/1/epoch1',
    family: 'A/cell/epoch1',
    sheet: 'Main',
    cell: 'A1',
    receipt,
    active: true,
    writable: true,
    read: async () => ({ ...raw, ...receipt }),
    apply: async (command) => writes.push(command),
    ...patch,
  };
  const render = async (value) => {
    frame = { ...frame, ...value };
    await act(async () => root.render(React.createElement(WorkbookHyperlink, { frame, sheets: ['Main', 'Dest'] })));
    await flush();
  };
  await render({});
  t.after(async () => act(async () => root.unmount()));
  return { writes, render };
}
async function change(node, value) {
  await act(async () => props(node).onChange({ target: { value } }));
  await flush();
}
async function click(node) {
  assert(node);
  await act(async () => node.dispatchEvent(new window.MouseEvent('click', { bubbles: true, cancelable: true })));
  await flush();
}
async function submit() {
  await act(async () =>
    document.querySelector('form').dispatchEvent(new window.Event('submit', { bubbles: true, cancelable: true }))
  );
  await flush();
}
test('local URL draft and Reset do not write; submit preserves cell identity and sends one command', async (t) => {
  const ui = await mount(t);
  await change(label('Web address'), 'https://example.test/a?x=1&y=2');
  await change(label('Screen tip'), 'Visit');
  assert.deepEqual(ui.writes, []);
  await click(button('Reset link form'));
  assert.equal(label('Web address').value, '');
  assert.deepEqual(ui.writes, []);
  await change(label('Web address'), 'https://example.test/');
  await submit();
  assert.deepEqual(ui.writes, [
    [
      {
        type: 'setHyperlink',
        sheet: 'Main',
        cell: 'A1',
        expectedHyperlinkRevision: null,
        link: { kind: 'url', target: 'https://example.test/' },
        tooltip: '',
      },
    ],
  ]);
});
test('internal draft uses actual worksheet names and target cell, not a display-only URL', async (t) => {
  const ui = await mount(t);
  await change(label('Link destination'), 'internal');
  await change(label('Target worksheet'), 'Dest');
  await change(label('Target cell'), 'B2');
  await submit();
  assert.deepEqual(ui.writes[0][0].link, { kind: 'internal', sheet: 'Dest', cell: 'B2' });
});
test('edit/remove use the displayed revision and readonly forms never call the writer', async (t) => {
  const ui = await mount(t, {
    read: async () => ({
      ...raw,
      ...receipt,
      hyperlinkRevision: 'a'.repeat(64),
      link: { kind: 'url', target: 'https://example.test/old' },
      tooltip: 'Old',
    }),
  });
  assert.equal(label('Web address').value, 'https://example.test/old');
  await click(button('Remove link'));
  assert.deepEqual(ui.writes, [
    [{ type: 'removeHyperlink', sheet: 'Main', cell: 'A1', expectedHyperlinkRevision: 'a'.repeat(64) }],
  ]);
  await ui.render({ writable: false, reason: 'Read-only' });
  assert.equal(document.querySelector('fieldset').disabled, true);
  await submit();
  assert.equal(ui.writes.length, 1);
});
test('owner/selection retirement discards typed draft and stale submit callback cannot accept', async (t) => {
  const ui = await mount(t);
  await change(label('Web address'), 'https://old.test/');
  const old = props(document.querySelector('form')).onSubmit;
  await ui.render({
    key: 'B/1/epoch2',
    family: 'B/cell/epoch2',
    sheet: 'Dest',
    cell: 'B2',
    receipt: { ...receipt, documentId: 'doc-B' },
    read: async () => ({ ...raw, ...receipt, documentId: 'doc-B', sheet: 'Dest', cell: 'B2' }),
  });
  assert.equal(label('Web address').value, '');
  await act(async () => old({ preventDefault() {} }));
  await flush();
  assert.deepEqual(ui.writes, []);
});
test('unsupported imported record explains the boundary without granting apply/remove', async (t) => {
  const ui = await mount(t, {
    read: async () => ({
      ...raw,
      ...receipt,
      hyperlinkRevision: 'a'.repeat(64),
      canEdit: false,
      protectedReason: 'Imported range link',
    }),
  });
  assert.match(document.body.textContent, /Imported range link/);
  assert.equal(document.querySelector('fieldset').disabled, true);
  assert.equal(button('Remove link').disabled, true);
  assert.deepEqual(ui.writes, []);
});
test('blank-cell label defaults to the address and plain-text labels are explicit captured edits', async (t) => {
  const ui = await mount(t, { readCell: async () => null });
  assert(label('Text to display'));
  await change(label('Web address'), 'https://example.test/new');
  await submit();
  assert.equal(ui.writes[0].length, 2);
  assert.deepEqual(ui.writes[0][1], {
    sheet: 'Main',
    cell: 'A1',
    value: 'https://example.test/new',
    expectedValue: null,
    expectedFormula: null,
  });
});
