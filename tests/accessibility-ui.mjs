import './support/scratch-dir.mjs';
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
import { createAccessibilityController } from '../ui/accessibilityController.ts';
import { bindAccessibility, ACCESSIBILITY_RULES } from '../core/accessibility.mjs';
const out = await fs.mkdtemp(path.resolve('.build/tests/accessibility-ui-')),
  bundle = path.join(out, 'ui.mjs');
await build({
  stdin: {
    contents: 'export {WorkbookAccessibility} from "./ui/WorkbookAccessibility"',
    resolveDir: process.cwd(),
    loader: 'tsx',
  },
  bundle: true,
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
  IS_REACT_ACT_ENVIRONMENT: true,
});
const { createRoot } = await import('react-dom/client'),
  { WorkbookAccessibility } = await import(pathToFileURL(bundle));
after(() => dom.window.close());
const receipt = { documentId: 'doc', ownerId: 'owner', state: 'head', revision: 'revision' },
  blank = (cell = 'A1') => ({
    rule: 'blank-table-header',
    message: 'This header is blank.',
    sheet: 'Main',
    cell,
    objectId: null,
    tableName: 'Data',
    relatedCells: [],
  });
const raw = (patch = {}) => ({
    schema: 'workbook.xlsx.accessibility.v1',
    scope: 'workbook',
    status: 'complete',
    rules: [...ACCESSIBILITY_RULES],
    findings: [
      blank(),
      { ...blank('B1'), rule: 'duplicate-table-header', message: 'These headers repeat.', relatedCells: ['B1', 'C1'] },
      {
        ...blank('Z80'),
        rule: 'missing-alt-text',
        message: 'This object has no alternative text.',
        sheet: 'Details',
        objectId: 'object',
        tableName: null,
      },
    ],
    coverage: {
      objects: { total: 1, checked: 1, skipped: 0 },
      tables: { total: 1, checked: 1, skipped: 0 },
      headerCellsRead: 3,
      skipped: [],
    },
    ...patch,
  }),
  result = (patch = {}, binding = receipt) => bindAccessibility(raw(patch), binding),
  frame = (patch = {}) => ({
    key: 'owner/head/pane',
    receipt,
    active: true,
    visibleSheets: ['Main', 'Details'],
    read: async () => result(),
    navigate: () => {},
    ...patch,
  });
const props = (n) => Object.entries(n).find(([k]) => k.startsWith('__reactProps$'))[1],
  label = (s) => document.querySelector('[aria-label="' + s + '"]'),
  button = (s) => [...document.querySelectorAll('button')].find((n) => n.textContent === s),
  flush = async () => {
    for (let i = 0; i < 4; i++) await act(async () => new Promise((r) => setTimeout(r, 0)));
  },
  click = async (s) => {
    assert(button(s), s);
    await act(async () => button(s).click());
    await flush();
  },
  gate = () => {
    let resolve;
    const promise = new Promise((r) => (resolve = r));
    return { promise, resolve };
  };
async function mount(t, patch = {}) {
  const root = createRoot(document.getElementById('root')),
    moves = [];
  let value = frame({ navigate: (...v) => moves.push(v), ...patch });
  const render = async (p) => {
    value = { ...value, ...p };
    await act(async () => root.render(React.createElement(WorkbookAccessibility, { frame: value })));
    await flush();
  };
  await render({});
  t.after(async () => act(async () => root.unmount()));
  return { moves, render, frame: () => value };
}
test('readonly findings show exact rules/coverage and navigate cross-sheet and related header locations without a write surface', async (t) => {
  const ui = await mount(t);
  assert.match(document.body.textContent, /Objects checked: 1 \/ 1/);
  assert.match(document.body.textContent, /Tables checked: 1 \/ 1/);
  assert.match(document.body.textContent, /Missing alternative text/);
  assert.match(document.body.textContent, /Blank table header/);
  assert.equal(document.querySelector('form'), null);
  assert.equal(document.querySelector('input,textarea'), null);
  for (const name of ['Go to finding Main!A1', 'Go to related header Main!C1', 'Go to finding Details!Z80'])
    await act(async () => label(name).click());
  assert.deepEqual(ui.moves, [
    ['Main', 'A1'],
    ['Main', 'C1'],
    ['Details', 'Z80'],
  ]);
});
test('partial unknown coverage and skipped reasons never claim a clean workbook; current refresh clears prior findings', async (t) => {
  const partial = {
      findings: [],
      status: 'partial',
      coverage: {
        objects: { total: null, checked: 0, skipped: 0 },
        tables: { total: 1, checked: 0, skipped: 1 },
        headerCellsRead: 0,
        skipped: [
          { kind: 'objects', sheet: null, name: null, reason: 'Drawing inventory unavailable' },
          { kind: 'table', sheet: 'Main', name: 'Data', reason: 'Formula headers skipped' },
        ],
      },
    },
    ui = await mount(t, { read: async () => result(partial) });
  assert.match(document.body.textContent, /Some content was not checked/);
  assert.match(document.body.textContent, /Objects checked: 0 \/ unknown/);
  assert.match(document.body.textContent, /skipped content still needs review/);
  assert.match(document.body.textContent, /Formula headers skipped/);
  const hold = gate();
  await ui.render({ read: async () => hold.promise });
  await click('Check again');
  assert.match(document.body.textContent, /Checking workbook content/);
  assert.doesNotMatch(document.body.textContent, /Formula headers skipped/);
  await act(async () => hold.resolve(result()));
  await flush();
  assert.match(document.body.textContent, /Findings \(3\)/);
});
test('finding and skipped-item paging is bounded and retains original report indices for navigation', async (t) => {
  const findings = Array.from({ length: 70 }, (_, i) => blank('A' + (i + 1))),
    skipped = Array.from({ length: 70 }, (_, i) => ({
      kind: 'objects',
      sheet: 'Main',
      name: 'Object ' + i,
      reason: 'Unmanaged',
    })),
    ui = await mount(t, {
      read: async () =>
        result({
          findings,
          status: 'partial',
          coverage: {
            objects: { total: 70, checked: 0, skipped: 70 },
            tables: { total: 1, checked: 1, skipped: 0 },
            headerCellsRead: 70,
            skipped,
          },
        }),
    });
  assert.equal(document.querySelectorAll('ol>li').length, 50);
  assert.equal(label('Go to finding Main!A51'), null);
  assert.equal(document.querySelectorAll('details li').length, 50);
  await click('More findings');
  assert.equal(document.querySelectorAll('ol>li').length, 20);
  await act(async () => label('Go to finding Main!A51').click());
  assert.deepEqual(ui.moves, [['Main', 'A51']]);
  await click('More skipped items');
  assert.equal(document.querySelectorAll('details li').length, 20);
  await click('Previous findings');
  assert.equal(document.querySelectorAll('ol>li').length, 50);
});
test('hidden worksheet findings retain their locations/explanations with navigation disabled', async (t) => {
  const ui = await mount(t, { visibleSheets: ['Main'] });
  assert(label('Go to finding Details!Z80').disabled);
  assert.match(document.body.textContent, /hidden worksheet/);
  await act(async () => label('Go to finding Details!Z80').click());
  assert.deepEqual(ui.moves, []);
});
test('owner/head/pane replacement masks old data and joins the retired physical read before the latest frame', async () => {
  const controller = createAccessibilityController(),
    hold = gate(),
    reads = [],
    moves = [],
    old = frame({
      read: async () => {
        reads.push('old');
        return hold.promise;
      },
    }),
    nextReceipt = { ...receipt, state: 'new-head', revision: 'new-revision' },
    fresh = frame({
      key: 'fresh',
      receipt: nextReceipt,
      read: async () => {
        reads.push('fresh');
        return result({ findings: [blank('D4')] }, nextReceipt);
      },
      navigate: (...v) => moves.push(v),
    });
  controller.setFrame(old);
  controller.setFrame(fresh);
  assert.deepEqual(reads, ['old']);
  assert.equal(controller.getSnapshot().report, null);
  assert.throws(() => controller.navigate(old, 0), /changed/);
  hold.resolve(result());
  await flush();
  assert.deepEqual(reads, ['old', 'fresh']);
  assert.equal(controller.getSnapshot().report.findings[0].cell, 'D4');
  assert.throws(() => controller.navigate(old, 0), /changed/);
  assert.throws(() => controller.navigate(fresh, 0, 'E9'), /current finding/);
  controller.navigate(fresh, 0);
  assert.deepEqual(moves, [['Main', 'D4']]);
  controller.dispose();
});
test('wrong-owner replies are visible errors with no guessed-empty results, and explicit retry uses the live reader', async (t) => {
  const ui = await mount(t, { read: async () => result({}, { ...receipt, ownerId: 'other' }) });
  assert.match(document.body.textContent, /retired owner\/head/);
  assert.doesNotMatch(document.body.textContent, /No findings in checked content/);
  assert.equal(label('Go to finding Main!A1'), null);
  await ui.render({ read: async () => result() });
  await click('Check again');
  assert(label('Go to finding Main!A1'));
});
test('held old location and refresh callbacks cannot affect a closed or replacement pane', async (t) => {
  const ui = await mount(t),
    move = props(label('Go to finding Main!A1')).onClick,
    refresh = props(button('Check again')).onClick;
  await ui.render({ key: 'closed', active: false });
  await act(async () => {
    move();
    refresh();
  });
  await flush();
  assert.deepEqual(ui.moves, []);
  assert.equal(label('Go to finding Main!A1'), null);
  await ui.render({ key: 'reopened', active: true, read: async () => result({ findings: [blank('D4')] }) });
  await act(async () => move());
  await flush();
  assert.deepEqual(ui.moves, []);
  assert(label('Go to finding Main!D4'));
});
