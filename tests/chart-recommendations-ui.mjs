import './support/scratch-dir.mjs';
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
import { createChartRecommendationsController } from '../ui/chartRecommendationsController.ts';
const out = await fs.mkdtemp(path.resolve('.build/tests/chart-recommendations-ui-')),
  file = path.join(out, 'ui.mjs');
await build({
  stdin: {
    contents: 'export {WorkbookObjects} from "./ui/WorkbookObjects";',
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
const { WorkbookObjects } = await import(pathToFileURL(file)),
  dom = new JSDOM('<div id="root"></div>', { url: 'http://localhost' });
Object.assign(globalThis, {
  window: dom.window,
  document: dom.window.document,
  HTMLElement: dom.window.HTMLElement,
  Element: dom.window.Element,
  IS_REACT_ACT_ENVIRONMENT: true,
});
const { createRoot } = await import('react-dom/client');
const cells = {
    A1: { value: 'Category', dataType: 'text' },
    B1: { value: 'Sales', dataType: 'text' },
    A2: { value: 'Alpha', dataType: 'text' },
    A3: { value: 'Beta', dataType: 'text' },
    B2: { value: 2, dataType: 'number', rawValue: '2' },
    B3: { value: 5, dataType: 'number', rawValue: '5' },
  },
  deferred = () => {
    let resolve;
    const promise = new Promise((r) => (resolve = r));
    return { promise, resolve };
  },
  flush = async () => {
    for (let i = 0; i < 5; i++) await act(async () => new Promise((r) => setTimeout(r, 0)));
  },
  props = (n) => Object.entries(n).find(([k]) => k.startsWith('__reactProps$'))[1],
  button = (s) => [...document.querySelectorAll('button')].find((n) => n.textContent === s),
  click = async (s) => {
    assert(button(s), s);
    await act(async () => button(s).click());
    await flush();
  },
  input = async (label, value) => {
    const n = document.querySelector('[aria-label="' + label + '"]');
    assert(n);
    await act(async () => props(n).onChange({ target: { value } }));
    await flush();
  },
  submit = async () => {
    await act(async () =>
      props(document.querySelector('[aria-label="Recommended Charts"] form')).onSubmit({ preventDefault() {} })
    );
    await flush();
  };
const frame = (patch = {}) => ({
    key: 'owner/head/selection/pane',
    sheet: 'Data',
    range: 'A1:B3',
    active: true,
    view: { sheets: [{ name: 'Data' }] },
    read: async () => structuredClone(cells),
    apply: async () => {},
    ...patch,
  }),
  base = (f) => ({
    documentId: 'doc',
    ownerIdentity: 'owner',
    revision: 'head',
    sheet: 'Data',
    selection: 'A1:B3',
    mode: 'recommendations',
    recommendationFrame: f,
    objects: async () => ({ objects: [], total: 0, nextCursor: null }),
    object: async () => {},
    onBatch: async () => {},
  });
async function mount(t, more = {}) {
  const root = createRoot(document.getElementById('root')),
    writes = [],
    f = frame({ apply: async (c) => writes.push(c) });
  let args = { ...base(f), ...more };
  const render = async (p) => {
    args = { ...args, ...p };
    await act(async () => root.render(React.createElement(WorkbookObjects, args)));
    await flush();
  };
  await render({});
  t.after(async () => act(async () => root.unmount()));
  return {
    writes,
    render,
    get args() {
      return args;
    },
  };
}
after(() => dom.window.close());
test('three proposals show actual selected data and style choices create captured editable charts', async (t) => {
  const ui = await mount(t);
  assert.equal(document.querySelectorAll('[data-chart-proposal]').length, 3);
  assert.equal(document.querySelectorAll('[data-chart-proposal] svg').length, 3);
  assert.equal(document.querySelector('img,image'), null);
  for (const [label, kind, style] of [
    ['Compare categories', 'column', 'colorful'],
    ['Compare labels', 'bar', 'monochrome'],
    ['Follow row order', 'line', 'outline'],
  ]) {
    await click(label);
    await submit();
    const c = ui.writes.at(-1);
    assert.equal(c.operation, 'createChart');
    assert.equal(c.expectedObjectRevision, null);
    assert.equal(c.chart.kind, kind);
    assert.equal(c.chart.style, style);
    assert.equal(c.chart.range, 'A1:B3');
  }
  assert.equal(ui.writes.length, 3);
});
test('a style choice and anchor change only the selected draft before one accepted write', async (t) => {
  const ui = await mount(t);
  await input('Chart style', 'outline');
  await input('Recommended chart anchor', 'H7');
  assert.equal(ui.writes.length, 0);
  await submit();
  assert.equal(ui.writes.length, 1);
  assert.equal(ui.writes[0].chart.style, 'outline');
  assert.equal(ui.writes[0].anchor.cell, 'H7');
});
test('original owner/head/selection replacement retires the preview and its held submit callback', async (t) => {
  const ui = await mount(t);
  const old = props(document.querySelector('[aria-label="Recommended Charts"] form')).onSubmit;
  await ui.render({
    revision: 'other-head',
    recommendationFrame: {
      ...ui.args.recommendationFrame,
      key: 'other-head/range',
      range: 'B1:C3',
      read: async () => ({}),
    },
  });
  assert.equal(document.querySelectorAll('[data-chart-proposal]').length, 0);
  await act(async () => old({ preventDefault() {} }));
  assert.equal(ui.writes.length, 0);
  assert.match(document.body.textContent, /header|literal/);
});
test('readonly/busy and invalid source profiles cannot insert or silently substitute cached values', async (t) => {
  const f = frame({ reason: 'This workbook is read-only.' }),
    ui = await mount(t, { readOnly: true, recommendationFrame: f });
  await submit();
  assert.equal(ui.writes.length, 0);
  assert(document.querySelector('[aria-label="Recommended Charts"] fieldset').disabled);
  const bad = structuredClone(cells);
  bad.B2 = {
    value: 99,
    formula: '1+1',
    dataType: 'number',
    calculation: { status: 'calculated', pending: true, value: 2 },
  };
  await ui.render({ readOnly: false, recommendationFrame: frame({ key: 'pending-values', read: async () => bad }) });
  assert.match(document.body.textContent, /unresolved/);
  assert.equal(document.querySelectorAll('[data-chart-proposal]').length, 0);
});
test('physical write lane persists across recommendation-to-manual mode/head retirement', async (t) => {
  const gate = deferred(),
    writes = [],
    f = frame({
      apply: async (c) => {
        writes.push(c);
        await gate.promise;
      },
    }),
    ui = await mount(t, { recommendationFrame: f, onBatch: async (c) => writes.push(...c) });
  await act(async () =>
    props(document.querySelector('[aria-label="Recommended Charts"] form')).onSubmit({ preventDefault() {} })
  );
  await flush();
  assert.equal(writes.length, 1);
  await ui.render({
    mode: 'manage',
    revision: 'new-head',
    recommendationFrame: { ...f, key: 'retired', active: false },
  });
  assert(button('Create chart').disabled);
  await click('Create chart');
  assert.equal(writes.length, 1);
  await act(async () => gate.resolve());
  await flush();
  assert(!button('Create chart').disabled);
});
test('manual chart forms expose three real styles and retain chosen style in edit commands', async (t) => {
  const writes = [],
    entry = {
      id: 'chart',
      name: 'Saved',
      kind: 'chart',
      revision: 'r1',
      managed: true,
      canEdit: true,
      canDelete: true,
      anchor: { cell: 'E2' },
      chart: { kind: 'column', range: 'A1:B3', title: 'Saved', style: 'monochrome', legend: 'bottom' },
    },
    ui = await mount(t, {
      mode: 'manage',
      objects: async () => ({ objects: [entry], total: 1, nextCursor: null }),
      object: async () => entry,
      onBatch: async (c) => writes.push(...c),
    });
  await click('Create chart');
  await input('Chart style', 'outline');
  await act(async () => props(document.querySelector('form')).onSubmit({ preventDefault() {} }));
  await flush();
  assert.equal(writes[0].chart.style, 'outline');
  await click('Saved (chart)');
  await click('Edit chart');
  assert.equal(document.querySelector('[aria-label="Chart style"]').value, 'monochrome');
  await input('Chart style', 'colorful');
  await act(async () => props(document.querySelector('form')).onSubmit({ preventDefault() {} }));
  await flush();
  assert.equal(writes[1].patch.style, 'colorful');
  assert.equal(writes[1].expectedObjectRevision, 'r1');
});
test('one physical source read joins retirement, then reads only the latest frame; stale replies never grant Apply', async () => {
  const gate = deferred(),
    reads = [],
    controller = createChartRecommendationsController(),
    old = frame({
      read: async () => {
        reads.push('old');
        return gate.promise;
      },
    });
  controller.setFrame(old);
  controller.setFrame(
    frame({
      key: 'next',
      read: async () => {
        reads.push('next');
        return structuredClone(cells);
      },
    })
  );
  assert.deepEqual(reads, ['old']);
  assert.equal(controller.getSnapshot().result, null);
  gate.resolve(cells);
  await flush();
  assert.deepEqual(reads, ['old', 'next']);
  assert.equal(controller.getSnapshot().key, 'next');
  assert.throws(() => controller.capture(old, 'column-comparison', 'colorful'));
  controller.dispose();
});

test('legacy unstyled charts disclose the current default and can explicitly choose Colorful', async (t) => {
  const writes = [],
    entry = {
      id: 'legacy',
      name: 'Legacy',
      kind: 'chart',
      revision: 'r1',
      managed: true,
      canEdit: true,
      canDelete: true,
      anchor: { cell: 'E2' },
      chart: { kind: 'column', range: 'A1:B3', title: 'Legacy', legend: 'bottom' },
    };
  await mount(t, {
    mode: 'manage',
    objects: async () => ({ objects: [entry], total: 1, nextCursor: null }),
    object: async () => entry,
    onBatch: async (c) => writes.push(...c),
  });
  await click('Legacy (chart)');
  await click('Edit chart');
  assert.equal(document.querySelector('[aria-label="Chart style"]').value, '');
  assert.equal(document.querySelector('[aria-label="Chart style"] option:checked').textContent, 'Current default');
  await input('Chart style', 'colorful');
  await act(async () => props(document.querySelector('form')).onSubmit({ preventDefault() {} }));
  await flush();
  assert.equal(writes[0].patch.style, 'colorful');
});
