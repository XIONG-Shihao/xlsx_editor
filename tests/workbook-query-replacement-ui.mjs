import './support/scratch-dir.mjs';
/** Actual query replacement pane, controlled host receipts; no native/app claim. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { randomUUID } from 'node:crypto';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
const component = fileURLToPath(new URL('../', import.meta.url)),
  output = await fs.mkdtemp(path.join(component, '.build/tests/query-replacement-ui-')),
  checks = [];
const entry = path.join(output, 'entry.ts'),
  bundle = path.join(output, 'test.mjs');
await fs.writeFile(
  entry,
  `export {WorkbookQueries} from ${JSON.stringify(path.join(component, 'ui/WorkbookQueries.tsx'))}`
);
await build({
  entryPoints: [entry],
  bundle: true,
  format: 'esm',
  platform: 'browser',
  jsx: 'automatic',
  outfile: bundle,
  external: ['react', 'react/jsx-runtime', 'react-dom', 'react-dom/client'],
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
  { WorkbookQueries } = await import(pathToFileURL(bundle));
const check = (name, fn) => {
  test(name, { timeout: 10000 }, async (t) => {
    try {
      await fn(t);
    } catch (error) {
      await fs.writeFile(path.join(output, 'failure-' + checks.length + '.html'), document.documentElement.outerHTML);
      throw error;
    }
  });
};
after(async () => {
  dom.window.close();
});
const props = (n) => n[Object.keys(n).find((key) => key.startsWith('__reactProps'))],
  field = (label) => [...document.querySelectorAll('[aria-label]')].find((n) => n.getAttribute('aria-label') === label),
  button = (name) => [...document.querySelectorAll('button')].find((n) => n.textContent === name),
  flush = async () => act(async () => new Promise((r) => setTimeout(r, 4))),
  defer = () => {
    let resolve;
    const promise = new Promise((r) => (resolve = r));
    return { promise, resolve };
  };
async function change(node, value) {
  assert(node);
  await act(async () => props(node).onChange({ target: { value, checked: value } }));
  await flush();
}
async function click(name) {
  const node = typeof name === 'string' ? button(name) : name;
  assert(node);
  assert.equal(node.disabled, false);
  await act(async () => node.click());
  await flush();
}
function fixture(patch = {}) {
  const ids = [randomUUID(), randomUUID()],
    records = ids.map((id, i) => ({
      id,
      revision: 1,
      name: 'Q' + i,
      target: { sheet: 'S', start: 'B3' },
      query: { columns: [1], filters: [], keepHeader: false, policies: ['number'] },
      source: { id: randomUUID(), name: 'source' + i + '.csv' },
    })),
    calls = [],
    review = randomUUID();
  const preview = (id = ids[0], offset = 0) => ({
    connectionId: id,
    revision: records.find((r) => r.id === id).revision,
    source: { id: 'source', name: 'source.csv', sha256: 'a'.repeat(64), bytes: 100 },
    outputRows: 105,
    outputColumns: 1,
    rows: [['1'], ['2']],
    truncated: true,
    target: { sheet: 'S', start: 'B3' },
    conversionInspection: {
      schema: 1,
      revision: review,
      date1904: false,
      policies: ['number'],
      rejected: 0,
      changed: 0,
      totalIssues: 0,
      offset: 0,
      nextOffset: null,
      issues: [],
      expiresAt: Date.now() + 600000,
    },
    replacementInspection: {
      schema: 1,
      revision: review,
      occupied: 105,
      blocked: 0,
      totalNewPositions: 105,
      offset,
      nextOffset: offset === 0 ? 100 : null,
      rows: Array.from({ length: Math.min(100, 105 - offset) }, (_, i) => ({
        index: offset + i,
        sourceRef: 'A' + (offset + i + 1),
        destinationRef: 'B' + (offset + i + 3),
        sourceText: String(offset + i + 1),
        outputKind: 'number',
        outputDisplay: String(offset + i + 1),
        current:
          i === 0 && offset === 0
            ? {
                kind: 'formula',
                display: '=SUM(Z1)',
                formula: 'SUM(Z1)',
                rawValue: '913',
                format: 0,
                pending: true,
                storedDisplay: 'Stored last result: 912',
              }
            : {
                kind: 'text',
                display: JSON.stringify('Original ' + (offset + i)),
                rawValue: 'Original ' + (offset + i),
                formula: null,
                format: 0,
              },
      })),
    },
  });
  const p = {
    documentId: 'doc',
    ownerId: 'owner',
    revision: 'head',
    sheet: 'S',
    selection: 'A1',
    list: async () => ({ entries: structuredClone(records) }),
    create: async () => ({ cancelled: true }),
    update: async (id, revision, draft) => {
      const i = records.findIndex((r) => r.id === id);
      records[i] = { ...records[i], ...draft, revision: revision + 1 };
      return { connection: structuredClone(records[i]) };
    },
    remove: async () => ({ removed: true }),
    preview: async (id, revision, requestId, options) => {
      calls.push({ kind: 'preview', id, revision, options });
      return { queryPreview: preview(id, options?.replacementOffset ?? 0) };
    },
    refresh: async (id, revision, requestId, options) => {
      calls.push({ kind: 'refresh', id, revision, options });
      return { refreshed: true, outputRows: 105, outputColumns: 1 };
    },
    cancel: async () => ({ outcome: 'cancelled' }),
    ...patch,
  };
  return { props: p, calls, ids, records, review, preview };
}
async function mount(t, f) {
  const root = createRoot(document.getElementById('root'));
  let p = f.props;
  await act(async () => root.render(React.createElement(WorkbookQueries, p)));
  await flush();
  t.after(async () => act(async () => root.unmount()));
  await change(field('Saved query'), f.ids[0]);
  return {
    render: async (next) => {
      p = { ...p, ...next };
      await act(async () => root.render(React.createElement(WorkbookQueries, p)));
      await flush();
    },
  };
}
check(
  'explicit replacement action exposes all paged originals and exact authored/pending formula before one acknowledged refresh',
  async (t) => {
    const f = fixture();
    await mount(t, f);
    await click('Preview query');
    assert(button('Refresh query').disabled);
    assert.equal(field('Query replacement rows').querySelectorAll('tbody tr').length, 100);
    assert.match(field('Query replacement B3').textContent, /formula · =SUM\(Z1\)/);
    assert.match(field('Query replacement B3').textContent, /Pending calculation. Stored last result: 912/);
    await click('More replacements');
    assert.equal(field('Query replacement rows').querySelectorAll('tbody tr').length, 5);
    assert.match(field('Query replacement B107').textContent, /Original 104/);
    assert.equal(f.calls[1].options.replacementOffset, 100);
    assert.equal(f.calls[1].options.expectedConversionRevision, f.review);
    await click('Refresh and replace 105 occupied cells');
    const call = f.calls.find((c) => c.kind === 'refresh');
    assert.deepEqual(call.options, { expectedConversionRevision: f.review, expectedReplacementRevision: f.review });
    assert.equal(f.calls.filter((c) => c.kind === 'refresh').length, 1);
  }
);
check(
  'retained replace handlers cannot dispatch after type,target or saved-source edits and do not overwrite new settings',
  async (t) => {
    const f = fixture();
    await mount(t, f);
    for (const edit of [
      async () => change(field('Output field 1 type'), 'text'),
      async () => change(field('Query start cell'), 'G8'),
      async () => change(field('Saved query'), f.ids[1]),
    ]) {
      if (button('Preview query').disabled) await click('Save query settings');
      await click('Preview query');
      const old = props(button('Refresh and replace 105 occupied cells')).onClick;
      await edit();
      await act(async () => old());
      await flush();
      assert.equal(f.calls.filter((c) => c.kind === 'refresh').length, 0);
      assert.equal(field('Query destination replacements'), undefined);
    }
    assert.equal(field('Saved query').value, f.ids[1]);
  }
);
check(
  'rejections and known special targets block entire explicit replacement, including a direct stale DOM invocation',
  async (t) => {
    const f = fixture({
      preview: async (id) => {
        const p = f.preview(id);
        p.replacementInspection.blocked = 1;
        p.replacementInspection.rows[0].current = {
          kind: 'protected',
          display: 'Stored array',
          rawValue: '913',
          formula: 'SUM(Z1)',
          format: 0,
          reason: 'Protected array',
        };
        return { queryPreview: p };
      },
    });
    await mount(t, f);
    await click('Preview query');
    assert(button('Refresh and replace 105 occupied cells').disabled);
    await act(async () => props(button('Refresh and replace 105 occupied cells')).onClick());
    assert.equal(f.calls.filter((c) => c.kind === 'refresh').length, 0);
    assert.match(field('Query destination replacements').textContent, /protected destination cells block/);
  }
);
check(
  'one physical replacement awaits same-owner durable result; owner retirement masks pages and late receipts',
  async (t) => {
    const gate = defer(),
      f = fixture({
        refresh: async (...args) => {
          f.calls.push({ kind: 'refresh', args });
          return gate.promise;
        },
      });
    const m = await mount(t, f);
    await click('Preview query');
    const old = props(button('Refresh and replace 105 occupied cells')).onClick;
    await act(async () => {
      old();
      old();
    });
    await flush();
    assert.equal(f.calls.filter((c) => c.kind === 'refresh').length, 1);
    await m.render({ ownerId: 'replacement', revision: 'head' });
    await act(async () => gate.resolve({ refreshed: true, outputRows: 105, outputColumns: 1 }));
    await flush();
    assert.equal(field('Query destination replacements'), undefined);
    assert(!document.body.textContent.includes('Refreshed 105 rows'));
  }
);
