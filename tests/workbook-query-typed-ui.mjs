import './support/scratch-dir.mjs';
/** Actual source query pane; controlled host Promises, not native/app proof. */
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
  output = await fs.mkdtemp(path.join(component, '.build/tests/typed-query-ui-')),
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
    let resolve, reject;
    const promise = new Promise((a, b) => {
      resolve = a;
      reject = b;
    });
    return { promise, resolve, reject };
  };
async function change(n, value) {
  assert(n);
  await act(async () => props(n).onChange({ target: { value, checked: value } }));
  await flush();
}
async function click(name) {
  const n = typeof name === 'string' ? button(name) : name;
  assert(n);
  assert.equal(n.disabled, false);
  await act(async () => n.click());
  await flush();
}
function fixture(patch = {}) {
  let record = {
    id: randomUUID(),
    revision: 1,
    name: 'Imported data',
    target: { sheet: 'S', start: 'B3' },
    query: { columns: [1, 2, 3, 4], filters: [], keepHeader: true },
    source: { id: randomUUID(), name: 'data.csv' },
  };
  const calls = [],
    revision = randomUUID(),
    preview = (offset = 0) => ({
      connectionId: record.id,
      revision: record.revision,
      source: { id: record.source.id, name: 'data.csv', bytes: 20, sha256: 'a'.repeat(64) },
      target: record.target,
      outputRows: 206,
      outputColumns: 12,
      rows: [
        ['id', 'n', 'flag', 'day'],
        ['00123', '-0', 'TRUE', '2011-07-05'],
      ],
      truncated: true,
      conversionInspection: {
        schema: 1,
        revision,
        date1904: false,
        policies: Array.from({ length: 12 }, (_, i) => record.query.policies?.[i] ?? 'text'),
        rejected: 0,
        changed: 205,
        totalIssues: 205,
        offset,
        nextOffset: offset === 0 ? 100 : offset === 100 ? 200 : null,
        issues: Array.from({ length: Math.min(100, 205 - offset) }, (_, n) => ({
          index: offset + n,
          sourceRow: offset + n + 2,
          sourceColumn: 2,
          outputRow: offset + n + 2,
          outputColumn: 2,
          sourceRef: 'B' + (offset + n + 2),
          destinationRef: 'C' + (offset + n + 4),
          raw: '-0',
          policy: 'number',
          status: 'converted',
          kind: 'number',
          display: '-0',
          notices: ['Numeric spelling changes.'],
        })),
        expiresAt: Date.now() + 600000,
      },
    });
  let config = {
    documentId: 'doc',
    ownerId: 'owner',
    revision: 'head',
    sheet: 'S',
    selection: 'A1',
    list: async () => ({ entries: [structuredClone(record)] }),
    create: async () => {
      throw Error('Unexpected create');
    },
    update: async (id, prior, draft) => {
      calls.push({ kind: 'update', id, prior, draft });
      assert.equal(prior, record.revision);
      record = { ...record, ...draft, revision: prior + 1 };
      return { connection: structuredClone(record) };
    },
    remove: async () => ({ removed: true }),
    preview: async (id, prior, requestId, options) => {
      calls.push({ kind: 'preview', id, prior, requestId, options });
      return { queryPreview: preview(options?.offset ?? 0) };
    },
    refresh: async (id, prior, requestId, options) => {
      calls.push({ kind: 'refresh', id, prior, requestId, options });
      return { refreshed: true, outputRows: 206, outputColumns: 4 };
    },
    cancel: async () => ({ outcome: 'cancelled' }),
    ...patch,
  };
  return {
    calls,
    preview,
    props: config,
    record: () => record,
    typed: () => {
      record.query.policies = ['text', 'number', 'boolean', 'iso-date'];
    },
  };
}
async function mount(t, f) {
  const root = createRoot(document.getElementById('root'));
  let p = f.props;
  await act(async () => root.render(React.createElement(WorkbookQueries, p)));
  await flush();
  t.after(async () => act(async () => root.unmount()));
  const render = async (patch) => {
    p = { ...p, ...patch };
    await act(async () => root.render(React.createElement(WorkbookQueries, p)));
    await flush();
  };
  await change(field('Saved query'), f.record().id);
  return { render };
}
check(
  'legacy Text settings and refresh remain compatible; pane marks draft zone and never creates formula edits',
  async (t) => {
    const f = fixture();
    await mount(t, f);
    assert.equal(field('Output field 1 type').value, 'text');
    assert.equal(field('Queries and refresh').getAttribute('data-formula-draft-zone'), 'true');
    await click('Refresh query');
    assert.deepEqual(f.calls.filter((c) => c.kind === 'refresh')[0].options, {});
    assert.match(document.body.textContent, /Refreshed 206 rows/);
    assert.equal(f.calls.filter((c) => c.kind === 'update').length, 0);
  }
);
check(
  'policy edits retire preview but keep other field controls; saving exact output order requires a fresh typed review',
  async (t) => {
    const f = fixture();
    await mount(t, f);
    await click('Preview query');
    assert(field('Query preview'));
    await change(field('Output field 2 type'), 'number');
    assert.equal(field('Query preview'), undefined);
    assert(field('Output field 3 type'));
    await change(field('Output field 3 type'), 'boolean');
    await change(field('Output field 4 type'), 'iso-date');
    assert(button('Preview query').disabled);
    assert(button('Refresh query').disabled);
    await click('Save query settings');
    assert.deepEqual(f.record().query.policies, ['text', 'number', 'boolean', 'iso-date']);
    assert(button('Refresh query').disabled);
    await click('Preview query');
    await click('Refresh query');
    const call = f.calls.find((c) => c.kind === 'refresh');
    assert.equal(call.options.expectedConversionRevision, f.preview().conversionInspection.revision);
    assert.equal(f.calls.filter((c) => c.kind === 'refresh').length, 1);
  }
);
check(
  'all paged issues are accessible with original/source refs, signed display and exact reviewed token',
  async (t) => {
    const f = fixture();
    f.typed();
    await mount(t, f);
    await click('Preview query');
    assert.equal(field('Query conversion issues').querySelectorAll('tbody tr').length, 100);
    assert.match(field('Query conversion issue C4').textContent, /-0/);
    await click('Next conversion issues');
    assert(field('Query conversion issue C104'));
    await click('Next conversion issues');
    assert.equal(field('Query conversion issues').querySelectorAll('tbody tr').length, 5);
    assert(field('Query conversion issue C208'));
    assert.equal(button('Next conversion issues'), undefined);
    const pages = f.calls.filter((c) => c.kind === 'preview');
    assert.deepEqual(
      pages.map((c) => c.options.offset),
      [undefined, 100, 200]
    );
    assert.equal(pages[1].options.expectedConversionRevision, pages[2].options.expectedConversionRevision);
    await click('Previous conversion issues');
    assert(field('Query conversion issue C104'));
  }
);
check(
  'rejected conversions disable whole Refresh and remain visible beyond first text preview; edits mask receipt immediately',
  async (t) => {
    const f = fixture();
    f.typed();
    f.props.preview = async () => {
      const p = f.preview();
      p.conversionInspection.rejected = 103;
      p.conversionInspection.issues[0] = {
        ...p.conversionInspection.issues[0],
        status: 'rejected',
        kind: undefined,
        display: undefined,
        reason: 'Integer is outside the exact safe-integer range.',
      };
      return { queryPreview: p };
    };
    await mount(t, f);
    await click('Preview query');
    assert(button('Refresh query').disabled);
    assert.match(field('Query conversion issue C4').textContent, /safe-integer/);
    const old = props(button('Refresh query')).onClick;
    await act(async () => old());
    assert.equal(f.calls.filter((c) => c.kind === 'refresh').length, 0);
    await change(field('Output field 2 type'), 'text');
    assert.equal(field('Query conversion inspection'), undefined);
    assert(field('Output field 3 type'));
  }
);
check(
  'held preview masks on head or runtime replacement; old handlers and cancelled replies cannot adopt',
  async (t) => {
    const gate = defer(),
      f = fixture({ preview: () => gate.promise });
    const m = await mount(t, f);
    await click('Preview query');
    const oldPolicy = props(field('Output field 2 type')).onChange;
    await m.render({ ownerId: 'other', revision: 'head' });
    await act(async () => oldPolicy({ target: { value: 'number' } }));
    await act(async () => gate.resolve({ queryPreview: f.preview() }));
    await flush();
    assert.equal(field('Query preview'), undefined);
    assert(!document.body.textContent.includes('Refreshed'));
    const next = fixture(),
      gate2 = defer();
    next.props.preview = () => gate2.promise;
    await m.render(next.props);
    await change(field('Saved query'), next.record().id);
    await click('Preview query');
    await click('Cancel query');
    await act(async () => gate2.resolve({ queryPreview: next.preview() }));
    await flush();
    assert.equal(field('Query preview'), undefined);
  }
);
check(
  'one physical Refresh survives same-owner accepted head, finishing cancellation and warning; duplicate is fenced',
  async (t) => {
    const gate = defer(),
      f = fixture({
        refresh: async (...args) => {
          f.calls.push({ kind: 'refresh', args });
          return gate.promise;
        },
        cancel: async () => ({ outcome: 'finishing' }),
      });
    f.typed();
    const m = await mount(t, f);
    await click('Preview query');
    const handler = props(button('Refresh query')).onClick;
    await act(async () => {
      handler();
      handler();
    });
    await flush();
    assert.equal(f.calls.filter((c) => c.kind === 'refresh').length, 1);
    await click('Cancel query');
    assert(button('Cancel query').disabled);
    await m.render({ revision: 'accepted-head' });
    await act(async () =>
      gate.resolve({ refreshed: true, warning: 'Data accepted; metadata update remains pending.' })
    );
    await flush();
    assert.match(document.body.textContent, /Data accepted; metadata/);
    assert.equal(button('Cancel query'), undefined);
  }
);
check(
  'late Refresh/error on replacement owner is invisible and read-only settings/output remain blocked',
  async (t) => {
    const gate = defer(),
      f = fixture({ refresh: () => gate.promise });
    const m = await mount(t, f);
    await click('Refresh query');
    await m.render({ ownerId: 'new-owner' });
    await act(async () => gate.reject(Error('Old owner error')));
    await flush();
    assert(!document.body.textContent.includes('Old owner error'));
    await m.render({ readOnly: true });
    assert(button('Refresh query').disabled);
    assert(button('Preview query').disabled);
    assert(field('Output field 2 type')?.disabled ?? true);
  }
);
check(
  'discovered all-column fields stay available after type1 edit/save and field paging; configuration retires issue approval',
  async (t) => {
    const f = fixture();
    f.record().query.columns = [];
    await mount(t, f);
    await click('Preview query');
    await change(field('Output field 1 type'), 'number');
    assert(field('Output field 2 type'));
    await click('More output fields');
    assert(field('Output field 12 type'));
    await change(field('Output field 12 type'), 'boolean');
    await click('Save query settings');
    await click('More output fields');
    assert(field('Output field 12 type'));
    assert.equal(f.record().query.policies.length, 12);
    await click('Preview query');
    await change(field('Keep first row as header'), false);
    assert.equal(field('Query conversion inspection'), undefined);
    assert(button('Refresh query').disabled);
    assert.equal(f.calls.filter((c) => c.kind === 'refresh').length, 0);
  }
);
