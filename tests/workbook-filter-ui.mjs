import './support/scratch-dir.mjs';
import { rangeBounds } from '../core/workbook-commands.mjs';
/** Actual current Filter controls/helpers in React/JSDOM, controlled host only. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
const component = fileURLToPath(new URL('../', import.meta.url)),
  output = await fs.mkdtemp(path.join(component, '.build/tests/filter-ui-')),
  checks = [];
const entry = path.join(output, 'entry.ts'),
  bundle = path.join(output, 'test.mjs');
await fs.writeFile(
  entry,
  `export {SpreadsheetGrid} from ${JSON.stringify(path.join(component, 'ui/SpreadsheetGrid.tsx'))};export {configureXlsxHost} from ${JSON.stringify(path.join(component, 'ui/host.tsx'))};export * from ${JSON.stringify(path.join(component, 'ui/filterSelection.ts'))};`
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
  HTMLDetailsElement: dom.window.HTMLDetailsElement,
  Node: dom.window.Node,
  IS_REACT_ACT_ENVIRONMENT: true,
});
Object.defineProperties(window.HTMLElement.prototype, {
  clientWidth: {
    get() {
      return 1600;
    },
  },
  clientHeight: {
    get() {
      return 600;
    },
  },
});
globalThis.ResizeObserver = class {
  constructor(fn) {
    this.fn = fn;
  }
  observe() {
    this.fn([{ contentRect: { width: 1600, height: 600 } }]);
  }
  disconnect() {}
};
globalThis.requestAnimationFrame = (fn) => setTimeout(fn, 0);
globalThis.cancelAnimationFrame = clearTimeout;
const { createRoot } = await import('react-dom/client'),
  ui = await import(pathToFileURL(bundle));
ui.configureXlsxHost({
  api: { copyText: async () => {}, readClipboardText: async () => '' },
  Icon: () => React.createElement('svg'),
  useI18n: () => ({ locale: 'en' }),
});
const check = (name, fn) => {
  test(name, { timeout: 20000 }, async (t) => {
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
const props = (node) => node[Object.keys(node).find((key) => key.startsWith('__reactProps'))],
  field = (name) =>
    [...document.querySelectorAll('[aria-label]')].find((node) => node.getAttribute('aria-label') === name) ?? null,
  button = (name) =>
    [...document.querySelectorAll('button')].find((n) => n.textContent === name && !n.closest('[hidden]'));
const defer = () => {
  let resolve, reject;
  const promise = new Promise((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
};
const flush = async (ms = 8) => act(async () => new Promise((resolve) => setTimeout(resolve, ms)));
async function click(n) {
  n = typeof n === 'string' ? button(n) : n;
  assert(n);
  assert.equal(n.disabled, false);
  await act(async () => n.click());
  await flush();
  return n;
}
async function change(n, value) {
  assert(n);
  await act(async () => props(n).onChange({ target: { value, checked: value } }));
  await flush();
}
async function go(range) {
  await change(field('Go to cell'), range);
  await act(async () =>
    field('Go to cell').dispatchEvent(
      new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })
    )
  );
  await flush();
}
const revision = 'a'.repeat(64),
  nextRevision = 'b'.repeat(64),
  definition = (patch = {}) => ({
    range: 'A1:B6',
    column: 'A',
    values: ['West', 1, false, null],
    owned: true,
    revision,
    canReapply: true,
    canClear: true,
    ...patch,
  });
const page = (patch = {}) => ({
  values: [
    { key: 'text-west', value: 'West', count: 2 },
    { key: 'number-zero', value: 0, count: 1 },
    { key: 'boolean-false', value: false, count: 1 },
    { key: 'blank', value: null, count: 1 },
    { key: 'error-looking-text', value: '#N/A', count: 1 },
  ],
  total: 5,
  nextOffset: null,
  protected: { errors: 0, unresolved: 0, special: 0, unsafeNumbers: 0 },
  ...patch,
});
async function mount(t, patch = {}) {
  const root = createRoot(document.getElementById('root')),
    accepted = [],
    legacy = [],
    reads = [],
    data = {
      A1: { value: 'Region' },
      B1: { value: 'Amount' },
      A2: { value: 'West' },
      B2: { value: 10 },
      A3: { value: 0 },
      B3: { value: 20 },
      A4: { value: false },
      B4: { value: 30 },
      A5: { value: null },
      B5: { value: 5 },
      A6: { value: '#N/A' },
      B6: { value: 9 },
    };
  const sheet = (name) => ({
    name,
    path: 'xl/worksheets/' + name + '.xml',
    state: 'visible',
    rows: 40,
    columns: 18,
    cells: structuredClone(data),
    merges: [],
    widths: [],
    hiddenRows: [],
    rowHeights: [],
    freeze: { rows: 0, columns: 0 },
    filter: name === 'Sheet1' ? definition() : undefined,
  });
  let input = {
    sheets: [sheet('Sheet1'), sheet('Other')],
    documentId: 'doc',
    ownerIdentity: 'owner',
    storageKey: 'head',
    active: true,
    readOnly: false,
    onEdit: (value) => legacy.push(value),
    onBatch: (value) => legacy.push(value),
    onCommandBatch: async (commands) => accepted.push(commands),
    readRange: async (_sheet, range) => {
      const area = rangeBounds(range);
      return Object.fromEntries(
        Object.entries(input.sheets.find((s) => s.name === _sheet)?.cells ?? {}).filter(([ref]) => {
          const cell = rangeBounds(ref).first;
          return (
            cell.row >= area.first.row &&
            cell.row <= area.last.row &&
            cell.col >= area.first.col &&
            cell.col <= area.last.col
          );
        })
      );
    },
    filterValues: async (...args) => {
      reads.push(args);
      return page();
    },
    ...patch,
  };
  const render = async (patch) => {
    input = { ...input, ...patch };
    await act(async () => root.render(React.createElement(ui.SpreadsheetGrid, input)));
    await flush();
  };
  await render();
  await click([...document.querySelectorAll('[role="tab"]')].find((n) => n.textContent === 'Data'));
  let closed = false;
  const close = async () => {
    if (!closed) {
      closed = true;
      await act(async () => root.unmount());
    }
  };
  t.after(close);
  return {
    accepted,
    legacy,
    reads,
    data,
    render,
    close,
    input: () => input,
    setDefinition: async (value) =>
      render({ sheets: input.sheets.map((s) => (s.name === 'Sheet1' ? { ...s, filter: value } : s)) }),
  };
}
check(
  'accepted metadata forwards exact revision-only commands, independent Clear capability and truthful unavailable reasons',
  () => {
    assert.deepEqual(ui.acceptedFilterCommand('S', definition(), 'reapply'), {
      operation: 'reapplyFilter',
      sheet: 'S',
      expectedFilterRevision: revision,
    });
    assert.deepEqual(ui.acceptedFilterCommand('S', definition(), 'clear'), {
      operation: 'clearFilter',
      sheet: 'S',
      expectedFilterRevision: revision,
    });
    const protectedFilter = definition({ canReapply: false, canClear: true, reason: 'A key is pending calculation.' });
    assert.equal(ui.acceptedFilterReason(protectedFilter, 'reapply'), 'A key is pending calculation.');
    assert.equal(ui.acceptedFilterReason(protectedFilter, 'clear'), undefined);
    assert.equal(ui.acceptedFilterReason(undefined, 'clear'), 'This worksheet has no current filter.');
    for (const invalid of [
      definition({ owned: false, revision: null, reason: 'Imported criteria are protected.' }),
      definition({ revision: 'bad' }),
      definition({ values: [Infinity] }),
      definition({ extra: true }),
      definition({ canClear: 'yes' }),
    ])
      assert.throws(() => ui.acceptedFilterCommand('S', invalid, 'clear'));
    assert.notEqual(
      ui.filterDefinitionKey(definition({ values: [-0] })),
      ui.filterDefinitionKey(definition({ values: [0] }))
    );
    assert.doesNotThrow(() => ui.filterDefinitionKey(definition({ values: 'malformed' })));
  }
);
check(
  'typed pages preserve zero/false/blank and error-looking text; malformed outer/inner pages never become a clean preview',
  () => {
    const result = ui.validateFilterValuePage(page(), 0);
    assert.equal(result.values[1].value, 0);
    assert.equal(result.values[2].value, false);
    assert.equal(result.values[3].value, null);
    assert.equal(result.values[4].value, '#N/A');
    for (const bad of [
      null,
      [],
      page({ extra: 1 }),
      page({ total: '5' }),
      page({ nextOffset: 0 }),
      page({ protected: { errors: 0, unresolved: 0, special: 0 } }),
      page({ values: [null] }),
      page({ values: [{ key: 'x', count: 1, value: Infinity }] }),
      page({
        values: [
          { key: 'x', count: 1, value: 2 },
          { key: 'x', count: 1, value: 3 },
        ],
      }),
      page({ values: [{ key: 'x', count: 1, value: 'ok', extra: true }] }),
    ])
      assert.throws(() => ui.validateFilterValuePage(bad, 0));
  }
);
check(
  'actual Reapply/Clear use the current accepted range/revision even with unrelated selection and edited choice form',
  async (t) => {
    const h = await mount(t);
    await go('D20');
    await change(field('Search filter values'), 'unrelated typed query');
    assert.equal(button('Reapply filter').disabled, false);
    assert.match(field('Current filter').textContent, /Sheet1!A1:B6.*Column A/);
    assert.match(field('Current filter').textContent, /false \(boolean\)/);
    assert.equal(field('Current filter').dataset.filterRevision, revision);
    await click('Reapply filter');
    assert.deepEqual(h.accepted.pop(), [
      { operation: 'reapplyFilter', sheet: 'Sheet1', expectedFilterRevision: revision },
    ]);
    await h.render({ storageKey: 'reapplied-head' });
    await click('Clear filter');
    assert.deepEqual(h.accepted.pop(), [
      { operation: 'clearFilter', sheet: 'Sheet1', expectedFilterRevision: revision },
    ]);
    await h.setDefinition(
      definition({ values: [], canReapply: false, canClear: false, reason: 'No active filter criteria' })
    );
    assert(button('Reapply filter').disabled);
    assert.equal(button('Clear filter').title, 'No active filter criteria');
    assert.match(field('Current filter').textContent, /0 selected values/);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'one strict Apply/Reapply acknowledgement is awaited, duplicate physical writes are blocked and normal accepted head advance stays successful',
  async (t) => {
    const gate = defer(),
      h = await mount(t, {
        onCommandBatch: (commands) => {
          h.accepted.push(commands);
          return gate.promise;
        },
      }),
      submit = props(button('Reapply filter')).onClick;
    await act(async () => {
      submit();
      submit();
    });
    await flush();
    assert.equal(h.accepted.length, 1);
    assert(button('Reapply filter').disabled);
    assert(button('Clear filter').disabled);
    await h.render({ storageKey: 'accepted-head' });
    await act(async () => gate.resolve());
    await flush();
    assert.equal(field('Filter status').textContent, 'Filter reapplied.');
    await act(async () => submit());
    await flush();
    assert.equal(h.accepted.length, 1);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'definition-only replacement and old action handlers cannot park a new draft or write with an obsolete definition revision',
  async (t) => {
    const h = await mount(t),
      old = props(button('Reapply filter')).onClick;
    await h.setDefinition(definition({ revision: nextRevision, values: ['East'] }));
    const draft = field('Formula or value');
    await act(async () => {
      draft.focus();
      props(draft).onChange({ target: { value: '=SUM(' } });
    });
    draft.setSelectionRange(5, 5);
    await act(async () => old());
    await flush();
    assert.equal(draft.value, '=SUM(');
    assert.equal(draft.selectionStart, 5);
    assert.equal(document.activeElement, draft);
    assert.deepEqual(h.accepted, []);
    assert.equal(document.querySelector('.sheet-formula-bar [role="alert"]'), null);
    assert.equal(field('Current filter').dataset.filterRevision, nextRevision);
    assert.match(field('Current filter').textContent, /East/);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'paged values capture the original query/header body and submit exact selected primitive criteria in one command',
  async (t) => {
    const h = await mount(t);
    await go('A1:B6');
    await change(field('Search filter values'), 'we');
    await click('Load filter values');
    assert.deepEqual(h.reads, [['Sheet1', 'A2:B6', 'A', 'we', 0]]);
    await change(field('Filter value 0 (number)'), true);
    await change(field('Filter value false (boolean)'), true);
    await click('Apply selected values');
    assert.deepEqual(h.accepted, [
      [{ operation: 'filterRows', sheet: 'Sheet1', range: 'A1:B6', column: 'A', values: [0, false] }],
    ]);
    assert.equal(field('Filter status').textContent, 'Filter applied.');
    assert.equal(field('Filter value 0 (number)'), null);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'held value replies retire on owner/head/definition/sheet/selection roundtrip/activation and cannot paint or call a replacement callback',
  async (t) => {
    for (const kind of ['owner', 'head', 'definition', 'sheet', 'selection', 'active', 'mode']) {
      const gate = defer(),
        calls = [],
        h = await mount(t, {
          filterValues: async (...args) => {
            calls.push(args);
            return gate.promise;
          },
        });
      await go('A1:B6');
      await click('Load filter values');
      assert.equal(calls.length, 1);
      await h.render({
        filterValues: async () => {
          throw Error('Retargeted current reader');
        },
      });
      if (kind === 'owner') await h.render({ ownerIdentity: 'replacement' });
      else if (kind === 'head') await h.render({ storageKey: 'head-new' });
      else if (kind === 'definition') await h.setDefinition(definition({ revision: nextRevision }));
      else if (kind === 'sheet') {
        await click(button('Other'));
        await click(button('Sheet1'));
      } else if (kind === 'selection') {
        await go('C1:D6');
        await go('A1:B6');
      } else if (kind === 'active') {
        await h.render({ active: false });
        await h.render({ active: true });
      } else {
        const settings = {
            area: 'A1:B6',
            paper: 'Letter',
            orientation: 'portrait',
            marginsPt: { top: 36, right: 36, bottom: 36, left: 36, header: 18, footer: 18 },
            scaling: { mode: 'scale', percent: 100 },
            repeatRows: null,
            repeatColumns: null,
            gridlines: false,
          },
          printed = {
            number: 1,
            rowIndices: [0],
            columnIndices: [0],
            cells: [{ ref: 'A1', text: 'Region', style: { fontSize: 11 }, bbox: [36, 36, 100, 20] }],
            objects: [],
          };
        await h.render({
          printControls: {
            documentId: 'print-owner',
            revision: 'head',
            definition: async () => ({
              sheet: 'Sheet1',
              revision: null,
              settings,
              canEdit: true,
              protectedReasons: [],
            }),
            preview: async (_sheet, _settings, _id, geometryProfile) => ({
              geometryProfile,
              id: 'accepted-pages',
              state: 'head',
              units: 'pt',
              pageSizePt: [612, 792],
              scale: 1,
              sourceFingerprint: 'a'.repeat(64),
              revision: 'source',
              sheet: 'Sheet1',
              settings,
              pageCount: 1,
              pages: [printed],
            }),
            page: async () => printed,
            onBatch: async () => {},
            publish: async () => {},
            release: async () => {},
            cancel: async () => ({ granted: true }),
          },
        });
        await click([...document.querySelectorAll('[role="tab"]')].find((n) => n.textContent === 'View'));
        await click(document.querySelector('[data-ribbon-command="view.page-layout"]'));
        assert(document.querySelector('.sheet-page-layout'));
      }
      await act(async () =>
        gate.resolve(page({ values: [{ key: 'old', value: 'Obsolete reply', count: 1 }], total: 1 }))
      );
      await flush();
      assert.equal(field('Filter value "Obsolete reply" (text)'), null, kind);
      assert.deepEqual(h.accepted, [], kind);
      assert.equal(calls.length, 1);
      assert(!document.body.textContent.includes('Retargeted current reader'));
      await h.close();
    }
  }
);
check(
  'Cancel/query retirement keeps one physical read until settlement and old page Apply cannot use a new query form',
  async (t) => {
    const gate = defer(),
      calls = [],
      h = await mount(t, {
        filterValues: async (...args) => {
          calls.push(args);
          return calls.length === 1 ? gate.promise : page();
        },
      });
    await go('A1:B6');
    await click('Load filter values');
    assert(button('Load filter values').disabled);
    await click('Cancel value loading');
    assert(button('Load filter values').disabled);
    await act(async () => gate.resolve(page()));
    await flush();
    assert.equal(field('Filter value 0 (number)'), null);
    await click('Load filter values');
    assert.equal(calls.length, 2);
    await change(field('Filter value false (boolean)'), true);
    const oldApply = props(button('Apply selected values')).onClick;
    await change(field('Search filter values'), 'new query');
    await act(async () => oldApply());
    await flush();
    assert.deepEqual(h.accepted, []);
    assert.equal(field('Filter value false (boolean)'), null);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'late rejection masks obsolete errors; readonly/imported/absent/pending capabilities retain precise independent action reasons',
  async (t) => {
    const gate = defer(),
      h = await mount(t, {
        onCommandBatch: (commands) => {
          h.accepted.push(commands);
          return gate.promise;
        },
      });
    await click('Clear filter');
    await h.render({ ownerIdentity: 'new-owner', documentId: 'new-doc' });
    await act(async () => gate.reject(Error('Obsolete filter error')));
    await flush();
    assert(!document.body.textContent.includes('Obsolete filter error'));
    await h.render({ onCommandBatch: async () => {}, readOnly: true });
    assert(button('Clear filter').disabled);
    assert.equal(button('Clear filter').title, 'This workbook is read-only.');
    await h.render({ readOnly: false });
    await h.setDefinition(definition({ canReapply: false, canClear: true, reason: 'A key is pending calculation.' }));
    assert(button('Reapply filter').disabled);
    assert.equal(button('Reapply filter').title, 'A key is pending calculation.');
    assert.equal(button('Clear filter').disabled, false);
    await h.setDefinition({
      range: 'A1:B6',
      owned: false,
      revision: null,
      canReapply: false,
      canClear: false,
      reason: 'Imported criteria are protected.',
    });
    assert(button('Clear filter').disabled);
    assert.equal(button('Clear filter').title, 'Imported criteria are protected.');
    await h.setDefinition(undefined);
    assert.equal(button('Reapply filter').title, 'This worksheet has no current filter.');
    for (const patch of [
      { active: false },
      { active: true, busy: true },
      { busy: false, onCommandBatch: undefined },
      { onCommandBatch: async () => {}, ownerIdentity: undefined },
      { ownerIdentity: 'new-owner', passiveNotice: 'Passive inspector' },
    ]) {
      await h.render(patch);
      assert(button('Reapply filter').disabled);
      assert(button('Reapply filter').title);
    }
    await h.render({ passiveNotice: undefined, literal: true });
    assert.equal(button('Reapply filter'), undefined);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'filter menus/search/options/Cancel remain pure and preserve parked drafts while Sort remains one dock; explicit protected pages cannot Apply',
  async (t) => {
    const h = await mount(t, {
      filterValues: async () => page({ protected: { errors: 1, unresolved: 0, special: 0, unsafeNumbers: 0 } }),
    });
    await go('A1:B6');
    const draft = field('Formula or value');
    await act(async () => {
      draft.focus();
      props(draft).onChange({ target: { value: '=SUM(' } });
    });
    draft.setSelectionRange(5, 5);
    await change(field('Search filter values'), 'West');
    await click('Load filter values');
    assert(field('Filter value "West" (text)').disabled);
    assert(button('Apply selected values').disabled);
    assert.equal(draft.value, '=SUM(');
    assert.equal(draft.selectionStart, 5);
    await act(async () => draft.focus());
    await click('Sort…');
    assert.equal(document.querySelectorAll('.sheet-task-pane').length, 1);
    await click(field('Close workbook tools'));
    assert.equal(draft.value, '=SUM(');
    assert.equal(draft.selectionStart, 5);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
  }
);

check(
  'accepted worksheet Reapply/Clear remain usable under a genuine special selection while new Apply/Load keep rectangular guards',
  async (t) => {
    const h = await mount(t);
    await go('A1:B6');
    await click([...document.querySelectorAll('[role="tab"]')].find((n) => n.textContent === 'Home'));
    await click(document.querySelector('[data-home-command="home.select-blanks"]'));
    await click('Preview matching cells');
    await click('Select matching cells');
    assert(document.querySelector('.sheet-special-status'));
    assert(document.querySelector('.sheet-special-status'));
    await click([...document.querySelectorAll('[role="tab"]')].find((n) => n.textContent === 'Data'));
    assert.match(field('Current filter').textContent, /Sheet1!A1:B6/);
    assert.equal(button('Reapply filter').disabled, false);
    assert.equal(button('Clear filter').disabled, false);
    assert(button('Load filter values').disabled);
    await click('Reapply filter');
    await click('Clear filter');
    assert.deepEqual(h.accepted, [
      [{ operation: 'reapplyFilter', sheet: 'Sheet1', expectedFilterRevision: revision }],
      [{ operation: 'clearFilter', sheet: 'Sheet1', expectedFilterRevision: revision }],
    ]);
    assert.deepEqual(h.legacy, []);
  }
);
