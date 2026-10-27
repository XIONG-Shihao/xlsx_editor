import './support/scratch-dir.mjs';
import { rangeBounds } from '../core/workbook-commands.mjs';
/** Actual sort Grid/pane in React/JSDOM with simulated owner reads/acknowledgements. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
const component = fileURLToPath(new URL('../', import.meta.url)),
  output = await fs.mkdtemp(path.join(component, '.build/tests/sort-ui-')),
  checks = [];
const entry = path.join(output, 'entry.ts'),
  bundle = path.join(output, 'test.mjs');
await fs.writeFile(
  entry,
  `export {SpreadsheetGrid} from ${JSON.stringify(path.join(component, 'ui/SpreadsheetGrid.tsx'))};export {configureXlsxHost} from ${JSON.stringify(path.join(component, 'ui/host.tsx'))};export {WorkbookSort} from ${JSON.stringify(path.join(component, 'ui/WorkbookSort.tsx'))};export * from ${JSON.stringify(path.join(component, 'ui/sortSelection.ts'))};`
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
  field = (name) => document.querySelector('[aria-label="' + name + '"]'),
  buttons = () => [...document.querySelectorAll('button')],
  button = (name) => buttons().find((n) => n.textContent === name && !n.closest('[hidden]'));
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
async function open() {
  await click([...document.querySelectorAll('[role="tab"]')].find((n) => n.textContent === 'Data'));
  await click('Sort…');
  await flush();
}
const binding = {
  documentId: 'doc',
  state: 'head',
  sheet: 'Sheet1',
  selection: 'A1:C5',
  activationEpoch: 0,
  ownerIdentity: 'owner',
  selectedCell: 'A1',
  navigationGeneration: 0,
  modeGeneration: 0,
  specialSelectionEpoch: 0,
  mode: 'normal',
  specialSelection: null,
};
function context(patch = {}) {
  return {
    binding: { ...binding },
    serial: 1,
    caption: 'Sheet1!A1:C5',
    initialColumn: 'A',
    initialHeader: true,
    isCurrent: () => true,
    isRuntimeCurrent: () => true,
    disabledReason: () => undefined,
    read: async () => ({ A1: { value: 'Region' }, B1: { value: 'Amount' }, C1: { value: 'Code' } }),
    accept: async () => {},
    ...patch,
  };
}
async function mount(t, patch = {}) {
  const root = createRoot(document.getElementById('root')),
    accepted = [],
    legacy = [],
    reads = [],
    data = {
      A1: { value: 'Region' },
      B1: { value: 'Amount' },
      C1: { value: 'Code' },
      A2: { value: 'West' },
      B2: { value: 10 },
      C2: { value: 'A' },
      A3: { value: 'East' },
      B3: { value: 20 },
      C3: { value: 'B' },
      A4: { value: 'West' },
      B4: { value: 30 },
      C4: { value: 'C' },
      A5: { value: 'East' },
      B5: { value: 5 },
      C5: { value: 'D' },
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
  });
  let readerOverride = null,
    input = {
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
      ...patch,
    };
  const reader = async (sheet, range) => {
    reads.push({ sheet, range });
    if (readerOverride) return readerOverride(sheet, range);
    const match = /^([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)$/.exec(range);
    assert(match);
    return Object.fromEntries(
      Object.entries(data).filter(
        ([ref]) =>
          ref.endsWith(match[2]) && ref.replace(/[0-9]+$/, '') >= match[1] && ref.replace(/[0-9]+$/, '') <= match[3]
      )
    );
  };
  const render = async (patch) => {
    input = { ...input, ...patch };
    await act(async () => root.render(React.createElement(ui.SpreadsheetGrid, input)));
    await flush();
  };
  await render();
  t.after(async () => {
    await act(async () => root.unmount());
  });
  return { accepted, legacy, reads, data, reader, holdRead: (fn) => (readerOverride = fn), render, input: () => input };
}
check(
  'ordered key model produces one exact header-excluded command; duplicate/outside/mixed shapes and unsafe dimensions refuse',
  () => {
    const keys = [
      { column: 'C', direction: 'desc' },
      { column: 'A', direction: 'asc' },
      { column: 'B', direction: 'desc' },
    ];
    assert.deepEqual(ui.buildSelectionSort('S', 'A1:C5', true, keys), {
      operation: 'sortRows',
      sheet: 'S',
      range: 'A2:C5',
      keys,
    });
    assert.equal(ui.buildSelectionSort('S', 'A1:C5', false, keys).range, 'A1:C5');
    for (const invalid of [
      [
        { column: 'A', direction: 'asc' },
        { column: 'A', direction: 'desc' },
      ],
      [{ column: 'D', direction: 'asc' }],
      [{ column: 'A', direction: 'up' }],
      [{ column: 'A', direction: 'asc', extra: 1 }],
      [],
    ])
      assert.throws(() => ui.buildSelectionSort('S', 'A1:C5', true, invalid));
    for (const range of ['A1:C2', 'C5:A1', 'A1:ZZ100', 'A0:C5', 'A1:XFE3'])
      assert.throws(() => ui.buildSelectionSort('S', range, true, [{ column: 'A', direction: 'asc' }]));
  }
);
check(
  'header reads are bounded and original-only; pending/formula labels and malformed out-of-range receipts are never promoted',
  async () => {
    const calls = [],
      ctx = context({
        binding: { ...binding, selection: 'A1001:ALM1003' },
        read: async (range) => {
          calls.push(range);
          return range.startsWith('A1001:')
            ? {
                A1001: { value: 'Region' },
                B1001: { formula: '1+1', calculation: { status: 'calculated', value: 2 } },
                C1001: { value: 'Old', calculation: { pending: true, status: 'stored' } },
              }
            : { ALM1001: { value: 'End' } };
        },
      });
    assert.deepEqual(await ui.readSortHeaders(ctx), { A: 'Region', ALM: 'End' });
    assert.equal(calls.length, 2);
    assert.match(calls[0], /^A1001:ALL1001$/); // ALL is column1000; first chunk stays at the transport bound.
    for (const cells of [{ A1: null }, { A1: [] }, { A1: 'bad' }, { D1: { value: 'foreign' } }])
      await assert.rejects(ui.readSortHeaders(context({ read: async () => cells })), /header reply/);
  }
);
check(
  'real dock edits levels independently, reorders/removes them and accepts one mixed-direction command for the captured range',
  async (t) => {
    const h = await mount(t, {
      tables: [
        {
          name: 'Sales',
          sheet: 'Sheet1',
          ref: 'A1:C5',
          columns: [{ name: 'Region' }, { name: 'Amount' }, { name: 'Code' }],
        },
      ],
    });
    await h.render({ readRange: h.reader });
    await go('A1:C5');
    await open();
    assert.equal(document.querySelectorAll('.sheet-task-pane').length, 1);
    assert.match(field('Sort').textContent, /Sheet1!A1:C5 · Sales/);
    assert.match(field('Sort column 1').textContent, /A — Region/);
    await click('Add level');
    await change(field('Sort order 2'), 'desc');
    await click('Add level');
    await change(field('Sort order 3'), 'desc');
    await click(field('Move sort level 3 up'));
    await click(field('Move sort level 2 up'));
    assert.equal(field('Sort column 1').value, 'C');
    await click(field('Remove sort level 3'));
    await click('Add level');
    await change(field('Sort order 3'), 'desc');
    await click('Apply sort');
    assert.deepEqual(h.accepted, [
      [
        {
          operation: 'sortRows',
          sheet: 'Sheet1',
          range: 'A2:C5',
          keys: [
            { column: 'C', direction: 'desc' },
            { column: 'A', direction: 'asc' },
            { column: 'B', direction: 'desc' },
          ],
        },
      ],
    ]);
    assert.match(field('Sort').textContent, /Rows sorted/);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'quick ascending/descending use one acknowledged key and explicit header choice, never chained secondary sorts',
  async (t) => {
    const h = await mount(t);
    await go('A1:C5');
    await click([...document.querySelectorAll('[role="tab"]')].find((n) => n.textContent === 'Data'));
    await change(field('Primary sort column'), 'B');
    await click('Sort descending');
    assert.deepEqual(h.accepted.pop(), [
      { operation: 'sortRows', sheet: 'Sheet1', range: 'A2:C5', keys: [{ column: 'B', direction: 'desc' }] },
    ]);
    const header = [...document.querySelectorAll('label')]
      .find((n) => n.textContent === 'Header row')
      .querySelector('input');
    await change(header, false);
    await click('Sort ascending');
    assert.deepEqual(h.accepted.pop(), [
      { operation: 'sortRows', sheet: 'Sheet1', range: 'A1:C5', keys: [{ column: 'B', direction: 'asc' }] },
    ]);
    assert.equal(field('Secondary sort column'), null);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'held header replies and retained Apply handlers retire on head/runtime/sheet/selection round trips, never sorting a new target',
  async (t) => {
    const h = await mount(t);
    for (const kind of ['head', 'owner', 'sheet', 'selection']) {
      await go('A1:C5');
      const held = defer();
      await h.render({ readRange: async (_sheet, range) => (range === 'A1:C1' ? held.promise : {}) });
      await open();
      const oldApply = props(button('Apply sort')).onClick;
      if (kind === 'head') await h.render({ storageKey: 'head-' + kind });
      else if (kind === 'owner') await h.render({ ownerIdentity: 'owner-' + kind });
      else if (kind === 'sheet') {
        await click(button('Other'));
        await click(button('Sheet1'));
      } else {
        await go('B1:C5');
        await go('A1:C5');
      }
      await act(async () => held.resolve({ A1: { value: 'Obsolete' } }));
      await flush();
      assert.match(field('Sort').textContent, /changed. Reopen Sort/);
      if (button('Apply sort')) assert(button('Apply sort').disabled);
      await act(async () => oldApply());
      await flush();
      assert.deepEqual(h.accepted, []);
      assert(!field('Sort').textContent.includes('Obsolete'));
      await click(field('Close workbook tools'));
    }
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'one pending owner acceptance disables duplicate Apply; truthful accepted head advance reports success without old-head resubmit',
  async (t) => {
    const held = defer(),
      h = await mount(t, {
        onCommandBatch: (commands) => {
          h.accepted.push(commands);
          return held.promise;
        },
      });
    await go('A1:C5');
    await open();
    const submit = props(button('Apply sort')).onClick;
    await act(async () => {
      submit();
      submit();
    });
    assert.equal(h.accepted.length, 1);
    assert(button('Apply sort').disabled);
    await h.render({ storageKey: 'accepted-head' });
    await act(async () => held.resolve());
    await flush();
    assert.match(field('Sort').textContent, /Rows sorted/);
    assert(button('Apply sort').disabled);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'late original-owner failures never land on a replacement owner and stale quick handlers cannot disturb its current draft',
  async (t) => {
    const held = defer(),
      h = await mount(t, {
        onCommandBatch: (commands) => {
          h.accepted.push(commands);
          return held.promise;
        },
      });
    await go('A1:C5');
    await click([...document.querySelectorAll('[role="tab"]')].find((n) => n.textContent === 'Data'));
    const oldQuick = props(button('Sort ascending')).onClick,
      oldOpen = props(button('Sort…')).onClick;
    await open();
    await click('Apply sort');
    await h.render({ ownerIdentity: 'replacement', documentId: 'replacement-doc' });
    const input = field('Formula or value');
    await act(async () => {
      input.focus();
      props(input).onChange({ target: { value: '=SUM(' } });
    });
    input.setSelectionRange(5, 5);
    await act(async () => {
      oldQuick();
      oldOpen();
      held.reject(Error('Obsolete sort failed'));
    });
    await flush();
    assert.equal(input.value, '=SUM(');
    assert.equal(input.selectionStart, 5);
    assert.equal(document.activeElement, input);
    assert(!document.body.textContent.includes('Obsolete sort failed'));
    assert.equal(h.accepted.length, 1);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'readonly, missing strict/runtime capability, inactive/passive/literal and special selections refuse sort but keep real reasons',
  async (t) => {
    const h = await mount(t);
    await go('A1:C5');
    await click([...document.querySelectorAll('[role="tab"]')].find((n) => n.textContent === 'Data'));
    for (const patch of [
      { readOnly: true },
      { readOnly: false, busy: true },
      { busy: false, onCommandBatch: undefined },
      { onCommandBatch: async () => {}, ownerIdentity: undefined },
      { ownerIdentity: 'owner', active: false },
      { active: true, passiveNotice: 'Passive inspector' },
    ]) {
      await h.render(patch);
      assert(button('Sort ascending').disabled);
      assert(button('Sort ascending').title);
      assert(button('Sort…').disabled);
    }
    await h.render({
      passiveNotice: undefined,
      sheets: h.input().sheets.map((sheet) => ({ ...sheet, cells: { ...sheet.cells, B2: { value: null } } })),
      onCommandBatch: async () => {},
    });
    await click([...document.querySelectorAll('[role="tab"]')].find((n) => n.textContent === 'Home'));
    await click(document.querySelector('[data-home-command="home.select-blanks"]'));
    await click('Preview matching cells');
    await click('Select matching cells');
    assert(document.querySelector('.sheet-special-status'));
    await click([...document.querySelectorAll('[role="tab"]')].find((n) => n.textContent === 'Data'));
    assert(button('Sort ascending').disabled);
    assert.match(button('Sort ascending').title, /rectangular/);
    await h.render({ passiveNotice: undefined, literal: true });
    assert.equal(button('Sort ascending'), undefined);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'pure Sort opening, header/level changes, Cancel and dock close preserve a parked formula draft and do not opportunistically commit a live grid input',
  async (t) => {
    const h = await mount(t);
    await go('A1:C5');
    await click([...document.querySelectorAll('[role="tab"]')].find((n) => n.textContent === 'Data'));
    const input = field('Formula or value');
    await act(async () => {
      input.focus();
      props(input).onChange({ target: { value: '=SUM(' } });
    });
    input.setSelectionRange(5, 5);
    await click('Sort…');
    await click('Add level');
    await change(field('Sort order 2'), 'desc');
    await click('Cancel sort');
    await click(field('Close workbook tools'));
    assert.equal(input.value, '=SUM(');
    assert.equal(input.selectionStart, 5);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
    const oldOpen = props(button('Sort…')).onClick,
      cell = field('Sheet1 A1');
    await act(async () => {
      cell.focus();
      cell.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'F2', bubbles: true, cancelable: true }));
    });
    await act(async () => oldOpen());
    await flush();
    assert(document.querySelector('.sheet-task-pane').hidden);
    assert.equal(cell.dataset.editing, 'true');
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
  }
);
check('pane enforces16 distinct levels and explicit duplicate-key refusal without calling the owner', async (t) => {
  const root = createRoot(document.getElementById('root')),
    accepted = [],
    ctx = context({
      binding: { ...binding, selection: 'A1:Q3' },
      read: async () => ({}),
      accept: async (commands) => accepted.push(commands),
    });
  t.after(async () => act(async () => root.unmount()));
  await act(async () => root.render(React.createElement(ui.WorkbookSort, { context: ctx, visible: true })));
  await flush();
  for (let i = 1; i < 16; i++) await click('Add level');
  assert.equal(document.querySelectorAll('[data-sort-level]').length, 16);
  assert(button('Add level').disabled);
  await change(field('Sort column 2'), 'A');
  assert(button('Apply sort').disabled);
  assert.match(field('Sort').textContent, /Duplicate sort key/);
  assert.deepEqual(accepted, []);
  await change(field('Sort column 2'), 'B');
  await change(field('Sort order 16'), 'desc');
  await click('Apply sort');
  assert.equal(accepted.length, 1);
  assert.equal(accepted[0].length, 1);
  assert.equal(accepted[0][0].keys.length, 16);
  assert.deepEqual(accepted[0][0].keys.at(-1), { column: 'P', direction: 'desc' });
});
