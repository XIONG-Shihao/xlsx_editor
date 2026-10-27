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
  folder = path.join(component, '.build/tests'),
  runs = await fs.mkdtemp(path.join(folder, 'watch-ui-')),
  entry = path.join(runs, 'entry.ts'),
  bundle = path.join(runs, 'test.mjs');
await fs.writeFile(
  entry,
  `export {WorkbookWatch} from ${JSON.stringify(path.join(component, 'ui/WorkbookWatch.tsx'))};export {createScopedWatchReader} from ${JSON.stringify(path.join(component, 'ui/watchAdapter.ts'))};export {SpreadsheetGrid} from ${JSON.stringify(path.join(component, 'ui/SpreadsheetGrid.tsx'))};export {configureXlsxHost} from ${JSON.stringify(path.join(component, 'ui/host.tsx'))};`
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
      return 500;
    },
  },
  clientHeight: {
    get() {
      return 260;
    },
  },
});
globalThis.ResizeObserver = class {
  constructor(fn) {
    this.fn = fn;
  }
  observe() {
    this.fn([{ contentRect: { width: 500, height: 260 } }]);
  }
  disconnect() {}
};
globalThis.requestAnimationFrame = () => 1;
globalThis.cancelAnimationFrame = () => {};
const { createRoot } = await import('react-dom/client'),
  { WorkbookWatch, createScopedWatchReader, SpreadsheetGrid, configureXlsxHost } = await import(pathToFileURL(bundle));
configureXlsxHost({
  api: { copyText: async () => {}, readClipboardText: async () => '' },
  Icon: () => React.createElement('svg'),
  useI18n: () => ({ locale: 'en' }),
});
after(async () => {
  dom.window.close();
});
const deferred = () => {
  let resolve, reject;
  const promise = new Promise((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
};
const scope = () => ({
  documentId: 'doc-A',
  ownerIdentity: 'owner-A',
  state: 'base',
  activationEpoch: 0,
  active: true,
  sheets: ['Sheet1'],
});
const button = (name) =>
  [...document.querySelectorAll('button')].find(
    (node) => !node.closest('[hidden]') && (node.getAttribute('aria-label') ?? node.textContent) === name
  );
async function click(node) {
  assert(node);
  assert.equal(node.disabled, false);
  await act(async () => node.click());
}
const value = (ref) => document.querySelector('output[aria-label="Watch value Sheet1!' + ref + '"]')?.textContent;
async function mount(t, extra = {}) {
  const root = createRoot(document.getElementById('root'));
  let props = {
    scope: scope(),
    read: async (bound, target) => ({ scope: bound, target, cell: { value: 1 } }),
    selected: { sheet: 'Sheet1', reference: 'A1' },
    ...extra,
  };
  const draw = () => React.createElement(WorkbookWatch, props);
  await act(async () => root.render(extra.strict ? React.createElement(React.StrictMode, null, draw()) : draw()));
  t.after(async () => act(async () => root.unmount()));
  return {
    render: async (patch) => {
      props = { ...props, ...patch };
      await act(async () => root.render(extra.strict ? React.createElement(React.StrictMode, null, draw()) : draw()));
    },
  };
}
test('watch pane distinguishes typed scalar/error/unresolved/stored and protected raw precision without workbook mutations', async (t) => {
  const cells = {
    A1: { value: false },
    A2: { value: '' },
    A3: { value: '#N/A', dataType: 'text' },
    A4: { value: '#N/A', dataType: 'error' },
    A5: { value: null, formula: 'SUM(A1:A2)', calculation: { status: 'circular', error: '#CIRC!' } },
    A6: { value: null, formula: 'UNKNOWN()', calculation: { status: 'unresolved', error: 'Unknown function' } },
    A7: { value: 99, formula: 'UNKNOWN()', cachedValue: 99, calculation: { status: 'stored', value: 99 } },
    A8: { value: 9007199254740992, rawValue: '9007199254740993', dataType: 'number' },
  };
  const ui = await mount(t, {
    read: async (bound, target) => ({ scope: bound, target, cell: cells[target.reference] }),
  });
  for (let row = 1; row <= 8; row++) {
    await ui.render({ selected: { sheet: 'Sheet1', reference: 'A' + row } });
    await click(button('Add selected cell'));
  }
  assert.equal(value('A1'), 'FALSE');
  assert.equal(value('A2'), '""');
  assert.equal(value('A3'), '"#N/A"');
  assert.equal(value('A4'), '#N/A');
  assert.equal(value('A5'), '#CIRC!');
  assert.equal(value('A6'), 'Unknown function');
  assert.equal(value('A7'), 'Stored result; not recalculated.');
  assert.equal(value('A8'), 'Protected number: 9007199254740993');
  assert(!document.body.textContent.includes('9007199254740992'));
});
test('accepted head change masks old values immediately and adopts only the new accepted reply', async (t) => {
  const gate = deferred(),
    ui = await mount(t, { read: async (bound, target) => ({ scope: bound, target, cell: { value: 77 } }) });
  await click(button('Add selected cell'));
  assert.equal(value('A1'), '77');
  await ui.render({
    scope: { ...scope(), state: 'next' },
    read: async (bound, target) => ({ scope: bound, target, cell: await gate.promise }),
  });
  assert.notEqual(value('A1'), '77');
  await act(async () => gate.resolve({ value: 88 }));
  assert.equal(value('A1'), '88');
});
test('hidden/inactive pane masks values and pauses; actual owner replacement clears its targets', async (t) => {
  let reads = 0;
  const ui = await mount(t, {
    read: async (bound, target) => {
      reads++;
      return { scope: bound, target, cell: { value: 41 } };
    },
  });
  await click(button('Add selected cell'));
  assert.equal(value('A1'), '41');
  await ui.render({ scope: { ...scope(), active: false } });
  assert.notEqual(value('A1'), '41');
  assert.match(document.body.textContent, /Paused/);
  assert.equal(reads, 1);
  await ui.render({ scope: { ...scope(), activationEpoch: 2 } });
  assert.equal(value('A1'), '41');
  assert.equal(reads, 2);
  await ui.render({ scope: { ...scope(), ownerIdentity: 'owner-replaced' } });
  assert.equal(document.querySelector('[data-watch-target]'), null);
});
test('ignored abort never overlaps reads and Cancel does not silently restart a retired refresh', async (t) => {
  const gate = deferred(),
    calls = [];
  let active = 0,
    peak = 0;
  const ui = await mount(t, {
    read: async (bound, target) => {
      calls.push(bound.state);
      active++;
      peak = Math.max(peak, active);
      const cell = calls.length === 1 ? await gate.promise : { value: 2 };
      active--;
      return { scope: bound, target, cell };
    },
  });
  await click(button('Add selected cell'));
  await click(button('Cancel refresh'));
  assert.deepEqual(calls, ['base']);
  await act(async () => gate.resolve({ value: 1 }));
  assert.deepEqual(calls, ['base']);
  assert.notEqual(value('A1'), '1');
  await click(button('Refresh watches'));
  assert.equal(value('A1'), '2');
  assert.equal(peak, 1);
});
test('scope auto-refresh joins a late old physical read before requesting the latest owner head', async (t) => {
  const gate = deferred(),
    calls = [];
  let active = 0,
    peak = 0;
  const ui = await mount(t, {
    read: async (bound, target) => {
      calls.push(bound.state);
      active++;
      peak = Math.max(peak, active);
      const cell = calls.length === 1 ? await gate.promise : { value: 22 };
      active--;
      return { scope: bound, target, cell };
    },
  });
  await click(button('Add selected cell'));
  await ui.render({ scope: { ...scope(), state: 'next' } });
  assert.deepEqual(calls, ['base']);
  await act(async () => gate.resolve({ value: 11 }));
  assert.deepEqual(calls, ['base', 'next']);
  assert.equal(value('A1'), '22');
  assert.equal(peak, 1);
});
test('scoped adapter fences original reader/head before and after one exact-cell read and refuses foreign replies', async () => {
  const gate = deferred(),
    calls = [];
  let frame = {
    scope: scope(),
    read: async (sheet, range) => {
      calls.push({ sheet, range });
      return gate.promise;
    },
  };
  const reader = createScopedWatchReader(() => frame),
    signal = new AbortController().signal,
    target = { sheet: 'Sheet1', reference: 'A1' };
  const pending = reader(scope(), target, signal);
  frame = { ...frame, scope: { ...scope(), state: 'next' } };
  gate.resolve({ A1: { value: 1 } });
  await assert.rejects(pending, /accepted state changed/);
  assert.deepEqual(calls, [{ sheet: 'Sheet1', range: 'A1:A1' }]);
  await assert.rejects(reader(scope(), target, signal), /accepted state changed/);
  assert.equal(calls.length, 1);
  frame = { scope: scope(), read: async () => ({ B1: { value: 9 } }) };
  await assert.rejects(reader(scope(), target, signal), /watched target/);
});
test('regular canonical selection, duplicate/cap/remove/clear behavior is bounded and readonly', async (t) => {
  let reads = 0;
  const ui = await mount(t, {
    read: async (bound, target) => {
      reads++;
      return { scope: bound, target, cell: { value: 1 } };
    },
  });
  await click(button('Add selected cell'));
  await click(button('Add selected cell'));
  assert.equal(document.querySelectorAll('[data-watch-target]').length, 1);
  assert.equal(reads, 1);
  await ui.render({ selected: { sheet: 'Sheet1', reference: 'A1:A2' } });
  await click(button('Add selected cell'));
  assert.match(document.querySelector('[role="alert"]').textContent, /canonical cell/);
  for (let row = 2; row <= 32; row++) {
    await ui.render({ selected: { sheet: 'Sheet1', reference: 'A' + row } });
    await click(button('Add selected cell'));
  }
  assert.equal(document.querySelectorAll('[data-watch-target]').length, 32);
  assert.equal(button('Add selected cell').disabled, true);
  await click(button('Remove watch Sheet1!A1'));
  assert.equal(document.querySelectorAll('[data-watch-target]').length, 31);
  await click(button('Clear watches'));
  assert.equal(document.querySelectorAll('[data-watch-target]').length, 0);
});
test('StrictMode effect rehearsal does not dispose the live retained watch controller', async (t) => {
  await mount(t, { strict: true });
  assert(!document.body.textContent.includes('unavailable'));
  await click(button('Add selected cell'));
  assert.equal(value('A1'), '1');
});
test('production Grid uses one dock/auditing group and a real scoped readonly reader; range/merge/special selection cannot add', async (t) => {
  const root = createRoot(document.getElementById('root')),
    reads = [],
    writes = [];
  const sheet = {
    name: 'Sheet1',
    state: 'visible',
    path: 'xl/worksheets/sheet1.xml',
    rows: 64,
    columns: 12,
    cells: {},
    merges: ['C3:D3'],
    widths: [],
    hiddenRows: [],
    rowHeights: [],
    freeze: { rows: 0, columns: 0 },
  };
  await act(async () =>
    root.render(
      React.createElement(SpreadsheetGrid, {
        sheets: [sheet],
        documentId: 'doc-A',
        ownerIdentity: 'owner-A',
        storageKey: 'owner-A:base',
        readOnly: true,
        onEdit: () => writes.push('edit'),
        onBatch: () => writes.push('batch'),
        readRange: async (name, range) => {
          reads.push({ name, range });
          return range === 'A1:A1' ? { A1: { value: 0 } } : {};
        },
      })
    )
  );
  t.after(async () => act(async () => root.unmount()));
  await click([...document.querySelectorAll('[role="tab"]')].find((node) => node.textContent === 'Formulas'));
  await click(button('Watch Window'));
  const groups = [
    ...document.querySelector('.sheet-ribbon-tabpanel:not([hidden])').querySelectorAll('.sheet-ribbon-group'),
  ].filter((group) => group.querySelector('small')?.textContent === 'Formula auditing');
  assert.equal(groups.length, 1);
  assert.equal(document.querySelectorAll('aside.sheet-task-pane:not([hidden])').length, 1);
  await click(button('Add selected cell'));
  assert.equal(value('A1'), '0');
  assert(reads.some((item) => item.range === 'A1:A1'));
  assert.deepEqual(writes, []);
  const go = async (ref) => {
    const input = document.querySelector('[aria-label="Go to cell"]');
    await act(async () => {
      Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(input, ref);
      input.dispatchEvent(new window.Event('input', { bubbles: true }));
    });
    await act(async () =>
      input.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
    );
  };
  await go('A1:A2');
  assert.equal(button('Add selected cell').disabled, true);
  await go('C3');
  assert.equal(button('Add selected cell').disabled, true);
  assert.match(document.body.textContent, /unmerged/);
  await go('A1:A2');
  await click([...document.querySelectorAll('[role="tab"]')].find((node) => node.textContent === 'Home'));
  await click(button('Editing'));
  const menu = [...document.querySelectorAll('details')].find(
    (node) => !node.closest('[hidden]') && node.textContent.includes('Find & Select')
  );
  menu.open = true;
  await click(button('Select blanks'));
  assert.equal(button('Add selected cell').disabled, true);
  assert.match(document.body.textContent, /regular cell/);
  assert.deepEqual(writes, []);
});

test('protected long decimal keeps its exact stored spelling while ordinary zero and signed zero remain typed', async (t) => {
  const raw = '0.12345678901234567890123456789',
    cells = {
      A1: { value: Number(raw), rawValue: raw, dataType: 'number' },
      A2: { value: 0, rawValue: '0', dataType: 'number' },
      A3: { value: -0, rawValue: '-0', dataType: 'number' },
    };
  const ui = await mount(t, {
    read: async (bound, target) => ({ scope: bound, target, cell: cells[target.reference] }),
  });
  for (let row = 1; row <= 3; row++) {
    await ui.render({ selected: { sheet: 'Sheet1', reference: 'A' + row } });
    await click(button('Add selected cell'));
  }
  assert.equal(value('A1'), 'Protected number: ' + raw);
  assert.equal(value('A2'), '0');
  assert.equal(value('A3'), '-0');
  assert(!document.body.textContent.includes('0.12345678901234568'));
});

test('opaque UI state token is never parsed and the original host reader supplies its actual accepted head', async () => {
  const gate = deferred(),
    calls = [],
    bound = { ...scope(), state: 'opaque:owner-with:colons:head-token', sheets: ['Sheet1', 'Sheet2'] },
    target = { sheet: 'Sheet2', reference: 'C90' };
  const capturedHost = (nativeState) => async (sheet, range) => {
    calls.push({ nativeState, sheet, range });
    return gate.promise;
  };
  let frame = { scope: bound, read: capturedHost('native-journal-head-A') };
  const reader = createScopedWatchReader(() => frame),
    signal = new AbortController().signal;
  const pending = reader(bound, target, signal);
  frame = { scope: bound, read: capturedHost('must-not-retarget') };
  gate.resolve({ C90: { value: 5 } });
  assert.equal((await pending).cell.value, 5);
  assert.deepEqual(calls, [{ nativeState: 'native-journal-head-A', sheet: 'Sheet2', range: 'C90:C90' }]);
  frame = { scope: { ...bound, sheets: ['Sheet1'] }, read: capturedHost('unavailable-sheet') };
  await assert.rejects(reader(bound, target, signal), /accepted state changed/);
  assert.equal(calls.length, 1);
});

test('actual retained Grid masks on accepted head/hide and clears on replaced owner without any workbook commands', async (t) => {
  const root = createRoot(document.getElementById('root')),
    gate = deferred(),
    calls = [],
    writes = [];
  const sheet = {
    name: 'Sheet1',
    state: 'visible',
    path: 'xl/worksheets/sheet1.xml',
    rows: 64,
    columns: 12,
    cells: {},
    merges: [],
    widths: [],
    hiddenRows: [],
    rowHeights: [],
    freeze: { rows: 0, columns: 0 },
  };
  let state = 'opaque-A',
    owner = 'worker-A';
  const reader = (head) => async (name, range) => {
    if (range !== 'A1:A1') return {};
    calls.push(head);
    return { [range.split(':')[0]]: head === 'opaque-A' ? { value: 41 } : await gate.promise };
  };
  const render = async () =>
    act(async () =>
      root.render(
        React.createElement(SpreadsheetGrid, {
          sheets: [sheet],
          documentId: 'doc-A',
          ownerIdentity: owner,
          storageKey: state,
          readOnly: true,
          onBatch: () => writes.push('batch'),
          onCommandBatch: async () => writes.push('command'),
          readRange: reader(state),
        })
      )
    );
  await render();
  t.after(async () => act(async () => root.unmount()));
  await click([...document.querySelectorAll('[role="tab"]')].find((node) => node.textContent === 'Formulas'));
  await click(button('Watch Window'));
  await click(button('Add selected cell'));
  assert.equal(value('A1'), '41');
  state = 'opaque-B';
  await render();
  assert.notEqual(value('A1'), '41');
  await click(button('Close workbook tools'));
  await act(async () => gate.resolve({ value: 42 }));
  assert.notEqual(value('A1'), '42');
  await click(button('Watch Window'));
  assert.equal(value('A1'), '42');
  owner = 'worker-B';
  await render();
  assert.equal(document.querySelector('[data-watch-target]'), null);
  assert.deepEqual(writes, []);
  assert.deepEqual(calls, ['opaque-A', 'opaque-B', 'opaque-B']);
});

test('central Page Layout pauses and masks the retained watch pane until Normal returns, without editing', async (t) => {
  const root = createRoot(document.getElementById('root')),
    writes = [],
    reads = [];
  const sheet = {
      name: 'Sheet1',
      state: 'visible',
      path: 'xl/worksheets/sheet1.xml',
      rows: 64,
      columns: 12,
      cells: {},
      merges: [],
      widths: [],
      hiddenRows: [],
      rowHeights: [],
      freeze: { rows: 0, columns: 0 },
    },
    settings = {
      area: 'A1:B3',
      paper: 'Letter',
      orientation: 'portrait',
      marginsPt: { top: 36, right: 36, bottom: 36, left: 36, header: 18, footer: 18 },
      scaling: { mode: 'scale', percent: 100 },
      repeatRows: null,
      repeatColumns: null,
      gridlines: false,
    };
  const page = {
      number: 1,
      rowIndices: [0],
      columnIndices: [0],
      cells: [{ ref: 'A1', text: '41', style: { fontSize: 11 }, bbox: [36, 36, 100, 30] }],
      objects: [],
    },
    controls = {
      documentId: 'runtime-A',
      revision: 'base',
      definition: async () => ({ sheet: 'Sheet1', revision: null, settings, canEdit: true, protectedReasons: [] }),
      preview: async (_sheet, requested) => ({
        id: 'accepted-page',
        state: 'base',
        units: 'pt',
        pageSizePt: [612, 792],
        scale: 1,
        sourceFingerprint: 'a'.repeat(64),
        revision: 'source-A',
        sheet: 'Sheet1',
        settings: requested,
        pageCount: 1,
        pages: [page],
      }),
      page: async () => page,
      onBatch: async () => writes.push('print'),
      publish: async () => {},
      release: async () => {},
      cancel: async () => ({ granted: true }),
    };
  await act(async () =>
    root.render(
      React.createElement(SpreadsheetGrid, {
        sheets: [sheet],
        documentId: 'doc-A',
        ownerIdentity: 'worker-A',
        storageKey: 'worker-A:base',
        readRange: async (_sheet, range) => {
          if (range === 'A1:A1') {
            reads.push(range);
            return { A1: { value: 41 } };
          }
          return {};
        },
        printControls: controls,
        onEdit: () => writes.push('edit'),
        onCommandBatch: async () => writes.push('command'),
      })
    )
  );
  t.after(async () => act(async () => root.unmount()));
  const tab = (name) => [...document.querySelectorAll('[role="tab"]')].find((node) => node.textContent === name);
  await click(tab('Formulas'));
  await click(button('Watch Window'));
  await click(button('Add selected cell'));
  assert.equal(value('A1'), '41');
  await click(tab('View'));
  await click(document.querySelector('[data-ribbon-command="view.page-layout"]'));
  assert(document.querySelector('.sheet-page-layout'));
  assert.notEqual(value('A1'), '41');
  await click(tab('Formulas'));
  await click(button('Watch Window'));
  assert.equal(button('Add selected cell').disabled, true);
  assert.match(document.querySelector('section[aria-label="Watch Window"]').textContent, /Paused/);
  await click(document.querySelector('.sheet-page-layout button'));
  assert.equal(document.querySelector('.sheet-page-layout'), null);
  assert.equal(button('Add selected cell').disabled, false);
  assert.equal(value('A1'), '41');
  assert.deepEqual(writes, []);
  assert.equal(reads.length, 2);
});
