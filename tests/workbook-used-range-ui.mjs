import './support/scratch-dir.mjs';
/** Actual Grid/pane with controlled provider reads. No native/app qualification. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createHash } from 'node:crypto';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
import {
  providerReceipt,
  initialCells,
  completeUsedRange,
  incompleteUsedRange,
  contains,
} from './support/used-range-selection-fixture.mjs';

const component = fileURLToPath(new URL('../', import.meta.url)),
  output = await fs.mkdtemp(path.join(component, '.build/tests/used-range-grid-ui-')),
  checks = [];
const files = [
  'ui/SpreadsheetGrid.tsx',
  'ui/SpreadsheetRibbon.tsx',
  'ui/homeCommands.ts',
  'ui/WorkbookSpecialSelection.tsx',
  'ui/specialSelection.ts',
  'core/contracts.ts',
  'core/used-range.mjs',
  'tests/workbook-used-range-ui.mjs',
  'tests/support/used-range-selection-fixture.mjs',
];
const pins = () =>
    Promise.all(
      files.map(async (relative) => {
        const file = path.join(component, relative),
          bytes = await fs.readFile(file);
        return { relative, path: file, sha256: createHash('sha256').update(bytes).digest('hex'), bytes: bytes.length };
      })
    ),
  before = await pins();
const entry = path.join(output, 'entry.ts'),
  bundle = path.join(output, 'grid.mjs');
await fs.writeFile(
  entry,
  `export {SpreadsheetGrid} from ${JSON.stringify(path.join(component, 'ui/SpreadsheetGrid.tsx'))};export {configureXlsxHost} from ${JSON.stringify(path.join(component, 'ui/host.tsx'))};`
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
const props = (node) => node[Object.keys(node).find((key) => key.startsWith('__reactProps'))],
  field = (name) => document.querySelector('[aria-label=' + JSON.stringify(name) + ']'),
  command = (id) =>
    [...document.querySelectorAll('[data-home-command=' + JSON.stringify(id) + ']')].find(
      (node) => !node.closest('[hidden]')
    );
const pane = () => field('Go To Special'),
  button = (name) => [...pane().querySelectorAll('button')].find((node) => node.textContent === name),
  flush = async (ms = 12) => act(async () => new Promise((resolve) => setTimeout(resolve, ms))),
  defer = () => {
    let resolve, reject;
    const promise = new Promise((yes, no) => {
      resolve = yes;
      reject = no;
    });
    return { promise, resolve, reject };
  };
async function click(node) {
  assert(node, 'Missing control');
  assert.equal(node.disabled, false);
  await act(async () => node.click());
  await flush();
  return node;
}
async function change(node, value) {
  assert(node);
  await act(async () => props(node).onChange({ target: { value, checked: value } }));
  await flush();
}
async function go(reference) {
  await change(field('Go to cell'), reference);
  await act(async () =>
    field('Go to cell').dispatchEvent(
      new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })
    )
  );
  await flush();
}
async function settled() {
  for (let attempt = 0; attempt < 60; attempt++) {
    if (pane().querySelector('[aria-label="Special selection summary"],[role="alert"]')) return;
    await flush(15);
  }
  throw Error('UsedRange preview did not settle.');
}
async function preview(kind = 'constants') {
  await click(command('home.special-open'));
  await change(field('Selection type'), kind);
  await click(button('Preview matching cells'));
  await settled();
}
const targets = () =>
    [...field('Special selection targets').querySelectorAll('tbody tr')].map((row) => row.cells[0].textContent),
  alert = () => pane().querySelector('[role=alert]')?.textContent;

async function mount(t, patch = {}) {
  const root = createRoot(document.getElementById('root')),
    accepted = [],
    legacy = [],
    history = [],
    reads = [],
    extentReads = [],
    data = initialCells();
  let reply = completeUsedRange(),
    readOverride = null,
    extentOverride = null;
  const reader = async (sheet, range) => {
    reads.push({ sheet, range });
    if (readOverride) return readOverride(sheet, range);
    return Object.fromEntries(
      Object.entries(data)
        .filter(([ref]) => contains(range, ref))
        .map(([ref, value]) => [ref, structuredClone(value)])
    );
  };
  const extent = async (sheet) => {
    extentReads.push(sheet);
    return extentOverride ? extentOverride(sheet) : structuredClone(reply);
  };
  const source = {
    name: 'Sheet1',
    path: 'xl/worksheets/original-path.xml',
    state: 'visible',
    rows: 40,
    columns: 14,
    cells: structuredClone(data),
    merges: [],
    widths: [],
    hiddenRows: [],
    rowHeights: [],
    freeze: { rows: 0, columns: 0 },
    capabilities: { insertRows: true, insertRowsReason: null, deleteRows: true, referenceColumns: true },
  };
  let input = {
      sheets: [source],
      documentId: 'doc',
      ownerIdentity: 'standalone-registry-alias',
      storageKey: 'registry:accepted-head',
      active: true,
      readOnly: false,
      readRange: reader,
      usedRangeControls: { receipt: { ...providerReceipt }, read: extent },
      onEdit: (value) => legacy.push(value),
      onBatch: (value) => legacy.push(value),
      onCommandBatch: async (commands) => accepted.push(commands),
      onCommandUndo: async () => history.push('undo'),
      onCommandRedo: async () => history.push('redo'),
      canUndo: true,
      canRedo: true,
      ...patch,
    },
    mounted = true;
  const render = async (patch) => {
      input = { ...input, ...patch };
      await act(async () => root.render(React.createElement(ui.SpreadsheetGrid, input)));
      await flush();
    },
    unmount = async () => {
      if (!mounted) return;
      mounted = false;
      await act(async () => root.unmount());
      await flush();
    };
  await render();
  t.after(unmount);
  return {
    accepted,
    legacy,
    history,
    reads,
    extentReads,
    data,
    reader,
    extent,
    render,
    unmount,
    input: () => input,
    setReply: (value) => (reply = structuredClone(value)),
    holdExtent: (fn) => (extentOverride = fn),
    holdRead: (fn) => (readOverride = fn),
    pure: (before) => {
      assert.deepEqual(accepted, []);
      assert.deepEqual(legacy, []);
      assert.deepEqual(history, []);
      assert.deepEqual(data, before);
    },
  };
}
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

check(
  'Single-cell defaults use original provider receipt despite Home aliases; Apply is pure and preserves focus',
  async (t) => {
    const h = await mount(t);
    await go('A5');
    const before = structuredClone(h.data);
    await preview();
    assert.equal(alert(), undefined);
    assert.deepEqual(targets(), ['Sheet1!A1', 'Sheet1!C4', 'Sheet1!A5', 'Sheet1!C5', 'Sheet1!C6']);
    assert.match(field('Special selection summary').textContent, /Source positions18/);
    assert.equal(h.extentReads.length, 1);
    const focus = document.activeElement;
    await click(button('Select matching cells'));
    assert.equal(h.extentReads.length, 2);
    assert.equal(field('Go to cell').placeholder, 'A1');
    assert.equal(field('Formula or value').value, '7');
    assert.equal(document.activeElement, focus);
    assert.match(pane().textContent, /Matching cells selected/);
    h.pure(before);
    await click(field('Close workbook tools'));
    assert.equal(field('Go to cell').placeholder, 'A1');
    assert.equal(document.activeElement, field('Go to cell'));
    h.pure(before);
  }
);
check('Observed default Formulas and Blanks select the correct complete targets and active cursors', async (t) => {
  for (const [kind, expected, cursor] of [
    ['formulas', ['A3', 'C3'], 'A3'],
    ['blanks', ['B1', 'C1', 'A2', 'B2', 'C2', 'B3', 'A4', 'B4', 'B5', 'A6', 'B6'], 'A2'],
  ]) {
    const h = await mount(t);
    await go('A5');
    const before = structuredClone(h.data);
    await preview(kind);
    assert.equal(alert(), undefined);
    assert.deepEqual(
      targets(),
      expected.map((ref) => 'Sheet1!' + ref)
    );
    await click(button('Select matching cells'));
    assert.equal(field('Go to cell').placeholder, cursor);
    h.pure(before);
    await h.unmount();
  }
});
check(
  'Explicit rectangle and Visible single-cell preserve local scope; Last Cell ignores explicit rectangle',
  async (t) => {
    const h = await mount(t);
    await go('A1:B2');
    await preview();
    assert.deepEqual(targets(), ['Sheet1!A1']);
    assert.equal(h.extentReads.length, 0);
    await click(field('Close workbook tools'));
    await go('A5');
    await preview('visible');
    assert.deepEqual(targets(), ['Sheet1!A5']);
    assert.equal(h.extentReads.length, 0);
    await click(field('Close workbook tools'));
    await go('A1:B2');
    await preview('last-cell');
    assert.deepEqual(targets(), ['Sheet1!C6']);
    assert.equal(h.extentReads.length, 1);
    const before = structuredClone(h.data);
    await click(button('Select matching cells'));
    assert.equal(field('Go to cell').placeholder, 'C6');
    h.pure(before);
  }
);
check('Last Cell handles a huge worksheet extent without reading cells during Preview', async (t) => {
  const h = await mount(t);
  h.setReply(completeUsedRange('A1:XFD1048576'));
  await go('A1:B2');
  await click(command('home.special-open'));
  await change(field('Selection type'), 'last-cell');
  h.reads.length = 0;
  await click(button('Preview matching cells'));
  await settled();
  assert.equal(alert(), undefined);
  assert.deepEqual(targets(), ['Sheet1!XFD1048576']);
  assert.deepEqual(h.reads, []);
  assert.match(field('Special selection summary').textContent, /Source positions1Matching cells1/);
  assert.deepEqual(h.accepted, []);
  assert.deepEqual(h.history, []);
});
check('Raw merged Last Cell expands one full area with anchor cursor and no workbook mutation', async (t) => {
  const h = await mount(t);
  for (const key of Object.keys(h.data)) delete h.data[key];
  h.data.C10 = { value: 9, dataType: 'number' };
  h.setReply(completeUsedRange('C8:I10', { merges: 'G8:I10' }));
  await h.render({
    sheets: h.input().sheets.map((sheet) => ({ ...sheet, cells: structuredClone(h.data), merges: ['G8:I10'] })),
  });
  await go('G8:I10');
  const before = structuredClone(h.data);
  await preview();
  assert.deepEqual(targets(), ['Sheet1!C10']);
  await click(field('Close workbook tools'));
  await go('A1:B2');
  await preview('last-cell');
  assert.deepEqual(
    targets(),
    ['G8', 'H8', 'I8', 'G9', 'H9', 'I9', 'G10', 'H10', 'I10'].map((ref) => 'Sheet1!' + ref)
  );
  const focus = document.activeElement;
  await click(button('Select matching cells'));
  assert.equal(field('Go to cell').placeholder, 'G8');
  assert.equal(document.activeElement, focus);
  h.pure(before);
});
check(
  'Missing, incomplete, foreign, stale-head and wrong-sheet provider authority cannot select a fallback display extent',
  async (t) => {
    const bad = [
      undefined,
      incompleteUsedRange(),
      completeUsedRange('A1:C6', { receipt: { ...providerReceipt, ownerId: 'foreign' } }),
      completeUsedRange('A1:C6', { receipt: { ...providerReceipt, state: 'retired' } }),
      completeUsedRange('A1:C6', { sheet: 'Other' }),
    ];
    for (const reply of bad) {
      const h = await mount(t);
      if (reply) h.setReply(reply);
      else await h.render({ usedRangeControls: undefined });
      await go('A5');
      const before = structuredClone(h.data);
      await preview();
      assert.match(alert(), /UsedRange|owner|retired|source\/version/);
      assert.equal(field('Special selection targets'), null);
      assert.equal(field('Go to cell').placeholder, 'A5');
      h.pure(before);
      await h.unmount();
    }
  }
);
check(
  'Save-like provider revision change at unchanged public head retires old Preview and retained Apply handler',
  async (t) => {
    const h = await mount(t);
    await go('A5');
    const head = h.input().storageKey,
      before = structuredClone(h.data);
    await preview();
    const stale = props(button('Select matching cells')).onClick,
      receipt = { ...providerReceipt, revision: 'post-save-view-revision' };
    h.setReply(completeUsedRange('A1:C6', { receipt, version: 0, sourceRevision: 'b'.repeat(64) }));
    await h.render({ usedRangeControls: { receipt, read: h.extent } });
    assert.equal(h.input().storageKey, head);
    assert.equal(field('Special selection targets'), null);
    await act(async () => stale());
    await flush();
    assert.equal(field('Go to cell').placeholder, 'A5');
    assert.equal(h.extentReads.length, 1);
    h.pure(before);
    await click(field('Close workbook tools'));
    await preview();
    assert.deepEqual(targets(), ['Sheet1!A1', 'Sheet1!C4', 'Sheet1!A5', 'Sheet1!C5', 'Sheet1!C6']);
    h.pure(before);
  }
);
check(
  'Apply re-reads native identity and refuses a changed source/version despite identical target cells',
  async (t) => {
    const h = await mount(t);
    await go('A5');
    const before = structuredClone(h.data);
    await preview();
    h.setReply(completeUsedRange('A1:C6', { version: 3, sourceRevision: 'b'.repeat(64) }));
    await click(button('Select matching cells'));
    assert.match(alert(), /matching cells changed/);
    assert.equal(field('Go to cell').placeholder, 'A5');
    assert.equal(h.extentReads.length, 2);
    h.pure(before);
  }
);
check('Cancel joins a held UsedRange query and ignores its late result before a new Preview', async (t) => {
  const h = await mount(t),
    held = defer();
  let calls = 0;
  h.holdExtent(() => (++calls === 1 ? held.promise : Promise.resolve(completeUsedRange())));
  await go('A5');
  await click(command('home.special-open'));
  h.reads.length = 0;
  await click(button('Preview matching cells'));
  assert.equal(calls, 1);
  await click(button('Cancel selection'));
  assert(button('Preview matching cells').disabled);
  await act(async () => props(button('Preview matching cells')).onClick());
  assert.equal(calls, 1);
  await act(async () => held.resolve(completeUsedRange()));
  await flush();
  assert.equal(field('Special selection targets'), null);
  assert.deepEqual(h.reads, []);
  await click(button('Preview matching cells'));
  await settled();
  assert.equal(calls, 2);
  assert.equal(targets().length, 5);
  assert.deepEqual(h.accepted, []);
  assert.deepEqual(h.history, []);
});
check('Late provider query after Save retirement cannot surface stale cells, errors or acknowledgements', async (t) => {
  const h = await mount(t),
    held = defer();
  h.holdExtent(() => held.promise);
  await go('A5');
  const before = structuredClone(h.data);
  await click(command('home.special-open'));
  await click(button('Preview matching cells'));
  const receipt = { ...providerReceipt, revision: 'new-saved-source' };
  let replacementCalls = 0;
  await h.render({
    usedRangeControls: {
      receipt,
      read: async () => {
        replacementCalls++;
        return completeUsedRange('A1:C6', { receipt });
      },
    },
  });
  await act(async () => held.reject(Error('Retired source failure')));
  await flush();
  assert.equal(field('Special selection targets'), null);
  assert(!document.body.textContent.includes('Retired source failure'));
  assert.equal(replacementCalls, 0);
  assert.equal(field('Go to cell').placeholder, 'A5');
  h.pure(before);
});
check('Hiding the pane while Apply rereads extent cannot commit or show hidden selection success', async (t) => {
  const h = await mount(t);
  await go('A5');
  const before = structuredClone(h.data);
  await preview();
  const held = defer();
  h.holdExtent(() => held.promise);
  await click(button('Select matching cells'));
  await click(field('Close workbook tools'));
  assert(pane().closest('[hidden]'));
  await act(async () => held.resolve(completeUsedRange()));
  await flush();
  assert.equal(field('Go to cell').placeholder, 'A5');
  assert(!pane().textContent.includes('Matching cells selected.'));
  h.pure(before);
});
check(
  '1004 default UsedRange targets page without more provider queries and Apply remains selection-only',
  async (t) => {
    const h = await mount(t);
    for (const key of Object.keys(h.data)) delete h.data[key];
    Object.assign(
      h.data,
      Object.fromEntries(Array.from({ length: 1004 }, (_, i) => ['A' + (i + 1), { value: i, dataType: 'number' }]))
    );
    h.setReply(completeUsedRange('A1:A1004'));
    await h.render({ sheets: h.input().sheets.map((sheet) => ({ ...sheet, rows: 1004, cells: {} })) });
    await go('A5');
    const before = structuredClone(h.data);
    await preview();
    const refs = [];
    let pages = 0;
    do {
      refs.push(...targets());
      pages++;
      if (button('More matching cells').disabled) break;
      await click(button('More matching cells'));
    } while (true);
    assert.equal(pages, 11);
    assert.equal(refs.length, 1004);
    assert.equal(new Set(refs).size, 1004);
    assert.equal(refs.at(-1), 'Sheet1!A1004');
    assert.equal(h.extentReads.length, 1);
    await click(button('Select matching cells'));
    assert.equal(h.extentReads.length, 2);
    assert.equal(field('Go to cell').placeholder, 'A1');
    h.pure(before);
  }
);
