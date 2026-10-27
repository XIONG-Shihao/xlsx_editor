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
  runs = await fs.mkdtemp(path.join(component, '.build/tests/trace-grid-ui-'));
const entry = path.join(runs, 'entry.ts'),
  bundle = path.join(runs, 'test.mjs');
await fs.writeFile(
  entry,
  `export {SpreadsheetGrid} from ${JSON.stringify(path.join(component, 'ui/SpreadsheetGrid.tsx'))};export {configureXlsxHost} from ${JSON.stringify(path.join(component, 'ui/host.tsx'))};export {traceBounds,tracePosition} from ${JSON.stringify(path.join(component, 'ui/formulaTrace.ts'))};`
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
      return 600;
    },
  },
  clientHeight: {
    get() {
      return 400;
    },
  },
});
globalThis.ResizeObserver = class {
  constructor(fn) {
    this.fn = fn;
  }
  observe() {
    this.fn([{ contentRect: { width: 600, height: 400 } }]);
  }
  disconnect() {}
};
globalThis.requestAnimationFrame = () => 1;
globalThis.cancelAnimationFrame = () => {};
const { createRoot } = await import('react-dom/client'),
  { SpreadsheetGrid, configureXlsxHost, traceBounds, tracePosition } = await import(pathToFileURL(bundle));
configureXlsxHost({
  api: { copyText: async () => {}, readClipboardText: async () => '' },
  Icon: () => React.createElement('svg'),
  useI18n: () => ({ locale: 'en' }),
});
const check = (name, fn) => {
  test(name, async () => {
    await fn();
  });
};
after(async () => {
  dom.window.close();
});
const button = (name) =>
  [...document.querySelectorAll('button')].find(
    (n) => !n.closest('[hidden]') && (n.getAttribute('aria-label') ?? n.textContent) === name
  );
async function click(node) {
  assert(node);
  assert.equal(node.disabled, false);
  await act(async () => node.click());
}
const tab = (name) => [...document.querySelectorAll('[role="tab"]')].find((n) => n.textContent === name);
const command = (id) =>
  document.querySelector('.sheet-ribbon-tabpanel:not([hidden]) [data-ribbon-command="' + id + '"]');
const defer = () => {
  let resolve;
  const promise = new Promise((yes) => (resolve = yes));
  return { promise, resolve };
};
async function go(ref) {
  const input = document.querySelector('[aria-label="Go to cell"]');
  await act(async () => {
    Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(input, ref);
    input.dispatchEvent(new window.Event('input', { bubbles: true }));
  });
  await act(async () =>
    input.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
  );
}
const baseCells = {
  A1: { value: 1 },
  B1: { formula: 'A1', value: 1, calculation: { status: 'calculated', value: 1 } },
  C1: { formula: 'B1', value: 1, calculation: { status: 'calculated', value: 1 } },
  D1: { value: 4 },
};
const sheet = (name) => ({
  name,
  state: 'visible',
  path: 'xl/worksheets/' + (name === 'Sheet1' ? 'sheet1' : 'sheet2') + '.xml',
  rows: 100,
  columns: 12,
  cells: { ...baseCells },
  merges: [],
  widths: [],
  hiddenRows: [],
  rowHeights: [],
  freeze: { rows: 0, columns: 0 },
});
const edge = (a, b, s = 'Sheet1') => ({
  from: { sheet: s, range: a },
  to: { sheet: s, range: b },
  kind: 'reference',
  via: 'formula',
});
async function mount(extra = {}) {
  const root = createRoot(document.getElementById('root')),
    calls = [],
    writes = [],
    sheets = [sheet('Sheet1'), sheet('Sheet2')];
  let state = 'real-head-A',
    owner = 'indexed-A',
    doc = 'grant-A',
    active = true,
    callback = extra.callback,
    receiptPatch = {};
  const render = async (patch) => {
    if (patch) {
      state = patch.state ?? state;
      owner = patch.owner ?? owner;
      doc = patch.doc ?? doc;
      active = patch.active ?? active;
      receiptPatch = { ...receiptPatch, ...patch.receipt };
    }
    const receipt = { documentId: doc, ownerId: owner, state, revision: 'source-sha-A', ...receiptPatch };
    const trace = async (name, cell, direction, after, limit) => {
      const request = { ...receipt, sheet: name, cell, direction, after, limit };
      calls.push(request);
      const edges = callback
        ? await callback(request)
        : direction === 'precedents'
          ? cell === 'C1'
            ? [edge('B1', 'C1', name)]
            : cell === 'B1'
              ? [edge('A1', 'B1', name)]
              : []
          : cell === 'A1'
            ? [edge('A1', 'B1', name)]
            : cell === 'B1'
              ? [edge('B1', 'C1', name)]
              : [];
      return {
        schema: 1,
        ...receipt,
        sheet: name,
        cell,
        direction,
        edges,
        markers: [],
        nextCursor: null,
        coverage: { mode: 'direct-static-references', complete: true, scannedFormulas: 1, pages: 1, limited: false },
      };
    };
    await act(async () =>
      root.render(
        React.createElement(SpreadsheetGrid, {
          sheets,
          documentId: doc,
          ownerIdentity: owner,
          storageKey: 'opaque-ui/' + owner + '/' + state,
          active,
          readOnly: extra.readOnly ?? false,
          readRange: async (_name, range) => {
            const b = traceBounds(range);
            return Object.fromEntries(
              Object.entries(baseCells).filter(([ref]) => {
                const p = tracePosition(ref);
                return p.row >= b.first.row && p.row <= b.last.row && p.col >= b.first.col && p.col <= b.last.col;
              })
            );
          },
          traceControls: { receipt, trace },
          onEdit: (x) => writes.push(x),
          onBatch: (x) => writes.push(x),
          onCommandBatch: async (x) => writes.push(x),
        })
      )
    );
  };
  await render();
  return {
    calls,
    writes,
    render,
    storageKey: () => 'opaque-ui/' + owner + '/' + state,
    close: async () => {
      await act(async () => root.unmount());
      await new Promise((resolve) => setTimeout(resolve, 2));
    },
  };
}
check('actual ribbon repeated trace expands without toggling away; one dock and one copy of each command', async () => {
  const ui = await mount();
  try {
    await go('C1');
    await click(tab('Formulas'));
    assert.equal(
      document.querySelectorAll(
        '.sheet-ribbon-tabpanel:not([hidden]) [data-ribbon-command="formulas.trace-precedents"]'
      ).length,
      1
    );
    await click(command('formulas.trace-precedents'));
    assert.equal(document.querySelectorAll('aside.sheet-task-pane:not([hidden])').length, 1);
    assert.equal(document.querySelectorAll('[data-trace-arrow]').length, 1);
    await click(command('formulas.trace-precedents'));
    assert.equal(document.querySelectorAll('aside.sheet-task-pane:not([hidden])').length, 1);
    assert.equal(document.querySelectorAll('[data-trace-arrow]').length, 2);
    assert.deepEqual(
      ui.calls.map((x) => x.cell),
      ['C1', 'B1']
    );
    assert.equal(document.querySelector('[aria-label="Formula tracing"]').getAttribute('data-trace-root'), 'Sheet1!C1');
    await click(command('formulas.remove-arrows'));
    assert.equal(document.querySelector('[data-trace-arrow]'), null);
    assert.deepEqual(ui.writes, []);
  } finally {
    await ui.close();
  }
});
check('same-sheet Go keeps original graph/root; explicit new ribbon Trace chooses current selection', async () => {
  const ui = await mount();
  try {
    await go('C1');
    await click(tab('Formulas'));
    await click(command('formulas.trace-precedents'));
    await click(button('Go to trace source Sheet1!B1'));
    assert.equal(document.querySelector('[aria-label="Formula tracing"]').getAttribute('data-trace-root'), 'Sheet1!C1');
    assert.equal(document.querySelectorAll('[data-trace-arrow]').length, 1);
    await click(command('formulas.trace-precedents'));
    assert.equal(document.querySelector('[aria-label="Formula tracing"]').getAttribute('data-trace-root'), 'Sheet1!B1');
    assert.deepEqual(
      ui.calls.map((x) => x.cell),
      ['C1', 'B1']
    );
    assert.equal(document.querySelectorAll('[data-trace-arrow]').length, 1);
    assert.deepEqual(ui.writes, []);
  } finally {
    await ui.close();
  }
});
check('head owner inactive and worksheet changes mask existing graph; read-only trace remains usable', async () => {
  const ui = await mount({ readOnly: true });
  try {
    await go('C1');
    await click(tab('Formulas'));
    await click(command('formulas.trace-precedents'));
    assert.equal(document.querySelectorAll('[data-trace-arrow]').length, 1);
    await ui.render({ state: 'real-head-B' });
    assert.equal(document.querySelector('[data-trace-arrow]'), null);
    assert.equal(document.querySelector('[data-trace-edge]'), null);
    await click(command('formulas.trace-precedents'));
    await ui.render({ owner: 'indexed-B' });
    assert.equal(document.querySelector('[data-trace-arrow]'), null);
    await click(command('formulas.trace-precedents'));
    await ui.render({ active: false });
    assert.equal(document.querySelector('[data-trace-arrow]'), null);
    await ui.render({ active: true });
    await click(command('formulas.trace-precedents'));
    await click([...document.querySelectorAll('.sheet-tabs button')].find((x) => x.textContent === 'Sheet2'));
    assert.equal(document.querySelector('[data-trace-arrow]'), null);
    assert.equal(document.querySelector('[data-trace-edge]'), null);
    assert.deepEqual(ui.writes, []);
  } finally {
    await ui.close();
  }
});
check(
  'retained pane can be hidden without cancelling accepted same-sheet arrows and overlay is outside table extent',
  async () => {
    const ui = await mount();
    try {
      await go('C1');
      await click(tab('Formulas'));
      const table = document.querySelector('table.sheet-grid'),
        width = table.style.width;
      await click(command('formulas.trace-precedents'));
      await click(button('Close workbook tools'));
      assert.equal(document.querySelectorAll('[data-trace-arrow]').length, 1);
      const overlay = document.querySelector('.sheet-trace-overlay'),
        scroll = document.querySelector('.sheet-grid-scroll'),
        wrapper = document.querySelector('.sheet-grid-viewport');
      assert.equal(overlay.parentElement, wrapper);
      assert.equal(scroll.parentElement, wrapper);
      assert.equal(overlay.style.position, 'absolute');
      assert.equal(table.style.width, width);
      assert(!scroll.contains(overlay));
      assert.equal(document.querySelectorAll('table.sheet-grid').length, 1);
      assert.deepEqual(ui.writes, []);
    } finally {
      await ui.close();
    }
  }
);
check('held first reply survives pane mount cleanup but cannot adopt after accepted head change', async () => {
  const gate = defer();
  let count = 0;
  const ui = await mount({
    callback: async (r) => {
      count++;
      if (count === 1) await gate.promise;
      return [edge('A1', r.cell, r.sheet)];
    },
  });
  try {
    await go('C1');
    await click(tab('Formulas'));
    await click(command('formulas.trace-precedents'));
    assert.equal(ui.calls.length, 1);
    assert.equal(button('Cancel trace')?.disabled, false);
    await ui.render({ state: 'real-head-B' });
    assert.equal(document.querySelector('[data-trace-arrow]'), null);
    await act(async () => gate.resolve());
    assert.equal(document.querySelector('[data-trace-arrow]'), null);
    assert.equal(ui.calls.length, 1);
    await click(command('formulas.trace-precedents'));
    assert.equal(ui.calls[1].state, 'real-head-B');
    assert.equal(document.querySelectorAll('[data-trace-arrow]').length, 1);
    assert.deepEqual(ui.writes, []);
  } finally {
    await ui.close();
  }
});

check(
  'actual Grid masks on every backend receipt field change with unchanged UI storage key and refuses old navigation',
  async () => {
    const ui = await mount();
    try {
      await go('C1');
      await click(tab('Formulas'));
      const key = ui.storageKey();
      for (const receipt of [
        { revision: 'source-sha-B' },
        { ownerId: 'indexed-owner-B' },
        { documentId: 'grant-B' },
        { state: 'real-head-B' },
      ]) {
        await click(command('formulas.trace-precedents'));
        assert.equal(document.querySelectorAll('[data-trace-arrow]').length, 1);
        const oldToken = document.querySelector('[aria-label="Formula tracing"]').getAttribute('data-trace-state'),
          node = button('Go to trace source Sheet1!B1'),
          stale = node[Object.keys(node).find((k) => k.startsWith('__reactProps$'))].onClick,
          count = ui.calls.length;
        await ui.render({ receipt });
        assert.equal(ui.storageKey(), key);
        assert.equal(
          document.querySelector('[aria-label="Formula tracing"]').getAttribute('data-trace-owner'),
          'indexed-A'
        );
        assert.notEqual(
          document.querySelector('[aria-label="Formula tracing"]').getAttribute('data-trace-state'),
          oldToken
        );
        assert.equal(document.querySelector('[data-trace-arrow]'), null);
        assert.equal(document.querySelector('[data-trace-edge]'), null);
        assert.equal(ui.calls.length, count);
        stale();
        assert.equal(document.querySelector('[aria-label="Go to cell"]').getAttribute('placeholder'), 'C1');
      }
      assert.deepEqual(ui.writes, []);
    } finally {
      await ui.close();
    }
  }
);
