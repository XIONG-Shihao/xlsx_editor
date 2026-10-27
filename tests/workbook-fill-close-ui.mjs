import './support/scratch-dir.mjs';
/** Close-focus regression: actual Grid code, controlled owner + geometry only. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
import { rangeBounds } from '../core/workbook-commands.mjs';
const component = fileURLToPath(new URL('../', import.meta.url)),
  output = await fs.mkdtemp(path.join(component, '.build/tests/fill-close-ui-')),
  checks = [];
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
  field = (name) =>
    [...document.querySelectorAll('[aria-label]')].find(
      (n) => n.getAttribute('aria-label') === name && !n.closest('[hidden]')
    ),
  button = (name) =>
    [...document.querySelectorAll('button')].find(
      (n) => (n.getAttribute('aria-label') ?? n.textContent) === name && !n.closest('[hidden]')
    ),
  opener = () => document.querySelector('[data-home-command="home.fill-series"]'),
  flush = () => act(async () => new Promise((resolve) => setTimeout(resolve, 16)));
const defer = () => {
  let resolve;
  const promise = new Promise((yes) => (resolve = yes));
  return { promise, resolve };
};
async function click(n) {
  n = typeof n === 'string' ? button(n) : n;
  assert(n);
  assert(!n.disabled);
  await act(async () => n.click());
  await flush();
}
async function focusClick(n) {
  assert(n);
  await act(async () => {
    if (n.dispatchEvent(new window.MouseEvent('pointerdown', { bubbles: true, cancelable: true }))) n.focus();
    n.click();
  });
  await flush();
}
async function go(value) {
  const n = field('Go to cell');
  await act(async () => {
    props(n).onChange({ target: { value } });
  });
  await flush();
  await act(async () =>
    n.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
  );
  await flush();
}
async function mount(t) {
  const root = createRoot(document.getElementById('root')),
    held = defer(),
    accepted = [],
    legacy = [],
    data = {
      E1: { value: 1, rawValue: '1' },
      F1: { value: 3, rawValue: '3' },
      G1: { value: null },
      H1: { value: null },
    };
  const sheet = {
    name: 'Sheet1',
    path: 'xl/worksheets/sheet1.xml',
    state: 'visible',
    rows: 30,
    columns: 12,
    cells: structuredClone(data),
    merges: [],
    widths: [],
    hiddenRows: [],
    rowHeights: [],
    freeze: { rows: 0, columns: 0 },
  };
  let input = {
    sheets: [sheet],
    documentId: 'original-doc',
    ownerIdentity: 'original-owner',
    storageKey: 'original-head',
    active: true,
    readOnly: false,
    onEdit: (value) => legacy.push(value),
    onBatch: (value) => legacy.push(value),
    readRange: async (_name, range) => {
      const b = rangeBounds(range);
      return Object.fromEntries(
        Object.entries(data).filter(([ref]) => {
          const p = rangeBounds(ref).first;
          return p.row >= b.first.row && p.row <= b.last.row && p.col >= b.first.col && p.col <= b.last.col;
        })
      );
    },
    onCommandBatch: (commands) => {
      accepted.push(structuredClone(commands));
      return held.promise;
    },
  };
  const render = async (patch) => {
    input = { ...input, ...patch };
    await act(async () => root.render(React.createElement(ui.SpreadsheetGrid, input)));
    await flush();
  };
  await render({});
  t.after(async () => {
    await act(async () => root.unmount());
  });
  return { render, held, accepted, legacy };
}
const check = (name, fn) => {
  test(name, { timeout: 20000 }, async (t) => {
    let diagnostics;
    try {
      diagnostics = await fn(t);
    } catch (error) {
      await fs.writeFile(path.join(output, 'failure-' + checks.length + '.html'), document.documentElement.outerHTML);
      throw error;
    }
  });
};
after(async () => {
  dom.window.close();
});
async function acceptedClose(t, simulateFalsePositive) {
  const h = await mount(t);
  await go('E1:H1');
  const input = field('Formula or value');
  await act(async () => {
    input.focus();
    props(input).onChange({ target: { value: '=SUM(' } });
  });
  input.setSelectionRange(3, 5);
  await focusClick([...document.querySelectorAll('[role=tab]')].find((n) => n.textContent === 'Home'));
  await focusClick(opener().closest('details').querySelector('summary'));
  await focusClick(opener());
  await click('Preview series');
  assert(field('Series summary'));
  assert.deepEqual(h.accepted, []);
  assert.deepEqual(h.legacy, []);
  await click('Apply series');
  assert.deepEqual(h.accepted, [[{ type: 'fillSeries', sheet: 'Sheet1', range: 'E1:H1' }]]);
  await h.render({ storageKey: 'accepted-head' });
  await act(async () => h.held.resolve());
  await flush();
  assert.match(field('Fill series').textContent, /Series filled/);
  assert.equal(input.value, '1', 'Accepted head must retire the old draft, not rebase it.');
  const trigger = opener(),
    details = trigger.closest('details');
  details.removeAttribute('open');
  let refusedFocusAttempts = 0;
  if (simulateFalsePositive) {
    trigger.getClientRects = () => [{ x: 10, y: 10, width: 100, height: 30 }];
    trigger.focus = () => {
      refusedFocusAttempts++;
    };
  }
  const close = button('Close workbook tools');
  close.focus();
  await click(close);
  await flush();
  const active = document.activeElement,
    diagnostics = {
      simulatedNonemptyRectFocusRefusal: simulateFalsePositive,
      refusedFocusAttempts,
      detailsOpen: details.hasAttribute('open'),
      activeElement: {
        tag: active?.tagName,
        ariaLabel: active?.getAttribute('aria-label'),
        homeCommand: active?.getAttribute('data-home-command'),
        workbookTool: active?.getAttribute('data-workbook-tool'),
      },
      originalAcceptedCommand: h.accepted[0],
    };
  assert(
    active?.matches('.sheet-grid input,.sheet-cell-reference'),
    'Close must focus the current grid or Go field after a closed trigger refuses focus.'
  );
  assert.equal(input.value, '1');
  assert.notEqual(document.activeElement, trigger);
  assert.deepEqual(h.accepted, [[{ type: 'fillSeries', sheet: 'Sheet1', range: 'E1:H1' }]]);
  assert.deepEqual(h.legacy, []);
  return diagnostics;
}
check(
  'accepted-head Close retires the parked old draft and follows current grid/Go fallback with ordinary JSDOM empty rectangles',
  (t) => acceptedClose(t, false)
);
check(
  'closed-details trigger with nonempty rectangle but refusing focus still falls back after accepted-head draft retirement',
  (t) => acceptedClose(t, true)
);
