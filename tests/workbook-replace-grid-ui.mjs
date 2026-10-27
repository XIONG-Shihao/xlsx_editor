import './support/scratch-dir.mjs';
/** Current Grid/Home/Replace integration; controlled owner reads and acknowledgements. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createHash } from 'node:crypto';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
import { rangeBounds } from '../core/workbook-commands.mjs';
const component = fileURLToPath(new URL('../', import.meta.url)),
  output = await fs.mkdtemp(path.join(component, '.build/tests/replace-grid-ui-')),
  checks = [];
const files = [
  'ui/SpreadsheetGrid.tsx',
  'ui/SpreadsheetRibbon.tsx',
  'ui/WorkbookReplace.tsx',
  'ui/replacePreview.ts',
  'ui/homeCommands.ts',
  'ui/ribbonCommands.ts',
  'ui/host.tsx',
  'ui/words.ts',
  'ui/styles/01.css',
  'core/contracts.ts',
  'core/workbook-commands.mjs',
  'tests/workbook-replace-grid-ui.mjs',
];
const pins = () =>
    Promise.all(
      files.map(async (relative) => {
        const p = path.join(component, relative),
          b = await fs.readFile(p);
        return { relative, path: p, sha256: createHash('sha256').update(b).digest('hex'), bytes: b.length };
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
  field = (name) =>
    [...document.querySelectorAll('[aria-label]')].find(
      (node) => node.getAttribute('aria-label') === name && !node.closest('[hidden]')
    ),
  node = (id) => [...document.querySelectorAll(`[data-home-command="${id}"]`)].find((n) => !n.closest('[hidden]')),
  button = (name) =>
    [...document.querySelectorAll('button')].find(
      (n) => (n.getAttribute('aria-label') ?? n.textContent) === name && !n.closest('[hidden]')
    ),
  flush = async (ms = 6) => act(async () => new Promise((resolve) => setTimeout(resolve, ms))),
  // Waits until the screen has done something, however loaded the machine is.
  until = async (done, ms = 10000) => {
    for (const end = Date.now() + ms; !done() && Date.now() < end; ) await flush(20);
  };
const defer = () => {
  let resolve, reject;
  const promise = new Promise((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
};
async function click(n) {
  n = typeof n === 'string' ? button(n) : n;
  assert(n);
  assert(!n.disabled);
  await act(async () => n.click());
  await flush();
  return n;
}
/** JSDOM has no pointer default focus. Explicitly execute production capture,
 * then model the focus transfer that capture permits, before the click. */
async function focusClick(n) {
  assert(n);
  assert(!n.disabled);
  await act(async () => {
    const event = new window.MouseEvent('pointerdown', { bubbles: true, cancelable: true });
    if (n.dispatchEvent(event)) n.focus();
    n.click();
  });
  await flush();
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
  await click(node('home.find-open'));
  await change(field('Find'), 'TOKEN');
  await change(field('Replace with'), 'VALUE');
  await click(node('home.replace-preview'));
  assert(field('Replacement preview'));
  assert.equal(document.querySelectorAll('aside.sheet-task-pane:not([hidden])').length, 1);
}
async function preview() {
  await click('Preview all replacements');
  for (
    let i = 0;
    i < 100 && !field('Replacement targets') && !field('Replacement preview')?.querySelector('[role="alert"]');
    i++
  )
    await flush(10);
  assert(field('Replacement targets'), field('Replacement preview')?.textContent);
}
const rows = () =>
  [...field('Replacement targets').querySelectorAll('tbody tr')].map((n) => ({
    cell: n.cells[0].textContent,
    before: n.cells[1].textContent,
    after: n.cells[2].textContent,
  }));
function selectedCells(data, range) {
  const b = rangeBounds(range),
    cells = {};
  for (let r = b.first.row; r <= b.last.row; r++)
    for (let c = b.first.col; c <= b.last.col; c++) {
      let column = '',
        value = c + 1;
      while (value) {
        column = String.fromCharCode(65 + ((value - 1) % 26)) + column;
        value = Math.floor((value - 1) / 26);
      }
      const ref = column + (r + 1);
      if (data[ref] !== undefined) cells[ref] = structuredClone(data[ref]);
    }
  return cells;
}
async function mount(t, patch = {}) {
  const root = createRoot(document.getElementById('root')),
    accepted = [],
    legacy = [],
    reads = [],
    queries = [],
    data = {};
  for (let n = 1; n <= 1004; n++) data['A' + n] = { value: 'TOKEN ' + n, dataType: 'text', rawValue: 'TOKEN ' + n };
  data.C1 = { value: '=TOKEN(A1)', dataType: 'text', rawValue: '=TOKEN(A1)' };
  data.D1 = { value: 3, formula: 'SUM(1,2)', rawValue: '3', calculation: { status: 'calculated', value: 3 } };
  data.P20 = { value: 'Outside' };
  const sheet = (name) => ({
    name,
    path: 'xl/worksheets/' + name + '.xml',
    state: 'visible',
    rows: 1100,
    columns: 20,
    cells: structuredClone(data),
    merges: [],
    widths: [],
    hiddenRows: [],
    rowHeights: [],
    freeze: { rows: 0, columns: 0 },
  });
  let override = null,
    queryOverride = null;
  const query = async (sheet, text, mode, after, options) => {
    queries.push({ sheet, text, mode, after, options });
    if (queryOverride) return queryOverride(sheet, text, mode, after, options);
    const entries = Object.entries(
      sheet === 'Other' ? { A1: { value: 'TOKEN Other', rawValue: 'TOKEN Other', dataType: 'text' } } : data
    )
      .filter(([ref, value]) => {
        const content =
          mode === 'formula' ? (value.formula ? '=' + value.formula : '') : String(value.rawValue ?? value.value ?? '');
        return (
          content.toLowerCase().includes(text.toLowerCase()) &&
          (!after ||
            rangeBounds(ref).first.row > rangeBounds(after).first.row ||
            (rangeBounds(ref).first.row === rangeBounds(after).first.row &&
              rangeBounds(ref).first.col > rangeBounds(after).first.col))
        );
      })
      .sort(([a], [b]) => {
        const aa = rangeBounds(a).first,
          bb = rangeBounds(b).first;
        return aa.row - bb.row || aa.col - bb.col;
      });
    const page = entries.slice(0, options?.matchLimit ?? 1000);
    return {
      cells: Object.fromEntries(page.map(([ref, value]) => [ref, structuredClone(value)])),
      matches: page.length,
      nextCursor: entries.length > page.length ? page.at(-1)[0] : null,
    };
  };
  const reader = async (sheet, range) => {
    reads.push({ sheet, range });
    return override ? override(sheet, range) : selectedCells(data, range);
  };
  let input = {
      sheets: [sheet('Sheet1'), sheet('Other')],
      documentId: 'doc',
      ownerIdentity: 'original-owner',
      storageKey: 'original-head',
      active: true,
      readOnly: false,
      readRange: reader,
      query,
      onEdit: (value) => legacy.push(value),
      onBatch: (value) => legacy.push(value),
      onCommandBatch: async (commands) => accepted.push(structuredClone(commands)),
      ...patch,
    },
    mounted = true;
  const render = async (patch) => {
    input = { ...input, ...patch };
    await act(async () => root.render(React.createElement(ui.SpreadsheetGrid, input)));
    await flush();
  };
  const unmount = async () => {
    if (!mounted) return;
    mounted = false;
    await act(async () => root.unmount());
    await flush();
  };
  await render({});
  t.after(unmount);
  return {
    accepted,
    legacy,
    reads,
    queries,
    data,
    reader,
    render,
    unmount,
    input: () => input,
    holdRead: (fn) => (override = fn),
    holdQuery: (fn) => (queryOverride = fn),
  };
}
const check = (name, fn) => {
  test(name, { timeout: 25000 }, async (t) => {
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
  'Complete all-sheet Preview includes unloaded tail and second sheet; pages and Cancel never submit edits',
  async (t) => {
    const h = await mount(t);
    await open();
    await change(field('Search scope'), 'workbook');
    await click(node('home.replace-preview'));
    assert.equal(h.queries.length, 0);
    assert.equal(field('Replacement targets'), undefined);
    await preview();
    assert.equal(field('Replacement cells shown').textContent, '1–100 of 1006');
    assert(h.queries.every((q) => q.options.matchLimit === 100));
    const reads = h.queries.length,
      all = [];
    while (true) {
      all.push(...rows());
      if (button('More replacement cells').disabled) break;
      await click('More replacement cells');
    }
    assert.equal(all.length, 1006);
    assert.equal(new Set(all.map((r) => r.cell)).size, 1006);
    assert(all.some((r) => r.cell === 'Sheet1!A1004' && r.after === 'VALUE 1004'));
    assert.deepEqual(all.at(-1), { cell: 'Other!A1', before: 'TOKEN Other', after: 'VALUE Other' });
    assert.equal(h.queries.length, reads);
    await click('Cancel replacements');
    assert.equal(field('Replacement summary'), undefined);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Apply re-collects complete original reader and uses one captured batch after same-binding callbacks change',
  async (t) => {
    const h = await mount(t),
      held = defer(),
      newQueries = [],
      newAccepted = [];
    await h.render({
      onCommandBatch: (commands) => {
        h.accepted.push(structuredClone(commands));
        return held.promise;
      },
    });
    await open();
    await change(field('Search scope'), 'workbook');
    await click(node('home.replace-preview'));
    await preview();
    const first = h.queries.length;
    await h.render({
      query: async (...args) => {
        newQueries.push(args);
        return { cells: {}, matches: 0, nextCursor: null };
      },
      onCommandBatch: async (commands) => newAccepted.push(commands),
    });
    const invoke = props(button('Apply replacements')).onClick;
    await act(async () => {
      invoke();
      invoke();
    });
    await until(() => h.accepted.length > 0);
    await flush(50);
    assert.equal(h.accepted.length, 1);
    assert.equal(h.accepted[0].length, 3);
    assert.equal(
      h.accepted[0].reduce((n, c) => n + c.cells.length, 0),
      1006
    );
    assert.equal(h.queries.length, first * 2);
    assert.deepEqual(newQueries, []);
    assert.deepEqual(newAccepted, []);
    assert(button('Cancel replacements').disabled);
    await h.render({ storageKey: 'accepted-head' });
    await act(async () => held.resolve());
    await flush();
    assert.match(field('Replacement preview').textContent, /Replacements applied. Use Undo/);
    assert.deepEqual(h.legacy, []);
  }
);
check('Changed late target with unchanged head invalidates reviewed plan before any submission', async (t) => {
  const h = await mount(t);
  await open();
  await preview();
  h.data.A1004.value = 'TOKEN changed';
  h.data.A1004.rawValue = 'TOKEN changed';
  await click('Apply replacements');
  for (let i = 0; i < 100 && !field('Replacement preview').querySelector('[role="alert"]'); i++) await flush(10);
  assert.deepEqual(h.accepted, []);
  assert.match(field('Replacement preview').textContent, /matching cells changed/);
});
check('Cancellation joins late original read and prevents replacement submission', async (t) => {
  const h = await mount(t),
    held = defer();
  await open();
  h.holdQuery(() => held.promise);
  await click('Preview all replacements');
  assert.equal(h.queries.length, 1);
  await click('Cancel replacements');
  assert(button('Preview all replacements').disabled);
  await act(async () =>
    held.resolve({ cells: { A1: { value: 'TOKEN 1', rawValue: 'TOKEN 1' } }, matches: 1, nextCursor: null })
  );
  await flush();
  assert.equal(field('Replacement targets'), undefined);
  assert.deepEqual(h.accepted, []);
  assert.equal(button('Preview all replacements').disabled, false);
});
check('Search draft, owner and head retirement mask late preview and never redirect an old acceptance', async (t) => {
  for (const kind of ['search', 'owner', 'head']) {
    const h = await mount(t),
      held = defer();
    await open();
    h.holdQuery(() => held.promise);
    await click('Preview all replacements');
    if (kind === 'search') await change(field('Replace with'), 'NEXT');
    else await h.render(kind === 'owner' ? { ownerIdentity: 'new-owner' } : { storageKey: 'new-head' });
    await act(async () => held.reject(Error('Late original search failure')));
    await flush();
    assert(!document.body.textContent.includes('Late original search failure'));
    assert.equal(field('Replacement targets'), undefined);
    assert.deepEqual(h.accepted, []);
    await h.unmount();
  }
});
check('Readonly supports full inspection while Apply and direct Replace remain disabled', async (t) => {
  const h = await mount(t, { readOnly: true });
  await open();
  await preview();
  assert(field('Replacement summary'));
  assert(button('Apply replacements').disabled);
  assert(node('home.replace-one').disabled);
  assert.deepEqual(h.accepted, []);
  await click('Close workbook tools');
  assert.equal(document.querySelectorAll('aside.sheet-task-pane:not([hidden])').length, 0);
});
check('Direct single replacement preserves formula-looking literal text and exact source preconditions', async (t) => {
  const h = await mount(t);
  await open();
  await click('Close workbook tools');
  await change(field('Find'), 'TOKEN(A1)');
  await change(field('Replace with'), 'SUM(A1)');
  await click(node('home.find-search'));
  await click(node('home.replace-one'));
  assert.deepEqual(h.accepted, [
    [
      {
        type: 'setCells',
        sheet: 'Sheet1',
        cells: [
          {
            cell: 'C1',
            expectedValue: '=TOKEN(A1)',
            expectedFormula: null,
            expectedRawValue: '=TOKEN(A1)',
            value: '=SUM(A1)',
          },
        ],
      },
    ],
  ]);
  assert.deepEqual(h.legacy, []);
});
check('Authored formula search includes equals prefix and keeps a formula-only replacement proposal', async (t) => {
  const h = await mount(t);
  await open();
  await click('Close workbook tools');
  await change(field('Search content'), 'formula');
  await change(field('Find'), '=SUM');
  await change(field('Replace with'), '=MAX');
  await click(node('home.replace-preview'));
  await preview();
  assert.deepEqual(rows(), [{ cell: 'Sheet1!D1', before: '=SUM(1,2)', after: '=MAX(1,2)' }]);
  await click('Apply replacements');
  assert.deepEqual(h.accepted, [
    [
      {
        type: 'setCells',
        sheet: 'Sheet1',
        cells: [
          { cell: 'D1', formula: 'MAX(1,2)', expectedFormula: 'SUM(1,2)', expectedValue: 3, expectedRawValue: '3' },
        ],
      },
    ],
  ]);
});
check('Same-head Close restores parked formula draft and caret with no mutation', async (t) => {
  const h = await mount(t);
  await click(node('home.find-open'));
  await change(field('Find'), 'TOKEN');
  await change(field('Replace with'), 'VALUE');
  const formula = field('Formula or value');
  await act(async () => {
    formula.focus();
    props(formula).onChange({ target: { value: '=SUM(' } });
  });
  formula.setSelectionRange(5, 5);
  await focusClick(node('home.replace-preview'));
  assert.equal(formula.value, '=SUM(');
  await click('Close workbook tools');
  assert.equal(document.activeElement, formula);
  assert.equal(formula.value, '=SUM(');
  assert.equal(formula.selectionStart, 5);
  assert.equal(formula.selectionEnd, 5);
  assert.deepEqual(h.accepted, []);
  assert.deepEqual(h.legacy, []);
});
