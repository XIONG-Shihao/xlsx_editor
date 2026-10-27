import './support/scratch-dir.mjs';
/** Current Grid/Home/Fill integration; controlled owner reads and acknowledgements. */
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
  output = await fs.mkdtemp(path.join(component, '.build/tests/fill-grid-ui-')),
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
      (node) => node.getAttribute('aria-label') === name && !node.closest('[hidden]')
    ),
  node = (id) => [...document.querySelectorAll(`[data-home-command="${id}"]`)].find((n) => !n.closest('[hidden]')),
  button = (name) =>
    [...document.querySelectorAll('button')].find(
      (n) => (n.getAttribute('aria-label') ?? n.textContent) === name && !n.closest('[hidden]')
    ),
  flush = async (ms = 6) => act(async () => new Promise((resolve) => setTimeout(resolve, ms)));
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
  await click(node('home.fill-series'));
  assert(field('Fill series'));
  assert.equal(document.querySelectorAll('aside.sheet-task-pane:not([hidden])').length, 1);
}
const seriesReads = (h) => h.reads.filter((r) => ['E1:H1', 'A1:A1000', 'A1001:A1004'].includes(r.range));
const rows = () =>
  [...field('Series destinations').querySelectorAll('tbody tr')].map((n) => ({
    cell: n.cells[0].textContent,
    value: Number(n.cells[1].textContent),
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
    data = {
      A1: { value: 1, rawValue: '1' },
      A2: { value: 3, rawValue: '3' },
      E1: { value: 1, rawValue: '1' },
      F1: { value: 3, rawValue: '3' },
      G1: { value: null, bold: true },
      H1: { value: null, bold: true },
      P20: { value: 'Outside' },
    };
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
  let override = null;
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
    data,
    reader,
    render,
    unmount,
    input: () => input,
    holdRead: (fn) => (override = fn),
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
  'Home Fill opener is pure and exact1004-position reads cover unloaded destinations; paging uses no new read',
  async (t) => {
    const h = await mount(t);
    await go('A1:A1004');
    await open();
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(seriesReads(h), []);
    await click('Preview series');
    assert.deepEqual(
      seriesReads(h).map((r) => r.range),
      ['A1:A1000', 'A1001:A1004']
    );
    assert(seriesReads(h).every((r) => rangeBounds(r.range).size <= 1000));
    assert.equal(rows().length, 100);
    assert.deepEqual(rows()[0], { cell: 'A3', value: 5 });
    assert.equal(field('Series cells shown').textContent, '1–100 of 1002');
    const all = [];
    while (true) {
      all.push(...rows());
      if (button('More series cells').disabled) break;
      await click('More series cells');
    }
    assert.equal(all.length, 1002);
    assert.equal(new Set(all.map((r) => r.cell)).size, 1002);
    assert.deepEqual(all.at(-1), { cell: 'A1004', value: 2007 });
    await click('Previous series cells');
    assert.equal(field('Series cells shown').textContent, '901–1000 of 1002');
    assert.equal(seriesReads(h).length, 2);
    await click('Cancel series');
    assert.equal(field('Series summary'), undefined);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Apply re-reads exact original range and awaits one original callback even after same-binding callback replacement',
  async (t) => {
    const h = await mount(t);
    await go('E1:H1');
    await open();
    await click('Preview series');
    assert.deepEqual(rows(), [
      { cell: 'G1', value: 5 },
      { cell: 'H1', value: 7 },
    ]);
    const held = defer(),
      newReads = [],
      newAccepted = [];
    const original = h.input().onCommandBatch;
    // Explicit fresh invocation captures this strict callback; then replacing its
    // same-binding props must not retarget the already reviewed context.
    await click('Close workbook tools');
    await h.render({
      onCommandBatch: (commands) => {
        original(commands);
        return held.promise;
      },
    });
    await go('E1:H1');
    await open();
    await click('Preview series');
    await h.render({
      readRange: async (...args) => {
        newReads.push(args);
        return {};
      },
      onCommandBatch: async (commands) => newAccepted.push(commands),
    });
    const submit = props(button('Apply series')).onClick;
    await act(async () => {
      submit();
      submit();
    });
    await flush();
    assert.deepEqual(h.accepted, [[{ type: 'fillSeries', sheet: 'Sheet1', range: 'E1:H1' }]]);
    assert.equal(h.reads.filter((r) => r.range === 'E1:H1').length, 3);
    assert.deepEqual(
      newReads.filter((args) => args[1] === 'E1:H1'),
      []
    );
    assert.deepEqual(newAccepted, []);
    assert(button('Cancel series').disabled);
    await h.render({ storageKey: 'accepted-head' });
    await act(async () => held.resolve());
    await flush();
    assert.match(field('Fill series').textContent, /Series filled. Use Undo/);
    assert.deepEqual(h.legacy, []);
  }
);
check('before-Apply re-read detects changed seeds or newly occupied targets and never sends a mutation', async (t) => {
  for (const kind of ['seed', 'target']) {
    const h = await mount(t);
    await go('E1:H1');
    await open();
    await click('Preview series');
    if (kind === 'seed') h.data.F1 = { value: 4, rawValue: '4' };
    else h.data.G1 = { value: 913, rawValue: '913' };
    await click('Apply series');
    assert.deepEqual(h.accepted, []);
    assert.match(
      field('Fill series').textContent,
      kind === 'seed' ? /selected cells changed.*Preview/i : /overwrite G1/
    );
    assert.equal(h.reads.filter((r) => r.range === 'E1:H1').length, 2);
    assert.deepEqual(h.legacy, []);
    await h.unmount();
  }
});
check(
  'readonly allows numeric Preview only; missing strict callback remains inspectable and literal/passive/inactive refuse',
  async (t) => {
    const h = await mount(t, { readOnly: true });
    await go('E1:H1');
    assert(!node('home.fill-series').disabled);
    await open();
    await click('Preview series');
    assert(button('Apply series').disabled);
    assert.equal(button('Apply series').title, 'This workbook is read-only.');
    assert.deepEqual(h.accepted, []);
    await click('Close workbook tools');
    await h.render({ readOnly: false, onCommandBatch: undefined });
    await go('E1:H1');
    await open();
    await click('Preview series');
    assert(button('Apply series').disabled);
    assert.match(button('Apply series').title, /commands.*unavailable/i);
    await click('Close workbook tools');
    for (const patch of [
      { literal: true },
      { literal: false, passiveNotice: 'Passive workbook', readOnly: true },
      { passiveNotice: undefined, readOnly: false, active: false },
      { active: true, busy: true },
    ]) {
      await h.render(patch);
      const opener = node('home.fill-series');
      assert(!opener || opener.disabled);
      if (opener) assert(opener.title);
      assert.deepEqual(h.accepted, []);
      assert.deepEqual(h.legacy, []);
    }
  }
);
check(
  'known source merge/table/filter/output and unsafe seed protections remain readable before acceptance',
  async (t) => {
    for (const kind of ['merge', 'table', 'filter', 'output', 'precision', 'occupied']) {
      const h = await mount(t);
      const source = h.input().sheets.map((s) =>
        s.name === 'Sheet1'
          ? {
              ...s,
              ...(kind === 'merge' ? { merges: ['G1:H1'] } : {}),
              ...(kind === 'filter'
                ? {
                    filter: {
                      range: 'E1:H6',
                      owned: false,
                      revision: null,
                      canReapply: false,
                      canClear: false,
                      reason: 'Imported filter',
                    },
                  }
                : {}),
            }
          : s
      );
      if (kind === 'output') h.data.G1 = { value: null, formulaOutput: { schema: 1, master: 'G1', range: 'G1:H1' } };
      if (kind === 'precision') h.data.E1 = { value: 9007199254740992, rawValue: '9007199254740993' };
      if (kind === 'occupied') h.data.G1 = { value: 913 };
      await h.render({
        sheets: source,
        ...(kind === 'table'
          ? {
              tables: [
                {
                  name: 'OwnedTable',
                  sheet: 'Sheet1',
                  ref: 'G1:H4',
                  path: 'xl/tables/table1.xml',
                  columns: ['One', 'Two'],
                },
              ],
            }
          : {}),
      });
      await go('E1:H1');
      await open();
      await click('Preview series');
      const message = field('Fill series').textContent;
      assert.match(
        message,
        {
          merge: /merged/,
          table: /table/,
          filter: /filtered/,
          output: /protected output|protected|stored/,
          precision: /exact literal numeric seeds/,
          occupied: /overwrite G1/,
        }[kind]
      );
      assert.equal(field('Series summary'), undefined);
      assert.equal(button('Apply series'), undefined);
      assert.deepEqual(h.accepted, []);
      assert.deepEqual(h.legacy, []);
      await h.unmount();
    }
  }
);
check(
  'held preview replies retire on actual owner/head/navigation roundtrip/sheet/activation and never use new readers',
  async (t) => {
    for (const kind of ['owner', 'head', 'navigation', 'sheet', 'active']) {
      const h = await mount(t),
        held = defer(),
        newReads = [];
      await go('E1:H1');
      h.holdRead((_sheet, range) => (range === 'E1:H1' ? held.promise : {}));
      await open();
      await click('Preview series');
      assert.equal(seriesReads(h).length, 1);
      await h.render({
        readRange: async (...args) => {
          newReads.push(args);
          return {};
        },
      });
      if (kind === 'owner') await h.render({ ownerIdentity: 'replacement', documentId: 'replacement' });
      else if (kind === 'head') await h.render({ storageKey: 'next-head' });
      else if (kind === 'navigation') {
        await go('J1:J3');
        await go('E1:H1');
      } else if (kind === 'sheet') {
        await click(button('Other'));
        await click(button('Sheet1'));
      } else {
        await h.render({ active: false });
        await h.render({ active: true });
      }
      await act(async () => held.resolve({ E1: { value: 1, rawValue: '1' }, F1: { value: 3, rawValue: '3' } }));
      await flush();
      assert.equal(field('Series summary'), undefined);
      assert.equal(button('Apply series'), undefined);
      assert.match(field('Fill series').textContent, /changed. Reopen Fill series/);
      assert.deepEqual(h.accepted, []);
      assert.deepEqual(
        newReads.filter((args) => args[1] === 'E1:H1'),
        []
      );
      assert.deepEqual(h.legacy, []);
      await h.unmount();
    }
  }
);
check(
  'Cancel or hiding the retained dock masks a held read; reopening cannot overlap the old physical query',
  async (t) => {
    for (const mode of ['cancel', 'hide']) {
      const h = await mount(t),
        held = defer();
      await go('E1:H1');
      h.holdRead((_sheet, range) =>
        range === 'E1:H1' && seriesReads(h).length === 1 ? held.promise : selectedCells(h.data, range)
      );
      await open();
      await click('Preview series');
      if (mode === 'cancel') await click('Cancel series');
      else {
        await click('Close workbook tools');
        await go('E1:H1');
        await open();
      }
      assert(button('Preview series').disabled);
      await act(async () => held.resolve(selectedCells(h.data, 'E1:H1')));
      await flush();
      assert.equal(field('Series summary'), undefined);
      await click('Preview series');
      assert.equal(seriesReads(h).length, 2);
      assert(field('Series summary'));
      assert.deepEqual(h.accepted, []);
      assert.deepEqual(h.legacy, []);
      await h.unmount();
    }
  }
);
check(
  'actual Home-tab/menu focus parks the same-head formula draft and Close restores exact text/caret without blur commit',
  async (t) => {
    const h = await mount(t);
    await go('E1:H1');
    const input = field('Formula or value');
    await act(async () => {
      input.focus();
      props(input).onChange({ target: { value: '=SUM(' } });
    });
    input.setSelectionRange(3, 5);
    const home = [...document.querySelectorAll('[role=tab]')].find((n) => n.textContent === 'Home');
    await focusClick(home);
    const opener = node('home.fill-series'),
      summary = opener.closest('details').querySelector('summary');
    await focusClick(summary);
    assert.notEqual(document.activeElement, input);
    assert.equal(input.value, '=SUM(');
    await focusClick(opener);
    await click('Preview series');
    assert(field('Series summary'));
    await click('Cancel series');
    await focusClick(button('Close workbook tools'));
    await flush(16);
    assert.equal(input.value, '=SUM(');
    assert.equal(input.selectionStart, 3);
    assert.equal(input.selectionEnd, 5);
    assert.equal(document.activeElement, input);
    assert.equal(field('Go to cell').getAttribute('placeholder'), 'E1:H1');
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'retained old opener and late acceptance failure cannot park or disturb a replacement owner formula draft',
  async (t) => {
    const h = await mount(t);
    await go('E1:H1');
    const oldOpen = props(node('home.fill-series')).onClick;
    await open();
    await click('Preview series');
    const held = defer();
    await click('Close workbook tools');
    await h.render({
      onCommandBatch: (commands) => {
        h.accepted.push(structuredClone(commands));
        return held.promise;
      },
    });
    await go('E1:H1');
    await open();
    await click('Preview series');
    await click('Apply series');
    await h.render({ documentId: 'replacement', ownerIdentity: 'replacement' });
    const input = field('Formula or value');
    await act(async () => {
      input.focus();
      props(input).onChange({ target: { value: '=NEW(' } });
    });
    input.setSelectionRange(5, 5);
    await act(async () => {
      oldOpen();
      held.reject(Error('Old owner refused fill'));
    });
    await flush();
    assert.equal(input.value, '=NEW(');
    assert.equal(input.selectionStart, 5);
    assert.equal(document.activeElement, input);
    assert(!document.body.textContent.includes('Old owner refused fill'));
    assert.equal(document.querySelector('.sheet-grid-error'), null);
    assert.equal(h.accepted.length, 1);
    assert.deepEqual(h.legacy, []);
  }
);

check(
  'Dates mode re-reads the captured source then submits content and tail date formatting in one acknowledgement',
  async (t) => {
    const h = await mount(t);
    h.data.E1 = { value: 45350, rawValue: '45350', format: 14 };
    h.data.F1 = { value: 45351, rawValue: '45351', format: 14 };
    await go('E1:H1');
    await open();
    await change(field('Series type'), 'date');
    await click('Preview series');
    assert.match(field('Series summary').textContent, /2024-02-28/);
    assert.match(field('Series destinations').textContent, /2024-03-01/);
    await click('Apply series');
    assert.deepEqual(h.accepted, [
      [
        { type: 'fillSeries', sheet: 'Sheet1', range: 'E1:H1' },
        { type: 'formatRange', sheet: 'Sheet1', range: 'G1:H1', format: { numberFormat: 'date' } },
      ],
    ]);
    assert.equal(seriesReads(h).length, 2);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Date preview rejects text/times and rechecks destination occupancy before accepting both commands',
  async (t) => {
    const h = await mount(t);
    h.data.E1 = { value: 45350, rawValue: '45350', format: 14 };
    h.data.F1 = { value: 45351.5, rawValue: '45351.5', format: 22 };
    await go('E1:H1');
    await open();
    await change(field('Series type'), 'date');
    await click('Preview series');
    assert.match(field('Fill series').textContent, /whole-day/);
    assert.deepEqual(h.accepted, []);
    h.data.F1 = { value: 45351, rawValue: '45351', format: 14 };
    await click('Preview series');
    h.data.H1 = { value: 'keep' };
    await click('Apply series');
    assert.match(field('Fill series').textContent, /overwrite H1/);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
  }
);
