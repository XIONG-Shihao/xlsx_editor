import './support/scratch-dir.mjs';
/** Current Across Worksheets Grid/Home integration; controlled owner acknowledgements, no native/app execution. */
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
import { fillAcrossFixture, acrossRange, fillAcrossCommands } from './support/fill-across-fixture.mjs';
import { readWorkbook } from '../core/workbook-package.mjs';
const component = fileURLToPath(new URL('../', import.meta.url)),
  output = await fs.mkdtemp(path.join(component, '.build/tests/fill-across-ui-')),
  checks = [];
const files = [
  'ui/SpreadsheetGrid.tsx',
  'ui/SpreadsheetRibbon.tsx',
  'ui/WorkbookFillAcross.tsx',
  'ui/fillAcrossPreview.ts',
  'ui/homeCommands.ts',
  'ui/ribbonCommands.ts',
  'ui/styles/02.css',
  'tests/workbook-fill-across-ui.mjs',
  'tests/support/fill-across-fixture.mjs',
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
  `export {SpreadsheetGrid} from ${JSON.stringify(path.join(component, 'ui/SpreadsheetGrid.tsx'))};export {configureXlsxHost} from ${JSON.stringify(path.join(component, 'ui/host.tsx'))};export * from ${JSON.stringify(path.join(component, 'ui/fillAcrossPreview.ts'))};export * from ${JSON.stringify(path.join(component, 'ui/homeCommands.ts'))};`
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
  await click(node('home.fill-across'));
  assert(field('Across Worksheets'));
  assert.equal(document.querySelectorAll('aside.sheet-task-pane:not([hidden])').length, 1);
}
async function destinations(...names) {
  for (const name of names) await change(field('Destination worksheet ' + name), true);
}
const targetRows = () =>
  [...field('Across Worksheets destinations').querySelectorAll('tbody tr')].map((row) => ({
    sheet: row.cells[0].textContent,
    range: row.cells[1].textContent,
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
    fixture = readWorkbook(fillAcrossFixture()),
    data = Object.fromEntries(fixture.sheets.map((sheet) => [sheet.name, structuredClone(sheet.cells)]));
  const hidden = (state) => ({
    ...fixture.sheets[3],
    name: state === 'hidden' ? 'Hidden' : 'VeryHidden',
    state,
    cells: {},
  });
  let override = null;
  const reader = async (sheet, range) => {
    reads.push({ sheet, range });
    return override ? override(sheet, range) : selectedCells(data[sheet] ?? {}, range);
  };
  let input = {
      sheets: [
        ...fixture.sheets.map((sheet) => ({ ...sheet, rows: 1100, columns: 20 })),
        hidden('hidden'),
        hidden('veryHidden'),
      ],
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
  'Across has closed typed preview/apply arguments and an editable native-only route between Left and Series',
  async () => {
    const binding = {
        documentId: 'doc',
        ownerIdentity: 'owner',
        state: 'head',
        sheet: 'Template',
        selection: acrossRange,
        selectedCell: 'A1',
        activationEpoch: 0,
        navigationGeneration: 0,
        modeGeneration: 0,
        specialSelectionEpoch: 0,
        mode: 'normal',
        specialSelection: null,
      },
      frame = { binding, profile: 'xlsx', active: true, canEdit: true, busy: false, acceptBatch: async () => {} };
    assert(ui.homeCommandDefinition('home.fill-across').order > ui.homeCommandDefinition('home.fill-left').order);
    assert(ui.homeCommandDefinition('home.fill-across').order < ui.homeCommandDefinition('home.fill-series').order);
    for (const id of ['home.fill-across', 'home.fill-across-preview', 'home.fill-across-apply']) {
      assert.equal(ui.homeCommandDefinition(id).requiresEdit, true);
      assert.equal(ui.homeCommandCapability(id, frame).state, 'enabled');
      for (const patch of [
        { canEdit: false },
        { busy: true },
        { active: false },
        { binding: { ...binding, ownerIdentity: null } },
        { binding: { ...binding, mode: 'pageLayout' } },
        { binding: { ...binding, specialSelection: ['A1'] } },
      ])
        assert.equal(ui.homeCommandCapability(id, { ...frame, ...patch }).state, 'disabled');
      for (const profile of ['delimited', 'passive'])
        assert.equal(ui.homeCommandCapability(id, { ...frame, profile }).state, 'hidden');
    }
    for (const id of ['home.fill-across-preview', 'home.fill-across-apply']) {
      const input = { targets: ['January', 'February'], mode: 'contents' },
        copy = ui.normalizeHomeArguments(id, input);
      input.targets.push('Untouched');
      assert.deepEqual(copy, { targets: ['January', 'February'], mode: 'contents' });
      assert(Object.isFrozen(copy.targets));
      for (const args of [
        undefined,
        {},
        { targets: [], mode: 'all' },
        { targets: ['January', 'January'], mode: 'all' },
        { targets: ['January'], mode: 'values' },
        { targets: ['January'], mode: 'all', sourceSheet: 'foreign' },
        { targets: ['January'], mode: 'all', range: 'B2' },
      ])
        assert.throws(() => ui.normalizeHomeArguments(id, args), /Invalid Home/);
    }
    const sheets = [
      { name: 'Template', state: 'visible' },
      { name: 'January', state: 'visible' },
      { name: 'February', state: 'visible' },
      { name: 'Hidden', state: 'hidden' },
    ];
    for (const targets of [['Template'], ['Hidden'], ['Unknown']])
      assert.throws(() => ui.createFillAcrossPreview(binding, sheets, { targets, mode: 'all' }), /visible destination/);
    assert.throws(
      () =>
        ui.createFillAcrossPreview({ ...binding, selection: 'A1:A6000' }, sheets, {
          targets: ['January', 'February'],
          mode: 'all',
        }),
      /10,000 destination/
    );
  }
);
check(
  'Picker excludes source/hidden worksheets; every mode submits exact workbook-order destinations in one captured batch',
  async (t) => {
    for (const [mode, label] of [
      ['all', 'All'],
      ['contents', 'Contents'],
      ['formats', 'Formats'],
    ]) {
      const h = await mount(t),
        before = structuredClone(h.data);
      await go(acrossRange);
      await open();
      assert.equal(field('Destination worksheet Template'), undefined);
      assert.equal(field('Destination worksheet Hidden'), undefined);
      assert.equal(field('Destination worksheet VeryHidden'), undefined);
      assert.deepEqual(h.accepted, []);
      await destinations('February', 'January');
      await change(field(label), true);
      await click('Preview destinations');
      assert.deepEqual(targetRows(), [
        { sheet: 'January', range: acrossRange },
        { sheet: 'February', range: acrossRange },
      ]);
      assert.match(field('Across Worksheets summary').textContent, /Template!A1:C4/);
      assert.match(field('Across Worksheets summary').textContent, /Destination cells24/);
      assert.deepEqual(h.accepted, []);
      await click('Apply Across Worksheets');
      assert.deepEqual(h.accepted, [fillAcrossCommands(mode)]);
      assert.match(field('Across Worksheets').textContent, /Destination worksheets filled. Use Undo/);
      assert.equal(field('Go to cell').placeholder, acrossRange);
      assert.deepEqual(h.data, before);
      assert.deepEqual(h.legacy, []);
      await h.unmount();
    }
  }
);
check(
  'Choice changes and Cancel invalidate retained Apply callbacks; switching Fill panes keeps one dock and pure state',
  async (t) => {
    const h = await mount(t),
      before = structuredClone(h.data);
    await go(acrossRange);
    await open();
    await destinations('January');
    await click('Preview destinations');
    const oldApply = props(button('Apply Across Worksheets')).onClick;
    await change(field('Formats'), true);
    await act(async () => oldApply());
    await flush();
    assert.equal(field('Across Worksheets summary'), undefined);
    assert.deepEqual(h.accepted, []);
    await click('Preview destinations');
    const cancelledApply = props(button('Apply Across Worksheets')).onClick;
    await click('Cancel Across Worksheets');
    await act(async () => cancelledApply());
    await flush();
    assert.equal(field('Across Worksheets summary'), undefined);
    assert.deepEqual(h.accepted, []);
    await click(node('home.fill-series'));
    assert(field('Fill series'));
    assert.equal(field('Across Worksheets'), undefined);
    assert.equal(document.querySelectorAll('aside.sheet-task-pane:not([hidden])').length, 1);
    await open();
    assert.equal(field('Fill series'), undefined);
    assert.equal(document.querySelectorAll('aside.sheet-task-pane:not([hidden])').length, 1);
    assert.deepEqual(h.data, before);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Retained Apply is retired by owner/head/selection roundtrip/sheet/activation/worksheet-inventory changes',
  async (t) => {
    for (const kind of ['owner', 'head', 'selection', 'sheet', 'activation', 'inventory']) {
      const h = await mount(t);
      await go(acrossRange);
      await open();
      await destinations('January', 'February');
      await click('Preview destinations');
      const stale = props(button('Apply Across Worksheets')).onClick;
      if (kind === 'owner') await h.render({ documentId: 'replacement', ownerIdentity: 'replacement' });
      else if (kind === 'head') await h.render({ storageKey: 'replacement-head' });
      else if (kind === 'selection') {
        await go('D9:D10');
        await go(acrossRange);
      } else if (kind === 'sheet') {
        await click('January');
        await click('Template');
      } else if (kind === 'activation') {
        await h.render({ active: false });
        await h.render({ active: true });
      } else
        await h.render({
          sheets: h.input().sheets.map((sheet) => (sheet.name === 'February' ? { ...sheet, state: 'hidden' } : sheet)),
        });
      await act(async () => stale());
      await flush();
      assert.deepEqual(h.accepted, [], kind);
      assert.equal(field('Across Worksheets summary'), undefined);
      assert.match(field('Across Worksheets').textContent, /changed. Reopen Across Worksheets/);
      assert.deepEqual(h.legacy, []);
      await h.unmount();
    }
  }
);
check(
  'One original callback is awaited once despite duplicate Apply and same-binding callback replacement; acknowledged head advance stays truthful',
  async (t) => {
    const h = await mount(t),
      held = defer(),
      replacement = [];
    await h.render({
      onCommandBatch: (commands) => {
        h.accepted.push(structuredClone(commands));
        return held.promise;
      },
    });
    await go(acrossRange);
    await open();
    await destinations('January', 'February');
    await click('Preview destinations');
    await h.render({ onCommandBatch: async (commands) => replacement.push(commands) });
    const submit = props(button('Apply Across Worksheets')).onClick;
    await act(async () => {
      submit();
      submit();
    });
    await flush();
    assert.deepEqual(h.accepted, [fillAcrossCommands('all')]);
    assert.deepEqual(replacement, []);
    assert(button('Cancel Across Worksheets').disabled);
    assert(!field('Across Worksheets').textContent.includes('Destination worksheets filled.'));
    await h.render({ storageKey: 'accepted-head' });
    await act(async () => held.resolve());
    await flush();
    assert.match(field('Across Worksheets').textContent, /Destination worksheets filled. Use Undo/);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Backend reasons remain visible on the original owner; a late rejection cannot leak into a replacement owner',
  async (t) => {
    const h = await mount(t);
    await h.render({
      onCommandBatch: async () => {
        throw Error('Destination intersects a protected merged area.');
      },
    });
    await go(acrossRange);
    await open();
    await destinations('February');
    await click('Preview destinations');
    await click('Apply Across Worksheets');
    assert.match(field('Across Worksheets').querySelector('[role=alert]').textContent, /protected merged area/);
    assert(!field('Across Worksheets').textContent.includes('Destination worksheets filled.'));
    assert.deepEqual(h.legacy, []);
    await h.unmount();
    const late = await mount(t),
      held = defer();
    await late.render({
      onCommandBatch: (commands) => {
        late.accepted.push(structuredClone(commands));
        return held.promise;
      },
    });
    await go(acrossRange);
    await open();
    await destinations('January');
    await click('Preview destinations');
    await click('Apply Across Worksheets');
    await late.render({ documentId: 'replacement', ownerIdentity: 'replacement' });
    await act(async () => held.reject(Error('Original owner refused Across')));
    await flush();
    assert(!document.body.textContent.includes('Original owner refused Across'));
    assert.deepEqual(late.legacy, []);
  }
);
check(
  'Aggregate destination budget refuses before Apply and readonly/missing-owner/callback/literal/passive/inactive profiles cannot open the pane',
  async (t) => {
    const h = await mount(t);
    await go('A1:A6000');
    await open();
    await destinations('January', 'February');
    await click('Preview destinations');
    assert.match(field('Across Worksheets').querySelector('[role=alert]').textContent, /10,000 destination/);
    assert.equal(button('Apply Across Worksheets'), undefined);
    assert.deepEqual(h.accepted, []);
    await click('Close workbook tools');
    for (const patch of [
      { readOnly: true },
      { readOnly: false, ownerIdentity: undefined },
      { ownerIdentity: 'original-owner', onCommandBatch: undefined },
      { onCommandBatch: async (commands) => h.accepted.push(commands), literal: true },
      { literal: false, passiveNotice: 'Passive workbook' },
      { passiveNotice: undefined, active: false },
      { active: true, busy: true },
    ]) {
      await h.render(patch);
      assert(node('home.fill-across').disabled);
      assert(node('home.fill-across').title);
      const stale = props(node('home.fill-across')).onClick;
      await act(async () => stale());
      await flush();
      assert.equal(field('Across Worksheets'), undefined);
      assert.deepEqual(h.accepted, []);
      assert.deepEqual(h.legacy, []);
    }
  }
);
check(
  'Across menu pointer focus preserves exact parked formula text/caret and Close restores it without a blur commit',
  async (t) => {
    const h = await mount(t);
    await go(acrossRange);
    const input = field('Formula or value');
    await act(async () => {
      input.focus();
      props(input).onChange({ target: { value: '=SUM(' } });
    });
    input.setSelectionRange(3, 5);
    const home = [...document.querySelectorAll('[role=tab]')].find((node) => node.textContent === 'Home');
    await focusClick(home);
    const opener = node('home.fill-across'),
      summary = opener.closest('details').querySelector('summary');
    await focusClick(summary);
    await focusClick(opener);
    await destinations('January');
    await click('Preview destinations');
    await click('Cancel Across Worksheets');
    await focusClick(button('Close workbook tools'));
    await flush(16);
    assert.equal(input.value, '=SUM(');
    assert.equal(input.selectionStart, 3);
    assert.equal(input.selectionEnd, 5);
    assert.equal(document.activeElement, input);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
  }
);
