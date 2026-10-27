import './support/scratch-dir.mjs';
/** Actual current Home Grid/Ribbon in React/JSDOM; controlled reads/acknowledgements,
 * no native/clipboard-OS/default browser keyboard/Undo/file-authority claim. */
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
  output = await fs.mkdtemp(path.join(component, '.build/tests/home-ui-')),
  checks = [];
const files = [
  'ui/SpreadsheetGrid.tsx',
  'ui/SpreadsheetRibbon.tsx',
  'ui/homeCommands.ts',
  'ui/ribbonCommands.ts',
  'ui/WorkbookFill.tsx',
  'ui/fillSeriesPreview.ts',
  'ui/WorkbookReplace.tsx',
  'ui/replacePreview.ts',
  'ui/WorkbookSpecialSelection.tsx',
  'ui/specialSelection.ts',
  'ui/searchPages.ts',
  'core/value-paste.mjs',
  'core/used-range.mjs',
  'tests/workbook-home-ui.mjs',
];
const pins = async () =>
    Promise.all(
      files.map(async (relative) => {
        const p = path.join(component, relative),
          b = await fs.readFile(p);
        return { relative, path: p, sha256: createHash('sha256').update(b).digest('hex'), bytes: b.length };
      })
    ),
  before = await pins();
const entry = path.join(output, 'entry.ts'),
  bundle = path.join(output, 'test.mjs');
await fs.writeFile(
  entry,
  `export {SpreadsheetGrid} from ${JSON.stringify(path.join(component, 'ui/SpreadsheetGrid.tsx'))};export {configureXlsxHost} from ${JSON.stringify(path.join(component, 'ui/host.tsx'))};export * from ${JSON.stringify(path.join(component, 'ui/homeCommands.ts'))};`
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
let clipboard = '',
  clipboardRead = async () => clipboard,
  clipboardWrite = async (text) => {
    clipboard = text;
  };
ui.configureXlsxHost({
  api: { copyText: (text) => clipboardWrite(text), readClipboardText: () => clipboardRead() },
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
  node = (id) => [...document.querySelectorAll('[data-home-command="' + id + '"]')].find((n) => !n.closest('[hidden]'));
const defer = () => {
  let resolve, reject;
  const promise = new Promise((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
};
const flush = async (ms = 12) => act(async () => new Promise((resolve) => setTimeout(resolve, ms)));
async function click(id) {
  const n = typeof id === 'string' ? node(id) : id;
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
async function go(value) {
  await change(field('Go to cell'), value);
  await act(async () =>
    field('Go to cell').dispatchEvent(
      new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })
    )
  );
  await flush();
}
const specialPane = () => document.querySelector('[aria-label="Go To Special"]'),
  specialButton = (name) => [...specialPane().querySelectorAll('button')].find((button) => button.textContent === name);
async function specialSettled() {
  for (let attempt = 0; attempt < 50; attempt++) {
    if (specialPane().querySelector('[aria-label="Special selection summary"],[role="alert"]')) return;
    await flush(20);
  }
  throw Error('The Special selection preview did not settle.');
}
async function selectSpecial(id) {
  await click(id);
  await click(specialButton('Preview matching cells'));
  await specialSettled();
  await click(specialButton('Select matching cells'));
  await click(field('Close workbook tools'));
}

async function mount(t, patch = {}) {
  clipboard = '';
  clipboardRead = async () => clipboard;
  clipboardWrite = async (text) => {
    clipboard = text;
  };
  const root = createRoot(document.getElementById('root')),
    legacy = [],
    accepted = [],
    reads = [],
    history = [],
    data = {
      A1: { value: 'Alpha', dataType: 'text' },
      A2: { value: null },
      A3: { value: null, formula: 'SUM(B1:B3)', calculation: { status: 'calculated', value: 6 } },
      B1: { value: 1 },
      B2: { value: 2 },
      B3: { value: 3 },
      C1: { value: '#N/A', dataType: 'text' },
    };
  const sheet = (name) => ({
    name,
    path: 'xl/worksheets/' + name + '.xml',
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
  });
  let readOverride = null;
  const reader = async (name, range) => {
    reads.push({ name, range });
    if (readOverride) return readOverride(name, range);
    const b = rangeBounds(range);
    return Object.fromEntries(
      Object.entries(data).filter(([ref]) => {
        const p = rangeBounds(ref).first;
        return p.row >= b.first.row && p.row <= b.last.row && p.col >= b.first.col && p.col <= b.last.col;
      })
    );
  };
  let input = {
    sheets: [sheet('Sheet1'), sheet('Other')],
    documentId: 'doc',
    ownerIdentity: 'runtime',
    storageKey: 'opaque-head',
    active: true,
    readOnly: false,
    readRange: reader,
    onEdit: (value) => legacy.push(value),
    onBatch: (value) => legacy.push(value),
    onCommandBatch: async (commands) => {
      accepted.push(commands);
    },
    onCommandUndo: async () => history.push('undo'),
    onCommandRedo: async () => history.push('redo'),
    canUndo: true,
    canRedo: true,
    ...patch,
  };
  const render = async (patch) => {
    input = { ...input, ...patch };
    await act(async () => root.render(React.createElement(ui.SpreadsheetGrid, input)));
    await flush();
  };
  await render();
  t.after(async () => {
    await act(async () => root.unmount());
    clipboardRead = async () => clipboard;
    clipboardWrite = async (text) => {
      clipboard = text;
    };
  });
  return {
    accepted,
    legacy,
    reads,
    history,
    data,
    input: () => input,
    reader,
    holdRead: (fn) => (readOverride = fn),
    render,
  };
}
const conditional = () => ({
  documentId: 'hostdoc',
  revision: 'head',
  inventory: async (sheet) => ({ sheet, revision: 'rules', canEdit: true, protectedReasons: [], rules: [] }),
  onBatch: async () => {},
});
check(
  'actual seven Home groups/control inventory maps real actions, including local drafts, menus, history and deferred gestures',
  async (t) => {
    await mount(t);
    for (const name of ['Clipboard', 'Font', 'Alignment', 'Number', 'Styles', 'Cells', 'Editing'])
      assert(document.body.textContent.includes(name));
    // The screen is split into modules under ui/grid; the command ids may live in any of them.
    const gridFiles = (await fs.readdir(path.join(component, 'ui/grid'), { recursive: true })).filter((name) =>
      /\.tsx?$/.test(name)
    );
    const source = (
      await Promise.all(
        ['ui/SpreadsheetGrid.tsx', ...gridFiles.map((name) => 'ui/grid/' + name)].map((file) =>
          fs.readFile(path.join(component, file), 'utf8')
        )
      )
    ).join('\n');
    for (const id of ui.homeCommandDefinitions.map((d) => d.id))
      assert(source.includes("'" + id + "'") || source.includes('"' + id + '"'), id);
    for (const n of document.querySelectorAll(
      '.sheet-ribbon-tabpanel:not([hidden]) button,.sheet-ribbon-tabpanel:not([hidden]) input,.sheet-ribbon-tabpanel:not([hidden]) select,.sheet-ribbon-tabpanel:not([hidden]) summary'
    )) {
      if (n.closest('.sheet-ribbon-group-toggle')) continue;
      assert(n.getAttribute('data-home-command') || n.hasAttribute('data-ribbon-command'), n.outerHTML);
    }
    assert(node('home.undo') && node('home.redo'));
  }
);
check(
  'Font/Alignment/Number and preset actions submit exact one awaited strict original-target format batch',
  async (t) => {
    const h = await mount(t);
    await go('B1:B3');
    const values = [
      ['home.font-family', 'Arial', { fontFamily: 'Arial' }],
      ['home.font-size', '14', { fontSize: 14 }],
      ['home.bold', null, { bold: true }],
      ['home.italic', null, { italic: true }],
      ['home.underline', null, { underline: true }],
      ['home.borders', null, { border: 'all' }],
      ['home.fill-color', '#11aa22', { fillColor: '11aa22' }],
      ['home.font-color', '#aa1122', { fontColor: 'aa1122' }],
      ['home.wrap', null, { wrapText: true }],
      ['home.align-top', null, { verticalAlign: 'top' }],
      ['home.align-middle', null, { verticalAlign: 'center' }],
      ['home.align-bottom', null, { verticalAlign: 'bottom' }],
      ['home.align-left', null, { horizontalAlign: 'left' }],
      ['home.align-center', null, { horizontalAlign: 'center' }],
      ['home.align-right', null, { horizontalAlign: 'right' }],
      ['home.number-format', 'date', { numberFormat: 'date' }],
      ['home.accounting', null, { numberFormat: 'accounting' }],
      ['home.percent', null, { numberFormat: 'percent' }],
      ['home.thousands', null, { numberFormat: 'thousands' }],
      ['home.decimals-increase', null, { numberFormat: 'decimal1' }],
      ['home.decimals-decrease', null, { numberFormat: 'decimal0' }],
      ['home.style-good', null, { bold: true, fontColor: '006100', fillColor: 'C6EFCE', border: 'none' }],
      ['home.style-bad', null, { bold: true, fontColor: '9C0006', fillColor: 'FFC7CE', border: 'none' }],
      ['home.style-heading', null, { bold: true, fontColor: 'FFFFFF', fillColor: '244062', border: 'none' }],
    ];
    for (const [id, value, format] of values) {
      if (value !== null) await change(node(id), value);
      else await click(id);
      assert.deepEqual(h.accepted.at(-1), [{ type: 'formatRange', sheet: 'Sheet1', range: 'B1:B3', format }], id);
    }
    assert.equal(h.accepted.length, values.length);
    assert.deepEqual(h.legacy, []);
  }
);
check('merge/unmerge, table and conditional commands retain exact fields and source capability refusals', async (t) => {
  const h = await mount(t);
  await go('A1:B2');
  await click('home.merge-center');
  assert.deepEqual(h.accepted.pop(), [
    { sheet: 'Sheet1', operation: 'mergeCells', range: 'A1:B2' },
    {
      type: 'formatRange',
      sheet: 'Sheet1',
      range: 'A1',
      format: { horizontalAlign: 'center', verticalAlign: 'center' },
    },
  ]);
  await h.render({ sheets: h.input().sheets.map((s) => ({ ...s, merges: ['A1:B2'] })) });
  await click('home.unmerge');
  assert.equal(h.accepted.at(-1)[0].operation, 'unmergeCells');
  await change(node('home.table-draft'), 'Table_New');
  await click('home.create-table');
  assert.deepEqual(h.accepted.at(-1), [
    { sheet: 'Sheet1', operation: 'createTable', range: 'A1:B2', tableName: 'Table_New' },
  ]);
  await change(field('Conditional threshold'), '2');
  await click('home.conditional-apply');
  assert.equal(h.accepted.at(-1)[0].threshold, 2);
  assert.deepEqual(h.legacy, []);
});
check('Cells structural actions and dimension drafts await strict acceptance without legacy point edits', async (t) => {
  const h = await mount(t);
  await go('C4');
  for (const [id, operation, index] of [
    ['home.insert-row', 'insertRow', 3],
    ['home.delete-row', 'deleteRow', 3],
    ['home.insert-column', 'insertColumn', 2],
    ['home.delete-column', 'deleteColumn', 2],
  ]) {
    await click(id);
    assert.deepEqual(h.accepted.at(-1), [{ sheet: 'Sheet1', operation, index }]);
  }
  const width = node('home.column-width');
  await change(width, '18');
  await act(async () => props(width).onBlur());
  await flush();
  assert.deepEqual(h.accepted.at(-1), [{ sheet: 'Sheet1', operation: 'resizeColumn', index: 2, width: 18 }]);
  const height = node('home.row-height');
  await change(height, '24');
  await act(async () => props(height).onBlur());
  await flush();
  assert.deepEqual(h.accepted.at(-1), [{ sheet: 'Sheet1', operation: 'resizeRow', index: 3, height: 24 }]);
  assert.deepEqual(h.legacy, []);
});
check(
  'Fill and Clear controls keep declared command semantics and one strict batch, including special formats/contents',
  async (t) => {
    const h = await mount(t);
    h.data.B3 = { value: null };
    await h.render({
      readRange: h.reader,
      sheets: h.input().sheets.map((sheet) => ({ ...sheet, cells: { ...sheet.cells, B3: { value: null } } })),
    });
    await go('B1:C3');
    await click('home.fill-down');
    assert.deepEqual(h.accepted.at(-1), [{ type: 'fillRange', sheet: 'Sheet1', source: 'B1:C1', range: 'B2:C3' }]);
    await click('home.fill-right');
    assert.deepEqual(h.accepted.at(-1), [{ type: 'fillRange', sheet: 'Sheet1', source: 'B1:B3', range: 'C1:C3' }]);
    await go('B1:B3');
    const beforeSeries = h.accepted.length;
    await click('home.fill-series');
    assert.equal(h.accepted.length, beforeSeries);
    const seriesButton = (name) =>
      [...field('Fill series').querySelectorAll('button')].find((button) => button.textContent === name);
    await click(seriesButton('Preview series'));
    assert.equal(h.accepted.length, beforeSeries);
    assert(h.reads.some((read) => read.range === 'B1:B3'));
    assert.deepEqual(
      [...field('Series destinations').querySelectorAll('tbody tr')].map((row) => [
        row.cells[0].textContent,
        row.cells[1].textContent,
      ]),
      [['B3', '3']]
    );
    await click(seriesButton('Apply series'));
    assert.equal(h.accepted.length, beforeSeries + 1);
    assert.deepEqual(h.accepted.at(-1), [{ type: 'fillSeries', sheet: 'Sheet1', range: 'B1:B3' }]);
    await click('home.clear-contents');
    assert.deepEqual(h.accepted.at(-1), [{ type: 'clearRange', sheet: 'Sheet1', range: 'B1:B3' }]);
    await click('home.clear-all');
    assert.equal(h.accepted.at(-1).length, 2);
    await go('A1:A3');
    await selectSpecial('home.select-blanks');
    await click('home.bold');
    assert.deepEqual(h.accepted.at(-1), [
      { type: 'formatRange', sheet: 'Sheet1', range: 'A2', format: { bold: true } },
    ]);
    assert(node('home.paste').disabled);
    assert.match(node('home.paste').title, /Special selection/);
    await click('home.clear-contents');
    assert.deepEqual(h.accepted.at(-1), [{ type: 'setCells', sheet: 'Sheet1', cells: [{ cell: 'A2', value: null }] }]);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Fill Up/Left submit one compact bottom/right-edge command without materializing source types or styles',
  async (t) => {
    const h = await mount(t);
    Object.assign(h.data, {
      A4: { value: '#N/A', dataType: 'error', styleId: 2, fontFamily: 'Arial' },
      B4: { value: '#N/A', dataType: 'text', styleId: 1 },
      D1: { formula: 'H1+$H1+H$1+$H$1', styleId: 3, calculation: { status: 'stored', pending: true } },
      D2: { value: true, styleId: 2 },
      A10000: { value: '00123', dataType: 'text', styleId: 1 },
    });
    await h.render({
      readRange: h.reader,
      sheets: h.input().sheets.map((sheet) => ({ ...sheet, rows: 10001, cells: structuredClone(h.data) })),
    });
    const source = structuredClone(h.data);
    const vectors = [
      ['home.fill-up', 'A1:B4', 'A4:B4', 'A1:B3'],
      ['home.fill-left', 'A1:D2', 'D1:D2', 'A1:C2'],
      ['home.fill-up', 'A1:A10000', 'A10000:A10000', 'A1:A9999'],
      ['home.fill-left', 'A1:NTP1', 'NTP1:NTP1', 'A1:NTO1'],
    ];
    for (const [id, selection, edge, target] of vectors) {
      await go(selection);
      const count = h.accepted.length;
      await click(id);
      assert.equal(h.accepted.length, count + 1);
      assert.deepEqual(h.accepted.at(-1), [{ type: 'fillRange', sheet: 'Sheet1', source: edge, range: target }]);
      assert.deepEqual(h.data, source);
      assert.equal(document.querySelector('.spreadsheet-editor').dataset.homePending, 'false');
    }
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Fill Up/Left refuse missing destination dimension, oversized selection and unsupported write profiles',
  async (t) => {
    const h = await mount(t);
    for (const [id, selection, reason] of [
      ['home.fill-up', 'A1:C1', /at least two rows to fill up/],
      ['home.fill-left', 'A1:A3', /at least two columns to fill left/],
      ['home.fill-up', 'A1:A10001', /10,000-cell limit/],
      ['home.fill-left', 'A1:A10001', /10,000-cell limit/],
    ]) {
      await go(selection);
      await click(id);
      assert.deepEqual(h.accepted, []);
      assert.match(document.body.textContent, reason);
    }
    await go('A1:A3');
    await selectSpecial('home.select-blanks');
    for (const id of ['home.fill-up', 'home.fill-left']) {
      assert(node(id).disabled);
      assert.match(node(id).title, /Special selection/);
    }
    await go('A1:B3');
    await h.render({ readOnly: true });
    for (const id of ['home.fill-up', 'home.fill-left']) {
      assert(node(id).disabled);
      assert.match(node(id).title, /read-only/);
    }
    await h.render({
      readOnly: false,
      literal: true,
      storageKey: undefined,
      acceptedCommandState: 'csv-head',
      ownerIdentity: 'csv-owner',
    });
    for (const id of ['home.fill-up', 'home.fill-left']) {
      assert(node(id).disabled);
      assert.match(node(id).title, /format does not support/);
    }
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Fill Up/Left keep original receipt, await one acknowledgement and retire old owner/selection handlers',
  async (t) => {
    const h = await mount(t);
    await go('A1:B4');
    const staleUp = props(node('home.fill-up')).onClick;
    await h.render({ ownerIdentity: 'other-owner' });
    await act(async () => staleUp());
    await flush();
    assert.deepEqual(h.accepted, []);
    const staleLeft = props(node('home.fill-left')).onClick;
    await go('C1:D2');
    await act(async () => staleLeft());
    await flush();
    assert.deepEqual(h.accepted, []);
    const ack = defer();
    await h.render({
      onCommandBatch: (commands) => {
        h.accepted.push(commands);
        return ack.promise;
      },
    });
    await go('A1:B4');
    await click('home.fill-up');
    assert.deepEqual(h.accepted, [[{ type: 'fillRange', sheet: 'Sheet1', source: 'A4:B4', range: 'A1:B3' }]]);
    assert.equal(document.querySelector('.spreadsheet-editor').dataset.homePending, 'true');
    assert(node('home.fill-left').disabled);
    await h.render({ ownerIdentity: 'next-owner', documentId: 'next-doc' });
    await act(async () => ack.resolve());
    await flush();
    assert.equal(h.accepted.length, 1);
    assert.equal(document.querySelector('.spreadsheet-editor').dataset.homePending, 'false');
    assert.equal(document.querySelector('.sheet-grid-error'), null);
    const reject = defer();
    await h.render({ onCommandBatch: () => reject.promise });
    await go('A1:D2');
    await click('home.fill-left');
    await h.render({ ownerIdentity: 'replacement-owner' });
    await act(async () => reject.reject(Error('Original Fill Left refused')));
    await flush();
    assert(!document.body.textContent.includes('Original Fill Left refused'));
    assert.equal(h.accepted.length, 1);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Go To Special reads the full original rectangle, pages offscreen constants, filters types and selects without history',
  async (t) => {
    const h = await mount(t);
    for (const key of Object.keys(h.data)) delete h.data[key];
    Object.assign(
      h.data,
      Object.fromEntries(
        Array.from({ length: 1004 }, (_, i) => ['A' + (i + 1), { value: i + 1, dataType: 'number', styleId: 1 }])
      ),
      {
        A2: { value: '00123', dataType: 'text' },
        A3: { value: true, dataType: 'boolean' },
        A4: { value: '#N/A', dataType: 'error' },
        A5: { value: '#N/A', dataType: 'text' },
        A6: { value: null, formula: '1', calculation: { status: 'stored', pending: true } },
      }
    );
    await h.render({ sheets: h.input().sheets.map((s) => ({ ...s, rows: 1004, cells: {} })) });
    await go('A1:A1004');
    const before = structuredClone(h.data);
    h.reads.length = 0;
    await click('home.special-open');
    assert.equal(document.querySelectorAll('.sheet-task-pane').length, 1);
    assert.equal(field('Selection type').value, 'constants');
    assert.deepEqual(h.accepted, []);
    await click(specialButton('Preview matching cells'));
    await specialSettled();
    const expected = Array.from({ length: 1004 }, (_, i) => 'A' + (i + 1)).filter((ref) => ref !== 'A6'),
      seen = [];
    do {
      seen.push(
        ...[...field('Special selection targets').querySelectorAll('tbody tr')].map((row) => row.cells[0].textContent)
      );
      if (specialButton('More matching cells').disabled) break;
      await click(specialButton('More matching cells'));
    } while (true);
    assert.deepEqual(
      seen,
      expected.map((ref) => 'Sheet1!' + ref)
    );
    assert(h.reads.some((r) => r.range === 'A1:A1000'));
    assert(h.reads.some((r) => r.range === 'A1001:A1004'));
    assert.equal(node('home.fill-up').disabled, false);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.history, []);
    for (const name of ['Text', 'Logicals', 'Errors'])
      await act(async () => props(field(name)).onChange({ target: { checked: false } }));
    await flush();
    assert.equal(field('Special selection targets'), null);
    await click(specialButton('Preview matching cells'));
    await specialSettled();
    const numbers = expected.filter((ref) => !['A2', 'A3', 'A4', 'A5'].includes(ref));
    await click(specialButton('Select matching cells'));
    assert.match(document.body.textContent, /Matching cells selected/);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.history, []);
    assert.deepEqual(h.data, before);
    await click(field('Close workbook tools'));
    assert(node('home.fill-up').disabled);
    assert.match(node('home.fill-up').title, /Special selection/);
    await click('home.bold');
    assert.deepEqual(h.accepted, [
      numbers
        .slice()
        .sort()
        .map((range) => ({ type: 'formatRange', sheet: 'Sheet1', range, format: { bold: true } })),
    ]);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Go To Special uses source hidden axes and authored rule membership, including blank and known opaque ranges',
  async (t) => {
    const h = await mount(t);
    for (const key of Object.keys(h.data)) delete h.data[key];
    h.data.A1 = { value: 2, dataType: 'number' };
    const conditional = {
        sheet: 'Sheet1',
        revision: 'rules',
        canEdit: false,
        protectedReasons: ['Unsupported authored rule'],
        protectedRanges: ['$A$1'],
        rules: [{ range: '$B$2:$C$4', kind: 'greaterThan', values: [999], fillColor: 'FFFF00', stopIfTrue: false }],
      },
      validation = {
        sheet: 'Sheet1',
        revision: 'validation',
        canEdit: false,
        protectedReasons: ['Unsupported multi-area validation'],
        protectedRanges: ['$B$2', '$C$4'],
        rules: [],
      };
    let conditionalReply = conditional;
    await h.render({
      sheets: h.input().sheets.map((s) => ({
        ...s,
        cells: {},
        hiddenRows: [1],
        rowHeights: [{ row: 2, height: 0 }],
        widths: [{ min: 2, max: 2, width: 14, hidden: true }],
      })),
      conditionalControls: {
        documentId: 'doc',
        revision: 'opaque-head',
        inventory: async () => structuredClone(conditionalReply),
        onBatch: async () => assert.fail('Selection must not write rules'),
      },
      validationControls: {
        documentId: 'doc',
        revision: 'opaque-head',
        inventory: async () => structuredClone(validation),
        onBatch: async () => assert.fail('Selection must not write rules'),
      },
    });
    await go('A1:C4');
    await click('home.special-open');
    await change(field('Selection type'), 'visible');
    await click(specialButton('Preview matching cells'));
    await specialSettled();
    assert.deepEqual(
      [...field('Special selection targets').querySelectorAll('tbody tr')].map((r) => r.cells[0].textContent),
      ['Sheet1!A1', 'Sheet1!C1', 'Sheet1!A4', 'Sheet1!C4']
    );
    await change(field('Selection type'), 'conditional-formats');
    await click(specialButton('Preview matching cells'));
    await specialSettled();
    assert.deepEqual(
      [...field('Special selection targets').querySelectorAll('tbody tr')].map((r) => r.cells[0].textContent),
      ['Sheet1!A1', 'Sheet1!B2', 'Sheet1!C2', 'Sheet1!B3', 'Sheet1!C3', 'Sheet1!B4', 'Sheet1!C4']
    );
    await change(field('Selection type'), 'data-validation');
    await click(specialButton('Preview matching cells'));
    await specialSettled();
    assert.deepEqual(
      [...field('Special selection targets').querySelectorAll('tbody tr')].map((r) => r.cells[0].textContent),
      ['Sheet1!B2', 'Sheet1!C4']
    );
    conditionalReply = { ...conditional, protectedRanges: undefined };
    await change(field('Selection type'), 'conditional-formats');
    await click(specialButton('Preview matching cells'));
    await specialSettled();
    assert.match(specialPane().querySelector('[role=alert]').textContent, /opaque rule.*range/i);
    assert.equal(field('Special selection targets'), null);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.history, []);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Formula type selection captures the real owner revision when worksheet metadata has no calculationRevision, and retires foreign or changed receipts',
  async (t) => {
    const h = await mount(t);
    for (const ref of Object.keys(h.data)) delete h.data[ref];
    Object.assign(h.data, {
      A1: { value: 2, formula: '1+1', calculation: { status: 'calculated', value: 2, revision: 'owner-revision' } },
      A2: { value: '', formula: '""', calculation: { status: 'calculated', value: '', revision: 'owner-revision' } },
    });
    const controls = {
      receipt: { documentId: 'doc', ownerId: 'runtime', state: 'base', revision: 'owner-revision' },
      policy: {
        schema: 1,
        mode: 'auto',
        recalculateBeforeSave: false,
        pendingFormulaCells: 0,
        canCalculateSheet: true,
        canChangeOptions: true,
      },
      optionsScope: 'workbook',
      run: async () => {
        throw Error('Selection must not calculate');
      },
      cancel: async () => {
        throw Error('Selection must not calculate');
      },
    };
    await h.render({
      storageKey: 'owner-fixture',
      calculationControls: controls,
      sheets: h.input().sheets.map((s) => ({ ...s, cells: {}, calculationRevision: undefined })),
    });
    await go('A1:A2');
    await click('home.select-formulas');
    for (const name of ['Text', 'Logicals', 'Errors'])
      await act(async () => props(field(name)).onChange({ target: { checked: false } }));
    await flush();
    await click(specialButton('Preview matching cells'));
    await specialSettled();
    assert.deepEqual(
      [...field('Special selection targets').querySelectorAll('tbody tr')].map((row) => row.cells[0].textContent),
      ['Sheet1!A1']
    );
    const stale = props(specialButton('Select matching cells')).onClick;
    await h.render({
      calculationControls: { ...controls, receipt: { ...controls.receipt, revision: 'new-owner-revision' } },
    });
    await act(async () => stale());
    await flush();
    assert.equal(field('Special selection targets'), null);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.history, []);
    await click(field('Close workbook tools'));
    await h.render({
      calculationControls: { ...controls, receipt: { ...controls.receipt, ownerId: 'foreign-owner' } },
    });
    await go('A1:A2');
    await click('home.select-formulas');
    for (const name of ['Text', 'Logicals', 'Errors'])
      await act(async () => props(field(name)).onChange({ target: { checked: false } }));
    await flush();
    await click(specialButton('Preview matching cells'));
    await specialSettled();
    assert.match(specialPane().querySelector('[role=alert]').textContent, /current defined result|type authority/);
    assert.equal(field('Special selection targets'), null);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.history, []);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Go To Special Formulas opens its mask pane and refuses stored type authority while all authored formulas remain inspectable',
  async (t) => {
    const h = await mount(t);
    for (const key of Object.keys(h.data)) delete h.data[key];
    const current = (value) => ({ status: 'calculated', value, revision: 'current-calculation' });
    Object.assign(h.data, {
      A1: { value: null, formula: '1', calculation: current(1) },
      A2: { value: null, formula: '""', calculation: current('') },
      A3: { value: null, formula: 'TRUE', calculation: current(true) },
      A4: {
        value: null,
        formula: '1/0',
        calculation: { status: 'error', error: '#DIV/0!', revision: 'current-calculation' },
      },
      A5: {
        value: 913,
        formula: 'A99',
        calculation: { status: 'stored', pending: true },
        cachedResult: { authority: 'stored', status: 'calculated', value: 913 },
      },
    });
    await h.render({
      readOnly: true,
      onCommandBatch: undefined,
      sheets: h.input().sheets.map((s) => ({ ...s, cells: {}, calculationRevision: 'current-calculation' })),
    });
    await go('A1:A5');
    await click('home.select-formulas');
    assert.equal(field('Selection type').value, 'formulas');
    await click(specialButton('Preview matching cells'));
    await specialSettled();
    assert.equal(field('Special selection targets').querySelectorAll('tbody tr').length, 5);
    for (const name of ['Text', 'Logicals', 'Errors'])
      await act(async () => props(field(name)).onChange({ target: { checked: false } }));
    await flush();
    await click(specialButton('Preview matching cells'));
    await specialSettled();
    assert.match(specialPane().querySelector('[role=alert]').textContent, /stored or pending|current defined result/);
    assert.equal(field('Special selection targets'), null);
    assert.deepEqual(h.accepted, []);
    for (const name of ['Text', 'Logicals', 'Errors'])
      await act(async () => props(field(name)).onChange({ target: { checked: true } }));
    await flush();
    await click(specialButton('Preview matching cells'));
    await specialSettled();
    await click(specialButton('Select matching cells'));
    assert.match(document.body.textContent, /Matching cells selected/);
    assert.deepEqual(h.history, []);
    assert.deepEqual(h.accepted, []);
    assert(node('home.bold').disabled);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Go To Special refuses missing/incomplete/foreign reads and rechecks matching cells before selection-only Apply',
  async (t) => {
    const h = await mount(t);
    await go('A1:A3');
    await h.render({ readRange: undefined });
    assert(node('home.special-open').disabled);
    assert.match(node('home.special-open').title, /complete original.*unavailable/);
    assert.deepEqual(h.accepted, []);
    await h.render({ readRange: h.reader });
    await click('home.special-open');
    h.holdRead((_name, range) =>
      range === 'A1:A3' ? Promise.resolve({ B1: { value: 7, dataType: 'number' } }) : Promise.resolve({})
    );
    await click(specialButton('Preview matching cells'));
    await specialSettled();
    assert.match(specialPane().querySelector('[role=alert]').textContent, /foreign or malformed/);
    assert.equal(field('Special selection targets'), null);
    h.holdRead(null);
    await click(specialButton('Preview matching cells'));
    await specialSettled();
    assert.deepEqual(
      [...field('Special selection targets').querySelectorAll('tbody tr')].map((r) => r.cells[0].textContent),
      ['Sheet1!A1']
    );
    h.data.A1 = { value: null };
    await click(specialButton('Select matching cells'));
    assert.match(specialPane().querySelector('[role=alert]').textContent, /matching cells changed/);
    assert.equal(node('home.fill-up').disabled, false);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.history, []);
    await click(specialButton('Preview matching cells'));
    await specialSettled();
    assert(specialButton('Select matching cells').disabled);
    assert.match(document.body.textContent, /No cells match/);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Go To Special cancellation joins held reads, stale heads cannot apply, and Close restores an unchanged parked draft',
  async (t) => {
    const h = await mount(t);
    await go('B1:B3');
    const formula = field('Formula or value');
    await act(async () => {
      formula.focus();
      props(formula).onChange({ target: { value: '=SUM(' } });
    });
    formula.setSelectionRange(5, 5);
    await click('home.special-open');
    assert.equal(formula.value, '=SUM(');
    const held = defer();
    let calls = 0;
    h.holdRead((_name, range) => (range === 'B1:B3' ? (calls++, held.promise) : Promise.resolve({})));
    await click(specialButton('Preview matching cells'));
    await click(specialButton('Cancel selection'));
    assert(specialButton('Preview matching cells').disabled);
    await act(async () => props(specialButton('Preview matching cells')).onClick());
    assert.equal(calls, 1);
    await act(async () => held.resolve({ B1: { value: 1 } }));
    await flush();
    assert.equal(field('Special selection targets'), null);
    assert.deepEqual(h.accepted, []);
    h.holdRead(null);
    await click(field('Close workbook tools'));
    assert.equal(formula.value, '=SUM(');
    assert.equal(formula.selectionStart, 5);
    assert.equal(document.activeElement, formula);
    assert.deepEqual(h.accepted, []);
    await click('home.special-open');
    await click(specialButton('Preview matching cells'));
    await specialSettled();
    const stale = props(specialButton('Select matching cells')).onClick;
    await h.render({ storageKey: 'new-head' });
    await act(async () => stale());
    await flush();
    assert.equal(field('Special selection targets'), null);
    assert.match(document.body.textContent, /workbook or selection changed/);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.history, []);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Merged Special selections retain Excel-observed complete areas and active cursors without workbook changes',
  async (t) => {
    const h = await mount(t);
    for (const key of Object.keys(h.data)) delete h.data[key];
    const calc = (value) => ({ status: 'calculated', value, revision: 'merged-calculation' });
    Object.assign(h.data, {
      A1: { value: 7, dataType: 'number' },
      B1: { value: null, dataType: 'blank' },
      A2: { value: null, dataType: 'blank' },
      B2: { value: null, dataType: 'blank' },
      A3: { value: 2, dataType: 'number', formula: '1+1', calculation: calc(2) },
      B3: { value: null, dataType: 'blank' },
      C3: { value: '', dataType: 'text', formula: '""', calculation: calc('') },
    });
    await h.render({
      storageKey: 'merged-fixture-head',
      sheets: h.input().sheets.map((s) => ({
        ...s,
        cells: structuredClone(h.data),
        merges: ['A1:B1', 'A2:B2', 'A3:B3'],
        calculationRevision: 'merged-calculation',
      })),
    });
    const before = structuredClone(h.data);
    const cases = [
      ['A1:C3', 'constants', ['A1', 'B1'], 'A1', '7'],
      ['A1:C3', 'formulas', ['A3', 'B3', 'C3'], 'A3', '=1+1'],
      ['A1:C3', 'blanks', ['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3'], 'C1', ''],
      ['A1:C3', 'visible', ['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3'], 'A1', '7'],
      ['A1:B3', 'blanks', ['A1', 'B1', 'A2', 'B2', 'A3', 'B3'], 'A1', '7'],
      ['A1:B4', 'blanks', ['A1', 'B1', 'A2', 'B2', 'A3', 'B3', 'A4', 'B4'], 'A1', '7'],
      ['B1:C3', 'constants', ['A1', 'B1'], 'A1', '7'],
    ];
    for (const [range, kind, references, cursor, formula] of cases) {
      await go(range);
      await click('home.special-open');
      await change(field('Selection type'), kind);
      await click(specialButton('Preview matching cells'));
      await specialSettled();
      assert.deepEqual(
        [...field('Special selection targets').querySelectorAll('tbody tr')].map((row) => row.cells[0].textContent),
        references.map((ref) => 'Sheet1!' + ref),
        range + ' ' + kind
      );
      await click(specialButton('Select matching cells'));
      assert.match(document.body.textContent, /Matching cells selected/);
      assert.equal(field('Go to cell').placeholder, cursor);
      assert.equal(field('Formula or value').value, formula);
      await click(field('Close workbook tools'));
      assert.equal(field('Go to cell').placeholder, cursor);
      assert.equal(field('Formula or value').value, formula);
      assert.equal(document.activeElement, field('Go to cell'));
      assert.deepEqual(h.accepted, []);
      assert.deepEqual(h.history, []);
      assert.deepEqual(h.legacy, []);
      assert.deepEqual(h.data, before);
    }
  }
);
check(
  'Unmerged Special selection moves its active cursor to the first row-major match and Close preserves cursor and exact targets',
  async (t) => {
    const h = await mount(t);
    await go('A1:A3');
    assert.equal(field('Formula or value').value, 'Alpha');
    await selectSpecial('home.select-blanks');
    assert.equal(field('Go to cell').placeholder, 'A2');
    assert.equal(field('Formula or value').value, '');
    assert(node('home.fill-up').disabled);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.history, []);
    for (const key of Object.keys(h.data)) delete h.data[key];
    Object.assign(h.data, {
      B1: { value: 222, dataType: 'number' },
      A2: { value: 111, dataType: 'number' },
      C2: { value: true, dataType: 'boolean' },
    });
    await h.render({
      storageKey: 'cursor-fixture-head',
      sheets: h.input().sheets.map((s) => ({ ...s, cells: structuredClone(h.data) })),
    });
    await go('A1:C2');
    const before = structuredClone(h.data);
    await click('home.special-open');
    await click(specialButton('Preview matching cells'));
    await specialSettled();
    assert.deepEqual(
      [...field('Special selection targets').querySelectorAll('tbody tr')].map((r) => r.cells[0].textContent),
      ['Sheet1!B1', 'Sheet1!A2', 'Sheet1!C2']
    );
    await click(specialButton('Select matching cells'));
    assert.match(document.body.textContent, /Matching cells selected/);
    assert.equal(field('Go to cell').placeholder, 'B1');
    assert.equal(field('Formula or value').value, '222');
    await click(field('Close workbook tools'));
    assert.equal(field('Go to cell').placeholder, 'B1');
    assert.equal(field('Formula or value').value, '222');
    assert.equal(document.activeElement, field('Go to cell'));
    assert(node('home.fill-up').disabled);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.history, []);
    assert.deepEqual(h.data, before);
    await click('home.bold');
    assert.deepEqual(h.accepted, [
      [
        { type: 'formatRange', sheet: 'Sheet1', range: 'A2', format: { bold: true } },
        { type: 'formatRange', sheet: 'Sheet1', range: 'B1', format: { bold: true } },
        { type: 'formatRange', sheet: 'Sheet1', range: 'C2', format: { bold: true } },
      ],
    ]);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Special Apply requests a source-based distant reveal without cell focus, navigation or workbook changes (controlled offsets)',
  async (t) => {
    const h = await mount(t);
    for (const key of Object.keys(h.data)) delete h.data[key];
    h.data.A999 = { value: 999, dataType: 'number' };
    await h.render({
      storageKey: 'distant-source-head',
      sheets: h.input().sheets.map((s) => ({
        ...s,
        rows: 1004,
        cells: {},
        sheetFormat: { defaultRowHeight: 15, rowsHiddenByDefault: false },
        freeze: { rows: 0, columns: 0 },
      })),
    });
    await go('A1:A1004');
    const scroll = document.querySelector('.sheet-grid-scroll');
    assert.equal(scroll.scrollTop, 0);
    assert.equal(scroll.scrollLeft, 0);
    const before = structuredClone(h.data);
    await click('home.special-open');
    await click(specialButton('Preview matching cells'));
    await specialSettled();
    assert.deepEqual(
      [...field('Special selection targets').querySelectorAll('tbody tr')].map((row) => row.cells[0].textContent),
      ['Sheet1!A999']
    );
    const focusBefore = document.activeElement;
    await click(specialButton('Select matching cells'));
    // 999th row bottom 19980px - (600px mock viewport - 20px clearance -
    // 32px headings) + 8px reveal padding. This is not browser visibility proof.
    assert.equal(scroll.scrollTop, 19440);
    assert.equal(scroll.scrollLeft, 0);
    assert.equal(document.activeElement, focusBefore);
    assert.equal(field('Go to cell').placeholder, 'A999');
    assert.equal(field('Formula or value').value, '999');
    assert.match(document.body.textContent, /Matching cells selected/);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.history, []);
    assert.deepEqual(h.data, before);
    await click(field('Close workbook tools'));
    assert.equal(scroll.scrollTop, 19440);
    assert.equal(field('Go to cell').placeholder, 'A999');
    assert(node('home.fill-up').disabled);
    assert.equal(document.activeElement, field('Go to cell'));
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'toolbar and native clipboard events share bounded original reads rather than treating unloaded cells as empty',
  async (t) => {
    const h = await mount(t, {
      sheets: [
        { name: 'Sheet1', path: 'x', state: 'visible', rows: 1100, columns: 5, cells: {}, merges: [], widths: [] },
      ],
    });
    h.data.B1001 = { value: 'remote' };
    await h.render({ readRange: h.reader });
    await go('B1001:B1002');
    h.reads.length = 0;
    await click('home.copy');
    assert.equal(clipboard, '"remote"\n""');
    assert(h.reads.some((r) => r.range === 'B1001:B1002'));
    h.reads.length = 0;
    const cell = document.querySelector('.sheet-grid input'),
      root = document.querySelector('.spreadsheet-editor');
    await act(async () =>
      props(root).onCopy({
        target: cell,
        preventDefault() {},
        clipboardData: {
          setData() {
            throw Error('Synchronous cached copy bypassed host');
          },
        },
      })
    );
    await flush();
    assert.equal(clipboard, '"remote"\n""');
    assert(h.reads.some((r) => r.range === 'B1001:B1002'));
    assert(h.reads.every((r) => rangeBounds(r.range).size <= 1000));
    assert.deepEqual(h.accepted, []);
  }
);
check(
  'clipboard and target replies cannot retarget after navigation roundtrip, head, sheet or runtime replacement',
  async (t) => {
    const h = await mount(t, { readRange: async () => ({}) });
    for (const change of ['navigation', 'head', 'owner', 'sheet']) {
      await go('B1');
      const gate = defer();
      clipboardRead = () => gate.promise;
      const before = h.accepted.length;
      await click('home.paste');
      if (change === 'navigation') {
        await go('C2');
        await go('B1');
      } else if (change === 'head') await h.render({ storageKey: 'head-new' });
      else if (change === 'owner') await h.render({ ownerIdentity: 'new-runtime' });
      else await click([...document.querySelectorAll('button')].find((n) => n.textContent === 'Other'));
      await act(async () => gate.resolve('7'));
      await flush();
      assert.equal(h.accepted.length, before, change);
      assert.equal(document.querySelector('.spreadsheet-editor').dataset.homePending, 'false');
    }
    assert.deepEqual(h.legacy, []);
  }
);
check('late target-read failure stays on original scope and a real acknowledgement is awaited only once', async (t) => {
  const h = await mount(t);
  await go('B1');
  const read = defer();
  await h.render({ readRange: (_sheet, range) => (range === 'B1:B1' ? read.promise : Promise.resolve({})) });
  clipboard = '9';
  await click('home.paste');
  await h.render({ ownerIdentity: 'replacement' });
  await act(async () => read.reject(Error('Original read failed')));
  await flush();
  assert(!document.body.textContent.includes('Original read failed'));
  assert.deepEqual(h.accepted, []);
  const ack = defer();
  await h.render({
    readRange: undefined,
    onCommandBatch: (commands) => {
      h.accepted.push(commands);
      return ack.promise;
    },
  });
  await click('home.bold');
  assert.equal(h.accepted.length, 1);
  assert.equal(document.querySelector('.spreadsheet-editor').dataset.homePending, 'true');
  assert(node('home.italic').disabled);
  assert.equal(node('home.find-open').disabled, false);
  await click('home.find-open');
  assert(document.querySelector('[role="search"]'));
  await act(async () => ack.resolve());
  await flush();
  assert.equal(document.querySelector('.spreadsheet-editor').dataset.homePending, 'false');
  assert.deepEqual(h.legacy, []);
});
check(
  'retained old Home handlers do not park or disturb a new owner draft, error, focus/caret or pending state',
  async (t) => {
    const h = await mount(t),
      oldBold = props(node('home.bold')).onClick,
      oldMenu = props(node('home.menu')).onClick;
    await h.render({ ownerIdentity: 'replacement', documentId: 'other-doc' });
    const formula = field('Formula or value');
    await act(async () => {
      formula.focus();
      props(formula).onChange({ target: { value: '=SUM(' } });
    });
    formula.setSelectionRange(5, 5);
    await act(async () => {
      oldBold();
      oldMenu({
        preventDefault() {
          throw Error('Stale menu touched current default');
        },
      });
    });
    await flush();
    assert.equal(formula.value, '=SUM(');
    assert.equal(formula.selectionStart, 5);
    assert.equal(document.activeElement, formula);
    assert.equal(document.querySelector('.spreadsheet-editor').dataset.homePending, 'false');
    assert.equal(document.querySelector('.sheet-grid-error'), null);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
  }
);
check('normal accepted head advance stays success, while rejection never lands on a replacement owner', async (t) => {
  const h = await mount(t),
    ack = defer();
  await h.render({
    onCommandBatch: (commands) => {
      h.accepted.push(commands);
      return ack.promise;
    },
  });
  await click('home.bold');
  await h.render({ storageKey: 'accepted-head' });
  await act(async () => ack.resolve());
  await flush();
  assert.equal(document.querySelector('.sheet-grid-error'), null);
  const rejected = defer();
  await h.render({ onCommandBatch: () => rejected.promise });
  await click('home.italic');
  await h.render({ ownerIdentity: 'next-runtime' });
  await act(async () => rejected.reject(Error('Old accepted owner refusal')));
  await flush();
  assert(!document.body.textContent.includes('Old accepted owner refusal'));
  assert.deepEqual(h.legacy, []);
});
check(
  'literal Cut/Paste is one acknowledged text destination/source-clear batch; rejection preserves cut intent',
  async (t) => {
    const h = await mount(t, {
      literal: true,
      storageKey: undefined,
      acceptedCommandState: 'literal-head',
      ownerIdentity: 'literal-owner',
    });
    await go('A1');
    await click('home.cut');
    assert.equal(clipboard, '"Alpha"');
    await go('C3');
    await click('home.paste');
    assert.deepEqual(h.accepted.at(-1), [
      {
        type: 'setCells',
        sheet: 'Sheet1',
        cells: [
          { cell: 'C3', value: 'Alpha' },
          { cell: 'A1', value: null },
        ],
      },
    ]);
    assert(node('home.bold').disabled);
    await go('B1');
    await click('home.cut');
    await go('C4');
    const failure = defer();
    await h.render({ onCommandBatch: () => failure.promise });
    await click('home.paste');
    await act(async () => failure.reject(Error('Literal refused')));
    await flush();
    assert(document.body.textContent.includes('Literal refused'));
    await h.render({ onCommandBatch: async (commands) => h.accepted.push(commands) });
    await click('home.paste');
    assert.equal(h.accepted.at(-1)[0].cells.at(-1).cell, 'B1');
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'native copy+cut/paste and all Paste Special modes retain declared exact payloads and reject stale source heads',
  async (t) => {
    const h = await mount(t);
    await go('B1:B2');
    await click('home.copy');
    await go('D1');
    clipboardRead = async () => clipboard;
    await click('home.paste-values');
    assert.equal(h.accepted.at(-1)[0].type, 'pasteValues');
    await go('B1:B2');
    await click('home.copy');
    await go('D2');
    await click('home.paste-formulas');
    assert.equal(h.accepted.at(-1)[0].type, 'pasteFormulas');
    await go('B1');
    await click('home.copy');
    await go('D4');
    await click('home.paste-formatting');
    assert.equal(h.accepted.at(-1)[0].type, 'formatRange');
    await go('B1');
    await click('home.cut');
    await go('E4');
    await click('home.paste');
    assert.deepEqual(h.accepted.at(-1), [
      { type: 'fillRange', sheet: 'Sheet1', source: 'B1', range: 'E4' },
      { type: 'clearRange', sheet: 'Sheet1', range: 'B1' },
    ]);
    await go('B1');
    await click('home.copy');
    await go('D1');
    await h.render({ storageKey: 'changed-source-head' });
    const count = h.accepted.length;
    await click('home.paste-values');
    assert.equal(h.accepted.length, count);
    assert.match(document.body.textContent, /copied workbook changed/);
  }
);
check(
  'readonly/missing strict callbacks/inactive/passive capabilities refuse writes while safe inspection remains available',
  async (t) => {
    const h = await mount(t);
    for (const patch of [
      { readOnly: true },
      { readOnly: false, onCommandBatch: undefined },
      { onCommandBatch: async () => {}, ownerIdentity: undefined },
      { ownerIdentity: 'runtime', active: false },
    ]) {
      await h.render(patch);
      assert(node('home.bold').disabled);
      assert(node('home.paste').disabled);
      if (patch.active !== false) {
        assert.equal(node('home.find-open').disabled, false);
        await click('home.find-open');
      }
    }
    await h.render({
      active: true,
      readOnly: true,
      ownerIdentity: 'runtime',
      onCommandBatch: async () => {},
      conditionalControls: conditional(),
    });
    assert.equal(node('styles.conditional').disabled, false);
    await click('styles.conditional');
    assert.equal(document.querySelectorAll('aside.sheet-task-pane:not([hidden])').length, 1);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Home draft zone and conditional inspection preserve parked FormulaBar text/caret with zero writes',
  async (t) => {
    const h = await mount(t, { conditionalControls: conditional() }),
      formula = field('Formula or value');
    await act(async () => {
      formula.focus();
      props(formula).onChange({ target: { value: '=SUM(' } });
    });
    formula.setSelectionRange(5, 5);
    const menu = node('home.menu');
    assert(menu.closest('[data-formula-draft-zone]'));
    await act(async () => {
      props(formula).onBlur({ currentTarget: formula, relatedTarget: menu });
      menu.focus();
    });
    await click('styles.conditional');
    await click(
      [...document.querySelectorAll('button')].find((n) => n.getAttribute('aria-label') === 'Close workbook tools')
    );
    assert.equal(formula.value, '=SUM(');
    assert.equal(formula.selectionStart, 5);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Find drafts, read-only search, exact replacement preview/page and writes use the captured state without legacy acceptance',
  async (t) => {
    const h = await mount(t);
    await h.render({
      query: async (_sheet, text) => ({
        cells: text === 'Alpha' ? { A1: structuredClone(h.data.A1) } : {},
        matches: text === 'Alpha' ? 1 : 0,
        nextCursor: null,
      }),
    });
    await click('home.find-open');
    await change(field('Find'), 'Alpha');
    await click('home.find-search');
    assert.match(document.body.textContent, /1 \/ 1/);
    await change(field('Replace with'), 'Beta');
    await click('home.replace-preview');
    assert(document.querySelector('[aria-label="Replacement preview"]'));
    await click(document.querySelector('[aria-label="Close workbook tools"]'));
    await change(field('Found cell'), '0');
    await click('home.replace-one');
    assert.deepEqual(h.accepted.at(-1), [
      {
        type: 'setCells',
        sheet: 'Sheet1',
        cells: [{ cell: 'A1', expectedValue: 'Alpha', expectedFormula: null, expectedRawValue: null, value: 'Beta' }],
      },
    ]);
    assert.deepEqual(h.legacy, []);
  }
);
check('Format Painter deferred destination and drag-fill release use strict original gesture scope', async (t) => {
  const h = await mount(t);
  await go('B1');
  await click('home.format-painter');
  const cell = document.querySelector('.sheet-grid input[aria-label="Sheet1 C1"]'),
    td = cell.closest('td');
  await act(async () => props(td).onMouseDown({ shiftKey: false, preventDefault() {} }));
  await flush();
  assert.equal(h.accepted.at(-1)[0].type, 'formatRange');
  assert.equal(h.accepted.at(-1)[0].range, 'C1');
  await go('B1:B2');
  const grip = document.querySelector('.sheet-fill-handle');
  assert(grip);
  await act(async () => props(grip).onMouseDown({ preventDefault() {}, stopPropagation() {} }));
  const destination = document.querySelector('.sheet-grid input[aria-label="Sheet1 B4"]').closest('td');
  await act(async () => props(destination).onMouseEnter());
  await act(async () => window.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true })));
  await flush();
  assert.deepEqual(h.accepted.at(-1), [{ type: 'fillRange', sheet: 'Sheet1', source: 'B1:B2', range: 'B3:B4' }]);
  assert.deepEqual(h.legacy, []);
  await go('B1:B2');
  await act(async () =>
    props(document.querySelector('.sheet-fill-handle')).onMouseDown({ preventDefault() {}, stopPropagation() {} })
  );
  await act(async () => props(destination).onMouseEnter());
  await h.render({ storageKey: 'new-head' });
  const n = h.accepted.length;
  await act(async () => window.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true })));
  await flush();
  assert.equal(h.accepted.length, n);
});
check(
  'strict quick-access history and large-paste cancellation await original promises without source-side fire-and-forget',
  async (t) => {
    const h = await mount(t);
    const history = defer();
    await h.render({
      onCommandUndo: () => {
        h.history.push('held');
        return history.promise;
      },
    });
    await click('home.undo');
    assert.equal(document.querySelector('.spreadsheet-editor').dataset.homePending, 'true');
    assert(node('home.redo').disabled);
    await act(async () => history.resolve());
    await flush();
    assert.equal(document.querySelector('.spreadsheet-editor').dataset.homePending, 'false');
    await click('home.redo');
    assert.deepEqual(h.history, ['held', 'redo']);
    let options;
    const imported = defer();
    await h.render({
      importPaste: (_text, _inspection, value) => {
        options = value;
        return imported.promise;
      },
    });
    clipboard = Array.from({ length: 10001 }, () => '7').join('\n');
    await click('home.paste');
    assert(options);
    await click('home.cancel-paste');
    assert.equal(options.signal.aborted, true);
    await act(async () => imported.reject(Error('Paste cancelled')));
    await flush();
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
  }
);
