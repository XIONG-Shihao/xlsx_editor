import './support/scratch-dir.mjs';
/** Actual Grid/Home clipboard routes in React/JSDOM, controlled host I/O only. */
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
  output = await fs.mkdtemp(path.join(component, '.build/tests/clipboard-text-ui-')),
  checks = [];
const entry = path.join(output, 'entry.ts'),
  bundle = path.join(output, 'test.mjs');
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
let clipboard = '',
  clipboardReads = 0,
  clipboardWrites = [],
  clipboardRead = async () => clipboard,
  clipboardWrite = async (text) => {
    clipboard = text;
  };
ui.configureXlsxHost({
  api: {
    copyText: (text) => {
      clipboardWrites.push(text);
      return clipboardWrite(text);
    },
    readClipboardText: () => {
      clipboardReads++;
      return clipboardRead();
    },
  },
  Icon: () => React.createElement('svg'),
  useI18n: () => ({ locale: 'en' }),
});
const props = (node) => node[Object.keys(node).find((key) => key.startsWith('__reactProps$'))],
  field = (name) => document.querySelector(`[aria-label=${JSON.stringify(name)}]`),
  node = (id) =>
    [...document.querySelectorAll(`[data-home-command=${JSON.stringify(id)}]`)].find((n) => !n.closest('[hidden]'));
const flush = () => act(async () => new Promise((resolve) => setTimeout(resolve, 10))),
  defer = () => {
    let resolve, reject;
    const promise = new Promise((yes, no) => {
      resolve = yes;
      reject = no;
    });
    return { promise, resolve, reject };
  };
const click = async (id) => {
  const n = typeof id === 'string' ? node(id) : id;
  assert(n);
  assert(!n.disabled);
  await act(async () => n.click());
  await flush();
};
const change = async (n, value) => {
  assert(n);
  await act(async () => props(n).onChange({ target: { value } }));
  await flush();
};
async function go(value) {
  await change(field('Go to cell'), value);
  await act(async () =>
    field('Go to cell').dispatchEvent(
      new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })
    )
  );
  await flush();
}
const text = (value) => ({ value, dataType: 'text' });
const column = (index) => {
  let n = index + 1,
    result = '';
  while (n) {
    n--;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
};
async function mount(t, { data: provided, ...patch } = {}) {
  clipboard = '';
  clipboardReads = 0;
  clipboardWrites = [];
  clipboardRead = async () => clipboard;
  clipboardWrite = async (value) => {
    clipboard = value;
  };
  const root = createRoot(document.getElementById('root')),
    accepted = [],
    legacy = [],
    reads = [],
    imports = [];
  const data = provided ?? {
    A1: text('left\tright'),
    B1: text('say "hi"'),
    C1: text('line1\nline2'),
    D1: text("'00123"),
    A2: { value: null, dataType: 'blank' },
    C2: text(''),
    D2: text(' tail '),
    A3: { formula: 'SUM(B1:B3)', calculation: { status: 'calculated', value: 6 } },
    B3: { value: true, dataType: 'boolean' },
    C3: { value: '#N/A', dataType: 'error', error: '#N/A' },
    D3: text('00123'),
    G1: text('Move me'),
  };
  const sheet = (name) => ({
    name,
    path: 'xl/worksheets/' + name + '.xml',
    state: 'visible',
    rows: 1100,
    columns: 30,
    cells: structuredClone(data),
    merges: [],
    widths: [],
    hiddenRows: [],
    rowHeights: [],
    freeze: { rows: 0, columns: 0 },
  });
  let readOverride = null,
    mounted = true;
  const reader = async (name, range) => {
    reads.push({ name, range });
    if (readOverride) return readOverride(name, range);
    const b = rangeBounds(range),
      cells = {};
    for (let r = b.first.row; r <= b.last.row; r++)
      for (let c = b.first.col; c <= b.last.col; c++) {
        let n = c + 1,
          col = '';
        while (n) {
          n--;
          col = String.fromCharCode(65 + (n % 26)) + col;
          n = Math.floor(n / 26);
        }
        const ref = col + (r + 1);
        if (data[ref] !== undefined) cells[ref] = structuredClone(data[ref]);
      }
    return cells;
  };
  let input = {
    sheets: [sheet('Sheet1'), sheet('Other')],
    documentId: 'doc',
    ownerIdentity: 'runtime',
    storageKey: 'opaque-head',
    active: true,
    readOnly: false,
    onEdit: (value) => legacy.push(value),
    onBatch: (value) => legacy.push(value),
    onCommandBatch: async (commands) => {
      accepted.push(structuredClone(commands));
    },
    importPaste: async (...args) => imports.push(args),
    ...patch,
  };
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
  await render();
  t.after(unmount);
  return {
    data,
    accepted,
    legacy,
    reads,
    imports,
    reader,
    render,
    unmount,
    input: () => input,
    holdRead: (fn) => (readOverride = fn),
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
  'Copy quotes every rectangular field and preserves tabs, quotes, newlines, apostrophe, formulas and three blank presences',
  async (t) => {
    const h = await mount(t),
      before = structuredClone(h.data);
    await h.render({ readRange: h.reader });
    await go('A1:D3');
    await click('home.copy');
    assert.equal(
      clipboard,
      '"left\tright"\t"say ""hi"""\t"line1\nline2"\t"\'00123"\n""\t""\t""\t" tail "\n"=SUM(B1:B3)"\t"true"\t"#N/A"\t"00123"'
    );
    assert.equal(clipboardWrites.length, 1);
    assert(h.reads.some((r) => r.range === 'A1:D3'));
    assert.deepEqual(h.data, before);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Copy of an off-viewport unloaded selection reads the original rectangle and emits an explicit quoted blank',
  async (t) => {
    const h = await mount(t, {
      data: { B1001: text('remote') },
      sheets: [
        { name: 'Sheet1', path: 'x', state: 'visible', rows: 1100, columns: 30, cells: {}, merges: [], widths: [] },
      ],
    });
    await h.render({ readRange: h.reader });
    await go('B1001:B1002');
    await click('home.copy');
    assert.equal(clipboard, '"remote"\n""');
    assert(h.reads.some((r) => r.range === 'B1001:B1002'));
    await go('B1002');
    await click('home.copy');
    assert.equal(clipboard, '""');
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Paste as text forwards exact string literals without apostrophe stripping, formula/Boolean/error/numeric inference',
  async (t) => {
    const h = await mount(t);
    await go('F20');
    clipboard = "'00123\t=SUM(A1)\tTRUE\t#DIV/0!\t00123\t9007199254740993\t1.00\t  padded  \t😀\tcafe\u0301";
    await click('home.paste-text');
    assert.deepEqual(h.accepted, [
      [
        {
          type: 'setCells',
          sheet: 'Sheet1',
          cells: [
            { cell: 'F20', value: "'00123" },
            { cell: 'G20', value: '=SUM(A1)' },
            { cell: 'H20', value: 'TRUE' },
            { cell: 'I20', value: '#DIV/0!' },
            { cell: 'J20', value: '00123' },
            { cell: 'K20', value: '9007199254740993' },
            { cell: 'L20', value: '1.00' },
            { cell: 'M20', value: '  padded  ' },
            { cell: 'N20', value: '😀' },
            { cell: 'O20', value: 'cafe\u0301' },
          ],
        },
      ],
    ]);
    assert.equal(h.imports.length, 0);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Ragged holes and trailing empty fields produce dense explicit empty strings with quoted tabs/newlines intact',
  async (t) => {
    const h = await mount(t);
    await go('D10');
    clipboard = '"a\tb"\tc\t"line1\nline2"\nshort\n';
    await click('home.paste-text');
    assert.deepEqual(h.accepted.at(-1), [
      {
        type: 'setCells',
        sheet: 'Sheet1',
        cells: [
          { cell: 'D10', value: 'a\tb' },
          { cell: 'E10', value: 'c' },
          { cell: 'F10', value: 'line1\nline2' },
          { cell: 'D11', value: 'short' },
          { cell: 'E11', value: '' },
          { cell: 'F11', value: '' },
        ],
      },
    ]);
    await go('G20');
    clipboard = 'value\t\n""';
    await click('home.paste-text');
    assert.deepEqual(h.accepted.at(-1)[0].cells, [
      { cell: 'G20', value: 'value' },
      { cell: 'H20', value: '' },
      { cell: 'G21', value: '' },
      { cell: 'H21', value: '' },
    ]);
    assert.deepEqual(h.legacy, []);
  }
);
check('Empty OS clipboard is a pure no-op, while a copied quoted blank writes one empty text cell', async (t) => {
  const h = await mount(t),
    before = structuredClone(h.data);
  await go('D10');
  clipboard = '';
  await click('home.paste-text');
  assert.deepEqual(h.accepted, []);
  assert.deepEqual(h.data, before);
  assert.deepEqual(h.imports, []);
  assert.equal(document.querySelector('.spreadsheet-editor').dataset.homePending, 'false');
  clipboard = '""';
  await click('home.paste-text');
  assert.deepEqual(h.accepted, [[{ type: 'setCells', sheet: 'Sheet1', cells: [{ cell: 'D10', value: '' }] }]]);
  assert.deepEqual(h.legacy, []);
});
check(
  'The dense10000-position boundary accepts once;10001+ and ragged dense overflow refuse without staged-import fallback',
  async (t) => {
    const h = await mount(t);
    await go('A1');
    clipboard = Array.from({ length: 100 }, () => Array(100).fill('x').join('\t')).join('\n');
    await click('home.paste-text');
    assert.equal(h.accepted.length, 1);
    assert.equal(h.accepted[0].length, 10);
    assert(
      h.accepted[0].every(
        (command) => command.type === 'setCells' && command.sheet === 'Sheet1' && command.cells.length === 1000
      )
    );
    const cells = h.accepted[0].flatMap((command) => command.cells);
    assert.equal(cells.length, 10000);
    assert.equal(new Set(cells.map((c) => c.cell)).size, 10000);
    assert.deepEqual(cells[0], { cell: 'A1', value: 'x' });
    assert.deepEqual(cells.at(-1), { cell: 'CV100', value: 'x' });
    assert(
      cells.every(
        (c) =>
          typeof c.value === 'string' &&
          Object.keys(c).sort().join(',') === 'cell,value' &&
          rangeBounds(c.cell).first.row < 100 &&
          rangeBounds(c.cell).first.col < 100
      )
    );
    await go('A1');
    for (const payload of [
      Array.from({ length: 101 }, () => Array(100).fill('x').join('\t')).join('\n'),
      Array(6000).fill('x').join('\t') + '\nshort',
    ]) {
      clipboard = payload;
      await click('home.paste-text');
      assert.equal(h.accepted.length, 1);
      assert.match(document.body.textContent, /10,000 positions/);
    }
    assert.deepEqual(h.imports, []);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Clipboard reads captured before navigation roundtrip, head, owner or worksheet retirement never apply or retarget',
  async (t) => {
    for (const kind of ['navigation', 'head', 'owner', 'sheet']) {
      const h = await mount(t);
      await go('D10');
      const held = defer();
      clipboardRead = () => held.promise;
      await click('home.paste-text');
      if (kind === 'navigation') {
        await go('E11');
        await go('D10');
      } else if (kind === 'head') await h.render({ storageKey: 'new-head' });
      else if (kind === 'owner') await h.render({ ownerIdentity: 'new-runtime', documentId: 'new-doc' });
      else await click([...document.querySelectorAll('.sheet-tabs button')].find((n) => n.textContent === 'Other'));
      await act(async () => held.resolve('=A1'));
      await flush();
      assert.deepEqual(h.accepted, [], kind);
      assert.deepEqual(h.legacy, []);
      assert.equal(document.querySelector('.spreadsheet-editor').dataset.homePending, 'false');
      assert.equal(document.querySelector('.sheet-grid-error'), null);
      await h.unmount();
    }
  }
);
check(
  'A held original Copy source read retired by head/owner/navigation produces no clipboard output or new-owner error',
  async (t) => {
    for (const kind of ['head', 'owner', 'navigation']) {
      const h = await mount(t);
      await go('A1:B2');
      const held = defer();
      h.holdRead((_name, range) => (range === 'A1:B2' ? held.promise : {}));
      await h.render({ readRange: h.reader });
      await click('home.copy');
      if (kind === 'head') await h.render({ storageKey: 'new-head' });
      else if (kind === 'owner') await h.render({ ownerIdentity: 'replacement', documentId: 'replacement' });
      else {
        await go('C3');
        await go('A1:B2');
      }
      await act(async () => held.resolve({ A1: text('obsolete') }));
      await flush();
      assert.deepEqual(clipboardWrites, []);
      assert.deepEqual(h.accepted, []);
      assert.deepEqual(h.legacy, []);
      assert.equal(document.querySelector('.sheet-grid-error'), null);
      await h.unmount();
    }
  }
);
check(
  'Cut then Paste as text refuses without consuming intent; acknowledged ordinary Paste still moves and clears exact source once',
  async (t) => {
    const h = await mount(t),
      before = structuredClone(h.data);
    await go('G1');
    await click('home.cut');
    const copied = clipboard;
    assert.equal(copied, '"Move me"');
    await go('E4');
    await click('home.paste-text');
    assert.deepEqual(h.accepted, []);
    assert.equal(clipboard, copied);
    assert.deepEqual(h.data, before);
    assert.match(document.body.textContent, /Paste as text does not move a cut range/);
    await click('home.paste');
    assert.deepEqual(h.accepted.at(-1), [
      { type: 'fillRange', sheet: 'Sheet1', source: 'G1', range: 'E4' },
      { type: 'clearRange', sheet: 'Sheet1', range: 'G1' },
    ]);
    await go('F5');
    await click('home.paste-text');
    assert.deepEqual(h.accepted.at(-1), [
      { type: 'setCells', sheet: 'Sheet1', cells: [{ cell: 'F5', value: 'Move me' }] },
    ]);
    assert.deepEqual(h.legacy, []);
  }
);
check('A matching stale Copy head refuses literal Paste; an unrelated external clipboard remains usable', async (t) => {
  const h = await mount(t);
  await go('D3');
  await click('home.copy');
  await go('E4');
  await h.render({ storageKey: 'changed-source-head' });
  await click('home.paste-text');
  assert.deepEqual(h.accepted, []);
  assert.match(document.body.textContent, /The copied workbook changed/);
  clipboard = 'external=literal';
  await click('home.paste-text');
  assert.deepEqual(h.accepted.at(-1), [
    { type: 'setCells', sheet: 'Sheet1', cells: [{ cell: 'E4', value: 'external=literal' }] },
  ]);
  assert.deepEqual(h.legacy, []);
});
check(
  'Text write acknowledgement is awaited once and late failure does not overwrite a replacement formula draft',
  async (t) => {
    const held = defer(),
      h = await mount(t);
    await go('D10');
    await h.render({
      onCommandBatch: (commands) => {
        h.accepted.push(structuredClone(commands));
        return held.promise;
      },
    });
    clipboard = '=A1';
    const submit = props(node('home.paste-text')).onClick;
    await act(async () => {
      submit();
      submit();
    });
    await flush();
    assert.equal(h.accepted.length, 1);
    assert(node('home.paste-text').disabled);
    assert.equal(document.querySelector('.spreadsheet-editor').dataset.homePending, 'true');
    await h.render({ ownerIdentity: 'replacement', documentId: 'replacement' });
    const input = field('Formula or value');
    await act(async () => {
      input.focus();
      props(input).onChange({ target: { value: '=SUM(' } });
    });
    input.setSelectionRange(5, 5);
    await act(async () => held.reject(Error('Original text paste failed')));
    await flush();
    assert.equal(input.value, '=SUM(');
    assert.equal(input.selectionStart, 5);
    assert.equal(document.activeElement, input);
    assert(!document.body.textContent.includes('Original text paste failed'));
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Readonly/inactive/busy/passive or missing strict-owner capability refuses text writes; literal CSV remains exact text',
  async (t) => {
    const h = await mount(t),
      old = props(node('home.paste-text')).onClick;
    for (const patch of [
      { readOnly: true },
      { readOnly: false, busy: true },
      { busy: false, active: false },
      { active: true, onCommandBatch: undefined },
      { onCommandBatch: async () => {}, ownerIdentity: undefined },
      { ownerIdentity: 'runtime', passiveNotice: 'Original passive source' },
    ]) {
      await h.render(patch);
      const n = node('home.paste-text');
      assert(!n || n.disabled);
      if (n) assert(n.title);
      clipboard = '=A1';
      await act(async () => old());
      await flush();
      assert.deepEqual(h.accepted, []);
      assert.deepEqual(h.legacy, []);
    }
    await h.render({
      passiveNotice: undefined,
      readOnly: false,
      busy: false,
      active: true,
      ownerIdentity: 'literal-owner',
      literal: true,
      storageKey: undefined,
      acceptedCommandState: 'literal-head',
      onCommandBatch: async (commands) => h.accepted.push(structuredClone(commands)),
    });
    await go('D10');
    clipboard = "00123\t=SUM(A1)\tTRUE\t#N/A\t'quoted";
    await click('home.paste-text');
    assert.deepEqual(h.accepted.at(-1), [
      {
        type: 'setCells',
        sheet: 'Sheet1',
        cells: [
          { cell: 'D10', value: '00123' },
          { cell: 'E10', value: '=SUM(A1)' },
          { cell: 'F10', value: 'TRUE' },
          { cell: 'G10', value: '#N/A' },
          { cell: 'H10', value: "'quoted" },
        ],
      },
    ]);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Copy/Cut quoted payload above16MiB refuses before host write and preserves the prior Copy intent without arming Cut',
  async (t) => {
    const large = 'a'.repeat(32767),
      data = Object.fromEntries(Array.from({ length: 513 }, (_, index) => ['A' + (index + 1), text(large)]));
    data.G1 = text('safe prior copy');
    const h = await mount(t, { data });
    await h.render({ readRange: h.reader });
    await go('G1');
    await click('home.copy');
    assert.equal(clipboard, '"safe prior copy"');
    const prior = clipboard,
      priorWrites = clipboardWrites.slice();
    await go('A1:A513');
    const expectedQuoted = Array(513)
      .fill('"' + large + '"')
      .join('\n');
    assert.equal(large.length, 32767);
    assert(expectedQuoted.length < 64 * 1024 * 1024);
    assert(Buffer.byteLength(expectedQuoted, 'utf8') > 16 * 1024 * 1024);
    for (const action of ['home.copy', 'home.cut']) {
      await click(action);
      assert.deepEqual(clipboardWrites, priorWrites);
      assert.equal(clipboard, prior);
      assert.deepEqual(h.accepted, []);
      assert.match(document.body.textContent, /Copy text exceeds 16 MiB/);
      assert.equal(document.querySelector('.spreadsheet-editor').dataset.homePending, 'false');
    }
    await go('E4');
    await click('home.paste');
    assert.deepEqual(h.accepted, [[{ type: 'fillRange', sheet: 'Sheet1', source: 'G1', range: 'E4' }]]);
    assert.equal(h.data.G1.value, 'safe prior copy');
    assert.equal(h.data.A513.value, large);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Paste as text refuses actual serialized2MiB command overflow including JSON escaping and multibyte Unicode before acknowledgement',
  async (t) => {
    const h = await mount(t),
      before = structuredClone(h.data);
    await go('A1');
    const cases = [
      { fields: 100, value: 'a'.repeat(32767), label: 'ASCII' },
      { fields: 33, value: '\n'.repeat(32767), label: 'JSON-escaped newline' },
      { fields: 22, value: '\u0800'.repeat(32767), label: 'UTF-8 multibyte' },
    ];
    for (const item of cases) {
      const row = Array(item.fields).fill(item.value);
      clipboard = row.map((value) => '"' + value.replaceAll('"', '""') + '"').join('\t');
      const expected = [
          {
            type: 'setCells',
            sheet: 'Sheet1',
            cells: row.map((value, index) => ({ cell: column(index) + '1', value })),
          },
        ],
        serialized = JSON.stringify(expected);
      assert(item.fields <= 10000);
      assert(item.value.length === 32767);
      assert(Buffer.byteLength(clipboard, 'utf8') < 64 * 1024 * 1024);
      assert(Buffer.byteLength(serialized, 'utf8') > 2 * 1024 * 1024, item.label);
      if (item.label === 'JSON-escaped newline') assert(Buffer.byteLength(clipboard, 'utf8') < 2 * 1024 * 1024);
      if (item.label === 'UTF-8 multibyte') assert(serialized.length < 2 * 1024 * 1024);
      await click('home.paste-text');
      assert.deepEqual(h.accepted, [], item.label);
      assert.deepEqual(h.imports, [], item.label);
      assert.deepEqual(h.data, before);
      assert.match(document.body.textContent, /Paste as text exceeds its 2 MiB action limit/);
      assert.equal(document.querySelector('.spreadsheet-editor').dataset.homePending, 'false');
    }
    assert.deepEqual(h.legacy, []);
  }
);
