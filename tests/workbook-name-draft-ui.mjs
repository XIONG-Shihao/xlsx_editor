import './support/scratch-dir.mjs';
/** Actual source Grid/formula input/registry/picker; mocked owner acknowledgement. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createHash } from 'node:crypto';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
const component = fileURLToPath(new URL('../', import.meta.url)),
  output = await fs.mkdtemp(path.join(component, '.build/tests/name-draft-ui-')),
  checks = [];
const entry = path.join(output, 'entry.ts'),
  bundle = path.join(output, 'test.mjs');
await fs.writeFile(
  entry,
  `export {SpreadsheetGrid} from ${JSON.stringify(path.join(component, 'ui/SpreadsheetGrid.tsx'))};export {configureXlsxHost} from ${JSON.stringify(path.join(component, 'ui/host.tsx'))};export * from ${JSON.stringify(path.join(component, 'ui/namedFormulaDraft.ts'))};`
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
      return 1000;
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
    this.fn([{ contentRect: { width: 1000, height: 600 } }]);
  }
  disconnect() {}
};
globalThis.requestAnimationFrame = (callback) => setTimeout(callback, 0);
globalThis.cancelAnimationFrame = clearTimeout;
const { createRoot } = await import('react-dom/client'),
  ui = await import(pathToFileURL(bundle));
ui.configureXlsxHost({
  api: { copyText: async () => {}, readClipboardText: async () => '' },
  Icon: () => React.createElement('svg'),
  useI18n: () => ({ locale: 'en' }),
});
const files = [
  'ui/SpreadsheetGrid.tsx',
  'ui/SpreadsheetRibbon.tsx',
  'ui/ribbonCommands.ts',
  'ui/namedFormulaDraft.ts',
  'ui/WorkbookNamePicker.tsx',
  'ui/styles/01.css',
  'tests/workbook-name-draft-ui.mjs',
];
const pins = async () =>
  Promise.all(
    files.map(async (relative) => {
      const p = path.join(component, relative),
        bytes = await fs.readFile(p);
      return { relative, path: p, sha256: createHash('sha256').update(bytes).digest('hex'), bytes: bytes.length };
    })
  );
const before = await pins();
after(async () => {
  dom.window.close();
});
const check = (name, fn) => {
  test(name, { timeout: 10000 }, async () => {
    try {
      await fn();
    } catch (error) {
      await fs.writeFile(
        path.join(output, 'failure-' + checks.length + '.json'),
        JSON.stringify(checks.at(-1), null, 2)
      );
      throw error;
    }
  });
};
const names = [
  { name: 'Revenue', formula: "'Sheet1'!$B$1:$B$3" },
  { name: 'Tax', formula: "'Sheet1'!$B$1" },
  { name: 'Tax', sheet: 'Sheet1', formula: "'Sheet1'!$B$2" },
  { name: 'OtherLocal', sheet: 'Other Sheet', formula: "'Other Sheet'!$A$1" },
];
const sheet = (name) => ({
  name,
  state: 'visible',
  path: 'xl/worksheets/' + name + '.xml',
  rows: 5,
  columns: 4,
  cells: { B1: { value: 2 }, B2: { value: 4 }, B3: { value: 6 } },
  merges: [],
  widths: [],
  hiddenRows: [],
  rowHeights: [],
  freeze: { rows: 0, columns: 0 },
});
const props = (node) => node[Object.keys(node).find((key) => key.startsWith('__reactProps'))],
  field = () => document.querySelector('[aria-label="Formula or value"]'),
  command = () =>
    document.querySelector('.sheet-ribbon-tabpanel:not([hidden]) [data-ribbon-command="names.use-in-formula"]'),
  picker = () => document.querySelector('[role="dialog"][aria-label="Use in Formula"]'),
  button = (label) =>
    [...document.querySelectorAll('button')].find(
      (node) => node.getAttribute('aria-label') === label || node.textContent === label
    );
const flush = async () => act(async () => new Promise((resolve) => setTimeout(resolve, 10)));
async function mount(patch = {}) {
  const root = createRoot(document.getElementById('root')),
    legacy = [],
    accepted = [],
    reads = [];
  let input = {
    sheets: [sheet('Sheet1'), sheet('Other Sheet')],
    documentId: 'doc',
    ownerIdentity: 'runtime',
    storageKey: 'opaque-head',
    active: true,
    readOnly: false,
    namedRanges: names,
    onEdit: (value) => legacy.push(value),
    onBatch: (value) => legacy.push(value),
    onCommandBatch: async (commands) => accepted.push(commands),
    readRange: async (name, range) => {
      reads.push({ name, range });
      return Object.fromEntries(
        Object.entries(input.sheets.find((sheet) => sheet.name === name).cells).filter(
          ([ref]) => range === ref + ':' + ref || range !== 'A1:A1'
        )
      );
    },
    ...patch,
  };
  const render = async (patch) => {
    input = { ...input, ...patch };
    await act(async () => root.render(React.createElement(ui.SpreadsheetGrid, input)));
  };
  await render();
  await act(async () =>
    [...document.querySelectorAll('[role="tab"]')].find((node) => node.textContent === 'Formulas').click()
  );
  await flush();
  return { legacy, accepted, reads, render, input: () => input, close: async () => act(async () => root.unmount()) };
}
async function type(text, start = text.length, end = start) {
  await act(async () => {
    field().focus();
    props(field()).onChange({ target: { value: text } });
  });
  field().setSelectionRange(start, end);
}
async function key(key) {
  await act(async () =>
    field().dispatchEvent(new window.KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }))
  );
}
async function open() {
  const node = command();
  assert(node && !node.disabled);
  await act(async () => {
    const event = new window.MouseEvent('pointerdown', { bubbles: true, cancelable: true });
    node.dispatchEvent(event);
    assert(event.defaultPrevented, 'Name opener must prevent accidental blur');
    node.click();
  });
  await flush();
  assert(picker());
}
async function choose(label) {
  await act(async () => button(label).click());
  await flush();
  assert.equal(picker(), null);
}
check(
  'effective static names preserve quoting/local shadow while duplicate, invalid and unsupported definitions cannot be selected',
  () => {
    const list = ui.namedFormulaChoices(
      [
        ...names,
        { name: 'Invalid', formula: "'Sheet1'!$A$1", scopeInvalid: true },
        { name: 'Unknown', formula: "'Sheet1'!$A$1", sheet: undefined },
        { name: 'Relative', formula: 'Sheet1!A1' },
        { name: 'Dynamic', formula: 'OFFSET(A1,1,1)' },
        { name: 'Quoted', formula: "'O''Brien'!$A$1:$B$2" },
      ],
      'Sheet1',
      ['Sheet1', 'Other Sheet', "O'Brien"]
    );
    assert(list[1].disabledReason.includes('takes precedence'));
    assert.equal(list[2].disabledReason, undefined);
    assert(list[3].disabledReason.includes('Switch'));
    assert(list[4].disabledReason.includes('invalid'));
    assert(list[5].disabledReason.includes('invalid'));
    assert(list[6].disabledReason && list[7].disabledReason);
    assert.equal(list[8].formula, "'O''Brien'!$A$1:$B$2");
    assert.equal(list[8].disabledReason, undefined);
    const duplicate = ui.namedFormulaChoices(
      [
        { name: 'Rate', formula: "'Sheet1'!$A$1" },
        { name: 'rate', formula: "'Sheet1'!$A$2" },
      ],
      'Sheet1',
      ['Sheet1']
    );
    assert(duplicate.every((item) => item.disabledReason));
    assert.throws(() =>
      ui.namedFormulaChoices(
        Array.from({ length: 1001 }, () => names[0]),
        'Sheet1',
        ['Sheet1']
      )
    );
  }
);
check(
  'name insertion replaces exactly a caret/reference span, preserves surrounding draft and refuses strings/selectors/partial tokens',
  () => {
    const choice = ui.namedFormulaChoices([names[0]], 'Sheet1', ['Sheet1'])[0],
      snapshot = { text: '=SUM()+10', start: 5, end: 5, generation: 1, composing: false };
    assert.deepEqual(ui.insertFormulaName(snapshot, choice), { text: '=SUM(Revenue)+10', caret: 12 });
    assert.deepEqual(ui.insertFormulaName({ ...snapshot, text: '=SUM(B1:B3)+10', start: 5, end: 10 }, choice), {
      text: '=SUM(Revenue)+10',
      caret: 12,
    });
    assert.deepEqual(ui.insertFormulaName({ ...snapshot, text: '', start: 0, end: 0 }, choice), {
      text: '=Revenue',
      caret: 8,
    });
    for (const value of [
      { ...snapshot, text: 'text', start: 4, end: 4 },
      { ...snapshot, text: '="hi"', start: 3, end: 3 },
      { ...snapshot, text: '=Tab[Value]', start: 6, end: 6 },
      { ...snapshot, text: '=SUM(A1)', start: 2, end: 2 },
      { ...snapshot, composing: true },
    ])
      assert.throws(() => ui.insertFormulaName(value, choice));
  }
);
check(
  'actual formula input/pointer opener/caret replacement makes zero owner calls until one strict Enter acknowledgement',
  async () => {
    const view = await mount();
    try {
      await type('=SUM()+10', 5);
      await open();
      assert.equal(field().value, '=SUM()+10');
      assert.deepEqual(view.legacy, []);
      assert.deepEqual(view.accepted, []);
      await choose('Insert name Revenue · Workbook');
      assert.equal(field().value, '=SUM(Revenue)+10');
      assert.equal(field().selectionStart, 12);
      assert.equal(document.activeElement, field());
      assert.deepEqual(view.accepted, []);
      await key('Enter');
      await flush();
      assert.deepEqual(view.accepted, [[{ sheet: 'Sheet1', cell: 'A1', formula: 'SUM(Revenue)+10' }]]);
      assert.deepEqual(view.legacy, []);
      assert(view.reads.some((row) => row.range === 'A1:A1'));
      await key('Enter');
      await flush();
      assert.equal(view.accepted.length, 1);
    } finally {
      await view.close();
    }
  }
);
check(
  'keyboard Tab to name opener parks draft; Escape and Cancel return original text/caret without owner history',
  async () => {
    const view = await mount();
    try {
      await type('=SUM()+10', 5);
      await key('Tab');
      await act(async () => command().focus());
      assert.deepEqual(view.legacy, []);
      await act(async () => command().click());
      await flush();
      await act(async () =>
        picker().dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }))
      );
      await flush();
      assert.equal(field().value, '=SUM()+10');
      assert.equal(field().selectionStart, 5);
      assert.equal(document.activeElement, field());
      await open();
      await act(async () => button('Cancel name insertion').click());
      await flush();
      assert.equal(field().value, '=SUM()+10');
      assert.deepEqual(view.legacy, []);
      assert.deepEqual(view.accepted, []);
    } finally {
      await view.close();
    }
  }
);
check(
  'scope/head/owner/inactive/selection roundtrip retire captured picker and stale choice handlers cannot affect current target',
  async () => {
    for (const patch of [{ storageKey: 'new-head' }, { ownerIdentity: 'other-runtime' }, { active: false }]) {
      const view = await mount();
      try {
        await type('=SUM()', 5);
        await open();
        const old = props(button('Insert name Revenue · Workbook')).onClick;
        await view.render(patch);
        assert.equal(picker(), null);
        await act(async () => old());
        assert.deepEqual(view.legacy, []);
        assert.deepEqual(view.accepted, []);
        assert.equal(field().value, '');
      } finally {
        await view.close();
      }
    }
    const view = await mount();
    try {
      await type('=SUM()', 5);
      await open();
      const old = props(button('Insert name Revenue · Workbook')).onClick;
      await act(async () => {
        const control = document.querySelector('[aria-label="Worksheet"]');
        props(control).onChange({ target: { value: 'Other Sheet' } });
      });
      await act(async () => {
        const control = document.querySelector('[aria-label="Worksheet"]');
        props(control).onChange({ target: { value: 'Sheet1' } });
      });
      await act(async () => old());
      assert.deepEqual(view.accepted, []);
      assert.equal(picker(), null);
    } finally {
      await view.close();
    }
  }
);
check(
  'IME, readonly/literal/missing strict owner and another live input cannot launch an opportunistic edit',
  async () => {
    for (const patch of [
      { readOnly: true },
      { onCommandBatch: undefined },
      { ownerIdentity: '' },
      { literal: true, readRange: undefined },
    ]) {
      const view = await mount(patch);
      try {
        assert(!command() || command().disabled);
        assert.deepEqual(view.accepted, []);
      } finally {
        await view.close();
      }
    }
    const view = await mount();
    try {
      await type('=SUM()', 5);
      await act(async () => props(field()).onCompositionStart({ currentTarget: field() }));
      assert(command().disabled);
      await act(async () => props(field()).onCompositionEnd({ currentTarget: field() }));
      const other = document.querySelector('[aria-label="Go to cell"]');
      await act(async () => other.focus());
      const before = view.legacy.length;
      await open().catch(() => {});
      assert.equal(picker(), null);
      assert.equal(view.legacy.length, before);
      assert.deepEqual(view.accepted, []);
    } finally {
      await view.close();
    }
  }
);
check(
  'strict Enter waits once, preserves rejection, and a late captured read cannot retarget a replacement owner',
  async () => {
    let release;
    const held = new Promise((resolve) => (release = resolve)),
      calls = [],
      view = await mount({
        onCommandBatch: async (commands) => {
          calls.push(commands);
          await held;
          throw Error('Original owner refused');
        },
      });
    try {
      await type('=SUM()', 5);
      await open();
      await choose('Insert name Revenue · Workbook');
      await key('Enter');
      await key('Enter');
      await flush();
      assert.equal(calls.length, 1);
      assert.equal(field().value, '=SUM(Revenue)');
      await act(async () => release());
      await flush();
      assert.match(document.querySelector('.sheet-grid-error').textContent, /Original owner refused/);
      assert.deepEqual(view.legacy, []);
    } finally {
      release();
      await view.close();
    }
    let resolve,
      hold = false,
      readCalls = 0;
    const read = new Promise((done) => (resolve = done)),
      next = await mount({
        readRange: async (name, range) => {
          if (hold && range === 'A1:A1') {
            readCalls++;
            return read;
          }
          return {};
        },
      });
    try {
      await type('=SUM()', 5);
      await open();
      await choose('Insert name Revenue · Workbook');
      hold = true;
      await key('Enter');
      await flush();
      assert.equal(readCalls, 1);
      await next.render({
        ownerIdentity: 'new-runtime',
        readRange: async () => {
          throw Error('Wrong new reader');
        },
      });
      await act(async () => resolve({}));
      await flush();
      assert.deepEqual(next.accepted, []);
      assert.equal(field().value, '');
      assert.equal(document.querySelector('.sheet-grid-error'), null);
    } finally {
      resolve({});
      await next.close();
    }
  }
);
check('actual target read refuses protected spill/merge/Text-formatted targets without owner writes', async () => {
  for (const kind of ['spill', 'merge', 'text']) {
    const source = sheet('Sheet1');
    if (kind === 'spill') source.cells.A1 = { spill: { master: 'B1' } };
    if (kind === 'merge') source.merges = ['A1:B1'];
    if (kind === 'text') source.cells.A1 = { format: 49 };
    const view = await mount({
      sheets: [source],
      readRange: async (name, range) => (range === 'A1:A1' ? { A1: source.cells.A1 } : source.cells),
    });
    try {
      await type('=SUM()', 5);
      await open();
      await choose('Insert name Revenue · Workbook');
      await key('Enter');
      await flush();
      assert.equal(view.accepted.length, 0);
      assert.match(document.querySelector('.sheet-grid-error').textContent, /array|unmerged|number format/);
    } finally {
      await view.close();
    }
  }
});
check(
  'TRUE/FALSE and Unicode confusables cannot shadow a qualified ASCII name or serialize as intended named references',
  () => {
    const definitions = [
        { name: 'TRUE', formula: "'Sheet1'!$A$1" },
        { name: 'false', formula: "'Sheet1'!$A$2" },
        { name: 'Kelvin', formula: "'Sheet1'!$A$1" },
        { name: 'Kelvin', sheet: 'Sheet1', formula: "'Sheet1'!$A$2" },
      ],
      choices = ui.namedFormulaChoices(definitions, 'Sheet1', ['Sheet1']);
    assert(choices[0].disabledReason && choices[1].disabledReason && choices[3].disabledReason);
    assert.equal(choices[2].disabledReason, undefined);
    for (const name of ['TRUE', 'fAlSe'])
      assert.throws(() =>
        ui.insertFormulaName(
          { text: '=', start: 1, end: 1, generation: 0, composing: false },
          { id: name, name, formula: "'Sheet1'!$A$1", scope: 'Workbook' }
        )
      );
  }
);
check(
  'Alt+ArrowDown opens the actual Formulas dropdown directly from a draft without native blur/owner writes',
  async () => {
    const view = await mount();
    try {
      await act(async () =>
        [...document.querySelectorAll('[role="tab"]')].find((node) => node.textContent === 'Home').click()
      );
      await type('=SUM()+10', 5);
      await act(async () =>
        field().dispatchEvent(
          new window.KeyboardEvent('keydown', { key: 'ArrowDown', altKey: true, bubbles: true, cancelable: true })
        )
      );
      await flush();
      assert(picker());
      assert.equal(document.querySelector('[role="tab"][aria-selected="true"]').textContent, 'Formulas');
      assert.equal(field().value, '=SUM()+10');
      assert.deepEqual(view.legacy, []);
      assert.deepEqual(view.accepted, []);
      await choose('Insert name Revenue · Workbook');
      assert.equal(field().value, '=SUM(Revenue)+10');
      await key('Escape');
      assert.equal(field().value, '');
      assert.deepEqual(view.accepted, []);
    } finally {
      await view.close();
    }
  }
);
check(
  'close/reopen same scope retires the old restoration RAF instead of stealing new picker search focus',
  async () => {
    const view = await mount(),
      original = globalThis.requestAnimationFrame,
      held = [];
    try {
      await type('=SUM()', 5);
      await open();
      globalThis.requestAnimationFrame = (callback) => {
        held.push(callback);
        return 900;
      };
      await act(async () => button('Cancel name insertion').click());
      assert.equal(held.length, 1);
      await open();
      const search = document.querySelector('[aria-label="Search names"]');
      assert.equal(document.activeElement, search);
      await act(async () => held[0]());
      assert.equal(document.activeElement, search);
      assert.deepEqual(view.legacy, []);
      assert.deepEqual(view.accepted, []);
    } finally {
      globalThis.requestAnimationFrame = original;
      await view.close();
    }
  }
);
check(
  'invalid-scope existing manager rows are labelled and cannot update or remove a legitimate global of the same name',
  async () => {
    const view = await mount({
      nameCommands: true,
      namedRanges: [
        { name: 'Rate', formula: "'Sheet1'!$B$1" },
        { name: 'Rate', formula: "'Sheet1'!$B$2", scopeInvalid: true },
      ],
    });
    try {
      await act(async () => document.querySelector('[data-ribbon-command="names.manager"]').click());
      const rows = [...document.querySelector('[aria-label="Named ranges"]').children],
        bad = rows.find((node) => node.textContent.includes('Invalid worksheet scope'));
      assert(bad);
      const update = [...bad.querySelectorAll('button')].find((node) => node.textContent.startsWith('Use selection')),
        remove = [...bad.querySelectorAll('button')].find((node) => node.textContent.startsWith('Remove'));
      assert(update.disabled && remove.disabled);
      await act(async () => {
        props(update).onClick();
        props(remove).onClick();
      });
      assert.deepEqual(view.legacy, []);
      const good = rows.find((node) => !node.textContent.includes('Invalid worksheet scope'));
      assert.equal(
        [...good.querySelectorAll('button')].find((node) => node.textContent.startsWith('Use selection')).disabled,
        false
      );
    } finally {
      await view.close();
    }
  }
);
check(
  'worksheet-qualified prefix cannot silently retarget a chosen local name; entire qualified-reference replacement stays draft-only',
  async () => {
    const choices = ui.namedFormulaChoices(names, 'Sheet1', ['Sheet1', 'Other Sheet']),
      tax = choices[2];
    for (const text of ["='Other Sheet'!", "='Other Sheet'!  "]) {
      assert.throws(
        () =>
          ui.insertFormulaName({ text, start: text.length, end: text.length, generation: 0, composing: false }, tax),
        /entire worksheet-qualified reference/
      );
    }
    const text = "=SUM('Other Sheet'!$A$1)+10",
      a = text.indexOf("'Other"),
      b = text.indexOf(')');
    assert.equal(
      ui.insertFormulaName({ text, start: a, end: b, generation: 0, composing: false }, tax).text,
      '=SUM(Tax)+10'
    );
    const view = await mount();
    try {
      await type("='Other Sheet'!");
      await open();
      await act(async () => button('Insert name Tax · Sheet1').click());
      await flush();
      assert.equal(field().value, "='Other Sheet'!");
      assert(picker());
      assert.match(document.querySelector('.sheet-grid-error').textContent, /entire worksheet-qualified reference/);
      assert.deepEqual(view.legacy, []);
      assert.deepEqual(view.accepted, []);
    } finally {
      await view.close();
    }
  }
);
check(
  'accepted-head delayed viewport refresh restores the pristine focused formula bar without a second owner edit',
  async () => {
    let finishAcceptance, finishTile;
    const acceptance = new Promise((resolve) => (finishAcceptance = resolve)),
      tile = new Promise((resolve) => (finishTile = resolve)),
      calls = [];
    const view = await mount({
      onCommandBatch: async (commands) => {
        calls.push(commands);
        await acceptance;
      },
    });
    try {
      await type('=SUM()+10', 5);
      await open();
      await choose('Insert name Revenue · Workbook');
      await key('Enter');
      await flush();
      assert.equal(calls.length, 1);
      let tileReads = 0;
      await view.render({
        storageKey: 'acknowledged-head',
        readRange: async (name) => {
          if (name === 'Sheet1') {
            tileReads++;
            return tile;
          }
          return sheet(name).cells;
        },
      });
      await flush();
      assert(tileReads > 0);
      assert.equal(field().value, '');
      assert.equal(document.activeElement, field());
      await act(async () => finishAcceptance());
      await flush();
      const cells = {
        ...sheet('Sheet1').cells,
        A1: { value: 22, formula: 'SUM(Revenue)+10', calculation: { status: 'calculated', value: 22 } },
      };
      await act(async () => finishTile(cells));
      await flush();
      assert.equal(document.querySelector('[aria-label="Sheet1 A1"]').value, '22');
      assert.equal(field().value, '=SUM(Revenue)+10');
      assert.equal(document.activeElement, field());
      await key('Enter');
      await flush();
      assert.equal(calls.length, 1);
      assert.deepEqual(view.legacy, []);
    } finally {
      finishAcceptance();
      finishTile({});
      await view.close();
    }
  }
);
check('accepted-head delayed tile cannot overwrite a newly typed or composing formula draft', async () => {
  for (const composing of [false, true]) {
    let finishTile;
    const tile = new Promise((resolve) => (finishTile = resolve)),
      view = await mount();
    try {
      await act(async () => field().focus());
      await view.render({
        storageKey: 'new-accepted-head',
        readRange: async (name) => (name === 'Sheet1' ? tile : sheet(name).cells),
      });
      await flush();
      assert.equal(field().value, '');
      if (composing) await act(async () => props(field()).onCompositionStart({ currentTarget: field() }));
      else await type('=SUM(B1:B3)+100');
      await act(async () =>
        finishTile({
          ...sheet('Sheet1').cells,
          A1: { value: 22, formula: 'SUM(Revenue)+10', calculation: { status: 'calculated', value: 22 } },
        })
      );
      await flush();
      if (composing) {
        assert.equal(field().value, '');
        await act(async () => props(field()).onChange({ target: { value: '=SUM(税)' } }));
        await act(async () => props(field()).onCompositionEnd({ currentTarget: field() }));
        await flush();
        assert.equal(field().value, '=SUM(税)');
      } else assert.equal(field().value, '=SUM(B1:B3)+100');
      assert.deepEqual(view.accepted, []);
      assert.deepEqual(view.legacy, []);
      assert.equal(document.activeElement, field());
    } finally {
      finishTile({});
      await view.close();
    }
  }
});
