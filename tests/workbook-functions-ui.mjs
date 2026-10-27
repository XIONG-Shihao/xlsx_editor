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
  prefix = path.join(component, '.build/tests/functions-ui-');
await fs.mkdir(path.dirname(prefix), { recursive: true });
const runs = await fs.mkdtemp(prefix),
  entry = path.join(runs, 'entry.ts'),
  bundle = path.join(runs, 'test.mjs');
await fs.writeFile(
  entry,
  `export {WorkbookFunctions} from ${JSON.stringify(path.join(component, 'ui/WorkbookFunctions.tsx'))};export * from ${JSON.stringify(path.join(component, 'ui/functionInsertion.ts'))};export {SpreadsheetGrid} from ${JSON.stringify(path.join(component, 'ui/SpreadsheetGrid.tsx'))};export {configureXlsxHost} from ${JSON.stringify(path.join(component, 'ui/host.tsx'))};`
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
  {
    WorkbookFunctions,
    SpreadsheetGrid,
    configureXlsxHost,
    insertFunction,
    FunctionInsertionRetired,
    FunctionInsertionCancelled,
    targetFingerprint,
  } = await import(pathToFileURL(bundle));
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
const binding = () =>
  Object.freeze({
    documentId: 'doc-A',
    grantDocumentId: 'grant-A',
    ownerKey: 'worker-A',
    state: 'base',
    sheet: 'Sheet1',
    target: 'C33',
    selection: 'C33',
    navigationGeneration: 1,
    activationEpoch: 0,
  });
const scope = () => ({ active: true, readOnly: false, busy: false, literal: false, specialSelection: false });
const visible = (node) => !node.closest('[hidden]');
const button = (name) =>
  [...document.querySelectorAll('button')].find(
    (node) => visible(node) && (node.getAttribute('aria-label') ?? node.textContent) === name
  );
const field = (name) =>
  [...document.querySelectorAll('input,select')].find(
    (node) => visible(node) && node.getAttribute('aria-label') === name
  );
async function input(node, value) {
  assert(node);
  await act(async () => {
    Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(node, value);
    node.dispatchEvent(new window.Event('input', { bubbles: true }));
  });
}
async function select(node, value) {
  assert(node);
  await act(async () => {
    node.value = value;
    node.dispatchEvent(new window.Event('change', { bubbles: true }));
  });
}
async function click(node) {
  assert(node);
  assert.equal(node.disabled, false);
  await act(async () => node.click());
}
async function mount(t, options = {}) {
  window.localStorage.clear();
  const root = createRoot(document.getElementById('root'));
  let current = true,
    target = options.target ?? {},
    available = scope(),
    read = options.read ?? (async () => target);
  const writes = [],
    outcomes = [];
  const context = {
    binding: binding(),
    isCurrent: () => current,
    availability: () => available,
    readTarget: () => read(),
    accept: options.accept ?? (async (commands) => writes.push(commands)),
  };
  let props = { context, visible: true, onOutcome: (value) => outcomes.push(value) };
  await act(async () => root.render(React.createElement(WorkbookFunctions, props)));
  t.after(async () => act(async () => root.unmount()));
  return {
    context,
    writes,
    outcomes,
    changeCurrent: (value) => {
      current = value;
    },
    changeScope: (patch) => {
      available = { ...available, ...patch };
    },
    target: (value) => {
      target = value;
    },
    setRead: (value) => {
      read = value;
    },
    render: async (patch) => {
      props = { ...props, ...patch };
      await act(async () => root.render(React.createElement(WorkbookFunctions, props)));
    },
  };
}
test('pane browses actual populated categories and typed text yields an escaped closed formula preview', async (t) => {
  await mount(t);
  assert.equal(
    [...field('Function category').options].some((item) => item.text === 'Financial'),
    true
  );
  assert.equal(field('Function category').options.length, 9);
  await input(field('Search functions'), 'LEN');
  assert.equal(field('Function').value, 'LEN');
  await select(field('Text type'), 'text');
  await input(field('Text'), 'a"b');
  assert.equal(document.querySelector('output[aria-label="Formula preview"]').textContent, '=LEN("a""b")');
  assert.equal(button('Insert formula into C33').disabled, false);
});
test('number drafts refuse silent rounding and optional blank arguments are omitted', async (t) => {
  await mount(t);
  await select(field('Function'), 'ABS');
  await select(field('Number type'), 'number');
  await input(field('Number'), '0.10000000000000001');
  assert.equal(button('Insert formula into C33').disabled, true);
  await input(field('Number'), '1.25');
  assert.equal(document.querySelector('output[aria-label="Formula preview"]').textContent, '=ABS(1.25)');
  await select(field('Function'), 'LEFT');
  await select(field('Text type'), 'text');
  await input(field('Text'), 'hello');
  await click(button('Add argument'));
  assert.equal(document.querySelector('output[aria-label="Formula preview"]').textContent, '=LEFT("hello")');
  assert.equal(button('Insert formula into C33').disabled, false);
});
test('ordinary zero false and empty text show explicit replacement actions without an extra confirmation', async (t) => {
  for (const value of [0, false, ''])
    await t.test(JSON.stringify(value), async (child) => {
      const ui = await mount(child, { target: { cell: { value } } });
      await input(field('Values 1'), 'A1:A2');
      assert.match(
        document.querySelector('output[aria-label="Existing target"]').textContent,
        /Number: 0|Boolean: false|Text: ""/
      );
      assert.equal(document.querySelector('input[type="checkbox"]'), null);
      await click(button('Replace C33 with formula'));
      assert.deepEqual(ui.writes, [
        [{ type: 'setCells', sheet: 'Sheet1', cells: [{ cell: 'C33', formula: 'SUM(A1:A2)' }] }],
      ]);
    });
});
test('fresh target preflight refuses a changed occupied cell before the captured owner callback', async (t) => {
  const ui = await mount(t, { target: { cell: { value: 9 } } });
  await input(field('Values 1'), 'A1:A2');
  ui.target({ cell: { value: 10 } });
  await click(button('Replace C33 with formula'));
  assert.deepEqual(ui.writes, []);
  assert.match(document.body.textContent, /Review the existing cell/);
});
test('Cancel retires a delayed preflight and Check target permits a new insertion without history', async (t) => {
  const gate = deferred(),
    ui = await mount(t);
  await input(field('Values 1'), 'A1:A2');
  ui.setRead(() => gate.promise);
  await click(button('Insert formula into C33'));
  await click(button('Cancel insertion'));
  await act(async () => gate.resolve({}));
  assert.deepEqual(ui.writes, []);
  assert.match(document.body.textContent, /Insertion cancelled/);
  ui.setRead(async () => ({}));
  await click(button('Check target'));
  await click(button('Insert formula into C33'));
  assert.equal(ui.writes.length, 1);
});
test('protected array/spill/merge targets and read-only or special mode never reach acceptance', async (t) => {
  for (const [name, target] of [
    ['array', { cell: { value: null, formulaType: 'array' } }],
    [
      'formula-output',
      { cell: { value: 'wrld', dataType: 'text', formulaOutput: { schema: 1, master: 'C32', range: 'C32:C33' } } },
    ],
    ['spill', { cell: { value: 1, spill: { schema: 1, master: 'C33', range: 'C33:D34', role: 'follower' } } }],
    ['merge', { protectedReason: 'Choose an unmerged cell.' }],
  ])
    await t.test(name, async (child) => {
      const ui = await mount(child, { target });
      await input(field('Values 1'), 'A1:A2');
      assert.equal(document.querySelector('.sheet-functions button[type="submit"]').disabled, true);
      assert.deepEqual(ui.writes, []);
    });
  const ui = await mount(t);
  await input(field('Values 1'), 'A1:A2');
  ui.changeScope({ readOnly: true });
  await ui.render({});
  assert.equal(button('Insert formula into C33').disabled, true);
  ui.changeScope({ readOnly: false, specialSelection: true });
  await ui.render({});
  assert.equal(button('Insert formula into C33').disabled, true);
  assert.deepEqual(ui.writes, []);
});
test('foreign scope immediately masks target/draft and delayed reads cannot insert or display a stale error', async (t) => {
  const gate = deferred(),
    ui = await mount(t);
  await input(field('Values 1'), 'A1:A2');
  ui.setRead(() => gate.promise);
  await click(button('Insert formula into C33'));
  ui.changeCurrent(false);
  await ui.render({});
  assert.equal(document.querySelector('output[aria-label="Existing target"]'), null);
  assert.match(document.body.textContent, /Reopen Insert Function/);
  await act(async () => gate.reject(Error('Old owner read failed')));
  assert.deepEqual(ui.writes, []);
  assert(!document.body.textContent.includes('Old owner read failed'));
});
test('one strict owner acceptance is awaited; Finishing is not UI cancellation and late acceptance keeps captured truth', async (t) => {
  const owner = deferred(),
    writes = [],
    ui = await mount(t, {
      accept: (commands) => {
        writes.push(commands);
        return owner.promise;
      },
    });
  await input(field('Values 1'), 'A1:A2');
  await click(button('Insert formula into C33'));
  assert.equal(writes.length, 1);
  assert.equal(ui.outcomes.length, 0);
  assert.equal(button('Finishing…').disabled, true);
  ui.changeCurrent(false);
  await ui.render({});
  await act(async () => owner.resolve());
  assert.equal(ui.outcomes.length, 1);
  assert.equal(ui.outcomes[0].outcome, 'accepted');
  assert.equal(ui.outcomes[0].current, false);
  assert.equal(ui.outcomes[0].binding.target, 'C33');
  assert(!document.body.textContent.includes('Inserted into'));
});
test('current owner refusal leaves the draft retryable while an old owner failure cannot surface in a new scope', async (t) => {
  for (const stale of [false, true])
    await t.test(stale ? 'late stale refusal' : 'current refusal', async (child) => {
      const owner = deferred(),
        ui = await mount(child, { accept: () => owner.promise });
      await input(field('Values 1'), 'A1:A2');
      await click(button('Insert formula into C33'));
      if (stale) {
        ui.changeCurrent(false);
        await ui.render({});
      }
      await act(async () => owner.reject(Error('Owner protected this formula')));
      assert.equal(ui.outcomes.length, 0);
      if (stale) assert(!document.body.textContent.includes('Owner protected this formula'));
      else {
        assert.match(document.querySelector('[role="alert"]').textContent, /Owner protected/);
        assert.equal(button('Insert formula into C33').disabled, false);
        assert(!document.body.textContent.includes('Inserting formula…'));
      }
    });
});
test('retained pane hiding cannot latch Finishing after the exact pending owner operation settles', async (t) => {
  for (const failed of [false, true])
    await t.test(failed ? 'hidden refusal' : 'hidden acceptance', async (child) => {
      const owner = deferred(),
        ui = await mount(child, { accept: () => owner.promise });
      await input(field('Values 1'), 'A1:A2');
      await click(button('Insert formula into C33'));
      assert.equal(button('Finishing…').disabled, true);
      await ui.render({ visible: false });
      await ui.render({ visible: true });
      assert.equal(button('Finishing…').disabled, true);
      await act(async () => (failed ? owner.reject(Error('Hidden owner refusal')) : owner.resolve()));
      if (failed) {
        assert.equal(button('Cancel insertion').disabled, false);
        assert.equal(button('Insert formula into C33').disabled, false);
      } else {
        assert.match(document.body.textContent, /Inserted into Sheet1!C33/);
        assert.equal(button('Insert formula into C33'), undefined);
      }
      assert.equal(ui.outcomes.length, failed ? 0 : 1);
    });
});
const sheet = () => ({
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
});
async function go(range) {
  await input(field('Go to cell'), range);
  await act(async () =>
    field('Go to cell').dispatchEvent(
      new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })
    )
  );
}
async function mountGrid(t, extra = {}) {
  window.localStorage.clear();
  const root = createRoot(document.getElementById('root')),
    writes = [],
    reads = [],
    props = {
      sheets: [sheet()],
      documentId: 'doc-A',
      ownerIdentity: 'owner-A',
      storageKey: 'owner-A:base',
      readOnly: false,
      onEdit: () => {},
      onCommandBatch: async (commands) => writes.push(commands),
      readRange: async (_sheet, range) => {
        reads.push(range);
        return {};
      },
      ...extra,
    };
  await act(async () => root.render(React.createElement(SpreadsheetGrid, props)));
  t.after(async () => act(async () => root.unmount()));
  const tab = (name) => [...document.querySelectorAll('[role="tab"]')].find((node) => node.textContent === name);
  return {
    writes,
    reads,
    open: async () => {
      await click(tab('Formulas'));
      await click(button('Insert Function'));
    },
    render: async (patch) => act(async () => root.render(React.createElement(SpreadsheetGrid, { ...props, ...patch }))),
  };
}
test('actual source Grid opens one docked function pane, reads offscreen target and submits exactly one captured command', async (t) => {
  const ui = await mountGrid(t);
  await go('C33');
  await ui.open();
  assert.equal(document.querySelectorAll('aside.sheet-task-pane:not([hidden])').length, 1);
  assert.equal(document.querySelector('[data-workbook-panel="functions"] section').dataset.functionTarget, 'C33');
  assert.equal(ui.reads.filter((range) => range === 'C33:C33').length, 1);
  await input(field('Values 1'), 'A1:A2');
  await click(button('Insert formula into C33'));
  assert.equal(ui.reads.filter((range) => range === 'C33:C33').length, 3);
  assert.deepEqual(ui.writes, [
    [{ type: 'setCells', sheet: 'Sheet1', cells: [{ cell: 'C33', formula: 'SUM(A1:A2)' }] }],
  ]);
});
test('actual Grid navigation away and back cannot revive a pending function draft or retarget the original acceptance', async (t) => {
  const gate = deferred(),
    writes = [];
  let targetReads = 0;
  const ui = await mountGrid(t, {
    onCommandBatch: async (commands) => writes.push(commands),
    readRange: async (_sheet, range) => (range === 'C33:C33' && ++targetReads === 2 ? gate.promise : {}),
  });
  await go('C33');
  await ui.open();
  await input(field('Values 1'), 'A1:A2');
  await click(button('Insert formula into C33'));
  await go('A1');
  await go('C33');
  assert.match(document.body.textContent, /Reopen Insert Function/);
  await act(async () => gate.resolve({}));
  assert.deepEqual(writes, []);
  await click(button('Close workbook tools'));
  await ui.open();
  await input(field('Values 1'), 'A1:A2');
  await click(button('Insert formula into C33'));
  assert.equal(writes.length, 1);
});
test('actual source Grid fences pending insertion across head, worksheet round-trip, activation and same-range special selection', async (t) => {
  for (const change of ['head', 'worksheet', 'activation', 'special'])
    await t.test(change, async (child) => {
      const gate = deferred(),
        writes = [];
      let count = 0;
      const sheets = [sheet(), { ...sheet(), name: 'Other', path: 'xl/worksheets/sheet2.xml' }];
      const ui = await mountGrid(child, {
        sheets,
        onCommandBatch: async (commands) => writes.push(commands),
        readRange: async (name, range) =>
          name === 'Sheet1' && range === 'C33:C33' && ++count === 2 ? gate.promise : {},
      });
      await go(change === 'special' ? 'C33:C34' : 'C33');
      await ui.open();
      await input(field('Values 1'), 'A1:A2');
      await click(button('Insert formula into C33'));
      if (change === 'head') await ui.render({ storageKey: 'owner-A:next' });
      if (change === 'worksheet') {
        await click([...document.querySelectorAll('.sheet-tabs button')].find((node) => node.textContent === 'Other'));
        await click([...document.querySelectorAll('.sheet-tabs button')].find((node) => node.textContent === 'Sheet1'));
        await go('C33');
      }
      if (change === 'activation') {
        await ui.render({ active: false });
        await ui.render({ active: true });
      }
      if (change === 'special') {
        await click([...document.querySelectorAll('[role="tab"]')].find((node) => node.textContent === 'Home'));
        await click(button('Editing'));
        const menu = [...document.querySelectorAll('details')].find(
          (node) => visible(node) && node.textContent.includes('Find & Select')
        );
        assert(menu);
        menu.open = true;
        await click(button('Select blanks'));
      }
      await act(async () => gate.resolve({}));
      assert.deepEqual(writes, []);
      assert.match(
        document.body.textContent,
        change === 'special' ? /Open Insert Function before inserting|Reopen Insert Function/ : /Reopen Insert Function/
      );
    });
});
test('actual source Grid refuses merged/protected/spill targets and a foreign-cell owner reply', async (t) => {
  for (const kind of ['merge', 'array', 'spill', 'formula-output', 'foreign'])
    await t.test(kind, async (child) => {
      const writes = [],
        metadata = sheet();
      if (kind === 'merge') metadata.merges = ['C33:D34'];
      const ui = await mountGrid(child, {
        sheets: [metadata],
        onCommandBatch: async (commands) => writes.push(commands),
        readRange: async (_sheet, range) =>
          range === 'C33:C33'
            ? kind === 'foreign'
              ? { D33: { value: 9 } }
              : {
                  C33:
                    kind === 'array'
                      ? { value: null, formulaType: 'array' }
                      : kind === 'formula-output'
                        ? {
                            value: 'wrld',
                            dataType: 'text',
                            formulaOutput: { schema: 1, master: 'C32', range: 'C32:C33' },
                          }
                        : kind === 'spill'
                          ? { value: 1, spill: { schema: 1, master: 'C33', range: 'C33:D34', role: 'master' } }
                          : { value: 2 },
                }
            : {},
      });
      await go('C33');
      await ui.open();
      await input(field('Values 1'), 'A1:A2');
      assert.equal(document.querySelector('.sheet-functions button[type="submit"]').disabled, true);
      assert.deepEqual(writes, []);
      assert.match(
        document.body.textContent,
        kind === 'foreign'
          ? /target reply did not match/
          : kind === 'merge'
            ? /unmerged/
            : kind === 'array'
              ? /protected formula/
              : kind === 'formula-output'
                ? /formula output/
                : /array output/
      );
    });
});
test('actual source Grid retains the originally captured strict callback during delayed preflight and harmless rerender', async (t) => {
  const gate = deferred(),
    original = [],
    replacement = [];
  let count = 0;
  const ui = await mountGrid(t, {
    onCommandBatch: async (commands) => original.push(commands),
    readRange: async (_sheet, range) => (range === 'C33:C33' && ++count === 2 ? gate.promise : {}),
  });
  await go('C33');
  await ui.open();
  await input(field('Values 1'), 'A1:A2');
  await click(button('Insert formula into C33'));
  await ui.render({ onCommandBatch: async (commands) => replacement.push(commands) });
  await act(async () => gate.resolve({}));
  assert.equal(original.length, 1);
  assert.equal(replacement.length, 0);
});
test('actual source Grid read-only and missing strict callbacks permit catalog inspection but disable insertion', async (t) => {
  for (const kind of ['read-only', 'no-strict'])
    await t.test(kind, async (child) => {
      const ui = await mountGrid(child, {
        ...(kind === 'read-only' ? { readOnly: true } : { onCommandBatch: undefined }),
      });
      await go('C33');
      await ui.open();
      await input(field('Values 1'), 'A1:A2');
      assert.equal(button('Insert formula into C33').disabled, true);
      assert.match(
        document.body.textContent,
        kind === 'read-only' ? /This workbook is read-only/ : /Workbook commands are unavailable/
      );
      assert.deepEqual(ui.writes, []);
    });
});
test('delimited and passive source Grid profiles never expose the function insertion workflow', async (t) => {
  for (const props of [{ literal: true }, { passiveNotice: 'Original format is preserved' }])
    await t.test(JSON.stringify(props), async (child) => {
      await mountGrid(child, props);
      assert.equal(document.querySelector('[data-workbook-tool="functions"]'), null);
      assert.equal(document.querySelector('[data-ribbon-command="formulas.insert-function"]'), null);
    });
});
test('accepted state advancement shows scoped insertion success and actual typed statuses without enabling the old draft', async (t) => {
  for (const calculation of [
    { status: 'calculated', value: false },
    { status: 'calculated', value: '#N/A' },
    { status: 'error', error: '#N/A' },
    { status: 'circular', error: '#CIRC!' },
    { status: 'unresolved', error: 'Unsupported reference' },
  ])
    await t.test(JSON.stringify(calculation), async (child) => {
      const owner = deferred(),
        writes = [];
      const ui = await mountGrid(child, {
        onCommandBatch: (commands) => {
          writes.push(commands);
          return owner.promise;
        },
      });
      await go('C33');
      await ui.open();
      await input(field('Values 1'), 'A1:A2');
      await click(button('Insert formula into C33'));
      const cell = { value: null, formula: 'SUM(A1:A2)', calculation };
      await ui.render({
        storageKey: 'owner-A:accepted1',
        readRange: async (_sheet, range) => (range === 'C33:C33' ? { C33: cell } : {}),
      });
      assert.match(document.body.textContent, /Finishing insertion/);
      await act(async () => owner.resolve());
      assert.match(document.body.textContent, /Inserted into Sheet1!C33/);
      assert.equal(button('Insert formula into C33'), undefined);
      assert.equal(writes.length, 1);
      const result = document.querySelector('[aria-label="Accepted formula result"]');
      assert.equal(result.dataset.status, calculation.status);
      if (calculation.status === 'calculated')
        assert.match(
          result.textContent,
          calculation.value === false ? /Calculated \(boolean\): false/ : /Calculated \(string\): "#N\/A"/
        );
      else
        assert.match(
          result.textContent,
          new RegExp(
            calculation.status === 'error'
              ? 'Error: #N/A'
              : calculation.status === 'circular'
                ? 'Circular reference'
                : 'Unresolved formula'
          )
        );
      await click(button('Start another insertion'));
      assert.equal(field('Values 1').value, '');
      assert.match(document.querySelector('output[aria-label="Existing target"]').textContent, /Formula: =SUM/);
    });
});
test('mismatched or late accepted-cell reads never invent a result or expose success after owner replacement', async (t) => {
  const owner = deferred(),
    late = deferred();
  let readCount = 0;
  const ui = await mountGrid(t, { onCommandBatch: () => owner.promise });
  await go('C33');
  await ui.open();
  await input(field('Values 1'), 'A1:A2');
  await click(button('Insert formula into C33'));
  await ui.render({
    storageKey: 'owner-A:accepted1',
    readRange: async (_sheet, range) =>
      range === 'C33:C33'
        ? { C33: { value: 999, formula: 'SUM(A1:A3)', calculation: { status: 'calculated', value: 999 } } }
        : {},
  });
  await act(async () => owner.resolve());
  assert.match(document.body.textContent, /Inserted into Sheet1!C33/);
  assert.match(document.querySelector('[aria-label="Accepted formula result"]').textContent, /Result unavailable/);
  assert(!document.body.textContent.includes('Calculated (number): 999'));
  await ui.render({
    storageKey: 'owner-A:accepted2',
    readRange: async (_sheet, range) => (range === 'C33:C33' ? (++readCount === 1 ? late.promise : {}) : {}),
  });
  assert.match(document.querySelector('[aria-label="Accepted formula result"]').textContent, /Reading|unavailable/);
  await ui.render({ ownerIdentity: 'owner-new', storageKey: 'owner-new:base' });
  assert.equal(document.querySelector('[aria-label="Accepted formula result"]'), null);
  await act(async () =>
    late.resolve({ C33: { value: 3, formula: 'SUM(A1:A2)', calculation: { status: 'calculated', value: 3 } } })
  );
  assert(!document.body.textContent.includes('Inserted into Sheet1!C33'));
});
const recentIds = () => JSON.parse(window.localStorage.getItem('xlsx.function-recents.v1') ?? '{"ids":[]}').ids;
test('Recently Used route uses the same pane and persists only accepted known function IDs, never arguments', async (t) => {
  const ui = await mountGrid(t);
  await go('C33');
  await ui.open();
  await input(field('Values 1'), 'A1:A2');
  assert.deepEqual(recentIds(), []);
  await click(button('Insert formula into C33'));
  assert.deepEqual(recentIds(), ['SUM']);
  await click(button('Recently Used'));
  assert.equal(document.querySelectorAll('aside.sheet-task-pane:not([hidden])').length, 1);
  assert.equal(field('Function category').value, 'recent');
  assert.deepEqual(
    [...field('Function').options].map((option) => option.value),
    ['SUM']
  );
  assert.equal(field('Values 1').value, '');
  assert.deepEqual(JSON.parse(window.localStorage.getItem('xlsx.function-recents.v1')), { version: 1, ids: ['SUM'] });
});
test('function preview, local cancellation and owner refusal do not create a recent entry', async (t) => {
  for (const action of ['preview', 'cancel', 'refuse'])
    await t.test(action, async (child) => {
      const gate = deferred(),
        ui = await mount(child, {
          accept: async () => {
            throw Error('Owner refused the input');
          },
        });
      await input(field('Values 1'), 'A1:A2');
      if (action === 'cancel') {
        ui.setRead(() => gate.promise);
        await click(button('Insert formula into C33'));
        await click(button('Cancel insertion'));
        await act(async () => gate.resolve({}));
      }
      if (action === 'refuse') await click(button('Insert formula into C33'));
      assert.deepEqual(recentIds(), []);
    });
});
test('durable late acceptance records recents without showing the retired initiating target', async (t) => {
  const owner = deferred(),
    ui = await mount(t, { accept: () => owner.promise });
  await input(field('Values 1'), 'A1:A2');
  await click(button('Insert formula into C33'));
  ui.changeCurrent(false);
  await ui.render({});
  await act(async () => owner.resolve());
  assert.deepEqual(recentIds(), ['SUM']);
  assert.equal(ui.outcomes[0].current, false);
  assert(!document.body.textContent.includes('Inserted into Sheet1!C33'));
});
test('common criteria controls accept text or reference criteria and preserve optional sum-range omission', async (t) => {
  await mount(t);
  await select(field('Function'), 'SUMIF');
  await input(field('Criteria range'), 'A1:A3');
  await select(field('Criteria type'), 'text');
  await input(field('Criteria'), '>=2.5');
  assert.equal(document.querySelector('output[aria-label="Formula preview"]').textContent, '=SUMIF(A1:A3,">=2.5")');
  await click(button('Add argument'));
  await input(field('Sum range (optional)'), 'B1:B3');
  assert.equal(
    document.querySelector('output[aria-label="Formula preview"]').textContent,
    '=SUMIF(A1:A3,">=2.5",B1:B3)'
  );
  await input(field('Sum range (optional)'), 'B1:B2');
  assert.equal(button('Insert formula into C33').disabled, true);
  await select(field('Function'), 'COUNTIF');
  await input(field('Criteria range'), 'A1:A3');
  await input(field('Criteria'), 'D1');
  assert.equal(document.querySelector('output[aria-label="Formula preview"]').textContent, '=COUNTIF(A1:A3,D1)');
});
test('guided VLOOKUP shows its exact profile and accepts one closed formula with FALSE', async (t) => {
  const ui = await mount(t);
  await select(field('Function'), 'VLOOKUP');
  await select(field('Lookup value type'), 'text');
  await input(field('Lookup value'), 'sku');
  await input(field('Table range'), 'A1:B3');
  await select(field('Return column type'), 'number');
  await input(field('Return column'), '2');
  assert.equal(
    document.querySelector('output[aria-label="Formula preview"]').textContent,
    '=VLOOKUP("sku",A1:B3,2,FALSE)'
  );
  assert.match(document.body.textContent, /Exact match/);
  assert.equal(button('Add argument').disabled, true);
  await click(button('Insert formula into C33'));
  assert.deepEqual(ui.writes, [
    [{ type: 'setCells', sheet: 'Sheet1', cells: [{ cell: 'C33', formula: 'VLOOKUP("sku",A1:B3,2,FALSE)' }] }],
  ]);
});
test('guided XLOOKUP preserves an explicitly chosen empty-text fallback and refuses mismatched vector shapes', async (t) => {
  const ui = await mount(t);
  await select(field('Function'), 'XLOOKUP');
  await select(field('Lookup value type'), 'text');
  await input(field('Lookup value'), 'key');
  await input(field('Lookup range'), 'A1:A3');
  await input(field('Return range'), 'B1:D1');
  assert.equal(button('Insert formula into C33').disabled, true);
  await input(field('Return range'), 'B1:B3');
  await click(button('Add argument'));
  await select(field('If not found (optional) type'), 'text');
  assert.equal(
    document.querySelector('output[aria-label="Formula preview"]').textContent,
    '=XLOOKUP("key",A1:A3,B1:B3,"")'
  );
  await click(button('Insert formula into C33'));
  assert.equal(ui.writes[0][0].cells[0].formula, 'XLOOKUP("key",A1:A3,B1:B3,"")');
});
