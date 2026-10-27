import './support/scratch-dir.mjs';
/** Actual planner/pane/Grid in React, with simulated trusted reads/acceptance; native proof is separate. */
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
  output = await fs.mkdtemp(path.join(component, '.build/tests/create-names-ui-'));
const entry = path.join(output, 'entry.ts'),
  bundle = path.join(output, 'test.mjs');
await fs.writeFile(
  entry,
  `export * from ${JSON.stringify(path.join(component, 'ui/createNamesFromSelection.ts'))};export {WorkbookCreateNames} from ${JSON.stringify(path.join(component, 'ui/WorkbookCreateNames.tsx'))};export {SpreadsheetGrid} from ${JSON.stringify(path.join(component, 'ui/SpreadsheetGrid.tsx'))};export {configureXlsxHost} from ${JSON.stringify(path.join(component, 'ui/host.tsx'))};`
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
globalThis.requestAnimationFrame = (fn) => setTimeout(fn, 0);
globalThis.cancelAnimationFrame = clearTimeout;
const { createRoot } = await import('react-dom/client'),
  ui = await import(pathToFileURL(bundle));
ui.configureXlsxHost({
  api: { copyText: async () => {}, readClipboardText: async () => '' },
  Icon: () => React.createElement('svg'),
  useI18n: () => ({ locale: 'en' }),
});
const check = (name, fn) => {
  test(name, { timeout: 15000 }, async (t) => {
    await fn(t);
  });
};
after(async () => {
  dom.window.close();
});
const edges = { top: true, left: false, bottom: false, right: false },
  text = (value) => ({ value, dataType: 'text' }),
  binding = (selection) => ({
    documentId: 'document',
    ownerIdentity: 'owner',
    state: 'head',
    sheet: "O'Brien Data",
    selection: selection ?? 'B1:B4',
    navigationGeneration: 0,
    activationEpoch: 0,
    modeEpoch: 0,
  }),
  access = () => ({ active: true, readOnly: false, busy: false, literal: false, specialSelection: false });
const data = { B1: text('Revenue'), B2: { value: 2 }, B3: { value: 4 }, B4: { value: 6 } };
const defer = () => {
  let resolve, reject;
  const promise = new Promise((a, b) => {
    resolve = a;
    reject = b;
  });
  return { promise, resolve, reject };
};
function fixture(patch = {}) {
  const cells = structuredClone(data),
    reads = [],
    accepts = [];
  let current = true,
    runtime = true,
    inventory = { namedRanges: [], tables: [] };
  const context = {
    binding: binding(),
    isCurrent: () => current,
    isRuntimeCurrent: () => runtime,
    availability: access,
    merges: [],
    read: async (range) => {
      reads.push(range);
      const area = rangeBounds(range);
      return Object.fromEntries(
        Object.entries(cells).filter(([r]) => {
          const p = rangeBounds(r).first;
          return p.row >= area.first.row && p.row <= area.last.row && p.col >= area.first.col && p.col <= area.last.col;
        })
      );
    },
    inventory: async () => inventory,
    accept: async (commands) => accepts.push(commands),
    ...patch,
  };
  return {
    context,
    cells,
    reads,
    accepts,
    retire: () => (current = false),
    replace: () => {
      current = false;
      runtime = false;
    },
    inventory: (value) => (inventory = value),
  };
}
const props = (node) => node[Object.keys(node).find((key) => key.startsWith('__reactProps'))],
  field = (name) => document.querySelector('[aria-label="' + name + '"]'),
  button = (name) =>
    [...document.querySelectorAll('button')].find(
      (node) => !node.closest('[hidden]') && (node.getAttribute('aria-label') ?? node.textContent) === name
    );
const flush = async () => act(async () => new Promise((resolve) => setTimeout(resolve, 10)));
async function click(node) {
  assert(node);
  assert.equal(node.disabled, false);
  await act(async () => node.click());
  await flush();
}
async function change(node, value) {
  assert(node);
  await act(async () => props(node).onChange({ target: { value, checked: value } }));
  await flush();
}
async function mount(t, context, extra = {}) {
  const root = createRoot(document.getElementById('root'));
  let args = { context, visible: true, ...extra };
  await act(async () => root.render(React.createElement(ui.WorkbookCreateNames, args)));
  t.after(async () => act(async () => root.unmount()));
  return {
    render: async (patch) => {
      args = { ...args, ...patch };
      await act(async () => root.render(React.createElement(ui.WorkbookCreateNames, args)));
    },
  };
}
check(
  'all four edges and fifteen combinations name the shared data interior and explicitly exclude intersecting corners',
  () => {
    const cells = {};
    for (let r = 1; r <= 5; r++)
      for (let c = 0; c < 5; c++) cells[String.fromCharCode(65 + c) + r] = text('Label_' + r + '_' + c);
    for (let bits = 1; bits < 16; bits++) {
      const flags = Object.fromEntries(['top', 'left', 'bottom', 'right'].map((k, i) => [k, !!(bits & (1 << i))])),
        preview = ui.planNamesFromSelection(binding('A1:E5'), flags, cells, { namedRanges: [], tables: [] });
      const firstRow = 1 + Number(flags.top),
        lastRow = 5 - Number(flags.bottom),
        firstCol = Number(flags.left),
        lastCol = 4 - Number(flags.right);
      assert.equal(
        preview.dataRange,
        String.fromCharCode(65 + firstCol) + firstRow + ':' + String.fromCharCode(65 + lastCol) + lastRow
      );
      assert.equal(
        preview.proposals.length,
        (lastCol - firstCol + 1) * (Number(flags.top) + Number(flags.bottom)) +
          (lastRow - firstRow + 1) * (Number(flags.left) + Number(flags.right))
      );
      for (const proposal of preview.proposals) {
        assert(!preview.corners.includes(proposal.labelCell));
        assert.match(proposal.formula, /^'O''Brien Data'!\$[A-E]\$[1-5]:\$[A-E]\$[1-5]$/);
      }
    }
    const all = ui.planNamesFromSelection(
      binding('A1:E5'),
      { top: true, left: true, bottom: true, right: true },
      cells,
      { namedRanges: [], tables: [] }
    );
    assert.equal(all.proposals.length, 12);
    assert.deepEqual(all.corners, ['A1', 'E1', 'A5', 'E5']);
    assert.equal(all.proposals.find((p) => p.labelCell === 'B1').range, 'B2:B4');
    assert.equal(all.proposals.find((p) => p.labelCell === 'A2').range, 'B2:D2');
  }
);
check(
  'conversion is explicit; reserved/duplicate/global/local/table/invalid-scope conflicts require editing or exclusion',
  async () => {
    const h = fixture();
    h.cells.B1 = text(' Net sales ');
    let p = await ui.previewCreateNames(h.context, edges);
    assert.equal(p.proposals[0].name, 'Net_sales');
    assert.equal(p.proposals[0].converted, true);
    assert.equal(ui.createNameChoiceReason(p, p.proposals[0], {}), null);
    for (const name of ['TRUE', 'false', 'A1', 'R1C1', 'C', '_xlnm.Print_Area', '9name', 'Ω']) {
      h.cells.B1 = text(name);
      p = await ui.previewCreateNames(h.context, edges);
      assert(ui.createNameChoiceReason(p, p.proposals[0], {}));
      assert.equal(
        ui.createNameChoiceReason(p, p.proposals[0], { [p.proposals[0].id]: { name: 'Safe_Name', included: true } }),
        null
      );
    }
    h.cells.B1 = text('Revenue');
    for (const inventory of [
      { namedRanges: [{ name: 'revenue', formula: 'A1' }], tables: [] },
      { namedRanges: [{ name: 'Revenue', formula: 'A1', sheet: "O'Brien Data" }], tables: [] },
      { namedRanges: [{ name: 'Revenue', formula: 'A1', scopeInvalid: true }], tables: [] },
      { namedRanges: [], tables: [{ name: 'REVENUE' }] },
    ]) {
      h.inventory(inventory);
      p = await ui.previewCreateNames(h.context, edges);
      assert(ui.createNameChoiceReason(p, p.proposals[0], {}));
    }
    h.inventory({ namedRanges: [{ name: 'Revenue', formula: 'A1', sheet: 'Other' }], tables: [] });
    p = await ui.previewCreateNames(h.context, edges);
    assert.equal(ui.createNameChoiceReason(p, p.proposals[0], {}), null);
    h.context.binding = binding('B1:C4');
    h.cells.C1 = text('revenue');
    h.inventory({ namedRanges: [], tables: [] });
    p = await ui.previewCreateNames(h.context, edges);
    assert(ui.createNameChoiceReason(p, p.proposals[0], {}));
    const choices = { [p.proposals[1].id]: { name: 'revenue', included: false } };
    await ui.acceptCreateNames(h.context, p, choices);
    assert.equal(h.accepts.length, 1);
    assert.equal(h.accepts[0].length, 1);
  }
);
check(
  'only accepted text labels are eligible; pending/spill/merged/numeric labels cannot become names through coercion',
  async () => {
    for (const cell of [
      { value: 12, dataType: 'number', rawValue: '12' },
      { value: 'Revenue', spill: { role: 'follower' } },
      { value: 'Revenue', calculation: { status: 'stored', pending: true } },
      { value: '#N/A', dataType: 'error', error: '#N/A' },
      { formula: 'UNKNOWN()', value: null, calculation: { status: 'unresolved' } },
    ]) {
      const h = fixture();
      h.cells.B1 = cell;
      const p = await ui.previewCreateNames(h.context, edges);
      assert(p.proposals[0].sourceReason);
      await assert.rejects(ui.acceptCreateNames(h.context, p, {}));
      assert.equal(h.accepts.length, 0);
    }
    const merged = fixture({ merges: ['B1:C1'] });
    const p = await ui.previewCreateNames(merged.context, edges);
    assert.match(p.proposals[0].sourceReason, /unmerged/);
    const valid = fixture();
    valid.cells.B1 = {
      formula: 'IF(TRUE,"Revenue","")',
      value: null,
      calculation: { status: 'calculated', value: 'Revenue' },
    };
    assert.equal((await ui.previewCreateNames(valid.context, edges)).proposals[0].name, 'Revenue');
  }
);
check(
  'off-viewport labels use bounded real callback rectangles, never loaded Grid cells; empty/oversized selection refuses before reads',
  async () => {
    const h = fixture();
    h.context.binding = binding('H1001:H1004');
    h.cells.H1001 = text('Remote_Total');
    const p = await ui.previewCreateNames(h.context, edges);
    assert.equal(p.proposals[0].formula, "'O''Brien Data'!$H$1002:$H$1004");
    assert.deepEqual(h.reads, ['H1001:H1001']);
    const long = fixture();
    long.context.binding = binding('A1:B1000');
    for (let row = 1; row <= 1000; row++) long.cells['A' + row] = text('Row_' + row);
    await ui.previewCreateNames(long.context, { top: true, left: true, bottom: false, right: false });
    assert(long.reads.every((range) => rangeBounds(range).size <= 1000));
    for (const selection of ['A1', 'A1:XFD1048576', 'A1:B3000']) {
      const bad = fixture();
      bad.context.binding = binding(selection);
      await assert.rejects(
        ui.previewCreateNames(
          bad.context,
          selection === 'A1:B3000' ? { top: false, left: true, bottom: false, right: false } : edges
        )
      );
      assert.equal(bad.reads.length, 0);
    }
  }
);
check(
  'labels and full accepted inventory are re-read exactly before one immutable strict batch and acknowledged outcome',
  async () => {
    const h = fixture(),
      p = await ui.previewCreateNames(h.context, edges);
    h.cells.B1.rawValue = 'changed metadata';
    await assert.rejects(ui.acceptCreateNames(h.context, p, {}), /changed/);
    assert.equal(h.accepts.length, 0);
    delete h.cells.B1.rawValue;
    h.inventory({ namedRanges: [{ name: 'Other', formula: 'A1' }], tables: [] });
    await assert.rejects(ui.acceptCreateNames(h.context, p, {}), /changed/);
    assert.equal(h.accepts.length, 0);
    h.inventory({ namedRanges: [], tables: [] });
    const held = defer();
    h.context.accept = async (commands) => {
      h.accepts.push(commands);
      await held.promise;
    };
    let finished = false;
    const accepting = ui.acceptCreateNames(h.context, p, {}).then((value) => {
      finished = true;
      return value;
    });
    await flush();
    assert.equal(h.accepts.length, 1);
    assert.equal(finished, false);
    assert.deepEqual(h.accepts[0], [
      {
        operation: 'defineName',
        sheet: "O'Brien Data",
        name: 'Revenue',
        range: 'B2:B4',
        nameScope: 'workbook',
        expectedNameFormula: null,
      },
    ]);
    held.resolve();
    assert.equal((await accepting).outcome, 'accepted');
    assert.equal(finished, true);
  }
);
check(
  'late labels/inventory and cancellation retire before any acceptance; malformed scoped cell receipts refuse',
  async () => {
    for (const where of ['read', 'inventory']) {
      const gate = defer(),
        h = fixture({ [where]: () => gate.promise });
      const pending = ui.previewCreateNames(h.context, edges);
      await flush();
      h.retire();
      gate.resolve(where === 'read' ? { B1: text('Revenue') } : { namedRanges: [], tables: [] });
      await assert.rejects(pending, ui.CreateNamesRetired);
      assert.equal(h.accepts.length, 0);
    }
    const h = fixture(),
      p = await ui.previewCreateNames(h.context, edges),
      abort = new AbortController();
    await assert.rejects(
      ui.acceptCreateNames(h.context, p, {}, { signal: abort.signal, beforeAcceptance: () => abort.abort() }),
      ui.CreateNamesCancelled
    );
    assert.equal(h.accepts.length, 0);
    for (const reply of [null, [], { B1: null }, { B1: [] }, { B1: 'Revenue' }, { Z99: text('Revenue') }]) {
      const bad = fixture({ read: async () => reply });
      await assert.rejects(ui.previewCreateNames(bad.context, edges));
      assert.equal(bad.accepts.length, 0);
    }
  }
);
check(
  'readonly/inactive/literal/special/busy contexts cannot preview or accept; label/inventory byte budgets are real',
  async () => {
    for (const patch of [
      { readOnly: true },
      { active: false },
      { literal: true },
      { specialSelection: true },
      { busy: true },
    ]) {
      const h = fixture({ availability: () => ({ ...access(), ...patch }) });
      await assert.rejects(ui.previewCreateNames(h.context, edges));
      assert.equal(h.accepts.length, 0);
    }
    const huge = fixture();
    huge.context.binding = binding('A1:AO2');
    for (let col = 0; col < 41; col++)
      huge.cells[(col < 26 ? String.fromCharCode(65 + col) : 'A' + String.fromCharCode(65 + col - 26)) + '1'] = text(
        'x'.repeat(32767)
      );
    huge.context.read = async () => Object.fromEntries(Object.entries(huge.cells).filter(([ref]) => ref.endsWith('1')));
    await assert.rejects(ui.previewCreateNames(huge.context, edges), /supported size/);
    assert.equal(huge.accepts.length, 0);
  }
);
check(
  'pane offers editable/excluded exact proposals and one acknowledgement; accepted head advance stays truthful without resubmit',
  async (t) => {
    const h = fixture();
    h.cells.B1 = text('TRUE');
    await mount(t, h.context);
    await click(button('Preview names'));
    assert(field('Proposed name B1'));
    assert.equal(button('Create names').disabled, true);
    await change(field('Proposed name B1'), 'Enabled_Value');
    assert.equal(field('Named range B1').textContent, "'O''Brien Data'!$B$2:$B$4");
    const gate = defer();
    h.context.accept = async (commands) => {
      h.accepts.push(commands);
      await gate.promise;
      h.retire();
    };
    await click(button('Create names'));
    assert.equal(h.accepts.length, 1);
    assert.match(document.body.textContent, /Finishing name creation/);
    assert.equal(button('Finishing…').disabled, true);
    await act(async () => gate.resolve());
    await flush();
    assert.match(document.body.textContent, /Created 1 workbook names/);
    assert.equal(button('Create names'), undefined);
    assert.equal(h.accepts.length, 1);
  }
);
check(
  'Cancel and hidden replacement join ignored reads, never adopt late rows, and permit retry only after physical settlement',
  async (t) => {
    const gate = defer(),
      h = fixture({ read: () => gate.promise }),
      mounted = await mount(t, h.context);
    await click(button('Preview names'));
    await click(button('Cancel name creation'));
    assert.equal(button('Preview names').disabled, true);
    assert.match(document.body.textContent, /Stopping/);
    await mounted.render({ visible: false });
    const other = fixture();
    other.context.binding = { ...binding(), documentId: 'other', ownerIdentity: 'new' };
    await mounted.render({ visible: true, context: other.context });
    assert.equal(field('Names preview'), null);
    assert.equal(button('Preview names').disabled, true);
    assert.equal(other.reads.length, 0);
    await act(async () => gate.resolve({ B1: text('old') }));
    await flush();
    assert.equal(field('Names preview'), null);
    assert.equal(button('Preview names').disabled, false);
    await click(button('Preview names'));
    assert(field('Names preview'));
    assert.equal(other.accepts.length, 0);
  }
);
check(
  'head retirement masks rows/errors immediately; rejected old acceptance never reports success or error on a new owner',
  async (t) => {
    const h = fixture(),
      outcomes = [],
      mounted = await mount(t, h.context, { onOutcome: (value) => outcomes.push(value) });
    await click(button('Preview names'));
    const gate = defer();
    h.context.accept = () => gate.promise;
    await click(button('Create names'));
    h.replace();
    await mounted.render({ context: h.context });
    assert.equal(field('Names preview'), null);
    await act(async () => gate.reject(Error('Old owner refused')));
    await flush();
    assert.equal(document.querySelector('[role="alert"]'), null);
    assert.equal(outcomes.length, 0);
    assert.match(document.body.textContent, /changed/);
  }
);
check('durable late acceptance remains an original outcome but never paints names on a retired owner', async (t) => {
  const h = fixture(),
    outcomes = [],
    mounted = await mount(t, h.context, { onOutcome: (value) => outcomes.push(value) });
  await click(button('Preview names'));
  const gate = defer();
  h.context.accept = () => gate.promise;
  await click(button('Create names'));
  h.replace();
  await mounted.render({ visible: false, context: h.context });
  await act(async () => gate.resolve());
  await flush();
  assert.equal(outcomes.length, 1);
  assert.equal(outcomes[0].outcome, 'accepted');
  assert.equal(outcomes[0].runtimeCurrent, false);
  assert.doesNotMatch(document.body.textContent, /Created 1 workbook/);
});
async function grid(t, patch = {}) {
  const root = createRoot(document.getElementById('root')),
    writes = [],
    accepts = [],
    reads = [],
    cells = structuredClone(data),
    sheet = {
      name: "O'Brien Data",
      path: 'xl/worksheets/sheet1.xml',
      state: 'visible',
      rows: 1004,
      columns: 14,
      cells: {},
      merges: [],
      widths: [],
      hiddenRows: [],
      rowHeights: [],
      freeze: { rows: 0, columns: 0 },
    };
  let args = {
    sheets: [sheet, { ...sheet, name: 'Other', path: 'xl/worksheets/sheet2.xml' }],
    active: true,
    documentId: 'document',
    ownerIdentity: 'owner',
    storageKey: 'owner:head',
    nameCommands: true,
    namedRanges: [],
    tables: [],
    onEdit: (edits) => writes.push(edits),
    onBatch: (commands) => writes.push(commands),
    onCommandBatch: async (commands) => accepts.push(commands),
    readRange: async (name, range) => {
      reads.push({ name, range });
      const b = rangeBounds(range);
      return Object.fromEntries(
        Object.entries(cells).filter(([r]) => {
          const p = rangeBounds(r).first;
          return p.row >= b.first.row && p.row <= b.last.row && p.col >= b.first.col && p.col <= b.last.col;
        })
      );
    },
    ...patch,
  };
  const render = async (patch) => {
    args = { ...args, ...patch };
    await act(async () => root.render(React.createElement(ui.SpreadsheetGrid, args)));
    await flush();
  };
  await render();
  t.after(async () => act(async () => root.unmount()));
  await click([...document.querySelectorAll('[role="tab"]')].find((tab) => tab.textContent === 'Formulas'));
  return { writes, accepts, reads, cells, render };
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
check(
  'actual Grid opens one dock, preserves parked draft/caret through preview/cancel/close and never commits another input',
  async (t) => {
    const h = await grid(t);
    await go('B1:B4');
    const formula = field('Formula or value');
    await act(async () => {
      formula.focus();
      props(formula).onChange({ target: { value: '=SUM(' } });
    });
    formula.setSelectionRange(5, 5);
    const command = document.querySelector('[data-ribbon-command="names.create-from-selection"]'),
      down = new window.MouseEvent('pointerdown', { bubbles: true, cancelable: true });
    await act(async () => command.dispatchEvent(down));
    assert.equal(down.defaultPrevented, true);
    await click(command);
    assert.equal(document.querySelectorAll('aside.sheet-task-pane:not([hidden])').length, 1);
    await click(button('Preview names'));
    assert.equal(field('Source label B1').textContent, '"Revenue"');
    assert.equal(formula.value, '=SUM(');
    assert.deepEqual(h.writes, []);
    assert.deepEqual(h.accepts, []);
    await click(button('Cancel name creation'));
    await click(button('Close workbook tools'));
    await flush();
    assert.equal(formula.value, '=SUM(');
    assert.equal(formula.selectionStart, 5);
    assert.equal(document.activeElement, formula);
    assert.deepEqual(h.writes, []);
    assert.deepEqual(h.accepts, []);
  }
);
check(
  'actual Grid strict batch uses original offscreen selection and owner callback once, not legacy sends',
  async (t) => {
    const h = await grid(t);
    h.cells.H1001 = text('Remote_Total');
    await go('H1001:H1004');
    await click(button('Create from Selection'));
    await click(button('Preview names'));
    assert.equal(field('Named range H1001').textContent, "'O''Brien Data'!$H$1002:$H$1004");
    await click(button('Create names'));
    assert.equal(h.accepts.length, 1);
    assert.deepEqual(h.accepts[0], [
      {
        operation: 'defineName',
        sheet: "O'Brien Data",
        name: 'Remote_Total',
        range: 'H1002:H1004',
        nameScope: 'workbook',
        expectedNameFormula: null,
      },
    ]);
    assert.deepEqual(h.writes, []);
  }
);
check(
  'actual Grid same selection/head cannot revive a preview after navigation away and back or a runtime replacement',
  async (t) => {
    const h = await grid(t);
    await go('B1:B4');
    await click(button('Create from Selection'));
    await click(button('Preview names'));
    assert(field('Names preview'));
    await go('D2');
    await go('B1:B4');
    assert.equal(field('Names preview'), null);
    const retry = button('Preview names');
    if (retry) assert.equal(retry.disabled, true);
    await h.render({ ownerIdentity: 'replacement' });
    assert.equal(field('Names preview'), null);
    assert.deepEqual(h.accepts, []);
    assert.deepEqual(h.writes, []);
  }
);
check(
  'actual Grid source capabilities disable readonly/busy and omit literal/passive or missing strict-owner creation',
  async (t) => {
    const h = await grid(t);
    const selector = '[data-ribbon-command="names.create-from-selection"]';
    await h.render({ readOnly: true });
    assert(document.querySelector(selector).disabled);
    await h.render({ readOnly: false, busy: true });
    assert(document.querySelector(selector).disabled);
    for (const patch of [
      { busy: false, literal: true },
      { literal: false, passiveNotice: 'Stored passive view' },
      { passiveNotice: undefined, ownerIdentity: undefined },
      { ownerIdentity: 'owner', onCommandBatch: undefined },
    ]) {
      await h.render(patch);
      assert.equal(document.querySelector(selector), null);
    }
    assert.deepEqual(h.writes, []);
    assert.deepEqual(h.accepts, []);
  }
);

check(
  'actual Grid accepted name head retires the parked formula draft without committing or rebasing it',
  async (t) => {
    let accepted;
    const h = await grid(t, {
      onCommandBatch: (commands) =>
        new Promise((resolve) => {
          accepted = { commands, resolve };
        }),
    });
    await go('B1:B4');
    const formula = field('Formula or value');
    await act(async () => {
      formula.focus();
      props(formula).onChange({ target: { value: '=SUM(' } });
    });
    formula.setSelectionRange(5, 5);
    await click(button('Create from Selection'));
    await click(button('Preview names'));
    await click(button('Create names'));
    assert(accepted);
    assert.equal(accepted.commands[0].name, 'Revenue');
    assert.equal(formula.value, '=SUM(');
    await h.render({
      storageKey: 'owner:accepted',
      namedRanges: [{ name: 'Revenue', formula: "'O''Brien Data'!$B$2:$B$4" }],
    });
    assert.equal(formula.value, 'Revenue');
    await act(async () => accepted.resolve());
    await flush();
    assert.match(document.body.textContent, /Created 1 workbook names/);
    await click(button('Close workbook tools'));
    assert.equal(formula.value, 'Revenue');
    await act(async () =>
      formula.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
    );
    await flush();
    assert.deepEqual(h.writes, []);
    assert.equal(accepted.commands.length, 1);
  }
);
check(
  'actual Grid refuses a live worksheet-cell draft instead of committing it when opening name creation',
  async (t) => {
    const h = await grid(t);
    await go('B1:B4');
    const cell = document.querySelector('.sheet-grid input[aria-label="O\'Brien Data B1"]');
    assert(cell);
    await act(async () => {
      cell.focus();
      props(cell).onDoubleClick?.({ currentTarget: cell });
      props(cell).onChange({ target: { value: 'uncommitted' } });
    });
    assert.equal(cell.getAttribute('data-editing'), 'true');
    await click(button('Create from Selection'));
    assert.equal(document.querySelector('.sheet-create-names'), null);
    assert.match(document.body.textContent, /Finish or cancel the current input/);
    assert.deepEqual(h.writes, []);
    assert.deepEqual(h.accepts, []);
  }
);
