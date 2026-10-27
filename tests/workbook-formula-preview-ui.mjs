import './support/scratch-dir.mjs';
/** Actual source preview controller/Grid/Functions; simulated trusted host replies only. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
const component = fileURLToPath(new URL('../', import.meta.url)),
  output = await fs.mkdtemp(path.join(component, '.build/tests/formula-preview-ui-'));
const entry = path.join(output, 'entry.ts'),
  bundle = path.join(output, 'test.mjs');
await fs.writeFile(
  entry,
  `export {SpreadsheetGrid} from ${JSON.stringify(path.join(component, 'ui/SpreadsheetGrid.tsx'))};export {configureXlsxHost} from ${JSON.stringify(path.join(component, 'ui/host.tsx'))};export * from ${JSON.stringify(path.join(component, 'ui/formulaDraftPreview.ts'))};export {FormulaDraftPreview} from ${JSON.stringify(path.join(component, 'ui/WorkbookFormulaPreview.tsx'))};`
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
  test(name, { timeout: 15000 }, async () => {
    await fn();
  });
};
after(async () => {
  dom.window.close();
});
const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms)),
  flush = async (ms = 10) => act(async () => wait(ms));
const defer = () => {
  let resolve, reject;
  const promise = new Promise((a, b) => {
    resolve = a;
    reject = b;
  });
  return { promise, resolve, reject };
};
const receipt = { documentId: 'document', ownerId: 'indexed-owner', state: 'public-head', revision: 'revision' };
const scope = () => ({
  documentId: 'document',
  ownerIdentity: 'runtime',
  state: 'opaque-grid-token',
  sheet: 'Sheet1',
  cell: 'A1',
  navigationGeneration: 0,
  activationEpoch: 0,
  modeEpoch: 0,
  active: true,
  normal: true,
  busy: false,
  profile: 'xlsx',
  receipt: { ...receipt },
});
const input = (frame, formula = 'SUM(B1:B3)', generation = 1) => ({
  scope: frame.scope,
  formula,
  generation,
  active: true,
});
const reply = (call, result = { status: 'value', type: 'number', value: 12 }) => ({
  schema: 1,
  requestId: call.requestId,
  receipt: { ...call.receipt },
  sheet: call.sheet,
  cell: call.cell,
  formula: call.formula,
  authority: 'draft-preview',
  result,
});
function model() {
  const calls = [],
    cancellations = [];
  let frame = {
    scope: scope(),
    controls: {
      receipt: { ...receipt },
      preview: async (sheet, cell, formula, requestId) => {
        const held = defer(),
          call = { sheet, cell, formula, requestId, receipt: { ...frame.controls.receipt }, held };
        calls.push(call);
        return held.promise;
      },
      cancel: async (id) => {
        cancellations.push(id);
      },
    },
  };
  const controller = ui.createFormulaDraftPreviewController(() => frame, { debounceMs: 0 });
  return {
    controller,
    calls,
    cancellations,
    frame: () => frame,
    set: (patch) => {
      frame = { ...frame, ...patch };
    },
    stop: () => {
      controller.dispose();
      for (const call of calls) call.held.resolve(reply(call));
    },
  };
}
check(
  'typed scalars preserve text that resembles an error, Boolean, signed zero and true errors; malformed receipts refuse',
  () => {
    const frame = { scope: scope() },
      candidate = input(frame),
      call = { ...receipt, receipt, requestId: 'id', sheet: 'Sheet1', cell: 'A1', formula: candidate.formula };
    for (const result of [
      { status: 'value', type: 'number', value: -0 },
      { status: 'value', type: 'text', value: '#N/A' },
      { status: 'value', type: 'boolean', value: false },
      { status: 'value', type: 'error', value: '#DIV/0!' },
      { status: 'unavailable', reason: 'Protected input' },
    ])
      assert.deepEqual(ui.validateFormulaPreviewReply(reply(call, result), candidate, 'id').result, result);
    for (const change of [
      { receipt: { ...receipt, ownerId: 'foreign' } },
      { formula: 'SUM(C1:C3)' },
      { cell: 'A2' },
      { authority: 'accepted-cache' },
      { result: { status: 'value', type: 'number', value: Infinity } },
      { result: { status: 'value', type: 'boolean', value: 'false' } },
      { result: { status: 'value', type: 'error', value: 'text #N/A' } },
      { result: { status: 'unavailable', reason: 'bad', value: 12 } },
    ])
      assert.throws(() => ui.validateFormulaPreviewReply({ ...reply(call), ...change }, candidate, 'id'));
  }
);
check(
  'one physical preview retires an old draft and runs only the latest replacement after its real promise settles',
  async () => {
    const h = model();
    try {
      h.controller.setInput('bar', input(h.frame()));
      await wait(10);
      assert.equal(h.calls.length, 1);
      const old = h.calls[0];
      h.controller.setInput('bar', input(h.frame(), 'SUM(B1:B3)+1', 2));
      h.controller.setInput('bar', input(h.frame(), 'SUM(B1:B3)+2', 3));
      await wait(10);
      assert.equal(h.calls.length, 1);
      assert.deepEqual(h.cancellations, [old.requestId]);
      old.held.resolve(reply(old));
      await wait(10);
      assert.equal(h.calls.length, 2);
      assert.equal(h.calls[1].formula, 'SUM(B1:B3)+2');
      assert.equal(h.controller.getSnapshot().bar.reply, null);
      h.calls[1].held.resolve(reply(h.calls[1], { status: 'value', type: 'number', value: 14 }));
      await wait(10);
      assert.equal(h.controller.getSnapshot().bar.reply.result.value, 14);
    } finally {
      h.stop();
    }
  }
);
check(
  'Cancel suppresses an ignored-abort reply, does not duplicate physical work, and new input can retry after settlement',
  async () => {
    const h = model();
    try {
      const original = input(h.frame());
      h.controller.setInput('bar', original);
      await wait(10);
      const first = h.calls[0];
      h.controller.cancel('bar', ui.formulaPreviewInputKey(original));
      h.controller.cancel('bar', ui.formulaPreviewInputKey(original));
      assert.equal(h.controller.getSnapshot().bar.phase, 'cancelled');
      assert.equal(h.cancellations.length, 1);
      h.controller.setInput('bar', input(h.frame(), 'SUM(B1:B3)+1', 2));
      await wait(10);
      assert.equal(h.calls.length, 1);
      first.held.resolve(reply(first));
      await wait(10);
      assert.equal(h.calls.length, 2);
      assert.equal(h.controller.getSnapshot().bar.reply, null);
    } finally {
      h.stop();
    }
  }
);
check(
  'public receipt changes and runtime/head/navigation/mode/activation retirement never adopt a held result',
  async () => {
    for (const patch of [
      { ownerIdentity: 'replacement' },
      { state: 'new-opaque-token' },
      { sheet: 'Other' },
      { cell: 'B1' },
      { navigationGeneration: 1 },
      { activationEpoch: 2 },
      { modeEpoch: 1, normal: false },
      { active: false },
      { receipt: { ...receipt, revision: 'changed' } },
      { receipt: { ...receipt, ownerId: 'replacement-index' } },
    ]) {
      const h = model();
      try {
        h.controller.setInput('bar', input(h.frame()));
        await wait(5);
        const call = h.calls[0];
        h.set({ scope: { ...h.frame().scope, ...patch } });
        call.held.resolve(reply(call));
        await wait(5);
        assert.equal(h.controller.getSnapshot().bar.reply, null);
      } finally {
        h.stop();
      }
    }
  }
);
check(
  'generated Functions draft has priority but shares the same physical gate; invalid/IME/private profiles never call host',
  async () => {
    const h = model();
    try {
      h.controller.setInput('bar', input(h.frame()));
      await wait(5);
      const old = h.calls[0];
      h.controller.setInput('functions', input(h.frame(), 'LEN("#N/A")'));
      await wait(5);
      assert.equal(h.calls.length, 1);
      old.held.resolve(reply(old));
      await wait(5);
      assert.equal(h.calls.length, 2);
      assert.equal(h.calls[1].formula, 'LEN("#N/A")');
      h.calls[1].held.resolve(reply(h.calls[1], { status: 'value', type: 'number', value: 4 }));
      await wait(5);
      assert.equal(h.controller.getSnapshot().functions.reply.result.value, 4);
      assert.equal(h.controller.getSnapshot().bar.reply, null);
    } finally {
      h.stop();
    }
    for (const patch of [
      { formula: '' },
      { formula: 'a'.repeat(8193) },
      { composing: true },
      { scope: { ...scope(), profile: 'delimited' } },
      { scope: { ...scope(), profile: 'passive' } },
      { scope: { ...scope(), busy: true } },
    ]) {
      const h = model();
      try {
        h.controller.setInput('bar', { ...input(h.frame()), ...patch });
        await wait(5);
        assert.equal(h.calls.length, 0);
      } finally {
        h.stop();
      }
    }
  }
);
check(
  'shared Grid route identity and trusted durable receipt remain independent without weakening either retirement fence',
  async () => {
    const h = model();
    try {
      h.set({ scope: { ...h.frame().scope, documentId: 'project:entry' } });
      h.controller.setInput('bar', input(h.frame()));
      await wait(5);
      assert.equal(h.calls.length, 1);
      const call = h.calls[0];
      assert.equal(call.receipt.documentId, 'document');
      call.held.resolve(reply(call));
      await wait(5);
      assert.equal(h.controller.getSnapshot().bar.reply.result.value, 12);
      h.controller.setInput('bar', input(h.frame(), 'SUM(B1:B3)+1', 2));
      await wait(5);
      const late = h.calls[1];
      h.set({ scope: { ...h.frame().scope, documentId: 'another-project:entry' } });
      late.held.resolve(reply(late));
      await wait(5);
      assert.equal(h.controller.getSnapshot().bar.reply, null);
    } finally {
      h.stop();
    }
  }
);
const source = (name) => ({
  name,
  path: 'xl/worksheets/' + name + '.xml',
  state: 'visible',
  rows: 4,
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
  outputNode = () => document.querySelector('[aria-label="Formula result preview"]');
async function mount(patch = {}, strictMode = false) {
  const root = createRoot(document.getElementById('root')),
    calls = [],
    cancels = [],
    writes = [],
    strict = [];
  let held = false,
    values = { status: 'value', type: 'number', value: 12 },
    parameters;
  const controls = {
    receipt: { ...receipt },
    preview: async (sheet, cell, formula, requestId) => {
      const call = { sheet, cell, formula, requestId, receipt: { ...controls.receipt }, held: defer() };
      calls.push(call);
      return held ? call.held.promise : reply(call, values);
    },
    cancel: async (id) => {
      cancels.push(id);
    },
  };
  parameters = {
    sheets: [source('Sheet1'), source('Other')],
    active: true,
    documentId: 'document',
    ownerIdentity: 'runtime',
    storageKey: 'opaque-grid-token',
    namedRanges: [{ name: 'Revenue', formula: "'Sheet1'!$B$1:$B$3" }],
    formulaPreviewControls: controls,
    readRange: async (name, range) => {
      const cells = parameters.sheets.find((sheet) => sheet.name === name).cells,
        [first, last = first] = range.split(':');
      return first === last ? Object.fromEntries(Object.entries(cells).filter(([ref]) => ref === first)) : cells;
    },
    onEdit: (value) => writes.push(value),
    onBatch: (value) => writes.push(value),
    onCommandBatch: async (commands) => strict.push(commands),
    ...patch,
  };
  const render = async (patch) => {
    parameters = { ...parameters, ...patch };
    await act(async () =>
      root.render(
        strictMode
          ? React.createElement(React.StrictMode, null, React.createElement(ui.SpreadsheetGrid, parameters))
          : React.createElement(ui.SpreadsheetGrid, parameters)
      )
    );
  };
  await render();
  await flush();
  return {
    calls,
    cancels,
    writes,
    strict,
    controls,
    render,
    setHeld: (value) => {
      held = value;
    },
    setResult: (result) => {
      values = result;
    },
    close: async () => {
      await act(async () => root.unmount());
      for (const call of calls) call.held.resolve(reply(call));
      await wait(5);
    },
  };
}
async function type(text) {
  await act(async () => {
    field().focus();
    props(field()).onChange({ target: { value: text } });
  });
}
async function key(key) {
  await act(async () =>
    field().dispatchEvent(new window.KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }))
  );
}
check(
  'actual formula bar debounces exact text and displays a typed preview with no edit/history/recents action',
  async () => {
    const view = await mount();
    try {
      await type('=SUM(B1:B3)');
      await flush(240);
      assert.equal(view.calls.length, 1);
      assert.equal(view.calls[0].formula, 'SUM(B1:B3)');
      assert.equal(outputNode().textContent, 'Number: 12');
      assert.deepEqual(view.writes, []);
      assert.deepEqual(view.strict, []);
      assert.equal(window.localStorage.getItem('xlsx.function-recents.v1'), null);
      await key('Escape');
      assert.equal(outputNode()?.textContent ?? null, null);
      assert.deepEqual(view.writes, []);
    } finally {
      await view.close();
    }
  }
);
check(
  'ordinary value drafts stay clean and replacing a formula with literal text immediately retires its held preview without writes',
  async () => {
    const view = await mount();
    view.setHeld(true);
    try {
      await type('12');
      await flush(240);
      assert.equal(view.calls.length, 0);
      assert.equal(outputNode(), null);
      await type('=SUM(B1:B3)');
      await flush(240);
      assert.equal(view.calls.length, 1);
      assert(outputNode());
      const held = view.calls[0];
      await type('00123');
      assert.equal(outputNode(), null);
      assert.deepEqual(view.cancels, [held.requestId]);
      await flush(240);
      assert.equal(view.calls.length, 1);
      await act(async () => held.held.resolve(reply(held)));
      await flush();
      assert.equal(outputNode(), null);
      assert.equal(field().value, '00123');
      assert.deepEqual(view.writes, []);
      assert.deepEqual(view.strict, []);
      await type('=SUM(B1:B3)+1');
      await flush(240);
      assert.equal(view.calls.length, 2);
      const next = view.calls[1];
      await act(async () => next.held.resolve(reply(next, { status: 'value', type: 'number', value: 13 })));
      await flush();
      assert.equal(outputNode().textContent, 'Number: 13');
      await type('');
      assert.equal(outputNode(), null);
      assert.deepEqual(view.writes, []);
      assert.deepEqual(view.strict, []);
    } finally {
      await view.close();
    }
  }
);
check(
  'actual formula bar preserves text/error/Boolean/-0 distinction and incomplete/native-unavailable reasons',
  async () => {
    const view = await mount();
    try {
      let index = 0;
      for (const [result, text] of [
        [{ status: 'value', type: 'text', value: '#N/A' }, 'Text: "#N/A"'],
        [{ status: 'value', type: 'error', value: '#N/A' }, 'Error: #N/A'],
        [{ status: 'value', type: 'boolean', value: false }, 'Boolean: FALSE'],
        [{ status: 'value', type: 'number', value: -0 }, 'Number: -0'],
        [{ status: 'invalid', reason: 'Incomplete formula' }, 'Incomplete formula'],
        [{ status: 'unavailable', reason: 'This input is protected.' }, 'This input is protected.'],
      ]) {
        view.setResult(result);
        await type('=SUM(B1:B3)+' + index++);
        await flush(240);
        assert.equal(outputNode().textContent, text);
      }
      assert.deepEqual(view.writes, []);
      assert.deepEqual(view.strict, []);
    } finally {
      await view.close();
    }
  }
);
check('actual IME and cancelled preview preserve the live draft while late replies remain masked', async () => {
  const view = await mount();
  try {
    await act(async () => {
      field().focus();
      props(field()).onCompositionStart({ currentTarget: field() });
      props(field()).onChange({ target: { value: '=SUM(B1:B3)' } });
    });
    await flush(240);
    assert.equal(view.calls.length, 0);
    await act(async () => props(field()).onCompositionEnd({ currentTarget: field() }));
    view.setHeld(true);
    await flush(240);
    assert.equal(view.calls.length, 1);
    const call = view.calls[0];
    const cancel = [...document.querySelectorAll('button')].find((button) => button.textContent === 'Cancel preview');
    await act(async () => cancel.click());
    assert.match(outputNode().textContent, /cancelled/);
    assert.equal(field().value, '=SUM(B1:B3)');
    await act(async () => call.held.resolve(reply(call)));
    await flush();
    assert.match(outputNode().textContent, /cancelled/);
    assert.deepEqual(view.writes, []);
    assert.deepEqual(view.strict, []);
  } finally {
    await view.close();
  }
});
check(
  'actual same-opaque-head receipt replacement, hidden formula bar and inactive owner mask old preview immediately',
  async () => {
    const view = await mount();
    try {
      await type('=SUM(B1:B3)');
      await flush(240);
      assert.equal(outputNode().textContent, 'Number: 12');
      await act(async () =>
        [...document.querySelectorAll('[role="tab"]')].find((tab) => tab.textContent === 'View').click()
      );
      const toggle = document.querySelector('[data-ribbon-command="view.formula-bar"]');
      await act(async () => props(toggle).onChange({ target: { checked: false } }));
      assert.equal(outputNode()?.textContent ?? null, null);
      await act(async () => props(toggle).onChange({ target: { checked: true } }));
      await flush(240);
      assert.equal(outputNode().textContent, 'Number: 12');
      await view.render({
        formulaPreviewControls: { ...view.controls, receipt: { ...receipt, revision: 'new-revision' } },
      });
      assert.equal(outputNode()?.textContent ?? null, null);
      await type('=SUM(B1:B3)+1');
      view.setHeld(true);
      await flush(240);
      const call = view.calls.at(-1);
      await view.render({ active: false });
      assert.equal(outputNode()?.textContent ?? null, null);
      await act(async () => call.held.resolve(reply(call)));
      await flush();
      assert.equal(outputNode()?.textContent ?? null, null);
      assert.deepEqual(view.writes, []);
      assert.deepEqual(view.strict, []);
    } finally {
      await view.close();
    }
  }
);
check(
  'name chooser parking updates preview at the exact caret without writes; explicit Enter alone invokes original strict callback',
  async () => {
    const view = await mount();
    try {
      await act(async () =>
        [...document.querySelectorAll('[role="tab"]')].find((tab) => tab.textContent === 'Formulas').click()
      );
      await type('=SUM()+10');
      field().setSelectionRange(5, 5);
      const command = document.querySelector('[data-ribbon-command="names.use-in-formula"]');
      await act(async () => command.click());
      await flush();
      const choice = document.querySelector('[aria-label="Insert name Revenue · Workbook"]');
      await act(async () => choice.click());
      await flush(240);
      assert.equal(field().value, '=SUM(Revenue)+10');
      assert.equal(field().selectionStart, 12);
      assert.equal(view.calls.at(-1).formula, 'SUM(Revenue)+10');
      assert.deepEqual(view.strict, []);
      assert.deepEqual(view.writes, []);
      await key('Enter');
      await flush();
      assert.deepEqual(view.strict, [[{ sheet: 'Sheet1', cell: 'A1', formula: 'SUM(Revenue)+10' }]]);
      assert.equal(outputNode()?.textContent ?? null, null);
      await key('Enter');
      await flush();
      assert.equal(view.strict.length, 1);
    } finally {
      await view.close();
    }
  }
);
check(
  'actual readonly Functions inspector previews serialized arguments while insertion stays disabled; hidden pane retires preview',
  async () => {
    const view = await mount({ readOnly: true });
    try {
      await act(async () =>
        [...document.querySelectorAll('[role="tab"]')].find((tab) => tab.textContent === 'Formulas').click()
      );
      await act(async () => document.querySelector('[data-ribbon-command="formulas.insert-function"]').click());
      await flush();
      const pane = document.querySelector('.sheet-functions'),
        argument = pane.querySelector('[aria-label="Values 1"]');
      assert(argument);
      await act(async () => props(argument).onChange({ target: { value: 'B1:B3' } }));
      await flush(240);
      assert.equal(outputNode().textContent, 'Number: 12');
      assert.equal(view.calls.at(-1).formula, 'SUM(B1:B3)');
      const insert = [...pane.querySelectorAll('button')].find(
        (button) => button.textContent === 'Insert formula into A1'
      );
      assert(insert.disabled);
      assert.deepEqual(view.strict, []);
      await act(async () =>
        [...document.querySelectorAll('button')]
          .find((button) => button.getAttribute('aria-label') === 'Close workbook tools')
          .click()
      );
      assert.equal(outputNode()?.textContent ?? null, null);
    } finally {
      await view.close();
    }
  }
);
check(
  'StrictMode and disposal keep a single physical request and cancel only its captured id without owner writes',
  async () => {
    const view = await mount({}, true);
    view.setHeld(true);
    await type('=SUM(B1:B3)');
    await flush(240);
    assert.equal(view.calls.length, 1);
    const call = view.calls[0];
    await view.close();
    assert.deepEqual(view.cancels, [call.requestId]);
    assert.deepEqual(view.writes, []);
    assert.deepEqual(view.strict, []);
    assert.equal(document.querySelector('.sheet-formula-result-preview'), null);
  }
);
check('preview-zone blur and Cancel never commit a draft in an editor without the strict name bridge', async () => {
  const view = await mount({ onCommandBatch: undefined });
  view.setHeld(true);
  try {
    await type('=SUM(B1:B3)');
    await flush(240);
    const cancel = [...document.querySelectorAll('button')].find((button) => button.textContent === 'Cancel preview');
    assert(cancel);
    await act(async () => {
      props(field()).onBlur({ currentTarget: field(), relatedTarget: cancel });
      cancel.click();
    });
    assert.deepEqual(view.writes, []);
    assert.equal(field().value, '=SUM(B1:B3)');
    await view.render({ ownerIdentity: 'replacement-runtime' });
    assert.equal(outputNode()?.textContent ?? null, null);
    assert.equal(field().value, '');
    assert.deepEqual(view.writes, []);
  } finally {
    await view.close();
  }
});
