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
  runs = await fs.mkdtemp(path.join(component, '.build/tests/calculate-now-ui-'));
const entry = path.join(runs, 'entry.ts'),
  bundle = path.join(runs, 'test.mjs');
await fs.writeFile(
  entry,
  `export {SpreadsheetGrid} from ${JSON.stringify(path.join(component, 'ui/SpreadsheetGrid.tsx'))};export {configureXlsxHost} from ${JSON.stringify(path.join(component, 'ui/host.tsx'))};export {ribbonCommand,ribbonCapability} from ${JSON.stringify(path.join(component, 'ui/ribbonCommands.ts'))};`
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
      return 600;
    },
  },
  clientHeight: {
    get() {
      return 400;
    },
  },
});
globalThis.ResizeObserver = class {
  constructor(fn) {
    this.fn = fn;
  }
  observe() {
    this.fn([{ contentRect: { width: 600, height: 400 } }]);
  }
  disconnect() {}
};
globalThis.requestAnimationFrame = () => 1;
globalThis.cancelAnimationFrame = () => {};
const { createRoot } = await import('react-dom/client'),
  { SpreadsheetGrid, configureXlsxHost, ribbonCommand, ribbonCapability } = await import(pathToFileURL(bundle));
configureXlsxHost({
  api: { copyText: async () => {}, readClipboardText: async () => '' },
  Icon: () => React.createElement('svg'),
  useI18n: () => ({ locale: 'en' }),
});
const check = (name, fn) => {
  test(name, async () => {
    await fn();
  });
};
after(async () => {
  dom.window.close();
});
const defer = () => {
  let resolve, reject;
  const promise = new Promise((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
};
const command = () =>
  document.querySelector('.sheet-ribbon-tabpanel:not([hidden]) [data-ribbon-command="formulas.calculate-now"]');
const notice = () => document.querySelector('[aria-label="Calculation status"]');
const sheet = (name) => ({
  name,
  state: 'visible',
  path: 'xl/worksheets/' + name + '.xml',
  rows: 3,
  columns: 2,
  cells: { A1: { value: 2 }, B1: { formula: 'A1+1', calculation: { status: 'calculated', value: 3 } } },
  merges: [],
  widths: [],
  hiddenRows: [],
  rowHeights: [],
  freeze: { rows: 0, columns: 0 },
});
async function mount(initial = {}) {
  const root = createRoot(document.getElementById('root')),
    writes = [],
    legacy = [];
  let props = {
    sheets: [sheet('Sheet1'), sheet('Sheet2')],
    documentId: 'doc-A',
    ownerIdentity: 'owner-A',
    storageKey: 'opaque-head-A',
    active: true,
    readOnly: false,
    onEdit: (x) => legacy.push(x),
    onBatch: (x) => legacy.push(x),
    onCommandBatch: async (x) => writes.push(x),
    ...initial,
  };
  const render = async (patch) => {
    props = { ...props, ...patch };
    await act(async () => root.render(React.createElement(SpreadsheetGrid, props)));
  };
  await render();
  await act(async () =>
    [...document.querySelectorAll('[role="tab"]')].find((node) => node.textContent === 'Formulas').click()
  );
  return {
    writes,
    legacy,
    render,
    close: async () => {
      await act(async () => root.unmount());
      await new Promise((resolve) => setTimeout(resolve, 2));
    },
  };
}
async function click(node) {
  assert(node);
  assert.equal(node.disabled, false);
  await act(async () => node.click());
}
async function go(ref) {
  const input = document.querySelector('[aria-label="Go to cell"]');
  await act(async () => {
    Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(input, ref);
    input.dispatchEvent(new window.Event('input', { bubbles: true }));
  });
  await act(async () =>
    input.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
  );
}

check('registered Calculate Now is active-workbook only, editable XLSX and momentary', async () => {
  const d = ribbonCommand('formulas.calculate-now');
  assert.deepEqual(d.matrixIds, ['XL-F14']);
  assert.equal(d.group, 'Calculation');
  assert.equal(d.tab, 'Formulas');
  assert.equal(d.kind, 'workbook');
  assert.equal(d.requiresEdit, true);
  assert.deepEqual(d.supportedProfiles, ['xlsx']);
  assert.match(d.hint, /all worksheets in the current workbook/);
  assert.match(d.hint, /one Undo step/);
  const ui = await mount();
  try {
    assert(command());
    assert.equal(command().textContent, 'Calculate Now');
    assert.equal(command().hasAttribute('aria-pressed'), false);
    assert.match(command().title, /current workbook/);
    assert(!document.querySelector('[data-ribbon-command="formulas.calculate-sheet"]'));
    assert.equal(notice(), null);
  } finally {
    await ui.close();
  }
});
check(
  'actual click captures one exact strict batch, awaits acceptance and suppresses duplicate before host busy',
  async () => {
    const held = defer(),
      calls = [],
      replacement = [];
    const ui = await mount({
      onCommandBatch: async (batch) => {
        calls.push(batch);
        await held.promise;
      },
    });
    try {
      const button = command();
      await act(async () => {
        button.click();
        button.click();
      });
      assert.deepEqual(calls, [[{ type: 'recalculate' }]]);
      assert.equal(command().disabled, true);
      assert.equal(notice().dataset.calculationPhase, 'running');
      assert(!notice().textContent.includes('finished'));
      await ui.render({ storageKey: 'opaque-accepted-head', onCommandBatch: async (batch) => replacement.push(batch) });
      assert.equal(notice().dataset.calculationPhase, 'running');
      await act(async () => held.resolve());
      assert.equal(notice().dataset.calculationPhase, 'completed');
      assert.equal(notice().textContent, 'Current workbook calculation finished.');
      assert.equal(command().disabled, false);
      assert.deepEqual(replacement, []);
      assert.deepEqual(ui.legacy, []);
    } finally {
      held.resolve();
      await ui.close();
    }
  }
);
check('owner rejection is not success and routes real current error without legacy writes', async () => {
  const held = defer(),
    calls = [];
  const ui = await mount({
    onCommandBatch: async (batch) => {
      calls.push(batch);
      await held.promise;
    },
  });
  try {
    await click(command());
    await act(async () => held.reject(Error('Captured owner revision changed')));
    assert.equal(notice().dataset.calculationPhase, 'failed');
    assert.match(notice().textContent, /not accepted/);
    assert.match(document.querySelector('[role="alert"]').textContent, /Captured owner revision changed/);
    assert.equal(command().disabled, false);
    assert.equal(calls.length, 1);
    assert.deepEqual(ui.legacy, []);
  } finally {
    held.resolve();
    await ui.close();
  }
});
check(
  'retired owner completion masks status and does not call replacement callback; latest owner retries after physical settlement',
  async () => {
    const held = defer(),
      a = [],
      b = [];
    const ui = await mount({
      onCommandBatch: async (batch) => {
        a.push(batch);
        await held.promise;
      },
    });
    try {
      await click(command());
      await ui.render({
        documentId: 'doc-B',
        ownerIdentity: 'owner-B',
        storageKey: 'opaque-head-B',
        onCommandBatch: async (batch) => b.push(batch),
      });
      assert.equal(notice(), null);
      assert.equal(command().disabled, true);
      await act(async () => held.resolve());
      assert.equal(notice(), null);
      assert.deepEqual(b, []);
      assert.equal(command().disabled, false);
      await click(command());
      assert.deepEqual(b, [[{ type: 'recalculate' }]]);
      assert.equal(notice().dataset.calculationPhase, 'completed');
      assert.equal(a.length, 1);
      assert.deepEqual(ui.legacy, []);
    } finally {
      held.resolve();
      await ui.close();
    }
  }
);
check(
  'inactive round trip, worksheet round trip and changed selection cannot revive old accepted feedback',
  async () => {
    for (const kind of ['activation', 'sheet', 'selection']) {
      const held = defer();
      const ui = await mount({ onCommandBatch: async () => held.promise });
      try {
        await click(command());
        if (kind === 'activation') {
          await ui.render({ active: false });
          assert.equal(notice(), null);
          await ui.render({ active: true });
        } else if (kind === 'sheet') {
          await click([...document.querySelectorAll('.sheet-tabs button')].find((n) => n.textContent === 'Sheet2'));
          assert.equal(notice(), null);
          await click([...document.querySelectorAll('.sheet-tabs button')].find((n) => n.textContent === 'Sheet1'));
        } else {
          await go('B1');
          await go('A1');
        }
        await act(async () => held.resolve());
        assert.equal(notice(), null);
        assert.deepEqual(ui.legacy, []);
      } finally {
        held.resolve();
        await ui.close();
      }
    }
  }
);
check(
  'readonly busy inactive missing strict callback and passive/delimited profiles never invoke calculation',
  async () => {
    for (const [patch, reason, hidden] of [
      [{ readOnly: true }, /read-only/, false],
      [{ busy: true }, /Wait/, false],
      [{ active: false }, /Activate/, false],
      [{ onCommandBatch: undefined }, /unavailable/, false],
      [{ ownerIdentity: undefined }, /unavailable/, false],
      [{ ownerIdentity: '' }, /unavailable/, false],
      [{ ownerIdentity: ' ' }, /unavailable/, false],
      [{ literal: true }, null, true],
      [{ passiveNotice: 'Stored passive workbook' }, null, true],
    ]) {
      const calls = [];
      const ui = await mount({ onCommandBatch: async (x) => calls.push(x), ...patch });
      try {
        if (hidden) assert.equal(command(), null);
        else {
          assert.equal(command().disabled, true);
          assert.match(command().title, reason);
          await act(async () => command().click());
        }
        assert.deepEqual(calls, []);
        assert.deepEqual(ui.legacy, []);
        assert.equal(notice(), null);
      } finally {
        await ui.close();
      }
    }
    const base = { profile: 'xlsx', active: true, canEdit: true, busy: false, hasHandler: true };
    assert.equal(ribbonCapability('formulas.calculate-now', { ...base, canEdit: false }).state, 'disabled');
    assert.equal(ribbonCapability('formulas.calculate-now', { ...base, profile: 'delimited' }).state, 'hidden');
  }
);
check('a completed acceptance after unmount settles without retained UI or another command', async () => {
  const held = defer(),
    calls = [];
  const ui = await mount({
    onCommandBatch: async (x) => {
      calls.push(x);
      await held.promise;
    },
  });
  await click(command());
  await ui.close();
  await act(async () => held.resolve());
  assert.deepEqual(calls, [[{ type: 'recalculate' }]]);
  assert.equal(document.querySelector('[aria-label="Calculation status"]'), null);
});
check(
  'late rejection after owner-only replacement with the same document/head cannot land an error on the new owner',
  async () => {
    const held = defer();
    const ui = await mount({ onCommandBatch: async () => held.promise });
    try {
      await click(command());
      await ui.render({ ownerIdentity: 'owner-B' });
      assert.equal(notice(), null);
      await act(async () => held.reject(Error('Obsolete owner refused calculation')));
      assert.equal(notice() === null, true);
      assert.equal(document.querySelector('[role="alert"]') === null, true);
    } finally {
      held.resolve();
      await ui.close();
    }
  }
);
