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
  runs = await fs.mkdtemp(path.join(component, '.build/tests/calculation-lifecycle-ui-'));
const entry = path.join(runs, 'entry.ts'),
  bundle = path.join(runs, 'test.mjs');
await fs.writeFile(
  entry,
  `export {SpreadsheetGrid} from ${JSON.stringify(path.join(component, 'ui/SpreadsheetGrid.tsx'))};export {configureXlsxHost} from ${JSON.stringify(path.join(component, 'ui/host.tsx'))};export * from ${JSON.stringify(path.join(component, 'ui/calculationJobs.ts'))};export {WorkbookObjects} from ${JSON.stringify(path.join(component, 'ui/WorkbookObjects.tsx'))};export {WorkbookWatch} from ${JSON.stringify(path.join(component, 'ui/WorkbookWatch.tsx'))};export {classifyWatchCell} from ${JSON.stringify(path.join(component, 'ui/watchWindow.ts'))};export {acceptedRuleValue,conditionalFill} from ${JSON.stringify(path.join(component, 'core/rules-model.mjs'))};`
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
  {
    SpreadsheetGrid,
    configureXlsxHost,
    createCalculationJobController,
    calculationFrameKey,
    CalculationJobRetired,
    validateCalculationReceipt,
    WorkbookWatch,
    WorkbookObjects,
    classifyWatchCell,
    acceptedRuleValue,
    conditionalFill,
  } = await import(pathToFileURL(bundle));
let hostCopyText = async () => {};
configureXlsxHost({
  api: { copyText: (text) => hostCopyText(text), readClipboardText: async () => '' },
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
const original = { documentId: 'durable-A', ownerId: 'native-A', state: 'head-A', revision: 'revision-A' };
const policy = {
  schema: 1,
  mode: 'auto',
  recalculateBeforeSave: true,
  pendingFormulaCells: 0,
  canCalculateSheet: true,
  canChangeOptions: true,
};
function result(action, requestId, patch = {}) {
  return {
    schema: 1,
    requestId,
    initiator: { ...original },
    action: { ...action },
    consistency: 'independent-workbook-commits',
    outcome: 'completed',
    results: [
      {
        documentId: original.documentId,
        ownerId: original.ownerId,
        name: 'A.xlsx',
        beforeState: original.state,
        afterState: 'head-A-accepted',
        status: 'accepted',
      },
    ],
    ...patch,
  };
}
function controls(patch = {}) {
  return {
    receipt: { ...original },
    policy: { ...policy },
    optionsScope: 'application',
    applicationMode: { mode: 'auto', revision: 0 },
    run: async (action, id) => result(action, id),
    cancel: async () => ({ outcome: 'cancelled' }),
    ...patch,
  };
}
const command = (id) => document.querySelector(`.sheet-ribbon-tabpanel:not([hidden]) [data-ribbon-command="${id}"]`);
const notice = () => document.querySelector('[aria-label="Calculation status"]');
const pane = () => document.querySelector('section[aria-label="Calculation Options"]');
const label = (name) => pane()?.querySelector(`[aria-label="${name}"]`);
const sheet = (name, pending = false) => ({
  name,
  state: 'visible',
  path: 'xl/worksheets/' + name + '.xml',
  rows: 3,
  columns: 2,
  cells: {
    A1: { value: 2 },
    B1: pending
      ? {
          formula: 'A1+1',
          calculation: { status: 'stored', pending: true, error: 'Formula result is pending calculation' },
          cachedResult: { authority: 'stored', status: 'calculated', value: 3 },
          cachedValue: 999,
        }
      : { formula: 'A1+1', calculation: { status: 'calculated', value: 3 } },
  },
  merges: [],
  widths: [],
  hiddenRows: [],
  rowHeights: [],
  freeze: { rows: 0, columns: 0 },
});
async function mount(initial = {}) {
  const root = createRoot(document.getElementById('root')),
    legacy = [],
    strict = [];
  let props = {
    sheets: [sheet('Sheet1'), sheet('Sheet2')],
    documentId: 'ui-A',
    ownerIdentity: 'runtime-A',
    storageKey: 'opaque-UI-head-A',
    active: true,
    readOnly: false,
    onEdit: (x) => legacy.push(x),
    onBatch: (x) => legacy.push(x),
    onCommandBatch: async (x) => strict.push(x),
    calculationControls: controls(),
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
    legacy,
    strict,
    render,
    props: () => props,
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
async function open() {
  await click(document.querySelector('[data-workbook-tool="calculation"]'));
}
function uiFrame(c = controls()) {
  return {
    scope: {
      documentId: 'ui-A',
      ownerIdentity: 'runtime-A',
      state: 'opaque-UI-head-A',
      sheet: 'Sheet1',
      activationEpoch: 0,
      modeEpoch: 0,
      navigationGeneration: 0,
      active: true,
      normal: true,
      readOnly: false,
      busy: false,
      profile: 'xlsx',
    },
    controls: c,
  };
}

check(
  'Name Box preserves live and invalid drafts, then displays the actual accepted cell/range after navigation or worksheet/selection change',
  async () => {
    const ui = await mount(),
      field = () => document.querySelector('[aria-label="Go to cell"]');
    const draft = async (value) =>
      act(async () => {
        const input = field();
        Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(input, value);
        input.dispatchEvent(new window.Event('input', { bubbles: true }));
      });
    const go = async () =>
      act(async () =>
        field().dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
      );
    try {
      await draft('B2');
      await ui.render({ storageKey: 'same-owner-new-head' });
      assert.equal(field().value, 'B2', 'A same-selection render must not clear an unfinished draft');
      await go();
      assert.equal(field().value, '');
      assert.equal(field().placeholder, 'B2');
      await draft('B2');
      await go();
      assert.equal(field().value, '', 'A successful Go to the already-selected cell must also retire its draft');
      await draft('A1:B2');
      await go();
      assert.equal(field().value, '');
      assert.equal(field().placeholder, 'A1:B2');
      await draft('not a cell');
      await go();
      assert.equal(field().value, 'not a cell');
      assert.match(document.querySelector('.sheet-formula-bar [role="alert"]').textContent, /Enter a cell/);
      assert.equal(field().placeholder, 'A1:B2');
      await click([...document.querySelectorAll('.sheet-tabs button')].find((n) => n.textContent === 'Sheet2'));
      assert.equal(field().value, '');
      assert.equal(field().placeholder, 'A1');
      assert.equal(document.querySelector('.sheet-formula-bar [role="alert"]'), null);
      await draft('C3');
      await ui.render({ storageKey: 'another-head-same-selection' });
      assert.equal(field().value, 'C3');
      const cell = document.querySelector('[aria-label="Sheet2 B1"]').closest('td');
      await act(async () => cell.dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true })));
      assert.equal(field().value, '');
      assert.equal(field().placeholder, 'B1');
      assert.deepEqual(ui.strict, []);
      assert.deepEqual(ui.legacy, []);
    } finally {
      await ui.close();
    }
  }
);

check(
  'real lifecycle ribbon binds all-open Now, Sheet and one retained Options pane without changing local views or legacy callbacks',
  async () => {
    const calls = [],
      ui = await mount({
        calculationControls: controls({
          run: async (action, id) => {
            calls.push(action);
            return result(action, id);
          },
        }),
      });
    try {
      assert.match(command('formulas.calculate-now').title, /all open editable/);
      assert.equal(command('formulas.calculate-now').hasAttribute('aria-pressed'), false);
      assert(command('formulas.calculate-sheet'));
      await open();
      assert.equal(document.querySelectorAll('.sheet-task-pane').length, 1);
      assert(label('Automatic').checked);
      await click(command('formulas.calculate-sheet'));
      assert.deepEqual(calls, [{ type: 'calculateSheet', sheet: 'Sheet1' }]);
      assert.equal(notice().dataset.calculationPhase, 'completed');
      assert.equal(document.querySelector('[data-calculation-result="accepted"]').textContent, 'A.xlsx: accepted');
      assert.deepEqual(ui.strict, []);
      assert.deepEqual(ui.legacy, []);
    } finally {
      await ui.close();
    }
  }
);
check(
  'Manual is not optimistic; accepted metadata controls mode/save/pending and explicit application mismatch',
  async () => {
    const held = defer(),
      calls = [],
      c = controls({
        run: async (action, id) => {
          calls.push(action);
          await held.promise;
          return result(action, id);
        },
      }),
      ui = await mount({ calculationControls: c });
    try {
      await open();
      await click(label('Manual'));
      assert.deepEqual(calls, [{ type: 'setOptions', mode: 'manual' }]);
      assert(label('Automatic').checked);
      assert.equal(label('Manual').checked, false);
      assert(label('Recalculate before saving').checked);
      assert(label('Manual').disabled || label('Manual').closest('fieldset').disabled);
      await ui.render({
        storageKey: 'opaque-new-head',
        calculationControls: {
          ...c,
          receipt: { ...original, state: 'new-head', revision: 'new-revision' },
          policy: { ...policy, mode: 'manual', pendingFormulaCells: 2 },
          applicationMode: { mode: 'manual', revision: 1 },
        },
      });
      await act(async () => held.resolve());
      assert(label('Manual').checked);
      assert.equal(label('Pending formulas').dataset.pendingFormulas, '2');
      assert.match(notice().textContent, /options updated/);
      await ui.render({
        calculationControls: {
          ...ui.props().calculationControls,
          policy: { ...policy, mode: 'auto' },
          applicationMode: { mode: 'manual', revision: 1 },
        },
      });
      assert.match(label('Application calculation mode').textContent, /different calculation mode/);
      assert(label('Manual').checked);
      assert.match(pane().textContent, /Current workbook: Automatic/);
    } finally {
      held.resolve();
      await ui.close();
    }
  }
);
check(
  'save-before-recalculation false sends exact explicit mode/options and keeps the old checked value until acknowledgement',
  async () => {
    const held = defer(),
      calls = [],
      c = controls({
        policy: { ...policy, mode: 'manual' },
        applicationMode: { mode: 'manual', revision: 1 },
        run: async (action, id) => {
          calls.push(action);
          await held.promise;
          return result(action, id);
        },
      }),
      ui = await mount({ calculationControls: c });
    try {
      await open();
      await click(label('Recalculate before saving'));
      assert.deepEqual(calls, [{ type: 'setOptions', mode: 'manual', recalculateBeforeSave: false }]);
      assert(label('Recalculate before saving').checked);
      await ui.render({
        calculationControls: {
          ...c,
          receipt: { ...original, state: 'accepted' },
          policy: { ...policy, mode: 'manual', recalculateBeforeSave: false },
        },
      });
      await act(async () => held.resolve());
      assert.equal(label('Recalculate before saving').checked, false);
    } finally {
      held.resolve();
      await ui.close();
    }
  }
);
check(
  'all-open results preserve accepted/skipped/refused rows and partial outcome, not cross-workbook rollback',
  async () => {
    const calls = [],
      ui = await mount({
        calculationControls: controls({
          run: async (action, id) => {
            calls.push(action);
            return result(action, id, {
              outcome: 'partial',
              results: [
                ...result(action, id).results,
                {
                  documentId: 'R',
                  ownerId: 'R-owner',
                  name: 'Read only.xlsx',
                  beforeState: 'R0',
                  status: 'skipped',
                  reason: 'Read-only workbook',
                },
                {
                  documentId: 'B',
                  ownerId: 'B-owner',
                  name: 'B.xlsx',
                  beforeState: 'B0',
                  status: 'refused',
                  reason: 'Source revision changed',
                },
              ],
            });
          },
        }),
      });
    try {
      await click(command('formulas.calculate-now'));
      assert.deepEqual(calls, [{ type: 'calculateAllOpen' }]);
      assert.equal(notice().dataset.calculationPhase, 'partial');
      assert.match(notice().textContent, /Accepted changes were kept/);
      assert.equal(document.querySelectorAll('[data-calculation-result]').length, 3);
      assert.match(
        document.querySelector('[aria-label="Calculation results"]').textContent,
        /1 workbooks accepted; 1 skipped/
      );
      assert.deepEqual(ui.strict, []);
    } finally {
      await ui.close();
    }
  }
);
check(
  'Cancel Finishing remains pending until the original actual job receipt settles; same-tick duplicate cannot launch twice',
  async () => {
    const held = defer(),
      calls = [],
      cancels = [],
      c = controls({
        run: async (action, id, progress) => {
          calls.push(action);
          progress({
            requestId: id,
            phase: 'calculating',
            completed: 0,
            total: 2,
            currentName: 'A.xlsx',
            cancellable: true,
          });
          await held.promise;
          return result(action, id);
        },
        cancel: async (id) => {
          cancels.push(id);
          return { outcome: 'finishing' };
        },
      }),
      ui = await mount({ calculationControls: c });
    try {
      const b = command('formulas.calculate-now');
      await act(async () => {
        b.click();
        b.click();
      });
      assert.equal(calls.length, 1);
      assert.equal(notice().dataset.calculationPhase, 'running');
      await click(document.querySelector('[aria-label="Cancel calculation"]'));
      assert.equal(notice().dataset.calculationPhase, 'finishing');
      assert.equal(command('formulas.calculate-now').disabled, true);
      assert.equal(document.querySelector('[aria-label="Cancel calculation"]').disabled, true);
      assert.equal(cancels.length, 1);
      await act(async () => held.resolve());
      assert.equal(notice().dataset.calculationPhase, 'completed');
      assert.equal(command('formulas.calculate-now').disabled, false);
    } finally {
      held.resolve();
      await ui.close();
    }
  }
);
check('granted cancel waits for terminal receipt and retains earlier accepted independent workbooks', async () => {
  const held = defer(),
    c = controls({
      run: async (action, id) => {
        await held.promise;
        return result(action, id, {
          outcome: 'cancelled',
          results: [
            ...result(action, id).results,
            {
              documentId: 'B',
              ownerId: 'B-owner',
              name: 'B.xlsx',
              beforeState: 'B0',
              status: 'notRun',
              reason: 'Calculation cancelled',
            },
          ],
        });
      },
    }),
    ui = await mount({ calculationControls: c });
  try {
    await click(command('formulas.calculate-now'));
    await click(document.querySelector('[aria-label="Cancel calculation"]'));
    assert.equal(notice().dataset.calculationPhase, 'stopping');
    assert.equal(command('formulas.calculate-now').disabled, true);
    await act(async () => held.resolve());
    assert.equal(notice().dataset.calculationPhase, 'cancelled');
    assert.match(document.querySelector('[data-calculation-result="accepted"]').textContent, /A.xlsx/);
    assert(document.querySelector('[data-calculation-result="notRun"]'));
  } finally {
    held.resolve();
    await ui.close();
  }
});
check(
  'head advancement keeps truthful acceptance, while owner-only/receipt-owner replacement masks late success and rejects stale DOM callbacks',
  async () => {
    for (const kind of ['runtime', 'receipt']) {
      const held = defer(),
        a = [],
        b = [],
        old = controls({
          run: async (action, id) => {
            a.push(action);
            await held.promise;
            return result(action, id);
          },
        }),
        ui = await mount({ calculationControls: old });
      try {
        await open();
        const radio = label('Manual'),
          key = Object.keys(radio).find((key) => key.startsWith('__reactProps$')),
          staleHandler = radio[key].onChange;
        await click(command('formulas.calculate-now'));
        const newer = controls({
          receipt: { ...original, ownerId: 'native-B' },
          run: async (action, id) => {
            b.push(action);
            return result(action, id, {
              initiator: { ...original, ownerId: 'native-B' },
              results: [
                {
                  documentId: original.documentId,
                  ownerId: 'native-B',
                  name: 'B.xlsx',
                  beforeState: original.state,
                  afterState: 'B-next',
                  status: 'accepted',
                },
              ],
            });
          },
        });
        await ui.render({ ...(kind === 'runtime' ? { ownerIdentity: 'runtime-B' } : {}), calculationControls: newer });
        assert.equal(notice(), null);
        assert.equal(command('formulas.calculate-now').disabled, true);
        await act(async () => staleHandler());
        assert.deepEqual(b, []);
        await act(async () => held.resolve());
        assert.equal(notice(), null);
        assert(!document.querySelector('[data-calculation-result]'));
        assert.deepEqual(a, [{ type: 'calculateAllOpen' }]);
        assert.deepEqual(b, []);
      } finally {
        held.resolve();
        await ui.close();
      }
    }
  }
);
check(
  'dock hiding and same-owner worksheet navigation preserve group progress; activation round trips retire feedback',
  async () => {
    const held = defer(),
      ui = await mount({
        calculationControls: controls({
          run: async (action, id) => {
            await held.promise;
            return result(action, id);
          },
        }),
      });
    try {
      await click(command('formulas.calculate-now'));
      await click(document.querySelector('[aria-label="Close workbook tools"]'));
      assert.equal(notice().dataset.calculationPhase, 'running');
      await act(async () => held.resolve());
      assert.equal(notice().dataset.calculationPhase, 'completed');
      await open();
      assert(document.querySelector('[data-calculation-result="accepted"]'));
    } finally {
      held.resolve();
      await ui.close();
    }
    for (const kind of ['activation', 'sheet']) {
      const held = defer(),
        ui = await mount({
          calculationControls: controls({
            run: async (action, id) => {
              await held.promise;
              return result(action, id);
            },
          }),
        });
      try {
        await click(command('formulas.calculate-now'));
        if (kind === 'activation') {
          await ui.render({ active: false });
          assert.equal(notice(), null);
          await ui.render({ active: true });
        } else {
          await click([...document.querySelectorAll('.sheet-tabs button')].find((n) => n.textContent === 'Sheet2'));
          assert.equal(notice().dataset.calculationPhase, 'running');
          assert(document.querySelector('[aria-label="Cancel calculation"]'));
          await click([...document.querySelectorAll('.sheet-tabs button')].find((n) => n.textContent === 'Sheet1'));
        }
        await act(async () => held.resolve());
        if (kind === 'sheet') assert.equal(notice().dataset.calculationPhase, 'completed');
        else assert.equal(notice(), null);
      } finally {
        held.resolve();
        await ui.close();
      }
    }
  }
);
check(
  'readonly is inspectable without writes; protected sheet/options are disabled; passive and literal lifecycle routes are absent',
  async () => {
    for (const patch of [
      { readOnly: true },
      {
        calculationControls: controls({
          policy: { ...policy, canCalculateSheet: false, canChangeOptions: false, reason: 'Unknown formula closure' },
        }),
      },
    ]) {
      const calls = [],
        ui = await mount({
          ...patch,
          calculationControls: {
            ...(patch.calculationControls ?? controls()),
            run: async (a) => {
              calls.push(a);
              throw Error('must not run');
            },
          },
        });
      try {
        await open();
        assert.equal(command('formulas.calculate-sheet').disabled, true);
        assert(label('Manual').closest('fieldset').disabled);
        assert.equal(label('Recalculate before saving').disabled, true);
        assert.deepEqual(calls, []);
      } finally {
        await ui.close();
      }
    }
    for (const patch of [{ literal: true }, { passiveNotice: 'Passive stored workbook' }]) {
      const ui = await mount(patch);
      try {
        assert.equal(command('formulas.calculate-now'), null);
        assert.equal(command('formulas.calculate-sheet'), null);
        assert.equal(document.querySelector('[data-workbook-tool="calculation"]'), null);
      } finally {
        await ui.close();
      }
    }
    for (const ownerIdentity of [undefined, '', ' ']) {
      const ui = await mount({ ownerIdentity });
      try {
        assert.equal(command('formulas.calculate-now').disabled, true);
        assert.equal(command('formulas.calculate-sheet').disabled, true);
      } finally {
        await ui.close();
      }
    }
  }
);
check(
  'pending displayed cache is clearly stored, has no current numeric aria value, and respects Show formulas',
  async () => {
    const ui = await mount({
      sheets: [sheet('Sheet1', true)],
      calculationControls: controls({ policy: { ...policy, mode: 'manual', pendingFormulaCells: 1 } }),
    });
    try {
      const input = document.querySelector('[aria-label="Sheet1 B1"]'),
        calculated = document.querySelector('[aria-label="Sheet1 B1 calculated result"]');
      assert.equal(input.value, 'Stored: 3');
      assert.match(input.getAttribute('aria-description'), /Pending.*stored and not current/);
      assert.equal(calculated.hasAttribute('data-value'), false);
      assert.equal(calculated.textContent, 'Stored: 3');
      assert.equal(document.querySelector('[aria-label="Workbook calculation policy"]').dataset.pendingFormulas, '1');
      await click(command('view.show-formulas'));
      assert.equal(document.querySelector('[aria-label="Sheet1 B1"]').value, '=A1+1');
    } finally {
      await ui.close();
    }
  }
);
check(
  'a late rejected job after a different accepted head cannot place obsolete errors or a stuck running notice on that head',
  async () => {
    const held = defer(),
      old = controls({ run: async () => held.promise }),
      ui = await mount({ calculationControls: old });
    try {
      await click(command('formulas.calculate-now'));
      await ui.render({
        storageKey: 'different-accepted-UI',
        calculationControls: { ...old, receipt: { ...original, state: 'another-head', revision: 'another-revision' } },
      });
      await act(async () => held.reject(Error('Original revision no longer current')));
      assert.equal(notice(), null);
      assert.equal(document.querySelector('[role="alert"]'), null);
      assert.equal(command('formulas.calculate-now').disabled, false);
    } finally {
      held.resolve();
      await ui.close();
    }
  }
);
check(
  'application Undo mismatch preserves the real default during save-option changes instead of silently resetting its mode',
  async () => {
    const calls = [],
      ui = await mount({
        calculationControls: controls({
          policy: { ...policy, mode: 'auto' },
          applicationMode: { mode: 'manual', revision: 8 },
          run: async (a, id) => {
            calls.push(a);
            return result(a, id);
          },
        }),
      });
    try {
      await open();
      assert(label('Manual').checked);
      assert.match(pane().textContent, /Current workbook: Automatic/);
      await click(label('Recalculate before saving'));
      assert.deepEqual(calls, [{ type: 'setOptions', mode: 'manual', recalculateBeforeSave: false }]);
    } finally {
      await ui.close();
    }
  }
);
check(
  'a directly invoked retired ribbon handler cannot retarget a replacement owner or use its run callback',
  async () => {
    const oldCalls = [],
      newCalls = [],
      ui = await mount({
        calculationControls: controls({
          run: async (a, id) => {
            oldCalls.push(a);
            return result(a, id);
          },
        }),
      });
    try {
      const button = command('formulas.calculate-now'),
        key = Object.keys(button).find((key) => key.startsWith('__reactProps$')),
        stale = button[key].onClick;
      await ui.render({
        ownerIdentity: 'replacement',
        calculationControls: controls({
          receipt: { ...original, ownerId: 'other' },
          run: async (a) => {
            newCalls.push(a);
            throw Error('must not retarget');
          },
        }),
      });
      await act(async () => stale());
      assert.deepEqual(oldCalls, []);
      assert.deepEqual(newCalls, []);
      assert.equal(notice(), null);
      assert.equal(document.querySelector('[role="alert"]'), null);
    } finally {
      await ui.close();
    }
  }
);
check(
  'F9 and Shift-F9 use the same awaited captured all-open/sheet jobs and never commit or blur the selected cell',
  async () => {
    const calls = [],
      ui = await mount({
        calculationControls: controls({
          run: async (a, id) => {
            calls.push(a);
            return result(a, id);
          },
        }),
      });
    const key = async (target, options = {}) => {
      let prevented;
      await act(async () => {
        const e = new window.KeyboardEvent('keydown', { key: 'F9', bubbles: true, cancelable: true, ...options });
        target.dispatchEvent(e);
        prevented = e.defaultPrevented;
      });
      return prevented;
    };
    try {
      const input = document.querySelector('[aria-label="Sheet1 A1"]');
      await act(async () => input.focus());
      assert.equal(await key(input), true);
      assert.deepEqual(calls, [{ type: 'calculateAllOpen' }]);
      assert.deepEqual(ui.legacy, []);
      assert.deepEqual(ui.strict, []);
      assert.equal(notice().dataset.calculationPhase, 'completed');
      await click([...document.querySelectorAll('.sheet-tabs button')].find((n) => n.textContent === 'Sheet2'));
      assert.equal(await key(document.querySelector('[aria-label="Sheet2 A1"]'), { shiftKey: true }), true);
      assert.deepEqual(calls, [{ type: 'calculateAllOpen' }, { type: 'calculateSheet', sheet: 'Sheet2' }]);
      assert.deepEqual(ui.legacy, []);
      assert.deepEqual(ui.strict, []);
    } finally {
      await ui.close();
    }
  }
);
check(
  'calculation shortcuts ignore live cell/formula/text drafts, composition and modifiers without triggering blur commits',
  async () => {
    const calls = [],
      ui = await mount({
        calculationControls: controls({
          run: async (a, id) => {
            calls.push(a);
            return result(a, id);
          },
        }),
      });
    const key = async (target, options = {}) => {
      await act(async () =>
        target.dispatchEvent(
          new window.KeyboardEvent('keydown', { key: 'F9', bubbles: true, cancelable: true, ...options })
        )
      );
    };
    try {
      const input = document.querySelector('[aria-label="Sheet1 A1"]');
      await act(async () => {
        input.focus();
        input.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'F2', bubbles: true, cancelable: true }));
      });
      assert.equal(input.dataset.editing, 'true');
      await act(async () => {
        Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(input, '=9+1');
        input.dispatchEvent(new window.Event('input', { bubbles: true }));
      });
      assert.equal(input.value, '=9+1');
      await key(input);
      assert.equal(document.activeElement, input);
      assert.deepEqual(calls, []);
      assert.deepEqual(ui.legacy, []);
      await act(async () =>
        input.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }))
      );
      assert.equal(input.dataset.editing, 'false');
      for (const options of [
        { isComposing: true },
        { ctrlKey: true },
        { metaKey: true },
        { altKey: true },
        { ctrlKey: true, shiftKey: true },
      ])
        await key(input, options);
      input.dataset.composing = 'true';
      await key(input);
      input.dataset.composing = 'false';
      for (const field of [
        document.querySelector('[aria-label="Go to cell"]'),
        document.querySelector('[aria-label="Formula or value"]'),
      ]) {
        assert(field);
        await act(async () => {
          field.focus();
          Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(field, '42');
          field.dispatchEvent(new window.Event('input', { bubbles: true }));
        });
        await key(field);
        assert.equal(document.activeElement, field);
      }
      assert.deepEqual(calls, []);
      assert.deepEqual(ui.legacy, []);
      assert.equal(notice(), null);
    } finally {
      await ui.close();
    }
  }
);
check(
  'shortcut duplicate/disabled and fallback routes never bypass the controller or pretend all-open support',
  async () => {
    const held = defer(),
      calls = [],
      ui = await mount({
        calculationControls: controls({
          run: async (a, id) => {
            calls.push(a);
            await held.promise;
            return result(a, id);
          },
        }),
      });
    try {
      const input = document.querySelector('[aria-label="Sheet1 A1"]');
      await act(async () => {
        input.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'F9', bubbles: true, cancelable: true }));
        input.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'F9', bubbles: true, cancelable: true }));
      });
      assert.equal(calls.length, 1);
      await act(async () => held.resolve());
    } finally {
      held.resolve();
      await ui.close();
    }
    for (const patch of [
      { readOnly: true },
      { busy: true },
      { active: false },
      { literal: true },
      { passiveNotice: 'Stored passive owner' },
      { ownerIdentity: undefined },
      { calculationControls: undefined },
    ]) {
      const calls = [],
        ui = await mount({
          calculationControls: controls({
            run: async (a) => {
              calls.push(a);
              throw Error('must not run');
            },
          }),
          ...patch,
        });
      try {
        await act(async () =>
          document
            .querySelector('[aria-label="Sheet1 A1"]')
            .dispatchEvent(new window.KeyboardEvent('keydown', { key: 'F9', bubbles: true, cancelable: true }))
        );
        assert.deepEqual(calls, []);
        assert.deepEqual(ui.strict, []);
        assert.deepEqual(ui.legacy, []);
        assert.equal(notice(), null);
      } finally {
        await ui.close();
      }
    }
  }
);
check('F9 never dispatches from the retained hidden Normal grid while Page Layout is active', async () => {
  const settings = {
      area: 'A1:B3',
      paper: 'Letter',
      orientation: 'portrait',
      marginsPt: { top: 36, right: 36, bottom: 36, left: 36, header: 18, footer: 18 },
      scaling: { mode: 'scale', percent: 100 },
      repeatRows: null,
      repeatColumns: null,
      gridlines: false,
    },
    page = {
      number: 1,
      rowIndices: [0, 1, 2],
      columnIndices: [0, 1],
      cells: ['A1', 'B1', 'A2', 'B2', 'A3', 'B3'].map((ref, n) => ({
        ref,
        text: ref,
        style: { fontSize: 11 },
        bbox: [36 + (n % 2) * 100, 36 + Math.floor(n / 2) * 24, 100, 24],
      })),
      objects: [],
    },
    calls = [];
  const ui = await mount({
    calculationControls: controls({
      run: async (a) => {
        calls.push(a);
        throw Error('must not run');
      },
    }),
    printControls: {
      documentId: 'runtime-A',
      revision: original.state,
      definition: async () => ({ sheet: 'Sheet1', revision: null, settings, canEdit: true, protectedReasons: [] }),
      preview: async (_sheet, _settings, _requestId, geometryProfile) => ({
        geometryProfile,
        id: 'accepted-page-job',
        state: original.state,
        units: 'pt',
        pageSizePt: [612, 792],
        scale: 1,
        sourceFingerprint: 'a'.repeat(64),
        revision: 'source-A',
        sheet: 'Sheet1',
        settings,
        pageCount: 1,
        pages: [page],
      }),
      page: async () => page,
      onBatch: async () => {},
      publish: async () => {},
      release: async () => {},
      cancel: async () => ({ granted: true }),
    },
  });
  try {
    await click([...document.querySelectorAll('[role="tab"]')].find((n) => n.textContent === 'View'));
    await click(command('view.page-layout'));
    assert(document.querySelector('.sheet-page-layout'));
    const input = document.querySelector('[aria-label="Sheet1 A1"]');
    assert(input.closest('.sheet-grid-viewport').hidden);
    await act(async () =>
      input.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'F9', bubbles: true, cancelable: true }))
    );
    assert.deepEqual(calls, []);
    assert.deepEqual(ui.legacy, []);
    assert.equal(notice(), null);
  } finally {
    await ui.close();
  }
});
check(
  'explicit pending is authoritative for masters and non-formula spill followers in Watch, conditional rules and real Grid rendering',
  async () => {
    const follower = {
      value: 2,
      dataType: 'number',
      spill: { schema: 1, master: 'A1', range: 'A1:A3', role: 'follower' },
      calculation: { status: 'stored', pending: true, error: 'Formula result is pending calculation' },
      cachedResult: { authority: 'stored', status: 'calculated', value: 2 },
    };
    const master = {
      ...follower,
      formula: 'SEQUENCE(3,1)',
      formulaType: 'array',
      spill: { ...follower.spill, role: 'master' },
    };
    const rules = [{ range: 'A1:A3', kind: 'greaterThan', values: [0], fillColor: '#FF0000', stopIfTrue: false }];
    for (const cell of [
      master,
      follower,
      { ...follower, calculation: { status: 'calculated', pending: true, value: 2 } },
    ]) {
      assert.equal(classifyWatchCell(cell).value.kind, 'stored');
      assert.deepEqual(acceptedRuleValue(cell), { qualified: false });
      assert.equal(conditionalFill(rules, 'A2', cell), undefined);
    }
    const root = createRoot(document.getElementById('root')),
      scope = {
        documentId: 'watch-A',
        ownerIdentity: 'watch-owner',
        state: 'manual-pending-head',
        activationEpoch: 0,
        active: true,
        sheets: ['Sheet1'],
      };
    try {
      await act(async () =>
        root.render(
          React.createElement(WorkbookWatch, {
            scope,
            read: async (bound, target) => ({ scope: bound, target, cell: follower }),
            selected: { sheet: 'Sheet1', reference: 'A2' },
          })
        )
      );
      await click([...document.querySelectorAll('button')].find((n) => n.textContent === 'Add selected cell'));
      assert.equal(
        document.querySelector('[aria-label="Watch value Sheet1!A2"]').textContent,
        'Stored result; not recalculated.'
      );
      assert(!document.body.textContent.includes('Number: 2'));
    } finally {
      await act(async () => root.unmount());
    }
    const current = sheet('Sheet1');
    current.cells.A2 = follower;
    const ui = await mount({ sheets: [current] });
    try {
      const input = document.querySelector('[aria-label="Sheet1 A2"]'),
        status = document.querySelector('[aria-label="Sheet1 A2 calculated result"]');
      assert.equal(input.value, 'Stored: 2');
      assert.match(input.getAttribute('aria-description'), /Pending calculation/);
      assert.equal(status.textContent, 'Stored: 2');
      assert.equal(status.hasAttribute('data-value'), false);
    } finally {
      await ui.close();
    }
  }
);
check(
  'object chart preview is masked in the very first commit of a new head or replacement session, before effect cleanup',
  async () => {
    for (const kind of ['head', 'owner']) {
      const root = createRoot(document.getElementById('root')),
        held = defer(),
        paints = [],
        entry = {
          id: '11111111-1111-4111-8111-111111111111',
          kind: 'chart',
          name: 'Sales trend',
          revision: 'a'.repeat(64),
          managed: true,
          canEdit: true,
          canDelete: true,
          anchor: { cell: 'E2' },
          chart: { kind: 'column', range: 'A1:B2', title: 'Accepted chart A', legend: 'bottom' },
        },
        preview = {
          ...entry,
          data: { categories: ['Old accepted category'], series: [{ name: 'Values', values: [2] }], hidden: [false] },
        },
        base = {
          documentId: 'object-A',
          ownerIdentity: 'runtime-A',
          revision: 'head-A',
          sheet: 'Sheet1',
          selection: 'A1:B2',
          objects: async () => ({ objects: [entry], total: 1, nextCursor: null }),
          object: async () => preview,
          onBatch: async () => {
            throw Error('No object write');
          },
        };
      const render = async (props, phase) =>
        act(async () =>
          root.render(
            React.createElement(
              React.Profiler,
              {
                id: 'object-paint',
                onRender: () => {
                  paints.push({
                    phase,
                    oldChart: !!document.querySelector('svg[aria-label="Accepted chart A"]'),
                    oldData: document.body.textContent.includes('Old accepted category'),
                  });
                },
              },
              React.createElement(WorkbookObjects, props)
            )
          )
        );
      try {
        await render(base, 'A');
        await click([...document.querySelectorAll('button')].find((n) => n.textContent === 'Sales trend (chart)'));
        assert(document.querySelector('svg[aria-label="Accepted chart A"]'));
        await render(
          {
            ...base,
            ...(kind === 'head' ? { revision: 'head-B' } : { ownerIdentity: 'runtime-B' }),
            objects: async () => held.promise,
          },
          'B'
        );
        const first = paints.find((p) => p.phase === 'B');
        assert(first);
        assert.equal(first.oldChart, false);
        assert.equal(first.oldData, false);
        assert.equal(document.querySelector('svg[aria-label="Accepted chart A"]'), null);
      } finally {
        held.resolve({ objects: [], total: 0, nextCursor: null });
        await act(async () => root.unmount());
      }
    }
  }
);
check(
  'Stored display provenance never leaks into actual selected-grid clipboard data or authored formulas',
  async () => {
    const follower = {
        value: 2,
        dataType: 'number',
        spill: { schema: 1, master: 'A1', range: 'A1:A3', role: 'follower' },
        calculation: { status: 'stored', pending: true },
        cachedResult: { authority: 'stored', status: 'calculated', value: 2 },
      },
      current = sheet('Sheet1', true);
    current.cells.A2 = follower;
    const ui = await mount({ sheets: [current] }),
      go = async (ref) => {
        const field = document.querySelector('[aria-label="Go to cell"]');
        await act(async () => {
          Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(field, ref);
          field.dispatchEvent(new window.Event('input', { bubbles: true }));
        });
        await act(async () =>
          field.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
        );
      },
      copied = [];
    try {
      if (command('view.show-formulas').checked) await click(command('view.show-formulas'));
      for (const [ref, expected] of [
        ['A2', '2'],
        ['B1', '=A1+1'],
      ]) {
        await go(ref);
        const input = document.querySelector('[aria-label="Sheet1 ' + ref + '"]');
        assert.match(input.value, /^Stored: /);
        const started = defer(),
          completed = defer();
        hostCopyText = async (value) => {
          copied.push(value);
          started.resolve();
          await completed.promise;
        };
        try {
          await act(async () => {
            const e = new window.Event('copy', { bubbles: true, cancelable: true });
            Object.defineProperty(e, 'clipboardData', {
              value: {
                setData() {
                  throw Error('Copy bypassed the captured asynchronous host route');
                },
              },
            });
            input.dispatchEvent(e);
            await started.promise;
          });
          assert.equal(
            document.querySelector('.spreadsheet-editor').dataset.homePending,
            'true',
            'Copy must await the actual host Promise'
          );
          assert.equal(copied.at(-1), expected);
          assert.equal(copied.at(-1).includes('Stored:'), false);
          await act(async () => completed.resolve());
          assert.equal(document.querySelector('.spreadsheet-editor').dataset.homePending, 'false');
        } finally {
          completed.resolve();
        }
      }
      assert.deepEqual(copied, ['2', '=A1+1']);
      assert.deepEqual(ui.legacy, []);
      assert.deepEqual(ui.strict, []);
    } finally {
      hostCopyText = async () => {};
      await ui.close();
    }
  }
);
check(
  'pending owned object caches are visibly stored even with no pending formulas and malformed counters cannot enable jobs',
  async () => {
    const c = controls({ policy: { ...policy, mode: 'manual', pendingFormulaCells: 0, pendingCacheObjects: 2 } }),
      ui = await mount({ calculationControls: c });
    try {
      await open();
      assert.equal(label('Pending object caches').dataset.pendingCacheObjects, '2');
      assert.match(label('Pending object caches').textContent, /Stored output is not current/);
      assert.equal(
        document.querySelector('[aria-label="Workbook calculation policy"]').dataset.pendingCacheObjects,
        '2'
      );
      assert.match(
        document.querySelector('[aria-label="Workbook calculation policy"]').textContent,
        /2 chart\/pivot caches pending/
      );
      for (const bad of [-1, NaN, 1.5, '2']) {
        await ui.render({ calculationControls: { ...c, policy: { ...c.policy, pendingCacheObjects: bad } } });
        assert.equal(command('formulas.calculate-now').disabled, true);
        assert.match(command('formulas.calculate-now').title, /settings are unavailable/);
      }
    } finally {
      await ui.close();
    }
  }
);
check(
  'a protected-source explicit full-calculation capability refuses ribbon and F9 while absent legacy capability stays supported',
  async () => {
    const calls = [],
      c = controls({
        policy: {
          ...policy,
          protectedSource: true,
          canRecalculate: false,
          reason: 'The imported calculation mode is protected.',
        },
        run: async (a) => {
          calls.push(a);
          throw Error('must not run');
        },
      }),
      ui = await mount({ calculationControls: c });
    try {
      assert.equal(command('formulas.calculate-now').disabled, true);
      assert.match(command('formulas.calculate-now').title, /imported calculation mode is protected/);
      await act(async () =>
        document
          .querySelector('[aria-label="Sheet1 A1"]')
          .dispatchEvent(new window.KeyboardEvent('keydown', { key: 'F9', bubbles: true, cancelable: true }))
      );
      assert.deepEqual(calls, []);
      for (const patch of [{ canRecalculate: 'no' }, { protectedSource: 'yes' }]) {
        await ui.render({ calculationControls: { ...c, policy: { ...policy, ...patch } } });
        assert.equal(command('formulas.calculate-now').disabled, true);
        assert.match(command('formulas.calculate-now').title, /settings are unavailable/);
      }
      await ui.render({ calculationControls: controls() });
      assert.equal(command('formulas.calculate-now').disabled, false);
    } finally {
      await ui.close();
    }
  }
);
check(
  'strict receipt refusal prevents false success for unrelated request/head/action, malformed rows, duplicate owners and hidden failures',
  () => {
    const action = { type: 'calculateAllOpen' },
      id = 'request',
      expected = { requestId: id, receipt: original, action },
      good = result(action, id);
    assert.equal(validateCalculationReceipt(good, expected).results[0].status, 'accepted');
    const mutants = [
      { ...good, requestId: 'other' },
      { ...good, initiator: { ...original, state: 'other' } },
      { ...good, action: { type: 'calculateSheet', sheet: 'Sheet2' } },
      { ...good, consistency: 'atomic' },
      { ...good, results: [null] },
      { ...good, results: [...good.results, ...good.results] },
      { ...good, results: [{ ...good.results[0], afterState: undefined }] },
      { ...good, results: [{ ...good.results[0], status: 'refused', afterState: undefined }] },
      { ...good, applicationMode: { mode: 'manual', revision: NaN } },
    ];
    for (const value of mutants) assert.throws(() => validateCalculationReceipt(value, expected));
  }
);
check(
  'controller captures original callback/opaque state, ignores unrelated advisory progress and stale pre-run microtasks; disposed work cannot publish',
  async () => {
    const held = defer(),
      calls = [],
      base = uiFrame(
        controls({
          run: async (action, id, progress) => {
            calls.push({ action, id });
            progress({ requestId: 'foreign', phase: 'calculating', completed: 128, total: 128, cancellable: false });
            await held.promise;
            return result(action, id);
          },
        })
      );
    let live = base;
    const controller = createCalculationJobController(() => live),
      key = calculationFrameKey(base);
    const pending = controller.run({ type: 'calculateAllOpen' }, key);
    await Promise.resolve();
    assert.equal(calls.length, 1);
    assert.equal(controller.getSnapshot().progress, null);
    live = {
      ...base,
      controls: {
        ...base.controls,
        run: async () => {
          throw Error('must not retarget');
        },
      },
    };
    controller.dispose();
    held.resolve();
    assert.equal((await pending).results[0].status, 'accepted');
    assert.equal(controller.getSnapshot().phase, 'idle');
    let next = uiFrame(),
      secondCalls = 0;
    next.controls.run = async (a, id) => {
      secondCalls++;
      return result(a, id);
    };
    const second = createCalculationJobController(() => next),
      p = second.run({ type: 'calculateAllOpen' }, calculationFrameKey(next));
    next = { ...next, scope: { ...next.scope, state: 'new UI token' } };
    await assert.rejects(p, CalculationJobRetired);
    assert.equal(secondCalls, 0);
    second.dispose();
  }
);
