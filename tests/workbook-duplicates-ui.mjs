import './support/scratch-dir.mjs';
/** Actual current Duplicate pane/Grid in React/JSDOM; controlled owner only. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
const component = fileURLToPath(new URL('../', import.meta.url)),
  output = await fs.mkdtemp(path.join(component, '.build/tests/duplicates-ui-')),
  checks = [];
const entry = path.join(output, 'entry.ts'),
  bundle = path.join(output, 'test.mjs');
await fs.writeFile(
  entry,
  `export {SpreadsheetGrid} from ${JSON.stringify(path.join(component, 'ui/SpreadsheetGrid.tsx'))};export {WorkbookDuplicates} from ${JSON.stringify(path.join(component, 'ui/WorkbookDuplicates.tsx'))};export {configureXlsxHost} from ${JSON.stringify(path.join(component, 'ui/host.tsx'))};export * from ${JSON.stringify(path.join(component, 'ui/duplicateSelection.ts'))};export * from ${JSON.stringify(path.join(component, 'ui/filterSelection.ts'))};`
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
const props = (node) => node[Object.keys(node).find((key) => key.startsWith('__reactProps'))],
  field = (name) =>
    [...document.querySelectorAll('[aria-label]')].find((node) => node.getAttribute('aria-label') === name) ?? null,
  button = (name) =>
    [...document.querySelectorAll('button')].find((n) => n.textContent === name && !n.closest('[hidden]'));
const defer = () => {
  let resolve, reject;
  const promise = new Promise((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
};
const flush = async (ms = 8) => act(async () => new Promise((resolve) => setTimeout(resolve, ms)));
async function click(n) {
  n = typeof n === 'string' ? button(n) : n;
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
async function go(range) {
  await change(field('Go to cell'), range);
  await act(async () =>
    field('Go to cell').dispatchEvent(
      new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })
    )
  );
  await flush();
}
const receipt = {
    documentId: 'native-doc',
    ownerId: 'native-owner',
    state: 'public-head',
    revision: 'native-revision',
  },
  fingerprint = 'a'.repeat(64),
  binding = {
    documentId: 'doc',
    state: 'head',
    sheet: 'Sheet1',
    selection: 'A1:E14',
    activationEpoch: 0,
    ownerIdentity: 'owner',
    selectedCell: 'A1',
    navigationGeneration: 0,
    modeGeneration: 0,
    specialSelectionEpoch: 0,
    mode: 'normal',
    specialSelection: null,
  };
const config = { sheet: 'Sheet1', range: 'A1:E14', hasHeaders: true, columns: ['A', 'B'] };
const page = (configuration = config, offset = 0, patch = {}) => ({
  schema: 1,
  ...receipt,
  ...configuration,
  fingerprint,
  totalRows: 13,
  keptRows: 10,
  removedRows: 3,
  comparison: 'exact-typed-values-v1',
  samples: [
    { row: 3, keptRow: 2 },
    { row: 11, keptRow: 2 },
    { row: 14, keptRow: 12 },
  ],
  offset,
  nextOffset: null,
  ...patch,
});
const context = (patch = {}) => ({
  binding: { ...binding },
  receipt: { ...receipt },
  serial: 1,
  caption: 'Sheet1!A1:E14',
  initialHeader: true,
  isCurrent: () => true,
  isRuntimeCurrent: () => true,
  disabledReason: () => undefined,
  read: async (c, o) => page(c, o),
  readHeaders: async () => ({ A1: { value: 'Region' }, B1: { value: 'Code' } }),
  accept: async () => {},
  ...patch,
});
async function mount(t, patch = {}) {
  const root = createRoot(document.getElementById('root')),
    accepted = [],
    legacy = [],
    reads = [],
    headers = [],
    data = {
      A1: { value: 'Region' },
      B1: { value: 'Code' },
      C1: { value: 'Amount' },
      D1: { value: 'Style' },
      E1: { value: 'Formula' },
      A2: { value: 'West' },
      B2: { value: 'A' },
    };
  const sheet = (name) => ({
    name,
    path: 'xl/worksheets/' + name + '.xml',
    state: 'visible',
    rows: 1200,
    columns: 200,
    cells: structuredClone(data),
    merges: [],
    widths: [],
    hiddenRows: [],
    rowHeights: [],
    freeze: { rows: 0, columns: 0 },
  });
  let input = {
    sheets: [sheet('Sheet1'), sheet('Other')],
    documentId: 'doc',
    ownerIdentity: 'owner',
    storageKey: 'head',
    active: true,
    readOnly: false,
    onEdit: (value) => legacy.push(value),
    onBatch: (value) => legacy.push(value),
    onCommandBatch: async (commands) => accepted.push(commands),
    duplicatePreviewReceipt: receipt,
    duplicatePreview: async (sheet, range, hasHeaders, columns, offset) => {
      reads.push({ sheet, range, hasHeaders, columns, offset });
      return page({ sheet, range, hasHeaders, columns }, offset);
    },
    readRange: async (sheet, range) => {
      headers.push({ sheet, range });
      return range === 'A1:E1' ? Object.fromEntries(Object.entries(data).filter(([ref]) => ref.endsWith('1'))) : {};
    },
    ...patch,
  };
  const render = async (patch) => {
    input = { ...input, ...patch };
    await act(async () => root.render(React.createElement(ui.SpreadsheetGrid, input)));
    await flush();
  };
  await render();
  await click([...document.querySelectorAll('[role="tab"]')].find((n) => n.textContent === 'Data'));
  let closed = false;
  const close = async () => {
    if (!closed) {
      closed = true;
      await act(async () => root.unmount());
    }
  };
  t.after(close);
  return { accepted, legacy, reads, headers, data, render, close, input: () => input };
}
async function open() {
  const opener = document.querySelector('[data-workbook-tool="duplicates"]');
  await act(async () => opener.focus());
  await click(opener);
  await flush();
}
async function selectAB() {
  await click('Clear selected columns');
  await change(field('Compare column A'), true);
  await change(field('Compare column B'), true);
}
check('bounded configs preserve ordered keys and physical-row receipts; closed malformed/foreign pages refuse', () => {
  const ctx = context();
  assert.deepEqual(ui.duplicateConfig(ctx, true, ['B', 'A']), { ...config, columns: ['B', 'A'] });
  assert.equal(ui.validateDuplicatePreview(page(), config, 0, receipt).samples[2].row, 14);
  for (const columns of [[], ['A', 'A'], ['F'], Array.from({ length: 17 }, (_, i) => String.fromCharCode(65 + i))])
    assert.throws(() => ui.duplicateConfig(ctx, true, columns));
  for (const range of ['A1:E200001', 'A0:E3', 'E14:A1']) assert.throws(() => ui.duplicateColumns(range));
  for (const bad of [
    null,
    [],
    page(config, 0, { extra: 1 }),
    page(config, 0, { ownerId: 'foreign' }),
    page(config, 0, { revision: 'foreign' }),
    page({ ...config, columns: ['B', 'A'] }),
    page(config, 0, { removedRows: 4 }),
    page(config, 0, { samples: [{ row: 3, keptRow: 3 }] }),
    page(config, 0, { nextOffset: 3 }),
    page(config, 0, { schema: 2 }),
  ])
    assert.throws(() => ui.validateDuplicatePreview(bad, config, 0, receipt));
});
check(
  'continuation validates exact counts/fingerprint/order and header reads stay at100 columns off viewport',
  async () => {
    const large = { sheet: 'Sheet1', range: 'A1:B202', hasHeaders: true, columns: ['A'] },
      first = page(large, 0, {
        totalRows: 201,
        keptRows: 2,
        removedRows: 199,
        samples: Array.from({ length: 100 }, (_, i) => ({ row: i + 3, keptRow: 2 })),
        nextOffset: 100,
      }),
      next = page(large, 100, {
        totalRows: 201,
        keptRows: 2,
        removedRows: 199,
        samples: Array.from({ length: 99 }, (_, i) => ({ row: i + 104, keptRow: 103 })),
      });
    assert.equal(ui.validateDuplicatePreview(next, large, 100, receipt, first).samples.length, 99);
    assert.throws(() =>
      ui.validateDuplicatePreview({ ...next, fingerprint: 'b'.repeat(64) }, large, 100, receipt, first)
    );
    assert.throws(() =>
      ui.validateDuplicatePreview(
        { ...next, samples: [{ row: 100, keptRow: 2 }, ...next.samples.slice(1)] },
        large,
        100,
        receipt,
        first
      )
    );
    const ranges = [],
      ctx = context({
        binding: { ...binding, selection: 'A1001:GR1003' },
        readHeaders: async (range) => {
          ranges.push(range);
          return range.startsWith('A1001')
            ? {
                A1001: { value: 'Region' },
                B1001: { formula: '1+1', calculation: { status: 'calculated', value: 2 } },
                C1001: { value: 'Old', calculation: { pending: true } },
              }
            : { CW1001: { value: 'Remote' } };
        },
      });
    assert.deepEqual(await ui.readDuplicateHeaders(ctx, 0), { A: 'Region' });
    assert.deepEqual(await ui.readDuplicateHeaders(ctx, 1), { CW: 'Remote' });
    assert.deepEqual(ranges, ['A1001:CV1001', 'CW1001:GR1001']);
  }
);
check(
  'actual one dock previews exact captured configuration and submits one acknowledged command without a renderer row map',
  async (t) => {
    const h = await mount(t);
    await go('A1:E14');
    await open();
    assert.equal(document.querySelectorAll('.sheet-task-pane').length, 1);
    assert.match(field('Remove duplicates').textContent, /Sheet1!A1:E14/);
    await selectAB();
    await click('Preview duplicates');
    assert.deepEqual(h.reads, [{ ...config, offset: 0 }]);
    assert.match(field('Duplicate preview counts').textContent, /13 data rows: 10 kept, 3 duplicates/);
    assert.deepEqual(
      [...field('Duplicate rows').querySelectorAll('tbody tr')].map((row) => row.textContent),
      ['32', '112', '1412']
    );
    assert.deepEqual(h.accepted, []);
    await click('Remove duplicates');
    assert.deepEqual(h.accepted, [[{ operation: 'removeDuplicates', ...config, expectedFingerprint: fingerprint }]]);
    assert.match(field('Remove duplicates').textContent, /3 duplicate rows removed/);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'column/header changes invalidate receipt and preserve selected order; wide selection pages100 headings and caps16 keys',
  async (t) => {
    const h = await mount(t);
    await go('A1:E14');
    await open();
    await selectAB();
    await click('Preview duplicates');
    await change(field('My data has headers'), false);
    assert.equal(field('Duplicate rows'), null);
    await change(field('My data has headers'), true);
    await change(field('Compare column A'), false);
    await change(field('Compare column A'), true);
    await click('Preview duplicates');
    assert.deepEqual(h.reads.at(-1).columns, ['B', 'A']);
    await click(field('Close workbook tools'));
    await go('A1001:GR1003');
    await open();
    assert.match(field('Remove duplicates').textContent, /0\/16/);
    assert.equal([...field('Remove duplicates').querySelectorAll('input[aria-label^="Compare column"]')].length, 100);
    assert(button('Select all columns').disabled);
    await click('More columns');
    assert(h.headers.some((item) => item.range === 'CW1001:GR1001'));
    for (const name of ['CW', 'CX', 'CY', 'CZ', 'DA', 'DB', 'DC', 'DD', 'DE', 'DF', 'DG', 'DH', 'DI', 'DJ', 'DK', 'DL'])
      await change(field('Compare column ' + name), true);
    assert(field('Compare column DM').disabled);
    assert.deepEqual(h.accepted, []);
  }
);
check(
  'held preview retires on head/owner/receipt-only/sheet/selection roundtrip/activation and never calls a replacement reader',
  async (t) => {
    for (const kind of ['head', 'owner', 'receipt', 'sheet', 'selection', 'active']) {
      const gate = defer(),
        calls = [],
        h = await mount(t, {
          duplicatePreview: async (...args) => {
            calls.push(args);
            return gate.promise;
          },
        });
      await go('A1:E14');
      await open();
      await selectAB();
      await click('Preview duplicates');
      assert.equal(calls.length, 1);
      await h.render({
        duplicatePreview: async () => {
          throw Error('Retargeted reader');
        },
      });
      if (kind === 'head') await h.render({ storageKey: 'next-head' });
      else if (kind === 'owner') await h.render({ ownerIdentity: 'other-owner' });
      else if (kind === 'receipt')
        await h.render({ duplicatePreviewReceipt: { ...receipt, revision: 'next-native-revision' } });
      else if (kind === 'sheet') {
        await click(button('Other'));
        await click(button('Sheet1'));
      } else if (kind === 'selection') {
        await go('B1:E14');
        await go('A1:E14');
      } else {
        await h.render({ active: false });
        await h.render({ active: true });
      }
      await act(async () => gate.resolve(page()));
      await flush();
      assert.equal(field('Duplicate rows'), null);
      assert.match(field('Remove duplicates').textContent, /changed. Reopen/);
      assert.deepEqual(h.accepted, []);
      await h.close();
    }
  }
);
check(
  'Cancel masks a pending request immediately, keeps one physical read until settle, and can then preview afresh',
  async (t) => {
    const gate = defer(),
      h = await mount(t, {
        duplicatePreview: async (...args) => {
          h.reads.push(args);
          return h.reads.length === 1 ? gate.promise : page();
        },
      });
    await go('A1:E14');
    await open();
    await selectAB();
    await click('Preview duplicates');
    await click('Cancel duplicate removal');
    assert(button('Preview duplicates').disabled);
    await act(async () => gate.resolve(page()));
    await flush();
    assert.equal(field('Duplicate rows'), null);
    await click('Preview duplicates');
    assert.equal(h.reads.length, 2);
    assert(field('Duplicate rows'));
    assert.deepEqual(h.accepted, []);
  }
);
check(
  'one held strict acceptance blocks duplicates and accepted head advance stays truthful without repeat submission',
  async (t) => {
    const gate = defer(),
      h = await mount(t, {
        onCommandBatch: (commands) => {
          h.accepted.push(commands);
          return gate.promise;
        },
      });
    await go('A1:E14');
    await open();
    await selectAB();
    await click('Preview duplicates');
    const submit = props(button('Remove duplicates')).onClick;
    await act(async () => {
      submit();
      submit();
    });
    assert.equal(h.accepted.length, 1);
    await h.render({
      storageKey: 'accepted-head',
      duplicatePreviewReceipt: { ...receipt, state: 'accepted-public-head', revision: 'accepted-native-revision' },
    });
    await act(async () => gate.resolve());
    await flush();
    assert.match(field('Remove duplicates').textContent, /3 duplicate rows removed/);
    await act(async () => submit());
    assert.equal(h.accepted.length, 1);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'late owner errors and retained open/remove handlers cannot affect a replacement owner draft or acceptance',
  async (t) => {
    const gate = defer(),
      h = await mount(t, {
        onCommandBatch: (commands) => {
          h.accepted.push(commands);
          return gate.promise;
        },
      });
    await go('A1:E14');
    const oldOpen = props(document.querySelector('[data-workbook-tool="duplicates"]')).onClick;
    await open();
    await selectAB();
    await click('Preview duplicates');
    const oldRemove = props(button('Remove duplicates')).onClick;
    await click('Remove duplicates');
    await h.render({ documentId: 'other-doc', ownerIdentity: 'other-owner' });
    const input = field('Formula or value');
    await act(async () => {
      input.focus();
      props(input).onChange({ target: { value: '=SUM(' } });
    });
    input.setSelectionRange(5, 5);
    await act(async () => {
      oldOpen();
      oldRemove();
      gate.reject(Error('Old owner failed'));
    });
    await flush();
    assert.equal(input.value, '=SUM(');
    assert.equal(input.selectionStart, 5);
    assert.equal(document.activeElement, input);
    assert(!document.body.textContent.includes('Old owner failed'));
    assert.equal(h.accepted.length, 1);
  }
);
check(
  'readonly native preview stays inspectable while removal is refused; missing receipt/passive/literal/inactive/busy are honest',
  async (t) => {
    const h = await mount(t, { readOnly: true });
    await go('A1:E14');
    await open();
    await selectAB();
    await click('Preview duplicates');
    assert(button('Remove duplicates').disabled);
    assert.equal(button('Remove duplicates').title, 'This workbook is read-only.');
    assert.deepEqual(h.accepted, []);
    await click(field('Close workbook tools'));
    for (const patch of [
      { readOnly: false, duplicatePreviewReceipt: undefined },
      { duplicatePreviewReceipt: receipt, active: false },
      { active: true, busy: true },
    ]) {
      await h.render(patch);
      assert(document.querySelector('[data-workbook-tool="duplicates"]').disabled);
    }
    await h.render({ busy: false, literal: true });
    assert.equal(document.querySelector('[data-workbook-tool="duplicates"]'), null);
    await h.render({ literal: false, passiveNotice: 'Passive inspector' });
    assert.equal(document.querySelector('[data-workbook-tool="duplicates"]'), null);
  }
);
check(
  'open/preview/options/cancel/close preserve the same-head formula draft/caret without an owner write',
  async (t) => {
    const h = await mount(t);
    await go('A1:E14');
    const input = field('Formula or value');
    await act(async () => {
      input.focus();
      props(input).onChange({ target: { value: '=SUM(' } });
    });
    input.setSelectionRange(5, 5);
    await click(document.querySelector('[data-workbook-tool="duplicates"]'));
    await selectAB();
    await click('Preview duplicates');
    await click('Cancel duplicate removal');
    await click(field('Close workbook tools'));
    await flush();
    assert.equal(input.value, '=SUM(');
    assert.equal(input.selectionStart, 5);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
    assert.equal(document.querySelectorAll('.sheet-task-pane').length, 1);
  }
);
check(
  'Remove filter uses only owned accepted revision for active/inactive filters and rejects stale/imported definitions',
  async (t) => {
    const revision = 'c'.repeat(64),
      definition = {
        range: 'A1:B6',
        column: 'A',
        values: [],
        owned: true,
        revision,
        canReapply: false,
        canClear: false,
        reason: 'No active filter criteria',
      };
    assert.equal(ui.acceptedFilterReason(definition, 'remove'), undefined);
    assert.deepEqual(ui.acceptedFilterCommand('Sheet1', definition, 'remove'), {
      operation: 'removeFilter',
      sheet: 'Sheet1',
      expectedFilterRevision: revision,
    });
    assert(ui.acceptedFilterReason({ ...definition, owned: false, revision: null }, 'remove'));
    const h = await mount(t);
    await h.render({
      sheets: h.input().sheets.map((sheet) => ({ ...sheet, filter: sheet.name === 'Sheet1' ? definition : undefined })),
    });
    await go('D20');
    const remove = button('Remove filter');
    assert.equal(remove.disabled, false);
    const old = props(remove).onClick;
    await click(remove);
    assert.deepEqual(h.accepted, [[{ operation: 'removeFilter', sheet: 'Sheet1', expectedFilterRevision: revision }]]);
    await h.render({
      sheets: h.input().sheets.map((sheet) => ({
        ...sheet,
        filter: sheet.name === 'Sheet1' ? { ...definition, revision: 'd'.repeat(64) } : undefined,
      })),
    });
    await act(async () => old());
    await flush();
    assert.equal(h.accepted.length, 1);
    await h.render({ sheets: h.input().sheets.map((sheet) => ({ ...sheet, filter: undefined })) });
    assert(button('Remove filter').disabled);
    assert.equal(button('Remove filter').title, 'This worksheet has no current filter.');
  }
);

check(
  'hiding the retained pane masks a held preview; reopening waits for the old physical read before a new request',
  async (t) => {
    const gate = defer(),
      h = await mount(t, {
        duplicatePreview: async (...args) => {
          h.reads.push(args);
          return h.reads.length === 1 ? gate.promise : page();
        },
      });
    await go('A1:E14');
    await open();
    await selectAB();
    await click('Preview duplicates');
    await click(field('Close workbook tools'));
    assert.equal(field('Duplicate rows'), null);
    await go('A1:E14');
    await open();
    assert(button('Preview duplicates').disabled);
    await act(async () => gate.resolve(page()));
    await flush();
    assert.equal(field('Duplicate rows'), null);
    await selectAB();
    await click('Preview duplicates');
    assert.equal(h.reads.length, 2);
    assert(field('Duplicate rows'));
    assert.deepEqual(h.accepted, []);
  }
);
check(
  'unrelated live grid drafts and IME cannot be opportunistically committed by opening the duplicate dock',
  async (t) => {
    const h = await mount(t);
    await go('A1:E14');
    const openHandler = props(document.querySelector('[data-workbook-tool="duplicates"]')).onClick;
    const cell = field('Sheet1 A1');
    await act(async () => {
      cell.focus();
      props(cell).onChange({ target: { value: 'Uncommitted' } });
    });
    await act(async () => openHandler());
    await flush();
    assert.equal(cell.value, 'Uncommitted');
    assert.equal(document.querySelector('.sheet-task-pane').hidden, true);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
    await act(async () =>
      cell.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }))
    );
    const formula = field('Formula or value');
    await act(async () => {
      formula.focus();
      props(formula).onChange({ target: { value: '=SUM(' } });
      props(formula).onCompositionStart({ currentTarget: formula });
    });
    await act(async () => openHandler());
    await flush();
    assert.equal(formula.value, '=SUM(');
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
    assert.equal(document.querySelector('.sheet-task-pane').hidden, true);
  }
);
check(
  'owned Filter Off remains available for special selections while source/definition protections stay exact',
  async (t) => {
    const revision = 'c'.repeat(64),
      filter = {
        range: 'A1:B6',
        column: 'A',
        values: ['West'],
        owned: true,
        revision,
        canReapply: true,
        canClear: true,
      };
    const h = await mount(t);
    await h.render({
      sheets: h.input().sheets.map((sheet) => ({ ...sheet, filter, cells: { ...sheet.cells, B2: { value: null } } })),
    });
    await go('A1:B3');
    await click([...document.querySelectorAll('[role="tab"]')].find((n) => n.textContent === 'Home'));
    await click(document.querySelector('[data-home-command="home.select-blanks"]'));
    await click('Preview matching cells');
    await click('Select matching cells');
    assert(document.querySelector('.sheet-special-status'));
    await click([...document.querySelectorAll('[role="tab"]')].find((n) => n.textContent === 'Data'));
    assert(button('Remove filter').disabled === false);
    assert(document.querySelector('[data-workbook-tool="duplicates"]').disabled);
    await click('Remove filter');
    assert.deepEqual(h.accepted, [[{ operation: 'removeFilter', sheet: 'Sheet1', expectedFilterRevision: revision }]]);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Page Layout retirement masks a pending duplicate preview and returning Normal does not revive its old receipt',
  async (t) => {
    const gate = defer(),
      h = await mount(t, { duplicatePreview: async () => gate.promise });
    await go('A1:E14');
    await open();
    await selectAB();
    await click('Preview duplicates');
    const settings = {
        area: 'A1:E14',
        paper: 'Letter',
        orientation: 'portrait',
        marginsPt: { top: 36, right: 36, bottom: 36, left: 36, header: 18, footer: 18 },
        scaling: { mode: 'scale', percent: 100 },
        repeatRows: null,
        repeatColumns: null,
        gridlines: false,
      },
      printed = {
        number: 1,
        rowIndices: [0],
        columnIndices: [0],
        cells: [{ ref: 'A1', text: 'Region', style: { fontSize: 11 }, bbox: [36, 36, 100, 20] }],
        objects: [],
      };
    await h.render({
      printControls: {
        documentId: 'print-owner',
        revision: 'head',
        definition: async () => ({ sheet: 'Sheet1', revision: null, settings, canEdit: true, protectedReasons: [] }),
        preview: async (_sheet, _settings, _id, geometryProfile) => ({
          geometryProfile,
          id: 'accepted-pages',
          state: 'head',
          units: 'pt',
          pageSizePt: [612, 792],
          scale: 1,
          sourceFingerprint: 'a'.repeat(64),
          revision: 'source',
          sheet: 'Sheet1',
          settings,
          pageCount: 1,
          pages: [printed],
        }),
        page: async () => printed,
        onBatch: async () => {},
        publish: async () => {},
        release: async () => {},
        cancel: async () => ({ granted: true }),
      },
    });
    await click([...document.querySelectorAll('[role="tab"]')].find((n) => n.textContent === 'View'));
    await click(document.querySelector('[data-ribbon-command="view.page-layout"]'));
    assert(document.querySelector('.sheet-page-layout'));
    await act(async () => gate.resolve(page()));
    await flush();
    assert.equal(field('Duplicate rows'), null);
    await click(document.querySelector('[data-ribbon-command="view.normal"]'));
    assert.match(field('Remove duplicates').textContent, /changed. Reopen/);
    assert.deepEqual(h.accepted, []);
  }
);
