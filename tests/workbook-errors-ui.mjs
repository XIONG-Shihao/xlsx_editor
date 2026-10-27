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
  folder = path.join(component, '.build/tests'),
  runs = await fs.mkdtemp(path.join(folder, 'errors-ui-')),
  entry = path.join(runs, 'entry.ts'),
  bundle = path.join(runs, 'test.mjs');
await fs.writeFile(
  entry,
  `export {WorkbookErrors} from ${JSON.stringify(path.join(component, 'ui/WorkbookErrors.tsx'))};export * from ${JSON.stringify(path.join(component, 'ui/errorChecking.ts'))};export * from ${JSON.stringify(path.join(component, 'ui/errorScanAdapter.ts'))};export {SpreadsheetGrid} from ${JSON.stringify(path.join(component, 'ui/SpreadsheetGrid.tsx'))};export {configureXlsxHost} from ${JSON.stringify(path.join(component, 'ui/host.tsx'))};`
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
    WorkbookErrors,
    createErrorCheckingController,
    createScopedErrorScanner,
    validateErrorScanPage,
    SpreadsheetGrid,
    configureXlsxHost,
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
const scope = () => ({
  documentId: 'doc-A',
  ownerIdentity: 'owner-A',
  state: 'opaque:ui:head-A',
  sheet: 'Sheet1',
  activationEpoch: 0,
  active: true,
});
const error = (code = '#N/A') => ({ value: code, dataType: 'error' });
const rows = () => [...document.querySelectorAll('[data-error-reference]')];
const button = (name) =>
  [...document.querySelectorAll('button')].find(
    (node) => !node.closest('[hidden]') && (node.getAttribute('aria-label') ?? node.textContent) === name
  );
async function click(node) {
  assert(node);
  assert.equal(node.disabled, false);
  await act(async () => node.click());
}
async function mount(t, extra = {}) {
  const root = createRoot(document.getElementById('root'));
  let props = {
    scope: scope(),
    scan: async (bound, after) => ({ scope: bound, after, cells: { A1: error() }, nextCursor: null }),
    goToIssue: () => {},
    ...extra,
  };
  const draw = () => React.createElement(WorkbookErrors, props);
  await act(async () => root.render(extra.strict ? React.createElement(React.StrictMode, null, draw()) : draw()));
  t.after(async () => act(async () => root.unmount()));
  return {
    render: async (patch) => {
      props = { ...props, ...patch };
      await act(async () => root.render(extra.strict ? React.createElement(React.StrictMode, null, draw()) : draw()));
    },
  };
}

test('manual typed-error scan excludes ordinary error-looking text and keeps formula statuses distinct without publishing stored caches', async (t) => {
  let scans = 0;
  const go = [],
    cells = {
      A1: { value: '#N/A', dataType: 'text' },
      A2: error(),
      A3: { value: null, formula: 'UNKNOWN()', calculation: { status: 'unresolved', error: 'Unknown function' } },
      A4: { value: null, formula: 'A4+1', calculation: { status: 'circular', error: '#CIRC!' } },
      A5: { value: 99, cachedValue: 99, formula: 'UNKNOWN()', calculation: { status: 'stored', value: 99 } },
      A6: { value: 2, formula: '1+1', calculation: { status: 'calculated', value: 2 } },
      A7: { value: null, formula: '1/0', calculation: { status: 'error', error: '#DIV/0!' } },
      A8: { value: '#N/A', formula: 'XLOOKUP(4,A1:A5,B1:B5)', calculation: { status: 'calculated', value: '#N/A' } },
    };
  await mount(t, {
    scan: async (bound, after) => {
      scans++;
      return { scope: bound, after, cells, nextCursor: null };
    },
    goToIssue: (ref) => go.push(ref),
  });
  assert.equal(scans, 0);
  await click(button('Scan worksheet'));
  assert.deepEqual(
    rows().map((row) => [row.dataset.errorReference, row.dataset.errorKind]),
    [
      ['A2', 'error'],
      ['A3', 'unresolved'],
      ['A4', 'circular'],
      ['A5', 'stored'],
      ['A7', 'error'],
    ]
  );
  assert(!document.querySelector('section[aria-label="Error Checking"]').textContent.includes('99'));
  assert.match(document.body.textContent, /Scan complete: 5 issues found/);
  await click(button('Next issue'));
  await click(button('Next issue'));
  assert.deepEqual(go, ['A2', 'A3']);
  await click(button('Go to issue Sheet1!A7'));
  assert.deepEqual(go, ['A2', 'A3', 'A7']);
});

test('100-issue action resumes exactly after last processed cell rather than discarding suffix of a larger native page', async (t) => {
  const calls = [],
    cells = Object.fromEntries(Array.from({ length: 135 }, (_, i) => ['B' + (i + 1), error()]));
  await mount(t, {
    scan: async (bound, after) => {
      calls.push(after);
      return {
        scope: bound,
        after,
        cells: Object.fromEntries(
          Object.entries(cells).filter(([ref]) => !after || Number(ref.slice(1)) > Number(after.slice(1)))
        ),
        nextCursor: null,
      };
    },
  });
  await click(button('Scan worksheet'));
  assert.equal(rows().length, 100);
  assert.deepEqual(calls, [null]);
  assert.match(document.body.textContent, /Scan incomplete/);
  assert.match(document.body.textContent, /resumes after B100/);
  await click(button('Continue scan'));
  assert.equal(rows().length, 35);
  assert.deepEqual(calls, [null, 'B100']);
  assert.equal(rows()[0].dataset.errorReference, 'B101');
  assert.equal(rows().at(-1).dataset.errorReference, 'B135');
  assert.match(document.body.textContent, /Scan complete: 135 issues found/);
});

test('64-page step is explicitly incomplete and an empty final issue page cannot erase cumulative errors', async (t) => {
  const calls = [];
  await mount(t, {
    scan: async (bound, after) => {
      const row = after ? Number(after.slice(1)) + 1 : 1;
      calls.push(row);
      return {
        scope: bound,
        after,
        cells: { ['A' + row]: row <= 2 ? error() : { value: row } },
        nextCursor: row === 72 ? null : 'A' + row,
      };
    },
  });
  await click(button('Scan worksheet'));
  assert.equal(calls.length, 64);
  assert.equal(rows().length, 2);
  assert.match(document.body.textContent, /step reached the scan limit/);
  assert.match(document.body.textContent, /Scan incomplete/);
  await click(button('Continue scan'));
  assert.equal(calls.length, 72);
  assert.equal(rows().length, 0);
  assert.match(document.body.textContent, /Scan complete: 2 issues found/);
  assert.match(document.body.textContent, /Earlier pages/);
  assert(!document.body.textContent.includes('no issues found'));
});

test('Cancel masks pending results, blocks duplicate action and never retries or overlaps an ignored-abort scan', async (t) => {
  const gate = deferred(),
    calls = [];
  let physical = 0,
    peak = 0;
  await mount(t, {
    scan: async (bound, after) => {
      calls.push(after);
      physical++;
      peak = Math.max(peak, physical);
      const cells = calls.length === 1 ? await gate.promise : {};
      physical--;
      return { scope: bound, after, cells, nextCursor: null };
    },
  });
  await click(button('Scan worksheet'));
  assert.equal(button('Scan worksheet').disabled, true);
  assert.equal(button('Restart scan').disabled, true);
  await click(button('Cancel scan'));
  assert.equal(rows().length, 0);
  assert.match(document.body.textContent, /Scan cancelled/);
  await act(async () => gate.resolve({ A1: error() }));
  assert.equal(calls.length, 1);
  assert.equal(rows().length, 0);
  await click(button('Scan worksheet'));
  assert.equal(calls.length, 2);
  assert.equal(peak, 1);
});

test('accepted head change immediately clears old issues and retires its pending reply; a new scope requires an explicit scan', async (t) => {
  const gate = deferred(),
    calls = [];
  const ui = await mount(t, {
    scan: async (bound, after) => {
      calls.push(bound.state);
      return {
        scope: bound,
        after,
        cells: calls.length === 2 ? await gate.promise : { A1: error() },
        nextCursor: null,
      };
    },
  });
  await click(button('Scan worksheet'));
  assert.equal(rows().length, 1);
  await click(button('Restart scan'));
  await ui.render({ scope: { ...scope(), state: 'opaque:ui:head-B' } });
  assert.equal(rows().length, 0);
  assert.equal(button('Scan worksheet').disabled, false);
  await act(async () => gate.resolve({ A2: error('#OLD!') }));
  assert.equal(rows().length, 0);
  assert.equal(calls.length, 2);
  assert(!document.body.textContent.includes('#OLD!'));
  await click(button('Scan worksheet'));
  assert.equal(calls.at(-1), 'opaque:ui:head-B');
});

test('hidden inactive worksheet, selected-sheet switch and runtime replacement mask results and reset cursor without navigation', async (t) => {
  const go = [];
  const ui = await mount(t, { goToIssue: (ref) => go.push(ref) });
  await click(button('Scan worksheet'));
  assert.equal(rows().length, 1);
  await ui.render({ scope: { ...scope(), active: false } });
  assert.equal(rows().length, 0);
  assert.equal(button('Scan worksheet').disabled, true);
  assert.match(document.body.textContent, /Paused/);
  await ui.render({ scope: { ...scope(), activationEpoch: 2 } });
  assert.equal(rows().length, 0);
  await click(button('Scan worksheet'));
  assert.equal(rows().length, 1);
  await ui.render({ scope: { ...scope(), activationEpoch: 2, sheet: 'Sheet2' } });
  assert.equal(rows().length, 0);
  await click(button('Scan worksheet'));
  assert.match(rows()[0].textContent, /Sheet2!A1/);
  await ui.render({ scope: { ...scope(), ownerIdentity: 'owner-reloaded' } });
  assert.equal(rows().length, 0);
  assert.deepEqual(go, []);
});

test('captured scan adapter preserves actual backend closure and opaque token, refusing before/after scope or missing receipt changes', async () => {
  const gate = deferred(),
    calls = [],
    bound = scope();
  let frame = {
    scope: bound,
    scan: async (sheet, after) => {
      calls.push({ nativeState: 'actual-journal-head', sheet, after });
      return gate.promise;
    },
  };
  const scanner = createScopedErrorScanner(() => frame),
    signal = new AbortController().signal;
  const pending = scanner(bound, null, signal);
  frame = {
    scope: bound,
    scan: async () => {
      throw Error('must not retarget');
    },
  };
  gate.resolve({ cells: { A1: error() }, nextCursor: null });
  assert.equal((await pending).cells.A1.dataType, 'error');
  assert.deepEqual(calls, [{ nativeState: 'actual-journal-head', sheet: 'Sheet1', after: undefined }]);
  for (const patch of [
    { state: 'new-head' },
    { ownerIdentity: 'new-runtime' },
    { sheet: 'Sheet2' },
    { activationEpoch: 1 },
    { active: false },
  ]) {
    const held = deferred();
    frame = { scope: bound, scan: async () => held.promise };
    const retired = scanner(bound, 'A1', signal);
    frame = { ...frame, scope: { ...bound, ...patch } };
    held.resolve({ cells: { A2: error() }, nextCursor: null });
    await assert.rejects(retired, /revision changed/);
  }
  frame = {
    scope: { ...bound, state: 'new-head' },
    scan: async () => {
      throw Error('must not request');
    },
  };
  await assert.rejects(scanner(bound, null, signal), /revision changed/);
  for (const invalid of [
    { cells: {} },
    { cells: null, nextCursor: null },
    { cells: { A1: error() }, nextCursor: 'XFE1' },
    { cells: { bad: error() }, nextCursor: null },
    ...['not a cell', null, undefined, []].map((cell) => ({ cells: { A1: cell }, nextCursor: null })),
  ])
    assert.throws(() => validateErrorScanPage(invalid), /invalid/);
});

test('provider failure is incomplete and cannot become a clean worksheet; accepted previous issue page survives cancellation', async (t) => {
  const gate = deferred();
  let count = 0;
  await mount(t, {
    scan: async (bound, after) => {
      count++;
      if (count === 1) return { scope: bound, after, cells: { A1: error() }, nextCursor: 'A1' };
      if (count === 2) throw Error('No reply');
      return { scope: bound, after, cells: await gate.promise, nextCursor: null };
    },
  });
  await click(button('Scan worksheet'));
  assert.match(document.body.textContent, /could not be scanned/);
  assert.match(document.body.textContent, /Scan incomplete/);
  assert(!document.body.textContent.includes('Scan complete'));
  // Failure publishes no partially checked page. A fresh controller preserves
  // only an actually accepted previous page when its next read is cancelled.
  let scan = 0;
  const controller = createErrorCheckingController(
    async (bound, after) => {
      scan++;
      if (scan === 1) return { scope: bound, after, cells: { B1: error() }, nextCursor: 'B1' };
      return { scope: bound, after, cells: await gate.promise, nextCursor: null };
    },
    { pages: 1 }
  );
  controller.setScope(scope());
  await controller.scanNext();
  assert.equal(controller.getSnapshot().issues.length, 1);
  const pending = controller.scanNext();
  await Promise.resolve();
  controller.cancel();
  assert.equal(controller.getSnapshot().issues[0].reference, 'B1');
  assert.equal(controller.getSnapshot().cursor, 'B1');
  gate.resolve({ B2: error() });
  assert.equal((await pending).outcome, 'retired');
  assert.equal(controller.getSnapshot().issues.length, 1);
  controller.dispose();
});

test('StrictMode rehearsal keeps a retained controller live; Restart is one explicit fresh action', async (t) => {
  let scans = 0;
  await mount(t, {
    strict: true,
    scan: async (bound, after) => {
      scans++;
      return { scope: bound, after, cells: { A1: error() }, nextCursor: null };
    },
  });
  await click(button('Scan worksheet'));
  assert.equal(rows().length, 1);
  await click(button('Restart scan'));
  assert.equal(scans, 2);
  assert.equal(rows().length, 1);
});

test('production readonly Grid has one real auditing pane and Go to issue navigates its actual offscreen cell without writes', async (t) => {
  const root = createRoot(document.getElementById('root')),
    calls = [],
    writes = [],
    sheet = {
      name: 'Sheet1',
      state: 'visible',
      path: 'xl/worksheets/sheet1.xml',
      rows: 100,
      columns: 12,
      cells: {},
      merges: [],
      widths: [],
      hiddenRows: [],
      rowHeights: [],
      freeze: { rows: 0, columns: 0 },
    };
  await act(async () =>
    root.render(
      React.createElement(SpreadsheetGrid, {
        sheets: [sheet],
        documentId: 'doc-A',
        ownerIdentity: 'owner-A',
        storageKey: 'opaque:ui:head-A',
        readOnly: true,
        onBatch: () => writes.push('write'),
        readRange: async () => ({}),
        scan: async (name, after) => {
          calls.push({ name, after });
          return { cells: { A66: error('#DIV/0!') }, nextCursor: null };
        },
      })
    )
  );
  t.after(async () => act(async () => root.unmount()));
  await click([...document.querySelectorAll('[role="tab"]')].find((node) => node.textContent === 'Formulas'));
  await click(button('Error Checking'));
  assert.equal(calls.length, 0);
  assert.equal(document.querySelectorAll('aside.sheet-task-pane:not([hidden])').length, 1);
  const groups = [
    ...document.querySelector('.sheet-ribbon-tabpanel:not([hidden])').querySelectorAll('.sheet-ribbon-group'),
  ].filter((group) => group.querySelector('small')?.textContent === 'Formula auditing');
  assert.equal(groups.length, 1);
  await click(button('Scan worksheet'));
  assert.deepEqual(calls, [{ name: 'Sheet1', after: undefined }]);
  assert.equal(document.querySelector('section[aria-label="Error Checking"]').dataset.errorState, 'opaque:ui:head-A');
  await click(button('Go to issue Sheet1!A66'));
  assert.equal(document.querySelector('[aria-label="Go to cell"]').placeholder, 'A66');
  assert.deepEqual(writes, []);
});

test('production Grid sheet change resets the old scan and non-XLSX or unscoped hosts expose no auditing command', async (t) => {
  const root = createRoot(document.getElementById('root')),
    base = {
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
    },
    sheets = [
      { ...base, name: 'Sheet1' },
      { ...base, name: 'Sheet2' },
    ],
    calls = [];
  let props = {
    sheets,
    documentId: 'doc-A',
    ownerIdentity: 'owner-A',
    storageKey: 'head-A',
    readOnly: true,
    readRange: async () => ({}),
    scan: async (name, after) => {
      calls.push(name);
      return { cells: { A1: error() }, nextCursor: null };
    },
  };
  const render = async () => act(async () => root.render(React.createElement(SpreadsheetGrid, props)));
  await render();
  t.after(async () => act(async () => root.unmount()));
  await click([...document.querySelectorAll('[role="tab"]')].find((node) => node.textContent === 'Formulas'));
  await click(button('Error Checking'));
  await click(button('Scan worksheet'));
  await click([...document.querySelectorAll('.sheet-tabs button')].find((node) => node.textContent === 'Sheet2'));
  assert.equal(rows().length, 0);
  await click(button('Scan worksheet'));
  assert.deepEqual(calls, ['Sheet1', 'Sheet2']);
  props = { ...props, literal: true };
  await render();
  assert.equal(button('Error Checking'), undefined);
  props = { ...props, literal: false, passiveNotice: 'Readonly passive' };
  await render();
  assert.equal(button('Error Checking'), undefined);
  props = { ...props, passiveNotice: undefined, scan: undefined };
  await render();
  assert.equal(button('Error Checking'), undefined);
});

test('controller malformed page scope, cursor and size refuse without advancing the accepted cursor or claiming completeness', async () => {
  let response;
  const controller = createErrorCheckingController(async (bound, after) => ({
    ...response,
    scope: response?.scope ?? bound,
    after: response?.after ?? after,
  }));
  controller.setScope(scope());
  const invalid = [
    { cells: { A1: error() }, nextCursor: 'A2' },
    { cells: { A1: error() }, nextCursor: null, scope: { ...scope(), sheet: 'Sheet2' } },
    { cells: Object.fromEntries(Array.from({ length: 2049 }, (_, i) => ['A' + (i + 1), error()])), nextCursor: null },
    { cells: { A1: error() }, nextCursor: null, after: 'B1' },
  ];
  for (response of invalid) {
    await assert.rejects(controller.scanNext());
    assert.equal(controller.getSnapshot().cursor, null);
    assert.equal(controller.getSnapshot().complete, false);
    assert.equal(controller.getSnapshot().incompleteReason, 'scan-failed');
  }
  controller.dispose();
});

test('central Page Layout immediately masks and pauses issue navigation; Normal return requires a fresh manual scan', async (t) => {
  const root = createRoot(document.getElementById('root')),
    writes = [],
    scans = [];
  const sheet = {
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
    },
    settings = {
      area: 'A1:B3',
      paper: 'Letter',
      orientation: 'portrait',
      marginsPt: { top: 36, right: 36, bottom: 36, left: 36, header: 18, footer: 18 },
      scaling: { mode: 'scale', percent: 100 },
      repeatRows: null,
      repeatColumns: null,
      gridlines: false,
    };
  const page = {
      number: 1,
      rowIndices: [0],
      columnIndices: [0],
      cells: [{ ref: 'A1', text: '#N/A', style: { fontSize: 11 }, bbox: [36, 36, 100, 30] }],
      objects: [],
    },
    controls = {
      documentId: 'runtime-A',
      revision: 'base',
      definition: async () => ({ sheet: 'Sheet1', revision: null, settings, canEdit: true, protectedReasons: [] }),
      preview: async (_sheet, requested) => ({
        id: 'accepted-page',
        state: 'base',
        units: 'pt',
        pageSizePt: [612, 792],
        scale: 1,
        sourceFingerprint: 'a'.repeat(64),
        revision: 'source-A',
        sheet: 'Sheet1',
        settings: requested,
        pageCount: 1,
        pages: [page],
      }),
      page: async () => page,
      onBatch: async () => writes.push('print'),
      publish: async () => {},
      release: async () => {},
      cancel: async () => ({ granted: true }),
    };
  await act(async () =>
    root.render(
      React.createElement(SpreadsheetGrid, {
        sheets: [sheet],
        documentId: 'doc-A',
        ownerIdentity: 'owner-A',
        storageKey: 'opaque:head-A',
        readRange: async () => ({}),
        scan: async () => {
          scans.push('scan');
          return { cells: { A1: error() }, nextCursor: null };
        },
        printControls: controls,
        onEdit: () => writes.push('edit'),
        onCommandBatch: async () => writes.push('command'),
      })
    )
  );
  t.after(async () => act(async () => root.unmount()));
  const tab = (name) => [...document.querySelectorAll('[role="tab"]')].find((node) => node.textContent === name);
  await click(tab('Formulas'));
  await click(button('Error Checking'));
  await click(button('Scan worksheet'));
  assert.equal(rows().length, 1);
  await click(tab('View'));
  await click(document.querySelector('[data-ribbon-command="view.page-layout"]'));
  assert(document.querySelector('.sheet-page-layout'));
  assert.equal(rows().length, 0);
  await click(tab('Formulas'));
  await click(button('Error Checking'));
  assert.equal(button('Scan worksheet').disabled, true);
  assert.match(document.querySelector('section[aria-label="Error Checking"]').textContent, /Paused/);
  await click(document.querySelector('.sheet-page-layout button'));
  assert.equal(button('Scan worksheet').disabled, false);
  assert.equal(rows().length, 0);
  assert.equal(scans.length, 1);
  await click(button('Scan worksheet'));
  assert.equal(rows().length, 1);
  assert.equal(scans.length, 2);
  assert.deepEqual(writes, []);
});

test('direct retained Go handler cannot revive after head/sheet/owner change or an away-and-back scope round trip', async (t) => {
  const go = [],
    ui = await mount(t, { goToIssue: (ref) => go.push(ref) });
  await click(button('Scan worksheet'));
  const node = button('Go to issue Sheet1!A1'),
    reactKey = Object.keys(node).find((key) => key.startsWith('__reactProps$')),
    retained = node[reactKey].onClick;
  assert.equal(typeof retained, 'function');
  for (const patch of [
    { state: 'other-head' },
    { sheet: 'Sheet2' },
    { ownerIdentity: 'replacement-owner' },
    { active: false, activationEpoch: 1 },
  ]) {
    await ui.render({ scope: { ...scope(), ...patch } });
    await act(async () => retained());
    assert.deepEqual(go, []);
    await ui.render({ scope: scope() });
    await act(async () => retained());
    assert.deepEqual(go, []);
  }
  await click(button('Scan worksheet'));
  await click(button('Go to issue Sheet1!A1'));
  assert.deepEqual(go, ['A1']);
});

test('Next issue remains useful on the accepted readonly issue page while a following scan is physically pending', async (t) => {
  const gate = deferred(),
    go = [];
  let calls = 0;
  await mount(t, {
    goToIssue: (ref) => go.push(ref),
    scan: async (bound, after) => {
      calls++;
      return {
        scope: bound,
        after,
        cells:
          calls === 1
            ? Object.fromEntries(Array.from({ length: 101 }, (_, i) => ['B' + (i + 1), error()]))
            : await gate.promise,
        nextCursor: null,
      };
    },
  });
  await click(button('Scan worksheet'));
  assert.equal(rows().length, 100);
  await click(button('Continue scan'));
  assert.equal(button('Next issue').disabled, false);
  await click(button('Next issue'));
  assert.deepEqual(go, ['B1']);
  await click(button('Cancel scan'));
  await act(async () => gate.resolve({ B101: error() }));
  assert.equal(rows().length, 100);
  assert.deepEqual(go, ['B1']);
});
