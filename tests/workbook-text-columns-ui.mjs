import './support/scratch-dir.mjs';
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createHash } from 'node:crypto';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
import { rangeBounds, columnName } from '../core/workbook-commands.mjs';
const component = fileURLToPath(new URL('../', import.meta.url)),
  folder = path.join(component, '.build/tests'),
  runs = await fs.mkdtemp(path.join(folder, 'text-columns-ui-')),
  entry = path.join(runs, 'entry.ts'),
  bundle = path.join(runs, 'test.mjs');
await fs.writeFile(
  entry,
  `export * from ${JSON.stringify(path.join(component, 'ui/textColumns.ts'))};export * from ${JSON.stringify(path.join(component, 'ui/textColumnsProtection.ts'))};export {WorkbookTextColumns} from ${JSON.stringify(path.join(component, 'ui/WorkbookTextColumns.tsx'))};export {SpreadsheetGrid} from ${JSON.stringify(path.join(component, 'ui/SpreadsheetGrid.tsx'))};export {configureXlsxHost} from ${JSON.stringify(path.join(component, 'ui/host.tsx'))};`
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
const sha = (bytes) => createHash('sha256').update(bytes).digest('hex'),
  sourceNames = [
    'ui/textColumns.ts',
    'ui/textColumnsProtection.ts',
    'ui/WorkbookTextColumns.tsx',
    'tests/workbook-text-columns-ui.mjs',
    'ui/SpreadsheetGrid.tsx',
    'ui/SpreadsheetRibbon.tsx',
    'ui/ribbonCommands.ts',
    'ui/styles/01.css',
    'ui/index.ts',
  ],
  sourcePins = Object.fromEntries(
    await Promise.all(sourceNames.map(async (name) => [name, sha(await fs.readFile(path.join(component, name)))]))
  );
const productionUnchanged = async () =>
  Object.fromEntries(
    await Promise.all(
      Object.entries(sourcePins).map(async ([name, hash]) => [
        name,
        sha(await fs.readFile(path.join(component, name))) === hash,
      ])
    )
  );
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
globalThis.requestAnimationFrame = (fn) => setTimeout(fn, 0);
globalThis.cancelAnimationFrame = clearTimeout;
const { createRoot } = await import('react-dom/client'),
  {
    splitTextRecord,
    previewTextColumns,
    applyTextColumns,
    TextColumnsRetired,
    TextColumnsCancelled,
    textColumnProtection,
    WorkbookTextColumns,
    SpreadsheetGrid,
    configureXlsxHost,
  } = await import(pathToFileURL(bundle));
configureXlsxHost({
  api: { copyText: async () => {}, readClipboardText: async () => '' },
  Icon: () => React.createElement('svg'),
  useI18n: () => ({ locale: 'en' }),
});
after(async () => {
  assert(Object.values(await productionUnchanged()).every(Boolean), 'Production source changed during UI checks');
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
const text = (value) => ({ value, dataType: 'text' }),
  binding = () => ({
    documentId: 'doc-A',
    grantDocumentId: 'grant-A',
    ownerIdentity: 'owner-A',
    ownerKey: 'owner-A:base',
    state: 'opaque:base',
    sheet: 'Sheet1',
    selection: 'A1:A2',
    navigationGeneration: 0,
    activationEpoch: 0,
  }),
  availability = () => ({ active: true, readOnly: false, busy: false, literal: false, specialSelection: false });
function fixture(patch = {}) {
  let current = true;
  const data = {
      A1: text('00123,=1+2,x'),
      A2: text('00456,tail'),
      C1: { value: 0, dataType: 'number', rawValue: '0' },
      D1: { value: false, dataType: 'boolean', rawValue: '0' },
      E1: text(''),
    },
    reads = [],
    accepts = [];
  const context = {
    binding: binding(),
    isCurrent: () => current,
    isRuntimeCurrent: () => true,
    availability,
    read: async (range) => {
      reads.push(range);
      const bounds = rangeBounds(range);
      return Object.fromEntries(
        Object.entries(data).filter(([ref]) => {
          const p = rangeBounds(ref).first;
          return (
            p.row >= bounds.first.row &&
            p.row <= bounds.last.row &&
            p.col >= bounds.first.col &&
            p.col <= bounds.last.col
          );
        })
      );
    },
    inspect: async () => ({ reason: null }),
    accept: async (commands) => {
      accepts.push(commands);
    },
    ...patch,
  };
  return { context, data, reads, accepts, retire: () => (current = false) };
}
const options = { delimiter: ',', quotes: 'double', destination: 'C1' },
  button = (name) =>
    [...document.querySelectorAll('button')].find(
      (node) => !node.closest('[hidden]') && (node.getAttribute('aria-label') ?? node.textContent) === name
    ),
  field = (name) => document.querySelector('[aria-label="' + name + '"]');
async function click(node) {
  assert(node);
  assert.equal(node.disabled, false);
  await act(async () => node.click());
}
async function change(node, value) {
  assert(node);
  await act(async () => {
    if (node instanceof window.HTMLSelectElement) {
      node.value = value;
      node.dispatchEvent(new window.Event('change', { bubbles: true }));
    } else {
      Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(node, value);
      node.dispatchEvent(new window.Event('input', { bubbles: true }));
    }
  });
}
async function mount(t, context, extra = {}) {
  const root = createRoot(document.getElementById('root'));
  let props = { context, visible: true, ...extra };
  await act(async () => root.render(React.createElement(WorkbookTextColumns, props)));
  t.after(async () => act(async () => root.unmount()));
  return {
    render: async (patch) => {
      props = { ...props, ...patch };
      await act(async () => root.render(React.createElement(WorkbookTextColumns, props)));
    },
  };
}

test('closed quotes/delimiters preserve literal zero/formula text, spaces, empty fields and exact quoted linebreaks', () => {
  assert.deepEqual(splitTextRecord('00123,=SUM(A1:A2),false,', ',', 'double'), ['00123', '=SUM(A1:A2)', 'false', '']);
  assert.deepEqual(splitTextRecord('"a,b","say ""hi""", "x"', ',', 'none'), ['"a', 'b"', '"say ""hi"""', ' "x"']);
  assert.deepEqual(splitTextRecord('"a,b","say ""hi""","x\r\ny"', ',', 'double'), ['a,b', 'say "hi"', 'x\r\ny']);
  assert.deepEqual(splitTextRecord('a😀b😀', '😀', 'none'), ['a', 'b', '']);
  assert.deepEqual(splitTextRecord(' 01 | =1+2 ', '|', 'none'), [' 01 ', ' =1+2 ']);
  for (const [value, delimiter, quotes] of [
    ['"unfinished', ',', 'double'],
    ['"a" trailing', ',', 'double'],
    ['a\nb', ',', 'double'],
    ['a', '||', 'none'],
    ['a', '"', 'none'],
    ['a\u0000b', ',', 'none'],
  ])
    assert.throws(() => splitTextRecord(value, delimiter, quotes));
});

test('exact rectangular preview shows padding and false/zero/empty collisions; unapproved replacement never invokes owner', async () => {
  const h = fixture(),
    preview = await previewTextColumns(h.context, options);
  assert.equal(preview.destination, 'C1:E2');
  assert.equal(preview.collisions, 3);
  assert.deepEqual(
    preview.outputs.map((row) => [row.reference, row.value, row.padded]),
    [
      ['C1', '00123', false],
      ['D1', '=1+2', false],
      ['E1', 'x', false],
      ['C2', '00456', false],
      ['D2', 'tail', false],
      ['E2', '', true],
    ]
  );
  assert.deepEqual(h.reads, ['A1:A2', 'C1:E2']);
  await assert.rejects(applyTextColumns(h.context, preview), /occupied destination/);
  assert.equal(h.accepts.length, 0);
  const result = await applyTextColumns(h.context, preview, { replaceFingerprint: preview.fingerprint });
  assert.equal(result.positions, 6);
  assert.equal(h.accepts.length, 1);
  const command = h.accepts[0][0];
  assert.equal(command.type, 'setCells');
  assert.equal(h.accepts[0].length, 1);
  assert(command.cells.every((cell) => typeof cell.value === 'string' && !Object.hasOwn(cell, 'formula')));
  assert.equal(command.cells[1].expectedValue, false);
  assert.equal(command.cells[2].expectedValue, '');
});

test('source and destination are re-read/fingerprinted; mutable provider objects cannot silently alter the approved snapshot', async () => {
  const h = fixture(),
    preview = await previewTextColumns(h.context, options);
  h.data.C1.value = 9;
  h.data.C1.rawValue = '9';
  assert.equal(preview.outputs[0].previous.value, 0);
  await assert.rejects(applyTextColumns(h.context, preview, { replaceFingerprint: preview.fingerprint }), /changed/);
  assert.equal(h.accepts.length, 0);
  h.data.C1.value = 0;
  h.data.C1.rawValue = '0';
  h.data.A1.value = 'new,source';
  await assert.rejects(applyTextColumns(h.context, preview, { replaceFingerprint: preview.fingerprint }), /changed/);
  assert.equal(h.accepts.length, 0);
});

test('ordinary formula/spill/error/scenario and unsafe numeric intersections refuse before any owner write', async () => {
  for (const cell of [
    { value: 3, formula: '1+2' },
    { value: 'source', spill: { schema: 1, master: 'A1', range: 'A1:A2', role: 'follower' } },
    { value: 'source', formulaType: 'array' },
    { value: '#N/A', dataType: 'error' },
    { value: 'source', scenarioResult: { schema: 1 } },
    { value: 9007199254740992, dataType: 'number', rawValue: '9007199254740993' },
    { value: 1.23, dataType: 'number', rawValue: '1.2300000000000001' },
    { value: Infinity, dataType: 'number' },
  ]) {
    const h = fixture();
    h.data.C1 = cell;
    await assert.rejects(previewTextColumns(h.context, options));
    assert.equal(h.accepts.length, 0);
  }
  const h = fixture();
  h.data.A1 = { value: 123, dataType: 'number', rawValue: '123' };
  await assert.rejects(previewTextColumns(h.context, options), /literal text/);
});

test('metadata inspection refuses selected protected or unknown inventories but permits known unrelated tables/rules with formulas elsewhere', () => {
  const sheet = { name: 'Sheet1', merges: [], cells: { E1: { formula: 'SUM(C1:C2)' } } },
    definition = { sheet: 'Sheet1', revision: null, canEdit: true, protectedReasons: [], rules: [] },
    base = { sheet, tables: [], pivots: [], conditional: definition, validation: definition };
  assert.equal(textColumnProtection(base, 'A1:A2', 'C1:D2').reason, null);
  assert.equal(
    textColumnProtection(
      {
        ...base,
        tables: [{ name: 'Elsewhere', sheet: 'Sheet1', ref: 'Z10:AA12' }],
        validation: { ...definition, rules: [{ range: 'G10:G12' }] },
      },
      'A1:A2',
      'C1:D2'
    ).reason,
    null
  );
  const protectedPivot = {
    name: 'Imported',
    sheet: 'Sheet1',
    sourceSheet: 'Sheet1',
    outputRange: 'M10:N12',
    sourceRange: 'R1:S10',
    revision: null,
    canEdit: false,
  };
  const protectedRule = {
    ...definition,
    canEdit: false,
    protectedReasons: ['Imported custom validation'],
    protectedRanges: ['G10:G12'],
  };
  assert.equal(
    textColumnProtection(
      {
        ...base,
        pivots: [protectedPivot],
        validation: protectedRule,
        conditional: { ...definition, rules: [{ range: 'Z10:Z12' }], evaluationReasons: ['Unsafe duplicate values'] },
      },
      'A1:A2',
      'C1:D2'
    ).reason,
    null
  );
  assert.equal(
    textColumnProtection({ ...base, conditional: { ...protectedRule, protectedRanges: ['Z10:Z12'] } }, 'A1:A2', 'C1:D2')
      .reason,
    null
  );
  for (const patch of [
    { pivots: [{ ...protectedPivot, outputRange: 'C1:D3' }] },
    { pivots: [{ ...protectedPivot, sourceRange: 'A1:A2' }] },
    { validation: { ...protectedRule, protectedRanges: ['D2:D3'] } },
    { conditional: { ...protectedRule, protectedRanges: ['C1:C3'] } },
  ])
    assert(textColumnProtection({ ...base, ...patch }, 'A1:A2', 'C1:D2').reason);
  for (const patch of [
    { sheet: { ...sheet, merges: ['C1:D1'] } },
    { tables: [{ name: 'Here', sheet: 'Sheet1', ref: 'C1:D3' }] },
    { validation: { ...definition, rules: [{ range: 'C1:C3' }] } },
    { conditional: { ...definition, canEdit: false, protectedReasons: ['Opaque rule'] } },
    { validation: undefined },
    { pivots: [{ name: 'Opaque', sheet: null, revision: null, canEdit: false }] },
  ])
    assert(textColumnProtection({ ...base, ...patch }, 'A1:A2', 'C1:D2').reason);
});

test('1000x10 output is allowed with sequential <=1000-cell reads; overflow source/output worksheet cases refuse', async () => {
  const h = fixture();
  h.context.binding = { ...binding(), selection: 'A1:A1000' };
  for (let row = 1; row <= 1000; row++) h.data['A' + row] = text('1,2,3,4,5,6,7,8,9,0');
  delete h.data.C1;
  delete h.data.D1;
  delete h.data.E1;
  const preview = await previewTextColumns(h.context, options);
  assert.equal(preview.outputs.length, 10000);
  assert.equal(preview.destination, 'C1:L1000');
  assert(h.reads.every((range) => rangeBounds(range).size <= 1000));
  assert.equal(h.reads.length, 11);
  for (const selection of ['A1:A1001', 'A1:B1']) {
    const bad = fixture();
    bad.context.binding = { ...binding(), selection };
    await assert.rejects(previewTextColumns(bad.context, options));
    assert.equal(bad.accepts.length, 0);
  }
  await assert.rejects(previewTextColumns(h.context, { ...options, destination: 'XFD1' }), /limits/);
  h.data.A1 = text('1,2,3,4,5,6,7,8,9,0,extra');
  await assert.rejects(previewTextColumns(h.context, options), /10,000/);
});

test('accepted-head/selection/runtime retirement and early cancellation cannot enqueue any stale batch', async () => {
  const gate = deferred(),
    h = fixture({ read: () => gate.promise });
  const pending = previewTextColumns(h.context, options);
  h.retire();
  gate.resolve({ A1: text('x,y'), A2: text('z,w') });
  await assert.rejects(pending, TextColumnsRetired);
  assert.equal(h.accepts.length, 0);
  const normal = fixture(),
    preview = await previewTextColumns(normal.context, options),
    abort = new AbortController();
  await assert.rejects(
    applyTextColumns(normal.context, preview, {
      replaceFingerprint: preview.fingerprint,
      signal: abort.signal,
      beforeAcceptance: () => abort.abort(),
    }),
    TextColumnsCancelled
  );
  assert.equal(normal.accepts.length, 0);
  normal.context.availability = () => ({ ...availability(), readOnly: true });
  await assert.rejects(
    applyTextColumns(normal.context, preview, { replaceFingerprint: preview.fingerprint }),
    /read-only/
  );
  normal.context.availability = () => ({ ...availability(), specialSelection: true });
  await assert.rejects(
    applyTextColumns(normal.context, preview, { replaceFingerprint: preview.fingerprint }),
    /regular cell/
  );
});

test('pane exposes exact values/replacements and explicit action, accepted head advance stays success without resubmitting old draft', async (t) => {
  const h = fixture();
  let current = true;
  h.context.isCurrent = () => current;
  h.context.accept = async (commands) => {
    h.accepts.push(commands);
    current = false;
  };
  await mount(t, h.context);
  await change(field('Delimiter'), 'comma');
  await change(field('Destination start'), 'C1');
  await click(button('Preview split'));
  assert.equal(field('Output value C1').textContent, '"00123"');
  assert.equal(field('Output value D1').textContent, '"=1+2"');
  assert.equal(field('Output value E2').textContent, '""');
  assert.match(field('Destination replacements').textContent, /C1 0/);
  assert.match(field('Destination replacements').textContent, /D1 FALSE/);
  await click(button('Split and replace 3 occupied cells'));
  assert.equal(h.accepts.length, 1);
  assert.match(document.body.textContent, /Split accepted/);
  assert.equal(button('Preview split').disabled, true);
  assert.equal(button('Split and replace 3 occupied cells'), undefined);
});

test('Cancel preview waits for the ignored physical read to retire and never adopts its late table', async (t) => {
  const gate = deferred(),
    h = fixture({ read: () => gate.promise });
  await mount(t, h.context);
  await change(field('Delimiter'), 'comma');
  await click(button('Preview split'));
  await click(button('Cancel split'));
  assert.equal(button('Preview split').disabled, true);
  assert.match(document.body.textContent, /Stopping/);
  await act(async () => gate.resolve({ A1: text('x,y'), A2: text('z,w') }));
  assert.equal(field('Split preview'), null);
  assert.equal(h.accepts.length, 0);
  assert.equal(button('Preview split').disabled, false);
});

test('owner refusal remains refusal; hidden late accepted batch reports original outcome without foreign target content', async (t) => {
  const h = fixture();
  h.context.accept = async () => {
    throw Error('Owner refused this output');
  };
  const outcomes = [],
    ui = await mount(t, h.context, { onOutcome: (result) => outcomes.push(result) });
  await change(field('Delimiter'), 'comma');
  await change(field('Destination start'), 'C1');
  await click(button('Preview split'));
  await click(button('Split and replace 3 occupied cells'));
  assert.match(document.querySelector('[role="alert"]').textContent, /Owner refused/);
  assert.equal(outcomes.length, 0);
  const gate = deferred();
  h.context.accept = async () => gate.promise;
  await click(button('Split and replace 3 occupied cells'));
  await ui.render({ visible: false });
  await act(async () => gate.resolve());
  assert.equal(outcomes.length, 1);
  assert.equal(outcomes[0].outcome, 'accepted');
  assert.equal(outcomes[0].destination, 'C1:E2');
});

test('current stale-head pane masks old preview immediately and cannot accept or read a new implicit target', async (t) => {
  const h = fixture(),
    ui = await mount(t, h.context);
  await change(field('Delimiter'), 'comma');
  await change(field('Destination start'), 'C1');
  await click(button('Preview split'));
  assert(field('Split preview'));
  const calls = h.reads.length;
  h.retire();
  await ui.render({ context: h.context });
  assert.equal(field('Split preview'), null);
  assert.match(document.body.textContent, /changed/);
  assert.equal(button('Preview split').disabled, true);
  assert.equal(h.reads.length, calls);
  assert.equal(h.accepts.length, 0);
});

test('current production Grid launches one Data tool workflow and uses captured read/inventory/strict callback without a second pane', async (t) => {
  const root = createRoot(document.getElementById('root')),
    calls = [],
    sheet = {
      name: 'Sheet1',
      state: 'visible',
      path: 'xl/worksheets/sheet1.xml',
      rows: 64,
      columns: 12,
      cells: { A1: text('00123,=1+2') },
      merges: [],
      widths: [],
      hiddenRows: [],
      rowHeights: [],
      freeze: { rows: 0, columns: 0 },
    },
    definition = { sheet: 'Sheet1', revision: null, canEdit: true, protectedReasons: [], rules: [] };
  await act(async () =>
    root.render(
      React.createElement(SpreadsheetGrid, {
        sheets: [sheet],
        documentId: 'doc-A',
        ownerIdentity: 'owner-A',
        storageKey: 'owner-A:base',
        nameCommands: true,
        readRange: async (_name, range) => {
          calls.push(['read', range]);
          return range.startsWith('A1') ? { A1: text('00123,=1+2') } : {};
        },
        onCommandBatch: async (commands) => calls.push(['accepted', commands]),
        conditionalControls: {
          documentId: 'owner-A',
          revision: 'base',
          inventory: async () => definition,
          onBatch: async () => {},
        },
        validationControls: {
          documentId: 'owner-A',
          revision: 'base',
          inventory: async () => definition,
          onBatch: async () => {},
        },
      })
    )
  );
  t.after(async () => act(async () => root.unmount()));
  await click([...document.querySelectorAll('[role="tab"]')].find((node) => node.textContent === 'Data'));
  await click(button('Text to Columns'));
  assert.equal(document.querySelectorAll('aside.sheet-task-pane:not([hidden])').length, 1);
  await change(field('Delimiter'), 'comma');
  await change(field('Destination start'), 'C1');
  await click(button('Preview split'));
  assert.equal(field('Output value C1').textContent, '"00123"');
  await click(button('Split into C1:D1'));
  const accepted = calls.filter((item) => item[0] === 'accepted');
  assert.equal(accepted.length, 1);
  assert.deepEqual(
    accepted[0][1][0].cells.map((cell) => cell.value),
    ['00123', '=1+2']
  );
});

test('strict rectangle receipts refuse missing/array/null/foreign records without false clean preview', async () => {
  for (const reply of [
    null,
    [],
    { A1: null },
    { A1: [] },
    { A1: 'x' },
    { A1: undefined },
    { Z9: text('x') },
    { a1: text('x') },
  ]) {
    const h = fixture({ read: async () => reply });
    await assert.rejects(previewTextColumns(h.context, options));
    assert.equal(h.accepts.length, 0);
  }
});

test('preview and final guarded command enforce byte budgets independently', async () => {
  const huge = fixture();
  huge.context.binding = { ...binding(), selection: 'A1:A40' };
  for (let row = 1; row <= 40; row++) huge.data['A' + row] = text('x'.repeat(32767));
  await assert.rejects(previewTextColumns(huge.context, { ...options, quotes: 'none' }), /preview exceeds/);
  assert.equal(huge.accepts.length, 0);
  const large = fixture();
  large.context.binding = { ...binding(), selection: 'A1:A1000' };
  for (let row = 1; row <= 1000; row++) large.data['A' + row] = text('1,2,3,4,5,6,7,8,9,0');
  delete large.data.C1;
  delete large.data.D1;
  delete large.data.E1;
  const ordinaryRead = large.context.read;
  large.context.read = async (range) => {
    const record = await ordinaryRead(range);
    if (!range.startsWith('A1:')) {
      const area = rangeBounds(range);
      for (let row = area.first.row; row <= area.last.row; row++)
        for (let col = area.first.col; col <= area.last.col; col++)
          record[columnName(col) + (row + 1)] = text('x'.repeat(40));
    }
    return record;
  };
  await assert.rejects(previewTextColumns(large.context, options), /guarded split command exceeds/);
  assert.equal(large.accepts.length, 0);
});

test('hidden old read keeps one physical preflight gate through a runtime replacement and cannot adopt foreign content', async (t) => {
  const gate = deferred(),
    h = fixture({ read: () => gate.promise }),
    ui = await mount(t, h.context);
  await change(field('Delimiter'), 'comma');
  await click(button('Preview split'));
  await ui.render({ visible: false });
  const replacement = fixture();
  replacement.context.binding = {
    ...binding(),
    documentId: 'doc-B',
    grantDocumentId: 'grant-B',
    ownerIdentity: 'owner-B',
    ownerKey: 'owner-B:base',
  };
  await ui.render({ visible: true, context: replacement.context });
  assert.equal(button('Preview split').disabled, true);
  assert.equal(replacement.reads.length, 0);
  assert.equal(field('Split preview'), null);
  await act(async () => gate.resolve({ A1: text('old,read'), A2: text('old,read') }));
  assert.equal(button('Preview split').disabled, false);
  assert.equal(field('Split preview'), null);
  await click(button('Preview split'));
  assert(field('Split preview'));
  assert.equal(replacement.reads.length, 2);
  assert.equal(h.accepts.length, 0);
  assert.equal(replacement.accepts.length, 0);
});

async function go(range) {
  await change(field('Go to cell'), range);
  await act(async () =>
    field('Go to cell').dispatchEvent(
      new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })
    )
  );
}
async function mountGrid(t, extra = {}) {
  window.localStorage.clear();
  const root = createRoot(document.getElementById('root')),
    sheets = [
      {
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
      {
        name: 'Other',
        state: 'visible',
        path: 'xl/worksheets/sheet2.xml',
        rows: 64,
        columns: 12,
        cells: {},
        merges: [],
        widths: [],
        hiddenRows: [],
        rowHeights: [],
        freeze: { rows: 0, columns: 0 },
      },
    ],
    definition = { sheet: 'Sheet1', revision: null, canEdit: true, protectedReasons: [], rules: [] },
    writes = [],
    data = { A1: text('00123,=1+2'), A2: text('00456,tail') };
  const props = {
    sheets,
    documentId: 'doc-A',
    ownerIdentity: 'owner-A',
    storageKey: 'owner-A:base',
    nameCommands: true,
    onEdit: () => {},
    readRange: async (name, range) => {
      const area = rangeBounds(range);
      return name === 'Sheet1'
        ? Object.fromEntries(
            Object.entries(data).filter(([ref]) => {
              const p = rangeBounds(ref).first;
              return (
                p.row >= area.first.row && p.row <= area.last.row && p.col >= area.first.col && p.col <= area.last.col
              );
            })
          )
        : {};
    },
    onCommandBatch: async (commands) => writes.push(commands),
    conditionalControls: {
      documentId: 'owner-A',
      revision: 'base',
      inventory: async () => definition,
      onBatch: async () => {},
    },
    validationControls: {
      documentId: 'owner-A',
      revision: 'base',
      inventory: async () => definition,
      onBatch: async () => {},
    },
    ...extra,
  };
  await act(async () => root.render(React.createElement(SpreadsheetGrid, props)));
  t.after(async () => act(async () => root.unmount()));
  return {
    writes,
    render: async (patch) => act(async () => root.render(React.createElement(SpreadsheetGrid, { ...props, ...patch }))),
    open: async () => {
      await click([...document.querySelectorAll('[role="tab"]')].find((node) => node.textContent === 'Data'));
      await click(button('Text to Columns'));
    },
  };
}

test('current production Grid retires pending split for head, selection/sheet round trips, activation and same-range special selection', async (t) => {
  for (const reason of ['head', 'selection', 'worksheet', 'activation', 'special'])
    await t.test(reason, async (child) => {
      const gate = deferred(),
        writes = [];
      let exactReads = 0;
      const ui = await mountGrid(child, {
        onCommandBatch: async (commands) => writes.push(commands),
        readRange: async (name, range) =>
          name === 'Sheet1' && range === 'A1:A2' && ++exactReads === 1 ? gate.promise : {},
      });
      await go('A1:A2');
      await ui.open();
      await change(field('Delimiter'), 'comma');
      await change(field('Destination start'), 'C1');
      await click(button('Preview split'));
      if (reason === 'head') await ui.render({ storageKey: 'owner-A:next' });
      if (reason === 'selection') {
        await go('C5');
        await go('A1:A2');
      }
      if (reason === 'worksheet') {
        await click([...document.querySelectorAll('.sheet-tabs button')].find((node) => node.textContent === 'Other'));
        await click([...document.querySelectorAll('.sheet-tabs button')].find((node) => node.textContent === 'Sheet1'));
        await go('A1:A2');
      }
      if (reason === 'activation') {
        await ui.render({ active: false });
        await ui.render({ active: true });
      }
      if (reason === 'special') {
        await click([...document.querySelectorAll('[role="tab"]')].find((node) => node.textContent === 'Home'));
        await click(button('Editing'));
        const menu = [...document.querySelectorAll('details')].find(
          (node) => !node.closest('[hidden]') && node.textContent.includes('Find & Select')
        );
        assert(menu);
        menu.open = true;
        await click(button('Select blanks'));
        await click(button('Preview matching cells'));
        await act(async () => new Promise((resolve) => setTimeout(resolve, 20)));
        await click(button('Select matching cells'));
        await act(async () => new Promise((resolve) => setTimeout(resolve, 20)));
        assert(document.querySelector('.sheet-special-status'));
        assert(exactReads >= 2, 'Special reads the original range after the physical split read settles');
      }
      await act(async () => gate.resolve({ A1: text('old,source'), A2: text('old,source') }));
      assert.equal(field('Split preview'), null);
      assert.equal(writes.length, 0);
      assert.match(document.body.textContent, /Reopen Text to Columns/);
    });
});

test('current production Grid preserves the originally captured strict callback during a harmless render', async (t) => {
  const gate = deferred(),
    original = [],
    replacement = [];
  let reads = 0;
  const ui = await mountGrid(t, {
    onCommandBatch: async (commands) => original.push(commands),
    readRange: async (name, range) =>
      name === 'Sheet1' && range === 'A1:A2' && ++reads === 1
        ? gate.promise
        : name === 'Sheet1' && range === 'A1:A2'
          ? { A1: text('00123,=1+2'), A2: text('00456,tail') }
          : {},
  });
  await go('A1:A2');
  await ui.open();
  await change(field('Delimiter'), 'comma');
  await change(field('Destination start'), 'C1');
  await click(button('Preview split'));
  await ui.render({ onCommandBatch: async (commands) => replacement.push(commands) });
  await act(async () => gate.resolve({ A1: text('00123,=1+2'), A2: text('00456,tail') }));
  await click(button('Split into C1:D2'));
  assert.equal(original.length, 1);
  assert.equal(replacement.length, 0);
});
