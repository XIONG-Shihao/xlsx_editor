import './support/scratch-dir.mjs';
/** Actual Grid/Ribbon/Review source in React/JSDOM; owner/provider/ack replies controlled. */
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
  output = await fs.mkdtemp(path.join(component, '.build/tests/review-grid-ui-')),
  checks = [];
const files = [
  'ui/SpreadsheetGrid.tsx',
  'ui/SpreadsheetRibbon.tsx',
  'ui/ribbonCommands.ts',
  'ui/WorkbookReview.tsx',
  'ui/workbookReviewController.ts',
  'ui/reviewScanAdapter.ts',
  'ui/reviewHostAdapter.ts',
  'core/workbook-review.mjs',
  'core/workbook-review.d.mts',
  'ui/styles/01.css',
  'tests/workbook-review-grid-ui.mjs',
  'tests/support/review-provider-availability.mjs',
];
const pins = () =>
  Promise.all(
    files.map(async (relative) => {
      const file = path.join(component, relative),
        bytes = await fs.readFile(file);
      return { relative, path: file, sha256: createHash('sha256').update(bytes).digest('hex'), bytes: bytes.length };
    })
  );
const before = await pins();
await fs.writeFile(path.join(output, 'source-before.json'), JSON.stringify(before, null, 2) + '\n');
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
      return 1200;
    },
  },
  clientHeight: {
    get() {
      return 650;
    },
  },
});
globalThis.ResizeObserver = class {
  constructor(fn) {
    this.fn = fn;
  }
  observe() {
    this.fn([{ contentRect: { width: 1200, height: 650 } }]);
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
const { reviewProviderAvailability } = await import('./support/review-provider-availability.mjs'),
  available = await reviewProviderAvailability();
const props = (node) => node[Object.keys(node).find((key) => key.startsWith('__reactProps$'))],
  field = (name) => document.querySelector(`[aria-label=${JSON.stringify(name)}]`),
  button = (name) =>
    [...document.querySelectorAll('button')].find(
      (node) => !node.closest('[hidden]') && (node.getAttribute('aria-label') ?? node.textContent) === name
    ),
  tab = (name) => [...document.querySelectorAll('[role="tab"]')].find((node) => node.textContent === name),
  command = (id) => document.querySelector('.sheet-ribbon-tabpanel:not([hidden]) [data-ribbon-command="' + id + '"]');
const flush = () => act(async () => new Promise((resolve) => setTimeout(resolve, 8))),
  defer = () => {
    let resolve, reject;
    const promise = new Promise((yes, no) => {
      resolve = yes;
      reject = no;
    });
    return { promise, resolve, reject };
  };
async function click(node) {
  node = typeof node === 'string' ? button(node) : node;
  assert(node);
  assert.equal(node.disabled, false);
  await act(async () => node.click());
  await flush();
}
async function go(reference) {
  const input = field('Go to cell');
  await act(async () => props(input).onChange({ target: { value: reference } }));
  await act(async () =>
    input.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
  );
  await flush();
}
const cell = (value) => ({
  value,
  dataType:
    typeof value === 'string'
      ? 'text'
      : typeof value === 'number'
        ? 'number'
        : typeof value === 'boolean'
          ? 'boolean'
          : 'blank',
});
const fixture = {
  Main: {
    A1: cell('helo wrld'),
    A2: cell(7),
    A3: {
      ...cell('helo'),
      formula: '"helo"',
      calculation: { status: 'calculated', value: 'helo', revision: 'revision-A' },
    },
  },
  Other: { A1: cell('hello') },
  Hidden: { A1: cell('wrld') },
};
const sheet = (name, data) => ({
  name,
  state: name === 'Hidden' ? 'hidden' : 'visible',
  path: 'xl/worksheets/' + name + '.xml',
  rows: 12,
  columns: 5,
  cells: structuredClone(data),
  merges: [],
  widths: [],
  hiddenRows: [],
  visibleRows: [],
  rowHeights: [],
  freeze: { rows: 0, columns: 0 },
});
function checked(request) {
  const dictionary = available.locales.find((item) => item.locale === request.locale).verified;
  return {
    schema: 1,
    requestId: request.requestId,
    locale: request.locale,
    provider: 'nspell',
    providerVersion: '2.1.5',
    providerId: available.providerId,
    status: 'completed',
    dictionaryId: dictionary.dictionaryId,
    dictionary,
    results: request.words.map((word) => ({
      word,
      correct: !['helo', 'wrld'].includes(word),
      suggestions: word === 'helo' ? ['hello'] : word === 'wrld' ? ['world'] : [],
    })),
  };
}
async function mount(t, patch = {}) {
  const root = createRoot(document.getElementById('root')),
    scans = [],
    reads = [],
    providerCalls = [],
    cancels = [],
    accepted = [],
    legacy = [],
    data = structuredClone(fixture);
  let input = {
      sheets: Object.entries(data).map(([name, cells]) => sheet(name, cells)),
      documentId: 'grant-A',
      ownerIdentity: 'runtime-A',
      storageKey: 'opaque-head-A',
      active: true,
      readOnly: false,
      nameCommands: true,
      ...patch,
    },
    receipt = { documentId: 'grant-A', ownerId: 'indexed-A', state: 'head-A', revision: 'revision-A' },
    scanOverride = null,
    checkOverride = null,
    readOverride = null,
    acceptOverride = null;
  const render = async (changes = {}) => {
    if (changes.receipt) {
      receipt = { ...receipt, ...changes.receipt };
      const { receipt: _discard, ...rest } = changes;
      changes = rest;
    }
    input = { ...input, ...changes };
    const captured = { ...receipt };
    const reviewControls = {
      receipt: captured,
      availability: async () => structuredClone(available),
      check: async (request) => {
        providerCalls.push({ receipt: captured, request });
        return checkOverride ? checkOverride(request, captured) : checked(request);
      },
      cancel: async (requestId) => {
        cancels.push({ receipt: captured, requestId });
      },
    };
    const scan = async (name, after, limit = 2048) => {
      scans.push({ receipt: captured, sheet: name, after, limit });
      if (scanOverride) return scanOverride(name, after, limit, captured);
      const all = Object.entries(data[name] ?? {}).filter(
          ([ref]) => after === undefined || Number(ref.slice(1)) > Number(after.slice(1))
        ),
        page = all.slice(0, limit);
      return { cells: Object.fromEntries(page), nextCursor: all.length > page.length ? page.at(-1)[0] : null };
    };
    const readRange = async (name, range) => {
      reads.push({ receipt: captured, sheet: name, range });
      if (readOverride) return readOverride(name, range, captured);
      if (/^[A-Z]+\d+$/.test(range)) return data[name]?.[range] ? { [range]: data[name][range] } : {};
      return structuredClone(data[name] ?? {});
    };
    const onCommandBatch = async (commands) => {
      accepted.push({ receipt: captured, commands });
      if (acceptOverride) await acceptOverride(commands, captured);
    };
    await act(async () =>
      root.render(
        React.createElement(ui.SpreadsheetGrid, {
          ...input,
          scan,
          readRange,
          reviewControls: input.withoutReview ? undefined : reviewControls,
          onEdit: (value) => legacy.push(value),
          onBatch: (value) => legacy.push(value),
          onCommandBatch: input.noWriter ? undefined : onCommandBatch,
        })
      )
    );
    await flush();
  };
  await render();
  t.after(async () => {
    await act(async () => root.unmount());
    await flush();
  });
  return {
    render,
    input: () => input,
    scans,
    reads,
    providerCalls,
    cancels,
    accepted,
    legacy,
    data,
    holdScan: (fn) => (scanOverride = fn),
    holdCheck: (fn) => (checkOverride = fn),
    holdRead: (fn) => (readOverride = fn),
    holdAccept: (fn) => (acceptOverride = fn),
  };
}
const open = async (mode) => {
  await click(tab('Review'));
  await click(command('review.' + mode));
};
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

check(
  'Review direct commands repeatedly select modes in one retained dock and close/reopen never starts an inspection',
  async (t) => {
    const h = await mount(t);
    await open('statistics');
    const review = field('Workbook Review'),
      dock = document.querySelector('aside.sheet-task-pane');
    assert.equal(review.dataset.reviewMode, 'statistics');
    assert.equal(document.querySelectorAll('aside.sheet-task-pane').length, 1);
    assert.equal(command('review.statistics').getAttribute('aria-pressed'), 'true');
    assert.equal(command('review.spelling').getAttribute('aria-pressed'), 'false');
    await click(command('review.statistics'));
    assert.equal(field('Workbook Review'), review);
    assert.equal(dock.hidden, false);
    await click(command('review.spelling'));
    assert.equal(review.dataset.reviewMode, 'spelling');
    await click(command('review.spelling'));
    assert.equal(dock.hidden, false);
    await click('Close workbook tools');
    assert(dock.hidden);
    await click(command('review.statistics'));
    assert.equal(field('Workbook Review'), review);
    assert.equal(dock.hidden, false);
    assert.equal(review.dataset.reviewMode, 'statistics');
    assert.deepEqual(h.scans, []);
    assert.deepEqual(h.providerCalls, []);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Native blur into Review tab/opener, pure scans, Cancel/close preserve the parked FormulaBar draft and caret',
  async (t) => {
    const h = await mount(t),
      formula = field('Formula or value');
    await act(async () => {
      formula.focus();
      props(formula).onChange({ target: { value: '=SUM(' } });
    });
    formula.setSelectionRange(5, 5);
    assert(tab('Review').closest('[data-formula-draft-zone]'));
    await act(async () => {
      props(formula).onBlur({ currentTarget: formula, relatedTarget: tab('Review') });
      tab('Review').focus();
    });
    await click(tab('Review'));
    assert(command('review.statistics').closest('[data-formula-draft-zone]'));
    await click(command('review.statistics'));
    await click('Scan workbook statistics');
    await click(command('review.spelling'));
    await click('Scan spelling');
    await click('Close workbook tools');
    assert.equal(formula.value, '=SUM(');
    assert.equal(formula.selectionStart, 5);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Grid passes bounded captured callbacks across all worksheets including hidden, then submits exactly one acknowledged text correction',
  async (t) => {
    const h = await mount(t);
    await open('statistics');
    await click('Scan workbook statistics');
    assert.equal(field('Accepted records').textContent, '5');
    assert.deepEqual(
      h.scans.map((value) => value.sheet),
      ['Main', 'Other', 'Hidden']
    );
    assert(h.scans.every((value) => value.limit === 256 && value.receipt.ownerId === 'indexed-A'));
    assert.equal(field('Formula cells').textContent, '1');
    assert.equal(field('Current formula results').textContent, '1');
    await click(command('review.spelling'));
    await click('Scan spelling');
    assert.equal(field('Spelling findings').children.length, 3);
    const first = field('Change word in Main!A1 0'),
      held = defer();
    h.holdAccept(() => held.promise);
    const submit = props(first).onClick;
    await act(async () => {
      submit();
      submit();
    });
    while (!h.accepted.length) await flush();
    assert.equal(h.accepted.length, 1);
    assert.deepEqual(h.accepted[0], {
      receipt: { documentId: 'grant-A', ownerId: 'indexed-A', state: 'head-A', revision: 'revision-A' },
      commands: [
        {
          type: 'setCells',
          sheet: 'Main',
          cells: [
            {
              cell: 'A1',
              value: 'hello wrld',
              expectedValue: 'helo wrld',
              expectedFormula: null,
              expectedRawValue: null,
            },
          ],
        },
      ],
    });
    assert(h.reads.some((value) => value.range === 'A1' && value.receipt.ownerId === 'indexed-A'));
    assert(first.disabled);
    await h.render({ storageKey: 'accepted-key', receipt: { state: 'accepted-head', revision: 'accepted-revision' } });
    await act(async () => held.resolve());
    await flush();
    assert.match(field('Workbook Review').textContent, /One Undo/);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Owner/head/receipt revision/selection/worksheet and inactive changes retire held Grid scans without stale rows or errors',
  async (t) => {
    const h = await mount(t);
    for (const kind of ['head', 'owner', 'revision', 'selection', 'sheet', 'inactive']) {
      await h.render({
        active: true,
        documentId: 'grant-A',
        ownerIdentity: 'runtime-' + kind,
        storageKey: 'opaque-' + kind,
        receipt: {
          documentId: 'grant-A',
          ownerId: 'indexed-' + kind,
          state: 'head-' + kind,
          revision: 'revision-' + kind,
        },
      });
      await go('A1');
      await open('statistics');
      const held = defer();
      h.holdScan(() => held.promise);
      await click('Scan workbook statistics');
      if (kind === 'head')
        await h.render({ storageKey: 'changed', receipt: { state: 'changed-head', revision: 'changed-revision' } });
      else if (kind === 'owner')
        await h.render({ ownerIdentity: 'replacement', receipt: { ownerId: 'replacement-native' } });
      else if (kind === 'revision') await h.render({ receipt: { revision: 'new-revision' } });
      else if (kind === 'selection') await go('B2');
      else if (kind === 'sheet') await click(button('Other'));
      else await h.render({ active: false });
      await act(async () => held.resolve({ cells: { A1: cell('obsolete') }, nextCursor: null }));
      await flush();
      assert.equal(field('Accepted records')?.textContent ?? '0', '0');
      assert(!field('Workbook Review').textContent.includes('obsolete'));
      assert.equal(field('Workbook Review').querySelector('[role="alert"]'), null);
      h.holdScan(null);
      await h.render({ active: true });
      await click('Close workbook tools');
    }
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Dock hiding cancels the captured provider token; a late original reply never populates a newly captured owner',
  async (t) => {
    const h = await mount(t),
      held = defer();
    h.holdCheck(() => held.promise);
    await open('spelling');
    await click('Scan spelling');
    assert.equal(h.providerCalls.length, 1);
    const original = h.providerCalls[0];
    await click('Close workbook tools');
    assert(
      h.cancels.some((value) => value.requestId === original.request.requestId && value.receipt.ownerId === 'indexed-A')
    );
    await h.render({
      ownerIdentity: 'replacement',
      receipt: { ownerId: 'replacement-native', state: 'replacement-head', revision: 'replacement-revision' },
    });
    await click(command('review.spelling'));
    await act(async () => held.resolve(checked(original.request)));
    await flush();
    assert.equal(field('Spelling findings').children.length, 0);
    assert.equal(field('Workbook Review').querySelector('[role="alert"]'), null);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Readonly permits inspection but disabled corrections; missing/literal/passive capabilities have no fake Review writer',
  async (t) => {
    const h = await mount(t, { readOnly: true });
    await open('spelling');
    await click('Scan spelling');
    assert(field('Change word in Main!A1 0').disabled);
    assert.match(field('Change word in Main!A1 0').title, /read-only/);
    await h.render({ readOnly: false, noWriter: true });
    await click(command('review.statistics'));
    await click('Scan workbook statistics');
    await click(command('review.spelling'));
    await click('Restart review');
    await click('Scan spelling');
    assert(field('Change word in Main!A1 0').disabled);
    for (const patch of [
      { withoutReview: true },
      { withoutReview: false, literal: true },
      { literal: false, passiveNotice: 'Stored original source' },
    ]) {
      await h.render(patch);
      assert.equal(command('review.spelling'), null);
    }
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
  }
);
check(
  'Switching to Page Layout retires retained Review and hidden Normal actions cannot inspect; returning Normal permits a fresh scan',
  async (t) => {
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
      page = { number: 1, rowIndices: [0, 1, 2], columnIndices: [0, 1], cells: [], objects: [] },
      h = await mount(t, {
        printControls: {
          documentId: 'grant-A',
          revision: 'head-A',
          definition: async () => ({ sheet: 'Main', revision: null, settings, canEdit: true, protectedReasons: [] }),
          preview: async (_sheet, _settings, _requestId, geometryProfile) => ({
            geometryProfile,
            id: 'page-job',
            state: 'head-A',
            units: 'pt',
            pageSizePt: [612, 792],
            scale: 1,
            sourceFingerprint: 'a'.repeat(64),
            revision: 'revision-A',
            sheet: 'Main',
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
    await open('statistics');
    await click('Scan workbook statistics');
    const stale = props(button('Restart review')).onClick;
    await click(tab('View'));
    await click(command('view.page-layout'));
    assert(document.querySelector('.sheet-page-layout'));
    assert(document.querySelector('.sheet-grid-viewport').hidden);
    assert.match(field('Review status').textContent, /paused/);
    const count = h.scans.length;
    await act(async () => stale());
    await flush();
    assert.equal(h.scans.length, count);
    await click(command('view.normal'));
    await click(tab('Review'));
    await click(command('review.statistics'));
    await click('Scan workbook statistics');
    assert.equal(h.scans.length, count + 3);
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(h.legacy, []);
  }
);
after(async () => {
  dom.window.close();
});
