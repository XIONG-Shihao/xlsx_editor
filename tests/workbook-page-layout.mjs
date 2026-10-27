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
  prefix = path.join(component, '.build/tests/page-layout-component-');
await fs.mkdir(path.dirname(prefix), { recursive: true });
const runs = await fs.mkdtemp(prefix),
  entry = path.join(runs, 'entry.ts'),
  bundle = path.join(runs, 'test.mjs');
await fs.writeFile(
  entry,
  `export {useAcceptedPrintPreview,isPrintPreviewRetired} from ${JSON.stringify(path.join(component, 'ui/useAcceptedPrintPreview.ts'))};export {WorkbookPageLayout} from ${JSON.stringify(path.join(component, 'ui/WorkbookPageLayout.tsx'))};export {WorkbookPrint} from ${JSON.stringify(path.join(component, 'ui/WorkbookPrint.tsx'))};export {SpreadsheetGrid} from ${JSON.stringify(path.join(component, 'ui/SpreadsheetGrid.tsx'))};export {configureXlsxHost} from ${JSON.stringify(path.join(component, 'ui/host.tsx'))};`
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
globalThis.requestAnimationFrame = () => 1;
globalThis.cancelAnimationFrame = () => {};
const { createRoot } = await import('react-dom/client'),
  {
    useAcceptedPrintPreview,
    isPrintPreviewRetired,
    WorkbookPrint,
    WorkbookPageLayout,
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
const settings = {
  area: 'A1:B3',
  paper: 'Letter',
  orientation: 'portrait',
  marginsPt: { top: 36, right: 36, bottom: 36, left: 36, header: 18, footer: 18 },
  scaling: { mode: 'scale', percent: 100 },
  repeatRows: null,
  repeatColumns: null,
  gridlines: false,
};
const page = (number) => ({
  number,
  rowIndices: number === 1 ? [0, 1] : [2],
  columnIndices: [0, 1],
  cells: (number === 1 ? [1, 2] : [3]).flatMap((row, r) =>
    ['A', 'B'].map((col, c) => ({
      ref: col + row,
      text: col + row,
      style: { fontSize: 11 },
      bbox: [36 + c * 100, 36 + r * 350, 100, 350],
    }))
  ),
  objects: [],
});
const job = (id = 'job-A', state = 'base') => ({
  id,
  state,
  units: 'pt',
  pageSizePt: [612, 792],
  scale: 1,
  sourceFingerprint: 'a'.repeat(64),
  revision: 'source-A',
  sheet: 'Sheet1',
  settings,
  pageCount: 2,
  pages: [page(1)],
});
const scope = () => ({
  documentId: 'runtime-A',
  grantDocumentId: 'durable-A',
  state: 'base',
  sheet: 'Sheet1',
  ownerKey: 'worker-A:base',
  activationEpoch: 0,
  active: true,
});
let api,
  seen = [];
function Probe({ scope: bound, callbacks, showPane = false, showPage = false }) {
  api = useAcceptedPrintPreview({ scope: bound, callbacks });
  seen.push({
    key: bound.ownerKey,
    active: bound.active,
    job: api.job?.id,
    page: api.page?.number,
    status: api.status,
    error: api.error,
  });
  return React.createElement(
    'div',
    null,
    showPane &&
      React.createElement(WorkbookPrint, {
        documentId: bound.documentId,
        revision: bound.state,
        sheet: bound.sheet,
        selection: settings.area,
        definition: async () => ({ sheet: 'Sheet1', revision: null, settings, canEdit: true, protectedReasons: [] }),
        onBatch: async () => {
          throw Error('No writes');
        },
        preview: callbacks.preview,
        page: callbacks.page,
        publish: callbacks.publish ?? (async () => {}),
        release: callbacks.release,
        cancel: callbacks.cancel,
        previewSession: api,
      }),
    showPage && React.createElement(WorkbookPageLayout, { preview: api, onNormal: () => api.abandonPreparation() })
  );
}
async function mount(t, props) {
  seen = [];
  const root = createRoot(document.getElementById('root'));
  await act(async () => root.render(React.createElement(Probe, props)));
  t.after(async () => act(async () => root.unmount()));
  return { render: async (patch) => act(async () => root.render(React.createElement(Probe, { ...props, ...patch }))) };
}
const defaults = (releases = []) => ({
  preview: async () => job(),
  page: async (_id, number) => page(number),
  release: async (id) => {
    releases.push(id);
  },
  cancel: async () => ({ granted: true }),
});
test('one accepted job feeds actual pane and central physical pages and validates later page content', async (t) => {
  const released = [],
    calls = [],
    callbacks = {
      ...defaults(released),
      preview: async () => {
        calls.push('prepare');
        return job();
      },
      page: async (_id, number) => {
        calls.push(number);
        return page(number);
      },
    };
  await mount(t, { scope: scope(), callbacks, showPane: true, showPage: true });
  await act(async () => {
    await api.prepare(settings, 'one', { origin: { kind: 'saved', range: settings.area } });
  });
  assert.equal(document.querySelectorAll('iframe').length, 2);
  assert(document.querySelector('.sheet-page-layout iframe').srcdoc.includes('data-print-ref="A1"'));
  await act(async () => api.choosePage(2));
  assert(document.querySelector('.sheet-page-layout iframe').srcdoc.includes('data-print-ref="A3"'));
  assert.deepEqual(calls, ['prepare', 2]);
  assert.deepEqual(released, []);
});
test('owner/session or activation changes immediately mask old accepted page/status in the transition render', async (t) => {
  const released = [],
    props = { scope: scope(), callbacks: defaults(released) },
    ui = await mount(t, props);
  await act(async () => api.prepare(settings));
  const old = api;
  await ui.render({ scope: { ...scope(), ownerKey: 'worker-B:base' } });
  const first = seen.find((value) => value.key === 'worker-B:base');
  assert.equal(first.job, undefined);
  assert.equal(first.page, undefined);
  assert.equal(first.status, '');
  assert.deepEqual(released, ['job-A']);
  await act(async () => api.prepare(settings));
  old.release();
  assert.equal(api.job.id, 'job-A');
  await ui.render({ scope: { ...scope(), ownerKey: 'worker-B:base', active: false, activationEpoch: 1 } });
  assert.equal(seen.find((value) => !value.active).job, undefined);
});
test('Normal abandons only preparation and late job retires silently through its captured old callback', async (t) => {
  const gate = deferred(),
    released = [],
    cancelled = [],
    callbacks = {
      ...defaults(released),
      preview: () => gate.promise,
      cancel: async (id) => {
        cancelled.push(id);
        return { granted: false, reason: 'completed' };
      },
    };
  await mount(t, { scope: scope(), callbacks });
  let pending;
  await act(async () => {
    pending = api.prepare(settings, 'normal-request');
    pending.catch(() => {});
  });
  await act(async () => api.abandonPreparation());
  assert.equal(api.preparing, false);
  assert(!api.status.includes('Preparing'));
  await act(async () => gate.resolve(job('late')));
  await assert.rejects(pending, isPrintPreviewRetired);
  assert.deepEqual(released, ['late']);
  assert.deepEqual(cancelled, ['normal-request']);
  assert.equal(api.job, null);
  assert.equal(api.error, '');
});
test('malformed geometry/settings/page refuses and releases before rendering/adoption', async (t) => {
  let response = job(),
    released = [];
  await mount(t, { scope: scope(), callbacks: { ...defaults(released), preview: async () => response } });
  for (const mutate of [
    (value) => delete value.pageSizePt,
    (value) => (value.scale = NaN),
    (value) => (value.settings = { ...settings, orientation: 'landscape' }),
    (value) => (value.pages[0].cells[0].bbox[2] = NaN),
  ]) {
    response = structuredClone(job());
    mutate(response);
    await act(async () => {
      await assert.rejects(api.prepare(settings));
    });
    assert.equal(api.job, null);
  }
  assert.equal(released.length, 4);
});
test('stale pending owner reply cannot adopt after another owner returns to the same base head', async (t) => {
  const gate = deferred(),
    released = [],
    props = { scope: scope(), callbacks: { ...defaults(released), preview: () => gate.promise } },
    ui = await mount(t, props);
  let pending;
  await act(async () => {
    pending = api.prepare(settings);
    pending.catch(() => {});
  });
  await ui.render({ scope: { ...scope(), grantDocumentId: 'durable-new', ownerKey: 'worker-new:base' } });
  await act(async () => gate.resolve(job('old-reload')));
  await assert.rejects(pending, isPrintPreviewRetired);
  assert.equal(api.job, null);
  assert.deepEqual(released, ['old-reload']);
});
test('Finishing cancellation keeps preparation pending and only explicit granted cancellation retires it', async (t) => {
  const gate = deferred();
  let granted = false;
  const props = {
      scope: scope(),
      callbacks: {
        ...defaults(),
        preview: () => gate.promise,
        cancel: async () => (granted ? { granted: true } : { granted: false, reason: 'finishing' }),
      },
    },
    ui = await mount(t, props);
  let pending;
  await act(async () => {
    pending = api.prepare(settings);
    pending.catch(() => {});
  });
  await act(async () => api.cancelPreparation());
  assert.equal(api.preparing, true);
  assert(api.status.includes('finishing'));
  granted = true;
  await act(async () => api.cancelPreparation());
  assert.equal(api.preparing, false);
  await act(async () => gate.resolve(job()));
  await assert.rejects(pending, isPrintPreviewRetired);
  assert.equal(api.job, null);
});
test('Normal and scope cleanup cannot cancel/release publication; exact old release waits for terminal output', async (t) => {
  const output = deferred(),
    released = [],
    cancelled = [],
    props = {
      scope: scope(),
      callbacks: {
        ...defaults(released),
        publish: () => output.promise,
        cancel: async (id) => {
          cancelled.push(id);
          return { granted: true };
        },
      },
    },
    ui = await mount(t, props);
  await act(async () => api.prepare(settings));
  let pending;
  await act(async () => {
    pending = api.publish('job-A', 'pdf', 'publish-A');
  });
  await act(async () => api.abandonPreparation());
  assert.equal(api.publishing, true);
  assert.deepEqual(cancelled, []);
  assert.deepEqual(released, []);
  await act(async () => {
    await assert.rejects(api.prepare(settings), /current print operation/);
  });
  await ui.render({ scope: { ...scope(), active: false, activationEpoch: 1 } });
  assert.deepEqual(released, []);
  await act(async () => output.resolve({ exported: true }));
  assert.deepEqual(await pending, { exported: true });
  assert.deepEqual(released, ['job-A']);
});
test('shared Grid Page Layout and Normal preserve mounted input/scroll and all six facets focus real pane fields without writes', async (t) => {
  const root = createRoot(document.getElementById('root')),
    calls = [],
    sheet = {
      name: 'Sheet1',
      state: 'visible',
      path: 'xl/worksheets/sheet1.xml',
      rows: 64,
      columns: 12,
      cells: { A1: { value: 1 }, A2: { value: 2 } },
      merges: [],
      widths: [],
      hiddenRows: [],
      rowHeights: [],
      freeze: { rows: 0, columns: 0 },
    },
    controls = {
      documentId: 'runtime-A',
      revision: 'base',
      definition: async () => ({ sheet: 'Sheet1', revision: null, settings, canEdit: true, protectedReasons: [] }),
      preview: async (_sheet, _settings, _requestId, geometryProfile) => {
        calls.push('prepare');
        return { ...job(), geometryProfile };
      },
      page: async (_id, n) => page(n),
      onBatch: async () => {
        calls.push('write');
      },
      publish: async () => {},
      release: async () => {},
      cancel: async () => ({ granted: true }),
    };
  await act(async () =>
    root.render(
      React.createElement(SpreadsheetGrid, {
        sheets: [sheet],
        documentId: 'durable-A',
        storageKey: 'worker-A:base',
        readOnly: false,
        onEdit: () => calls.push('write'),
        onBatch: () => calls.push('write'),
        onCommandBatch: async () => calls.push('write'),
        printControls: controls,
      })
    )
  );
  t.after(async () => act(async () => root.unmount()));
  const tab = (name) => [...document.querySelectorAll('[role="tab"]')].find((node) => node.textContent === name),
    command = (id) => document.querySelector('[data-ribbon-command="' + id + '"]');
  const original = document.querySelector('input[aria-label="Sheet1 A1"]'),
    scroll = document.querySelector('.sheet-grid-scroll');
  scroll.scrollTop = 137;
  scroll.scrollLeft = 27;
  await act(async () => tab('View').click());
  const visiblePanel = document.querySelector('.sheet-ribbon-tabpanel:not([hidden])'),
    viewGroups = [...visiblePanel.querySelectorAll('.sheet-ribbon-group')].filter(
      (group) => group.querySelector('small')?.textContent === 'Workbook views'
    );
  assert.equal(viewGroups.length, 1);
  assert.deepEqual(
    [...viewGroups[0].querySelectorAll('[data-ribbon-command]')].map((node) => node.dataset.ribbonCommand),
    ['view.normal', 'view.page-layout', 'view.page-break-preview']
  );
  await act(async () => command('view.page-layout').click());
  assert(document.querySelector('.sheet-page-layout'));
  assert.equal(document.querySelector('input[aria-label="Sheet1 A1"]'), original);
  assert.equal(scroll.closest('.sheet-grid-viewport').hidden, true);
  assert.equal(scroll.inert, true);
  assert.equal(original.readOnly, true);
  assert.equal(document.querySelector('.sheet-formula-bar').hidden, true);
  assert.equal(document.querySelector('button[aria-label="Bold"]').disabled, true);
  await act(async () => tab('Page Layout').click());
  await act(async () => command('print.margins').click());
  assert.equal(document.querySelector('.sheet-page-layout'), null);
  assert.equal(scroll.closest('.sheet-grid-viewport').hidden, false);
  assert.equal(scroll.inert, false);
  assert.equal(document.activeElement.dataset.printFacet, 'margins');
  await act(async () => tab('View').click());
  await act(async () => command('view.page-layout').click());
  await act(async () => command('view.normal').click());
  assert.equal(scroll.closest('.sheet-grid-viewport').hidden, false);
  assert.equal(scroll.inert, false);
  assert.equal(scroll.scrollTop, 137);
  assert.equal(scroll.scrollLeft, 27);
  assert.equal(document.querySelector('input[aria-label="Sheet1 A1"]'), original);
  await act(async () => tab('Page Layout').click());
  const pageGroups = [
      ...document.querySelector('.sheet-ribbon-tabpanel:not([hidden])').querySelectorAll('.sheet-ribbon-group'),
    ],
    setup = pageGroups.find((group) => group.querySelector('small')?.textContent === 'Page setup'),
    scaling = pageGroups.find((group) => group.querySelector('small')?.textContent === 'Scale to fit');
  assert.deepEqual(
    [...setup.querySelectorAll('[data-ribbon-command]')].map((node) => node.dataset.ribbonCommand),
    ['print.margins', 'print.orientation', 'print.paper', 'print.area', 'print.titles', 'print.setup']
  );
  assert.deepEqual(
    [...scaling.querySelectorAll('[data-ribbon-command]')].map((node) => node.dataset.ribbonCommand),
    ['print.scale']
  );
  for (const facet of ['margins', 'orientation', 'paper', 'area', 'titles', 'scale']) {
    await act(async () => command('print.' + facet).click());
    assert.equal(document.activeElement.dataset.printFacet, facet);
  }
  assert.deepEqual(calls, ['prepare']);
});

test('retained pane treats same-key Normal retirement silently and cannot latch frame loading on a retired page read', async (t) => {
  const prepareGate = deferred(),
    pageGate = deferred();
  let producer = () => prepareGate.promise;
  const callbacks = { ...defaults(), preview: () => producer(), page: () => pageGate.promise };
  await mount(t, { scope: scope(), callbacks, showPane: true });
  const button = (name) => [...document.querySelectorAll('button')].find((node) => node.textContent === name);
  await act(async () => button('Preview pages').click());
  await act(async () => api.abandonPreparation());
  await act(async () => prepareGate.resolve(job('retired-pane')));
  assert.equal(api.job, null);
  assert(!document.body.textContent.includes('This page preparation was retired.'));
  assert.equal(button('Preview pages').disabled, false);
  producer = async () => job('current-pane');
  await act(async () => button('Preview pages').click());
  const frame = document.querySelector('iframe'),
    doc = frame.contentDocument;
  doc.open();
  doc.write(frame.srcdoc);
  doc.close();
  await act(async () => frame.dispatchEvent(new window.Event('load')));
  await act(async () => button('Next page').click());
  assert.equal(api.pageLoading, true);
  await act(async () => api.abandonPreparation());
  await act(async () => pageGate.resolve(page(2)));
  assert.equal(api.page.number, 1);
  assert.equal(api.pageLoading, false);
  assert.equal(button('Next page').disabled, false);
  assert(!document.body.textContent.includes('This page preparation was retired.'));
});

test('retained draft cannot publish a central saved preview; Preview pages adopts the preserved draft settings', async (t) => {
  const root = createRoot(document.getElementById('root')),
    requests = [],
    published = [],
    saved = { ...settings, paper: 'A4', orientation: 'portrait' };
  const sheet = {
    name: 'Sheet1',
    state: 'visible',
    path: 'xl/worksheets/sheet1.xml',
    rows: 64,
    columns: 12,
    cells: { A1: { value: 1 }, A2: { value: 2 } },
    merges: [],
    widths: [],
    hiddenRows: [],
    rowHeights: [],
    freeze: { rows: 0, columns: 0 },
  };
  const one = {
    number: 1,
    rowIndices: [0, 1, 2],
    columnIndices: [0, 1],
    cells: [1, 2, 3].flatMap((row, r) =>
      ['A', 'B'].map((col, c) => ({
        ref: col + row,
        text: col + row,
        style: { fontSize: 11 },
        bbox: [36 + c * 100, 36 + r * 20, 100, 20],
      }))
    ),
    objects: [],
  };
  const controls = {
    documentId: 'runtime-A',
    revision: 'base',
    canPrint: true,
    definition: async () => ({ sheet: 'Sheet1', revision: null, settings: saved, canEdit: true, protectedReasons: [] }),
    preview: async (_sheet, requested, _requestId, geometryProfile) => {
      requests.push(structuredClone(requested));
      let size = requested.paper === 'A4' ? [(210 / 25.4) * 72, (297 / 25.4) * 72] : [612, 792];
      if (requested.orientation === 'landscape') size.reverse();
      return {
        ...job('settings-' + requests.length),
        geometryProfile,
        settings: requested,
        pageSizePt: size,
        pageCount: 1,
        pages: [one],
      };
    },
    page: async () => one,
    onBatch: async () => {
      throw Error('No settings write');
    },
    publish: async (...args) => {
      published.push(args);
      return { exported: true };
    },
    release: async () => {},
    cancel: async () => ({ granted: true }),
  };
  await act(async () =>
    root.render(
      React.createElement(SpreadsheetGrid, {
        sheets: [sheet],
        documentId: 'durable-A',
        storageKey: 'worker-A:base',
        readOnly: false,
        onEdit: () => {
          throw Error('No worksheet write');
        },
        onCommandBatch: async () => {
          throw Error('No workbook write');
        },
        printControls: controls,
      })
    )
  );
  t.after(async () => act(async () => root.unmount()));
  const tab = (name) => [...document.querySelectorAll('[role="tab"]')].find((node) => node.textContent === name),
    command = (id) => document.querySelector('[data-ribbon-command="' + id + '"]'),
    button = (name) =>
      [...document.querySelectorAll('button')].find((node) => !node.closest('[hidden]') && node.textContent === name);
  await act(async () => tab('Page Layout').click());
  await act(async () => command('print.orientation').click());
  for (const [name, value] of [
    ['Paper', 'Letter'],
    ['Orientation', 'landscape'],
  ])
    await act(async () => {
      const select = document.querySelector('select[aria-label="' + name + '"]');
      select.value = value;
      select.dispatchEvent(new window.Event('change', { bubbles: true }));
    });
  assert.equal(document.querySelector('select[aria-label="Paper"]').value, 'Letter');
  assert.equal(document.querySelector('select[aria-label="Orientation"]').value, 'landscape');
  await act(async () => tab('View').click());
  await act(async () => command('view.page-layout').click());
  assert.equal(requests[0].paper, 'A4');
  assert.equal(requests[0].orientation, 'portrait');
  await act(async () => tab('Page Layout').click());
  await act(async () => command('print.orientation').click());
  assert.equal(document.querySelector('.sheet-page-layout'), null);
  const load = async () => {
    const frame = document.querySelector('.sheet-print-preview iframe'),
      doc = frame.contentDocument;
    doc.open();
    doc.write(frame.srcdoc);
    doc.close();
    await act(async () => frame.dispatchEvent(new window.Event('load')));
  };
  await load();
  assert.equal(document.querySelector('select[aria-label="Paper"]').value, 'Letter');
  assert.equal(document.querySelector('select[aria-label="Orientation"]').value, 'landscape');
  assert.equal(button('Export PDF').disabled, true);
  assert.equal(button('Print').disabled, true);
  assert.match(document.body.textContent, /Preview pages to use these settings/);
  assert.deepEqual(published, []);
  await act(async () => button('Preview pages').click());
  await load();
  assert.equal(requests[1].paper, 'Letter');
  assert.equal(requests[1].orientation, 'landscape');
  assert.equal(button('Export PDF').disabled, false);
  assert.equal(button('Print').disabled, false);
  await act(async () => button('Export PDF').click());
  assert.equal(published.length, 1);
  assert.equal(published[0][0], 'settings-2');
});

test('actual Grid re-entry reuses matching accepted pages across Normal and facets; changed settings or head prepare anew', async (t) => {
  const root = createRoot(document.getElementById('root')),
    requests = [],
    released = [];
  let state = 'base',
    saved = structuredClone(settings);
  const sheet = {
    name: 'Sheet1',
    state: 'visible',
    path: 'xl/worksheets/sheet1.xml',
    rows: 64,
    columns: 12,
    cells: { A1: { value: 1 }, A2: { value: 2 } },
    merges: [],
    widths: [],
    hiddenRows: [],
    rowHeights: [],
    freeze: { rows: 0, columns: 0 },
  };
  const one = {
    number: 1,
    rowIndices: [0, 1, 2],
    columnIndices: [0, 1],
    cells: [1, 2, 3].flatMap((row, r) =>
      ['A', 'B'].map((col, c) => ({
        ref: col + row,
        text: col + row,
        style: { fontSize: 11 },
        bbox: [36 + c * 100, 36 + r * 20, 100, 20],
      }))
    ),
    objects: [],
  };
  const render = () =>
    act(async () =>
      root.render(
        React.createElement(SpreadsheetGrid, {
          sheets: [sheet],
          documentId: 'durable-A',
          storageKey: 'worker-A:' + state,
          readOnly: false,
          onCommandBatch: async () => {
            throw Error('No writes');
          },
          printControls: {
            documentId: 'runtime-A',
            revision: state,
            definition: async () => ({
              sheet: 'Sheet1',
              revision: null,
              settings: structuredClone(saved),
              canEdit: true,
              protectedReasons: [],
            }),
            preview: async (_sheet, requested, _requestId, geometryProfile) => {
              requests.push(structuredClone(requested));
              const size = requested.orientation === 'landscape' ? [792, 612] : [612, 792];
              return {
                ...job('reuse-' + requests.length, state),
                geometryProfile,
                settings: requested,
                pageSizePt: size,
                pageCount: 1,
                pages: [one],
              };
            },
            page: async () => one,
            release: async (id) => released.push(id),
            cancel: async () => ({ granted: true }),
            publish: async () => {
              throw Error('No output');
            },
          },
        })
      )
    );
  await render();
  t.after(async () => act(async () => root.unmount()));
  const tab = (name) => [...document.querySelectorAll('[role="tab"]')].find((node) => node.textContent === name),
    command = (id) => document.querySelector('[data-ribbon-command="' + id + '"]');
  const enter = async () => {
    await act(async () => tab('View').click());
    await act(async () => command('view.page-layout').click());
    return document.querySelector('.sheet-page-layout iframe')?.dataset.acceptedPrintJob;
  };
  const normal = () => act(async () => command('view.normal').click());
  assert.equal(await enter(), 'reuse-1');
  await normal();
  assert.equal(await enter(), 'reuse-1');
  assert.equal(requests.length, 1);
  await act(async () => tab('Page Layout').click());
  await act(async () => command('print.area').click());
  assert.equal(document.querySelector('.sheet-page-layout'), null);
  assert.equal(await enter(), 'reuse-1');
  assert.equal(requests.length, 1);
  assert.deepEqual(released, []);
  await normal();
  saved = { ...saved, orientation: 'landscape' };
  assert.equal(await enter(), 'reuse-2');
  assert.equal(requests.length, 2);
  assert.deepEqual(released, ['reuse-1']);
  state = 'accepted-next';
  await render();
  assert.equal(document.querySelector('.sheet-page-layout'), null);
  assert.deepEqual(released, ['reuse-1', 'reuse-2']);
  assert.equal(await enter(), 'reuse-3');
  assert.equal(requests.length, 3);
  assert.equal(document.querySelector('.sheet-page-layout iframe').dataset.acceptedPrintState, 'accepted-next');
});

test('reuseCurrent requires a live matching lease and normalized settings, updates truthful origin and never revives a foreign callback', async (t) => {
  const released = [],
    props = { scope: scope(), callbacks: defaults(released) },
    ui = await mount(t, props);
  await act(async () => api.prepare(settings, 'reuse-origin', { origin: { kind: 'draft', range: settings.area } }));
  const old = api;
  await act(async () => {
    assert.equal(
      api.reuseCurrent(
        { ...settings, marginsPt: { ...settings.marginsPt } },
        { origin: { kind: 'saved', range: settings.area } }
      ).id,
      'job-A'
    );
  });
  assert.deepEqual(api.origin, { kind: 'saved', range: settings.area });
  await act(async () => {
    assert.equal(api.reuseCurrent(settings, { origin: { kind: 'selection', range: settings.area } }).id, 'job-A');
  });
  assert.deepEqual(api.origin, { kind: 'selection', range: settings.area });
  assert.equal(api.reuseCurrent({ ...settings, orientation: 'landscape' }), null);
  assert.equal(api.reuseCurrent(settings, { guard: () => false }), null);
  await ui.render({ scope: { ...scope(), ownerKey: 'worker-new:base', grantDocumentId: 'durable-new' } });
  assert.equal(api.reuseCurrent(settings), null);
  assert.equal(old.reuseCurrent(settings), null);
  assert.deepEqual(released, ['job-A']);
  await act(async () => api.prepare(settings));
  const current = api;
  await act(async () => api.release());
  assert.equal(current.reuseCurrent(settings), null);
});

// V2 reuses the real compositor; mocks only the owner callback/lifecycle.
async function boundaryFixture() {
  const { composePrint } = await import('../core/print-layout.mjs');
  const { defaultPrintSettings } = await import('../core/print-model.mjs');
  const saved = { ...defaultPrintSettings(), area: 'A1:B40', repeatRows: { first: 1, last: 1 } },
    source = {
      schema: 1,
      sessionId: 'owner',
      state: 'base',
      revision: 'source',
      sheet: 'Sheet1',
      date1904: false,
      fingerprint: 'b'.repeat(64),
      settings: saved,
      rows: Array.from({ length: 40 }, (_, index) => ({ index, heightPt: 36, hidden: index === 2 })),
      columns: [0, 1].map((index) => ({ index, widthPt: 80, hidden: false })),
      cells: Array.from({ length: 40 }, (_, r) =>
        [0, 1].map((c) => ({ ref: (c ? 'B' : 'A') + (r + 1), text: 'r' + r + 'c' + c, style: { fontSize: 8 } }))
      ).flat(),
      merges: [],
      objects: [],
    };
  return {
    saved,
    make: (id, state, geometryProfile, settings = saved) => {
      const layout = composePrint({ ...source, settings });
      return { ...layout, id, state, geometryProfile, pageCount: layout.pages.length, pages: [layout.pages[0]] };
    },
  };
}
function boundarySheet() {
  return {
    name: 'Sheet1',
    state: 'visible',
    path: 'xl/worksheets/sheet1.xml',
    rows: 64,
    columns: 12,
    sheetFormat: { defaultRowHeight: 36 },
    cells: { A1: { value: 1 } },
    merges: [],
    widths: [],
    hiddenRows: [2],
    rowHeights: [],
    freeze: { rows: 1, columns: 1 },
  };
}
const boundaryTab = (name) => [...document.querySelectorAll('[role="tab"]')].find((node) => node.textContent === name),
  boundaryCommand = (id) => document.querySelector('[data-ribbon-command="' + id + '"]');
test('V2 shares one accepted lease across Page Layout/boundaries and restores Normal selection/scroll with no writes', async (t) => {
  const f = await boundaryFixture(),
    root = createRoot(document.getElementById('root')),
    prepared = [],
    writes = [],
    released = [],
    sheet = boundarySheet();
  const controls = {
    documentId: 'runtime-A',
    revision: 'base',
    definition: async () => ({
      sheet: 'Sheet1',
      revision: null,
      settings: f.saved,
      canEdit: true,
      protectedReasons: [],
    }),
    preview: async (_s, settings, _id, geometryProfile) => {
      prepared.push(settings);
      return f.make('boundary-job', 'base', geometryProfile, settings);
    },
    page: async () => {
      throw Error('Boundary summaries need no additional physical-page read');
    },
    release: async (id) => released.push(id),
    cancel: async () => ({ granted: true }),
  };
  await act(async () =>
    root.render(
      React.createElement(SpreadsheetGrid, {
        sheets: [sheet],
        documentId: 'durable-A',
        storageKey: 'worker-A:base',
        readOnly: false,
        onEdit: (x) => writes.push(x),
        onCommandBatch: async (x) => writes.push(x),
        printControls: controls,
      })
    )
  );
  t.after(async () => act(async () => root.unmount()));
  const scroll = document.querySelector('.sheet-grid-scroll');
  scroll.scrollTop = 137;
  scroll.scrollLeft = 27;
  await act(async () => boundaryTab('View').click());
  await act(async () => boundaryCommand('view.page-layout').click());
  assert(document.querySelector('.sheet-page-layout'));
  await act(async () => boundaryCommand('view.page-break-preview').click());
  assert(document.querySelector('[aria-label="Page Break Preview"]'));
  assert.equal(document.querySelector('.sheet-page-layout'), null);
  assert.equal(document.querySelector('.sheet-grid').inert, true);
  assert.equal(scroll.inert, false);
  assert.equal(boundaryCommand('view.normal').getAttribute('aria-pressed'), 'false');
  assert.equal(boundaryCommand('view.page-break-preview').getAttribute('aria-pressed'), 'true');
  assert.equal(prepared.length, 1);
  assert.equal(document.querySelector('.sheet-grid-viewport').hidden, false);
  assert.match(document.body.textContent, /Frozen panes are suspended/);
  const select = document.querySelector('select[aria-label="Go to page boundary"]');
  assert(select.options.length > 2);
  await act(async () => {
    select.value = '2';
    select.dispatchEvent(new window.Event('change', { bubbles: true }));
  });
  assert(scroll.scrollTop > 137);
  assert.equal(document.querySelector('[aria-label="Go to cell"]').placeholder, 'A1');
  await act(async () => boundaryCommand('view.normal').click());
  assert.equal(scroll.scrollTop, 137);
  assert.equal(scroll.scrollLeft, 27);
  assert.equal(document.querySelector('.sheet-grid').inert, false);
  assert.equal(document.querySelector('[aria-label="Go to cell"]').placeholder, 'A1');
  assert.deepEqual(writes, []);
  assert.deepEqual(released, []);
});
test('V2 pending work retires its original callbacks after an owner/head switch, with no foreign map or error', async (t) => {
  const f = await boundaryFixture(),
    root = createRoot(document.getElementById('root')),
    gate = deferred(),
    released = [];
  let state = 'base',
    owner = 'A',
    geometry;
  const render = () =>
    act(async () =>
      root.render(
        React.createElement(SpreadsheetGrid, {
          sheets: [boundarySheet()],
          documentId: 'durable-' + owner,
          storageKey: 'worker-' + owner + ':' + state,
          readOnly: false,
          onCommandBatch: async () => {
            throw Error('No writes');
          },
          printControls: {
            documentId: 'runtime-' + owner,
            revision: state,
            definition: async () => ({
              sheet: 'Sheet1',
              revision: null,
              settings: f.saved,
              canEdit: true,
              protectedReasons: [],
            }),
            preview: async (_s, _v, _r, g) => {
              geometry = g;
              return gate.promise;
            },
            page: async () => {
              throw Error('No page reads');
            },
            release: async (id) => released.push(id),
            cancel: async () => ({ granted: true }),
          },
        })
      )
    );
  await render();
  t.after(async () => act(async () => root.unmount()));
  await act(async () => boundaryTab('View').click());
  await act(async () => boundaryCommand('view.page-break-preview').click());
  owner = 'B';
  state = 'head-B';
  await render();
  await act(async () => gate.resolve(f.make('old-A', 'base', geometry)));
  assert.equal(document.querySelector('[aria-label="Page Break Preview"]'), null);
  assert.equal(document.querySelector('.sheet-page-boundary-overlay'), null);
  assert(!document.body.textContent.includes('This page preparation was retired'));
  assert.deepEqual(released, ['old-A']);
});
test('V2 head change removes the map and releases once; changed settings replace rather than reuse old summaries', async (t) => {
  const f = await boundaryFixture(),
    root = createRoot(document.getElementById('root')),
    released = [],
    prepared = [];
  let state = 'base',
    saved = f.saved;
  const render = () =>
    act(async () =>
      root.render(
        React.createElement(SpreadsheetGrid, {
          sheets: [boundarySheet()],
          documentId: 'durable-A',
          storageKey: 'worker-A:' + state,
          readOnly: false,
          onCommandBatch: async () => {
            throw Error('No writes');
          },
          printControls: {
            documentId: 'runtime-A',
            revision: state,
            definition: async () => ({
              sheet: 'Sheet1',
              revision: null,
              settings: saved,
              canEdit: true,
              protectedReasons: [],
            }),
            preview: async (_s, v, _r, g) => {
              prepared.push(v);
              return f.make('map-' + prepared.length, state, g, v);
            },
            page: async () => {
              throw Error('No page reads');
            },
            release: async (id) => released.push(id),
            cancel: async () => ({ granted: true }),
          },
        })
      )
    );
  await render();
  t.after(async () => act(async () => root.unmount()));
  const enter = async () => {
    await act(async () => boundaryTab('View').click());
    await act(async () => boundaryCommand('view.page-break-preview').click());
    assert(document.querySelector('[aria-label="Page Break Preview"]'));
  };
  await enter();
  await act(async () => boundaryCommand('view.normal').click());
  saved = { ...saved, scaling: { mode: 'fit', width: 1, height: 1 } };
  await enter();
  assert.equal(prepared.length, 2);
  assert.deepEqual(released, ['map-1']);
  assert.equal(document.querySelector('select[aria-label="Go to page boundary"]').options.length, 2);
  state = 'head-next';
  await render();
  assert.equal(document.querySelector('.sheet-page-boundary-overlay'), null);
  assert.equal(document.querySelector('[aria-label="Page Break Preview"]'), null);
  assert.deepEqual(released, ['map-1', 'map-2']);
});
