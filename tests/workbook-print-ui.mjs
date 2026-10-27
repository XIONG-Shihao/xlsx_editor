import './support/scratch-dir.mjs';
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
const directory = path.resolve(
  '.build/tests/taskpane-print-component'
);
await fs.mkdir(directory, { recursive: true });
const modulePath = path.join(directory, 'WorkbookPrint.mjs');
await build({
  entryPoints: ['ui/WorkbookPrint.tsx'],
  bundle: true,
  format: 'esm',
  platform: 'browser',
  jsx: 'automatic',
  outfile: modulePath,
  external: ['react', 'react/jsx-runtime'],
  logLevel: 'silent',
});
const { WorkbookPrint } = await import(pathToFileURL(modulePath));
const dom = new JSDOM('<!doctype html><div id="root"></div>', { url: 'http://localhost/' });
globalThis.window = dom.window;
globalThis.document = dom.window.document;
globalThis.HTMLElement = dom.window.HTMLElement;
globalThis.IS_REACT_ACT_ENVIRONMENT = true;
// A bounded DOM-only observer. Tests drive width notifications explicitly;
// they do not pretend that JSDOM measured a physical page or dock layout.
const observers = [];
globalThis.ResizeObserver = class {
  constructor(callback) {
    this.callback = callback;
    this.target = null;
    this.disconnected = false;
    this.disconnectCalls = 0;
    observers.push(this);
  }
  observe(target) {
    assert.equal(this.target, null, 'One preview container per observer');
    assert.equal(target.className, 'sheet-print-preview');
    this.target = target;
  }
  disconnect() {
    this.disconnected = true;
    this.disconnectCalls++;
  }
  notify(width) {
    assert.equal(this.disconnected, false);
    assert.ok(Number.isFinite(width) && width >= 0);
    this.callback([{ target: this.target, contentRect: { width } }]);
  }
};
const { createRoot } = await import('react-dom/client');
after(() => {
  dom.window.close();
  delete globalThis.window;
  delete globalThis.document;
  delete globalThis.HTMLElement;
  delete globalThis.ResizeObserver;
  delete globalThis.IS_REACT_ACT_ENVIRONMENT;
});
const settings = {
    area: 'A1:C50',
    paper: 'A4',
    orientation: 'portrait',
    marginsPt: { top: 36, right: 36, bottom: 36, left: 36, header: 18, footer: 18 },
    scaling: { mode: 'scale', percent: 100 },
    repeatRows: null,
    repeatColumns: null,
    gridlines: false,
  },
  definition = { sheet: 'Print', revision: null, canEdit: true, settings, protectedReasons: [] },
  page = (number) => ({
    number,
    rowIndices: [number === 1 ? 0 : 42],
    columnIndices: [0],
    cells: [
      { ref: number === 1 ? 'A1' : 'A43', text: number === 1 ? 'Item' : 'row-43', style: {}, bbox: [36, 36, 72, 18] },
    ],
    objects: [],
  }),
  job = {
    id: 'job1',
    units: 'pt',
    pageSizePt: [595.2755905512, 841.8897637795],
    scale: 1,
    sourceFingerprint: 'f'.repeat(64),
    revision: 'r1',
    sheet: 'Print',
    settings,
    pageCount: 2,
    pages: [page(1)],
  };
const base = () => ({
  documentId: 'doc1',
  revision: 'r1',
  sheet: 'Print',
  selection: 'A1:C50',
  definition: async () => definition,
  onBatch: async () => {},
  preview: async () => job,
  page: async (_id, number) => page(number),
  publish: async () => {},
  release: async () => {},
});
function deferred() {
  let resolve;
  const promise = new Promise((accept) => {
    resolve = accept;
  });
  return { promise, resolve };
}
const button = (text) => [...document.querySelectorAll('button')].find((node) => node.textContent === text);
async function click(text) {
  assert.ok(button(text), 'Missing ' + text);
  assert.equal(button(text).disabled, false, text + ' disabled');
  await act(async () => button(text).click());
}
async function mount(t, props) {
  const root = createRoot(document.getElementById('root'));
  let unmounted = false;
  await act(async () => root.render(React.createElement(WorkbookPrint, props)));
  const unmount = async () => {
    if (!unmounted) {
      unmounted = true;
      await act(async () => root.unmount());
    }
  };
  t.after(unmount);
  return { render: async (props) => act(async () => root.render(React.createElement(WorkbookPrint, props))), unmount };
}
async function loaded() {
  const frame = document.querySelector('iframe');
  await act(async () => {
    frame.contentDocument.open();
    frame.contentDocument.write(frame.srcdoc);
    frame.contentDocument.close();
    frame.dispatchEvent(new window.Event('load'));
  });
}
test('page2 replaces the iframe node and contains the requested physical page/A43, with load status before readiness', async (t) => {
  const requests = [],
    props = {
      ...base(),
      page: async (id, number) => {
        requests.push([id, number]);
        return page(number);
      },
    };
  await mount(t, props);
  await click('Preview pages');
  const first = document.querySelector('iframe');
  assert.ok(first.srcdoc.includes('data-print-page="1"'));
  assert.ok(first.srcdoc.includes('>Item<'));
  assert.ok(document.body.textContent.includes('Loading print page…'));
  assert.equal(button('Next page').disabled, true);
  await loaded();
  await click('Next page');
  const second = document.querySelector('iframe');
  assert.notEqual(second, first);
  assert.equal(first.isConnected, false);
  assert.deepEqual(requests, [['job1', 2]]);
  assert.ok(second.srcdoc.includes('data-print-page="2"'));
  assert.ok(second.srcdoc.includes('data-print-ref="A43"'));
  assert.ok(second.srcdoc.includes('>row-43<'));
  assert.ok(!second.srcdoc.includes('>Item<'));
  assert.ok(document.body.textContent.includes('Page 2 / 2'));
  assert.equal(button('Export PDF').disabled, true);
  await loaded();
  assert.equal(button('Export PDF').disabled, false);
});
test('mismatched page reply is refused and delayed replies cannot cross replacement lifetime', async (t) => {
  const pending = deferred(),
    released = [],
    props = { ...base(), page: async () => page(1), release: async (id) => released.push(id) },
    ui = await mount(t, props);
  await click('Preview pages');
  await loaded();
  await click('Next page');
  assert.ok(document.body.textContent.includes('did not match'));
  assert.ok(document.body.textContent.includes('Page 1 / 2'));
  await ui.render({
    ...props,
    revision: 'r2',
    preview: async () => ({ ...job, revision: 'r2' }),
    page: async () => pending.promise,
  });
  await click('Preview pages');
  await loaded();
  await click('Next page');
  await ui.render({ ...props, documentId: 'doc2', revision: 'r3' });
  await act(async () => pending.resolve(page(2)));
  assert.equal(document.querySelector('iframe'), null);
  assert.ok(!document.body.textContent.includes('Page 2 / 2'));
  assert.ok(released.length >= 1);
});
test('preview and PDF publication expose opaque cancellation requests and use accepted settings guards', async (t) => {
  const preview = deferred(),
    cancelled = [],
    writes = [],
    published = [],
    props = {
      ...base(),
      preview: async (_settings, id) => {
        assert.ok(id);
        return preview.promise;
      },
      cancel: async (id) => cancelled.push(id),
      onBatch: async (commands) => writes.push(commands),
    },
    ui = await mount(t, props);
  await click('Apply print settings');
  assert.equal(writes[0][0].expectedPrintRevision, null);
  assert.equal(writes[0][0].settings.area, 'A1:C50');
  await click('Preview pages');
  await click('Cancel print preparation');
  assert.equal(cancelled.length, 1);
  await act(async () => preview.resolve(job));
  assert.equal(document.querySelector('iframe'), null);
  await ui.render({
    ...props,
    revision: 'r2',
    preview: async () => ({ ...job, revision: 'r2' }),
    publish: async (...args) => {
      published.push(args);
      return { cancelled: true };
    },
  });
  await click('Preview pages');
  await loaded();
  assert.equal(button('Print').disabled, true);
  await click('Export PDF');
  assert.equal(published[0][0], 'job1');
  assert.equal(published[0][1], 'pdf');
  assert.ok(published[0][2]);
  assert.ok(document.body.textContent.includes('Print output cancelled.'));
});
test('denied finishing cancellation keeps the active publication and reports its actual successful result', async (t) => {
  const publication = deferred(),
    props = {
      ...base(),
      publish: async () => publication.promise,
      cancel: async () => ({ granted: false, reason: 'finishing' }),
    };
  await mount(t, props);
  await click('Preview pages');
  await loaded();
  await act(async () => button('Export PDF').click());
  await click('Cancel print preparation');
  assert.ok(document.body.textContent.includes('Print output is finishing…'));
  assert.ok(!document.body.textContent.includes('Print preparation cancelled.'));
  assert.equal(button('Export PDF').disabled, true);
  await act(async () => publication.resolve({ printPublished: { entryId: 'pdf1' } }));
  assert.ok(document.body.textContent.includes('PDF output finished.'));
  assert.equal(button('Export PDF').disabled, false);
  assert.equal(button('Cancel print preparation'), undefined);
});
test('fitted dock preview scales full physical iframe without changing page geometry; hidden zero/reopen and cleanup are bounded', async (t) => {
  const released = [],
    props = { ...base(), release: async (id) => released.push(id) },
    ui = await mount(t, props);
  await click('Preview pages');
  const frame = document.querySelector('iframe'),
    container = document.querySelector('.sheet-print-preview'),
    observer = observers.at(-1),
    physicalHtml = frame.srcdoc,
    naturalWidth = (job.pageSizePt[0] * 96) / 72,
    naturalHeight = (job.pageSizePt[1] * 96) / 72;
  assert.equal(observer.target, container);
  assert.equal(observers.filter((value) => !value.disconnected).length, 1);
  assert.equal(parseFloat(frame.style.width), naturalWidth);
  assert.equal(parseFloat(frame.style.height), naturalHeight);
  const assertFit = (width) => {
    const scale = Math.min(1, width / naturalWidth);
    assert.equal(frame.style.transform, `scale(${scale})`);
    assert.equal(parseFloat(container.style.height), naturalHeight * scale);
    assert.equal(parseFloat(frame.style.width), naturalWidth);
    assert.equal(parseFloat(frame.style.height), naturalHeight);
    assert.equal(frame.srcdoc, physicalHtml);
    assert.deepEqual(job.pageSizePt, [595.2755905512, 841.8897637795]);
    assert.equal(job.scale, 1);
  };
  await act(async () => observer.notify(280));
  assertFit(280);
  container.parentElement.hidden = true;
  await act(async () => observer.notify(0));
  assertFit(280);
  assert.equal(observer.disconnected, false, 'Hiding a retained pane does not release the preview');
  assert.deepEqual(released, []);
  container.parentElement.hidden = false;
  await act(async () => observer.notify(360));
  assertFit(360);
  await act(async () => observer.notify(1200));
  assertFit(1200);
  assert.equal(frame.style.transform, 'scale(1)', 'Preview never scales beyond the physical page');
  await ui.render({ ...props, revision: 'r2' });
  assert.equal(document.querySelector('iframe'), null);
  assert.equal(observer.disconnected, true);
  assert.equal(observer.disconnectCalls, 1);
  assert.deepEqual(released, ['job1']);
  await click('Preview pages');
  const next = observers.at(-1);
  assert.notEqual(next, observer);
  assert.equal(next.disconnected, false);
  assert.equal(observers.filter((value) => !value.disconnected).length, 1);
  await ui.unmount();
  assert.equal(next.disconnected, true);
  assert.equal(next.disconnectCalls, 1);
  assert.equal(observers.filter((value) => !value.disconnected).length, 0);
  assert.deepEqual(released, ['job1', 'job1']);
});

test('three header/footer presets and source selection share a pure draft then one captured settings command', async (t) => {
  const writes = [],
    props = { ...base(), selection: 'B2:D20', onBatch: async (commands) => writes.push(commands) };
  await mount(t, props);
  const select = document.querySelector('[aria-label="Header/footer preset"]');
  for (const [preset, header, footer] of [
    ['title', 'Print', ''],
    ['confidential', 'Print', 'Confidential'],
    ['draft', 'Draft', ''],
  ]) {
    await act(async () => {
      select.value = preset;
      select.dispatchEvent(new window.Event('change', { bubbles: true }));
    });
    assert.equal(document.querySelector('[aria-label="header center"]').value, header);
    assert.equal(document.querySelector('[aria-label="footer center"]').value, footer);
    assert.equal(writes.length, 0);
  }
  await click('Use selection for print area');
  await click('Apply print settings');
  assert.equal(writes.length, 1);
  assert.equal(writes[0].length, 1);
  const command = writes[0][0];
  assert.equal(command.sheet, 'Print');
  assert.equal(command.expectedPrintRevision, null);
  assert.equal(command.settings.area, 'B2:D20');
  assert.equal(command.settings.headerFooter.header.center, 'Draft');
  assert.equal(command.settings.headerFooter.footer.right, 'Print');
  const margins = [...document.querySelectorAll('label')].filter((label) => label.textContent.includes('margin (pt)'));
  assert.equal(margins.length, 6);
});
test('header/footer draft cannot apply to readonly source and switching owner retires the prior draft', async (t) => {
  const writes = [],
    props = { ...base(), readOnly: true, onBatch: async (commands) => writes.push(commands) },
    ui = await mount(t, props);
  assert.equal(button('Apply print settings').disabled, true);
  const select = document.querySelector('[aria-label="Header/footer preset"]');
  await act(async () => {
    select.value = 'confidential';
    select.dispatchEvent(new window.Event('change', { bubbles: true }));
  });
  assert.equal(writes.length, 0);
  await ui.render({ ...props, documentId: 'doc-other', revision: 'r-other', readOnly: false });
  assert.equal(document.querySelector('[aria-label="header center"]').value, '');
  await click('Apply print settings');
  assert.equal(writes.length, 1);
  assert.equal(writes[0][0].settings.headerFooter, undefined);
});
test('shared preview failure is shown once and never leaves duplicate alert messages', async (t) => {
  const session = {
    job: null,
    page: null,
    number: 1,
    preparing: false,
    pageLoading: false,
    status: '',
    error: 'Header does not fit',
    release: () => {},
    prepare: async () => {
      throw Error('Header does not fit');
    },
  };
  await mount(t, { ...base(), previewSession: session });
  assert.equal(document.querySelectorAll('[role="alert"]').length, 1);
  await click('Preview pages');
  assert.equal(document.querySelectorAll('[role="alert"]').length, 1);
  assert.equal(document.querySelector('[role="alert"]').textContent, 'Header does not fit');
});
