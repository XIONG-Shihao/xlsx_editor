import './support/scratch-dir.mjs';
/** Actual current typed TTC model/pane/Grid; controlled host, not native proof. */
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
  output = await fs.mkdtemp(path.join(component, '.build/tests/typed-ttc-ui-')),
  checks = [];
const entry = path.join(output, 'entry.ts'),
  bundle = path.join(output, 'test.mjs');
await fs.writeFile(
  entry,
  `export * from ${JSON.stringify(path.join(component, 'ui/textColumns.ts'))};export {WorkbookTextColumns} from ${JSON.stringify(path.join(component, 'ui/WorkbookTextColumns.tsx'))};export {SpreadsheetGrid} from ${JSON.stringify(path.join(component, 'ui/SpreadsheetGrid.tsx'))};export {configureXlsxHost} from ${JSON.stringify(path.join(component, 'ui/host.tsx'))};`
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
  test(name, { timeout: 20000 }, async (t) => {
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
const text = (value) => ({ value, dataType: 'text' }),
  binding = () => ({
    documentId: 'doc',
    grantDocumentId: 'grant',
    ownerIdentity: 'owner',
    ownerKey: 'owner:base',
    state: 'head',
    sheet: 'S',
    selection: 'A1:A2',
    navigationGeneration: 0,
    activationEpoch: 0,
  }),
  options = { delimiter: ',', quotes: 'double', destination: 'C1' };
function fixture(patch = {}) {
  let current = true,
    runtime = true;
  const data = { A1: text('00123,12.50,TRUE,2024-02-29'), A2: text('=1+2,9.25,FALSE,') },
    reads = [],
    accepted = [];
  const context = {
    binding: binding(),
    date1904: false,
    isCurrent: () => current,
    isRuntimeCurrent: () => runtime,
    availability: () => ({ active: true, readOnly: false, busy: false, literal: false, specialSelection: false }),
    read: async (range) => {
      reads.push(range);
      const b = rangeBounds(range);
      return Object.fromEntries(
        Object.entries(data).filter(([ref]) => {
          const p = rangeBounds(ref).first;
          return p.row >= b.first.row && p.row <= b.last.row && p.col >= b.first.col && p.col <= b.last.col;
        })
      );
    },
    inspect: async () => ({ reason: null }),
    accept: async (commands) => accepted.push(commands),
    ...patch,
  };
  return {
    context,
    data,
    reads,
    accepted,
    retire: () => {
      current = false;
    },
    replace: () => {
      current = false;
      runtime = false;
    },
  };
}
async function mount(t, context) {
  const root = createRoot(document.getElementById('root'));
  let properties = { context, visible: true };
  await act(async () => root.render(React.createElement(ui.WorkbookTextColumns, properties)));
  t.after(async () => act(async () => root.unmount()));
  return {
    render: async (patch) => {
      properties = { ...properties, ...patch };
      await act(async () => root.render(React.createElement(ui.WorkbookTextColumns, properties)));
      await flush();
    },
  };
}
async function initial() {
  await change(field('Delimiter'), 'comma');
  await change(field('Destination start'), 'C1');
  await click('Preview split');
}
const dateSerial = (year, month, day, date1904) =>
  Math.floor(
    (Date.UTC(year, month - 1, day) - Date.UTC(date1904 ? 1904 : 1899, date1904 ? 0 : 11, date1904 ? 1 : 31)) / 86400000
  ) + (date1904 ? 0 : 1);
check(
  'default Text preserves leading zeros/formula-looking/unsafe raw text and explicitly empty/padded output',
  async () => {
    const h = fixture();
    h.data.A1 = text('00123,=SUM(A1),9007199254740993');
    h.data.A2 = text('x,');
    const p = await ui.previewTextColumns(h.context, options);
    assert.deepEqual(
      p.outputs.map((o) => [o.value, o.conversion.kind, o.padded]),
      [
        ['00123', 'text', false],
        ['=SUM(A1)', 'text', false],
        ['9007199254740993', 'text', false],
        ['x', 'text', false],
        ['', 'text', false],
        ['', 'text', true],
      ]
    );
    assert.equal(p.rejections, 0);
    assert(p.commands[0].cells.every((c) => typeof c.value === 'string' && !Object.hasOwn(c, 'formula')));
    assert.deepEqual(h.accepted, []);
  }
);
check(
  'explicit mixed policies produce exact primitives/date format under both captured calendars in one guarded batch',
  async () => {
    for (const date1904 of [false, true]) {
      const h = fixture({ date1904 }),
        p = await ui.previewTextColumns(h.context, { ...options, policies: ['text', 'number', 'boolean', 'iso-date'] });
      assert.equal(p.rejections, 0);
      assert.deepEqual(
        p.commands[0].cells.map((c) => c.value),
        ['00123', 12.5, true, dateSerial(2024, 2, 29, date1904), '=1+2', 9.25, false, '']
      );
      assert.equal(p.commands[0].cells[3].numberFormat, 'date');
      assert.equal(p.commands[0].cells[7].numberFormat, undefined);
      assert(p.outputs[1].conversion.notices.some((n) => n.includes('12.50')));
      const result = await ui.applyTextColumns(h.context, p);
      assert.equal(result.outcome, 'accepted');
      assert.equal(h.accepted.length, 1);
      assert.equal(h.accepted[0].length, 1);
      assert.equal(h.accepted[0][0].cells[1].expectedFormula, null);
    }
  }
);
check(
  'all rejected fields remain visible in a model preview and block the whole action without partial commands',
  async () => {
    const h = fixture();
    h.data.A1 = text('9007199254740993,1,2026-02-29');
    h.data.A2 = text('1.2300000000000001,FALSE,1900-02-29');
    const p = await ui.previewTextColumns(h.context, { ...options, policies: ['number', 'boolean', 'iso-date'] });
    assert.equal(p.rejections, 5);
    assert.deepEqual(p.commands, []);
    assert.equal(p.outputs[4].conversion.status, 'converted');
    await assert.rejects(ui.applyTextColumns(h.context, p), /rejected conversions/);
    assert.deepEqual(h.accepted, []);
    const unknown = fixture({ date1904: undefined });
    const rejected = await ui.previewTextColumns(unknown.context, {
      ...options,
      policies: ['text', 'number', 'boolean', 'iso-date'],
    });
    assert.equal(rejected.rejections, 1);
  }
);
check(
  'closed policies reject extra fields/holes/unsupported types and stale calendar/source/destination before acceptance',
  async () => {
    const h = fixture(),
      typed = { ...options, policies: ['text', 'number', 'boolean', 'iso-date'] },
      p = await ui.previewTextColumns(h.context, typed);
    for (const bad of [
      { ...options, auto: true },
      { ...options, policies: ['general'] },
      { ...options, policies: Array(2) },
      { ...options, policies: ['text', 'text', 'text', 'text', 'text'] },
    ])
      await assert.rejects(ui.previewTextColumns(h.context, bad));
    h.context.date1904 = true;
    await assert.rejects(ui.applyTextColumns(h.context, p), ui.TextColumnsRetired);
    h.context.date1904 = false;
    h.data.A1 = text('00123,1,TRUE,2024-02-29');
    await assert.rejects(ui.applyTextColumns(h.context, p), /changed/);
    assert.deepEqual(h.accepted, []);
  }
);
check(
  'actual field-type edits retire preview while keeping remaining field controls usable and typed labels/notices exact',
  async (t) => {
    const h = fixture();
    await mount(t, h.context);
    await initial();
    assert.equal(field('Output value C1').textContent, '"00123"');
    await change(field('Field 2 type'), 'number');
    assert.equal(field('Split preview'), null);
    assert(field('Field 3 type'));
    await change(field('Field 3 type'), 'boolean');
    assert(field('Field 4 type'));
    await change(field('Field 4 type'), 'iso-date');
    await click('Preview split');
    assert.equal(field('Output type D1').textContent, 'number');
    assert.equal(field('Output value D1').textContent, '12.5');
    assert.equal(field('Output value E1').textContent, 'true');
    assert.equal(field('Output type F1').textContent, 'date');
    assert.equal(field('Output type F2').textContent, 'text');
    assert.match(field('Text to Columns').textContent, /Numeric spelling changes: 12.50/);
    assert.equal(field('Workbook date calendar').textContent, 'Date calendar: 1900');
    await click('Split into C1:F2');
    assert.equal(h.accepted.length, 1);
  }
);
check(
  'actual rejected preview exposes each source/destination issue, disables Apply, then a Text choice recovers without mutation',
  async (t) => {
    const h = fixture();
    h.data.A1 = text('9007199254740993,2026-02-29');
    h.data.A2 = text('1.2300000000000001,1900-02-29');
    await mount(t, h.context);
    await initial();
    await change(field('Field 1 type'), 'number');
    await change(field('Field 2 type'), 'iso-date');
    await click('Preview split');
    assert.match(field('Rejected conversions').textContent, /4 rejected/);
    for (const ref of ['C1', 'D1', 'C2', 'D2']) assert(field('Conversion issue ' + ref));
    assert(button('Split into C1:D2').disabled);
    const staleApply = props(button('Split into C1:D2')).onClick;
    await act(async () => staleApply());
    assert.deepEqual(h.accepted, []);
    await change(field('Field 1 type'), 'text');
    await change(field('Field 2 type'), 'text');
    await click('Preview split');
    assert.equal(field('Rejected conversions'), null);
    assert.equal(button('Split into C1:D2').disabled, false);
    assert.deepEqual(h.accepted, []);
  }
);
check(
  'completed Cancel restores Text defaults; pending Cancel keeps one physical read and masks its late reply',
  async (t) => {
    const h = fixture();
    const m = await mount(t, h.context);
    await initial();
    await change(field('Field 2 type'), 'number');
    await click('Preview split');
    await click('Cancel split');
    assert.equal(field('Split preview'), null);
    assert.equal(field('Field 2 type').value, 'text');
    assert.deepEqual(h.accepted, []);
    const gate = defer(),
      held = fixture({ read: () => gate.promise });
    await m.render({ context: held.context });
    await click('Preview split');
    await click('Cancel split');
    assert(button('Preview split').disabled);
    await act(async () => gate.resolve({ A1: text('x,y'), A2: text('z,w') }));
    await flush();
    assert.equal(field('Split preview'), null);
    assert.deepEqual(held.accepted, []);
  }
);
check(
  'head/owner/calendar and hide retire a held reply and old policy handlers cannot alter a new context',
  async (t) => {
    const h = fixture();
    const m = await mount(t, h.context);
    await initial();
    const old = props(field('Field 2 type')).onChange;
    h.retire();
    await m.render({ context: h.context });
    await act(async () => old({ target: { value: 'number' } }));
    assert.equal(field('Field 2 type').value, 'text');
    const gate = defer(),
      held = fixture({ read: () => gate.promise });
    await m.render({ context: held.context });
    await click('Preview split');
    const next = fixture({ binding: { ...binding(), ownerIdentity: 'new-owner', state: 'new-head' }, date1904: true });
    await m.render({ context: next.context });
    await act(async () => gate.resolve({ A1: text('old,source'), A2: text('old,source') }));
    await flush();
    assert.equal(field('Split preview'), null);
    assert.deepEqual(next.accepted, []);
    assert.equal(field('Workbook date calendar').textContent, 'Date calendar: 1904');
  }
);
check(
  'one strict acknowledgement blocks duplicate Apply and keeps a durable same-runtime success truthful',
  async (t) => {
    const gate = defer(),
      h = fixture({
        accept: (commands) => {
          h.accepted.push(commands);
          return gate.promise;
        },
      });
    const m = await mount(t, h.context);
    await initial();
    await change(field('Field 2 type'), 'number');
    await click('Preview split');
    const submit = props(button('Split into C1:F2')).onClick;
    await act(async () => {
      submit();
      submit();
    });
    await flush();
    assert.equal(h.accepted.length, 1);
    h.retire();
    await m.render({ context: h.context });
    await act(async () => gate.resolve());
    await flush();
    assert.match(field('Text to Columns').textContent, /Split accepted/);
    assert.deepEqual(
      h.accepted[0][0].cells.map((c) => typeof c.value),
      ['string', 'number', 'string', 'string', 'string', 'number', 'string', 'string']
    );
  }
);
check(
  'readonly/missing acknowledgement/protected target and cancelled revalidation cannot emit typed writes',
  async () => {
    const h = fixture(),
      p = await ui.previewTextColumns(h.context, { ...options, policies: ['text', 'number', 'boolean', 'iso-date'] });
    h.context.availability = () => ({
      active: true,
      readOnly: true,
      busy: false,
      literal: false,
      specialSelection: false,
    });
    await assert.rejects(ui.applyTextColumns(h.context, p), /read-only/);
    h.context.availability = () => ({
      active: true,
      readOnly: false,
      busy: false,
      literal: false,
      specialSelection: false,
    });
    const abort = new AbortController();
    await assert.rejects(
      ui.applyTextColumns(h.context, p, { signal: abort.signal, beforeAcceptance: () => abort.abort() }),
      ui.TextColumnsCancelled
    );
    assert.deepEqual(h.accepted, []);
    h.context.accept = () => undefined;
    await assert.rejects(ui.applyTextColumns(h.context, p), /acknowledged/);
    const protectedFixture = fixture({
      inspect: async () => ({ reason: 'Destination intersects a protected merge.' }),
    });
    await assert.rejects(ui.previewTextColumns(protectedFixture.context, options), /protected merge/);
  }
);
check(
  'actual Grid captures the accepted date calendar and retires it on metadata-only change; same-head draft survives pure workflow',
  async (t) => {
    const root = createRoot(document.getElementById('root')),
      accepted = [],
      legacy = [],
      data = { A1: text('00123,12.50,TRUE,2024-02-29') },
      sheet = {
        name: 'S',
        state: 'visible',
        path: 'xl/worksheets/sheet1.xml',
        rows: 64,
        columns: 12,
        cells: data,
        merges: [],
        widths: [],
        hiddenRows: [],
        rowHeights: [],
        date1904: false,
        freeze: { rows: 0, columns: 0 },
      },
      definition = { sheet: 'S', revision: null, canEdit: true, protectedReasons: [], rules: [] };
    let input = {
      sheets: [sheet],
      documentId: 'doc',
      ownerIdentity: 'owner',
      storageKey: 'owner:base',
      active: true,
      readOnly: false,
      nameCommands: true,
      onEdit: (edit) => legacy.push(edit),
      onBatch: (commands) => legacy.push(commands),
      onCommandBatch: async (commands) => accepted.push(commands),
      readRange: async (_name, range) => {
        const b = rangeBounds(range);
        return Object.fromEntries(
          Object.entries(data).filter(([ref]) => {
            const p = rangeBounds(ref).first;
            return p.row >= b.first.row && p.row <= b.last.row && p.col >= b.first.col && p.col <= b.last.col;
          })
        );
      },
      conditionalControls: {
        documentId: 'owner',
        revision: 'base',
        inventory: async () => definition,
        onBatch: async () => {},
      },
      validationControls: {
        documentId: 'owner',
        revision: 'base',
        inventory: async () => definition,
        onBatch: async () => {},
      },
    };
    const render = async (patch) => {
      input = { ...input, ...patch };
      await act(async () => root.render(React.createElement(ui.SpreadsheetGrid, input)));
      await flush();
    };
    await render();
    t.after(async () => act(async () => root.unmount()));
    await click([...document.querySelectorAll('[role="tab"]')].find((n) => n.textContent === 'Data'));
    const formula = field('Formula or value');
    await act(async () => {
      formula.focus();
      props(formula).onChange({ target: { value: '=SUM(' } });
    });
    formula.setSelectionRange(5, 5);
    await click(document.querySelector('[data-workbook-tool="textColumns"]'));
    await initial();
    await change(field('Field 4 type'), 'iso-date');
    await click('Preview split');
    assert.equal(field('Output value F1').textContent, String(dateSerial(2024, 2, 29, false)));
    await click('Cancel split');
    await click(field('Close workbook tools'));
    await flush();
    assert.equal(formula.value, '=SUM(');
    assert.equal(formula.selectionStart, 5);
    assert.deepEqual(accepted, []);
    assert.deepEqual(legacy, []);
    await click(document.querySelector('[data-workbook-tool="textColumns"]'));
    await initial();
    await change(field('Field 4 type'), 'iso-date');
    await click('Preview split');
    await render({ sheets: [{ ...sheet, date1904: true }] });
    assert.equal(field('Split preview'), null);
    assert.match(document.querySelector('.sheet-text-columns').textContent, /changed. Reopen/);
    assert.deepEqual(accepted, []);
    assert.equal(document.querySelectorAll('.sheet-task-pane').length, 1);
  }
);
check(
  'all output fields remain configurable through10-field pages after preview retirement; zero/nonempty whitespace are distinct',
  async (t) => {
    const h = fixture();
    h.data.A1 = text(Array.from({ length: 12 }, (_, i) => String(i)).join(','));
    h.data.A2 = text('');
    await mount(t, h.context);
    await initial();
    await click('More preview fields');
    assert(field('Field 11 type'));
    await change(field('Field 11 type'), 'number');
    assert.equal(field('Split preview'), null);
    assert(field('Field 12 type'));
    await change(field('Field 12 type'), 'number');
    await click('Previous field types');
    assert(field('Field 1 type'));
    await change(field('Field 1 type'), 'number');
    await click('Preview split');
    assert.equal(field('Output value C1').textContent, '0');
    assert.equal(field('Output type C1').textContent, 'number');
    assert.equal(field('Output value C2').textContent, '""');
    assert.equal(field('Output type C2').textContent, 'text');
    assert.deepEqual(h.accepted, []);
  }
);
check(
  'admitted negative zero keeps its sign in commands and visible typed preview without JSON display normalization',
  async (t) => {
    const h = fixture();
    h.data.A1 = text('-0,0');
    h.data.A2 = text('0,-0');
    const p = await ui.previewTextColumns(h.context, { ...options, policies: ['number', 'number'] });
    assert.equal(p.rejections, 0);
    assert(Object.is(p.commands[0].cells[0].value, -0));
    assert(Object.is(p.commands[0].cells[3].value, -0));
    const changedTarget = fixture();
    changedTarget.data.C1 = { value: -0, dataType: 'number', rawValue: '0' };
    const targetPreview = await ui.previewTextColumns(changedTarget.context, options);
    changedTarget.data.C1.value = 0;
    await assert.rejects(
      ui.applyTextColumns(changedTarget.context, targetPreview, { replaceFingerprint: targetPreview.fingerprint }),
      /changed/
    );
    assert.deepEqual(changedTarget.accepted, []);
    assert(!Object.is(p.commands[0].cells[1].value, -0));
    await mount(t, h.context);
    await initial();
    await change(field('Field 1 type'), 'number');
    await change(field('Field 2 type'), 'number');
    await click('Preview split');
    assert.equal(field('Output value C1').textContent, '-0');
    assert.equal(field('Output value D1').textContent, '0');
    await click('Split into C1:D2');
    assert(Object.is(h.accepted[0][0].cells[0].value, -0));
    assert(Object.is(h.accepted[0][0].cells[3].value, -0));
  }
);
