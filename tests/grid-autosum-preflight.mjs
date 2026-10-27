import './support/scratch-dir.mjs';
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';

const component = fileURLToPath(new URL('../', import.meta.url));
const prefix = path.join(component, '.build/tests/grid-autosum-preflight-');
await fs.mkdir(path.dirname(prefix), { recursive: true });
const directory = await fs.mkdtemp(prefix),
  entry = path.join(directory, 'entry.ts'),
  module = path.join(directory, 'grid.mjs');
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
  outfile: module,
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
      return 400;
    },
  },
  clientHeight: {
    get() {
      return 240;
    },
  },
});
globalThis.ResizeObserver = class {
  constructor(callback) {
    this.callback = callback;
  }
  observe() {
    this.callback([{ contentRect: { width: 400, height: 240 } }]);
  }
  disconnect() {}
};
let frames = [];
globalThis.requestAnimationFrame = (callback) => {
  frames.push(callback);
  return frames.length;
};
globalThis.cancelAnimationFrame = (id) => {
  frames[id - 1] = null;
};
const { SpreadsheetGrid, configureXlsxHost } = await import(pathToFileURL(module));
configureXlsxHost({
  api: { copyText: async () => {}, readClipboardText: async () => '' },
  Icon: () => React.createElement('svg'),
  useI18n: () => ({ locale: 'en' }),
});
const { createRoot } = await import('react-dom/client');
after(async () => {
  dom.window.close();
  for (const key of [
    'window',
    'document',
    'HTMLElement',
    'HTMLDetailsElement',
    'Node',
    'ResizeObserver',
    'requestAnimationFrame',
    'cancelAnimationFrame',
    'IS_REACT_ACT_ENVIRONMENT',
  ])
    delete globalThis[key];
});
const deferred = () => {
  let resolve, reject;
  const promise = new Promise((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
};
const column = (index) => {
  let value = '';
  for (let next = index + 1; next; next = Math.floor((next - 1) / 26))
    value = String.fromCharCode(65 + ((next - 1) % 26)) + value;
  return value;
};
const position = (ref) => {
  const [, letters, row] = /^([A-Z]+)(\d+)$/.exec(ref);
  return { row: Number(row), col: [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) };
};
const project = (cells, range) => {
  const [first, last = first] = range.split(':'),
    a = position(first),
    b = position(last);
  return Object.fromEntries(
    Object.entries(cells).filter(([ref]) => {
      const p = position(ref);
      return p.row >= a.row && p.row <= b.row && p.col >= a.col && p.col <= b.col;
    })
  );
};
const metadata = () => ({
  name: 'Sheet1',
  state: 'visible',
  path: 'xl/worksheets/sheet1.xml',
  rows: 64,
  columns: 12,
  cells: {},
  merges: [],
  widths: [],
  rowHeights: [],
  hiddenRows: [],
  freeze: { rows: 0, columns: 0 },
});
const visible = (node) => !node.closest('[hidden]');
const button = (name) =>
  [...document.querySelectorAll('button')].find(
    (node) => visible(node) && (node.getAttribute('aria-label') ?? node.textContent) === name
  );
const tab = (name) => [...document.querySelectorAll('[role="tab"]')].find((node) => node.textContent === name);
async function click(node) {
  assert(node);
  assert.equal(node.disabled, false);
  await act(async () => node.click());
}
async function mount(t, extra = {}) {
  frames = [];
  window.localStorage.clear();
  const root = createRoot(document.getElementById('root')),
    sheets = [metadata()],
    props = { sheets, documentId: 'A', storageKey: 'owner-A:base', readOnly: false, onEdit: () => {}, ...extra };
  await act(async () => root.render(React.createElement(SpreadsheetGrid, props)));
  t.after(async () => act(async () => root.unmount()));
  return {
    sheets,
    render: async (patch) => act(async () => root.render(React.createElement(SpreadsheetGrid, { ...props, ...patch }))),
  };
}
async function go(range) {
  const input = document.querySelector('input[aria-label="Go to cell"]');
  await act(async () => {
    Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(input, range);
    input.dispatchEvent(new window.Event('input', { bubbles: true }));
  });
  await act(async () =>
    input.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
  );
}
async function autosum(range) {
  await go(range);
  await click(tab('Formulas'));
  await click(button('AutoSum'));
}
const source = () => ({ G1: { value: 1 }, G2: { value: 2 } });

test('actual AutoSum reads delayed next-tile G33 before writing and refuses existing933 without adopting the reply', async (t) => {
  const gate = deferred(),
    reads = [],
    writes = [],
    cells = { ...source(), G33: { value: 933 } },
    before = JSON.stringify(cells);
  const ui = await mount(t, {
    readRange: async (sheet, range) => {
      reads.push({ sheet, range });
      return range === 'G33:G33' ? gate.promise : project(cells, range);
    },
    onCommandBatch: async (commands) => {
      writes.push(commands);
    },
  });
  await autosum('G1:G32');
  assert.equal(reads.filter((item) => item.range === 'G33:G33').length, 1);
  assert.equal(writes.length, 0);
  assert(!reads.some((item) => item.range === 'A33:AF64'));
  await act(async () => gate.resolve({ G33: cells.G33 }));
  assert.equal(writes.length, 0);
  assert.match(document.body.textContent, /AutoSum would overwrite G33/);
  assert.deepEqual(ui.sheets[0].cells, {});
  assert.equal(JSON.stringify(cells), before);
});
test('zero false empty text formula and formulaType targets remain occupied after a bounded owner read', async (t) => {
  for (const [name, target] of [
    ['zero', { value: 0 }],
    ['false', { value: false }],
    ['empty text', { value: '' }],
    ['formula', { value: null, formula: 'SUM(A1:A2)' }],
    ['special formula', { value: null, formulaType: 'array' }],
  ])
    await t.test(name, async (child) => {
      let writes = 0;
      const cells = { ...source(), G33: target };
      await mount(child, {
        readRange: async (_sheet, range) => project(cells, range),
        onCommandBatch: async () => {
          writes++;
        },
      });
      await autosum('G1:G32');
      assert.equal(writes, 0);
      assert.match(document.body.textContent, /AutoSum would overwrite G33/);
    });
});
test('blank/style-only output accepts original captured callback after delayed read and harmless rerender', async (t) => {
  const gate = deferred(),
    original = [],
    replacement = [];
  const ui = await mount(t, {
    readRange: async (_sheet, range) => (range === 'G33:G33' ? gate.promise : project(source(), range)),
    onCommandBatch: async (commands) => {
      original.push(commands);
    },
  });
  await autosum('G1:G32');
  assert.equal(original.length, 0);
  await ui.render({
    onCommandBatch: async (commands) => {
      replacement.push(commands);
    },
  });
  await act(async () => gate.resolve({ G33: { value: null, styleId: 2 } }));
  assert.deepEqual(original, [
    [{ type: 'setCells', sheet: 'Sheet1', cells: [{ cell: 'G33', formula: 'SUM(G1:G32)' }] }],
  ]);
  assert.equal(replacement.length, 0);
  assert(!document.body.textContent.includes('AutoSum would overwrite'));
  assert.deepEqual(ui.sheets[0].cells, {});
});
test('changed accepted head or selection refuses a pending preflight without a write or late active error', async (t) => {
  for (const change of ['head', 'selection', 'activation'])
    await t.test(change, async (child) => {
      const gate = deferred();
      let writes = 0;
      const ui = await mount(child, {
        readRange: async (_sheet, range) => (range === 'G33:G33' ? gate.promise : project(source(), range)),
        onCommandBatch: async () => {
          writes++;
        },
      });
      await autosum('G1:G32');
      if (change === 'head') await ui.render({ storageKey: 'owner-A:changed' });
      else if (change === 'selection') await go('A1:A2');
      else {
        await ui.render({ active: false });
        await ui.render({ active: true });
      }
      await act(async () => gate.resolve({}));
      assert.equal(writes, 0);
      assert(!document.body.textContent.includes('before AutoSum could be applied'));
    });
});
test('output row reads are sequential bounded1000-position chunks and normalized totals form one accepted batch', async (t) => {
  const gates = [deferred(), deferred(), deferred()],
    targets = [],
    writes = [];
  await mount(t, {
    sheets: [{ ...metadata(), columns: 2001 }],
    readRange: async (_sheet, range) => {
      const [a, b] = range.split(':').map(position);
      if (a.row === 3 && b.row === 3) {
        targets.push(range);
        return gates[targets.length - 1].promise;
      }
      return {};
    },
    onCommandBatch: async (commands) => {
      writes.push(commands);
    },
  });
  await autosum('A1:' + column(2000) + '2');
  assert.deepEqual(targets, ['A3:' + column(999) + '3']);
  assert.equal(writes.length, 0);
  await act(async () => gates[0].resolve({}));
  assert.deepEqual(targets, ['A3:' + column(999) + '3', column(1000) + '3:' + column(1999) + '3']);
  assert.equal(writes.length, 0);
  await act(async () => gates[1].resolve({}));
  assert.equal(targets.at(-1), column(2000) + '3:' + column(2000) + '3');
  assert.equal(writes.length, 0);
  await act(async () => gates[2].resolve({}));
  assert.equal(writes.length, 1);
  assert.deepEqual(
    writes[0].map((command) => command.cells.length),
    [1000, 1000, 1]
  );
  assert.equal(writes[0].at(-1).cells[0].formula, 'SUM(' + column(2000) + '1:' + column(2000) + '2)');
});
test('read failures surface through current command scope and cannot reach the owner', async (t) => {
  const gate = deferred();
  let writes = 0;
  await mount(t, {
    readRange: async (_sheet, range) => (range === 'G33:G33' ? gate.promise : project(source(), range)),
    onCommandBatch: async () => {
      writes++;
    },
  });
  await autosum('G1:G32');
  await act(async () => gate.reject(Error('Target row could not be read')));
  assert.equal(writes, 0);
  assert.match(document.body.textContent, /Target row could not be read/);
});
test('generic full-sheet model without range callback protects occupied totals too', async (t) => {
  let writes = 0;
  await mount(t, {
    sheets: [{ ...metadata(), cells: { ...source(), G33: { value: false } } }],
    onCommandBatch: async () => {
      writes++;
    },
  });
  await autosum('G1:G32');
  assert.equal(writes, 0);
  assert.match(document.body.textContent, /AutoSum would overwrite G33/);
});

test('selection and worksheet away then back cannot revive the original pending AutoSum', async (t) => {
  for (const change of ['selection-roundtrip', 'worksheet-roundtrip'])
    await t.test(change, async (child) => {
      const gate = deferred();
      let writes = 0;
      await mount(child, {
        sheets: [metadata(), { ...metadata(), name: 'Other', path: 'xl/worksheets/sheet2.xml' }],
        readRange: async (sheet, range) =>
          sheet === 'Sheet1' && range === 'G33:G33' ? gate.promise : project(source(), range),
        onCommandBatch: async () => {
          writes++;
        },
      });
      await autosum('G1:G32');
      if (change === 'selection-roundtrip') await go('A1:A2');
      else {
        await click([...document.querySelectorAll('.sheet-tabs button')].find((node) => node.textContent === 'Other'));
        await click([...document.querySelectorAll('.sheet-tabs button')].find((node) => node.textContent === 'Sheet1'));
      }
      await go('G1:G32');
      assert.match(document.querySelector('.sheet-a11y-status').textContent, /Range G1:G32 selected/);
      await act(async () => gate.resolve({}));
      assert.equal(writes, 0);
    });
});
test('previewed Go to Special blanks retires its pending AutoSum without a write', async (t) => {
  const gate = deferred();
  let writes = 0;
  await mount(t, {
    ownerIdentity: 'owner-A',
    readRange: async (_sheet, range) => (range === 'G33:G33' ? gate.promise : project(source(), range)),
    onCommandBatch: async () => {
      writes++;
    },
  });
  await autosum('G1:G32');
  await click(tab('Home'));
  await click(button('Editing'));
  const menu = [...document.querySelectorAll('details')].find(
    (node) => visible(node) && node.textContent.includes('Find & Select')
  );
  assert(menu);
  menu.open = true;
  await click(button('Select blanks'));
  await click(button('Preview matching cells'));
  for (let n = 0; n < 50 && !button('Select matching cells'); n++)
    await act(async () => {
      const ready = frames;
      frames = [];
      ready.forEach((fn) => fn?.());
      await Promise.resolve();
    });
  assert(button('Select matching cells'), document.querySelector('.sheet-task-pane')?.textContent);
  await click(button('Select matching cells'));
  for (let n = 0; n < 50 && !document.querySelector('.sheet-special-status'); n++)
    await act(async () => {
      const ready = frames;
      frames = [];
      ready.forEach((fn) => fn?.());
      await Promise.resolve();
    });
  assert.match(document.querySelector('.sheet-special-status').textContent, /matching cells selected/);
  const nameBox = document.querySelector('input[aria-label="Go to cell"]');
  assert.equal(nameBox.value, '');
  assert.equal(nameBox.placeholder, 'G3');
  assert.match(document.querySelector('.sheet-special-status').textContent, /30 matching cells selected/);
  await act(async () => gate.resolve({}));
  assert.equal(writes, 0);
});
