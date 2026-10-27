import './support/scratch-dir.mjs';
/** Production-source Financial arithmetic picker React/JSDOM; no actual calculation claim. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
const folder = fileURLToPath(new URL('../', import.meta.url)),
  out = await fs.mkdtemp(path.join(folder, '.build/tests/financial-expression-picker-'));
const bundle = path.join(out, 'picker.mjs');
await build({
  entryPoints: [path.join(folder, 'ui/WorkbookFunctions.tsx')],
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
  Node: dom.window.Node,
  IS_REACT_ACT_ENVIRONMENT: true,
});
const { createRoot } = await import('react-dom/client'),
  { WorkbookFunctions } = await import(pathToFileURL(bundle));
after(async () => {
  dom.window.close();
});
const field = (name) => document.querySelector(`[aria-label="${name}"]`),
  button = (name) => [...document.querySelectorAll('button')].find((b) => b.textContent === name),
  preview = () => field('Formula preview').textContent;
async function select(name, value) {
  const f = field(name);
  assert(f, name);
  await act(async () => {
    f.value = value;
    f.dispatchEvent(new window.Event('change', { bubbles: true }));
  });
}
async function input(name, value) {
  const f = field(name);
  assert(f, name);
  await act(async () => {
    Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(f, value);
    f.dispatchEvent(new window.Event('input', { bubbles: true }));
  });
}
async function numeric(name, value) {
  await select(name + ' type', 'number');
  await input(name, String(value));
}
async function click(name) {
  const b = button(name);
  assert(b, name);
  assert.equal(b.disabled, false, name);
  await act(async () => b.click());
}
async function mount(t) {
  window.localStorage.clear();
  const writes = [],
    root = createRoot(document.getElementById('root')),
    context = {
      binding: {
        documentId: 'doc-A',
        grantDocumentId: 'grant-A',
        ownerKey: 'session-A',
        ownerIdentity: 'runtime-A',
        state: 'head-A',
        sheet: 'Cases',
        target: 'C33',
        selection: 'C33',
        navigationGeneration: 0,
        activationEpoch: 0,
      },
      isCurrent: () => true,
      availability: () => ({ active: true, readOnly: false, busy: false, literal: false, specialSelection: false }),
      readTarget: async () => ({}),
      accept: async (commands) => {
        writes.push(commands);
      },
    };
  let props = { context, visible: true };
  await act(async () => root.render(React.createElement(WorkbookFunctions, props)));
  t.after(async () => act(async () => root.unmount()));
  return {
    writes,
    render: async (patch) => {
      props = { ...props, ...patch };
      await act(async () => root.render(React.createElement(WorkbookFunctions, props)));
    },
  };
}
test('populated Financial menu preserves initial SUM, actual domain/sign/default help and PMT contiguous optional fields', async (t) => {
  await mount(t);
  assert.equal(field('Function').value, 'SUM');
  await select('Function category', 'Financial');
  assert.deepEqual(
    [...field('Function').options].map((o) => o.value),
    ['PMT', 'PV', 'FV']
  );
  assert.match(document.body.textContent, /0\.005/);
  assert.match(document.body.textContent, /ordinary numeric input cell/);
  assert.match(document.body.textContent, /Results can differ from Excel in the last decimal places/);
  await numeric('Period rate', 0);
  await numeric('Number of periods', 12);
  await numeric('Present value', 1200);
  assert.equal(preview(), '=PMT(0,12,1200)');
  await click('Add argument');
  assert.equal(preview(), '=PMT(0,12,1200)');
  await click('Add argument');
  await numeric('Payment timing (optional; 0 = end, 1 = beginning)', 1);
  assert.equal(
    button('Insert formula into C33').disabled,
    true,
    'empty middle balance cannot be omitted before filled timing'
  );
  await numeric('Future value (optional; default 0)', 0);
  assert.equal(preview(), '=PMT(0,12,1200,0,1)');
  assert.equal(button('Insert formula into C33').disabled, false);
  assert.doesNotMatch(
    document.body.textContent,
    /qualified.*(?:helper|calculator)|retained21|exact-bit|owner proof|compatibility route|\.build\//i
  );
});
test('Financial PV/FV cash labels differ correctly; one accepted FV emits one original-target batch', async (t) => {
  const { writes } = await mount(t);
  await select('Function category', 'Financial');
  await select('Function', 'PV');
  assert(field('Payment each period'));
  await click('Add argument');
  assert(field('Future value (optional; default 0)'));
  await select('Function', 'FV');
  assert(field('Payment each period'));
  await numeric('Period rate', 0);
  await numeric('Number of periods', 12);
  await numeric('Payment each period', -100);
  assert.equal(preview(), '=FV(0,12,-100)');
  await click('Add argument');
  assert(field('Present value (optional; default 0)'));
  await numeric('Present value (optional; default 0)', 1200);
  assert.equal(preview(), '=FV(0,12,-100,1200)');
  await click('Insert formula into C33');
  assert.deepEqual(writes, [
    [{ type: 'setCells', sheet: 'Cases', cells: [{ cell: 'C33', formula: 'FV(0,12,-100,1200)' }] }],
  ]);
  assert.match(
    document.body.textContent,
    /Result unavailable/,
    'no mocked numeric result is passed off as owner calculation'
  );
  assert.deepEqual(JSON.parse(window.localStorage.getItem('xlsx.function-recents.v1')).ids, ['FV']);
});
test('typed Financial widgets refuse text coercion, ranges, raw decimal loss and type2 before callback', async (t) => {
  const { writes } = await mount(t);
  await select('Function category', 'Financial');
  assert.deepEqual(
    [...field('Period rate type').options].map((o) => o.value),
    ['number', 'reference', 'expression']
  );
  await numeric('Number of periods', 12);
  await numeric('Present value', 100);
  await input('Period rate', 'A1:A2');
  assert.equal(button('Insert formula into C33').disabled, true);
  await numeric('Period rate', '0.1234567890123456789012345');
  assert.equal(button('Insert formula into C33').disabled, true);
  await input('Period rate', '0.000000000001');
  assert.equal(preview(), '=PMT(1E-12,12,100)');
  await click('Add argument');
  await numeric('Future value (optional; default 0)', 0);
  await click('Add argument');
  await numeric('Payment timing (optional; 0 = end, 1 = beginning)', 2);
  assert.equal(button('Insert formula into C33').disabled, true);
  assert.equal(writes.length, 0);
  assert.equal(window.localStorage.getItem('xlsx.function-recents.v1'), null);
});

test('direct populated category routes clear stale search and select an actual first member without owner writes', async (t) => {
  const { writes, render } = await mount(t);
  await input('Search functions', 'LEFT');
  assert.equal(field('Function').value, 'LEFT');
  let serial = 0;
  for (const [mode, first] of [
    ['Financial', 'PMT'],
    ['Logical', 'IF'],
    ['Text', 'LEN'],
    ['DateTime', 'DATE'],
    ['Lookup', 'ROWS'],
    ['Math', 'SUM'],
    ['More', 'DEC2HEX'],
  ]) {
    await render({ browse: { mode, serial: ++serial } });
    assert.equal(field('Function category').value, mode);
    assert.equal(field('Search functions').value, '');
    assert.equal(field('Function').value, first);
    assert.equal(document.querySelector('[data-function-target]').getAttribute('data-function-target'), 'C33');
    assert.equal(writes.length, 0);
  }
});
test('direct category requests preserve drafts until a new explicit entry; normal SUM and Recent remain intact', async (t) => {
  const { writes, render } = await mount(t);
  const browse = { mode: 'Financial', serial: 1 };
  await render({ browse });
  await numeric('Period rate', 0);
  await numeric('Number of periods', 12);
  await numeric('Present value', 1200);
  assert.equal(preview(), '=PMT(0,12,1200)');
  await render({ browse: { ...browse } });
  assert.equal(preview(), '=PMT(0,12,1200)', 'same request rerender keeps form state');
  await render({ browse: { mode: 'Financial', serial: 2 } });
  assert.equal(field('Present value').value, '', 'explicit entry resets arguments');
  assert.equal(button('Insert formula into C33').disabled, true);
  await render({ browse: { mode: 'Logical', serial: 2 } });
  assert.equal(field('Function').value, 'IF', 'a changed mode is an explicit category even when serial is shared');
  await render({ browse: { mode: 'all', serial: 3 } });
  assert.equal(field('Function category').value, 'all');
  assert.equal(field('Function').value, 'SUM');
  window.localStorage.setItem('xlsx.function-recents.v1', JSON.stringify({ version: 1, ids: ['FV', 'IF'] }));
  await render({ browse: { mode: 'recent', serial: 4 } });
  assert.equal(field('Function category').value, 'recent');
  assert.equal(field('Function').value, 'FV');
  assert.deepEqual(
    [...field('Function').options].map((item) => item.value),
    ['FV', 'IF']
  );
  assert.equal(writes.length, 0);
});

async function expression(name, value) {
  await select(name + ' type', 'expression');
  await input(name, value);
}
test('Financial expression widgets preserve the authored mortgage inputs and one original-target acknowledged batch', async (t) => {
  const { writes, render } = await mount(t);
  await render({ browse: { mode: 'Financial', serial: 1 } });
  assert.deepEqual(
    [...field('Period rate type').options].map((o) => o.value),
    ['number', 'reference', 'expression']
  );
  await expression('Period rate', 'B1/12');
  await expression('Number of periods', 'B2*12');
  await expression('Present value', 'B3');
  assert.equal(preview(), '=PMT(B1/12,B2*12,B3)');
  assert.equal(button('Insert formula into C33').disabled, false);
  assert.match(document.body.textContent, /Choose Expression/);
  await click('Insert formula into C33');
  assert.deepEqual(writes, [
    [{ type: 'setCells', sheet: 'Cases', cells: [{ cell: 'C33', formula: 'PMT(B1/12,B2*12,B3)' }] }],
  ]);
  assert.match(
    document.body.textContent,
    /Result unavailable/,
    'a mocked acknowledgement is not a calculated numeric preview'
  );
});
test('expression optional blanks omit only at the tail and never coerce a required or middle operand to zero', async (t) => {
  const { writes, render } = await mount(t);
  await render({ browse: { mode: 'Financial', serial: 1 } });
  await expression('Period rate', 'B1/12');
  await expression('Number of periods', 'B2*12');
  await expression('Present value', 'B3');
  await click('Add argument');
  await expression('Future value (optional; default 0)', '');
  assert.equal(preview(), '=PMT(B1/12,B2*12,B3)');
  await click('Add argument');
  await expression('Payment timing (optional; 0 = end, 1 = beginning)', '1');
  assert.equal(button('Insert formula into C33').disabled, true);
  await expression('Future value (optional; default 0)', '0');
  assert.equal(preview(), '=PMT(B1/12,B2*12,B3,0,1)');
  await expression('Period rate', '');
  assert.equal(button('Insert formula into C33').disabled, true);
  assert.equal(writes.length, 0);
});
test('nested/formula/range/unsafe operands refuse in the actual form while syntax never pretends to evaluate role inputs', async (t) => {
  const { writes, render } = await mount(t);
  await render({ browse: { mode: 'Financial', serial: 1 } });
  await numeric('Number of periods', 12);
  await numeric('Present value', 100);
  await select('Period rate type', 'expression');
  for (const raw of [
    '=B1/12',
    'SUM(B1)',
    'Other!B1',
    'B1:B2',
    'B1^2',
    '0.123456789012345678901',
    '9007199254740993',
    'TRUE',
    'b1/12',
  ]) {
    await input('Period rate', raw);
    assert.equal(button('Insert formula into C33').disabled, true, raw);
  }
  await input('Period rate', 'B1/0');
  assert.equal(preview(), '=PMT(B1/0,12,100)', 'numeric/source proof is deliberately deferred to the actual owner');
  assert.equal(writes.length, 0);
  assert.equal(window.localStorage.getItem('xlsx.function-recents.v1'), null);
});
