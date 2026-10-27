import './support/scratch-dir.mjs';
/** Actual Go To Special pane/model in React/JSDOM; controlled selection callback. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
const component = fileURLToPath(new URL('../', import.meta.url)),
  output = await fs.mkdtemp(path.join(component, '.build/tests/special-selection-ui-')),
  checks = [];
const bundle = path.join(output, 'pane.mjs'),
  entry = path.join(output, 'entry.ts');
await fs.writeFile(
  entry,
  `export {WorkbookSpecialSelection} from ${JSON.stringify(path.join(component, 'ui/WorkbookSpecialSelection.tsx'))};export * from ${JSON.stringify(path.join(component, 'ui/specialSelection.ts'))};export {homeBindingMatches} from ${JSON.stringify(path.join(component, 'ui/homeCommands.ts'))};`
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
const api = await import(pathToFileURL(bundle)),
  dom = new JSDOM('<input id="parked" value="=SUM("><div id="root"></div>', { url: 'http://localhost' });
Object.assign(globalThis, {
  window: dom.window,
  document: dom.window.document,
  HTMLElement: dom.window.HTMLElement,
  Node: dom.window.Node,
  IS_REACT_ACT_ENVIRONMENT: true,
});
const { createRoot } = await import('react-dom/client');
const binding = {
  documentId: 'doc',
  ownerIdentity: 'original-owner',
  state: 'head',
  sheet: 'Data',
  selection: 'A1:A4',
  selectedCell: 'A1',
  activationEpoch: 0,
  navigationGeneration: 0,
  modeGeneration: 0,
  specialSelectionEpoch: 0,
  mode: 'normal',
  specialSelection: null,
};
const button = (name) => [...document.querySelectorAll('button')].find((n) => n.textContent === name),
  field = (name) => document.querySelector(`[aria-label=${JSON.stringify(name)}]`),
  props = (node) => node[Object.keys(node).find((k) => k.startsWith('__reactProps$'))],
  flush = () => act(async () => new Promise((resolve) => setTimeout(resolve, 0))),
  click = async (name) => {
    const n = button(name);
    assert(n);
    assert(!n.disabled);
    await act(async () => n.click());
    await flush();
    return n;
  };
const defer = () => {
  let resolve, reject;
  const promise = new Promise((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
};
async function model(b, options, count = 3) {
  const data = Object.fromEntries(
    Array.from({ length: count }, (_, i) => ['A' + (i + 1), { value: i, dataType: 'number' }])
  );
  return api.collectSpecialSelectionPreview({
    binding: b,
    options,
    geometry: { name: b.sheet, merges: [] },
    read: async (_sheet, range) => {
      const [a, last = a] = range.split(':'),
        first = Number(a.slice(1)),
        end = Number(last.slice(1));
      return {
        range,
        complete: true,
        cells: Object.fromEntries(
          Object.entries(data).filter(([ref]) => Number(ref.slice(1)) >= first && Number(ref.slice(1)) <= end)
        ),
      };
    },
    isCurrent: () => true,
  });
}
function owner(patch = {}) {
  const live = {
      binding: structuredClone(patch.binding ?? binding),
      current: true,
      runtime: true,
      previewReason: undefined,
      applyReason: undefined,
    },
    reads = [],
    accepted = [],
    signals = [];
  const context = {
    key: 'special-original',
    binding: structuredClone(binding),
    caption: 'Data — original range',
    isCurrent: () => live.current && api.homeBindingMatches(context.binding, live.binding),
    isRuntimeCurrent: () => live.runtime,
    disabledReason: (action) => (action === 'preview' ? live.previewReason : live.applyReason),
    preview: async (options, signal) => {
      reads.push(structuredClone(options));
      signals.push(signal);
      return model(context.binding, options);
    },
    apply: async (value) => {
      accepted.push(structuredClone(value));
    },
    ...patch,
  };
  return { context, live, reads, accepted, signals };
}
async function mount(t, o = owner()) {
  const root = createRoot(document.getElementById('root'));
  let input = { context: o.context, visible: true },
    mounted = true;
  const render = async (patch) => {
      input = { ...input, ...patch };
      await act(async () => root.render(React.createElement(api.WorkbookSpecialSelection, input)));
      await flush();
    },
    unmount = async () => {
      if (!mounted) return;
      mounted = false;
      await act(async () => root.unmount());
      await flush();
    };
  await render({});
  t.after(unmount);
  return { o, render, unmount };
}
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

check('Opening and Cancel are pure; an unrelated parked formula draft/caret remains exact', async (t) => {
  const o = owner(),
    parked = document.getElementById('parked');
  parked.focus();
  parked.setSelectionRange(5, 5);
  await mount(t, o);
  assert.deepEqual(o.reads, []);
  await click('Cancel selection');
  assert.deepEqual(o.accepted, []);
  assert.equal(parked.value, '=SUM(');
  assert.equal(parked.selectionStart, 5);
  assert.equal(parked.selectionEnd, 5);
  assert.equal(document.activeElement, parked);
});
check('Explicit Preview inspects references; one selection-only Apply uses the captured immutable plan', async (t) => {
  const o = owner();
  await mount(t, o);
  await click('Preview matching cells');
  assert.match(field('Special selection summary').textContent, /Source positions4Matching cells3/);
  assert.deepEqual(o.accepted, []);
  await click('Select matching cells');
  assert.equal(o.accepted.length, 1);
  assert.deepEqual(o.accepted[0].references, ['A1', 'A2', 'A3']);
  assert.equal(o.accepted[0].binding.ownerIdentity, 'original-owner');
  assert.equal(o.accepted[0].commands, undefined);
  assert.match(document.body.textContent, /values and history stay unchanged/);
  assert(button('Preview matching cells').disabled);
});
check('No-match Preview leaves Apply disabled and retained handlers cannot accept an empty selection', async (t) => {
  const o = owner({ preview: (options) => model(binding, options, 0) });
  await mount(t, o);
  await click('Preview matching cells');
  assert(button('Select matching cells').disabled);
  await act(async () => props(button('Select matching cells')).onClick());
  assert.deepEqual(o.accepted, []);
  assert.match(document.body.textContent, /current selection stays unchanged/);
});
check('Kinds and masks retire prior proposals; no enabled types cannot Preview', async (t) => {
  const o = owner();
  await mount(t, o);
  await click('Preview matching cells');
  await act(async () => props(field('Numbers')).onChange({ target: { checked: false } }));
  await flush();
  assert.equal(field('Special selection targets'), null);
  for (const name of ['Text', 'Logicals', 'Errors'])
    await act(async () => props(field(name)).onChange({ target: { checked: false } }));
  await flush();
  assert(button('Preview matching cells').disabled);
  await act(async () => props(button('Preview matching cells')).onClick());
  assert.equal(o.reads.length, 1);
  assert.deepEqual(o.accepted, []);
  assert.match(document.body.textContent, /Choose at least one/);
  assert([...field('Selection type').options].find((v) => v.value === 'notes').disabled);
});
check('1004 captured targets page in100 rows with no additional owner Preview calls', async (t) => {
  const b = { ...binding, selection: 'A1:A1004' },
    o = owner({ binding: b, preview: (options) => model(b, options, 1004) });
  await mount(t, o);
  await click('Preview matching cells');
  let total = 0;
  do {
    total += field('Special selection targets').querySelectorAll('tbody tr').length;
    if (button('More matching cells').disabled) break;
    await click('More matching cells');
  } while (true);
  assert.equal(total, 1004);
  assert.equal(field('Special selection cells shown').textContent, '1001–1004 of 1004');
  assert.deepEqual(o.accepted, []);
});
check('Cancel aborts a held Preview and joins its physical read before another Preview', async (t) => {
  const held = defer(),
    o = owner();
  o.context.preview = (options, signal) => {
    o.reads.push(options);
    o.signals.push(signal);
    return o.reads.length === 1 ? held.promise : model(binding, options);
  };
  await mount(t, o);
  await click('Preview matching cells');
  await click('Cancel selection');
  assert(o.signals[0].aborted);
  assert(button('Preview matching cells').disabled);
  await act(async () => props(button('Preview matching cells')).onClick());
  assert.equal(o.reads.length, 1);
  await act(async () => held.resolve(await model(binding, api.DEFAULT_SPECIAL_SELECTION_OPTIONS)));
  await flush();
  assert.equal(field('Special selection targets'), null);
  await click('Preview matching cells');
  assert.equal(o.reads.length, 2);
  assert.deepEqual(o.accepted, []);
});
check('Head, owner, selection, navigation and visibility retirement mask old success/failure', async (t) => {
  for (const change of [
    { state: 'new-head' },
    { ownerIdentity: 'new-owner' },
    { selection: 'B1:B4', selectedCell: 'B1' },
    { navigationGeneration: 1 },
    { activationEpoch: 1 },
    { mode: 'pageLayout', modeGeneration: 1 },
  ]) {
    const held = defer(),
      o = owner({ preview: () => held.promise }),
      h = await mount(t, o);
    await click('Preview matching cells');
    o.live.binding = { ...o.live.binding, ...change };
    await h.render({});
    await act(async () => held.resolve(await model(binding, api.DEFAULT_SPECIAL_SELECTION_OPTIONS)));
    await flush();
    assert.equal(field('Special selection targets'), null);
    assert.deepEqual(o.accepted, []);
    await h.unmount();
  }
  const held = defer(),
    o = owner({
      preview: (_options, signal) => {
        o.signals.push(signal);
        return held.promise;
      },
    }),
    h = await mount(t, o);
  await click('Preview matching cells');
  await h.render({ visible: false });
  assert(o.signals[0].aborted);
  await act(async () => held.reject(Error('old target failure')));
  await flush();
  assert.equal(document.querySelector('[role=alert]'), null);
  assert.deepEqual(o.accepted, []);
});
check('Selection-only acknowledgement may retire its original binding while staying in the same runtime', async (t) => {
  const held = defer(),
    o = owner({
      apply: async (value) => {
        o.accepted.push(value);
        await held.promise;
      },
    }),
    h = await mount(t, o);
  await click('Preview matching cells');
  await click('Select matching cells');
  assert.equal(o.accepted.length, 1);
  o.live.binding = { ...o.live.binding, specialSelection: ['A1', 'A2', 'A3'], specialSelectionEpoch: 1 };
  await h.render({});
  await act(async () => held.resolve());
  await flush();
  assert.match(document.body.textContent, /Matching cells selected/);
  assert.equal(o.accepted[0].binding.specialSelection, null);
  assert.equal(o.accepted[0].commands, undefined);
});
check(
  'Current capability, producer mutation and foreign plans cannot redirect or fabricate selection acceptance',
  async (t) => {
    const snapshot = structuredClone(await model(binding, api.DEFAULT_SPECIAL_SELECTION_OPTIONS)),
      o = owner({ preview: async () => snapshot }),
      h = await mount(t, o);
    await click('Preview matching cells');
    snapshot.references[0] = 'B99';
    snapshot.rows[0].cell = 'B99';
    o.live.applyReason = 'Wait for the current operation.';
    await h.render({});
    assert(button('Select matching cells').disabled);
    await act(async () => props(button('Select matching cells')).onClick());
    assert.deepEqual(o.accepted, []);
    o.live.applyReason = undefined;
    await h.render({});
    await click('Select matching cells');
    assert.deepEqual(o.accepted[0].references, ['A1', 'A2', 'A3']);
    await h.unmount();
    const foreign = owner({ preview: (options) => model({ ...binding, documentId: 'foreign' }, options) });
    await mount(t, foreign);
    await click('Preview matching cells');
    assert.match(document.querySelector('[role=alert]').textContent, /another workbook/);
    assert.equal(button('Select matching cells'), undefined);
    assert.deepEqual(foreign.accepted, []);
  }
);
