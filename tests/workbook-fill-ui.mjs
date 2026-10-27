import './support/scratch-dir.mjs';
/** Actual Fill pane/model in React/JSDOM, with controlled owner callbacks. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
const component = fileURLToPath(new URL('../', import.meta.url)),
  output = await fs.mkdtemp(path.join(component, '.build/tests/fill-ui-')),
  checks = [];
const entry = path.join(output, 'entry.ts'),
  bundle = path.join(output, 'pane.mjs');
await fs.writeFile(
  entry,
  `export {WorkbookFill} from ${JSON.stringify(path.join(component, 'ui/WorkbookFill.tsx'))};export * from ${JSON.stringify(path.join(component, 'ui/fillSeriesPreview.ts'))};export {homeBindingMatches} from ${JSON.stringify(path.join(component, 'ui/homeCommands.ts'))};`
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
const ui = await import(pathToFileURL(bundle)),
  dom = new JSDOM('<input id="parked" value="=SUM("><div id="root"></div>', { url: 'http://localhost' });
Object.assign(globalThis, {
  window: dom.window,
  document: dom.window.document,
  HTMLElement: dom.window.HTMLElement,
  Node: dom.window.Node,
  IS_REACT_ACT_ENVIRONMENT: true,
});
const { createRoot } = await import('react-dom/client');
const props = (node) => node[Object.keys(node).find((key) => key.startsWith('__reactProps$'))],
  field = (name) => document.querySelector(`[aria-label=${JSON.stringify(name)}]`),
  button = (name) => [...document.querySelectorAll('button')].find((node) => node.textContent === name),
  flush = () => act(async () => new Promise((resolve) => setTimeout(resolve, 0)));
const click = async (name) => {
  const node = typeof name === 'string' ? button(name) : name;
  assert(node);
  assert(!node.disabled);
  await act(async () => node.click());
  await flush();
  return node;
};
const defer = () => {
  let resolve, reject;
  const promise = new Promise((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
};
const binding = {
  documentId: 'doc',
  ownerIdentity: 'owner',
  state: 'head',
  sheet: 'Report',
  selection: 'A1:D1',
  selectedCell: 'A1',
  activationEpoch: 0,
  navigationGeneration: 0,
  modeGeneration: 0,
  specialSelectionEpoch: 0,
  mode: 'normal',
  specialSelection: null,
};
function model(b = binding, seedValues = [1, 3], patch = {}) {
  const cells = {};
  const vertical = b.selection.split(':')[1]?.startsWith('A') && b.selection.endsWith('1') === false;
  const second = vertical ? 'A2' : 'B1';
  cells.A1 = {
    value: seedValues[0],
    rawValue: Object.is(seedValues[0], -0) ? '-0' : String(seedValues[0]),
    dataType: 'number',
  };
  cells[second] = {
    value: seedValues[1],
    rawValue: Object.is(seedValues[1], -0) ? '-0' : String(seedValues[1]),
    dataType: 'number',
  };
  const result = ui.createFillSeriesPreview({
    binding: b,
    sheet: {
      name: b.sheet,
      path: 'xl/worksheets/sheet1.xml',
      rows: 12000,
      columns: 26,
      cells,
      merges: [],
      widths: [],
      hiddenRows: [],
      rowHeights: [],
      freeze: { rows: 0, columns: 0 },
    },
    cells,
    readRange: b.selection,
    ...patch,
  });
  assert.equal(result.status, 'ready', result.reason);
  return result.preview;
}
function owner(patch = {}) {
  const status = {
      current: true,
      runtime: true,
      binding: structuredClone(patch.binding ?? binding),
      reason: undefined,
      applyReason: undefined,
    },
    reads = [],
    accepted = [],
    signals = [];
  let snapshot;
  const context = {
    key: 'original-scope',
    binding: { ...binding },
    caption: 'Report!A1:D1',
    isCurrent: () => status.current && ui.homeBindingMatches(context.binding, status.binding),
    isRuntimeCurrent: () => status.runtime && ui.homeBindingMatches(context.binding, status.binding, true),
    disabledReason: (action) => status.reason ?? (action === 'apply' ? status.applyReason : undefined),
    preview: async (signal) => {
      reads.push(context.binding.selection);
      signals.push(signal);
      return snapshot ?? model(context.binding);
    },
    apply: async (value) => {
      accepted.push(structuredClone(value.command));
    },
    ...patch,
  };
  return { context, status, reads, accepted, signals, setSnapshot: (value) => (snapshot = value) };
}
async function mount(t, o = owner(), patch = {}) {
  const root = createRoot(document.getElementById('root'));
  let input = { context: o.context, visible: true, ...patch },
    mounted = true;
  const render = async (patch) => {
    input = { ...input, ...patch };
    await act(async () => root.render(React.createElement(ui.WorkbookFill, input)));
    await flush();
  };
  const unmount = async () => {
    if (!mounted) return;
    mounted = false;
    await act(async () => root.unmount());
    await flush();
  };
  await render({});
  t.after(unmount);
  return { o, render, unmount, input: () => input };
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

check('opening and Cancel are pure and preserve parked formula text/caret/focus', async (t) => {
  const o = owner(),
    parked = document.getElementById('parked');
  parked.focus();
  parked.setSelectionRange(5, 5);
  await mount(t, o);
  assert(field('Fill series').hasAttribute('data-formula-draft-zone'));
  assert.match(field('Fill series').textContent, /Report!A1:D1/);
  assert.deepEqual(o.reads, []);
  assert.equal(field('Series destinations'), null);
  assert.equal(button('Apply series'), undefined);
  await click('Cancel series');
  assert.deepEqual(o.accepted, []);
  assert.equal(parked.value, '=SUM(');
  assert.equal(parked.selectionStart, 5);
  assert.equal(document.activeElement, parked);
});
check(
  'explicit horizontal Preview describes1,3→5,7 and Apply uses one original acknowledged fillSeries command',
  async (t) => {
    const o = owner();
    await mount(t, o);
    await click('Preview series');
    assert.deepEqual(o.reads, ['A1:D1']);
    assert.match(field('Series summary').textContent, /A1: 1Second valueB1: 3Step2Destination cells2RangeReport!A1:D1/);
    assert.deepEqual(
      [...field('Series destinations').querySelectorAll('tbody tr')].map((n) => n.textContent),
      ['C15', 'D17']
    );
    assert.deepEqual(o.accepted, []);
    await click('Apply series');
    assert.deepEqual(o.accepted, [{ type: 'fillSeries', sheet: 'Report', range: 'A1:D1' }]);
    assert.match(field('Fill series').textContent, /Series filled. Use Undo/);
    assert(button('Preview series').disabled);
    assert.equal(button('Apply series'), undefined);
  }
);
check(
  'vertical off-viewport range exposes every exact destination in100-row local pages without extra owner reads',
  async (t) => {
    const b = { ...binding, selection: 'A1:A1004' },
      o = owner({ binding: b, caption: 'Report!A1:A1004' });
    await mount(t, o);
    await click('Preview series');
    let rows = field('Series destinations').querySelectorAll('tbody tr');
    assert.equal(rows.length, 100);
    assert.equal(rows[0].textContent, 'A35');
    assert.equal(rows[99].textContent, 'A102203');
    assert.equal(field('Series cells shown').textContent, '1–100 of 1002');
    for (let i = 0; i < 10; i++) await click('More series cells');
    rows = field('Series destinations').querySelectorAll('tbody tr');
    assert.deepEqual(
      [...rows].map((n) => n.textContent),
      ['A10032005', 'A10042007']
    );
    assert(button('More series cells').disabled);
    await click('Previous series cells');
    assert.equal(field('Series cells shown').textContent, '901–1000 of 1002');
    assert.equal(o.reads.length, 1);
    assert.deepEqual(o.accepted, []);
  }
);
check(
  'preview snapshot is immutable, preserves signed seed display and uses native-model precision for exact rows',
  async (t) => {
    const o = owner(),
      snapshot = structuredClone(model(binding, [-0, 0.1]));
    o.setSnapshot(snapshot);
    const h = await mount(t, o);
    await click('Preview series');
    assert.match(field('Series summary').textContent, /A1: -0/);
    assert.deepEqual(
      [...field('Series destinations').querySelectorAll('tbody tr')].map((n) => n.textContent),
      ['C10.2', 'D10.3']
    );
    snapshot.destinations[0].value = 913;
    snapshot.command.range = 'A1:E1';
    await h.render({});
    assert.deepEqual(
      [...field('Series destinations').querySelectorAll('tbody tr')].map((n) => n.textContent),
      ['C10.2', 'D10.3']
    );
    await click('Apply series');
    assert.deepEqual(o.accepted, [{ type: 'fillSeries', sheet: 'Report', range: 'A1:D1' }]);
    assert.equal(Object.is(snapshot.seeds[0].value, -0), true);
  }
);
check('Cancel aborts and masks a held preview immediately but keeps one physical read until settlement', async (t) => {
  const held = defer(),
    o = owner();
  o.context.preview = (signal) => {
    o.reads.push('original');
    o.signals.push(signal);
    return o.reads.length === 1 ? held.promise : Promise.resolve(model());
  };
  await mount(t, o);
  await click('Preview series');
  await click('Cancel series');
  assert(o.signals[0].aborted);
  assert(button('Preview series').disabled);
  await act(async () => held.resolve(model()));
  await flush();
  assert.equal(field('Series destinations'), null);
  await click('Preview series');
  assert.equal(o.reads.length, 2);
  assert(field('Series destinations'));
  assert.deepEqual(o.accepted, []);
});
check(
  'held replies retire on head, owner, selection, navigation, sheet, activation and mode without retargeting',
  async (t) => {
    for (const change of [
      { state: 'next-head' },
      { ownerIdentity: 'replacement' },
      { selection: 'A1:E1' },
      { navigationGeneration: 1 },
      { sheet: 'Other' },
      { activationEpoch: 1 },
      { mode: 'pageLayout', modeGeneration: 1 },
    ]) {
      const held = defer(),
        o = owner();
      o.context.preview = () => {
        o.reads.push('original');
        return held.promise;
      };
      const h = await mount(t, o);
      await click('Preview series');
      o.status.binding = { ...o.status.binding, ...change };
      await h.render({});
      await act(async () => held.resolve(model()));
      await flush();
      assert.equal(field('Series destinations'), null);
      assert.match(field('Fill series').textContent, /changed. Reopen Fill series/);
      assert.deepEqual(o.accepted, []);
      assert.equal(o.reads.length, 1);
      await h.unmount();
    }
  }
);
check(
  'hidden pane masks a held reply and reopening waits for the old physical slot before a fresh Preview',
  async (t) => {
    const held = defer(),
      o = owner();
    o.context.preview = (signal) => {
      o.signals.push(signal);
      o.reads.push('original');
      return o.reads.length === 1 ? held.promise : Promise.resolve(model());
    };
    const h = await mount(t, o);
    await click('Preview series');
    await h.render({ visible: false });
    assert.equal(field('Series destinations'), null);
    assert(o.signals[0].aborted);
    await h.render({ visible: true });
    assert(button('Preview series').disabled);
    await act(async () => held.resolve(model()));
    await flush();
    assert.equal(field('Series destinations'), null);
    await click('Preview series');
    assert.equal(o.reads.length, 2);
    assert(field('Series destinations'));
    assert.deepEqual(o.accepted, []);
  }
);
check(
  'duplicate Apply is blocked while strict acknowledgement is held; accepted head advancement stays truthful',
  async (t) => {
    const held = defer(),
      o = owner();
    o.context.apply = (value) => {
      o.accepted.push(structuredClone(value.command));
      return held.promise;
    };
    const h = await mount(t, o);
    await click('Preview series');
    const apply = props(button('Apply series')).onClick;
    await act(async () => {
      apply();
      apply();
    });
    assert.equal(o.accepted.length, 1);
    assert(button('Cancel series').disabled);
    o.status.current = false;
    await h.render({});
    await act(async () => held.resolve());
    await flush();
    assert.match(field('Fill series').textContent, /Series filled. Use Undo/);
    assert(!field('Fill series').textContent.includes('changed. Reopen'));
    await act(async () => apply());
    assert.equal(o.accepted.length, 1);
  }
);
check(
  'Apply retains the original preview callback across same-key props replacement and ignores obsolete owner failure',
  async (t) => {
    const held = defer(),
      o = owner();
    o.context.apply = (value) => {
      o.accepted.push(structuredClone(value.command));
      return held.promise;
    };
    const h = await mount(t, o);
    await click('Preview series');
    await h.render({
      context: {
        ...o.context,
        apply: async () => {
          throw Error('Retargeted Apply');
        },
      },
    });
    await click('Apply series');
    assert.equal(o.accepted.length, 1);
    const next = owner({
      key: 'replacement',
      binding: { ...binding, ownerIdentity: 'next-owner' },
      caption: 'Next range',
    });
    o.status.current = false;
    o.status.runtime = false;
    await h.render({ context: next.context });
    await act(async () => held.reject(Error('Obsolete original acceptance failed')));
    await flush();
    assert(!document.body.textContent.includes('Obsolete original acceptance failed'));
    assert(!document.body.textContent.includes('Retargeted Apply'));
    assert.deepEqual(next.accepted, []);
    await click('Preview series');
    assert(field('Series destinations'));
  }
);
check(
  'retained old Preview/Apply/Cancel/page handlers cannot change a replacement frame or its parked draft',
  async (t) => {
    const b = { ...binding, selection: 'A1:A104' },
      o = owner({ binding: b }),
      h = await mount(t, o);
    const oldPreview = props(button('Preview series')).onClick;
    await click('Preview series');
    const oldApply = props(button('Apply series')).onClick,
      oldCancel = props(button('Cancel series')).onClick,
      oldPage = props(button('More series cells')).onClick;
    const next = owner({ key: 'replacement', binding: { ...b, ownerIdentity: 'new-owner' } });
    o.status.current = false;
    o.status.runtime = false;
    await h.render({ context: next.context });
    await click('Preview series');
    const parked = document.getElementById('parked');
    parked.value = '=NEW(';
    parked.focus();
    parked.setSelectionRange(5, 5);
    await act(async () => {
      oldPreview();
      oldApply();
      oldCancel();
      oldPage();
    });
    await flush();
    assert(field('Series destinations'));
    assert.equal(field('Series cells shown').textContent, '1–100 of 102');
    assert.equal(next.reads.length, 1);
    assert.deepEqual(next.accepted, []);
    assert.equal(parked.value, '=NEW(');
    assert.equal(parked.selectionStart, 5);
    assert.equal(document.activeElement, parked);
  }
);
check(
  'capability/read errors remain readable, foreign previews refuse, and no fake Apply appears without a valid preview',
  async (t) => {
    const o = owner(),
      h = await mount(t, o, { context: null });
    assert.match(field('Fill series').textContent, /unavailable here/);
    assert.equal(button('Preview series'), undefined);
    await h.render({ context: o.context });
    o.status.reason = 'This workbook is read-only.';
    await h.render({});
    assert(button('Preview series').disabled);
    assert.equal(button('Preview series').title, o.status.reason);
    assert.match(field('Fill series').textContent, /read-only/);
    o.status.reason = undefined;
    o.context.preview = async () => {
      throw Error('Fill Series needs exact numeric seeds');
    };
    await h.render({});
    await click('Preview series');
    assert.match(document.querySelector('[role=alert]').textContent, /exact numeric seeds/);
    assert.equal(button('Apply series'), undefined);
    o.context.preview = async () => model({ ...binding, sheet: 'Foreign' });
    await h.render({});
    await click('Preview series');
    assert.match(document.querySelector('[role=alert]').textContent, /does not match/);
    assert.deepEqual(o.accepted, []);
  }
);
check(
  'readonly numeric Preview is inspectable while Apply is disabled with its own explicit capability reason',
  async (t) => {
    const o = owner();
    o.status.applyReason = 'This workbook is read-only.';
    await mount(t, o);
    assert(!button('Preview series').disabled);
    await click('Preview series');
    assert(field('Series destinations'));
    assert(button('Apply series').disabled);
    assert.equal(button('Apply series').title, 'This workbook is read-only.');
    assert.match(field('Fill series').textContent, /read-only/);
    const retained = props(button('Apply series')).onClick;
    await act(async () => retained());
    assert.deepEqual(o.accepted, []);
    await click('Cancel series');
    assert.equal(field('Series destinations'), null);
    assert.deepEqual(o.accepted, []);
  }
);
check(
  'scope changes between reviewed page and Apply retire the review; failure at the current owner leaves retry explicit',
  async (t) => {
    const o = owner(),
      h = await mount(t, o);
    await click('Preview series');
    const oldApply = props(button('Apply series')).onClick;
    o.status.current = false;
    await act(async () => oldApply());
    await h.render({});
    assert.deepEqual(o.accepted, []);
    assert.equal(field('Series destinations'), null);
    o.status.current = true;
    await h.render({
      context: {
        ...o.context,
        key: 'fresh-intent',
        apply: async (value) => {
          o.accepted.push(value.command);
          throw Error('A protected destination prevents Fill');
        },
      },
    });
    await click('Preview series');
    await click('Apply series');
    assert.match(document.querySelector('[role=alert]').textContent, /protected destination/);
    assert.equal(o.accepted.length, 1);
    assert(field('Series destinations'));
    assert.equal(button('Apply series').disabled, false);
  }
);

check('Dates mode shows readable dates and a day step; changing type retires an old Apply', async (t) => {
  const o = owner();
  o.context.preview = async (_signal, kind) => model(binding, [45350, 45351], { kind });
  await mount(t, o);
  await act(async () => props(field('Series type')).onChange({ target: { value: 'date' } }));
  await flush();
  await click('Preview series');
  assert.match(field('Series summary').textContent, /2024-02-28/);
  assert.match(field('Series summary').textContent, /Step \(days\)1/);
  assert.equal(field('Series destinations').querySelector('tbody td:nth-child(2)').textContent, '2024-03-01');
  const oldApply = props(button('Apply series')).onClick;
  await act(async () => props(field('Series type')).onChange({ target: { value: 'number' } }));
  await flush();
  assert.equal(button('Apply series'), undefined);
  await click('Preview series');
  await act(async () => oldApply());
  await flush();
  assert.deepEqual(o.accepted, []);
  assert.equal(field('Series destinations').querySelector('tbody td:nth-child(2)').textContent, '45352');
});
check('readonly Dates preview and Cancel remain pure with a current owner', async (t) => {
  const o = owner();
  o.status.applyReason = 'This workbook is read-only.';
  o.context.preview = async (_signal, kind) => model(binding, [45350, 45351], { kind });
  await mount(t, o);
  await act(async () => props(field('Series type')).onChange({ target: { value: 'date' } }));
  await flush();
  await click('Preview series');
  assert(button('Apply series').disabled);
  await click('Cancel series');
  assert.equal(field('Series destinations'), null);
  assert.deepEqual(o.accepted, []);
});
