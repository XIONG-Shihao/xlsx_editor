import './support/scratch-dir.mjs';
/** Actual Replace pane/model in React/JSDOM with controlled original-owner callbacks. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
const component = fileURLToPath(new URL('../', import.meta.url)),
  output = await fs.mkdtemp(path.join(component, '.build/tests/replace-ui-')),
  checks = [];
const entry = path.join(output, 'entry.ts'),
  bundle = path.join(output, 'pane.mjs');
await fs.writeFile(
  entry,
  `export {WorkbookReplace} from ${JSON.stringify(path.join(component, 'ui/WorkbookReplace.tsx'))};export * from ${JSON.stringify(path.join(component, 'ui/replacePreview.ts'))};export {homeBindingMatches} from ${JSON.stringify(path.join(component, 'ui/homeCommands.ts'))};`
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
  selection: 'A1:A4',
  selectedCell: 'A1',
  activationEpoch: 0,
  navigationGeneration: 0,
  modeGeneration: 0,
  specialSelectionEpoch: 0,
  mode: 'normal',
  specialSelection: null,
};
const configuration = { text: 'match', replacement: 'done', mode: 'raw', sheets: ['Report'], literal: false };
const queryFor = (sheets) => async (sheet, text, _mode, after, options) => {
  const matches = Object.entries(sheets[sheet] ?? {}).filter(
      ([ref, c]) =>
        (!after || +ref.match(/\d+$/)[0] > +after.match(/\d+$/)[0]) &&
        String(c.rawValue ?? c.value ?? '')
          .toLowerCase()
          .includes(text.toLowerCase())
    ),
    entries = matches.slice(0, options.matchLimit);
  return {
    cells: Object.fromEntries(entries),
    matches: entries.length,
    nextCursor: entries.length < matches.length ? entries.at(-1)[0] : null,
  };
};
const cell = (value) => ({ value, rawValue: String(value) });
async function model(b = binding, { count = 3, second = 0, skipped = false, config = configuration } = {}) {
  const sheets = {
    [b.sheet]: Object.fromEntries(Array.from({ length: count }, (_, i) => ['A' + (i + 1), cell('match ' + (i + 1))])),
  };
  if (skipped)
    sheets[b.sheet].A2 = {
      ...cell('match cache'),
      formula: '"match cache"',
      calculation: { status: 'calculated', value: 'match cache', revision: 'accepted' },
    };
  if (second)
    sheets.Other = Object.fromEntries(
      Array.from({ length: second }, (_, i) => ['B' + (i + 1), cell('match other ' + (i + 1))])
    );
  return ui.collectReplacementPreview({
    binding: b,
    configuration: { ...config, sheets: [b.sheet, ...(second ? ['Other'] : [])] },
    query: queryFor(sheets),
    isCurrent: () => true,
  });
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
    key: 'original-replace-scope',
    binding: structuredClone(binding),
    caption: 'Report — raw match → done',
    isCurrent: () => status.current && ui.homeBindingMatches(context.binding, status.binding),
    isRuntimeCurrent: () => status.runtime && ui.homeBindingMatches(context.binding, status.binding, true),
    disabledReason: (action) => status.reason ?? (action === 'apply' ? status.applyReason : undefined),
    preview: async (signal) => {
      reads.push('original');
      signals.push(signal);
      return snapshot ?? (await model(context.binding));
    },
    apply: async (value) => {
      accepted.push(structuredClone(value));
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
    await act(async () => root.render(React.createElement(ui.WorkbookReplace, input)));
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
const rows = () =>
  [...field('Replacement targets').querySelectorAll('tbody tr')].map((row) =>
    [...row.cells].map((cell) => cell.textContent)
  );
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

check('opening and Cancel are pure and do not alter an unrelated parked FormulaBar draft/caret', async (t) => {
  const o = owner(),
    parked = document.getElementById('parked');
  parked.focus();
  parked.setSelectionRange(5, 5);
  await mount(t, o);
  assert(field('Replacement preview').hasAttribute('data-formula-draft-zone'));
  assert.deepEqual(o.reads, []);
  assert.equal(field('Replacement targets'), null);
  await click('Cancel replacements');
  assert.deepEqual(o.accepted, []);
  assert.equal(parked.value, '=SUM(');
  assert.equal(parked.selectionStart, 5);
  assert.equal(parked.selectionEnd, 5);
  assert.equal(document.activeElement, parked);
});

check(
  'explicit complete Preview shows exact source/change and skipped reasons, then one captured acknowledged Apply',
  async (t) => {
    const o = owner();
    o.setSnapshot(await model(binding, { skipped: true }));
    await mount(t, o);
    await click('Preview all replacements');
    assert.deepEqual(rows(), [
      ['Report!A1', 'match 1', 'done 1'],
      ['Report!A2', 'match cache', 'Formula, calculated or stored results cannot be replaced directly.'],
      ['Report!A3', 'match 3', 'done 3'],
    ]);
    assert.match(field('Replacement summary').textContent, /Matched cells3Editable replacements2Skipped cells1/);
    assert.equal(o.accepted.length, 0);
    await click('Apply replacements');
    assert.equal(o.accepted.length, 1);
    assert.deepEqual(o.accepted[0].commands, [
      {
        type: 'setCells',
        sheet: 'Report',
        cells: [
          { cell: 'A1', expectedValue: 'match 1', expectedFormula: null, expectedRawValue: 'match 1', value: 'done 1' },
          { cell: 'A3', expectedValue: 'match 3', expectedFormula: null, expectedRawValue: 'match 3', value: 'done 3' },
        ],
      },
    ]);
    assert.equal(button('Apply replacements'), undefined);
    assert.match(field('Replacement preview').textContent, /Replacements applied. Use Undo/);
    assert(button('Preview all replacements').disabled);
  }
);

check(
  '1004 plus second-sheet targets are completely inspectable in100-row pages without extra owner reads',
  async (t) => {
    const o = owner();
    o.setSnapshot(await model(binding, { count: 1004, second: 7 }));
    await mount(t, o);
    await click('Preview all replacements');
    const all = [];
    do {
      all.push(...rows());
      if (button('More replacement cells').disabled) break;
      await click('More replacement cells');
    } while (true);
    assert.equal(all.length, 1011);
    assert.equal(new Set(all.map((row) => row[0])).size, 1011);
    assert.deepEqual(all[1003], ['Report!A1004', 'match 1004', 'done 1004']);
    assert.deepEqual(all[1004], ['Other!B1', 'match other 1', 'done other 1']);
    assert.equal(field('Replacement cells shown').textContent, '1001–1011 of 1011');
    await click('Previous replacement cells');
    assert.equal(field('Replacement cells shown').textContent, '901–1000 of 1011');
    assert.equal(o.reads.length, 1);
    assert.deepEqual(o.accepted, []);
  }
);

check(
  'reviewed snapshot is independent from a mutable producer and literal replacement strings render exactly',
  async (t) => {
    const o = owner(),
      snapshot = structuredClone(
        await model(binding, { count: 1, config: { ...configuration, replacement: '$& $1 =TRUE #N/A 00123' } })
      );
    o.setSnapshot(snapshot);
    const h = await mount(t, o);
    await click('Preview all replacements');
    assert.deepEqual(rows(), [['Report!A1', 'match 1', '$& $1 =TRUE #N/A 00123 1']]);
    snapshot.rows[0].after = 'forged';
    snapshot.commands[0].cells[0].value = 'forged';
    snapshot.configuration.replacement = 'forged';
    await h.render({});
    assert.deepEqual(rows(), [['Report!A1', 'match 1', '$& $1 =TRUE #N/A 00123 1']]);
    await click('Apply replacements');
    assert.equal(o.accepted[0].commands[0].cells[0].value, '$& $1 =TRUE #N/A 00123 1');
  }
);

check(
  'Cancel aborts a held complete preview, masks late success and joins one physical callback before another read',
  async (t) => {
    const held = defer(),
      o = owner();
    o.context.preview = (signal) => {
      o.reads.push('original');
      o.signals.push(signal);
      return o.reads.length === 1 ? held.promise : model();
    };
    await mount(t, o);
    await click('Preview all replacements');
    const retained = props(button('Preview all replacements')).onClick;
    await click('Cancel replacements');
    assert(o.signals[0].aborted);
    assert(button('Preview all replacements').disabled);
    await act(async () => retained());
    assert.equal(o.reads.length, 1);
    await act(async () => held.resolve(await model()));
    await flush();
    assert.equal(field('Replacement targets'), null);
    await click('Preview all replacements');
    assert.equal(o.reads.length, 2);
    assert(field('Replacement targets'));
    assert.deepEqual(o.accepted, []);
  }
);

check('hidden and unmounted panes abort/mask late success or failure without new-owner outcomes', async (t) => {
  for (const disposition of ['hidden', 'unmounted']) {
    const held = defer(),
      o = owner();
    o.context.preview = (signal) => {
      o.reads.push('original');
      o.signals.push(signal);
      return held.promise;
    };
    const h = await mount(t, o);
    await click('Preview all replacements');
    if (disposition === 'hidden') await h.render({ visible: false });
    else await h.unmount();
    assert(o.signals[0].aborted);
    await act(async () => held.reject(Error('old owner failure')));
    await flush();
    assert.equal(field('Replacement targets'), null);
    assert.equal(document.querySelector('[role=alert]'), null);
    assert.deepEqual(o.accepted, []);
    await h.unmount();
  }
});

check(
  'head/owner/selection/sheet/navigation/activation/mode changes retire a held reply without retargeting',
  async (t) => {
    for (const change of [
      { state: 'new-head' },
      { ownerIdentity: 'new-owner' },
      { selection: 'B1:B4', selectedCell: 'B1' },
      { sheet: 'Other' },
      { navigationGeneration: 1 },
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
      await click('Preview all replacements');
      o.status.binding = { ...o.status.binding, ...change };
      await h.render({});
      await act(async () => held.resolve(await model()));
      await flush();
      assert.equal(field('Replacement targets'), null);
      assert.match(field('Replacement preview').textContent, /changed. Reopen replacement/);
      assert.equal(o.reads.length, 1);
      assert.deepEqual(o.accepted, []);
      await h.unmount();
    }
  }
);

check('a fresh search context masks the old proposal and waits for its pending physical read', async (t) => {
  const held = defer(),
    old = owner();
  old.context.preview = (signal) => {
    old.reads.push('old');
    old.signals.push(signal);
    return held.promise;
  };
  const h = await mount(t, old);
  await click('Preview all replacements');
  const next = owner({ key: 'new-search-config', caption: 'new configuration' });
  await h.render({ context: next.context });
  assert(old.signals[0].aborted);
  assert(button('Preview all replacements').disabled);
  await act(async () => held.resolve(await model()));
  await flush();
  assert.equal(field('Replacement targets'), null);
  assert.deepEqual(old.accepted, []);
  await click('Preview all replacements');
  assert.equal(next.reads.length, 1);
  assert.deepEqual(next.accepted, []);
});

check('readonly can inspect while Apply and retained handlers obey the current capability', async (t) => {
  const o = owner();
  o.status.applyReason = 'This workbook is read-only.';
  await mount(t, o);
  assert(!button('Preview all replacements').disabled);
  await click('Preview all replacements');
  assert(field('Replacement targets'));
  assert(button('Apply replacements').disabled);
  assert.equal(button('Apply replacements').title, 'This workbook is read-only.');
  await act(async () => props(button('Apply replacements')).onClick());
  assert.deepEqual(o.accepted, []);
  await click('Cancel replacements');
  assert.equal(field('Replacement targets'), null);
});

check(
  'actual controlled acknowledgement may advance the captured head; acceptance remains truthful once and never retargets',
  async (t) => {
    const held = defer(),
      o = owner();
    o.context.apply = (value) => {
      o.accepted.push(structuredClone(value));
      return held.promise;
    };
    const h = await mount(t, o);
    await click('Preview all replacements');
    await click('Apply replacements');
    assert.equal(o.accepted.length, 1);
    o.status.binding = { ...o.status.binding, state: 'accepted-next' };
    await h.render({});
    await act(async () => held.resolve());
    await flush();
    assert.match(field('Replacement preview').textContent, /Replacements applied. Use Undo/);
    assert.equal(button('Apply replacements'), undefined);
    assert(button('Preview all replacements').disabled);
    assert.equal(o.accepted[0].binding.state, 'head');
  }
);

check(
  'late accepted old-runtime result is masked, and current native refusal stays explicit with guarded retry',
  async (t) => {
    const held = defer(),
      old = owner();
    old.context.apply = (value) => {
      old.accepted.push(value);
      return held.promise;
    };
    const h = await mount(t, old);
    await click('Preview all replacements');
    await click('Apply replacements');
    old.status.runtime = false;
    old.status.current = false;
    await h.render({});
    await act(async () => held.resolve());
    await flush();
    assert.doesNotMatch(field('Replacement preview').textContent, /Replacements applied/);
    assert.equal(old.accepted.length, 1);
    const next = owner({
      key: 'new-owner',
      apply: async () => {
        throw Error('Native protected destination refused');
      },
    });
    await h.render({ context: next.context });
    await click('Preview all replacements');
    await click('Apply replacements');
    assert.match(document.querySelector('[role=alert]').textContent, /Native protected/);
    assert(field('Replacement targets'));
    assert.equal(button('Apply replacements').disabled, false);
  }
);

check('foreign binding, invalid complete plan and all-skipped previews never enable an unreviewed Apply', async (t) => {
  const o = owner();
  o.setSnapshot(await model({ ...binding, documentId: 'foreign' }));
  const h = await mount(t, o);
  await click('Preview all replacements');
  assert.match(document.querySelector('[role=alert]').textContent, /different workbook/);
  assert.equal(button('Apply replacements'), undefined);
  const malformed = structuredClone(await model());
  malformed.matched = 913;
  o.setSnapshot(malformed);
  await h.render({});
  await click('Preview all replacements');
  assert.match(document.querySelector('[role=alert]').textContent, /Invalid replacement preview/);
  assert.equal(button('Apply replacements'), undefined);
  const skipped = await ui.collectReplacementPreview({
    binding,
    configuration,
    query: queryFor({
      Report: {
        A1: { ...cell('match cached'), formula: '"match cached"', calculation: { status: 'stored', pending: true } },
      },
    }),
    isCurrent: () => true,
  });
  o.setSnapshot(skipped);
  await h.render({});
  await click('Preview all replacements');
  assert(button('Apply replacements').disabled);
  await act(async () => props(button('Apply replacements')).onClick());
  assert.deepEqual(o.accepted, []);
});
