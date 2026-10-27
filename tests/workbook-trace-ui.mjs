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
  runs = await fs.mkdtemp(path.join(component, '.build/tests/trace-ui-'));
const entry = path.join(runs, 'entry.ts'),
  bundle = path.join(runs, 'test.mjs');
await fs.writeFile(
  entry,
  [
    'formulaTrace',
    'formulaTraceAdapter',
    'formulaTraceGeometry',
    'viewGeometry',
    'WorkbookTrace',
    'WorkbookTraceOverlay',
  ]
    .map(
      (name) =>
        `export * from ${JSON.stringify(path.join(component, 'ui', name + (name.startsWith('Workbook') ? '.tsx' : '.ts')))};`
    )
    .join('\n')
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
  {
    createFormulaTraceController,
    createScopedFormulaTraceQuery,
    createViewGeometry,
    traceEndpointPosition,
    formulaTraceArrows,
    validateFormulaTracePage,
    WorkbookTrace,
    WorkbookTraceOverlay,
  } = api;
const dom = new JSDOM('<div id="root"></div>', { url: 'http://localhost' });
Object.assign(globalThis, {
  window: dom.window,
  document: dom.window.document,
  HTMLElement: dom.window.HTMLElement,
  Node: dom.window.Node,
  IS_REACT_ACT_ENVIRONMENT: true,
});
const { createRoot } = await import('react-dom/client');
const check = (name, fn) => {
  test(name, async () => {
    await fn();
  });
};
after(async () => {
  dom.window.close();
});
const defer = () => {
  let resolve, reject;
  const promise = new Promise((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
};
const scope = (patch) => ({
  documentId: 'doc-A',
  ownerIdentity: 'runtime-A',
  state: 'opaque:owner:journal-head-A',
  sheet: 'Sheet1',
  cell: 'C1',
  activationEpoch: 0,
  active: true,
  sheets: ['Sheet1', 'Sheet2'],
  ...patch,
});
const receipt = { documentId: 'grant-A', ownerId: 'indexed-A', state: 'real-head-A', revision: 'source-sha-A' };
const edge = (from, to = 'C1', fromSheet = 'Sheet1', toSheet = 'Sheet1') => ({
  from: { sheet: fromSheet, range: from },
  to: { sheet: toSheet, range: to },
  kind: 'reference',
  via: 'formula',
});
const page = (target, direction, edges = [], patch = {}) => ({
  schema: 1,
  ...receipt,
  ...target,
  direction,
  edges,
  markers: [],
  nextCursor: null,
  coverage: { mode: 'direct-static-references', complete: true, scannedFormulas: 1, pages: 1, limited: false },
  ...patch,
});
const query = (fn) => async (bound, target, options, signal) => ({
  scope: bound,
  target,
  direction: options.direction,
  after: options.after,
  page: await fn(target, options, signal),
});
const controller = (fn) => {
  const c = createFormulaTraceController(query(fn));
  c.setScope(scope());
  return c;
};
check('repeated same-root tracing expands one real direction layer and cycle dedup closes known paths', async () => {
  const calls = [],
    c = controller(async (t, o) => {
      calls.push(t.cell);
      return page(
        t,
        o.direction,
        t.cell === 'C1' ? [edge('B1')] : t.cell === 'B1' ? [edge('A1', 'B1')] : [edge('C1', 'A1')]
      );
    });
  await c.trace('precedents');
  assert.equal(c.getSnapshot().depth, 1);
  assert.equal(c.getSnapshot().edges.length, 1);
  assert.equal(c.getSnapshot().canExpand, true);
  await c.trace('precedents');
  assert.equal(c.getSnapshot().edges.length, 2);
  await c.trace('precedents');
  assert.equal(c.getSnapshot().edges.length, 3);
  assert.equal(c.getSnapshot().complete, true);
  assert.deepEqual(calls, ['C1', 'B1', 'A1']);
  c.dispose();
});
check('direction change resets graph and dependent arrows retain source to formula orientation', async () => {
  const c = controller(async (t, o) =>
    page(t, o.direction, o.direction === 'precedents' ? [edge('A1', t.cell)] : [edge(t.cell, 'D1')])
  );
  await c.trace('precedents');
  await c.trace('dependents');
  assert.equal(c.getSnapshot().depth, 1);
  assert.equal(c.getSnapshot().edges.length, 1);
  assert.equal(c.getSnapshot().edges[0].from.range, 'C1');
  assert.equal(c.getSnapshot().edges[0].to.range, 'D1');
  c.dispose();
});
check('ranges remain rectangles and huge ranges stay unexpanded with explicit limits', async () => {
  const calls = [],
    c = controller(async (t, o) => {
      calls.push(t.cell);
      return page(t, o.direction, [edge('A1:A2000', t.cell)]);
    });
  await c.trace('precedents');
  assert.equal(c.getSnapshot().edges[0].from.range, 'A1:A2000');
  assert.equal(c.getSnapshot().markers[0].kind, 'limit');
  assert.equal(c.getSnapshot().canExpand, false);
  assert.equal(c.getSnapshot().complete, false);
  assert.deepEqual(calls, ['C1']);
  c.dispose();
});
check('continuation pages bind original target and unknown coverage never becomes a clean graph', async () => {
  const calls = [],
    c = controller(async (t, o) => {
      calls.push(o.after);
      return page(
        t,
        o.direction,
        o.after ? [edge('B1', t.cell)] : [edge('A1', t.cell)],
        o.after
          ? {
              markers: [
                {
                  sheet: 'Sheet1',
                  cell: 'C1',
                  kind: 'external',
                  reason: 'External workbook relationships are not covered.',
                },
              ],
              coverage: {
                mode: 'direct-static-references',
                complete: false,
                scannedFormulas: 4096,
                pages: 2,
                limited: false,
              },
            }
          : {
              nextCursor: 'opaque-page-2',
              coverage: {
                mode: 'direct-static-references',
                complete: false,
                scannedFormulas: 2048,
                pages: 1,
                limited: true,
              },
            }
      );
    });
  await c.trace('precedents');
  assert.deepEqual(calls, [null, 'opaque-page-2']);
  assert.equal(c.getSnapshot().edges.length, 2);
  assert.equal(c.getSnapshot().complete, false);
  assert.equal(c.getSnapshot().markers[0].kind, 'external');
  assert.equal(c.getSnapshot().canExpand, true);
  c.dispose();
});
check('ignored abort and duplicate trace never overlap physical reads or adopt cancelled arrows', async () => {
  const gate = defer();
  let reads = 0,
    active = 0,
    maximum = 0;
  const c = controller(async (t, o) => {
    reads++;
    active++;
    maximum = Math.max(maximum, active);
    if (reads === 1) await gate.promise;
    active--;
    return page(t, o.direction, [edge('A1', t.cell)]);
  });
  const first = c.trace('precedents'),
    joined = c.trace('precedents');
  assert.equal(first, joined);
  c.cancel();
  const cancelledJoin = c.trace('precedents');
  assert.equal(cancelledJoin, first);
  gate.resolve();
  assert.equal((await first).outcome, 'retired');
  assert.equal(c.getSnapshot().edges.length, 0);
  assert.equal(reads, 1);
  await c.trace('precedents');
  assert.equal(reads, 2);
  assert.equal(maximum, 1);
  assert.equal(c.getSnapshot().edges.length, 1);
  c.dispose();
  assert.equal(c.getSnapshot().scope, null);
});
check('head owner sheet activation and inactive changes atomically mask accepted edges before listeners', async () => {
  for (const patch of [
    { state: 'next-head' },
    { ownerIdentity: 'owner-B' },
    { sheet: 'Sheet2' },
    { activationEpoch: 2 },
    { active: false },
  ]) {
    const c = controller(async (t, o) => page(t, o.direction, [edge('A1', t.cell)]));
    await c.trace('precedents');
    const observed = [];
    c.subscribe(() => observed.push(c.getSnapshot()));
    c.setScope(scope(patch));
    assert(observed.length);
    assert(observed.every((s) => s.edges.length === 0 && s.markers.length === 0));
    assert.equal(c.getSnapshot().depth, 0);
    c.dispose();
  }
});
check('scope change during held reply rejects late data and joins before a new original callback', async () => {
  const gate = defer(),
    calls = [];
  const c = controller(async (t, o) => {
    calls.push(t.cell);
    if (calls.length === 1) await gate.promise;
    return page(t, o.direction, [edge('A1', t.cell)]);
  });
  const old = c.trace('precedents');
  c.setScope(scope({ cell: 'D1' }));
  const latest = c.trace('precedents');
  assert.notEqual(latest, old);
  assert.deepEqual(calls, ['C1']);
  gate.resolve();
  assert.equal((await old).outcome, 'retired');
  assert.equal((await latest).outcome, 'completed');
  assert.equal(c.getSnapshot().edges[0].to.range, 'D1');
  assert.deepEqual(calls, ['C1', 'D1']);
  c.dispose();
});
check('depth and page read budgets refuse further expansion with explicit partial coverage', async () => {
  const calls = [],
    c = controller(async (t, o) => {
      calls.push(t.cell);
      const n = Number(t.cell.slice(1));
      return page(t, o.direction, [edge('C' + (n + 1), t.cell)]);
    });
  for (let i = 0; i < 32; i++) await c.trace('precedents');
  assert.equal(c.getSnapshot().depth, 32);
  assert.equal(c.getSnapshot().canExpand, false);
  assert.equal(c.getSnapshot().complete, false);
  assert(c.getSnapshot().markers.some((x) => x.kind === 'limit'));
  await c.trace('precedents');
  assert.equal(calls.length, 32);
  c.dispose();
  const paged = controller(async (t, o) =>
    page(t, o.direction, [], {
      nextCursor: String(Number(o.after ?? 0) + 1),
      coverage: { mode: 'direct-static-references', complete: false, scannedFormulas: 2048, pages: 1, limited: true },
    })
  );
  await paged.trace('precedents');
  assert.equal(paged.getSnapshot().requests, 64);
  assert.equal(paged.getSnapshot().complete, false);
  assert(paged.getSnapshot().markers.length);
  paged.dispose();
});
check('strict page receipt rejects missing cursor malformed edges unknown sheets and unrelated roots', async () => {
  const target = { sheet: 'Sheet1', cell: 'C1' },
    valid = page(target, 'precedents', [edge('A1')]);
  assert.equal(validateFormulaTracePage(valid, scope(), target, 'precedents'), valid);
  const bad = [
    { ...valid, coverage: { ...valid.coverage, scannedFormulas: 131073 } },
    { ...valid, nextCursor: undefined },
    { ...valid, edges: [null] },
    { ...valid, edges: [edge('A1', 'D1')] },
    { ...valid, edges: [edge('A1', 'C1', 'ungranted')] },
    { ...valid, edges: [{ ...edge('A1'), via: undefined }] },
    { ...valid, coverage: { ...valid.coverage, complete: true, limited: true } },
    { ...valid, markers: [{ sheet: 'Sheet1', kind: 'external', reason: '' }] },
  ];
  for (const p of bad) assert.throws(() => validateFormulaTracePage(p, scope(), target, 'precedents'));
});
check('adapter captures callback and independently checks actual receipt without parsing opaque UI token', async () => {
  const gate = defer(),
    calls = [];
  let frame = {
    scope: scope(),
    receipt,
    trace: async (sheet, cell, direction, after, limit) => {
      calls.push({ actualState: receipt.state, sheet, cell, direction, after, limit });
      return gate.promise;
    },
  };
  const adapter = createScopedFormulaTraceQuery(() => frame),
    bound = scope(),
    target = { sheet: 'Sheet1', cell: 'C1' },
    options = { direction: 'precedents', after: null, limit: 100 };
  const pending = adapter(bound, target, options, new AbortController().signal);
  frame = {
    ...frame,
    trace: async () => {
      throw Error('must not retarget');
    },
  };
  gate.resolve(page(target, 'precedents', [edge('A1')]));
  assert.equal((await pending).page.state, 'real-head-A');
  assert.equal(calls.length, 1);
  assert.equal(calls[0].actualState, 'real-head-A');
  frame = { ...frame, trace: async () => ({ ...page(target, 'precedents'), state: 'another-real-head' }) };
  await assert.rejects(adapter(bound, target, options, new AbortController().signal), /another document/);
});
check('adapter rejects late real-shaped receipts on head and actual owner replacement', async () => {
  for (const change of ['scope', 'receipt']) {
    const gate = defer();
    let frame = { scope: scope(), receipt, trace: async () => gate.promise };
    const adapter = createScopedFormulaTraceQuery(() => frame),
      target = { sheet: 'Sheet1', cell: 'C1' };
    const pending = adapter(
      scope(),
      target,
      { direction: 'precedents', after: null, limit: 100 },
      new AbortController().signal
    );
    frame =
      change === 'scope'
        ? { ...frame, scope: scope({ state: 'opaque-B' }) }
        : { ...frame, receipt: { ...receipt, ownerId: 'indexed-B' } };
    gate.resolve(page(target, 'precedents', [edge('A1')]));
    await assert.rejects(pending, /accepted state changed/);
  }
});
check(
  'literal-profile zoom headers scroll and frozen strips use actual screen coordinates; hidden/offscreen cells get markers',
  () => {
    const sheet = { freeze: { rows: 1, columns: 1 }, hiddenRows: [2], widths: [], rowHeights: [] },
      g = createViewGeometry(sheet, { rows: 20, columns: 10, zoom: 150, literal: true }),
      v = { width: 600, height: 400, scrollLeft: 150, scrollTop: 87 };
    const frozen = traceEndpointPosition({ sheet: 'Sheet1', range: 'A1' }, 'Sheet1', g, v, sheet);
    assert.equal(frozen.visible, true);
    assert.equal(frozen.point.x, 147);
    assert.equal(frozen.point.y, 91.5);
    const visible = traceEndpointPosition({ sheet: 'Sheet1', range: 'D4' }, 'Sheet1', g, v, sheet);
    assert.equal(visible.visible, true);
    assert.equal(visible.point.x, 447);
    assert.equal(visible.point.y, 178.5);
    assert.equal(traceEndpointPosition({ sheet: 'Sheet1', range: 'B2' }, 'Sheet1', g, v, sheet).kind, 'offscreen');
    assert.equal(traceEndpointPosition({ sheet: 'Sheet1', range: 'A3' }, 'Sheet1', g, v, sheet).kind, 'hidden');
    assert.equal(traceEndpointPosition({ sheet: 'Sheet2', range: 'A1' }, 'Sheet1', g, v, sheet).kind, 'cross-sheet');
    const merged = { ...sheet, merges: ['B5:C5'] };
    assert.equal(traceEndpointPosition({ sheet: 'Sheet1', range: 'B5' }, 'Sheet1', g, v, merged).rectangle.width, 150);
    const split = { ...sheet, merges: ['A5:B5'] };
    assert.equal(traceEndpointPosition({ sheet: 'Sheet1', range: 'A5' }, 'Sheet1', g, v, split).kind, 'split-merge');
  }
);
check('literal-profile arrow direction remains input to formula and self-cycle has visible loop geometry', () => {
  const g = createViewGeometry(undefined, { rows: 20, columns: 10, literal: true }),
    v = { width: 700, height: 500, scrollLeft: 0, scrollTop: 0 };
  const a = formulaTraceArrows(
    [
      { ...edge('A1'), id: 'known' },
      { ...edge('C1'), id: 'cycle' },
      { ...edge('A1', 'C1', 'Sheet2'), id: 'other' },
    ],
    'Sheet1',
    g,
    v
  );
  assert.match(a[0].path, /^M 98 61 L 298 61$/);
  assert.match(a[1].path, / C /);
  assert.equal(a[2].from.kind, 'cross-sheet');
  assert.equal(a[2].to.visible, true);
});
check('ordinary source-profile arrows use source dimensions independently of literal-sheet defaults', () => {
  const sheet = {
    sheetFormat: { defaultRowHeight: 15 },
    widths: [
      { min: 1, max: 1, width: 8.43, hidden: false },
      { min: 2, max: 2, width: 8.43, hidden: false },
      { min: 3, max: 3, width: 8.43, hidden: false },
    ],
    rowHeights: [],
  };
  const g = createViewGeometry(sheet, { rows: 20, columns: 10 }),
    v = { width: 700, height: 500, scrollLeft: 0, scrollTop: 0 },
    a = formulaTraceArrows([{ ...edge('A1'), id: 'source' }], 'Sheet1', g, v, sheet);
  assert.equal(g.rowOffsets[1], 20);
  assert.equal(g.columnOffsets[1], 59);
  assert.equal(a[0].path, 'M 77.5 42 L 195.5 42');
  assert.equal(a[0].from.visible, true);
  assert.equal(a[0].to.visible, true);
});
const button = (name) =>
  [...document.querySelectorAll('button')].find((n) => (n.getAttribute('aria-label') ?? n.textContent) === name);
async function click(node) {
  assert(node);
  assert.equal(node.disabled, false);
  await act(async () => node.click());
}
async function mounted(fn, extra = {}) {
  const c = controller(fn),
    root = createRoot(document.getElementById('root'));
  let s = scope(),
    navigate = [];
  const geometry = createViewGeometry(undefined, { rows: 100, columns: 10 }),
    viewport = { width: 600, height: 400, scrollLeft: 0, scrollTop: 0 };
  const draw = () =>
    React.createElement(
      React.Fragment,
      null,
      React.createElement(WorkbookTrace, { controller: c, scope: s, onNavigate: (t) => navigate.push(t) }),
      React.createElement(WorkbookTraceOverlay, {
        controller: c,
        scope: s,
        geometry,
        viewport,
        onNavigate: (t) => navigate.push(t),
        onOpenList: () => navigate.push('list'),
      })
    );
  const render = async (patch) => {
    s = { ...s, ...patch };
    await act(async () => root.render(extra.strict ? React.createElement(React.StrictMode, null, draw()) : draw()));
  };
  await render({});
  return {
    c,
    root,
    render,
    navigate,
    close: async () => {
      await act(async () => root.unmount());
      c.dispose();
    },
  };
}
check(
  'real React pane renders known arrows cross-sheet/offscreen navigation and explicit partial reasons without writes',
  async () => {
    const ui = await mounted(async (t, o) =>
      page(t, o.direction, [edge('A1'), edge('A1', 'C1', 'Sheet2'), edge('A80')], {
        markers: [{ sheet: 'Sheet1', cell: 'C1', kind: 'dynamic', reason: 'Dynamic references are not covered.' }],
        coverage: { mode: 'direct-static-references', complete: false, scannedFormulas: 1, pages: 1, limited: false },
      })
    );
    try {
      await click(button('Trace Precedents'));
      assert.equal(document.querySelectorAll('[data-trace-arrow]').length, 3);
      assert.match(document.querySelector('[aria-label="Trace coverage limits"]').textContent, /Dynamic references/);
      await click(button('Go to trace Sheet2!A1'));
      assert.deepEqual(ui.navigate[0], { sheet: 'Sheet2', range: 'A1' });
      await click(button('Go to trace source Sheet1!A80'));
      assert.deepEqual(ui.navigate[1], { sheet: 'Sheet1', range: 'A80' });
      assert.equal(ui.c.getSnapshot().scope.cell, 'C1');
      await click(button('Remove Arrows'));
      assert.equal(document.querySelector('[aria-label="Formula trace arrows"]'), null);
    } finally {
      await ui.close();
    }
  }
);
check(
  'pane and overlay immediately mask head/sheet/owner/active scopes and stale direct navigation handlers refuse',
  async () => {
    const ui = await mounted(async (t, o) => page(t, o.direction, [edge('A1')]));
    try {
      await click(button('Trace Precedents'));
      const old = button('Go to trace source Sheet1!A1'),
        props = old[Object.keys(old).find((k) => k.startsWith('__reactProps$'))],
        stale = props.onClick;
      await ui.render({ state: 'opaque-next' });
      assert.equal(document.querySelector('[data-trace-arrow]'), null);
      assert.equal(document.querySelector('[data-trace-edge]'), null);
      stale();
      assert.deepEqual(ui.navigate, []);
      await ui.render({ active: false });
      assert.equal(button('Trace Precedents').disabled, true);
      assert.match(document.body.textContent, /Trace paused/);
      await ui.render({ active: true, ownerIdentity: 'runtime-B' });
      assert.equal(document.querySelector('[data-trace-edge]'), null);
    } finally {
      await ui.close();
    }
  }
);
check('retained pane Cancel and StrictMode preserve only prior layers and join ignored abort', async () => {
  const gate = defer();
  let reads = 0;
  const ui = await mounted(
    async (t, o) => {
      reads++;
      if (reads === 2) await gate.promise;
      return page(t, o.direction, [edge(reads === 1 ? 'B1' : 'A1', t.cell)]);
    },
    { strict: true }
  );
  try {
    await click(button('Trace Precedents'));
    assert.equal(ui.c.getSnapshot().edges.length, 1);
    await click(button('Trace Precedents'));
    await click(button('Cancel trace'));
    assert.equal(ui.c.getSnapshot().edges.length, 1);
    await act(async () => gate.resolve());
    assert.equal(ui.c.getSnapshot().edges.length, 1);
    assert.equal(reads, 2);
    assert.match(document.body.textContent, /Previously shown arrows/);
  } finally {
    await ui.close();
  }
});

check('small range frontier expands fully and preserves one rectangle while limits bound nodes and edges', async () => {
  const calls = [],
    c = controller(async (t, o) => {
      calls.push(t.cell);
      return page(t, o.direction, t.cell === 'C1' ? [edge('A1:B2')] : []);
    });
  await c.trace('precedents');
  await c.trace('precedents');
  assert.deepEqual(calls, ['C1', 'A1', 'B1', 'A2', 'B2']);
  assert.equal(c.getSnapshot().edges.length, 1);
  assert.equal(c.getSnapshot().edges[0].from.range, 'A1:B2');
  assert.equal(c.getSnapshot().complete, true);
  c.dispose();
  const large = controller(async (t, o) => {
    const start = Number(o.after ?? 0);
    return page(
      t,
      o.direction,
      Array.from({ length: 100 }, (_, i) => edge('A' + (start + i + 1))),
      {
        nextCursor: String(start + 100),
        coverage: { mode: 'direct-static-references', complete: false, scannedFormulas: 1, pages: 1, limited: true },
      }
    );
  });
  await large.trace('precedents');
  assert.equal(large.getSnapshot().edges.length, 1024);
  assert.equal(large.getSnapshot().canExpand, false);
  assert.equal(large.getSnapshot().complete, false);
  assert(large.getSnapshot().markers.some((m) => m.reason.includes('cell limit')));
  large.dispose();
});
check('explicit direction switch queues only latest intent after ignored-abort physical read', async () => {
  const gate = defer(),
    calls = [];
  let physical = 0,
    maximum = 0;
  const c = controller(async (t, o) => {
    calls.push(o.direction);
    physical++;
    maximum = Math.max(maximum, physical);
    if (calls.length === 1) await gate.promise;
    physical--;
    return page(t, o.direction, o.direction === 'precedents' ? [edge('A1')] : [edge('C1', 'D1')]);
  });
  const first = c.trace('precedents'),
    latest = c.trace('dependents'),
    duplicate = c.trace('dependents');
  assert.equal(latest, duplicate);
  assert.deepEqual(calls, ['precedents']);
  gate.resolve();
  assert.equal((await first).outcome, 'retired');
  assert.equal((await latest).outcome, 'completed');
  assert.deepEqual(calls, ['precedents', 'dependents']);
  assert.equal(maximum, 1);
  assert.equal(c.getSnapshot().direction, 'dependents');
  assert.equal(c.getSnapshot().edges[0].to.range, 'D1');
  c.dispose();
});

check('a retained unavailable pane shows its inventory reason and cannot start another query', async () => {
  const c = controller(async (t, o) => page(t, o.direction, [edge('A1')])),
    root = createRoot(document.getElementById('root'));
  let reads = 0;
  const reason = 'Formula tracing supports up to 1,024 worksheets in a workbook.';
  try {
    await act(async () =>
      root.render(
        React.createElement(WorkbookTrace, {
          controller: c,
          scope: scope(),
          disabledReason: reason,
          onNavigate: () => reads++,
        })
      )
    );
    assert.equal(button('Trace Precedents').disabled, true);
    assert.equal(button('Trace Dependents').disabled, true);
    assert.match(document.body.textContent, /up to 1,024 worksheets/);
    assert.equal(c.getSnapshot().requests, 0);
    assert.equal(reads, 0);
  } finally {
    await act(async () => root.unmount());
    c.dispose();
  }
});
