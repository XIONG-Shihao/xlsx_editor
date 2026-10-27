import assert from 'node:assert/strict';
import test from 'node:test';
import { diagnosticPayload, registerWorkbookDiagnostics } from '../ui/workbookDiagnostics.ts';
import { useRangeCells } from '../ui/rangeCells.ts';
import { WorkbookWorker } from '../ui/workbookWorker.ts';
import * as React from 'react';
import { createRoot } from 'react-dom/client';
import { JSDOM } from 'jsdom';

test('payload estimates deduplicate shared objects and cap work without emitting content', () => {
  const cell = { value: 'PRIVATE_CELL', count: 3 };
  const shared = diagnosticPayload([cell, cell]),
    copied = diagnosticPayload([cell, { ...cell }]);
  assert(shared.estimatedPayloadBytes < copied.estimatedPayloadBytes);
  assert.equal(shared.truncated, 0);
  const cycle = { value: 'PRIVATE_CELL' };
  cycle.self = cycle;
  assert.equal(diagnosticPayload(cycle).truncated, 0);
  const wide = Object.fromEntries(Array.from({ length: 1000 }, (_, i) => [i, i]));
  const limited = diagnosticPayload(wide, 16);
  assert.equal(limited.truncated, 1);
  assert(limited.visitedValues <= 16);
  assert(!JSON.stringify([shared, copied, limited]).includes('PRIVATE_CELL'));
});

test('diagnostics are opt-in, numeric-only, and cannot fail workbook operations', () => {
  const previous = globalThis.__platformWorkbookDiagnostics;
  try {
    delete globalThis.__platformWorkbookDiagnostics;
    let reads = 0;
    assert.equal(
      registerWorkbookDiagnostics('viewport', () => {
        reads++;
        return {};
      }),
      undefined
    );
    assert.equal(reads, 0);
    let snapshot;
    globalThis.__platformWorkbookDiagnostics = {
      register(_kind, read) {
        snapshot = read;
        return () => {
          throw Error('sink failure');
        };
      },
    };
    const release = registerWorkbookDiagnostics('viewport', () => ({
      valid: 3,
      text: 'PRIVATE',
      infinite: Infinity,
      'invalid key': 1,
    }));
    assert.deepEqual(snapshot(), { valid: 3 });
    assert.doesNotThrow(release);
    registerWorkbookDiagnostics('viewport', () => {
      throw Error('PRIVATE_FAILURE');
    });
    assert.deepEqual(snapshot(), { diagnosticError: 1 });
  } finally {
    if (previous === undefined) delete globalThis.__platformWorkbookDiagnostics;
    else globalThis.__platformWorkbookDiagnostics = previous;
  }
});

test('actual range hook accounts live tiles, releases evicted windows and unregisters on close', async () => {
  const dom = new JSDOM('<!doctype html><div id="root"></div>', { url: 'http://localhost/' });
  const previous = {
    window: globalThis.window,
    document: globalThis.document,
    act: globalThis.IS_REACT_ACT_ENVIRONMENT,
    diagnostics: globalThis.__platformWorkbookDiagnostics,
  };
  const owners = new Map();
  let sequence = 0,
    api,
    reads = 0;
  Object.assign(globalThis, {
    window: dom.window,
    document: dom.window.document,
    IS_REACT_ACT_ENVIRONMENT: true,
    __platformWorkbookDiagnostics: {
      register(kind, read) {
        const id = ++sequence;
        owners.set(id, { kind, read });
        return () => owners.delete(id);
      },
    },
  });
  const root = createRoot(dom.window.document.getElementById('root'));
  const sheet = {
    name: 'PRIVATE_SHEET',
    path: 'private.xml',
    rows: 1000,
    columns: 100,
    cells: {},
    merges: [],
    widths: [],
  };
  const reader = async (_sheet, range) => {
    reads++;
    return { [range.split(':')[0]]: { value: 'PRIVATE_CELL' } };
  };
  function Host({ revision }) {
    api = useRangeCells([sheet], reader, revision);
    return null;
  }
  const snapshot = () => [...owners.values()].find((owner) => owner.kind === 'viewport').read();
  async function load(range) {
    const key = api.request([{ sheet: sheet.name, range }]);
    await React.act(async () => {
      api.load(key);
      await new Promise(setImmediate);
    });
    assert(api.ready(key));
    return snapshot();
  }
  try {
    await React.act(async () => {
      root.render(React.createElement(Host, { revision: '1' }));
      await new Promise(setImmediate);
    });
    const first = await load('A1');
    assert.equal(first.retainedTiles, 1);
    assert.equal(first.publishedTiles, 1);
    assert.equal(first.truncated, 0);
    assert(first.estimatedPayloadBytes > 0);
    await load('A1');
    assert.equal(reads, 1, 'Instrumentation must not change tile reuse');
    const distant = await load('A65');
    assert.equal(distant.retainedTiles, 1);
    assert.equal(distant.publishedTiles, 1);
    assert(distant.estimatedPayloadBytes < first.estimatedPayloadBytes * 1.2, 'Old tile payload must not accumulate');
    assert(!JSON.stringify(distant).includes('PRIVATE'));
    await React.act(async () => {
      root.render(React.createElement(Host, { revision: '2' }));
      await new Promise(setImmediate);
    });
    await load('A65');
    assert(reads >= 3, 'New revision must not borrow an old tile');
    await React.act(async () => root.unmount());
    assert.equal(owners.size, 0);
  } finally {
    await React.act(async () => root.unmount());
    dom.window.close();
    globalThis.window = previous.window;
    globalThis.document = previous.document;
    globalThis.IS_REACT_ACT_ENVIRONMENT = previous.act;
    if (previous.diagnostics === undefined) delete globalThis.__platformWorkbookDiagnostics;
    else globalThis.__platformWorkbookDiagnostics = previous.diagnostics;
  }
});

test('indexed renderer-worker pending requests are counted and released on cancellation', async () => {
  const previous = globalThis.__platformWorkbookDiagnostics;
  const owners = new Map();
  let reply;
  globalThis.__platformWorkbookDiagnostics = {
    register(kind, read) {
      owners.set(kind, read);
      return () => owners.delete(kind);
    },
  };
  const worker = new WorkbookWorker(
    () =>
      new Promise((resolve) => {
        reply = resolve;
      })
  );
  const view = {
    revision: 'PRIVATE',
    storage: { kind: 'rust-indexed-v1', sessionId: 'PRIVATE', state: '0' },
    sheets: [],
  };
  try {
    const pending = worker.apply(view, []),
      refused = assert.rejects(pending, /closed/);
    assert.equal(owners.get('renderer-worker')().pendingOwnerRequests, 1);
    worker.dispose();
    await refused;
    assert.equal(owners.size, 0);
    reply({ view, edits: [] });
    await new Promise(setImmediate);
    assert.equal(owners.size, 0);
  } finally {
    worker.dispose();
    if (previous === undefined) delete globalThis.__platformWorkbookDiagnostics;
    else globalThis.__platformWorkbookDiagnostics = previous;
  }
});

test('duplicate summaries load on demand with whole-rule overlap, tile edges and stale-revision suppression', async () => {
  const dom = new JSDOM('<!doctype html><div id="root"></div>', { url: 'http://localhost/' });
  const previous = {
    window: globalThis.window,
    document: globalThis.document,
    act: globalThis.IS_REACT_ACT_ENVIRONMENT,
    diagnostics: globalThis.__platformWorkbookDiagnostics,
  };
  let api,
    version = 1,
    releaseOld,
    held = false,
    snapshot;
  const calls = [],
    remote = 'A1000:A1002',
    edge = 'A32:A34',
    overlapping = 'A31:A33';
  Object.assign(globalThis, {
    window: dom.window,
    document: dom.window.document,
    IS_REACT_ACT_ENVIRONMENT: true,
    __platformWorkbookDiagnostics: {
      register(_kind, read) {
        snapshot = read;
        return () => {};
      },
    },
  });
  const root = createRoot(dom.window.document.getElementById('root'));
  const sheet = {
    name: 'Sheet',
    path: '',
    rows: 2000,
    columns: 32,
    cells: {},
    merges: [],
    widths: [],
    conditionalRules: [remote, edge, overlapping].map((range) => ({ range, rule: 'duplicateValues' })),
  };
  const reader = async (_sheet, range) => {
    calls.push({ version, range });
    const [first, last = first] = range.split(':'),
      min = Number(first.match(/\d+$/)[0]),
      max = Number(last.match(/\d+$/)[0]);
    const cells = Object.fromEntries(
      [1, 31, 32, 33, 34, 1000, 1001, 1002]
        .filter((row) => row >= min && row <= max)
        .map((row) => [
          `A${row}`,
          { value: row === 31 || row === 34 || row === 1002 ? 'unique-' + row : 'version-' + version },
        ])
    );
    if (range === remote && version === 1 && !held) {
      held = true;
      return new Promise((resolve) => {
        releaseOld = () => resolve(cells);
      });
    }
    return cells;
  };
  function Host({ revision }) {
    api = useRangeCells([sheet], reader, revision);
    return null;
  }
  const load = async (range) => {
    const key = api.request([{ sheet: 'Sheet', range }]);
    await React.act(async () => {
      api.load(key);
      await new Promise(setImmediate);
    });
    return key;
  };
  try {
    await React.act(async () => {
      root.render(React.createElement(Host, { revision: 'one' }));
      await new Promise(setImmediate);
    });
    const initial = await load('A1');
    assert(api.ready(initial));
    assert(!calls.some((call) => call.range === remote), 'Offscreen rules must not delay initial readiness');
    assert(
      calls.some((call) => call.range === edge),
      'Rule touching row32 tile boundary must load'
    );
    assert(
      calls.some((call) => call.range === overlapping),
      'Overlapping rules each need their complete declared range'
    );
    assert.deepEqual(api.sheets[0].duplicateValues[edge], ['"version-1"']);
    assert.deepEqual(api.sheets[0].duplicateValues[overlapping], ['"version-1"']);
    const waiting = await load('A1000');
    assert(releaseOld);
    assert.equal(api.ready(waiting), false, 'Do not mark ready before the required summary is available');
    assert.equal(snapshot().inFlightTiles, 1);
    version = 2;
    await React.act(async () => {
      root.render(React.createElement(Host, { revision: 'two' }));
      await new Promise(setImmediate);
    });
    const current = await load('A1000');
    assert(api.ready(current));
    assert.deepEqual(api.sheets[0].duplicateValues[remote], ['"version-2"']);
    await React.act(async () => {
      releaseOld();
      await new Promise(setImmediate);
    });
    assert.deepEqual(
      api.sheets[0].duplicateValues[remote],
      ['"version-2"'],
      'Late old-revision summary must not replace current data'
    );
    const remoteCalls = calls.filter((call) => call.range === remote).length;
    await load('A1');
    await load('A1000');
    assert.equal(
      calls.filter((call) => call.range === remote).length,
      remoteCalls,
      'Returning to a visited rule reuses its current-revision summary'
    );
  } finally {
    await React.act(async () => root.unmount());
    dom.window.close();
    globalThis.window = previous.window;
    globalThis.document = previous.document;
    globalThis.IS_REACT_ACT_ENVIRONMENT = previous.act;
    if (previous.diagnostics === undefined) delete globalThis.__platformWorkbookDiagnostics;
    else globalThis.__platformWorkbookDiagnostics = previous.diagnostics;
  }
});

test('duplicate summary cache evicts least recently used entries by bytes and keeps sheet scopes independent', async () => {
  const { RangeSummaryCache } = await import('../ui/rangeSummaryCache.ts');
  const cache = new RangeSummaryCache(600, 2),
    a = ['a'.repeat(20)],
    b = ['b'.repeat(20)],
    c = ['c'.repeat(20)];
  cache.set('s', 'A1:A2', a);
  cache.set('t', 'A1:A2', b);
  assert.equal(cache.get('s', 'A1:A2'), a);
  cache.set('u', 'A1:A2', c);
  assert.equal(cache.get('t', 'A1:A2'), undefined);
  assert.equal(cache.get('s', 'A1:A2'), a);
  assert.equal(cache.get('u', 'A1:A2'), c);
  assert(cache.diagnostics().estimatedSummaryCacheBytes <= 600);
  cache.set('s', 'A1:A2', ['x'.repeat(1000)]);
  assert.equal(cache.get('s', 'A1:A2'), undefined);
  assert.equal(cache.get('u', 'A1:A2'), c);
  assert.equal(cache.diagnostics().retainedSummaries, 1);
  const countBound = new RangeSummaryCache(10000, 1);
  countBound.set('s', 'A1', []);
  countBound.set('s', 'B1', []);
  assert.equal(countBound.get('s', 'A1'), undefined);
  assert.deepEqual(countBound.get('s', 'B1'), []);
});

test('oversized duplicate summaries render correctly without accumulating offscreen or across revisions', async () => {
  const dom = new JSDOM('<!doctype html><div id="root"></div>', { url: 'http://localhost/' }),
    previous = {
      window: globalThis.window,
      document: globalThis.document,
      act: globalThis.IS_REACT_ACT_ENVIRONMENT,
      diagnostics: globalThis.__platformWorkbookDiagnostics,
    };
  let api,
    snapshot,
    version = 1,
    ruleReads = 0;
  Object.assign(globalThis, {
    window: dom.window,
    document: dom.window.document,
    IS_REACT_ACT_ENVIRONMENT: true,
    __platformWorkbookDiagnostics: {
      register(_kind, read) {
        snapshot = read;
        return () => {};
      },
    },
  });
  const root = createRoot(dom.window.document.getElementById('root')),
    rule = 'A1:A10000',
    sheet = {
      name: 'Sheet',
      path: '',
      rows: 10100,
      columns: 1,
      cells: {},
      merges: [],
      widths: [],
      conditionalRules: [{ range: rule, rule: 'duplicateValues' }],
    };
  const value = (row) => `${version}:${(row - 1) % 5000}:` + 'x'.repeat(240);
  const read = async (_sheet, range) => {
    if (range === rule) ruleReads++;
    const [a, b = a] = range.split(':'),
      first = Number(a.match(/\d+$/)[0]),
      last = Number(b.match(/\d+$/)[0]);
    return Object.fromEntries(
      Array.from({ length: last - first + 1 }, (_, i) => [`A${first + i}`, { value: value(first + i) }])
    );
  };
  function Host({ revision }) {
    api = useRangeCells([sheet], read, revision);
    return null;
  }
  const load = async (range) => {
    const key = api.request([{ sheet: 'Sheet', range }]);
    await React.act(async () => {
      api.load(key);
      await new Promise(setImmediate);
    });
    assert(api.ready(key));
  };
  try {
    await React.act(async () => {
      root.render(React.createElement(Host, { revision: 'one' }));
      await new Promise(setImmediate);
    });
    await load('A1');
    assert.equal(api.sheets[0].duplicateValues[rule].length, 5000);
    assert(api.sheets[0].duplicateValues[rule].includes(JSON.stringify(value(1))));
    assert.equal(snapshot().retainedSummaries, 0);
    assert.equal(snapshot().estimatedSummaryCacheBytes, 0);
    await load('A10049');
    assert.equal(api.sheets[0].duplicateValues[rule], undefined, 'Offscreen published summaries must be released');
    await load('A1');
    assert.equal(ruleReads, 2, 'An oversized summary is recomputed rather than cached beyond the bound');
    version = 2;
    await React.act(async () => {
      root.render(React.createElement(Host, { revision: 'two' }));
      await new Promise(setImmediate);
    });
    await load('A1');
    assert(api.sheets[0].duplicateValues[rule].includes(JSON.stringify(value(1))));
    assert(!api.sheets[0].duplicateValues[rule].some((v) => v.startsWith('"1:')));
  } finally {
    await React.act(async () => root.unmount());
    dom.window.close();
    globalThis.window = previous.window;
    globalThis.document = previous.document;
    globalThis.IS_REACT_ACT_ENVIRONMENT = previous.act;
    if (previous.diagnostics === undefined) delete globalThis.__platformWorkbookDiagnostics;
    else globalThis.__platformWorkbookDiagnostics = previous.diagnostics;
  }
});

test('revision changes request each visible tile once and ignore obsolete request keys', async () => {
  const dom = new JSDOM('<!doctype html><div id="root"></div>', { url: 'http://localhost/' });
  const previous = {
    window: globalThis.window,
    document: globalThis.document,
    act: globalThis.IS_REACT_ACT_ENVIRONMENT,
  };
  Object.assign(globalThis, { window: dom.window, document: dom.window.document, IS_REACT_ACT_ENVIRONMENT: true });
  const root = createRoot(dom.window.document.getElementById('root')),
    calls = [];
  const sheet = { name: 'Sheet', path: 'sheet.xml', rows: 100, columns: 32, cells: {}, merges: [], widths: [] };
  let api, currentKey;
  function Host({ revision }) {
    api = useRangeCells(
      [sheet],
      async (_sheet, range) => {
        calls.push({ revision, range });
        return { A1: { value: revision } };
      },
      revision
    );
    const key = api.request([{ sheet: 'Sheet', range: 'A1' }]);
    currentKey = key;
    React.useEffect(() => api.load(key), [key]);
    return null;
  }
  try {
    await React.act(async () => {
      root.render(React.createElement(Host, { revision: 'one' }));
      await new Promise(setImmediate);
    });
    assert(api.ready(currentKey));
    const oldKey = currentKey;
    await React.act(async () => {
      root.render(React.createElement(Host, { revision: 'two' }));
      await new Promise(setImmediate);
    });
    assert(api.ready(currentKey));
    assert.equal(api.sheets[0].cells.A1.value, 'two');
    assert.deepEqual(
      calls,
      [
        { revision: 'one', range: 'A1:AF32' },
        { revision: 'two', range: 'A1:AF32' },
      ],
      'The previous revision request must not start a redundant read using the new reader'
    );
    await React.act(async () => {
      api.load(oldKey);
      await new Promise(setImmediate);
    });
    assert.equal(calls.length, 2, 'A delayed obsolete request must not read the current revision');
    assert.equal(api.sheets[0].cells.A1.value, 'two');
  } finally {
    await React.act(async () => root.unmount());
    dom.window.close();
    globalThis.window = previous.window;
    globalThis.document = previous.document;
    globalThis.IS_REACT_ACT_ENVIRONMENT = previous.act;
  }
});
