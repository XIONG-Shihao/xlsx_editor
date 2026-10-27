import './support/scratch-dir.mjs';
/** Real controller/React with controlled owner/provider replies, never native results. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
const component = fileURLToPath(new URL('../', import.meta.url)),
  output = await fs.mkdtemp(path.join(component, '.build/tests/review-ui-'));
const bundle = path.join(output, 'test.mjs'),
  entry = path.join(output, 'entry.ts');
await fs.writeFile(
  entry,
  `export * from ${JSON.stringify(path.join(component, 'ui/workbookReviewController.ts'))};export * from ${JSON.stringify(path.join(component, 'ui/WorkbookReview.tsx'))};`
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
const ui = await import(pathToFileURL(bundle));
const { reviewProviderAvailability } = await import('./support/review-provider-availability.mjs'),
  availability = await reviewProviderAvailability();
const dom = new JSDOM('<input id="parked" value="=SUM("><div id="root"></div>', { url: 'http://localhost' });
Object.assign(globalThis, {
  window: dom.window,
  document: dom.window.document,
  HTMLElement: dom.window.HTMLElement,
  Node: dom.window.Node,
  IS_REACT_ACT_ENVIRONMENT: true,
});
const { createRoot } = await import('react-dom/client');
const defer = () => {
  let resolve, reject;
  const promise = new Promise((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
};
const flush = () => act(async () => new Promise((resolve) => setTimeout(resolve, 0)));
const props = (n) => n[Object.keys(n).find((key) => key.startsWith('__reactProps$'))];
const field = (name) => document.querySelector(`[aria-label=${JSON.stringify(name)}]`);
const button = (name) =>
  [...document.querySelectorAll('button')].find((node) => node.textContent === name && !node.closest('[hidden]'));
const click = async (name) => {
  const node = typeof name === 'string' ? button(name) : name;
  assert(node);
  assert(!node.disabled);
  await act(async () => node.click());
  await flush();
};
const change = async (node, value) => {
  await act(async () => props(node).onChange({ target: { value } }));
  await flush();
};
const owner = (scope) =>
  Object.fromEntries(['documentId', 'ownerId', 'state', 'revision'].map((key) => [key, scope[key]]));
const basic = {
  documentId: 'document',
  ownerId: 'native-owner',
  state: 'head',
  revision: 'revision',
  ownerIdentity: 'runtime',
  sheet: 'Sparse',
  selection: 'A1',
  activationEpoch: 1,
  navigationGeneration: 1,
  modeEpoch: 1,
  active: true,
  writable: true,
  sheets: [
    { name: 'Sparse', state: 'visible' },
    { name: 'Hidden', state: 'hidden' },
  ],
};
const cell = (value) => ({
  value,
  dataType:
    typeof value === 'string'
      ? 'text'
      : typeof value === 'number'
        ? 'number'
        : typeof value === 'boolean'
          ? 'boolean'
          : 'blank',
});
function provider(request) {
  const dictionary = availability.locales.find((item) => item.locale === request.locale).verified;
  return {
    schema: 1,
    requestId: request.requestId,
    locale: request.locale,
    provider: 'nspell',
    providerVersion: '2.1.5',
    providerId: availability.providerId,
    status: 'completed',
    dictionaryId: dictionary.dictionaryId,
    dictionary,
    results: request.words.map((word) => ({
      word,
      correct: !['helo', 'wrld'].includes(word),
      suggestions: word === 'helo' ? ['hello'] : word === 'wrld' ? ['world'] : [],
    })),
  };
}
function frame(data = { Sparse: { A1: cell('helo wrld') }, Hidden: {} }, patch = {}) {
  const scope = { ...structuredClone(basic), ...(patch.scope ?? {}) },
    reads = [],
    scans = [],
    applied = [],
    requests = [],
    cancelled = [],
    navigated = [];
  const result = {
    scope,
    callbacks: {
      scan: async (sheet, after, limit) => {
        scans.push({ sheet, after, limit });
        const entries = Object.entries(data[sheet] ?? {})
          .filter(([ref]) => after === undefined || Number(ref.replace(/\D/g, '')) > Number(after.replace(/\D/g, '')))
          .sort((a, b) => Number(a[0].replace(/\D/g, '')) - Number(b[0].replace(/\D/g, '')));
        const selected = entries.slice(0, limit);
        return {
          receipt: owner(scope),
          sheet,
          ...(after !== undefined ? { after } : {}),
          cells: Object.fromEntries(selected),
          nextCursor: entries.length > selected.length ? selected.at(-1)[0] : null,
        };
      },
      read: async (sheet, reference) => {
        reads.push({ sheet, cell: reference });
        return { receipt: owner(scope), sheet, cell: reference, value: data[sheet]?.[reference] };
      },
      availability: async () => structuredClone(availability),
      check: async (request) => {
        requests.push(request);
        return provider(request);
      },
      cancel: async (id) => {
        cancelled.push(id);
      },
      apply: async (commands) => {
        applied.push(commands);
      },
      navigate: (sheet, reference) => navigated.push({ sheet, cell: reference }),
      ...(patch.callbacks ?? {}),
    },
  };
  return { frame: result, data, reads, scans, applied, requests, cancelled, navigated };
}
const check = (name, fn) => {
  test(name, { timeout: 20000 }, async (t) => {
    await fn(t);
  });
};
const controller = (t, h) => {
  const value = ui.createWorkbookReviewController();
  value.setFrame(h.frame);
  t.after(() => value.dispose());
  return value;
};
async function mount(t, h, options = {}) {
  const root = createRoot(document.getElementById('root'));
  let configuration = options;
  const render = async (value, patch = {}) => {
    configuration = { ...configuration, ...patch };
    await act(async () => root.render(React.createElement(ui.WorkbookReview, { frame: value, ...configuration })));
    await flush();
  };
  await render(h.frame);
  t.after(async () => {
    await act(async () => root.unmount());
    await flush();
  });
  return { render, root };
}

check(
  'Statistics count sparse accepted records on hidden sheets and label non-formula record types separately from result freshness',
  async (t) => {
    const h = frame({
        Sparse: {
          A1: cell('hello'),
          A2: cell(5),
          A3: cell(true),
          A4: { ...cell(null), styleId: 2 },
          A5: { value: '#N/A', dataType: 'error', error: '#N/A' },
          A6: cell('#N/A'),
          A7: {
            value: 'old',
            formula: '"hello"',
            calculation: { status: 'calculated', value: 'hello', revision: 'revision' },
          },
          A8: {
            value: 'old',
            formula: '"hello"',
            cachedResult: { authority: 'stored', status: 'calculated', value: 'old' },
            calculation: { status: 'stored' },
          },
          A9: {
            ...cell(7),
            calculation: { status: 'stored', pending: true },
            cachedResult: { authority: 'stored', status: 'calculated', value: 7 },
          },
          A10: {
            value: null,
            formula: 'UNKNOWN()',
            calculation: { status: 'unresolved', error: 'unsupported', revision: 'revision' },
          },
          A11: {
            value: null,
            formula: 'A11',
            calculation: { status: 'circular', error: 'cycle', revision: 'revision' },
          },
          A12: {
            value: null,
            formula: '1/0',
            calculation: { status: 'error', error: '#DIV/0!', revision: 'revision' },
          },
          A13: {
            value: 'wrong',
            formula: '"helo"',
            calculation: { status: 'calculated', value: 'helo', revision: 'old' },
          },
        },
        Hidden: { B1: cell('wrld') },
      }),
      c = controller(t, h);
    await c.scanStatistics();
    const s = c.getSnapshot();
    assert(s.statisticsComplete);
    assert.equal(s.statisticsSheets, 2);
    assert.deepEqual(
      h.scans.map((item) => item.sheet),
      ['Sparse', 'Hidden']
    );
    assert.equal(s.statistics.authoredCells, 14);
    assert.equal(s.statistics.authoredBlankCells, 1);
    assert.equal(s.statistics.formulaCells, 6);
    assert.deepEqual(s.statistics.literalKinds, { blank: 0, text: 3, number: 2, boolean: 1, error: 1, unavailable: 0 });
    assert.deepEqual(s.statistics.formulaStates, {
      calculated: 1,
      stored: 1,
      pending: 0,
      unresolved: 1,
      error: 1,
      circular: 1,
      unqualified: 1,
    });
    assert.equal(s.statistics.results.pending.number, 1);
    assert.deepEqual(h.applied, []);
    await mount(t, h);
    await click('Scan workbook statistics');
    for (const [label, count] of [
      ['Text records without formulas', 3],
      ['Number records without formulas', 2],
      ['Boolean records without formulas', 1],
      ['Error records without formulas', 1],
      ['Records with explicit protection markers', 0],
    ])
      assert.equal(field(label).textContent, String(count));
    assert.equal(field('Literal numbers'), null);
    assert.equal(field('Known protected records'), null);
    assert.equal(field('Stored or pending results, including non-formula outputs').textContent, '2');
    assert.match(field('Workbook Review').textContent, /can include cached output records/);
    assert.match(
      field('Workbook Review').textContent,
      /read-only workbook does not mean each record has an explicit protection marker/
    );
    assert.deepEqual(h.applied, []);
  }
);
check(
  'Statistics continuation covers every accepted record once and never calls a partial scan complete',
  async (t) => {
    const h = frame(
        { Sparse: {} },
        {
          scope: { sheets: [{ name: 'Sparse', state: 'visible' }] },
          callbacks: {
            scan: async (sheet, after, limit) => {
              h.scans.push({ sheet, after, limit });
              const n = after ? Number(after.slice(1)) + 1 : 1;
              return {
                receipt: owner(h.frame.scope),
                sheet,
                ...(after ? { after } : {}),
                cells: { ['A' + n]: cell(n) },
                nextCursor: n === 130 ? null : 'A' + n,
              };
            },
          },
        }
      ),
      c = controller(t, h);
    await c.scanStatistics();
    assert.equal(c.getSnapshot().statistics.authoredCells, 64);
    assert(!c.getSnapshot().statisticsComplete);
    await c.scanStatistics();
    assert.equal(c.getSnapshot().statistics.authoredCells, 128);
    await c.scanStatistics();
    assert.equal(c.getSnapshot().statistics.authoredCells, 130);
    assert(c.getSnapshot().statisticsComplete);
    assert.equal(h.scans.length, 130);
    assert.equal(new Set(h.scans.map((value) => value.after)).size, 130);
  }
);
check(
  'Within-cell spelling continuation yields all122 exact original spans without repeating or dropping words',
  async (t) => {
    const text = Array(122).fill('wrld').join(' '),
      h = frame({
        Sparse: {
          A1: cell(text),
          A2: {
            value: 'helo',
            formula: '"helo"',
            calculation: { status: 'calculated', value: 'helo', revision: 'revision' },
          },
          A3: cell('hello'),
        },
        Hidden: {},
      }),
      c = controller(t, h);
    c.setMode('spelling');
    await c.scanSpelling();
    const first = c.getSnapshot();
    assert.equal(first.findings.length, 100);
    assert.equal(first.findingsFound, 100);
    assert(!first.spellingComplete);
    assert.equal(first.spellingRecords, 1);
    assert.equal(first.findings[99].finding.token.start, 495);
    await c.scanSpelling();
    const second = c.getSnapshot();
    assert.equal(second.findings.length, 22);
    assert.equal(second.findingsFound, 122);
    assert(second.spellingComplete);
    assert.equal(second.spellingRecords, 3);
    assert.equal(second.wordsChecked, 123);
    const spans = [...first.findings, ...second.findings].map((entry) => entry.finding.token);
    assert.deepEqual(
      spans,
      Array.from({ length: 122 }, (_, index) => ({ text: 'wrld', start: index * 5, end: index * 5 + 4 }))
    );
    assert(h.requests.every((value) => value.words.length <= 100));
    assert.deepEqual(h.applied, []);
  }
);
check(
  'Unicode offsets, punctuation, formula-looking/error-looking literal text and explicit numeric/oversized token exclusions survive review',
  async (t) => {
    const text = '😀 helo cafe\u0301 wrld',
      h = frame({
        Sparse: { A1: cell(text), A2: cell('=helo #N/A'), A3: cell('a'.repeat(300) + ' 123 A1 wrld') },
        Hidden: {},
      }),
      c = controller(t, h);
    c.setMode('spelling');
    await c.scanSpelling();
    const s = c.getSnapshot();
    assert(s.spellingComplete);
    assert.deepEqual(
      s.findings.filter((entry) => entry.finding.target.cell === 'A1').map((entry) => entry.finding.token),
      [
        { text: 'helo', start: 3, end: 7 },
        { text: 'wrld', start: 14, end: 18 },
      ]
    );
    assert(s.findings.some((entry) => entry.finding.target.cell === 'A2' && entry.finding.token.start === 1));
    assert(s.findings.some((entry) => entry.finding.target.cell === 'A3'));
    assert.equal(s.excludedTokens, 3);
    assert(h.requests.some((value) => value.words.includes('cafe\u0301')));
    assert.deepEqual(h.applied, []);
  }
);
check(
  'Non-formula pending/spill/protected records and cached strings are excluded from spelling without becoming current text',
  async (t) => {
    const h = frame({
        Sparse: {
          A1: cell('helo'),
          A2: { ...cell('helo'), calculation: { status: 'stored', pending: true } },
          A3: { ...cell('helo'), spill: { schema: 1, master: 'A3', range: 'A3:A4', role: 'follower' } },
          A4: { ...cell('helo'), protected: true },
          A5: { ...cell('helo'), formulaType: 'array' },
          A6: { ...cell('helo'), calculation: { status: 'calculated', value: 'helo', revision: 'revision' } },
          A7: cell(5),
          A8: { value: '#N/A', dataType: 'error' },
        },
        Hidden: {},
      }),
      c = controller(t, h);
    c.setMode('spelling');
    await c.scanSpelling();
    assert.equal(c.getSnapshot().findings.length, 1);
    assert.equal(c.getSnapshot().excludedCells, 7);
    assert.equal(c.getSnapshot().wordsChecked, 1);
    assert.deepEqual(h.applied, []);
  }
);
check(
  'Captured owner/head/selection/sheet/activation/mode and inactive transitions mask held reads before any stale cells paint',
  async (t) => {
    for (const patch of [
      { state: 'other-head', revision: 'other-revision' },
      { ownerId: 'replacement', ownerIdentity: 'replacement-runtime' },
      { selection: 'B9', navigationGeneration: 2 },
      { sheet: 'Hidden', navigationGeneration: 2 },
      { activationEpoch: 2 },
      { modeEpoch: 2 },
      { active: false },
    ]) {
      const held = defer(),
        h = frame({}, { callbacks: { scan: () => held.promise } }),
        c = ui.createWorkbookReviewController();
      c.setFrame(h.frame);
      const old = c.scanStatistics();
      await Promise.resolve();
      const next = { ...h.frame, scope: { ...h.frame.scope, ...patch } };
      c.setFrame(next);
      assert.equal(c.getSnapshot().statistics.authoredCells, 0);
      assert(c.getSnapshot().physicalBusy);
      held.resolve({
        receipt: owner(h.frame.scope),
        sheet: 'Sparse',
        cells: { A1: cell('obsolete') },
        nextCursor: null,
      });
      await old;
      assert.equal(c.getSnapshot().statistics.authoredCells, 0);
      assert.equal(c.getSnapshot().error, null);
      await c.dispose();
    }
  }
);
check(
  'Malformed/repeated/foreign pages refuse and byte limits shrink only the exact recoverable read failure',
  async (t) => {
    const h = frame(
        {},
        {
          callbacks: {
            scan: async (sheet, after, limit) => {
              h.scans.push({ sheet, after, limit });
              if (limit > 16) throw Error('Native workbook reply exceeds transport bound');
              return {
                receipt: owner(h.frame.scope),
                sheet,
                ...(after ? { after } : {}),
                cells: { A1: cell('hello') },
                nextCursor: null,
              };
            },
          },
        }
      ),
      c = controller(t, h);
    await c.scanStatistics();
    assert(h.scans.some((value) => value.limit === 16));
    assert(c.getSnapshot().statisticsComplete);
    for (const bad of [
      {
        receipt: { ...owner(h.frame.scope), ownerId: 'foreign' },
        sheet: 'Sparse',
        cells: { A1: cell(1) },
        nextCursor: null,
      },
      { receipt: owner(h.frame.scope), sheet: 'Sparse', cells: { A1: cell(1) }, nextCursor: 'B1' },
    ]) {
      const x = frame({}, { callbacks: { scan: async () => bad } }),
        d = ui.createWorkbookReviewController();
      d.setFrame(x.frame);
      await assert.rejects(d.scanStatistics(), /owner|cursor/);
      assert(!d.getSnapshot().statisticsComplete);
      await d.dispose();
    }
    const failure = frame(
        {},
        {
          callbacks: {
            scan: async () => {
              throw Error('Other owner error');
            },
          },
        }
      ),
      d = ui.createWorkbookReviewController();
    d.setFrame(failure.frame);
    await assert.rejects(d.scanStatistics(), /Other owner error/);
    await d.dispose();
  }
);
check(
  'Missing dictionary or mismatched typed/provider reply shows unavailable/incomplete rather than empty successful spelling',
  async (t) => {
    const missing = structuredClone(availability);
    missing.locales[0].status = 'unavailable';
    delete missing.locales[0].verified;
    missing.locales[0].reason = 'Dictionary is missing.';
    const h = frame(undefined, { callbacks: { availability: async () => missing } }),
      c = controller(t, h);
    c.setMode('spelling');
    await assert.rejects(c.scanSpelling(), /missing/);
    assert(!c.getSnapshot().spellingComplete);
    assert.equal(c.getSnapshot().wordsChecked, 0);
    assert.deepEqual(h.applied, []);
    for (const mutate of [
      (reply) => {
        reply.dictionaryId = '0'.repeat(64);
      },
      (reply) => {
        reply.results[0].correct = 'false';
      },
      (reply) => {
        reply.results[0].word = 'foreign';
      },
    ]) {
      const x = frame(undefined, {
          callbacks: {
            check: async (request) => {
              const result = provider(request);
              mutate(result);
              return result;
            },
          },
        }),
        d = ui.createWorkbookReviewController();
      d.setFrame(x.frame);
      d.setMode('spelling');
      await assert.rejects(d.scanSpelling());
      assert.equal(d.getSnapshot().findings.length, 0);
      assert(!d.getSnapshot().spellingComplete);
      assert.deepEqual(x.applied, []);
      await d.dispose();
    }
  }
);
check(
  'Locale change cancels the original provider token and late results cannot populate the new review',
  async (t) => {
    const held = defer(),
      h = frame(undefined, {
        callbacks: {
          check: (request) => {
            h.requests.push(request);
            return held.promise;
          },
        },
      }),
      c = controller(t, h);
    c.setMode('spelling');
    const old = c.scanSpelling();
    while (!h.requests.length) await new Promise((resolve) => setImmediate(resolve));
    c.setLocale('en-GB');
    await new Promise((resolve) => setImmediate(resolve));
    assert.deepEqual(h.cancelled, [h.requests[0].requestId]);
    held.resolve(provider(h.requests[0]));
    await old;
    assert.equal(c.getSnapshot().findings.length, 0);
    assert.equal(c.getSnapshot().locale, 'en-GB');
    assert.equal(c.getSnapshot().wordsChecked, 0);
    assert.equal(c.getSnapshot().error, null);
    assert.deepEqual(h.applied, []);
  }
);
check(
  'Explicit Cancel keeps only committed statistics and resumes after that exact cursor without overlapping physical reads',
  async (t) => {
    const held = defer();
    let calls = 0,
      live = 0,
      maximum = 0;
    const h = frame(
        {},
        {
          scope: { sheets: [{ name: 'Sparse', state: 'visible' }] },
          callbacks: {
            scan: async (sheet, after) => {
              calls++;
              live++;
              maximum = Math.max(maximum, live);
              try {
                if (calls === 1)
                  return { receipt: owner(h.frame.scope), sheet, cells: { A1: cell('hello') }, nextCursor: 'A1' };
                if (calls === 2) return await held.promise;
                assert.equal(after, 'A1');
                return { receipt: owner(h.frame.scope), sheet, after, cells: { A2: cell('world') }, nextCursor: null };
              } finally {
                live--;
              }
            },
          },
        }
      ),
      c = controller(t, h),
      work = c.scanStatistics();
    while (calls < 2) await new Promise((resolve) => setImmediate(resolve));
    await c.cancel();
    assert.equal(c.getSnapshot().statistics.authoredCells, 1);
    assert(!c.getSnapshot().statisticsComplete);
    assert.match(c.getSnapshot().incompleteReason, /cancelled/);
    const joining = c.scanStatistics();
    assert.equal(calls, 2);
    held.resolve({
      receipt: owner(h.frame.scope),
      sheet: 'Sparse',
      after: 'A1',
      cells: { A2: cell('world') },
      nextCursor: null,
    });
    await Promise.all([work, joining]);
    assert.equal(c.getSnapshot().statistics.authoredCells, 1);
    await c.scanStatistics();
    assert.equal(c.getSnapshot().statistics.authoredCells, 2);
    assert(c.getSnapshot().statisticsComplete);
    assert.equal(maximum, 1);
  }
);
check(
  'Confirmed correction rereads the exact literal and accepts one expected-value command without interpreting a leading equals sign',
  async (t) => {
    const h = frame({ Sparse: { A1: cell('=helo') }, Hidden: {} }),
      c = controller(t, h);
    c.setMode('spelling');
    await c.scanSpelling();
    const finding = c.getSnapshot().findings[0];
    assert.deepEqual(h.applied, []);
    const result = await c.changeWord(finding, 'hello');
    assert.equal(result.outcome, 'accepted');
    assert.equal(h.reads.length, 1);
    assert.deepEqual(h.applied, [
      [
        {
          type: 'setCells',
          sheet: 'Sparse',
          cells: [
            { cell: 'A1', value: '=hello', expectedValue: '=helo', expectedFormula: null, expectedRawValue: null },
          ],
        },
      ],
    ]);
    assert.equal(c.getSnapshot().findings.length, 0);
    assert.equal((await c.changeWord(finding, 'hello')).outcome, 'retired');
    assert.equal(h.applied.length, 1);
  }
);
check(
  'Changed raw/text/formula or foreign target receipt refuses a correction before the captured writer is called',
  async (t) => {
    for (const change of [
      () => ({ value: 'different', dataType: 'text' }),
      () => ({ value: 'helo wrld', dataType: 'text', rawValue: 'changed' }),
      () => ({ value: 'helo wrld', formula: '"helo wrld"' }),
      () => ({ value: 'helo wrld', dataType: 'text', calculation: { status: 'stored', pending: true } }),
    ]) {
      const h = frame(),
        c = ui.createWorkbookReviewController();
      c.setFrame(h.frame);
      c.setMode('spelling');
      await c.scanSpelling();
      const entry = c.getSnapshot().findings[0];
      h.data.Sparse.A1 = change();
      await assert.rejects(c.changeWord(entry, 'hello'));
      assert.deepEqual(h.applied, []);
      await c.dispose();
    }
    const h = frame(undefined, {
        callbacks: {
          read: async () => ({
            receipt: { ...owner(h.frame.scope), state: 'foreign' },
            sheet: 'Sparse',
            cell: 'A1',
            value: cell('helo wrld'),
          }),
        },
      }),
      c = controller(t, h);
    c.setMode('spelling');
    await c.scanSpelling();
    await assert.rejects(c.changeWord(c.getSnapshot().findings[0], 'hello'), /owner|head/);
    assert.deepEqual(h.applied, []);
  }
);
check(
  'Held correction re-read retires on navigation; accepted acknowledgement is original-owner truth and errors never land on replacements',
  async (t) => {
    const readHold = defer(),
      h = frame(undefined, { callbacks: { read: () => readHold.promise } }),
      c = controller(t, h);
    c.setMode('spelling');
    await c.scanSpelling();
    const old = c.changeWord(c.getSnapshot().findings[0], 'hello');
    await Promise.resolve();
    c.setFrame({ ...h.frame, scope: { ...h.frame.scope, selection: 'Z1', navigationGeneration: 2 } });
    readHold.resolve({ receipt: owner(h.frame.scope), sheet: 'Sparse', cell: 'A1', value: cell('helo wrld') });
    assert.equal((await old).outcome, 'retired');
    assert.deepEqual(h.applied, []);
    for (const reject of [false, true]) {
      const ack = defer(),
        x = frame(undefined, {
          callbacks: {
            apply: (commands) => {
              x.applied.push(commands);
              return ack.promise;
            },
          },
        }),
        d = ui.createWorkbookReviewController();
      d.setFrame(x.frame);
      d.setMode('spelling');
      await d.scanSpelling();
      const result = d.changeWord(d.getSnapshot().findings[0], 'hello');
      while (!x.applied.length) await new Promise((resolve) => setImmediate(resolve));
      d.setFrame({
        ...x.frame,
        scope: { ...x.frame.scope, documentId: 'new-doc', ownerId: 'new-owner', ownerIdentity: 'new-runtime' },
      });
      if (reject) {
        ack.reject(Error('Obsolete failure'));
        await assert.rejects(result, /Obsolete/);
      } else {
        ack.resolve();
        assert.deepEqual(await result, { outcome: 'accepted', runtimeCurrent: false });
      }
      assert.equal(d.getSnapshot().notice, null);
      assert.equal(d.getSnapshot().error, null);
      await d.dispose();
    }
  }
);
check(
  'Actual React pane exposes complete authority counts and readonly spelling while leaving parked input untouched',
  async (t) => {
    const h = frame(undefined, { scope: { writable: false } });
    const view = await mount(t, h),
      parked = document.getElementById('parked');
    parked.focus();
    parked.setSelectionRange(5, 5);
    assert(field('Workbook Review').hasAttribute('data-formula-draft-zone'));
    await click('Scan workbook statistics');
    assert.equal(field('Accepted records').textContent, '1');
    assert.match(field('Review status').textContent, /complete/);
    await click('Spelling');
    await click('Scan spelling');
    assert.equal(field('Spelling findings').children.length, 2);
    assert.match(field('Offline dictionary availability').textContent, /en-US: loaded dictionary-en 4.0.0/);
    assert(button('Change word').disabled);
    assert.match(button('Change word').title, /read-only/);
    assert.equal(parked.value, '=SUM(');
    assert.equal(parked.selectionStart, 5);
    assert.deepEqual(h.applied, []);
    await view.render({ ...h.frame, scope: { ...h.frame.scope, state: 'next-head', revision: 'next-revision' } });
    assert.equal(field('Spelling findings').children.length, 0);
    assert(!document.body.textContent.includes('helo'));
    assert.deepEqual(h.applied, []);
  }
);
check(
  'React correction remains awaited and single, duplicate or retained old handlers cannot edit a new owner',
  async (t) => {
    const held = defer(),
      h = frame(undefined, {
        callbacks: {
          apply: (commands) => {
            h.applied.push(commands);
            return held.promise;
          },
        },
      }),
      view = await mount(t, h);
    await click('Spelling');
    await click('Scan spelling');
    const oldButton = button('Change word'),
      submit = props(oldButton).onClick;
    await act(async () => {
      submit();
      submit();
    });
    while (!h.applied.length) await flush();
    assert.equal(h.applied.length, 1);
    assert(button('Change word').disabled);
    await view.render({
      ...h.frame,
      scope: { ...h.frame.scope, state: 'accepted-head', revision: 'accepted-revision' },
    });
    await act(async () => held.resolve());
    await flush();
    assert.match(document.body.textContent, /One Undo/);
    await view.render({
      ...h.frame,
      scope: { ...h.frame.scope, ownerId: 'replacement', ownerIdentity: 'replacement-runtime' },
    });
    await act(async () => submit());
    await flush();
    assert.equal(h.applied.length, 1);
    assert(!document.body.textContent.includes('One Undo'));
  }
);
check(
  'Repeated ribbon open intents select modes in one retained pane without scanning; stale pane handlers cannot cancel or scan a new owner',
  async (t) => {
    const h = frame(),
      view = await mount(t, h, { openIntent: { serial: 1, mode: 'spelling' } });
    assert.equal(field('Workbook Review').dataset.reviewMode, 'spelling');
    assert.deepEqual(h.scans, []);
    await click('Workbook Statistics');
    assert.equal(field('Workbook Review').dataset.reviewMode, 'statistics');
    await view.render(h.frame, { openIntent: { serial: 1, mode: 'spelling' } });
    assert.equal(field('Workbook Review').dataset.reviewMode, 'statistics');
    await view.render(h.frame, { openIntent: { serial: 2, mode: 'spelling' } });
    assert.equal(field('Workbook Review').dataset.reviewMode, 'spelling');
    assert.equal(document.querySelectorAll('[aria-label="Workbook Review"]').length, 1);
    assert.deepEqual(h.scans, []);
    const staleMode = props(button('Workbook Statistics')).onClick,
      staleScan = props(button('Scan spelling')).onClick,
      staleLocale = props(field('Review locale')).onChange;
    const held = defer(),
      next = frame(undefined, {
        scope: { documentId: 'replacement-doc', ownerId: 'replacement-owner', ownerIdentity: 'replacement-runtime' },
        callbacks: {
          check: (request) => {
            next.requests.push(request);
            return held.promise;
          },
        },
      });
    await view.render(next.frame, { openIntent: { serial: 3, mode: 'spelling' } });
    await click('Scan spelling');
    assert.equal(next.requests.length, 1);
    await act(async () => {
      staleMode();
      staleScan();
      staleLocale({ target: { value: 'en-GB' } });
    });
    await flush();
    assert.equal(field('Workbook Review').dataset.reviewMode, 'spelling');
    assert.equal(field('Review locale').value, 'en-US');
    assert.deepEqual(next.cancelled, []);
    assert.equal(next.requests.length, 1);
    await act(async () => held.resolve(provider(next.requests[0])));
    await flush();
    assert.equal(field('Spelling findings').children.length, 2);
    assert.deepEqual(next.applied, []);
  }
);
check(
  'A ribbon mode intent during strict acceptance is consumed without cancelling or retargeting the acknowledged correction',
  async (t) => {
    const ack = defer(),
      h = frame(undefined, {
        callbacks: {
          apply: (commands) => {
            h.applied.push(commands);
            return ack.promise;
          },
        },
      }),
      view = await mount(t, h, { openIntent: { serial: 1, mode: 'spelling' } });
    await click('Scan spelling');
    await click('Change word');
    assert.equal(h.applied.length, 1);
    await view.render(h.frame, { openIntent: { serial: 2, mode: 'statistics' } });
    assert.equal(field('Workbook Review').dataset.reviewMode, 'spelling');
    assert(button('Workbook Statistics').disabled);
    await act(async () => ack.resolve());
    await flush();
    assert.equal(field('Workbook Review').dataset.reviewMode, 'spelling');
    assert.deepEqual(h.cancelled, []);
    assert.equal(h.applied.length, 1);
    await view.render(h.frame, { openIntent: { serial: 3, mode: 'statistics' } });
    assert.equal(field('Workbook Review').dataset.reviewMode, 'statistics');
    assert.equal(h.applied.length, 1);
  }
);

after(async () => {
  dom.window.close();
});
