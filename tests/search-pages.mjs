import test from 'node:test';
import assert from 'node:assert/strict';
import { loadSearchPage, scanSearchQuery } from '../ui/searchPages.ts';

const value = (text) => ({ value: text });
const queryFor = (sheets, scan) =>
  scanSearchQuery({ cells: (sheet) => sheets[sheet], scan, content: (_sheet, _ref, cell) => String(cell.value) });
const read = (query, sheets, cursor, cancelled = () => false) =>
  loadSearchPage({ query, sheets, text: 'match', mode: 'raw', cursor, cancelled });

test('dense cross-sheet result pages visit every match exactly once and bound each request', async () => {
  const first = Object.fromEntries(Array.from({ length: 2500 }, (_, i) => ['A' + (i + 1), value('match first')]));
  const second = Object.fromEntries(Array.from({ length: 650 }, (_, i) => ['B' + (i + 1), value('match second')]));
  const actual = queryFor({ First: first, Second: second }),
    limits = [];
  const query = async (...args) => {
    limits.push(args[4].matchLimit);
    return actual(...args);
  };
  const pages = [],
    seen = [];
  let cursor;
  do {
    const page = await read(query, ['First', 'Second'], cursor);
    assert(page.matches.length <= 1000);
    pages.push(page);
    seen.push(...page.matches.map((m) => m.sheet + '!' + m.cell));
    cursor = page.next ?? undefined;
  } while (cursor);
  assert.deepEqual(
    pages.map((p) => p.matches.length),
    [1000, 1000, 1000, 150]
  );
  assert.equal(seen.length, 3150);
  assert.equal(new Set(seen).size, 3150);
  assert.deepEqual(seen.slice(2499, 2502), ['First!A2500', 'Second!B1', 'Second!B2']);
  assert(limits.every((n) => n >= 1 && n <= 1000));
  assert(limits.includes(500));
  const revisit = await read(query, ['First', 'Second'], pages[0].next);
  assert.deepEqual(revisit.matches, pages[1].matches);
});

test('scan fallback resumes inside a dense scan reply and traverses empty work pages', async () => {
  const sheets = {
    Sheet: Object.fromEntries(
      Array.from({ length: 9000 }, (_, i) => ['A' + (i + 1), value(i < 8192 ? 'empty' : 'match')])
    ),
  };
  const calls = [];
  const scan = async (sheet, after) => {
    const offset = after ? Number(after.slice(1)) : 0;
    calls.push(offset);
    const entries = Object.entries(sheets[sheet]).slice(offset, offset + 2048);
    return {
      cells: Object.fromEntries(entries),
      nextCursor: offset + entries.length < 9000 ? entries.at(-1)[0] : null,
    };
  };
  const page = await read(queryFor(sheets, scan), ['Sheet']);
  assert.equal(page.matches.length, 808);
  assert.equal(page.matches[0].cell, 'A8193');
  assert.equal(page.matches.at(-1).cell, 'A9000');
  assert.equal(page.next, null);
  assert.deepEqual(calls, [0, 2048, 4096, 6144, 8192]);
  const dense = { Sheet: Object.fromEntries(Array.from({ length: 2500 }, (_, i) => ['A' + (i + 1), value('match')])) };
  const denseScan = async (sheet, after) => {
    const offset = after ? Number(after.slice(1)) : 0;
    const entries = Object.entries(dense[sheet]).slice(offset, offset + 2048);
    return {
      cells: Object.fromEntries(entries),
      nextCursor: offset + entries.length < 2500 ? entries.at(-1)[0] : null,
    };
  };
  const p1 = await read(queryFor(dense, denseScan), ['Sheet']);
  const p2 = await read(queryFor(dense, denseScan), ['Sheet'], p1.next);
  const p3 = await read(queryFor(dense, denseScan), ['Sheet'], p2.next);
  assert.deepEqual([p1.matches.length, p2.matches.length, p3.matches.length], [1000, 1000, 500]);
  assert.equal(p2.matches[0].cell, 'A1001');
  assert.equal(p3.matches[0].cell, 'A2001');
});

test('cancelled pending query and cancelled work yield never publish partial result pages', async () => {
  let complete,
    stale = false,
    calls = 0;
  const pending = read(
    async () => {
      calls++;
      return new Promise((resolve) => (complete = resolve));
    },
    ['Sheet'],
    undefined,
    () => stale
  );
  await Promise.resolve();
  stale = true;
  complete({ cells: { A1: value('match') }, matches: 1, nextCursor: null });
  assert.equal(await pending, null);
  assert.equal(calls, 1);
  stale = false;
  calls = 0;
  const yielded = await loadSearchPage({
    sheets: ['Sheet'],
    text: 'match',
    mode: 'raw',
    cancelled: () => stale,
    query: async () => {
      calls++;
      return { cells: {}, matches: 0, nextCursor: 'A10' };
    },
    yieldWork: async () => {
      stale = true;
    },
  });
  assert.equal(yielded, null);
  assert.equal(calls, 1);
});

test('owner replies cannot silently omit counted matches, overflow or repeat cursors', async () => {
  const invalid = [
    { cells: { A1: value('match') }, matches: 2, nextCursor: 'A20' },
    {
      cells: Object.fromEntries(Array.from({ length: 1001 }, (_, i) => ['A' + (i + 1), value('match')])),
      matches: 1001,
      nextCursor: 'A1001',
    },
    { cells: { A2: value('match'), A1: value('match') }, matches: 2, nextCursor: null },
    { cells: {}, matches: 0, nextCursor: 'A1' },
  ];
  for (let i = 0; i < invalid.length; i++)
    await assert.rejects(
      read(async () => invalid[i], ['Sheet'], i === 3 ? { sheet: 0, after: 'A1' } : undefined),
      /omitted|exceeded|unordered|advancing/
    );
});
