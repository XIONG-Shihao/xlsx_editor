/** Independently declared accepted-value/token vectors; no provider/owner/app. */
import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { createHash } from 'node:crypto';
import {
  WORKBOOK_REVIEW_LIMITS,
  classifyReviewCell,
  reviewStatistics,
  addReviewStatistics,
  reviewTokens,
  validateReviewReceipt,
  captureSpellingFinding,
  planSpellingReplacement,
} from '../core/workbook-review.mjs';
const revision = 'accepted-r1',
  context = { revision },
  receipt = {
    documentId: 'document-A',
    ownerId: 'owner-A',
    state: 'head-A',
    revision,
    locale: 'en-US',
    providerId: '1'.repeat(64),
    dictionaryId: '2'.repeat(64),
  },
  target = { sheet: "O'Brien Data", cell: 'B7' };
const files = [
  'core/workbook-review.mjs',
  'core/workbook-review.d.mts',
  'core/calculation.mjs',
  'tests/workbook-review.mjs',
];
const record = async (relative) => {
  const file = path.resolve(relative),
    bytes = await fs.readFile(file);
  return { relative, path: file, sha256: createHash('sha256').update(bytes).digest('hex'), bytes: bytes.length };
};
const before = await Promise.all(files.map(record));
const check = (name, fn) =>
  test(name, () => {
    fn();
  });
after(async () => {
  return;
  const sourceAfter = await Promise.all(files.map(record)),
    sourceChanges = before.filter((r, i) => JSON.stringify(r) !== JSON.stringify(sourceAfter[i]));
});
const frozen = (value) => {
  if (value && typeof value === 'object') {
    for (const v of Object.values(value)) frozen(v);
    Object.freeze(value);
  }
  return value;
};

check('known formula-output cache remains special without inventing an authored formula or current authority', () => {
  const formulaOutput = { schema: 1, master: 'A1', range: 'A1:A2' },
    base = { value: 'wrld', rawValue: 'wrld', dataType: 'text', formulaOutput },
    before = structuredClone(base);
  const stored = classifyReviewCell(base, context);
  assert.equal(stored.authored, 'literal');
  assert.equal(stored.special, true);
  assert.equal(stored.authority, 'stored');
  assert.equal(stored.kind, 'text');
  assert.equal(stored.spellingEligible, false);
  assert.throws(
    () => captureSpellingFinding({ receipt, target, cell: base, token: { text: 'wrld', start: 0, end: 4 } }),
    /excluded/
  );
  for (const [calculation, authority] of [
    [{ status: 'calculated', value: 'world', revision }, 'current'],
    [{ status: 'stored', pending: true, revision }, 'pending'],
    [{ status: 'calculated', value: 'world', revision: 'old' }, 'unqualified'],
  ]) {
    const result = classifyReviewCell({ ...base, calculation }, context);
    assert.equal(result.authority, authority);
    assert.equal(result.spellingEligible, false);
    assert.equal(result.formulaState, null);
  }
  const stats = reviewStatistics([base], context);
  assert.equal(stats.formulaCells, 0);
  assert.equal(stats.literalCells, 1);
  assert.equal(stats.specialCells, 1);
  assert.equal(stats.results.stored.text, 1);
  assert.equal(stats.spellingEligibleCells, 0);
  assert.deepEqual(base, before);
  const ordinary = { value: 'wrld', rawValue: 'wrld', dataType: 'text' };
  assert.equal(classifyReviewCell(ordinary, context).spellingEligible, true);
});

check('authored presence/types, source error and current/stored/pending formula authorities remain distinct', () => {
  const cells = frozen([
    undefined,
    { value: null, dataType: 'blank' },
    { value: '', dataType: 'text' },
    { value: '#N/A', dataType: 'text' },
    { value: '=SUM(A1)', dataType: 'text' },
    { value: -0, dataType: 'number', rawValue: '-0' },
    { value: true, dataType: 'boolean' },
    { value: '#DIV/0!', dataType: 'error', error: '#DIV/0!' },
    { formula: '1+1', value: 913, cachedValue: 913, calculation: { status: 'calculated', value: 2, revision } },
    {
      formula: '"helo"',
      value: 'old',
      cachedValue: 'old',
      calculation: { status: 'calculated', value: 'helo', revision },
    },
    {
      formula: '1/0',
      value: '#N/A',
      cachedValue: '#N/A',
      dataType: 'error',
      calculation: { status: 'error', error: '#DIV/0!', revision },
    },
    { formula: 'A1', value: 7, cachedValue: 7, calculation: { status: 'stored', pending: true, revision } },
    { formula: 'UNKNOWN()', value: 'old', cachedValue: 'old', calculation: { status: 'unresolved', revision } },
    { formula: 'A1', value: 5, cachedValue: 5, calculation: { status: 'circular', revision } },
    { formula: '1', value: 5, cachedValue: 5, calculation: { status: 'calculated', value: 9, revision: 'old' } },
    { formula: '1/0', value: '#N/A', cachedValue: '#N/A', dataType: 'error' },
    { value: 'wrld', dataType: 'text', spill: { schema: 1, role: 'follower', master: 'A1', range: 'A1:B2' } },
    { value: 'wrld', dataType: 'text', protected: true },
  ]);
  const saved = structuredClone(cells),
    stats = reviewStatistics(cells, context);
  assert.deepEqual(cells, saved);
  assert.equal(stats.positions, 18);
  assert.equal(stats.authoredCells, 17);
  assert.equal(stats.absentPositions, 1);
  assert.equal(stats.authoredBlankCells, 1);
  assert.equal(stats.literalCells, 8);
  assert.equal(stats.formulaCells, 8);
  assert.equal(stats.specialCells, 1);
  assert.equal(stats.protectedCells, 1);
  assert.equal(stats.spellingEligibleCells, 3);
  assert.deepEqual(stats.literalKinds, { blank: 0, text: 5, number: 1, boolean: 1, error: 1, unavailable: 0 });
  assert.deepEqual(stats.formulaStates, {
    calculated: 2,
    stored: 1,
    pending: 1,
    unresolved: 1,
    error: 1,
    circular: 1,
    unqualified: 1,
  });
  assert.equal(stats.results.literal.error, 1);
  assert.equal(stats.results.current.error, 1);
  assert.equal(stats.results.stored.error, 1);
  assert.equal(stats.results.stored.text, 1);
  assert.equal(stats.results.pending.number, 1);
  assert.equal(stats.results.current.number, 1);
  assert.equal(stats.results.unqualified.unavailable, 1);
  const split = addReviewStatistics(reviewStatistics(cells.slice(0, 9), context), cells.slice(9), context);
  assert.deepEqual(split, stats);
  assert.equal(classifyReviewCell({ value: '', dataType: 'text' }, context).authored, 'literal');
  assert.equal(classifyReviewCell({ value: null, dataType: 'blank' }, context).authored, 'blank');
});
check('special, protected, calculated and unsafe scalar admission never grants a spelling correction', () => {
  for (const cell of [
    { value: 'helo', formula: '"helo"' },
    { value: 'helo', formulaType: 'shared' },
    { value: 'helo', calculation: { status: 'calculated', value: 'helo', revision } },
    { value: 'helo', calculation: { status: 'stored', pending: true } },
    { value: 'helo', scenarioResult: true },
    { value: 'helo', protection: {} },
    { value: 123, dataType: 'number' },
    { value: '#N/A', dataType: 'error', error: '#N/A' },
  ])
    assert.equal(classifyReviewCell(cell, context).spellingEligible, false);
  assert.equal(
    classifyReviewCell({ value: 'helo', dataType: 'text' }, { ...context, special: true }).spellingEligible,
    false
  );
  assert.equal(
    classifyReviewCell({ value: 'helo', dataType: 'text' }, { ...context, protected: true }).spellingEligible,
    false
  );
  assert.equal(
    classifyReviewCell({ value: 9007199254740992, rawValue: '9007199254740993', dataType: 'number' }, context)
      .unsafeNumber,
    true
  );
  assert.equal(
    classifyReviewCell({ formula: '1/0', value: '#N/A', cachedValue: '#N/A', dataType: 'error' }, context).authority,
    'stored'
  );
  assert.equal(
    classifyReviewCell(
      { formula: '1/0', calculation: { status: 'error', error: '#N/A', revision: 'retired' } },
      context
    ).authority,
    'unqualified'
  );
  assert.equal(classifyReviewCell({ value: 5, dataType: 'text' }, context).authority, 'unqualified');
  assert.throws(() => reviewStatistics(Array(257).fill(undefined), context), /bound/);
  assert.throws(() => addReviewStatistics(reviewStatistics([], context), [], { revision: 'old' }), /revision/);
  const wrong = reviewStatistics([], context);
  wrong.positions = 1;
  assert.throws(() => addReviewStatistics(wrong, [], context), { code: 'workbook-unsupported' });
});
const unicode = "😀 Hello, l’amour can't e\u0301lan—東京 foo42 00123 ‘word’ O'Brien.\n";
const expected = [
  { text: 'Hello', start: 3, end: 8 },
  { text: 'l’amour', start: 10, end: 17 },
  { text: "can't", start: 18, end: 23 },
  { text: 'e\u0301lan', start: 24, end: 29 },
  { text: '東京', start: 30, end: 32 },
  { text: 'word', start: 46, end: 50 },
  { text: "O'Brien", start: 52, end: 59 },
];
check('Unicode punctuation/apostrophes/combining marks retain independently declared original UTF16 spans', () => {
  const page = reviewTokens(unicode);
  assert.deepEqual(page.tokens, expected);
  assert.deepEqual(page.excluded, [
    { start: 33, end: 38, reason: 'numeric-or-identifier' },
    { start: 39, end: 44, reason: 'numeric-or-identifier' },
  ]);
  assert.equal(page.nextCursor, null);
  assert(page.coverage.complete);
  for (const t of page.tokens) assert.equal(unicode.slice(t.start, t.end), t.text);
  let cursor,
    tokens = [],
    excluded = [],
    pages = 0;
  do {
    const part = reviewTokens(unicode, cursor, { maxTokens: 2, maxWork: 2 });
    assert(part.coverage.workUnits <= 2);
    tokens.push(...part.tokens);
    excluded.push(...part.excluded);
    cursor = part.nextCursor;
    assert(++pages < 100);
  } while (cursor);
  assert.deepEqual(tokens, expected);
  assert.deepEqual(excluded, page.excluded);
  assert.equal(new Set(tokens.map((t) => t.start)).size, tokens.length);
});
check('within-cell work/byte boundaries emit each token once and fit actual UTF8 JSON provider budgets', () => {
  const word = '𐐀'.repeat(200),
    value = Array(80).fill(word).join(' ');
  let cursor,
    found = [],
    pages = 0;
  do {
    const p = reviewTokens(value, cursor, { maxTokens: 100, maxWork: 1000, maxBytes: 2048 });
    assert(Buffer.byteLength(JSON.stringify(p)) <= 2048);
    assert(
      Buffer.byteLength(
        JSON.stringify({ requestId: 'a'.repeat(36), locale: 'en-US', words: p.tokens.map((t) => t.text) })
      ) <= 32768
    );
    found.push(...p.tokens);
    cursor = p.nextCursor;
    assert(++pages < 150);
  } while (cursor);
  assert.equal(found.length, 80);
  for (let i = 0; i < 80; i++) {
    assert.equal(found[i].text, word);
    assert.equal(found[i].start, i * 401);
    assert.equal(found[i].end, i * 401 + 400);
  }
  assert.throws(() => reviewTokens('word', null), /cursor/);
  assert.throws(
    () => reviewTokens('word', { offset: 1, start: null, points: 0, hasNumber: false, oversized: false, ready: false }),
    /cursor/
  );
  assert.throws(
    () =>
      reviewTokens('😀word', { offset: 1, start: null, points: 0, hasNumber: false, oversized: false, ready: false }),
    /cursor/
  );
  assert.throws(() => reviewTokens('word', undefined, { maxWork: 1 }), /bound/);
});
check('oversized word and unavailable cell explanations do not omit the following ordinary word', () => {
  const value = 'x'.repeat(300) + ' good';
  let cursor,
    tokens = [],
    excluded = [],
    pages = 0;
  do {
    const p = reviewTokens(value, cursor, { maxWork: 20 });
    if (p.nextCursor) assert.equal(p.coverage.complete, false);
    tokens.push(...p.tokens);
    excluded.push(...p.excluded);
    cursor = p.nextCursor;
    assert(++pages < 30);
  } while (cursor);
  assert.deepEqual(tokens, [{ text: 'good', start: 301, end: 305 }]);
  assert.deepEqual(excluded, [{ start: 0, end: 300, reason: 'token-too-long' }]);
  const unavailable = reviewTokens('x'.repeat(WORKBOOK_REVIEW_LIMITS.textBytes + 1));
  assert.equal(unavailable.coverage.complete, false);
  assert.equal(unavailable.coverage.reason, 'cell-text-limit');
  assert.equal(unavailable.nextCursor, null);
  assert.equal(unavailable.excluded[0].end, 1048577);
  assert.throws(() => reviewTokens('word\u0000bad'), /XML/);
  assert.throws(() => reviewTokens('word\ud800'), /XML/);
});
check(
  'one literal correction preserves exact outside text/raw guards and refuses normalized or partial token aliases',
  () => {
    const cell = frozen({ value: unicode, dataType: 'text', rawValue: 'original\rraw', bold: true }),
      finding = captureSpellingFinding({
        receipt: frozen({ ...receipt }),
        target: frozen({ ...target }),
        cell,
        token: frozen(expected[0]),
      }),
      before = structuredClone(finding);
    const command = planSpellingReplacement(finding, {
      receipt: { ...receipt },
      cell,
      replacement: 'Hallo',
      writable: true,
    });
    assert.deepEqual(command, {
      type: 'setCells',
      sheet: "O'Brien Data",
      cells: [
        {
          cell: 'B7',
          value: "😀 Hallo, l’amour can't e\u0301lan—東京 foo42 00123 ‘word’ O'Brien.\n",
          expectedValue: unicode,
          expectedFormula: null,
          expectedRawValue: 'original\rraw',
        },
      ],
    });
    assert.deepEqual(finding, before);
    assert.equal(cell.value, unicode);
    assert.throws(
      () => captureSpellingFinding({ receipt, target, cell, token: { text: 'ello', start: 4, end: 8 } }),
      /original/
    );
    const combining = captureSpellingFinding({ receipt, target, cell, token: expected[3] });
    assert.throws(
      () =>
        planSpellingReplacement(combining, {
          receipt,
          cell: { ...cell, value: unicode.normalize('NFC') },
          replacement: 'élan',
          writable: true,
        }),
      /source changed/
    );
    assert.throws(
      () =>
        planSpellingReplacement(finding, {
          receipt,
          cell: { ...cell, rawValue: 'different' },
          replacement: 'Hallo',
          writable: true,
        }),
      { code: /^(internal-invalid-field|internal-missing-field|workbook-stale)$/ }
    );
    assert.throws(() => planSpellingReplacement(finding, { receipt, cell, replacement: 'Hello', writable: true }), {
      code: 'workbook-bad-input',
    });
    const literal = { value: 'helo', dataType: 'text', rawValue: null },
      literalFinding = captureSpellingFinding({
        receipt,
        target,
        cell: literal,
        token: { text: 'helo', start: 0, end: 4 },
      });
    assert.equal(
      planSpellingReplacement(literalFinding, { receipt, cell: literal, replacement: '00123', writable: true }).cells[0]
        .value,
      '00123'
    );
    assert.equal(
      planSpellingReplacement(literalFinding, { receipt, cell: literal, replacement: '=SUM(A1)', writable: true })
        .cells[0].value,
      '=SUM(A1)'
    );
  }
);
check('all seven captured receipt fields and current writable/target authority fence corrections', () => {
  const cell = { value: 'helo', dataType: 'text' },
    finding = captureSpellingFinding({ receipt, target, cell, token: { text: 'helo', start: 0, end: 4 } });
  for (const key of Object.keys(receipt)) {
    const current = { ...receipt, [key]: key === 'locale' ? 'en-GB' : key.endsWith('Id') ? '3'.repeat(64) : 'retired' };
    assert.throws(
      () => planSpellingReplacement(finding, { receipt: current, cell, replacement: 'hello', writable: true }),
      /retired/
    );
  }
  assert.throws(() => validateReviewReceipt({ ...receipt, path: '/caller' }), /fields/);
  assert.throws(() => validateReviewReceipt(receipt, null), /fields/);
  assert.throws(
    () =>
      planSpellingReplacement({ ...finding, receipt: null }, { receipt, cell, replacement: 'hello', writable: true }),
    /fields/
  );
  assert.throws(
    () => planSpellingReplacement(finding, { receipt, cell, replacement: 'hello', writable: false }),
    /read-only/
  );
  for (const modifier of [
    { cell: { ...cell, formula: '"helo"' } },
    { cell: { ...cell, calculation: { status: 'calculated', value: 'helo', revision } } },
    { special: true },
    { protected: true },
  ])
    assert.throws(
      () => planSpellingReplacement(finding, { receipt, cell, replacement: 'hello', writable: true, ...modifier }),
      /Only|special|protected|Stored cache/
    );
  assert.throws(
    () => planSpellingReplacement(finding, { receipt, cell, replacement: 'hello\u0000', writable: true }),
    /XML/
  );
  assert.throws(
    () =>
      captureSpellingFinding({
        receipt,
        target: { sheet: 'S', cell: '$A$1' },
        cell,
        token: { text: 'helo', start: 0, end: 4 },
      }),
    /canonical/
  );
});
check('page/statistics/correction bounds are explicit without accepting a truncated correction', () => {
  const cell = { value: 'helo ' + 'x'.repeat(350000), dataType: 'text', rawValue: 'y'.repeat(350000) },
    finding = captureSpellingFinding({ receipt, target, cell, token: { text: 'helo', start: 0, end: 4 } });
  assert.throws(
    () => planSpellingReplacement(finding, { receipt, cell, replacement: 'hello', writable: true }),
    /byte bound/
  );
  assert.throws(() => reviewStatistics([{ value: 'x'.repeat(1048576), dataType: 'text' }], context), /bound/);
  assert.equal(classifyReviewCell({ value: 'x'.repeat(1048577), dataType: 'text' }, context).spellingEligible, false);
  assert.throws(() => reviewTokens('word', undefined, { maxBytes: 2047 }), /bound/);
  assert.throws(() => reviewTokens('word', undefined, { maxTokens: 101 }), /bound/);
});
