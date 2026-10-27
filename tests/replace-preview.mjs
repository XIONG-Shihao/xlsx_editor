import test from 'node:test';
import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { fileURLToPath } from 'node:url';
const bundle = await build({
  entryPoints: [fileURLToPath(new URL('../ui/replacePreview.ts', import.meta.url))],
  bundle: true,
  write: false,
  platform: 'browser',
  format: 'esm',
  logLevel: 'silent',
});
const { replacementForMatch, collectReplacementPreview, replacementPreviewPage, REPLACE_PREVIEW_LIMITS } = await import(
  'data:text/javascript;base64,' +
    Buffer.from(bundle.outputFiles[0].text + '\n//# sourceURL=replace-preview-model.js').toString('base64')
);
const binding = {
  documentId: 'original-document',
  ownerIdentity: 'actual-owner',
  state: 'accepted-head',
  sheet: 'First',
  selection: 'A1:A4',
  selectedCell: 'A1',
  activationEpoch: 1,
  navigationGeneration: 2,
  modeGeneration: 3,
  specialSelectionEpoch: 4,
  mode: 'normal',
  specialSelection: null,
};
const config = (overrides = {}) => ({
  text: 'match',
  replacement: 'done',
  mode: 'raw',
  sheets: ['First'],
  literal: false,
  ...overrides,
});
const cell = (value, extra = {}) => ({
  value,
  rawValue: typeof value === 'string' ? value : value === null ? null : String(value),
  ...extra,
});
const row = (value, options = {}) => replacementForMatch({ sheet: 'First', cell: 'A1', value }, config(options));
function content(c, mode) {
  if (mode === 'formula') return c.formula ? '=' + c.formula : '';
  if (mode === 'raw') return String(c.rawValue ?? c.value ?? '');
  return c.formula ? (c.calculation?.value ?? c.calculation?.error ?? '') : String(c.value ?? '');
}
function queryFor(sheets, calls = []) {
  return async (sheet, text, mode, after, options) => {
    calls.push({ sheet, text, mode, after, limit: options.matchLimit });
    const refs = Object.keys(sheets[sheet] ?? {}).filter(
      (ref) => !after || +ref.match(/\d+$/)[0] > +after.match(/\d+$/)[0]
    );
    const matches = refs.filter((ref) =>
        String(content(sheets[sheet][ref], mode)).toLowerCase().includes(text.toLowerCase())
      ),
      chosen = matches.slice(0, options.matchLimit);
    return {
      cells: Object.fromEntries(chosen.map((ref) => [ref, sheets[sheet][ref]])),
      matches: chosen.length,
      nextCursor: chosen.length < matches.length ? chosen.at(-1) : null,
    };
  };
}
const collect = (query, configuration = config(), options = {}) =>
  collectReplacementPreview({ binding, configuration, query, isCurrent: () => true, ...options });

test('literal strings retain type and exact dollar, formula-looking, Boolean/error-looking/apostrophe/padded text', () => {
  for (const value of ['00123', '=SUM(A1)', 'TRUE', '#N/A', "'literal", '雪 😀 match']) {
    const result = row(cell(value), { text: value, replacement: '$& $1 =TRUE #N/A 00123' });
    assert.equal(result.before, value);
    assert.equal(result.after, '$& $1 =TRUE #N/A 00123');
    assert.equal(result.reason, null);
    assert.deepEqual(result.edit, {
      sheet: 'First',
      cell: 'A1',
      expectedValue: value,
      expectedFormula: null,
      expectedRawValue: value,
      value: '$& $1 =TRUE #N/A 00123',
    });
    assert.equal(result.edit.formula, undefined);
    assert.equal(result.edit.valueType, undefined);
  }
  const literal = row(cell('TRUE'), { text: 'TRUE', replacement: 'FALSE', literal: true });
  assert.equal(literal.edit.value, 'FALSE');
  assert.equal(typeof literal.edit.value, 'string');
  assert(Object.isFrozen(literal));
  assert(Object.isFrozen(literal.edit));
  assert.equal(row(cell('Match match'), { replacement: '$&' }).after, '$& $&');
  assert.match(row(cell(null)).reason, /Blank/);
  assert.match(row(cell('same'), { text: 'same', replacement: 'same' }).reason, /No effective/);
  assert.throws(() => row(cell('safe'), { replacement: '\u0000' }), /configuration/);
  assert.match(row(cell('x'.repeat(32767)), { text: 'x', replacement: 'xx' }).reason, /limits/);
});

test('formula mode uses the authored equals prefix and ordinary/owned master proposals retain exact source guards', () => {
  const source = cell(12, {
      formula: 'SUM(A1:A3)',
      rawValue: '12.0',
      calculation: { status: 'stored', pending: true },
    }),
    result = row(source, { mode: 'formula', text: '=SUM', replacement: '=AVERAGE' });
  assert.equal(result.before, '=SUM(A1:A3)');
  assert.equal(result.after, '=AVERAGE(A1:A3)');
  assert.deepEqual(result.edit, {
    sheet: 'First',
    cell: 'A1',
    formula: 'AVERAGE(A1:A3)',
    expectedFormula: 'SUM(A1:A3)',
    expectedValue: 12,
    expectedRawValue: '12.0',
  });
  const owned = row(
    cell(1, {
      formula: 'SEQUENCE(2,2)',
      formulaType: 'array',
      spill: { schema: 1, master: 'A1', range: 'A1:B2', role: 'master' },
    }),
    { mode: 'formula', text: '2,2', replacement: '3,2' }
  );
  assert.equal(owned.edit.formula, 'SEQUENCE(3,2)');
  assert.match(row(source, { mode: 'formula', text: '=', replacement: '' }).reason, /retain/);
  assert.match(row(source, { mode: 'formula', text: 'SUM', replacement: 'WEBSERVICE' }).reason, /supported/);
  assert.match(row(source, { mode: 'formula', text: 'SUM', replacement: 'x'.repeat(8192) }).reason, /limits/);
  assert.match(row(source, { mode: 'formula', text: 'SUM', replacement: 'AVERAGE', literal: true }).reason, /Literal/);
  for (const value of [
    cell(914, { formulaOutput: { schema: 1, master: 'R1000', range: 'A1:R1000' } }),
    cell(1, { formula: 'UNKNOWN()', formulaType: 'array' }),
    cell(2, { spill: { schema: 1, master: 'Z1', range: 'A1:Z1', role: 'follower' } }),
    cell(3, { scenarioResult: true }),
    cell('match', { protected: true }),
  ])
    assert.equal(row(value, { mode: 'formula', text: 'UNKNOWN', replacement: 'SUM' }).edit, null);
});

test('display/raw calculated and stored results are explicit skipped rows and never become replacement labels', () => {
  for (const mode of ['display', 'raw'])
    for (const state of ['calculated', 'stored', 'unresolved', 'circular', 'error']) {
      const result = row(
        cell(103, {
          formula: 'A1+1',
          rawValue: '103.0',
          calculation: { status: state, pending: state === 'stored', value: 103, error: '#N/A' },
          cachedResult: { authority: 'stored', status: 'calculated', value: 103 },
        }),
        { mode, text: '103', replacement: '6' }
      );
      assert.equal(result.edit, null);
      assert.equal(result.after, null);
      assert.match(result.reason, /results/);
    }
  assert.equal(
    row(cell(914, { formulaOutput: { schema: 1, master: 'R1100', range: 'A1:R1100' } }), {
      text: '914',
      replacement: '6',
    }).edit,
    null
  );
});

test('numbers and booleans change only within exact original types; genuine errors remain distinct from literal error-looking text', () => {
  assert.equal(row(cell(12, { rawValue: '12.00' }), { text: '12', replacement: '14' }).edit.value, 14);
  assert.equal(row(cell(0, { rawValue: '0' }), { text: '0', replacement: '-0' }).after, '-0');
  assert(Object.is(row(cell(0, { rawValue: '0' }), { text: '0', replacement: '-0' }).edit.value, -0));
  for (const replacement of ['TRUE', '00123wrong', '1e999', '9007199254740993'])
    assert.equal(row(cell(1), { text: '1', replacement }).edit, null);
  assert.match(
    row(cell(9007199254740992, { rawValue: '9007199254740993' }), { text: '993', replacement: '992' }).reason,
    /precision/
  );
  assert.match(row(cell(12, { format: 2 }), { mode: 'display', text: '12', replacement: '13' }).reason, /Formatted/);
  assert.equal(
    row(cell(12, { rawValue: '12.00', format: 2 }), { mode: 'display', text: '12', replacement: '13' }).edit.value,
    13
  );
  assert.equal(row(cell(true), { text: 'true', replacement: 'FALSE' }).edit.value, false);
  assert.equal(row(cell(false), { text: 'false', replacement: '1' }).edit, null);
  const genuine = row(cell('#N/A', { dataType: 'error' }), { text: '#N/A', replacement: '#DIV/0!' });
  assert.equal(genuine.edit.valueType, 'error');
  assert.equal(genuine.edit.expectedRawValue, '#N/A');
  assert.equal(
    row(cell('#N/A', { dataType: 'text' }), { text: '#N/A', replacement: '#DIV/0!' }).edit.valueType,
    undefined
  );
  assert.equal(row(cell('#N/A', { dataType: 'error' }), { text: '#N/A', replacement: 'ordinary' }).edit, null);
});

test('complete 1004-row plus second-sheet collection clamps owner replies, visits every match and normalizes1000-cell commands', async () => {
  const sheets = {
      First: Object.fromEntries(Array.from({ length: 1004 }, (_, i) => ['A' + (i + 1), cell('match ' + i)])),
      Second: Object.fromEntries(Array.from({ length: 17 }, (_, i) => ['B' + (i + 1), cell('match second ' + i)])),
    },
    calls = [],
    progress = [];
  const source = structuredClone(sheets),
    preview = await collect(queryFor(sheets, calls), config({ sheets: ['First', 'Second'] }), {
      progress: (value) => progress.push(value),
    });
  assert.equal(preview.matched, 1021);
  assert.equal(preview.editable, 1021);
  assert.equal(preview.skipped, 0);
  assert.deepEqual(
    preview.commands.map((c) => [c.sheet, c.cells.length]),
    [
      ['First', 1000],
      ['First', 4],
      ['Second', 17],
    ]
  );
  assert(calls.every((c) => c.limit <= 100 && c.limit > 0));
  assert.equal(preview.rows[1003].cell, 'A1004');
  assert.equal(preview.rows[1004].sheet, 'Second');
  assert.equal(new Set(preview.rows.map((r) => r.sheet + '!' + r.cell)).size, 1021);
  assert.deepEqual(sheets, source);
  let offset = 0;
  const all = [];
  do {
    const page = replacementPreviewPage(preview, offset);
    assert(page.rows.length <= 100);
    all.push(...page.rows);
    offset = page.nextOffset;
  } while (offset !== null);
  assert.deepEqual(all, preview.rows);
  assert.deepEqual(progress.at(-1), { matched: 1021, editable: 1021, skipped: 0 });
  for (const value of [
    preview,
    preview.binding,
    preview.configuration,
    preview.configuration.sheets,
    preview.rows,
    preview.rows[0],
    preview.rows[0].edit,
    preview.commands,
    preview.commands[0],
    preview.commands[0].cells,
  ])
    assert(Object.isFrozen(value));
  assert.throws(() => {
    preview.commands[0].cells[0].value = 'forged';
  }, TypeError);
});

test('exact10000 obtains a further real end read;10001 and late bad pages refuse the whole preview', async () => {
  const values = Object.fromEntries(Array.from({ length: 10000 }, (_, i) => ['A' + (i + 1), cell('match')])),
    calls = [];
  const preview = await collect(queryFor({ First: values }, calls));
  assert.equal(preview.matched, 10000);
  assert(
    calls.some((c) => c.after === 'A10000'),
    'Additional accepted end read'
  );
  values.A10001 = cell('match');
  await assert.rejects(collect(queryFor({ First: values })), /10,000/);
  let count = 0;
  await assert.rejects(
    collect(async (_s, _t, _m, after) => {
      count++;
      if (!after) return { cells: { A1: cell('match') }, matches: 1, nextCursor: 'A1' };
      return { cells: { A2: cell('match') }, matches: 2, nextCursor: null };
    }),
    /incomplete/
  );
  assert.equal(count, 2);
});

test('cancelled or retired actual pending query and work/progress callbacks publish no partial preview', async () => {
  for (const retire of ['cancel', 'owner']) {
    const controller = new AbortController();
    let current = true,
      resolve;
    const promise = collect(
      () =>
        new Promise((done) => {
          resolve = done;
        }),
      config(),
      { signal: controller.signal, isCurrent: () => current }
    );
    await Promise.resolve();
    if (retire === 'cancel') controller.abort();
    else current = false;
    resolve({ cells: { A1: cell('match') }, matches: 1, nextCursor: null });
    await assert.rejects(promise, retire === 'cancel' ? /cancelled/ : /changed/);
  }
  let current = true;
  await assert.rejects(
    collect(async () => ({ cells: {}, matches: 0, nextCursor: 'A100' }), config(), {
      isCurrent: () => current,
      yieldWork: async () => {
        current = false;
      },
    }),
    /changed/
  );
  current = true;
  await assert.rejects(
    collect(queryFor({ First: { A1: cell('match') } }), config(), {
      isCurrent: () => current,
      progress: () => {
        current = false;
      },
    }),
    /changed/
  );
});

test('captured configuration and accepted cell guards are independent from caller mutations', async () => {
  const settings = config(),
    original = cell('match', { bold: true });
  let resolve;
  const promise = collect(
    () =>
      new Promise((done) => {
        resolve = done;
      }),
    settings
  );
  await Promise.resolve();
  settings.text = 'different';
  settings.replacement = 'forged';
  settings.sheets.push('Foreign');
  resolve({ cells: { A1: original }, matches: 1, nextCursor: null });
  const preview = await promise;
  original.value = 'changed later';
  assert.equal(preview.configuration.text, 'match');
  assert.equal(preview.configuration.replacement, 'done');
  assert.deepEqual(preview.configuration.sheets, ['First']);
  assert.equal(preview.rows[0].edit.expectedValue, 'match');
  assert.equal(preview.rows[0].edit.value, 'done');
  assert.equal(preview.rows[0].edit.bold, undefined);
});

test('reader count/schema/foreign/cursor/order/prototype faults cannot manufacture a partial plan', async () => {
  const invalid = [
    { cells: { A1: cell('match') }, matches: 2, nextCursor: null },
    { cells: { A2: cell('match'), A1: cell('match') }, matches: 2, nextCursor: null },
    { cells: { A1: cell('other') }, matches: 1, nextCursor: null },
    { cells: { XFE1: cell('match') }, matches: 1, nextCursor: null },
    { cells: { A1: { value: 'match', sheet: 'Foreign' } }, matches: 1, nextCursor: null },
    { cells: { A1: cell('match') }, matches: 1, nextCursor: null, sheet: 'Foreign' },
    { cells: { A1: cell('match') }, matches: 1 },
    { cells: { A1: Object.assign(Object.create({ borrowed: true }), cell('match')) }, matches: 1, nextCursor: null },
    {
      cells: Object.fromEntries(Array.from({ length: 101 }, (_, i) => ['A' + (i + 1), cell('match')])),
      matches: 101,
      nextCursor: null,
    },
  ];
  for (const reply of invalid)
    await assert.rejects(
      collect(async () => reply),
      /owner|cursor|worksheet|unordered|page|replacement/i
    );
  await assert.rejects(
    collect(async () => ({ cells: { A1: cell('match') }, matches: 1, nextCursor: 'A1' })),
    /duplicate|unordered|advancing/
  );
  assert.throws(() => replacementForMatch({ sheet: 'Foreign', cell: 'A1', value: cell('match') }, config()), /foreign/);
});

test('escaped complete action envelope and captured skipped-row budgets refuse whole plans', async () => {
  const values = Object.fromEntries(Array.from({ length: 100 }, (_, i) => ['A' + (i + 1), cell('x'.repeat(10000))]));
  await assert.rejects(collect(queryFor({ First: values }), config({ text: 'x', replacement: '\\' })), /2 MiB/);
  const skipped = Object.fromEntries(
    Array.from({ length: 600 }, (_, i) => [
      'A' + (i + 1),
      cell('match' + 'a'.repeat(32762), {
        formula: '"cache"',
        calculation: { status: 'calculated', value: 'match' + 'a'.repeat(32762) },
      }),
    ])
  );
  await assert.rejects(collect(queryFor({ First: skipped })), /16 MiB captured-row/);
  assert.equal(REPLACE_PREVIEW_LIMITS.ownerReply, 100);
  assert.equal(REPLACE_PREVIEW_LIMITS.matches, 10000);
  assert.equal(REPLACE_PREVIEW_LIMITS.actionBytes, 2 * 1024 * 1024);
});

test('preview paging rejects foreign commands/counts/rows/style fields and does not expose mutable plan aliases', async () => {
  const preview = await collect(queryFor({ First: { A1: cell('match') } }));
  const page = replacementPreviewPage(preview);
  assert.notEqual(page.rows, preview.rows);
  assert.notEqual(page.rows[0], preview.rows[0]);
  assert(Object.isFrozen(page.rows[0].edit));
  const mutants = [
    { ...preview, editable: 0, skipped: 1 },
    { ...preview, commands: [] },
    { ...preview, rows: [{ ...preview.rows[0], sheet: 'Foreign' }] },
    { ...preview, rows: [{ ...preview.rows[0], edit: { ...preview.rows[0].edit, bold: true } }] },
    { ...preview, rows: [{ ...preview.rows[0], edit: { ...preview.rows[0].edit, value: 'wrong' } }] },
  ];
  for (const mutant of mutants) assert.throws(() => replacementPreviewPage(mutant), /Invalid|Replacement/);
  for (const [offset, limit] of [
    [-1, 100],
    [2, 100],
    [0, 101],
    [0, 0],
    [0.5, 1],
  ])
    assert.throws(() => replacementPreviewPage(preview, offset, limit), /Invalid/);
  assert.deepEqual(replacementPreviewPage(preview, 1), { offset: 1, total: 1, rows: [], nextOffset: null });
});
