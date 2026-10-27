/** Production-source closed expression syntax and mocked captured insertion. */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';
const folder = fileURLToPath(new URL('../', import.meta.url)),
  failures = [];
const load = async (entry) => {
  const result = await build({
    entryPoints: [entry],
    bundle: true,
    write: false,
    format: 'esm',
    platform: 'node',
    logLevel: 'silent',
  });
  return import('data:text/javascript;base64,' + Buffer.from(result.outputFiles[0].text).toString('base64'));
};
const catalog = await load(path.join(folder, 'ui/functionCatalog.ts')),
  insertion = await load(path.join(folder, 'ui/functionInsertion.ts')),
  old = await load(path.join(folder, 'tests/support/financial-scalar-catalog.ts'));
const e = (value) => ({ kind: 'expression', value }),
  n = (value) => ({ kind: 'number', value }),
  r = (value) => ({ kind: 'reference', value });
const check = (name, work) =>
  test(name, async () => {
    try {
      await work();
    } catch (error) {
      failures.push({ name, error: String(error.stack ?? error) });
      throw error;
    }
  });
check('all original19 definition bodies/order/default and scalar Financial serializers remain intact', () => {
  assert.equal(catalog.FUNCTION_CATALOG.length, 26);
  assert.equal(catalog.FUNCTION_CATALOG[0].id, 'SUM');
  assert.deepEqual(catalog.FUNCTION_CATALOG.slice(0, 19), old.FUNCTION_CATALOG.slice(0, 19));
  for (const id of ['PMT', 'PV', 'FV'])
    for (const args of [
      [n(0), n(12), n(1200)],
      [r('B1'), r('B2'), r('B3')],
    ])
      assert.deepEqual(catalog.serializeFunction(id, args), old.serializeFunction(id, args));
});
check(
  'Financial-only typed raw expressions preserve operators/grouping/anchors and require real owner validation',
  () => {
    for (const id of ['PMT', 'PV', 'FV']) {
      const result = catalog.serializeFunction(id, [e(' B1/12 '), e('(B2*12)'), r('$B$3')]);
      assert.equal(result.formula, id + '(B1/12,(B2*12),$B$3)');
      assert.equal(result.requiresOwnerValidation, true);
      assert(catalog.functionDefinition(id).arguments.every((arg) => arg.allowedKinds.includes('expression')));
    }
    for (const definition of catalog.FUNCTION_CATALOG.slice(0, 19)) {
      assert(definition.arguments.every((arg) => !arg.allowedKinds.includes('expression')));
      assert.throws(() =>
        catalog.serializeFunction(
          definition.id,
          Array.from({ length: definition.minArgs }, () => e('A1+1'))
        )
      );
    }
  }
);
check(
  'nested functions/names/ranges/other sheets/coercions/Unicode whitespace and unsafe numeric meaning refuse',
  () => {
    for (const raw of [
      '',
      ' ',
      '=B1/12',
      'b1/12',
      'Rate/12',
      'SUM(B1)/12',
      'IF(TRUE,1,0)',
      'B1:B2',
      'B:B',
      'Other!B1',
      "'Cases'!B1",
      '[x.xlsx]B1',
      'B1#',
      '@B1',
      'B1^2',
      '12%',
      'TRUE+1',
      '"1"+1',
      '[1]',
      'B1\u00a0/12',
      '9007199254740993',
      '0.123456789012345678901',
      '1e400',
      '1e-999',
      'XFE1',
      'A1048577',
    ])
      assert.throws(() => catalog.serializeFunction('PMT', [e(raw), n(12), n(100)]), raw);
    assert.equal(
      catalog.serializeFunction('PMT', [e('$B$1/12'), e('B2*12'), e('-B3')]).formula,
      'PMT($B$1/12,B2*12,-B3)'
    );
  }
);
const balanced = (items) =>
  items.length === 1
    ? items[0]
    : '(' + balanced(items.slice(0, items.length / 2)) + '+' + balanced(items.slice(items.length / 2)) + ')';
check('per-expression and whole-call bounds prevent fragmented reference/node-limit bypass', () => {
  const refs = balanced(Array.from({ length: 16 }, (_, i) => 'A' + (i + 1))),
    constants = balanced(Array(16).fill('1'));
  assert.throws(() => catalog.serializeFunction('PMT', [e(refs), e(refs), r('B1')]), /reference limits/);
  assert.throws(
    () => catalog.serializeFunction('PMT', [e(constants), e(constants), e(constants), e(constants), e(constants)]),
    /grouping or cell-reference limits/
  );
  for (const raw of ['('.repeat(9) + 'B1' + ')'.repeat(9), 'B1' + ' '.repeat(1024), '1+'.repeat(20) + '1'])
    assert.throws(() => catalog.serializeFunction('PMT', [e(raw), n(12), n(100)]));
});
check('untrusted AST/objects/accessors and optional gaps cannot become formula fragments', () => {
  let calls = 0;
  for (const value of [{ kind: 'binary', operator: '+', left: 1, right: 2 }, new String('B1/12'), null, 1])
    assert.throws(() => catalog.serializeFunction('PMT', [e(value), n(12), n(100)]));
  assert.throws(() =>
    catalog.serializeFunction('PMT', [
      {
        kind: 'expression',
        get value() {
          calls++;
          return 'B1/12';
        },
      },
      n(12),
      n(100),
    ])
  );
  assert.equal(calls, 0);
  assert.throws(() => catalog.serializeFunction('PMT', [e('B1/12'), n(12), n(100), undefined, e('1')]));
  assert.throws(() => catalog.serializeFunction('PMT', [e('B1/12'), n(12), n(100), e(''), e('1')]));
});
const binding = Object.freeze({
  documentId: 'doc',
  grantDocumentId: 'grant',
  ownerIdentity: 'owner',
  ownerKey: 'owner',
  state: 'head',
  sheet: 'Cases',
  target: 'C33',
  selection: 'C33',
  navigationGeneration: 1,
  activationEpoch: 1,
});
check('one original-target strict batch is awaited and late durable acceptance retains truthful outcome', async () => {
  let current = true,
    writes = [];
  const context = {
    binding,
    isCurrent: () => current,
    availability: () => ({ active: true, readOnly: false, busy: false, literal: false, specialSelection: false }),
    readTarget: async () => ({}),
    accept: async (commands) => {
      writes.push(commands);
      current = false;
    },
  };
  const result = await insertion.insertFunction(context, { id: 'PMT', args: [e('B1/12'), e('B2*12'), r('B3')] });
  assert.equal(result.outcome, 'accepted');
  assert.equal(result.current, false);
  assert.deepEqual(writes, [
    [{ type: 'setCells', sheet: 'Cases', cells: [{ cell: 'C33', formula: 'PMT(B1/12,B2*12,B3)' }] }],
  ]);
});
check('retired target/readonly/cancel/protected target still refuse before any expression write', async () => {
  for (const options of [
    { current: false },
    { readOnly: true },
    { cell: { formulaType: 'array' } },
    { cell: { spill: { schema: 1 } } },
    { protectedReason: 'Merged target is protected' },
  ]) {
    let writes = 0;
    const context = {
      binding,
      isCurrent: () => options.current !== false,
      availability: () => ({
        active: true,
        readOnly: options.readOnly ?? false,
        busy: false,
        literal: false,
        specialSelection: false,
      }),
      readTarget: async () => options,
      accept: async () => writes++,
    };
    await assert.rejects(insertion.insertFunction(context, { id: 'PV', args: [e('B1/12'), n(12), n(-100)] }));
    assert.equal(writes, 0);
  }
  let writes = 0;
  const controller = new AbortController();
  controller.abort();
  await assert.rejects(
    insertion.insertFunction(
      {
        binding,
        isCurrent: () => true,
        availability: () => ({ active: true, readOnly: false, busy: false, literal: false, specialSelection: false }),
        readTarget: async () => ({}),
        accept: async () => writes++,
      },
      { id: 'FV', args: [e('B1/12'), n(12), n(-100)] },
      { signal: controller.signal }
    )
  );
  assert.equal(writes, 0);
});
