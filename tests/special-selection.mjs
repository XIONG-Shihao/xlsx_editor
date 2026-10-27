import './support/scratch-dir.mjs';
/** Complete Go To Special source classification, never viewport or mutation proof. */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
const component = fileURLToPath(new URL('../', import.meta.url)),
  output = await fs.mkdtemp(path.join(component, '.build/tests/special-selection-model-'));
const bundle = path.join(output, 'model.mjs');
await build({
  entryPoints: [path.join(component, 'ui/specialSelection.ts')],
  bundle: true,
  format: 'esm',
  platform: 'node',
  outfile: bundle,
  logLevel: 'silent',
});
const api = await import(pathToFileURL(bundle));
const binding = {
  documentId: 'document',
  ownerIdentity: 'original-owner',
  state: 'original-head',
  sheet: 'Data',
  selection: 'A1:A12',
  selectedCell: 'A1',
  activationEpoch: 0,
  navigationGeneration: 0,
  modeGeneration: 0,
  specialSelectionEpoch: 0,
  mode: 'normal',
  specialSelection: null,
};
const options = (kind = 'constants', types = { number: true, text: true, logical: true, error: true }) => ({
  kind,
  types,
});
const position = (ref) => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  return { row: Number(m[2]), column: [...m[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) };
};
const contained = (ref, range) => {
  const [first, last = first] = range.split(':'),
    a = position(first),
    b = position(last),
    p = position(ref);
  return p.row >= a.row && p.row <= b.row && p.column >= a.column && p.column <= b.column;
};
const sourceReader =
  (cells, calls = []) =>
  async (sheet, range) => {
    calls.push({ sheet, range });
    return {
      range,
      complete: true,
      cells: Object.fromEntries(
        Object.entries(cells)
          .filter(([ref]) => contained(ref, range))
          .map(([ref, c]) => [ref, structuredClone(c)])
      ),
    };
  };
const collect = (cells, { kind = 'constants', types, selection = binding.selection, ...patch } = {}) =>
  api.collectSpecialSelectionPreview({
    binding: { ...binding, selection },
    options: options(kind, types),
    geometry: { name: binding.sheet, merges: [] },
    read: sourceReader(cells),
    isCurrent: () => true,
    ...patch,
  });
const literal = (value) => ({
  value,
  dataType:
    value === null ? 'blank' : typeof value === 'string' ? 'text' : typeof value === 'number' ? 'number' : 'boolean',
});
const formula = (value, type = 'calculated') => ({
  value,
  formula: 'A99',
  calculation: { status: type, ...(type === 'calculated' ? { value } : {}), revision: 'accepted-calculation' },
});
const defer = () => {
  let resolve;
  const promise = new Promise((yes) => (resolve = yes));
  return { promise, resolve };
};
const check = (name, fn) => {
  test(name, async () => {
    await fn();
  });
};

check(
  'Constants masks distinguish genuine errors, error-looking/empty text and derived or authored formulas',
  async () => {
    const cells = {
      A1: literal(12),
      A2: literal('00123'),
      A3: literal(false),
      A4: { value: '#N/A', dataType: 'error' },
      A5: literal('#N/A'),
      A6: literal(''),
      A7: literal(null),
      A8: {
        value: 913,
        formula: 'SUM(A1)',
        calculation: { status: 'stored', pending: true },
        cachedResult: { authority: 'stored', status: 'calculated', value: 913 },
      },
      A9: { value: 14, dataType: 'number', formulaOutput: { schema: 1, master: 'A8', range: 'A8:A9' } },
      A10: { value: 3, scenarioResult: true },
      A11: { value: null, formulaType: 'shared', sharedFormulaId: '7' },
      A12: { value: 5, spill: { schema: 1, role: 'follower', master: 'B1', range: 'A12:B12' } },
    };
    assert.deepEqual((await collect(cells)).references, ['A1', 'A2', 'A3', 'A4', 'A5', 'A6']);
    assert.deepEqual(
      (await collect(cells, { types: { number: false, text: false, logical: false, error: true } })).references,
      ['A4']
    );
    assert.deepEqual(
      (await collect(cells, { types: { number: false, text: true, logical: false, error: false } })).references,
      ['A2', 'A5', 'A6']
    );
    assert.deepEqual((await collect(cells, { kind: 'formulas' })).references, ['A8', 'A11']);
  }
);
check(
  'Formula masks use only defined current number/text/logical/error results, including empty formula text',
  async () => {
    const cells = {
      A1: formula(5),
      A2: formula(''),
      A3: formula(true),
      A4: {
        value: '#DIV/0!',
        formula: '1/0',
        calculation: { status: 'error', error: '#DIV/0!', revision: 'accepted-calculation' },
      },
      A5: literal('#DIV/0!'),
    };
    for (const [type, expected] of [
      ['number', ['A1']],
      ['text', ['A2']],
      ['logical', ['A3']],
      ['error', ['A4']],
    ]) {
      const types = { number: false, text: false, logical: false, error: false };
      types[type] = true;
      assert.deepEqual(
        (await collect(cells, { kind: 'formulas', types, calculationRevision: 'accepted-calculation' })).references,
        expected
      );
    }
  }
);
check(
  'Readonly selection uses authored/current error type authority without borrowing the narrower Paste Values error whitelist',
  async () => {
    const cells = {
      A1: { value: '#GETTING_DATA', dataType: 'error' },
      A2: literal('#GETTING_DATA'),
      A3: {
        value: '#FIELD!',
        formula: 'A99',
        calculation: { status: 'error', error: '#FIELD!', revision: 'accepted-calculation' },
      },
    };
    assert.deepEqual(
      (await collect(cells, { types: { number: false, text: false, logical: false, error: true } })).references,
      ['A1']
    );
    assert.deepEqual(
      (
        await collect(cells, {
          kind: 'formulas',
          types: { number: false, text: false, logical: false, error: true },
          calculationRevision: 'accepted-calculation',
        })
      ).references,
      ['A3']
    );
  }
);
check(
  'All formula masks include authored pending formulas; narrower masks refuse stored caches or stale result revisions',
  async () => {
    for (const cell of [
      {
        value: 913,
        formula: 'A1',
        cachedValue: 913,
        calculation: { status: 'stored', pending: true, revision: 'accepted-calculation' },
        cachedResult: { authority: 'stored', status: 'error', error: '#N/A' },
      },
      {
        value: '#N/A',
        formula: 'A1',
        dataType: 'error',
        calculation: { status: 'unresolved', pending: true, revision: 'accepted-calculation' },
      },
      { ...formula(7), calculation: { status: 'calculated', value: 7, revision: 'retired-calculation' } },
      {
        value: '',
        dataType: 'text',
        formula: '""',
        cachedValue: '',
        calculation: { status: 'stored', revision: 'accepted-calculation' },
        cachedResult: { authority: 'stored', status: 'calculated', value: '' },
      },
      {
        value: '',
        dataType: 'text',
        formula: '""',
        cachedValue: '',
        calculation: { status: 'unresolved', revision: 'accepted-calculation' },
        cachedResult: { authority: 'stored', status: 'calculated', value: '' },
      },
    ]) {
      assert.deepEqual((await collect({ A1: cell }, { kind: 'formulas' })).references, ['A1']);
      await assert.rejects(
        collect(
          { A1: cell },
          {
            kind: 'formulas',
            types: { number: true, text: false, logical: false, error: false },
            calculationRevision: 'accepted-calculation',
          }
        ),
        /current defined result|type authority/
      );
      await assert.rejects(
        collect(
          { A1: cell },
          {
            kind: 'formulas',
            types: { number: false, text: true, logical: false, error: false },
            calculationRevision: 'accepted-calculation',
          }
        ),
        /current defined result|type authority/
      );
      assert.deepEqual((await collect({ A1: cell }, { selection: 'A1:A2', kind: 'blanks' })).references, ['A2']);
      assert.deepEqual((await collect({ A1: cell }, { selection: 'A1:A2', kind: 'constants' })).references, []);
    }
    await assert.rejects(
      collect(
        { A1: formula(7) },
        { kind: 'formulas', types: { number: true, text: false, logical: false, error: false } }
      ),
      /current defined result/
    );
  }
);
check('10000-position full rectangle reads ten original windows and includes an unloaded last target', async () => {
  const calls = [],
    p = await collect(
      { A1: literal(1), CV100: literal(2) },
      { selection: 'A1:CV100', read: sourceReader({ A1: literal(1), CV100: literal(2) }, calls) }
    );
  assert.equal(p.positions, 10000);
  assert.equal(calls.length, 10);
  assert.deepEqual(calls[0], { sheet: 'Data', range: 'A1:CV10' });
  assert.deepEqual(calls.at(-1), { sheet: 'Data', range: 'A91:CV100' });
  assert.deepEqual(p.references, ['A1', 'CV100']);
  assert(Object.isFrozen(p) && Object.isFrozen(p.binding) && Object.isFrozen(p.rows) && Object.isFrozen(p.references));
});
check('Wide source rectangles are read without omissions using at most1000 positions per window', async () => {
  const calls = [],
    p = await collect(
      { BER5: literal('tail') },
      { selection: 'A1:BER5', read: sourceReader({ BER5: literal('tail') }, calls) }
    );
  assert.equal(p.positions, 7500);
  assert.equal(calls.length, 10);
  assert.deepEqual(calls.slice(0, 2), [
    { sheet: 'Data', range: 'A1:ALL1' },
    { sheet: 'Data', range: 'ALM1:BER1' },
  ]);
  assert.deepEqual(calls.at(-1), { sheet: 'Data', range: 'ALM5:BER5' });
  assert.deepEqual(p.references, ['BER5']);
});
check(
  'A maximum-size merged Blank selection expands exactly10000 physical positions without blocking for seconds',
  async () => {
    const usedRangeReceipt = {
        documentId: binding.documentId,
        ownerId: binding.ownerIdentity,
        state: binding.state,
        revision: 'accepted-source',
      },
      bounds = { minRow: 0, minColumn: 0, maxRow: 99, maxColumn: 99 };
    const usedRange = {
      schema: 1,
      policy: 'xlsx-used-range-v1',
      status: 'complete',
      receipt: usedRangeReceipt,
      sheet: 'Data',
      native: { version: 0, sourceRevision: 'c'.repeat(64), sheetPath: 'xl/worksheets/sheet1.xml' },
      scopeMinMax: bounds,
      range: 'A1:CV100',
      lastCell: 'CV100',
      empty: false,
      provenance: {
        factsSchema: 'workbook.xlsx.used-range-facts.v1',
        declaration: { status: 'valid', ref: 'A1:CV100', bounds },
        sourceCells: { minRow: 0, minColumn: 0, maxRow: 0, maxColumn: 0 },
        sourceMerges: bounds,
        effectiveCells: { minRow: 0, minColumn: 0, maxRow: 0, maxColumn: 0 },
        sourceAxisFormatting: null,
        acceptedHistory: null,
        sourceDefaultFormatting: false,
      },
      reasons: [],
    };
    const start = performance.now(),
      p = await collect(
        { A1: literal(7) },
        {
          selection: 'A1:CV100',
          kind: 'blanks',
          geometry: { name: 'Data', merges: ['A1:CV100'] },
          usedRange,
          usedRangeReceipt,
        }
      ),
      elapsedMs = performance.now() - start;
    assert.equal(p.references.length, 10000);
    assert.equal(new Set(p.references).size, 10000);
    assert.equal(p.references[0], 'A1');
    assert.equal(p.references.at(-1), 'CV100');
    assert.equal(p.activeCell, 'A1');
    assert.equal(p.rows.find((row) => row.cell === 'A1').classification, 'Matched merged area (Absent source cell)');
    assert(elapsedMs < 3000, 'The supported 10000-position merge blocked for ' + elapsedMs + 'ms');
  }
);
check('Over-limit, reversed, disjoint, special or retired owner ranges fail before any read', async () => {
  let reads = 0;
  const reader = async () => {
    reads++;
    throw Error('must not read');
  };
  for (const patch of [
    { selection: 'A1:A10001' },
    { selection: 'B2:A1' },
    { selection: 'A1,A3' },
    { mode: 'pageLayout' },
    { ownerIdentity: null },
    { specialSelection: ['A1'] },
  ])
    await assert.rejects(collect({}, { read: reader, binding: { ...binding, ...patch } }));
  await assert.rejects(collect({}, { read: reader, isCurrent: () => false }), /initiating workbook/);
  assert.equal(reads, 0);
});
check('Missing, foreign, malformed or overlapping merge geometry refuses before any owner read', async () => {
  let reads = 0;
  const read = async () => {
    reads++;
    throw Error('must not read');
  };
  for (const geometry of [
    undefined,
    { name: 'Other', merges: [] },
    { name: 'Data' },
    { name: 'Data', merges: ['B1:A1'] },
    { name: 'Data', merges: ['A0:B1'] },
    { name: 'Data', merges: ['A1'] },
    { name: 'Data', merges: ['A1:B2', 'B2:C3'] },
    { name: 'Data', merges: ['A1:B1', 'A1:B1'] },
    { name: 'Data', merges: Array(10001).fill('A1:B1') },
  ])
    await assert.rejects(
      collect({}, { read, geometry }),
      /merge geometry|merged area|merged areas|canonical|worksheet|first to last/
    );
  assert.equal(reads, 0);
});
check('Unrelated source merges and dimension warnings do not suppress qualified unmerged source matches', async () => {
  const geometry = { name: 'Data', merges: ['D1:E2'], geometryIssues: ['The display font is unavailable'] },
    cells = { A1: literal(7), A2: { value: 13, formula: 'A1+6', calculation: { status: 'stored', pending: true } } };
  assert.deepEqual((await collect(cells, { selection: 'A1:A3', geometry })).references, ['A1']);
  assert.deepEqual((await collect(cells, { selection: 'A1:A3', kind: 'formulas', geometry })).references, ['A2']);
  assert.deepEqual((await collect(cells, { selection: 'A1:A3', kind: 'blanks', geometry })).references, ['A3']);
});
check(
  'Saved Excel merged-area oracle selects full constants/formulas/blank areas and preserves the area-aware active cursor',
  async () => {
    const geometry = { name: 'Data', merges: ['A1:B1', 'A2:B2', 'A3:B3'] },
      cells = { A1: literal(7), A3: formula(17), C3: formula('') },
      visibility = { name: 'Data', widths: [] };
    const expected = {
      constants: { references: ['A1', 'B1'], activeCell: 'A1' },
      formulas: { references: ['A3', 'B3', 'C3'], activeCell: 'A3' },
      blanks: { references: ['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3'], activeCell: 'C1' },
      visible: { references: ['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3'], activeCell: 'A1' },
    };
    for (const [kind, result] of Object.entries(expected)) {
      const p = await collect(cells, { selection: 'A1:C3', kind, geometry, visibility });
      assert.deepEqual(p.references, result.references, kind);
      assert.equal(p.activeCell, result.activeCell, kind);
      assert.equal(p.positions, 9);
      assert.equal(api.specialSelectionPreviewPage(p).total, result.references.length);
      assert(Object.isFrozen(p.geometry) && Object.isFrozen(p.geometry.merges));
    }
    for (const selection of ['A1:B3', 'A1:B4']) {
      const p = await collect(cells, { selection, kind: 'blanks', geometry });
      assert.equal(p.activeCell, 'A1', selection);
      assert.equal(p.references.length, selection.endsWith('4') ? 8 : 6);
      assert.equal(p.references[0], 'A1');
      assert.equal(p.references.at(-1), selection.endsWith('4') ? 'B4' : 'B3');
    }
  }
);
check(
  'Partial source selections close transitively around merges before complete reads, with the original owner binding retained',
  async () => {
    const geometry = { name: 'Data', merges: ['A1:B1', 'A2:B2', 'A3:B3'] },
      cells = { A1: literal(7), A3: formula(17), C3: formula('') },
      calls = [];
    const p = await collect(cells, { selection: 'B1:C3', kind: 'blanks', geometry, read: sourceReader(cells, calls) });
    assert.equal(p.binding.selection, 'B1:C3');
    assert.equal(p.range, 'A1:C3');
    assert.equal(p.positions, 9);
    assert.equal(p.activeCell, 'C1');
    assert.deepEqual(calls, [{ sheet: 'Data', range: 'A1:C3' }]);
    assert.equal(api.specialSelectionPreviewPage(p).total, 8);
    const chain = await collect(
      { A1: literal(7), D3: literal(9) },
      { selection: 'B2:C3', geometry: { name: 'Data', merges: ['A1:B2', 'C1:D1', 'D3:E4'] } }
    );
    assert.equal(chain.range, 'A1:E4');
    assert.equal(chain.positions, 20);
    assert.deepEqual(chain.references, ['A1', 'B1', 'A2', 'B2', 'D3', 'E3', 'D4', 'E4']);
    let reads = 0;
    await assert.rejects(
      collect(
        {},
        {
          selection: 'A1',
          geometry: { name: 'Data', merges: ['A1:CW100'] },
          read: async () => {
            reads++;
            throw Error('must not read');
          },
        }
      ),
      /10,000 source positions/
    );
    assert.equal(reads, 0);
  }
);
check('Merged constant masks and full areas crossing read windows remain exact, unique and source-bound', async () => {
  const geometry = { name: 'Data', merges: ['A1:B1', 'A2:B2', 'A3:B3', 'A4:B4'] },
    cells = { A1: literal(7), A2: { value: '#N/A', dataType: 'error' }, A3: literal('#N/A'), A4: literal(false) };
  for (const [type, expected] of [
    ['number', ['A1', 'B1']],
    ['error', ['A2', 'B2']],
    ['text', ['A3', 'B3']],
    ['logical', ['A4', 'B4']],
  ]) {
    const types = { number: false, text: false, logical: false, error: false };
    types[type] = true;
    const p = await collect(cells, { selection: 'A1:B4', types, geometry });
    assert.deepEqual(p.references, expected);
    assert.equal(p.activeCell, expected[0]);
  }
  const calls = [],
    p = await collect(
      { A1000: literal(7) },
      {
        selection: 'A1:A1004',
        geometry: { name: 'Data', merges: ['A1000:A1002'] },
        read: sourceReader({ A1000: literal(7) }, calls),
      }
    );
  assert.deepEqual(p.references, ['A1000', 'A1001', 'A1002']);
  assert.equal(p.activeCell, 'A1000');
  assert.deepEqual(
    calls.map((c) => c.range),
    ['A1:A1000', 'A1001:A1004']
  );
  assert.equal(new Set(p.references).size, 3);
  assert.equal(api.specialSelectionPreviewPage(p).total, 3);
  for (const change of [
    (q) => (q.range = 'A1:A1000'),
    (q) => (q.geometry.merges = ['A1:A10000']),
    (q) => (q.geometry.name = 'Other'),
    (q) => (q.activeCell = 'A999'),
    (q) => (q.activeCell = null),
  ]) {
    const forged = structuredClone(p);
    change(forged);
    assert.throws(() => api.specialSelectionPreviewPage(forged));
  }
});
check('Incomplete, wrong-range, foreign-cell and malformed source replies cannot produce a selection', async () => {
  const replies = [
    { range: 'A1:A12', cells: {} },
    { range: 'A1:A12', complete: false, cells: {} },
    { range: 'B1:B12', complete: true, cells: {} },
    { range: 'A1:A12', complete: true, cells: { B1: literal(1) } },
    { range: 'A1:A12', complete: true, cells: { A1: { value: NaN } } },
    { range: 'A1:A12', complete: true, cells: { A1: { value: 'x', dataType: 'number' } } },
    { range: 'A1:A12', complete: true, cells: { A1: { value: 'x', unsupported: true } } },
    { range: 'A1:A12', complete: true, cells: {}, truncated: true },
  ];
  for (const reply of replies)
    await assert.rejects(collect({}, { read: async () => reply }), /incomplete|foreign|outside|malformed|source type/);
});
check('An absent sparse record is a blank only after the entire exact range has a complete receipt', async () => {
  const cells = {
    A2: literal(null),
    A3: {
      value: '',
      formula: '""',
      calculation: { status: 'calculated', value: '', revision: 'accepted-calculation' },
    },
    A4: literal('x'),
    A5: literal(''),
    A6: literal(' '),
  };
  const p = await collect(cells, { selection: 'A1:A6', kind: 'blanks' });
  assert.deepEqual(p.references, ['A1', 'A2']);
  assert.deepEqual(
    p.rows.map((r) => r.classification),
    ['Absent source cell', 'Authored blank cell']
  );
  assert.deepEqual(
    (
      await collect(cells, {
        selection: 'A1:A6',
        kind: 'constants',
        types: { number: false, text: true, logical: false, error: false },
      })
    ).references,
    ['A4', 'A5', 'A6']
  );
  for (const cell of [
    { value: '', dataType: 'blank' },
    { value: '', dataType: 'number' },
    { value: '', dataType: 'boolean' },
    { value: null, dataType: 'text' },
  ]) {
    await assert.rejects(collect({ A1: cell }, { selection: 'A1:A2', kind: 'blanks' }), /source type/);
    await assert.rejects(collect({ A1: cell }, { selection: 'A1:A2', kind: 'constants' }), /source type/);
  }
});
check('Observed explicit Z79:Z80 selection separates absent Blanks from canonical empty Text Constants', async () => {
  const textTypes = { number: false, text: true, logical: false, error: false };
  for (const [encoding, rawValue] of [
    ['shared', null],
    ['inline', null],
    ['str', ''],
  ]) {
    const cells = { Z80: { value: '', dataType: 'text', rawValue, cachedValue: null, styleId: 0 } },
      original = structuredClone(cells),
      calls = [];
    const text = await collect(cells, { selection: 'Z79:Z80', types: textTypes, read: sourceReader(cells, calls) });
    const blanks = await collect(cells, { selection: 'Z79:Z80', kind: 'blanks', read: sourceReader(cells, calls) });
    assert.deepEqual(text.references, ['Z80'], encoding);
    assert.equal(text.activeCell, 'Z80');
    assert.equal(text.rows[0].classification, 'Text constant');
    assert.deepEqual(blanks.references, ['Z79'], encoding);
    assert.equal(blanks.activeCell, 'Z79');
    assert.equal(blanks.rows[0].classification, 'Absent source cell');
    for (const preview of [text, blanks]) {
      assert.equal(preview.scope.kind, 'selection');
      assert.equal(preview.scope.usedRange, null);
      assert.equal(preview.binding.ownerIdentity, binding.ownerIdentity);
      assert.equal(preview.binding.state, binding.state);
      assert.equal(api.specialSelectionPreviewPage(preview).total, 1);
    }
    assert.deepEqual(calls, [
      { sheet: 'Data', range: 'Z79:Z80' },
      { sheet: 'Data', range: 'Z79:Z80' },
    ]);
    assert.deepEqual(cells, original);
  }
});
check(
  'Mixed sparse and merged empty-text matches use canonical types and preserve the original closed scope',
  async () => {
    const geometry = { name: 'Data', merges: ['A1:B1'] },
      cells = {
        A1: { ...literal(''), format: '0' },
        A2: { ...literal(null), styleId: 1, format: '@' },
        A3: formula(''),
        C1: { ...literal(' '), format: ';;;' },
        C2: { ...literal(0), format: ';;;' },
        C3: { ...literal('payload'), format: '0' },
      },
      original = structuredClone(cells),
      calls = [];
    const patch = { selection: 'B1:C3', geometry, read: sourceReader(cells, calls) },
      textTypes = { number: false, text: true, logical: false, error: false };
    const text = await collect(cells, { ...patch, types: textTypes });
    assert.deepEqual(text.references, ['A1', 'B1', 'C1', 'C3']);
    const numbers = await collect(cells, {
      ...patch,
      types: { number: true, text: false, logical: false, error: false },
    });
    assert.deepEqual(numbers.references, ['C2']);
    const formulas = await collect(cells, {
      ...patch,
      kind: 'formulas',
      types: textTypes,
      calculationRevision: 'accepted-calculation',
    });
    assert.deepEqual(formulas.references, ['A3']);
    const blanks = await collect(cells, { ...patch, kind: 'blanks' });
    assert.deepEqual(blanks.references, ['A1', 'B1', 'A2', 'B2', 'B3']);
    assert.equal(
      blanks.rows.find((row) => row.cell === 'A1').classification,
      'Matched merged area (Absent source cell)'
    );
    assert.equal(blanks.rows.find((row) => row.cell === 'A2').classification, 'Authored blank cell');
    for (const preview of [text, numbers, formulas, blanks]) {
      assert.equal(preview.binding.selection, 'B1:C3');
      assert.equal(preview.binding.ownerIdentity, binding.ownerIdentity);
      assert.equal(preview.binding.state, binding.state);
      assert.equal(preview.range, 'A1:C3');
      assert.equal(preview.positions, 9);
      assert.equal(new Set(preview.references).size, preview.references.length);
      assert.equal(api.specialSelectionPreviewPage(preview).total, preview.references.length);
    }
    assert.deepEqual(
      calls,
      Array.from({ length: 4 }, () => ({ sheet: 'Data', range: 'A1:C3' }))
    );
    assert.deepEqual(cells, original);
  }
);
check(
  'Source hidden rows/columns and hidden-by-default overrides select visible blank cells, independent of rendering',
  async () => {
    const v = {
      name: 'Data',
      hiddenRows: [1],
      visibleRows: [0, 2],
      widths: [{ min: 2, max: 2, width: 16, hidden: true }],
      sheetFormat: { rowsHiddenByDefault: false },
    };
    assert.deepEqual((await collect({}, { selection: 'A1:C3', kind: 'visible', visibility: v })).references, [
      'A1',
      'C1',
      'A3',
      'C3',
    ]);
    assert.deepEqual(
      (
        await collect(
          {},
          { selection: 'A1:C3', kind: 'visible', visibility: { ...v, sheetFormat: { rowsHiddenByDefault: true } } }
        )
      ).references,
      ['A1', 'C1', 'A3', 'C3']
    );
    assert.deepEqual(
      (
        await collect(
          {},
          {
            selection: 'A1:C3',
            kind: 'visible',
            visibility: {
              ...v,
              rowHeights: [{ row: 2, height: 0 }],
              widths: [...v.widths, { min: 3, max: 3, width: 0, hidden: false }],
            },
          }
        )
      ).references,
      ['A1']
    );
    await assert.rejects(collect({}, { kind: 'visible' }), /visibility.*unavailable/);
    await assert.rejects(
      collect({}, { kind: 'visible', visibility: { ...v, geometryIssues: ['opaque hidden axes'] } }),
      /visibility.*unavailable/
    );
  }
);
check('Rule membership includes blank cells and known opaque authored ranges without evaluating the rule', async () => {
  const conditional = {
    sheet: 'Data',
    revision: 'rules',
    canEdit: false,
    protectedReasons: ['Unsupported rule remains present'],
    protectedRanges: ['$A$1:$D$1'],
    evaluationReasons: ['No current formula result'],
    rules: [{ range: '$B$2:$C$4', kind: 'greaterThan', values: [99], fillColor: 'FFFF00', stopIfTrue: false }],
  };
  const validation = {
    ...conditional,
    rules: [
      {
        range: 'B2:C4',
        kind: 'list',
        values: ['Yes'],
        allowBlank: true,
        showDropdown: true,
        promptTitle: '',
        prompt: '',
        errorTitle: '',
        error: '',
      },
    ],
  };
  for (const [kind, definition] of [
    ['conditional-formats', conditional],
    ['data-validation', validation],
  ])
    assert.deepEqual(
      (await collect({}, { selection: 'A1:D4', kind, rules: { complete: true, definition } })).references,
      ['A1', 'B1', 'C1', 'D1', 'B2', 'C2', 'B3', 'C3', 'B4', 'C4']
    );
});
check('Opaque or missing rule ranges, partial rule inventory and foreign worksheet metadata refuse', async () => {
  const definition = { sheet: 'Data', revision: null, canEdit: false, protectedReasons: ['Opaque rule'], rules: [] };
  for (const receipt of [
    { complete: true, definition },
    { complete: false, definition },
    { complete: true, definition: { ...definition, protectedReasons: [], rules: [{ range: '', kind: 'list' }] } },
    { complete: true, definition: { ...definition, sheet: 'Other', protectedRanges: ['A1'] } },
  ])
    await assert.rejects(collect({}, { kind: 'data-validation', rules: receipt }), /inventory|opaque|source range/);
});
check('Cancellation and retirement are checked before and after asynchronous reads and work yields', async () => {
  const controller = new AbortController(),
    held = defer();
  let called = 0;
  const promise = collect(
    {},
    {
      signal: controller.signal,
      read: async (_sheet, range) => {
        called++;
        return { range, complete: true, cells: await held.promise };
      },
    }
  );
  controller.abort();
  held.resolve({ A1: literal(1) });
  await assert.rejects(promise, (e) => e.name === 'AbortError');
  assert.equal(called, 1);
  let current = true;
  const held2 = defer(),
    retired = collect(
      {},
      {
        isCurrent: () => current,
        read: async (_sheet, range) => ({ range, complete: true, cells: await held2.promise }),
      }
    );
  current = false;
  held2.resolve({});
  await assert.rejects(retired, /initiating workbook/);
  current = true;
  await assert.rejects(
    collect(
      {},
      {
        isCurrent: () => current,
        yieldWork: async () => {
          current = false;
        },
      }
    ),
    /initiating workbook/
  );
  const early = new AbortController();
  early.abort();
  await assert.rejects(
    collect(
      {},
      {
        signal: early.signal,
        read: async () => {
          throw Error('not called');
        },
      }
    ),
    (e) => e.name === 'AbortError'
  );
});
check('Mutable producer options/read callbacks cannot retarget a held original collector', async () => {
  const calls = [],
    held = defer();
  let redirected = 0;
  const input = {
    binding: { ...binding, selection: 'A1:A1004' },
    options: options(),
    read: async (sheet, range) => {
      calls.push(range);
      return { range, complete: true, cells: range === 'A1:A1000' ? await held.promise : { A1004: literal(4) } };
    },
    isCurrent: () => true,
  };
  input.geometry = { name: 'Data', merges: [] };
  const promise = api.collectSpecialSelectionPreview(input);
  input.read = async () => {
    redirected++;
    throw Error('new owner');
  };
  input.options.types.number = false;
  input.binding.state = 'retarget';
  input.geometry.merges.push('A1:A1004');
  held.resolve({ A1: literal(1) });
  const p = await promise;
  assert.deepEqual(p.references, ['A1', 'A1004']);
  assert.equal(p.binding.state, 'original-head');
  assert.equal(p.options.types.number, true);
  assert.equal(redirected, 0);
  assert.deepEqual(calls, ['A1:A1000', 'A1001:A1004']);
});
check('No-match plans and paged previews are complete; forged order/scope/receipt plans are rejected', async () => {
  const empty = await collect({}, { selection: 'A1:A2' });
  assert.deepEqual(empty.references, []);
  assert.equal(empty.activeCell, null);
  assert.equal(api.specialSelectionPreviewPage(empty).total, 0);
  const forgedEmpty = structuredClone(empty);
  forgedEmpty.activeCell = 'A1';
  assert.throws(() => api.specialSelectionPreviewPage(forgedEmpty));
  const cells = Object.fromEntries(Array.from({ length: 1004 }, (_, i) => ['A' + (i + 1), literal(i)])),
    p = await collect(cells, { selection: 'A1:A1004' }),
    all = [];
  let offset = 0;
  do {
    const page = api.specialSelectionPreviewPage(p, offset);
    all.push(...page.rows.map((r) => r.cell));
    if (page.nextOffset === null) break;
    offset = page.nextOffset;
  } while (true);
  assert.equal(all.length, 1004);
  assert.deepEqual(all.slice(-4), ['A1001', 'A1002', 'A1003', 'A1004']);
  for (const change of [
    (q) => (q.references[0] = 'B1'),
    (q) => q.rows.reverse(),
    (q) => (q.rows[0].cell = 'B1'),
    (q) => q.readRanges.pop(),
    (q) => (q.positions = 10001),
    (q) => (q.sheet = 'Other'),
  ]) {
    const bad = structuredClone(p);
    change(bad);
    assert.throws(() => api.specialSelectionPreviewPage(bad));
  }
  await assert.rejects(collect({}, { kind: 'notes' }), /not available/);
});
