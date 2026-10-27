/** Envelope/consumer authority checks. No native or Excel geometry qualification. */
import test from 'node:test';
import assert from 'node:assert/strict';
import { usedRangeFields, normalizeNativeUsedRange, bindUsedRange, validateUsedRange } from '../core/used-range.mjs';

const extent = { minRow: 3, minColumn: 3, maxRow: 99, maxColumn: 25 };
const native = { version: 7, sourceRevision: 'a'.repeat(64), sheetPath: 'xl/worksheets/sheet1.xml' };
const receipt = {
  documentId: 'trusted-document',
  ownerId: 'native-owner',
  state: 'accepted-head',
  revision: 'source:draft:owner:3',
};
const provenance = {
  factsSchema: 'workbook.xlsx.used-range-facts.v1',
  declaration: { status: 'absent', ref: null, bounds: null },
  sourceCells: { minRow: 3, minColumn: 3, maxRow: 5, maxColumn: 5 },
  sourceMerges: null,
  effectiveCells: extent,
  sourceAxisFormatting: null,
  acceptedHistory: extent,
  sourceDefaultFormatting: false,
};
const raw = () => ({
  schema: 'workbook.xlsx.used-range.v1',
  policy: 'xlsx-used-range-v1',
  status: 'complete',
  sheet: 'Data',
  ...native,
  scopeMinMax: extent,
  range: 'D4:Z100',
  lastCell: 'Z100',
  empty: false,
  provenance,
  reasons: [],
});
const expected = { sheet: 'Data', native };

test('UsedRange query cannot introduce worksheet/cell mutation or storage authority', () => {
  const query = { operation: 'usedRange', documentId: receipt.documentId, state: receipt.state, sheet: 'Data' };
  usedRangeFields(query);
  usedRangeFields({ ...query, kind: 'indexed', sessionId: receipt.ownerId }, { transport: true });
  for (const field of ['range', 'cell', 'path', 'version', 'revision', 'ownerId', 'commands', 'metadata', 'receipt'])
    assert.throws(() => usedRangeFields({ ...query, [field]: 'injected' }), /fields/);
  assert.throws(() => usedRangeFields({ ...query, kind: 'indexed' }), /fields/);
  assert.throws(() => usedRangeFields({ ...query, operation: 'read' }), /operation/);
  for (const sheet of ['', 'a'.repeat(32), 'Data\0'])
    assert.throws(() => usedRangeFields({ ...query, sheet }), /worksheet/);
});

test('native selected-version identity and public owner receipts remain separate and exact', () => {
  const inner = normalizeNativeUsedRange(raw(), expected),
    bound = bindUsedRange(inner, receipt, { sheet: 'Data' });
  assert.deepEqual(bound.native, native);
  assert.deepEqual(bound.receipt, receipt);
  validateUsedRange(bound, { ...expected, receipt });
  assert.equal(bound.range, 'D4:Z100');
  assert.equal(bound.lastCell, 'Z100', 'the bounding corner need not be a stored cell');
  for (const field of Object.keys(native))
    assert.throws(() => normalizeNativeUsedRange({ ...raw(), [field]: field === 'version' ? 8 : 'other' }, expected));
  assert.throws(() => normalizeNativeUsedRange({ ...raw(), sheet: 'Other' }, expected), /changed/);
  for (const field of Object.keys(receipt))
    assert.throws(() => validateUsedRange(bound, { receipt: { ...receipt, [field]: 'other' } }), /retired/);
  assert.throws(
    () =>
      validateUsedRange(bound, {
        receipt: { ...receipt, ownerId: 'registry-alias', state: 'registry-alias:accepted-head' },
      }),
    /retired/
  );
  assert.throws(() => normalizeNativeUsedRange({ ...raw(), receipt }, expected), /fields/);
  assert.throws(() => validateUsedRange({ ...bound, calculation: {} }), /fields/);
});

test('v2 facts keep the closed owner and geometry contract; legacy facts remain readable', () => {
  for (const factsSchema of ['workbook.xlsx.used-range-facts.v1', 'workbook.xlsx.used-range-facts.v2']) {
    const source = { ...raw(), provenance: { ...provenance, factsSchema } },
      inner = normalizeNativeUsedRange(source, expected),
      bound = bindUsedRange(inner, receipt);
    validateUsedRange(bound, { ...expected, receipt });
    assert.equal(bound.provenance.factsSchema, factsSchema);
    assert.throws(() => validateUsedRange(bound, { receipt: { ...receipt, state: 'retired' } }), /retired/);
    assert.throws(() => normalizeNativeUsedRange({ ...source, lastCell: 'A1' }, expected), /geometry/);
    assert.throws(
      () =>
        normalizeNativeUsedRange({ ...source, provenance: { ...source.provenance, callerBounds: extent } }, expected),
      /fields/
    );
  }
  for (const factsSchema of ['workbook.xlsx.used-range-facts.v3', 'v2', 2, null])
    assert.throws(
      () => normalizeNativeUsedRange({ ...raw(), provenance: { ...provenance, factsSchema } }, expected),
      /source facts/
    );
  const physical = { minRow: 0, minColumn: 0, maxRow: 79, maxColumn: 25 },
    small = { minRow: 0, minColumn: 0, maxRow: 5, maxColumn: 2 };
  const truncated = {
    ...raw(),
    scopeMinMax: small,
    range: 'A1:C6',
    lastCell: 'C6',
    provenance: { ...provenance, sourceCells: physical, effectiveCells: small, acceptedHistory: null },
  };
  normalizeNativeUsedRange(truncated, expected);
  assert.throws(
    () =>
      normalizeNativeUsedRange(
        { ...truncated, provenance: { ...truncated.provenance, factsSchema: 'workbook.xlsx.used-range-facts.v2' } },
        expected
      ),
    { code: 'layout-limit' }
  );
});

test('incomplete facts carry reasons and never authorize partial or padded geometry', () => {
  const incomplete = {
    ...raw(),
    status: 'incomplete',
    scopeMinMax: null,
    range: null,
    lastCell: null,
    empty: null,
    reasons: [{ code: 'row-formatting-unqualified', message: 'Row formatting needs qualification.' }],
    provenance: { ...provenance, sourceAxisFormatting: { minRow: 80, minColumn: 0, maxRow: 80, maxColumn: 16383 } },
  };
  const reply = bindUsedRange(normalizeNativeUsedRange(incomplete, expected), receipt);
  assert.equal(reply.status, 'incomplete');
  assert.equal(reply.range, null);
  for (const [field, value] of Object.entries({ scopeMinMax: extent, range: 'A1:H20', lastCell: 'H20', empty: false }))
    assert.throws(() => normalizeNativeUsedRange({ ...incomplete, [field]: value }, expected), {
      code: /^(layout-unsupported|note-unsupported)$/,
    });
  assert.throws(() => normalizeNativeUsedRange({ ...incomplete, reasons: [] }, expected), {
    code: /^(layout-unsupported|note-unsupported)$/,
  });
  assert.throws(() => normalizeNativeUsedRange({ ...raw(), reasons: incomplete.reasons }, expected), /complete/);
});

test('scope and last-cell receipts reject out-of-sheet, incoherent and oversized data', () => {
  for (const patch of [
    { minRow: -1 },
    { minColumn: -0 },
    { maxRow: 1048576 },
    { maxColumn: 16384 },
    { minRow: 100 },
    { maxRow: 1.5 },
  ])
    assert.throws(
      () => normalizeNativeUsedRange({ ...raw(), scopeMinMax: { ...extent, ...patch } }, expected),
      /bounds/
    );
  for (const range of ['D4:Z1000', 'Z100:D4', 'D4:Z100 D200', 'd4:Z100'])
    assert.throws(() => normalizeNativeUsedRange({ ...raw(), range }, expected), /geometry/);
  for (const lastCell of ['A100', 'Z1', 'Z1000'])
    assert.throws(() => normalizeNativeUsedRange({ ...raw(), lastCell }, expected), /geometry/);
  for (const patch of [
    { schema: 'other' },
    { policy: 'xlsx-view-extent' },
    { status: 'approximate' },
    { partial: true },
  ])
    assert.throws(() => normalizeNativeUsedRange({ ...raw(), ...patch }, expected));
  assert.throws(
    () => normalizeNativeUsedRange({ ...raw(), provenance: { ...provenance, cells: [] } }, expected),
    /fields/
  );
  assert.throws(
    () =>
      normalizeNativeUsedRange(
        { ...raw(), provenance: { ...provenance, sourceCells: { ...extent, maxColumn: 16384 } } },
        expected
      ),
    /bounds/
  );
  assert.throws(
    () =>
      normalizeNativeUsedRange(
        { ...raw(), provenance: { ...provenance, declaration: { status: 'valid', ref: 'A1', bounds: null } } },
        expected
      ),
    /provenance/
  );
  assert.throws(
    () =>
      normalizeNativeUsedRange(
        { ...raw(), provenance: { ...provenance, declaration: { status: 'valid', ref: 'A1', bounds: extent } } },
        expected
      ),
    { code: /^(internal-mismatch|layout-limit)$/ }
  );
  assert.throws(
    () =>
      normalizeNativeUsedRange(
        { ...raw(), scopeMinMax: { ...extent, maxColumn: 24 }, range: 'D4:Y100', lastCell: 'Y100' },
        expected
      ),
    { code: 'layout-limit' }
  );
  const edge = { minRow: 1048575, minColumn: 16383, maxRow: 1048575, maxColumn: 16383 };
  assert.equal(
    normalizeNativeUsedRange(
      {
        ...raw(),
        scopeMinMax: edge,
        range: 'XFD1048576',
        lastCell: 'XFD1048576',
        provenance: { ...provenance, effectiveCells: edge, acceptedHistory: edge },
      },
      expected
    ).lastCell,
    'XFD1048576'
  );
});

test('a complete empty-sheet receipt is distinct from unavailable facts and nonempty history', () => {
  const empty = {
    ...raw(),
    scopeMinMax: { minRow: 0, minColumn: 0, maxRow: 0, maxColumn: 0 },
    range: 'A1',
    lastCell: 'A1',
    empty: true,
    provenance: { ...provenance, sourceCells: null, effectiveCells: null, acceptedHistory: null },
  };
  assert.equal(normalizeNativeUsedRange(empty, expected).empty, true);
  const v2 = { ...empty, provenance: { ...empty.provenance, factsSchema: 'workbook.xlsx.used-range-facts.v2' } };
  assert.equal(normalizeNativeUsedRange(v2, expected).empty, true);
  const blankA1 = { ...v2, provenance: { ...v2.provenance, sourceCells: v2.scopeMinMax } };
  assert.throws(() => normalizeNativeUsedRange(blankA1, expected), /empty/);
  assert.equal(normalizeNativeUsedRange({ ...blankA1, empty: false }, expected).empty, false);
  assert.equal(
    normalizeNativeUsedRange(
      { ...blankA1, provenance: { ...blankA1.provenance, factsSchema: 'workbook.xlsx.used-range-facts.v1' } },
      expected
    ).empty,
    true,
    'Legacy effective-empty semantics stay readable'
  );
  for (const key of ['effectiveCells', 'sourceMerges', 'sourceAxisFormatting', 'acceptedHistory'])
    assert.throws(
      () => normalizeNativeUsedRange({ ...empty, provenance: { ...empty.provenance, [key]: extent } }, expected),
      /empty|excludes/
    );
  assert.throws(() => normalizeNativeUsedRange({ ...empty, range: 'A1:A1' }, expected), /geometry/);
});
