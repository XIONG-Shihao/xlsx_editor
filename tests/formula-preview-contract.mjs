import test from 'node:test';
import assert from 'node:assert/strict';
import {
  formulaPreviewFields,
  formulaPreviewTarget,
  formulaPreviewInner,
  bindFormulaPreview,
  validateFormulaPreview,
  formulaPreviewCancellation,
} from '../core/formula-preview.mjs';
import { encodeOwnerResult, decodeOwnerResult } from '../dist/core/transport.js';

const requestId = '018f1036-77cc-7aba-9561-125110477001';
const receipt = {
  documentId: 'trusted-document',
  ownerId: 'native-session',
  state: 'base',
  revision: 'source-revision',
};
const target = { requestId, sheet: "O'Brien Data", cell: 'XFD1048576', formula: 'SUM(B1:B3)' };
const request = { operation: 'previewFormula', documentId: receipt.documentId, state: receipt.state, ...target };
const inner = (result) => ({ schema: 1, ...target, authority: 'draft-preview', result });

test('preview requests cannot introduce write, path, metadata or calculation authority', () => {
  assert.equal(formulaPreviewFields(request), request);
  for (const [key, value] of Object.entries({
    commands: [],
    path: '/tmp/elsewhere',
    cells: [],
    metadata: {},
    nativeState: 'other',
    revision: 'fake',
    calculation: {},
    receipt,
  })) {
    assert.throws(() => formulaPreviewFields({ ...request, [key]: value }), /fields/);
  }
  assert.throws(() => formulaPreviewFields({ ...request, requestId: requestId.toUpperCase() }), /request ID/);
  assert.throws(
    () => formulaPreviewFields({ ...request, requestId: '018f1036-77cc-1aba-9561-125110477001' }),
    /request ID/
  );
  const transport = { ...request, kind: 'indexed', sessionId: receipt.ownerId };
  assert.equal(formulaPreviewFields(transport, { transport: true }), transport);
  assert.throws(() => formulaPreviewFields({ ...transport, kind: 'passive' }, { transport: true }), /operation/);
  const cancel = { operation: 'cancelFormulaPreview', documentId: receipt.documentId, state: receipt.state, requestId };
  formulaPreviewFields(cancel, { cancel: true });
  assert.throws(() => formulaPreviewFields({ ...cancel, formula: '1' }, { cancel: true }), /fields/);
});

test('targets obey canonical cell and UTF-8 budgets before evaluation', () => {
  formulaPreviewTarget(target.sheet, target.cell, '"' + '界'.repeat(2730) + '"'); // 8192 bytes exactly.
  assert.throws(() => formulaPreviewTarget(target.sheet, target.cell, '"' + '界'.repeat(2731) + '"'), /draft/);
  for (const cell of ['XFE1', 'A1048577', 'a1', 'A01', '$A$1', 'A1:B2', 'A0'])
    assert.throws(() => formulaPreviewTarget('Sheet', cell, '1'), /canonical/);
  for (const formula of ['=1', ' ', '1\0']) assert.throws(() => formulaPreviewTarget('Sheet', 'A1', formula), /draft/);
  assert.throws(() => formulaPreviewTarget('x'.repeat(32), 'A1', '1'), /worksheet/);
});

test('typed scalars remain distinct and incomplete or accepted results refuse', () => {
  for (const result of [
    { status: 'value', type: 'number', value: 12 },
    { status: 'value', type: 'text', value: '#N/A' },
    { status: 'value', type: 'error', value: '#N/A' },
    { status: 'value', type: 'boolean', value: false },
    ...['invalid', 'unavailable', 'cancelled'].map((status) => ({ status, reason: 'Bounded draft outcome.' })),
  ])
    formulaPreviewInner(inner(result), target);
  for (const result of [
    { status: 'value', type: 'number', value: Infinity },
    { status: 'value', type: 'number', value: NaN },
    { status: 'value', type: 'boolean', value: 1 },
    { status: 'value', type: 'error', value: 'unknown' },
    { status: 'value', type: 'text', value: '界'.repeat(10923) },
    { status: 'value', type: 'number', value: 1, partial: true },
    { status: 'unavailable', reason: '' },
  ])
    assert.throws(() => formulaPreviewInner(inner(result), target));
  assert.throws(
    () => formulaPreviewInner({ ...inner({ status: 'value', type: 'number', value: 12 }), authority: 'accepted' }),
    /authority/
  );
  assert.throws(
    () =>
      formulaPreviewInner({
        ...inner({ status: 'value', type: 'number', value: 12 }),
        calculation: { state: 'fresh' },
      }),
    /fields/
  );
});

test('owner, head, revision and exact draft receipts cannot be retargeted', () => {
  const bound = bindFormulaPreview(inner({ status: 'value', type: 'number', value: 12 }), receipt, target);
  validateFormulaPreview(bound, { ...target, receipt });
  for (const key of Object.keys(receipt))
    assert.throws(
      () => validateFormulaPreview(bound, { ...target, receipt: { ...receipt, [key]: 'different' } }),
      /retired/
    );
  for (const key of ['requestId', 'sheet', 'cell', 'formula'])
    assert.throws(() => validateFormulaPreview(bound, { ...target, [key]: 'different', receipt }), {
      code: 'formula-stale',
    });
  assert.throws(() => validateFormulaPreview({ ...bound, view: {} }), /fields/);
  assert.throws(() => validateFormulaPreview({ ...bound, receipt: { ...receipt, nativeVersion: 1 } }), /fields/);
  formulaPreviewCancellation({ schema: 1, requestId, cancelled: false }, requestId);
  assert.throws(
    () => formulaPreviewCancellation({ schema: 1, requestId, cancelled: true, ownerClosed: true }, requestId),
    /fields/
  );
  assert.throws(
    () =>
      formulaPreviewCancellation(
        { schema: 1, requestId: '018f1036-77cc-7aba-9561-125110477002', cancelled: true },
        requestId
      ),
    { code: /^(internal-invalid-field|operation-cancelled)$/ }
  );
});

test('the actual owner codec preserves signed zero through JSON without confusing text and errors', () => {
  for (const result of [
    { status: 'value', type: 'number', value: -0 },
    { status: 'value', type: 'number', value: 0 },
    { status: 'value', type: 'text', value: '#N/A' },
    { status: 'value', type: 'error', value: '#N/A' },
    { status: 'value', type: 'boolean', value: true },
  ]) {
    const bound = bindFormulaPreview(inner(result), receipt, target);
    const serialized = JSON.stringify(encodeOwnerResult({ formulaPreview: bound }));
    const restored = decodeOwnerResult(JSON.parse(serialized)).formulaPreview;
    validateFormulaPreview(restored, { ...target, receipt });
    assert.equal(restored.result.type, result.type);
    assert.ok(Object.is(restored.result.value, result.value));
  }
});
