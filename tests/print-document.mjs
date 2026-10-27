import test from 'node:test';
import assert from 'node:assert/strict';
import { validatePdfDocument } from '../adapters/print-document.mjs';

const document = () => ({ html: '<!doctype html><html><head></head><body></body></html>',
  pageWidthPt: 595, pageHeightPt: 842, pageCount: 1 });

test('native print validation preserves the accepted page document and rejects extra dispatch fields', () => {
  const accepted = document();
  assert.deepEqual(validatePdfDocument(accepted), accepted);
  assert.throws(() => validatePdfDocument({ ...accepted, destination: '/private/output.pdf' }), /Invalid bounded/);
  assert.throws(() => validatePdfDocument({ ...accepted, html: '<html><head></head></html>' }), /Invalid bounded/);
});

test('native print validation keeps page, geometry and byte limits before host dispatch', () => {
  for (const pageCount of [0, 101, 1.5])
    assert.throws(() => validatePdfDocument({ ...document(), pageCount }), /100 pages/);
  for (const pageWidthPt of [NaN, 99, 1501])
    assert.throws(() => validatePdfDocument({ ...document(), pageWidthPt }), /paper dimensions/);
  assert.throws(() => validatePdfDocument({ ...document(), html: document().html + 'x'.repeat(16 * 1024 * 1024) }),
    /Invalid bounded/);
});
