/** Validate the owner's bounded, self-contained print document for native hosts. */
import { refusal } from '../core/refusal-codes.mjs';

const MAX_BYTES = 16 * 1024 * 1024, MAX_PAGES = 100;

export function validatePdfDocument(value) {
  if (
    !value ||
    Object.keys(value).some((key) => !['html', 'pageWidthPt', 'pageHeightPt', 'pageCount'].includes(key)) ||
    typeof value.html !== 'string' ||
    Buffer.byteLength(value.html) > MAX_BYTES ||
    !/^<!doctype html>/i.test(value.html) ||
    !/<head>/i.test(value.html)
  )
    throw refusal('internal-invalid-field', 'Invalid bounded workbook print document.');
  for (const field of ['pageWidthPt', 'pageHeightPt'])
    if (!Number.isFinite(value[field]) || value[field] < 100 || value[field] > 1500)
      throw refusal('internal-invalid-field', 'Invalid workbook paper dimensions.');
  if (!Number.isInteger(value.pageCount) || value.pageCount < 1 || value.pageCount > MAX_PAGES)
    throw refusal('print-limit', 'Workbook print exceeds 100 pages.');
  return { ...value };
}
