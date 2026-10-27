import { indexedObjectReferenceGuard } from '../indexed-object-authoring.mjs';
import { indexedPrintReferenceGuard } from '../indexed-print.mjs';
import fs from 'node:fs/promises';
import path from 'node:path';
import { refusal } from '../refusal-codes.mjs';
const GUARDED_KINDS = [
  'renameSheet',
  'duplicateSheet',
  'insertRow',
  'deleteRow',
  'insertColumn',
  'deleteColumn',
  'sortRows',
  'removeDuplicates',
  'filterRows',
  'reapplyFilter',
  'clearFilter',
  'removeFilter',
];
/** Refuse a sheet, row, column, sort or filter change that would break a chart, image or print area. */
export async function referenceGuard(ctx, transform) {
  const { current } = ctx;
  const kind = transform?.operation;
  if (!GUARDED_KINDS.includes(kind)) return;
  const metadataBytes = await fs.readFile(path.join(current.root, 'metadata.xlsx'));
  const referenceCommand =
    kind === 'removeDuplicates'
      ? { ...transform, operation: 'sortRows' }
      : ['reapplyFilter', 'clearFilter', 'removeFilter'].includes(kind)
        ? { ...transform, operation: 'filterRows' }
        : transform;
  const reason =
    indexedObjectReferenceGuard(metadataBytes, current.view, referenceCommand) ??
    indexedPrintReferenceGuard(metadataBytes, current.view, referenceCommand);
  if (reason) throw refusal('workbook-refused', reason);
}
