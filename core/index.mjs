/** Public headless workbook engine. No Electron, app account or UI dependency. */
import { refusal } from './refusal-codes.mjs';

export {
  MAX_WORKBOOK_BYTES,
  readWorkbook,
  editWorkbook,
  editWorkbookResult,
  createWorkbook,
  columnName,
  cellPosition,
} from './workbook-package.mjs';
export async function readCalculatedWorkbook(bytes, options) {
  const { readCalculatedWorkbook: read } = await import('./workbook-service.mjs');
  return read(bytes, options);
}

export async function editCalculatedWorkbookResult(bytes, edits, expectedRevision) {
  const { WorkbookSession } = await import('./workbook-service.mjs');
  const session = new WorkbookSession(bytes);
  if (!edits.length) {
    const view = await session.read();
    if (expectedRevision && expectedRevision !== view.revision)
      throw refusal('workbook-stale', 'Conflict: this workbook changed. Reload or save a copy.');
    return { bytes, view };
  }
  return session.apply({
    operationId: 'desktop-command-batch',
    expectedRevision,
    commands: edits,
    ...(!expectedRevision ? { expectedRevision: (await session.read()).revision } : {}),
  });
}
