import type { WorkbookHistoryEdit } from '../core/contracts';
/** Persist the edit journal against its separately owned immutable base. Workbook
 * projections include large caches and are reconstructed, never sent to Rust state. */
export function compactWorkbookDraft(
  revision: string,
  edits: WorkbookHistoryEdit[]
): { revision: string; journalOnly: true; edits: WorkbookHistoryEdit[] } {
  return { revision, journalOnly: true, edits };
}
