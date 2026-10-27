// Questions the engine answers about a workbook without changing it (the list is
// the engine's own: native/read-only-operations.json). Test doubles that
// simulate edits send these to the real engine, so a new question needs no
// change to them; `workbook-view` depends only on the cell-free copy
// (metadata.xlsx) a double writes.
import { after } from 'node:test';
import { readFileSync } from 'node:fs';
import { nativeWorkbookOperationAsync, closeNativeWorkbookWorker } from '../../core/native-runtime.mjs';

after(() => closeNativeWorkbookWorker());

const questions = new Set(
  JSON.parse(readFileSync(new URL('../../native/read-only-operations.json', import.meta.url), 'utf8')).operations
);

/** The real answer when `input` is such a question and the double has none of
 * its own (`simulated`); undefined otherwise. */
export function engineQuestion(input, simulated = []) {
  return questions.has(input?.operation) && !simulated.includes(input.operation)
    ? nativeWorkbookOperationAsync(input)
    : undefined;
}
