import assert from 'node:assert/strict';
import { SheetHistory } from '@office-editors/xlsx-editor/ui/sheetHistory';
const sheet = {
  name: 'Sheet1',
  path: '',
  state: 'visible',
  rows: 2,
  columns: 2,
  cells: { A1: { value: 1 }, B2: { value: 2 } },
  merges: [],
  widths: [],
};
const history = new SheetHistory([sheet]),
  other = new SheetHistory([sheet]);
history.edit({ sheet: 'Sheet1', cell: 'A1', value: 5 });
history.edit({ sheet: 'Sheet1', cell: 'B2', value: 9 });
history.markSaved();
history.undo();
assert.equal(history.current.sheets[0].cells.B2.value, 2);
assert(history.dirty);
history.redo();
assert.equal(history.current.sheets[0].cells.B2.value, 9);
assert(!history.dirty);
history.edit({ sheet: 'Sheet1', operation: 'deleteRow', index: 0 });
history.undo();
assert.equal(history.current.sheets[0].cells.A1.value, 5);
history.undo();
history.edit({ sheet: 'Sheet1', cell: 'A1', value: 7 });
assert(!history.canRedo);
assert.equal(other.current.sheets[0].cells.A1.value, 1);

const longSession = new SheetHistory([sheet]);
for (let index = 0; index < 2_000; index++) {
  longSession.edit({ sheet: 'Sheet1', cell: 'A1', value: index });
  if (index % 100 === 99) longSession.markSaved();
}
assert.equal(longSession.current.edits.length, 1);
assert.equal(longSession.current.edits[0].value, 1_999);
assert.ok(
  longSession.past.reduce((total, frame) => total + frame.edits.length, 0) <= 100,
  'bounded undo must not retain a cumulative repeated-cell journal'
);

const formulaHistory = new SheetHistory([sheet]);
const compactFormat = { type: 'formatRange', sheet: 'Sheet1', range: 'B200:B10000', format: { bold: true } };
const indexedHistory = new SheetHistory([sheet]);
indexedHistory.batch([compactFormat], [sheet]);
assert.deepEqual(indexedHistory.current.edits, [compactFormat]);
assert.deepEqual(indexedHistory.undo().focus, { sheet: 'Sheet1', cell: 'B200' });
assert.deepEqual(indexedHistory.redo().edits, [compactFormat]);
formulaHistory.edit({ sheet: 'Sheet1', cell: 'A1', formula: '=B2' });
formulaHistory.edit({ sheet: 'Sheet1', cell: 'A1', value: 4 });
assert.deepEqual(formulaHistory.current.edits, [{ sheet: 'Sheet1', cell: 'A1', value: 4 }]);
formulaHistory.edit({ sheet: 'Sheet1', cell: 'A1', formula: 'B2*2' });
assert.deepEqual(formulaHistory.current.edits, [{ sheet: 'Sheet1', cell: 'A1', formula: 'B2*2' }]);
formulaHistory.undo();
assert.deepEqual(formulaHistory.current.edits, [{ sheet: 'Sheet1', cell: 'A1', value: 4 }]);
formulaHistory.redo();
assert.deepEqual(formulaHistory.current.edits, [{ sheet: 'Sheet1', cell: 'A1', formula: 'B2*2' }]);
console.log(
  'PASS sheet Undo/Redo, saved mark, compact journal, structural restoration, redo branch and document isolation'
);

let rawNodes = 0;
for (let node = longSession.journal; node; node = node.previous) rawNodes++;
assert(rawNodes <= 101, `retained raw journal must fit the undo window: ${rawNodes}`);
for (let i = 0; i < 100; i++) longSession.undo();
assert.equal(longSession.current.sheets[0].cells.A1.value, 1899);
assert.equal(longSession.current.edits.at(-1).value, 1899);
for (let i = 0; i < 100; i++) longSession.redo();
assert.equal(longSession.current.sheets[0].cells.A1.value, 1999);
assert(!longSession.dirty);

// Prefix checkpoints must preserve row shifts and edits on either side of them.
const structural = new SheetHistory([sheet]);
for (let i = 0; i < 450; i++) {
  structural.edit(
    i === 120
      ? { sheet: 'Sheet1', operation: 'insertRow', index: 0 }
      : i === 240
        ? { sheet: 'Sheet1', operation: 'deleteRow', index: 0 }
        : { sheet: 'Sheet1', cell: 'A1', value: i }
  );
  if (i % 50 === 49) structural.markSaved();
}
const { applyGridEdit, plainSheets } = await import('@office-editors/xlsx-editor/ui/sheetModel');
function verifyReplay(h) {
  let replay = sheet;
  for (const edit of h.current.edits) replay = applyGridEdit(replay, edit);
  assert.deepEqual(plainSheets([replay]), plainSheets(h.current.sheets));
}
verifyReplay(structural);
for (let i = 0; i < 100; i++) {
  structural.undo();
  verifyReplay(structural);
}
for (let i = 0; i < 100; i++) {
  structural.redo();
  verifyReplay(structural);
}
structural.undo();
structural.edit({ sheet: 'Sheet1', cell: 'B2', value: 'branch' });
assert(!structural.canRedo);
verifyReplay(structural);
console.log(
  'PASS old journal nodes compact into replayable checkpoints without changing saved marks, row shifts or the 100-step undo/redo window'
);

const batchHistory = new SheetHistory([sheet]);
batchHistory.batch([
  { sheet: 'Sheet1', cell: 'A1', value: 10 },
  { sheet: 'Sheet1', cell: 'B2', value: 20 },
]);
assert.equal(batchHistory.current.sheets[0].cells.A1.value, 10);
assert.equal(batchHistory.current.sheets[0].cells.B2.value, 20);
batchHistory.undo();
assert.equal(batchHistory.current.sheets[0].cells.A1.value, 1);
assert.equal(batchHistory.current.sheets[0].cells.B2.value, 2);
assert(!batchHistory.canUndo);
batchHistory.redo();
assert.equal(batchHistory.current.sheets[0].cells.B2.value, 20);
const guarded = new SheetHistory([sheet]);
guarded.batch([{ sheet: 'Sheet1', cell: 'A1', expectedValue: 1, value: 3 }]);
guarded.batch([{ sheet: 'Sheet1', cell: 'A1', expectedValue: 3, value: 5 }]);
assert.equal(guarded.current.edits.length, 2, 'guarded replay must retain each precondition');
console.log('PASS one undo frame per batch and guarded journal replay');

// Native retention follows the actual undo window, including redo branches.
const nativeHistory = new SheetHistory([sheet]);
for (let i = 0; i < 105; i++)
  nativeHistory.batch([{ sheet: 'Sheet1', cell: 'A1', value: i }], [sheet], {
    sheets: [sheet],
    storage: { kind: 'rust-indexed-v1', sessionId: 'native', state: String(i + 1) },
  });
assert.equal(nativeHistory.retainedNativeStates('native').length, 101);
assert.equal(nativeHistory.retainedNativeStates('native').includes('1'), false);
nativeHistory.undo();
assert.equal(nativeHistory.retainedNativeStates('native').includes('105'), true);
nativeHistory.batch([{ sheet: 'Sheet1', cell: 'A1', value: 999 }], [sheet], {
  sheets: [sheet],
  storage: { kind: 'rust-indexed-v1', sessionId: 'native', state: 'branch' },
});
assert.equal(nativeHistory.retainedNativeStates('native').includes('105'), false);
assert.deepEqual(nativeHistory.retainedNativeStates('other'), []);

// Staged data stays as one immutable descriptor and cannot be flattened into a
// fabricated cell edit. It is an ordered barrier between surrounding writes.
const stagedHistory = new SheetHistory([sheet]);
const descriptor = {
  type: 'importCells',
  id: '00000000-0000-4000-8000-000000000001',
  sha256: 'a'.repeat(64),
  bytes: 90000,
  count: 100000,
  shape: { sheet: 'Sheet1', start: 'C4', rows: 10000, columns: 10 },
  policy: 'grid',
};
stagedHistory.edit({ sheet: 'Sheet1', cell: 'A1', value: 3 });
assert.throws(() => stagedHistory.batch([descriptor]), /accepted workbook view/);
stagedHistory.batch([descriptor], [sheet], {
  sheets: [sheet],
  storage: { kind: 'rust-indexed-v1', sessionId: 'staged', state: '1' },
});
stagedHistory.edit({ sheet: 'Sheet1', cell: 'A1', value: 7 });
assert.deepEqual(stagedHistory.current.edits, [
  { sheet: 'Sheet1', cell: 'A1', value: 3 },
  descriptor,
  { sheet: 'Sheet1', cell: 'A1', value: 7 },
]);
stagedHistory.undo();
assert.deepEqual(stagedHistory.current.edits.at(-1), descriptor);
assert.deepEqual(stagedHistory.undo().focus, { sheet: 'Sheet1', cell: 'C4' });
stagedHistory.redo();
assert.equal(stagedHistory.current.edits.at(-1), descriptor);
assert.equal('sheet' in descriptor, false);
for (let index = 0; index < 160; index++) stagedHistory.edit({ sheet: 'Sheet1', cell: 'A1', value: index });
assert.equal(stagedHistory.current.edits.filter((edit) => edit.type === 'importCells').length, 1);
assert.equal(
  stagedHistory.current.edits.find((edit) => edit.type === 'importCells'),
  descriptor
);
console.log('PASS staged paste keeps one exact replay descriptor, focus and ordering across Undo/Redo');
