import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  NOTE_LIMITS,
  noteCommand,
  captureNoteCommands,
  noteInventory,
  notesFields,
  bindNotes,
  validateNotes,
} from '../core/notes.mjs';
const prior = '1'.repeat(64),
  set = (patch = {}) => ({
    type: 'setNote',
    sheet: 'Main',
    cell: 'A1',
    expectedNoteRevision: null,
    note: { author: 'Alice & Bob', text: '  Review <total>\r\n😀  ' },
    ...patch,
  }),
  empty = () => ({
    schema: 'workbook.xlsx.legacy-notes.v1',
    sheet: 'Main',
    noteRevision: null,
    notes: [],
    canEdit: true,
    protectedReason: null,
  }),
  receipt = { documentId: 'document', ownerId: 'owner', state: 'head', revision: 'revision' };
test('plain author/body and canonical cell commands preserve literal whitespace and Unicode within finite bounds', () => {
  const c = noteCommand(set());
  assert.equal(c.note.text, '  Review <total>\r\n😀  ');
  for (const cell of ['A1', 'XFD1048576']) assert.equal(noteCommand(set({ cell })).cell, cell);
  assert.equal(noteCommand(set({ note: { author: 'a'.repeat(128), text: 'x'.repeat(4096) } })).note.text.length, 4096);
  for (const cell of ['a1', 'A01', '$A$1', 'A1:B2', 'XFE1', 'A1048577'])
    assert.throws(() => noteCommand(set({ cell })));
  for (const note of [
    { author: '', text: 'body' },
    { author: 'a\nb', text: 'body' },
    { author: 'tc={00000000-0000-4000-8000-000000000000}', text: 'body' },
    { author: 'a'.repeat(129), text: 'body' },
    { author: 'a', text: '' },
    { author: 'a', text: '\u0000' },
    { author: 'a', text: '\ud800' },
    { author: 'a', text: 'x'.repeat(4097) },
  ])
    assert.throws(() => noteCommand(set({ note })));
  assert.throws(() => noteCommand({ ...set(), xml: 'caller' }));
  assert.throws(() => noteCommand(set({ expectedNoteRevision: 'unknown' })));
  assert.deepEqual(noteCommand({ type: 'removeNote', sheet: 'Main', cell: 'A1', expectedNoteRevision: prior }), {
    type: 'removeNote',
    sheet: 'Main',
    cell: 'A1',
    expectedNoteRevision: prior,
  });
});
test('capture detaches caller commands before queued writes and retains other commands without inventing a note grant', () => {
  const source = set(),
    other = { cell: 'B1', sheet: 'Main', value: 3 },
    result = captureNoteCommands([source, other], { sheets: [{ name: 'Main' }] });
  source.note.text = 'mutated';
  other.value = 9;
  assert.equal(result[0].note.text, '  Review <total>\r\n😀  ');
  assert.equal(result[1].value, 3);
  assert.throws(() => captureNoteCommands([set()], { sheets: [{ name: 'Other' }] }));
  assert.throws(() => captureNoteCommands([]));
  assert.throws(() =>
    captureNoteCommands(Array.from({ length: 600 }, () => set({ note: { author: 'Author', text: 'x'.repeat(4096) } })))
  );
});
test('inventory separates empty/editable models from protected/unknown graphs and requires complete sorted unique records', () => {
  assert.deepEqual(noteInventory(empty(), { sheet: 'Main' }).notes, []);
  const rows = [
      { cell: 'A1', author: 'Alice', text: 'First' },
      { cell: 'C9', author: 'Bob', text: 'Second' },
    ],
    valid = { ...empty(), noteRevision: prior, notes: rows };
  assert.equal(noteInventory(valid, { sheet: 'Main' }).notes.length, 2);
  const readonly = { ...valid, canEdit: false, protectedReason: 'Read-only workbook' };
  assert.equal(noteInventory(readonly, { sheet: 'Main' }).notes.length, 2);
  const protectedModel = { ...empty(), notes: null, canEdit: false, protectedReason: 'Threaded notes stay protected.' };
  assert.equal(noteInventory(protectedModel, { sheet: 'Main' }).notes, null);
  for (const patch of [
    { notes: rows },
    { noteRevision: prior },
    { noteRevision: prior, notes: [rows[1], rows[0]] },
    { noteRevision: prior, notes: [rows[0], rows[0]] },
    { canEdit: false, protectedReason: null },
    { notes: null },
    {
      noteRevision: prior,
      notes: Array.from({ length: NOTE_LIMITS.notes + 1 }, (_, i) => ({ ...rows[0], cell: 'A' + (i + 1) })),
    },
  ])
    assert.throws(() => noteInventory({ ...empty(), ...patch }, { sheet: 'Main' }));
  assert.throws(() => noteInventory(valid, { sheet: 'Other' }));
  assert.throws(() => noteInventory({ ...valid, callerAuthority: true }, { sheet: 'Main' }));
});
test('transport fields and bound replies keep the original document, owner, head, revision and sheet', () => {
  const request = { operation: 'notes', documentId: 'document', state: 'head', sheet: 'Main' };
  notesFields(request);
  notesFields({ ...request, kind: 'indexed', sessionId: 'owner' }, { transport: true });
  for (const patch of [
    { cell: 'A1' },
    { state: '' },
    { operation: 'noteMembership' },
    { kind: 'indexed', sessionId: 'owner' },
  ])
    assert.throws(() => notesFields({ ...request, ...patch }));
  const raw = empty(),
    bound = bindNotes(raw, receipt, { sheet: 'Main' });
  assert.equal(validateNotes(bound, receipt, { sheet: 'Main' }).ownerId, 'owner');
  raw.notes.push({ cell: 'A1', author: 'Caller', text: 'Changed' });
  assert.equal(bound.notes.length, 0);
  for (const key of Object.keys(receipt))
    assert.throws(() => validateNotes(bound, { ...receipt, [key]: 'retired' }, { sheet: 'Main' }));
  assert.throws(() => validateNotes({ ...bound, notesPath: '/external' }, receipt, { sheet: 'Main' }));
});
