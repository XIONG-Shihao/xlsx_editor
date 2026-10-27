import { FORMAT, matchesFormat } from './format.mjs';
/** Owned plain legacy-note commands and captured read DTOs; no author identity authentication. */
import { noteMembershipCell, noteMembershipSheet } from './note-membership.mjs';
import { refusal } from './refusal-codes.mjs';
export const NOTE_LIMITS = Object.freeze({ notes: 256, author: 128, text: 4096, batchBytes: 2 * 1024 * 1024 });
const fields = (v, keys, label) => {
  if (!v || typeof v !== 'object' || Array.isArray(v) || Object.keys(v).sort().join(',') !== [...keys].sort().join(','))
    throw refusal('internal-invalid-field', 'Invalid ' + label + ' fields.');
};
const text = (v, max, label) => {
  if (
    typeof v !== 'string' ||
    v.length > max ||
    !v.trim() ||
    (typeof v.isWellFormed === 'function' && !v.isWellFormed()) ||
    /[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/.test(v)
  )
    throw refusal('internal-invalid-field', 'Invalid ' + label + '.');
  return v;
};
const hash = (v) => typeof v === 'string' && /^[a-f0-9]{64}$/.test(v);
export const isNoteCommand = (v) => ['setNote', 'removeNote'].includes(v?.type);
export function noteCommand(v) {
  const remove = v?.type === 'removeNote';
  fields(v, ['type', 'sheet', 'cell', 'expectedNoteRevision', ...(!remove ? ['note'] : [])], 'note command');
  if (!isNoteCommand(v)) throw refusal('note-unsupported', 'Unsupported note command.');
  noteMembershipSheet(v.sheet);
  noteMembershipCell(v.cell);
  if (v.expectedNoteRevision !== null && !hash(v.expectedNoteRevision))
    throw refusal('internal-invalid-field', 'Invalid prior note revision.');
  let note;
  if (!remove) {
    fields(v.note, ['author', 'text'], 'note');
    note = {
      author: text(v.note.author, NOTE_LIMITS.author, 'note author'),
      text: text(v.note.text, NOTE_LIMITS.text, 'note text'),
    };
    if (/tc=/i.test(note.author)) throw refusal('note-unsupported', 'Threaded-comment author tokens are reserved.');
    if (/[\r\n\t]/.test(note.author)) throw refusal('note-bad-input', 'Use a single-line author label.');
  }
  return {
    type: v.type,
    sheet: v.sheet,
    cell: v.cell,
    expectedNoteRevision: v.expectedNoteRevision,
    ...(!remove ? { note } : {}),
  };
}
export function captureNoteCommands(commands, view) {
  if (!Array.isArray(commands) || !commands.length) throw refusal('internal-missing-field', 'Missing note commands.');
  const result = commands.map((v) => {
    if (!isNoteCommand(v)) return structuredClone(v);
    const c = noteCommand(v);
    if (view && !view.sheets.some((s) => s.name === c.sheet))
      throw refusal('internal-unknown-reference', 'Unknown note worksheet.');
    return c;
  });
  if (new TextEncoder().encode(JSON.stringify(result)).byteLength > NOTE_LIMITS.batchBytes)
    throw refusal('note-limit', 'Note batch exceeds its 2 MiB bound.');
  return result;
}
export function notesFields(request, { transport = false } = {}) {
  fields(
    request,
    ['operation', 'documentId', 'state', 'sheet', ...(transport ? ['kind', 'sessionId'] : [])],
    'notes request'
  );
  if (request.operation !== 'notes' || (transport && request.kind !== 'indexed'))
    throw refusal('internal-invalid-field', 'Invalid notes operation.');
  for (const k of ['documentId', 'state', ...(transport ? ['sessionId'] : [])]) text(request[k], 128, 'notes ' + k);
  noteMembershipSheet(request.sheet);
  return request;
}
const keys = ['schema', 'sheet', 'noteRevision', 'notes', 'canEdit', 'protectedReason'];
export function noteInventory(value, { sheet }) {
  fields(value, keys, 'note inventory');
  if (!matchesFormat(value.schema, 'SCHEMA_LEGACY_NOTES_V1') || value.sheet !== sheet)
    throw refusal('note-stale', 'Note inventory changed the captured worksheet.');
  if (typeof value.canEdit !== 'boolean' || (value.noteRevision !== null && !hash(value.noteRevision)))
    throw refusal('internal-invalid-field', 'Invalid note inventory authority.');
  if (value.canEdit) {
    if (value.protectedReason !== null)
      throw refusal('note-protected', 'Editable notes cannot have a protected reason.');
  } else {
    text(value.protectedReason, 1024, 'note refusal');
    if (value.notes === null) {
      if (value.noteRevision !== null)
        throw refusal('internal-unknown-reference', 'Unknown protected notes cannot grant a revision.');
      return value;
    }
  }
  if (
    !Array.isArray(value.notes) ||
    value.notes.length > NOTE_LIMITS.notes ||
    !!value.notes.length !== (value.noteRevision !== null)
  )
    throw refusal('internal-invalid-field', 'Invalid complete note inventory.');
  let prior = -1;
  for (const n of value.notes) {
    fields(n, ['cell', 'author', 'text'], 'note entry');
    const pos = noteMembershipCell(n.cell);
    if (pos <= prior) throw refusal('note-bad-input', 'Notes must have sorted unique cell targets.');
    prior = pos;
    noteCommand({
      type: 'setNote',
      sheet,
      cell: n.cell,
      expectedNoteRevision: value.noteRevision,
      note: { author: n.author, text: n.text },
    });
  }
  return value;
}
function receipt(value) {
  fields(value, ['documentId', 'ownerId', 'state', 'revision'], 'note receipt');
  for (const k of Object.keys(value)) text(value[k], k === 'revision' ? 512 : 128, 'note ' + k);
  return value;
}
export function bindNotes(value, binding, expected) {
  receipt(binding);
  noteInventory(value, expected);
  return { ...structuredClone(value), ...binding };
}
export function validateNotes(value, binding, expected) {
  receipt(binding);
  fields(value, [...keys, 'documentId', 'ownerId', 'state', 'revision'], 'bound notes');
  for (const k of Object.keys(binding))
    if (value[k] !== binding[k]) throw refusal('note-stale', 'Notes belong to a retired owner/head.');
  noteInventory(Object.fromEntries(keys.map((k) => [k, value[k]])), expected);
  return value;
}
