import { FORMAT, matchesFormat } from '../format.mjs';
/** Add, change or remove a cell note. */
import { captureNoteCommands, noteInventory } from '../notes.mjs';
import fs from 'node:fs/promises';
import path from 'node:path';
import { FACTS_V2, uuidV4, matchesHandles } from './shared.mjs';
import { refusal } from '../refusal-codes.mjs';
export const name = 'notes';
export const handles = { type: ['setNote', 'removeNote'] };
export const matches = (commands) => matchesHandles(handles, commands);
export async function run(ctx, commands) {
  const { current } = ctx;
  const [command] = captureNoteCommands(commands, current.view);
  if (!matchesFormat(current.index.sourceFactSchema, 'SCHEMA_USED_RANGE_FACTS_V2'))
    throw refusal('session-unsupported', 'Note authoring requires qualified native source ownership.');
  const source = await ctx.rewriteSource(true),
    candidate = source.candidate;
  try {
    const result = await ctx.native({
      action: 'workbook',
      operation: 'package-edit',
      root: current.root,
      version: current.version,
      source: path.basename(source.path),
      ...ctx.rewriteFields(source),
      command,
    });
    ctx.rewriteCheck(candidate);
    const fields = 'changed,contentChanged,file,priorNoteRevision,remainingNotes,rewriteId,schema';
    if (
      !result ||
      typeof result !== 'object' ||
      Array.isArray(result) ||
      Object.keys(result).sort().join(',') !== fields ||
      !matchesFormat(result.schema, 'SCHEMA_NOTE_WRITER_V1') ||
      result.file !== 'package.xlsx' ||
      !uuidV4.test(result.rewriteId) ||
      result.contentChanged !== false ||
      typeof result.changed !== 'boolean' ||
      result.priorNoteRevision !== command.expectedNoteRevision ||
      !Number.isSafeInteger(result.remainingNotes) ||
      result.remainingNotes < 0 ||
      result.remainingNotes > 256
    )
      throw refusal('internal-invalid-field', 'Invalid native note writer receipt.');
    if (!result.changed) {
      await fs.rm(candidate.root, { recursive: true, force: true });
      ctx.dropRoot(candidate.root);
      return { view: ctx.view(ctx.state), edits: [], unchanged: true };
    }
    const indexed = await ctx.reindexRewrite(
      source,
      result,
      current.view.calculationPolicy,
      (indexed) => ctx.clearHyperlinksCalculation(indexed, current, result.rewriteId, 'note-authoring'),
      true
    );
    const actual = noteInventory(
      await ctx.native({
        action: 'workbook',
        operation: 'notes',
        root: indexed.root,
        version: indexed.version,
        sourceRevision: indexed.index.revision,
        sheet: command.sheet,
      }),
      { sheet: command.sheet }
    );
    const entry = actual.notes?.find((n) => n.cell === command.cell);
    if (
      !actual.canEdit ||
      actual.notes.length !== result.remainingNotes ||
      (command.type === 'removeNote'
        ? !!entry
        : entry?.author !== command.note.author || entry?.text !== command.note.text)
    )
      throw refusal('internal-mismatch', 'Note candidate differs from its accepted command.');
    ctx.recheck();
    indexed.journal = { previous: null, edits: [] };
    ctx.publish(candidate.next, indexed);
    return { view: ctx.view(candidate.next), edits: [command], commands: [command] };
  } catch (error) {
    await fs.rm(candidate.root, { recursive: true, force: true });
    ctx.dropRoot(candidate.root);
    throw error;
  }
}
