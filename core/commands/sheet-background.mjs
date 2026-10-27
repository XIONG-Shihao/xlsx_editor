import { FORMAT, matchesFormat } from '../format.mjs';
/** Set or remove a sheet background picture. */
import { normalizeBackgroundCommands } from '../background-command.mjs';
import fs from 'node:fs/promises';
import path from 'node:path';
import { FACTS_V2, uuidV4, matchesHandles } from './shared.mjs';
import { refusal } from '../refusal-codes.mjs';
export const name = 'sheet-background';
export const handles = { type: ['setSheetBackground', 'removeSheetBackground'] };
export const matches = (commands) => matchesHandles(handles, commands);
export async function run(ctx, commands) {
  const { current } = ctx;
  const [command] = normalizeBackgroundCommands(commands);
  if (
    !matchesFormat(current.index.sourceFactSchema, 'SCHEMA_USED_RANGE_FACTS_V2') ||
    !current.view.sheets.some((sheet) => sheet.name === command.sheet)
  )
    throw refusal('session-unsupported', 'Background requires a qualified native worksheet owner.');
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
    if (
      !result ||
      typeof result !== 'object' ||
      Array.isArray(result) ||
      Object.keys(result).sort().join(',') !==
        'changed,contentChanged,file,priorBackgroundRevision,rewriteId,sample,schema,sheet' ||
      !matchesFormat(result.schema, 'SCHEMA_BACKGROUND_WRITER_V1') ||
      result.file !== 'package.xlsx' ||
      !uuidV4.test(result.rewriteId) ||
      result.contentChanged !== false ||
      typeof result.changed !== 'boolean' ||
      result.priorBackgroundRevision !== command.expectedBackgroundRevision ||
      result.sheet !== current.view.sheets.find((sheet) => sheet.name === command.sheet).path ||
      result.sample !== (command.type === 'setSheetBackground' ? command.sample : null)
    )
      throw refusal('internal-invalid-field', 'Invalid native background writer receipt.');
    if (!result.changed) {
      await fs.rm(candidate.root, { recursive: true, force: true });
      ctx.dropRoot(candidate.root);
      return { view: ctx.view(ctx.state), edits: [], unchanged: true };
    }
    const indexed = await ctx.reindexRewrite(
      source,
      result,
      current.view.calculationPolicy,
      (indexed) => ctx.clearHyperlinksCalculation(indexed, current, result.rewriteId, 'sheet-background'),
      true
    );
    const background = indexed.view.sheets.find((sheet) => sheet.name === command.sheet)?.background;
    if (
      command.type === 'setSheetBackground'
        ? background?.sample !== command.sample || !background?.canEdit
        : !!background
    )
      throw refusal('session-bad-input', 'Background candidate does not match the accepted operation.');
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
