import { FORMAT, matchesFormat } from '../format.mjs';
/** Apply a page theme. */
import { normalizePageThemeCommands } from '../page-theme-command.mjs';
import fs from 'node:fs/promises';
import path from 'node:path';
import { FACTS_V2, uuidV4, matchesHandles } from './shared.mjs';
import { refusal } from '../refusal-codes.mjs';
export const name = 'page-theme';
export const handles = { type: ['applyPageTheme'] };
export const matches = (commands) => matchesHandles(handles, commands);
export async function run(ctx, commands) {
  const { current } = ctx;
  const [command] = normalizePageThemeCommands(commands);
  if (!matchesFormat(current.index.sourceFactSchema, 'SCHEMA_USED_RANGE_FACTS_V2'))
    throw refusal('session-unsupported', 'Page theme requires qualified native source ownership.');
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
        'changed,contentChanged,file,priorThemeRevision,rewriteId,sample,schema,themeRevision' ||
      !matchesFormat(result.schema, 'SCHEMA_PAGE_THEME_WRITER_V1') ||
      result.file !== 'package.xlsx' ||
      !uuidV4.test(result.rewriteId) ||
      result.contentChanged !== false ||
      typeof result.changed !== 'boolean' ||
      result.priorThemeRevision !== command.expectedThemeRevision ||
      result.sample !== command.sample ||
      typeof result.themeRevision !== 'string' ||
      !/^[a-f0-9]{64}$/.test(result.themeRevision)
    )
      throw refusal('internal-invalid-field', 'Invalid native page theme writer receipt.');
    if (!result.changed) {
      await fs.rm(candidate.root, { recursive: true, force: true });
      ctx.dropRoot(candidate.root);
      return { view: ctx.view(ctx.state), edits: [], unchanged: true };
    }
    const indexed = await ctx.reindexRewrite(
      source,
      result,
      current.view.calculationPolicy,
      (indexed) => ctx.clearHyperlinksCalculation(indexed, current, result.rewriteId, 'page-theme'),
      true
    );
    if (
      indexed.view.pageTheme?.sample !== command.sample ||
      indexed.view.pageTheme?.revision !== result.themeRevision ||
      !indexed.view.pageTheme?.canApply
    )
      throw refusal('session-bad-input', 'Page theme candidate does not match the accepted sample.');
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
