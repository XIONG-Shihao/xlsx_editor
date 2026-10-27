import { FORMAT, matchesFormat } from '../format.mjs';
/** Clear the hyperlinks in a range. */
import fs from 'node:fs/promises';
import path from 'node:path';
import { normalizeClearHyperlinksCommands } from '../clear-hyperlinks.mjs';
import { FACTS_V2, uuidV4, matchesHandles } from './shared.mjs';
import { refusal } from '../refusal-codes.mjs';
export const name = 'clear-hyperlinks';
export const handles = { type: ['clearHyperlinks'] };
export const matches = (commands) => matchesHandles(handles, commands);
export async function run(ctx, commands) {
  const { current } = ctx;
  const [command] = normalizeClearHyperlinksCommands(commands, current.view);
  if (!matchesFormat(current.index.sourceFactSchema, 'SCHEMA_USED_RANGE_FACTS_V2'))
    throw refusal('session-unsupported', 'Clear Hyperlinks requires qualified native source ownership.');
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
    const fields =
      'changed,contentChanged,file,remainingHyperlinks,removedHyperlinks,removedRelationships,rewriteId,schema,trimmedHyperlinks';
    if (
      !result ||
      typeof result !== 'object' ||
      Array.isArray(result) ||
      Object.keys(result).sort().join(',') !== fields ||
      !matchesFormat(result.schema, 'SCHEMA_CLEAR_HYPERLINKS_WRITER_V1') ||
      result.file !== 'package.xlsx' ||
      !uuidV4.test(result.rewriteId) ||
      result.contentChanged !== false ||
      typeof result.changed !== 'boolean' ||
      ['removedHyperlinks', 'trimmedHyperlinks', 'removedRelationships', 'remainingHyperlinks'].some(
        (key) => !Number.isSafeInteger(result[key]) || result[key] < 0
      ) ||
      result.changed !== result.removedHyperlinks + result.trimmedHyperlinks > 0
    )
      throw refusal('internal-invalid-field', 'Invalid Clear Hyperlinks native writer receipt.');
    if (!result.changed) {
      await fs.rm(candidate.root, { recursive: true, force: true });
      ctx.dropRoot(candidate.root);
      return { view: ctx.view(ctx.state), edits: [], unchanged: true };
    }
    const indexed = await ctx.reindexRewrite(
      source,
      result,
      current.view.calculationPolicy,
      (indexed) => ctx.clearHyperlinksCalculation(indexed, current, result.rewriteId),
      true
    );
    ctx.recheck();
    indexed.journal = { previous: null, edits: [] };
    ctx.publish(candidate.next, indexed);
    return { view: ctx.view(candidate.next), edits: [command] };
  } catch (error) {
    await fs.rm(candidate.root, { recursive: true, force: true });
    ctx.dropRoot(candidate.root);
    throw error;
  }
}
