/** Sort, remove duplicates and filter rows. */
import { duplicateCommand } from '../duplicate-model.mjs';
import path from 'node:path';
import { matchesHandles } from './shared.mjs';
import { nativeTransform } from './transform.mjs';
import { referenceGuard } from './reference-guard.mjs';
import { refusal } from '../refusal-codes.mjs';
export const name = 'sort-filter';
export const handles = {
  operation: ['sortRows', 'removeDuplicates', 'filterRows', 'reapplyFilter', 'clearFilter', 'removeFilter'],
};
export const matches = (commands) => matchesHandles(handles, commands);
export async function run(ctx, commands) {
  const transform = commands[0];
  const kind = transform.operation;
  await referenceGuard(ctx, transform);
  if (kind === 'removeDuplicates') return removeDuplicates(ctx, commands);
  return nativeTransform(ctx, commands);
}
async function removeDuplicates(ctx, commands) {
  const { current } = ctx;
  const transform = commands[0];
  const command = duplicateCommand(transform),
    sheet = current.view.sheets.find((s) => s.name === command.sheet);
  if (!sheet) throw refusal('internal-unknown-reference', 'Unknown duplicate worksheet');
  const source = await ctx.rewriteSource();
  ctx.recheck();
  const result = await ctx.native({
    action: 'workbook',
    operation: 'dedup-apply',
    root: current.root,
    source: path.basename(source.path),
    ...ctx.rewriteFields(source),
    version: current.version,
    sheet: sheet.path,
    range: command.range,
    hasHeaders: command.hasHeaders,
    columns: command.columns,
    expectedFingerprint: command.expectedFingerprint,
  });
  ctx.rewriteCheck(source.candidate);
  const { next, root } = source.candidate;

  const indexed = await ctx.reindexRewrite(source, result, current.view.calculationPolicy);
  indexed.journal = { previous: null, edits: [] };
  ctx.recheck();
  ctx.publish(next, indexed);
  return { view: ctx.view(next), edits: [command], commands: [command] };
}
