/** Create, refresh and change pivot tables. */
import { indexedPivotPlan } from '../indexed-pivots.mjs';
import fs from 'node:fs/promises';
import path from 'node:path';
import { boundedObjectMedia, matchesHandles } from './shared.mjs';
import { refusal } from '../refusal-codes.mjs';
export const name = 'pivots';
export const handles = { operation: ['createPivot', 'refreshPivot', 'updatePivot', 'setPivotLayout'] };
export const matches = (commands) => matchesHandles(handles, commands);
export async function run(ctx, commands) {
  const { current } = ctx;
  const transform = commands[0];
  const source = await ctx.rewriteSource();
  const prepared = await indexedPivotPlan({
    metadataBytes: await fs.readFile(path.join(current.root, 'metadata.xlsx')),
    view: current.view,
    command: transform,
    readRange: (sheet, range) => ctx.read(current, sheet, range),
    readWorksheet: (part, { maxBytes }) => boundedObjectMedia(source.path, part, maxBytes, ctx.signal),
  });
  ctx.recheck();
  await ctx.native({
    action: 'workbook',
    operation: 'artifact-write',
    root: current.root,
    name: 'package-metadata.json',
    data: JSON.stringify(prepared.plan),
  });
  const result = await ctx.native({
    action: 'workbook',
    operation: 'package-edit',
    root: current.root,
    version: current.version,
    source: path.basename(source.path),
    ...ctx.rewriteFields(source),
  });
  ctx.rewriteCheck(source.candidate);
  const { next, root } = source.candidate;
  const indexed = await ctx.reindexRewrite(source, result, current.view.calculationPolicy);
  if (!indexed.view.pivots?.some((p) => p.name === prepared.command.name && p.canEdit))
    throw refusal(
      'session-unsupported',
      'Pivot candidate is not qualified: ' + indexed.view.pivots?.map((p) => p.reason ?? p.name).join('; ')
    );
  indexed.journal = { previous: null, edits: [] };
  ctx.recheck();
  ctx.publish(next, indexed);
  return { view: ctx.view(next), edits: [prepared.command], commands: [prepared.command] };
}
