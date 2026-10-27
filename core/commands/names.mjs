/** Defined names. */
import { indexedNamePlan } from '../indexed-names.mjs';
import fs from 'node:fs/promises';
import { unzipSync, zipSync } from 'fflate';
import path from 'node:path';
import { boundedObjectMedia, matchesHandles } from './shared.mjs';
export const name = 'names';
export const handles = { operation: ['defineName', 'removeName', 'updateName'] };
export const matches = (commands) => matchesHandles(handles, commands);
export async function run(ctx, commands) {
  const { current } = ctx;
  const transform = commands[0];
  const metadataBytes = await fs.readFile(path.join(current.root, 'metadata.xlsx'));
  // Validate the command before export. Its final raw workbook envelope
  // must also include the current accepted calculation-policy flags.
  indexedNamePlan(metadataBytes, current.view, transform);
  const source = await ctx.rewriteSource(),
    parts = unzipSync(metadataBytes);
  parts['xl/workbook.xml'] = await boundedObjectMedia(source.path, 'xl/workbook.xml', 16 * 1024 ** 2, ctx.signal);
  const metadata = indexedNamePlan(zipSync(parts), current.view, transform);
  await ctx.native({
    action: 'workbook',
    operation: 'artifact-write',
    root: current.root,
    name: 'package-metadata.json',
    data: JSON.stringify(metadata),
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
  indexed.journal = { previous: null, edits: [] };
  ctx.recheck();
  ctx.publish(next, indexed);
  return { view: ctx.view(next), edits: structuredClone(commands) };
}
