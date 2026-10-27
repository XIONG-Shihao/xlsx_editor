/** Page setup and print areas. */
import { indexedPrintDefinition, indexedPrintPlan } from '../indexed-print.mjs';
import fs from 'node:fs/promises';
import path from 'node:path';
import { matchesHandles } from './shared.mjs';
import { refusal } from '../refusal-codes.mjs';
export const name = 'print';
export const handles = { operation: ['setPrintSettings'] };
export const matches = (commands) => matchesHandles(handles, commands);
export async function run(ctx, commands) {
  const { current } = ctx;
  const transform = commands[0];
  const metadataBytes = await fs.readFile(path.join(current.root, 'metadata.xlsx'));
  let prepared = indexedPrintPlan(metadataBytes, current.view, transform);
  ctx.recheck();
  const source = await ctx.rewriteSource();
  // Calculation export can add calcPr readiness flags. Preserve that
  // trusted exported workbook XML when adding the reserved print names.
  if (path.basename(source.path) === 'export.xlsx') {
    const exportMetadata = JSON.parse(await fs.readFile(path.join(current.root, 'export-metadata.json'), 'utf8'));
    if (typeof exportMetadata.replacements?.['xl/workbook.xml'] === 'string')
      prepared = indexedPrintPlan(metadataBytes, current.view, transform, {
        workbookXml: exportMetadata.replacements['xl/workbook.xml'],
      });
  }
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
  const definition = indexedPrintDefinition(
    await fs.readFile(path.join(root, 'metadata.xlsx')),
    indexed.view,
    transform.sheet
  );
  if (!definition.canEdit) throw refusal('session-unsupported', 'Printed settings candidate is not qualified');
  indexed.journal = { previous: null, edits: [] };
  ctx.recheck();
  ctx.publish(next, indexed);
  return { view: ctx.view(next), edits: [prepared.command], commands: [prepared.command] };
}
