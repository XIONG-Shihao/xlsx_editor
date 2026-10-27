/** Add, rename and duplicate sheets. */
import { indexedCopyRelationships } from '../indexed-copy.mjs';
import { indexedRenameMetadata } from '../indexed-objects.mjs';
import fs from 'node:fs/promises';
import { unzipSync, zipSync } from 'fflate';
import path from 'node:path';
import { indexedPackageEdit, indexedPreservedSheetAppend } from '../workbook-package.mjs';
import { resolveWorkbookEdits } from '../workbook-commands.mjs';
import { empty, boundedObjectMedia, matchesHandles } from './shared.mjs';
import { finishPackageEdit } from './package-edit.mjs';
import { referenceGuard } from './reference-guard.mjs';
export const name = 'sheets';
export const handles = { operation: ['addSheet', 'renameSheet', 'duplicateSheet'] };
export const matches = (commands) => matchesHandles(handles, commands);
export async function run(ctx, commands) {
  const { current } = ctx;
  const transform = commands[0];
  const kind = transform.operation;
  await referenceGuard(ctx, transform);
  const renamed =
    kind === 'renameSheet'
      ? await ctx.expandShared(
          current,
          'Special formulas or extended references require rename qualification',
          transform.sheet
        )
      : kind === 'duplicateSheet'
        ? // A copy's formulas are rewritten for the new sheet name, so shared groups are written out first.
          await ctx.expandShared(current, '', transform.sheet).catch(() => current)
        : current;
  let source;
  let metadataBytes = await fs.readFile(path.join(current.root, 'metadata.xlsx'));
  if (kind === 'addSheet') {
    resolveWorkbookEdits(empty(current.view), [transform]);
    source = await ctx.rewriteSource();
    const parts = unzipSync(metadataBytes);
    for (const part of ['xl/workbook.xml', 'xl/_rels/workbook.xml.rels', '[Content_Types].xml'])
      parts[part] = await boundedObjectMedia(source.path, part, 16 * 1024 ** 2, ctx.signal);
    metadataBytes = zipSync(parts);
  }
  const metadata =
    kind === 'addSheet'
      ? indexedPreservedSheetAppend(metadataBytes, transform, current.view)
      : indexedPackageEdit(metadataBytes, transform, current.view, []);
  if (kind === 'duplicateSheet')
    await indexedCopyRelationships(
      await fs.readFile(path.join(current.root, 'metadata.xlsx')),
      metadata,
      transform,
      ctx.native,
      current.root
    );
  if (kind === 'renameSheet')
    await indexedRenameMetadata(
      await fs.readFile(path.join(current.root, 'metadata.xlsx')),
      metadata,
      ctx.native,
      current.root
    );
  return finishPackageEdit(ctx, commands, { metadata, source, base: renamed });
}
