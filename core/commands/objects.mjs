/** Charts, images, shapes, SmartArt, 3D models and ink strokes. */
import { isInkStrokeCommand } from '../ink-stroke.mjs';
import { isModel3DCommand } from '../model3d-authoring.mjs';
import { indexedObjectPlan, indexedObjectList, indexedObjectPreview } from '../indexed-object-authoring.mjs';
import fs from 'node:fs/promises';
import path from 'node:path';
import { boundedObjectMedia, matchesHandles } from './shared.mjs';
import { refusal } from '../refusal-codes.mjs';
export const name = 'objects';
export const handles = {
  operation: [
    'createChart',
    'updateChart',
    'insertImage',
    'updateImage',
    'deleteObject',
    'createShape',
    'updateShape',
    'deleteShape',
    'createSmartArt',
    'updateSmartArt',
    'deleteSmartArt',
    'createModel3D',
    'updateModel3D',
    'deleteModel3D',
    'createInkStroke',
    'updateInkStroke',
    'deleteInkStroke',
  ],
};
export const matches = (commands) => matchesHandles(handles, commands);
export async function run(ctx, commands) {
  const { current } = ctx;
  const transform = commands[0];
  const prepared = await indexedObjectPlan(
    {
      metadataBytes: await fs.readFile(path.join(current.root, 'metadata.xlsx')),
      view: current.view,
      readRange: (sheet, range) => ctx.read(current, sheet, range),
      readMedia: (part, { maxBytes }) =>
        boundedObjectMedia(path.join(current.root, 'source.xlsx'), part, maxBytes, ctx.signal),
    },
    transform
  );
  ctx.recheck();
  await ctx.native({
    action: 'workbook',
    operation: 'artifact-write',
    root: current.root,
    name: 'package-metadata.json',
    data: JSON.stringify(prepared.plan),
  });
  const source = await ctx.rewriteSource(),
    result = await ctx.native({
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
  // Refuse candidate metadata growth before publishing an unusable owner.
  const objectMetadata = await fs.readFile(path.join(root, 'metadata.xlsx'));
  indexedObjectList(objectMetadata, indexed.view, transform.sheet, { limit: 1 });
  if (isModel3DCommand(transform) && transform.operation !== 'deleteModel3D') {
    const proof = await indexedObjectPreview(
      {
        metadataBytes: objectMetadata,
        view: indexed.view,
        readMedia: (part, { maxBytes }) =>
          boundedObjectMedia(path.join(root, 'source.xlsx'), part, maxBytes, ctx.signal),
      },
      transform.sheet,
      transform.objectId
    );
    if (!proof.asset3d) throw refusal('session-bad-input', '3D candidate model was not validated');
  }
  if (isInkStrokeCommand(transform) && transform.operation !== 'deleteInkStroke') {
    const proof = await indexedObjectPreview(
        { metadataBytes: objectMetadata, view: indexed.view },
        transform.sheet,
        transform.objectId
      ),
      expected = prepared.command;
    const prior =
      transform.operation === 'updateInkStroke' ? await ctx.object(transform.sheet, transform.objectId) : null;
    if (
      proof.kind !== 'ink' ||
      JSON.stringify(proof.stroke) !== JSON.stringify(expected.stroke ?? prior?.stroke) ||
      JSON.stringify(proof.anchor) !== JSON.stringify(expected.anchor) ||
      proof.altText !== expected.altText
    )
      throw refusal('internal-mismatch', 'Native stroke candidate differs from its captured command');
  }
  indexed.journal = { previous: null, edits: [] };
  ctx.recheck();
  ctx.publish(next, indexed);
  return { view: ctx.view(next), edits: [prepared.command], commands: [prepared.command] };
}
