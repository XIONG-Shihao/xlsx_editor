import path from 'node:path';
/** Write the prepared metadata plan, rewrite the package from it and publish the result. */
export async function finishPackageEdit(ctx, commands, { metadata, source, base }) {
  const { current } = ctx;
  await ctx.native({
    action: 'workbook',
    operation: 'artifact-write',
    root: current.root,
    name: 'package-metadata.json',
    data: JSON.stringify(metadata),
  });
  source ??= await ctx.rewriteSource(true, base);
  const result = await ctx.native({
    action: 'workbook',
    operation: 'package-edit',
    root: current.root,
    version: base.version,
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
