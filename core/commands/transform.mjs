import { refusal, reworded } from '../refusal-codes.mjs';
import path from 'node:path';
import { normalizeSortKeys } from '../sort-keys.mjs';
import { empty, nativeValue } from './shared.mjs';
/** Sort, filter and plain row/column changes: Rust rewrites the sheet, then the result is reindexed. */
export async function nativeTransform(ctx, commands) {
  const { current } = ctx;
  const transform = commands[0];
  const kind = transform.operation;
  try {
    if (kind === 'sortRows') normalizeSortKeys(transform);
    if (
      ['reapplyFilter', 'clearFilter', 'removeFilter'].includes(kind) &&
      (Object.keys(transform).length !== 3 || !/^[a-f0-9]{64}$/.test(transform.expectedFilterRevision))
    )
      throw refusal('internal-invalid-field', 'Invalid accepted filter command fields');
    const sheet = current.view.sheets.find((s) => s.name === transform.sheet);
    if (!sheet) throw refusal('internal-unknown-reference', 'Unknown worksheet');
    // Sorting moves formulas to other rows, so shared groups are written out first.
    const sorted =
      kind === 'sortRows' ? await ctx.expandShared(current, '', transform.sheet).catch(() => current) : current;
    const source = await ctx.rewriteSource(true, sorted);
    const request = {
      ...nativeValue(transform),
      action: 'workbook',
      operation: 'transform',
      root: current.root,
      source: path.basename(source.path),
      version: sorted.version,
      ...ctx.rewriteFields(source),
      kind,
      sheet: sheet.path,
      view: empty(current.view),
    };
    const result = await ctx.native(request);
    ctx.rewriteCheck(source.candidate);
    const { next, root } = source.candidate;
    const indexed = await ctx.reindexRewrite(source, result, current.view.calculationPolicy);
    ctx.recheck();
    ctx.publish(next, indexed);
    return { view: ctx.view(next), edits: structuredClone(commands) };
  } catch (error) {
    throw reworded(
      error,
      String(error).replace('requires the shared structural writer', 'is not qualified for native preservation')
    );
  }
}
