/** Insert and delete rows and columns. */
import { withSpillDimensions } from '../indexed-spills.mjs';
import path from 'node:path';
import { matchesHandles } from './shared.mjs';
import { nativeTransform } from './transform.mjs';
import { referenceGuard } from './reference-guard.mjs';
export const name = 'structure';
export const handles = { operation: ['insertRow', 'deleteRow', 'insertColumn', 'deleteColumn'] };
export const matches = (commands) => matchesHandles(handles, commands);
// The refusal names the edit that was asked for, not only row insertion.
const refusedEdit = {
  insertRow: 'Inserting rows',
  deleteRow: 'Deleting rows',
  insertColumn: 'Inserting columns',
  deleteColumn: 'Deleting columns',
};
const extendedReason = 'cannot keep special formulas or extended worksheet references on this sheet correct';
export async function run(ctx, commands) {
  const { current } = ctx;
  const transform = commands[0];
  await referenceGuard(ctx, transform);
  const plain =
    !current.view.structuralReason &&
    !current.view.sheets.some((s) => s.storedFormulaCount) &&
    !current.view.namedRanges?.length &&
    !current.view.tables?.length;
  // Formulas, names or tables may refer to the moved cells, so those go through the reference rewriter.
  if (!plain || (transform.count ?? 1) !== 1) {
    const expanded = await ctx.expandShared(
      current,
      `${refusedEdit[transform.operation]} ${extendedReason}`,
      transform.sheet
    );
    const source = await ctx.rewriteSource(true, expanded);
    const prepared = await ctx.calculate({
      root: current.root,
      version: expanded.version,
      view: current.view,
      command: transform,
      ...(source.candidate.enabled ? { source: path.basename(source.path), ...ctx.rewriteFields(source) } : {}),
    });
    const result = await ctx.native({
      action: 'workbook',
      operation: 'insert-references',
      root: current.root,
      ...(source.candidate.enabled ? { version: expanded.version } : {}),
      source: path.basename(source.path),
      ...ctx.rewriteFields(source),
    });
    ctx.rewriteCheck(source.candidate);
    const { next, root } = source.candidate;
    const indexed = await ctx.reindexRewrite(source, result, current.view.calculationPolicy);
    indexed.journal = { previous: null, edits: [] };
    if (prepared.headers?.length) {
      const changes = prepared.headers.map((edit) => ({
        sheet: indexed.view.sheets.find((s) => s.name === edit.sheet).path,
        cell: edit.cell,
        value: { value: edit.value, dataType: 'text', rawValue: edit.value, cachedValue: null },
        edit,
      }));
      for (let start = 0; start < changes.length; start += 128) {
        const patch = await ctx.native({
          action: 'workbook',
          operation: 'patch',
          root,
          version: indexed.version,
          changes: changes.slice(start, start + 128),
        });
        indexed.version = patch.version;
      }
      if (indexed.view.sheets.some((s) => s.storedFormulaCount)) {
        const calculated = await ctx.calculate({
          root,
          version: indexed.version,
          view: indexed.view,
          changedRanges: prepared.headers.map((e) => ({ sheet: e.sheet, range: e.cell })),
          formulaChanged: false,
          revision: indexed.index.revision,
        });
        indexed.version = calculated.version;
        indexed.view = withSpillDimensions(indexed.view, calculated);
        indexed.spillOwnership ||= !!calculated.spillDimensions;
      }
    }
    ctx.recheck();
    ctx.publish(next, indexed);
    return { view: ctx.view(next), edits: structuredClone(commands) };
  }
  return nativeTransform(ctx, commands);
}
