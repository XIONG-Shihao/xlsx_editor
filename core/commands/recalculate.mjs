/** Recalculate the workbook or one sheet. */
import { withSpillDimensions } from '../indexed-spills.mjs';
import { matchesHandles } from './shared.mjs';
import { refusal } from '../refusal-codes.mjs';
export const name = 'recalculate';
export const handles = { type: ['recalculate', 'calculateSheet'] };
export const matches = (commands) => matchesHandles(handles, commands);
export async function run(ctx, commands) {
  const { current } = ctx;
  const command = commands[0];
  let scope;
  if (command.type === 'calculateSheet') {
    if (Object.keys(command).sort().join(',') !== 'sheet,type')
      throw refusal('session-bad-input', 'Calculate Sheet requires exactly type and sheet');
    scope = await ctx.native({
      action: 'workbook',
      operation: 'calculation-policy',
      method: 'scope',
      root: current.root,
      version: current.version,
      view: current.view,
      sheet: command.sheet,
    });
  }
  const fork = await ctx.native({
    action: 'workbook',
    operation: 'calculation-policy',
    method: 'fork',
    root: current.root,
    version: current.version,
    view: current.view,
  });
  const next = ctx.nextState(),
    result = await ctx.calculate({
      root: current.root,
      version: fork.version,
      view: current.view,
      calculationIntent: true,
      ...(scope ? { calculationSheet: command.sheet, calculationPlan: scope.sourcePlan ?? scope.plan } : {}),
      revision: `${current.index.revision}:draft:${ctx.id}:${next}`,
    });
  let working = {
    ...current,
    view: withSpillDimensions(current.view, result),
    spillOwnership: current.spillOwnership || !!result.spillDimensions,
    version: result.version,
  };
  if (scope?.pivots?.length) working = await ctx.finishSelectedPivots(working, command.sheet, scope.pivots);
  working = scope ? await ctx.finishCharts(working, command.sheet) : await ctx.finishDerived(working);
  ctx.recheck();
  ctx.publish(next, {
    ...working,
    journal: { previous: working.root === current.root ? current.journal : null, edits: structuredClone(commands) },
  });
  return { view: ctx.view(next), edits: structuredClone(commands) };
}
