/** Change the calculation mode. */
import { withSpillDimensions } from '../indexed-spills.mjs';
import { matchesHandles } from './shared.mjs';
import { refusal } from '../refusal-codes.mjs';
export const name = 'calculation-options';
export const handles = { type: ['setCalculationOptions'] };
export const matches = (commands) => matchesHandles(handles, commands);
export async function run(ctx, commands) {
  const { current } = ctx;
  const command = commands[0];
  if (Object.keys(command).some((k) => !['type', 'mode', 'recalculateBeforeSave'].includes(k)))
    throw refusal('internal-invalid-field', 'Invalid calculation options fields');
  const changed = await ctx.native({
    action: 'workbook',
    operation: 'calculation-policy',
    method: 'options',
    root: current.root,
    version: current.version,
    view: current.view,
    mode: command.mode,
    ...(command.recalculateBeforeSave !== undefined ? { recalculateBeforeSave: command.recalculateBeforeSave } : {}),
  });
  let result = { version: changed.version, calculationPolicy: changed.policy };
  const view = { ...current.view, calculationPolicy: changed.policy };
  if (changed.policy.mode === 'auto')
    result = await ctx.calculate({
      root: current.root,
      version: changed.version,
      view,
      calculationIntent: true,
      revision: `${current.index.revision}:options`,
    });
  let working = {
    ...current,
    version: result.version,
    view: withSpillDimensions(view, result),
    spillOwnership: current.spillOwnership || !!result.spillDimensions,
  };
  if (changed.policy.mode === 'auto') working = await ctx.finishDerived(working);
  const next = ctx.nextState();
  ctx.recheck();
  ctx.publish(next, {
    ...working,
    journal: { previous: working.root === current.root ? current.journal : null, edits: structuredClone(commands) },
  });
  return { view: ctx.view(next), edits: structuredClone(commands) };
}
