/** Paste staged cells from an import. */
import { applyIndexedImport } from '../indexed-import.mjs';
import { validatePasteDescriptor } from '../paste-protocol.mjs';
import { withSpillDimensions } from '../indexed-spills.mjs';
import { matchesHandles } from './shared.mjs';
import { refusal } from '../refusal-codes.mjs';
export const name = 'import-cells';
export const handles = { type: ['importCells'] };
export const matches = (commands) => matchesHandles(handles, commands);
export async function run(ctx, commands) {
  const { current } = ctx;
  const transform = commands[0];
  const descriptor = validatePasteDescriptor(transform);
  if (typeof ctx.context?.importPages !== 'function')
    throw refusal('session-bad-input', 'Staged paste requires a trusted owner attachment resolver');
  const check = () => {
    ctx.signal.throwIfAborted();
    ctx.recheck();
  };
  try {
    check();
    const pages = ctx.context.importPages(descriptor);
    check();
    const imported = await applyIndexedImport({
      descriptor,
      pages,
      current,
      native: ctx.native,
      checkCancelled: check,
      progress: (event) => {
        check();
        ctx.context.importProgress?.(event);
      },
    });
    const working = imported.working,
      next = ctx.nextState();
    check();
    if (
      working.view.sheets.some((sheet) => sheet.storedFormulaCount > 0) ||
      working.view.calculationPolicy?.mode === 'manual'
    ) {
      const calculated = await ctx.calculate({
        root: working.root,
        version: working.version,
        view: working.view,
        changedRanges: imported.changedRanges,
        formulaChanged: imported.formulaChanged,
        revision: `${current.index.revision}:draft:${ctx.id}:${next}`,
      });
      working.version = calculated.version;
      working.view = withSpillDimensions(working.view, calculated);
      working.spillOwnership ||= !!calculated.spillDimensions;
    }
    check();
    ctx.publish(next, { ...working, journal: { previous: current.journal, edits: [descriptor] } });
    return { view: ctx.view(next), edits: [descriptor], commands: [descriptor] };
  } catch (error) {
    // Failed pages/calculation must not leave private native versions as
    // new bases. Preserve every accepted old/redo state in every root.
    if (!ctx.isClosed()) {
      const accepted = new Map();
      for (const value of ctx.acceptedStates()) {
        if (!accepted.has(value.root)) accepted.set(value.root, new Set());
        accepted.get(value.root).add(value.version);
      }
      for (const [root, versions] of accepted) {
        try {
          await ctx.native({ action: 'workbook', operation: 'compact', root, keep: [...versions] });
        } catch {
          /* A closing/cancelled owner retains its original failure and cleanup path. */
        }
      }
    }
    throw error;
  }
}
