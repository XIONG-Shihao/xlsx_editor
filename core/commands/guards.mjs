import { indexedPivotMutationGuard } from '../indexed-pivots.mjs';
import { indexedValidationMutationGuard } from '../indexed-rules.mjs';
import fs from 'node:fs/promises';
import path from 'node:path';
import { refusal } from '../refusal-codes.mjs';
/** Refuse a command that touches something a pivot or a validation rule protects. */
export async function admitAgainstMetadata(ctx, commands) {
  const { current } = ctx;
  if (current.view.pivots?.length) {
    const metadata = await fs.readFile(path.join(current.root, 'metadata.xlsx'));
    for (const command of commands) {
      indexedPivotMutationGuard(metadata, current.view, command);
      if (command.type === 'fillAcrossSheets')
        indexedPivotMutationGuard(metadata, current.view, { ...command, sheet: command.sourceSheet });
    }
    ctx.recheck();
  }
  ctx.ruleDefinitions = await ctx.loadRuleDefinitions();
  ctx.recheck();
  if (
    commands.some((command) => {
      const name = command.sheet ?? command.shape?.sheet;
      return name && ctx.ruleDefinitions.validation[name]?.revision;
    })
  ) {
    const bytes = await fs.readFile(path.join(current.root, 'metadata.xlsx'));
    for (const command of commands) {
      const reason = indexedValidationMutationGuard(bytes, current.view, command);
      if (reason) throw refusal('workbook-refused', reason);
    }
  }
}
