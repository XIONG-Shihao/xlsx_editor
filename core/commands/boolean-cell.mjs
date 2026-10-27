import { FORMAT, matchesFormat } from '../format.mjs';
/** Insert, toggle or clear a checkbox cell; the value edit itself goes back through the ordinary cell path. */
import { normalizeBooleanCellCommands } from '../cell-insertions.mjs';
import { FACTS_V2, matchesHandles } from './shared.mjs';
import { refusal } from '../refusal-codes.mjs';
export const name = 'boolean-cell';
export const handles = { type: ['booleanCell'] };
export const matches = (commands) => matchesHandles(handles, commands);
export async function run(ctx, commands) {
  const { current } = ctx;
  const [command] = normalizeBooleanCellCommands(commands);
  if (!matchesFormat(current.index.sourceFactSchema, 'SCHEMA_USED_RANGE_FACTS_V2'))
    throw refusal('session-unsupported', 'Boolean-cell requires qualified native source ownership.');
  const admitted = await ctx.native({
    action: 'workbook',
    operation: 'boolean-cell-check',
    root: current.root,
    version: current.version,
    sourceRevision: current.index.revision,
    command,
  });
  ctx.recheck();
  const value = command.action === 'insert' ? false : command.action === 'toggle' ? !command.expectedValue : null;
  if (
    !admitted ||
    typeof admitted !== 'object' ||
    Array.isArray(admitted) ||
    Object.keys(admitted).sort().join(',') !== 'cell,changed,schema,sheet,value' ||
    !matchesFormat(admitted.schema, 'SCHEMA_BOOLEAN_CELL_ADMISSION_V1') ||
    admitted.sheet !== command.sheet ||
    admitted.cell !== command.cell ||
    admitted.value !== value ||
    admitted.changed !== (value !== command.expectedValue)
  )
    throw refusal('internal-invalid-field', 'Invalid native Boolean-cell admission receipt.');
  if (!admitted.changed) return { view: ctx.view(ctx.state), edits: [], unchanged: true };
  // Reuse ordinary validation/pivot/spill/value guards, calculation and formatting.
  // Durable replay retains the semantic command and repeats native graph admission.
  const result = await ctx.mutate([
    {
      sheet: command.sheet,
      cell: command.cell,
      value,
      expectedValue: command.expectedValue,
      expectedFormula: null,
    },
  ]);
  return { ...result, commands: [command] };
}
