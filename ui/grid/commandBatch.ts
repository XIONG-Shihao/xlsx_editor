import type { WorkbookCommand } from '../../core/contracts';

export const COMMAND_BATCH_UNAVAILABLE = 'Workbook commands are unavailable in this editor.';
/** Shared normalization preserves one owner batch while respecting each
 * setCells command's existing 1,000-position transport bound. */
export function normalizeGridCommandBatch(commands: WorkbookCommand[]): WorkbookCommand[] {
  return commands.flatMap((command) =>
    'type' in command && command.type === 'setCells' && command.cells.length > 1000
      ? Array.from({ length: Math.ceil(command.cells.length / 1000) }, (_, index) => ({
          ...command,
          cells: command.cells.slice(index * 1000, (index + 1) * 1000),
        }))
      : [command]
  );
}
export async function acceptGridCommandBatch(
  commands: WorkbookCommand[],
  scope: {
    active: boolean;
    readOnly: boolean;
    busy: boolean;
    literal?: boolean;
    specialSelection: boolean;
    onCommandBatch?: (commands: WorkbookCommand[]) => Promise<void>;
  },
  allowSpecial = false
): Promise<void> {
  if (!scope.active) throw Error('Activate this workbook before editing.');
  if (scope.readOnly) throw Error('This workbook is read-only.');
  if (scope.busy) throw Error('Wait for the workbook operation to finish.');
  if (scope.literal) throw Error('This command requires an editable XLSX workbook.');
  if (scope.specialSelection && !allowSpecial)
    throw Error('Special selection supports formatting and Clear. Click a cell for other actions.');
  if (!scope.onCommandBatch) throw Error(COMMAND_BATCH_UNAVAILABLE);
  await scope.onCommandBatch(normalizeGridCommandBatch(commands));
}
