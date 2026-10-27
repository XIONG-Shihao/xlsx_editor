import type { WorkbookCommand } from '../../core/contracts';
import { acceptGridCommandBatch, normalizeGridCommandBatch } from './commandBatch';

export interface CommandSendInputs {
  literal: boolean | undefined;
  normalMode: React.MutableRefObject<boolean>;
  onBatch: ((commands: WorkbookCommand[]) => void) | undefined;
  onCommandBatch: ((commands: WorkbookCommand[]) => Promise<void>) | undefined;
  readOnly: boolean;
  setGridError: React.Dispatch<React.SetStateAction<string>>;
  specialSelection: Set<string> | null;
  viewActive: React.MutableRefObject<boolean>;
  /** Read when the code runs, like the closed-over variables they stand for. */
  late: {
    readonly busy: boolean;
  };
}

export function buildCommandSend({
  literal,
  normalMode,
  onBatch,
  onCommandBatch,
  readOnly,
  setGridError,
  specialSelection,
  viewActive,
  late,
}: CommandSendInputs) {
  const send = (commands: WorkbookCommand[], allowSpecial = false): void => {
    if (!normalMode.current) {
      setGridError('Return to Normal view to edit the worksheet.');
      return;
    }
    if (specialSelection && !allowSpecial) {
      setGridError('Special selection supports formatting and Clear. Click a cell for other actions.');
      return;
    }
    // The shared MCP contract admits 1,000 cells per setCells command.
    // Keep larger user actions in one atomic batch containing bounded commands.
    const batch = normalizeGridCommandBatch(commands);
    if (!readOnly && !late.busy) onBatch?.(batch);
  };
  const sendCommand = (commands: WorkbookCommand[], allowSpecial = false): Promise<void> =>
    acceptGridCommandBatch(
      commands,
      {
        active: viewActive.current,
        readOnly: readOnly || !normalMode.current,
        busy: !!late.busy,
        literal,
        specialSelection: !!specialSelection,
        onCommandBatch,
      },
      allowSpecial
    );
  return { send, sendCommand } as const;
}
