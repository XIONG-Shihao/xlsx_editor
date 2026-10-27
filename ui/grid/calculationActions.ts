import {
  type CalculationAction,
  type CalculationFrame,
  CalculationJobRetired,
  calculationFrameKey,
} from '../calculationJobs';
import { CalculationViewRetired } from './CalculationViewRetired';
import type { GridCalculationController, GridSetCalculationNotice } from './gridTypes';
import type { RibbonBinding, RibbonDispatchScope } from '../ribbonCommands';
import type { WorkbookCommand } from '../../core/contracts';
import type { WorkbookTool } from '../SpreadsheetRibbon';
import { acceptGridCommandBatch } from './commandBatch';

export interface CalculationActionsInputs {
  calculationController: GridCalculationController;
  calculationFrame: React.MutableRefObject<CalculationFrame | null>;
  calculationKey: string;
  calculationLive: React.MutableRefObject<{ key: string; active: boolean; normal: boolean }>;
  calculationMounted: React.MutableRefObject<boolean>;
  calculationOwnerAvailable: boolean;
  calculationRun: React.MutableRefObject<{ key: string } | null>;
  cancelCellFocus: () => void;
  commandScope: RibbonDispatchScope;
  normalMode: React.MutableRefObject<boolean>;
  onCommandBatch: ((commands: WorkbookCommand[]) => Promise<void>) | undefined;
  passiveNotice: string | undefined;
  setActiveTool: React.Dispatch<React.SetStateAction<WorkbookTool | null>>;
  setCalculationBusy: React.Dispatch<React.SetStateAction<boolean>>;
  setCalculationNotice: GridSetCalculationNotice;
  setGridError: React.Dispatch<React.SetStateAction<string>>;
  setVisitedTools: React.Dispatch<React.SetStateAction<WorkbookTool[]>>;
  specialSelectionMode: React.MutableRefObject<{ selection: Set<string> | null }>;
}

export function buildCalculationActions({
  calculationController,
  calculationFrame,
  calculationKey,
  calculationLive,
  calculationMounted,
  calculationOwnerAvailable,
  calculationRun,
  cancelCellFocus,
  commandScope,
  normalMode,
  onCommandBatch,
  passiveNotice,
  setActiveTool,
  setCalculationBusy,
  setCalculationNotice,
  setGridError,
  setVisitedTools,
  specialSelectionMode,
}: CalculationActionsInputs) {
  const calculateCurrentWorkbook = async (binding: RibbonBinding): Promise<void> => {
    if (calculationRun.current) throw Error('Wait for the current calculation to finish.');
    if (!commandScope.isCurrent(binding) || !normalMode.current || calculationLive.current.key !== calculationKey)
      throw Error('The active workbook changed before calculation.');
    if (!calculationOwnerAvailable) throw Error('Calculation is unavailable in this editor.');
    const available = commandScope.availability('formulas.calculate-now'),
      batch = onCommandBatch;
    if (available.profile !== 'xlsx' || passiveNotice) throw Error('This command requires an editable XLSX workbook.');
    const run = { key: calculationKey },
      current = (): boolean =>
        calculationMounted.current &&
        calculationLive.current.active &&
        calculationLive.current.normal &&
        calculationLive.current.key === run.key;
    calculationRun.current = run;
    setCalculationBusy(true);
    setGridError('');
    setCalculationNotice({ key: run.key, phase: 'running', text: 'Calculating the current workbook…' });
    try {
      await acceptGridCommandBatch([{ type: 'recalculate' }], {
        active: available.active,
        readOnly: !available.canEdit,
        busy: available.busy,
        literal: available.profile !== 'xlsx',
        specialSelection: !!specialSelectionMode.current.selection,
        onCommandBatch: batch,
      });
      if (current())
        setCalculationNotice({ key: run.key, phase: 'completed', text: 'Current workbook calculation finished.' });
    } catch (error) {
      if (!current()) throw new CalculationViewRetired(error);
      setCalculationNotice({ key: run.key, phase: 'failed', text: 'Current workbook calculation was not accepted.' });
      throw error;
    } finally {
      if (calculationRun.current === run) calculationRun.current = null;
      if (calculationMounted.current) setCalculationBusy(false);
    }
  };
  const runCalculationAction = async (
    binding: RibbonBinding,
    action: CalculationAction,
    expectedKey: string | null
  ): Promise<void> => {
    if (
      !commandScope.isCurrent(binding) ||
      !calculationFrame.current ||
      !expectedKey ||
      calculationFrameKey(calculationFrame.current) !== expectedKey
    )
      throw new CalculationJobRetired();
    cancelCellFocus();
    setVisitedTools((previous) => (previous.includes('calculation') ? previous : [...previous, 'calculation']));
    setActiveTool('calculation');
    await calculationController.run(action, expectedKey);
  };
  return { calculateCurrentWorkbook, runCalculationAction } as const;
}
