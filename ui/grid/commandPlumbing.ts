import { COMMAND_BATCH_UNAVAILABLE } from './commandBatch';
import { type CalculationFrame, CalculationJobRetired, createCalculationJobController } from '../calculationJobs';
import { CalculationViewRetired } from './CalculationViewRetired';
import { type ErrorScanFrame, createScopedErrorScanner } from '../errorScanAdapter';
import { type FormulaTraceFrame, createScopedFormulaTraceQuery } from '../formulaTraceAdapter';
import type { PivotContextFrame } from '../pivotSelection';
import { type ReviewOwnerFrame, createScopedReviewReader } from '../reviewScanAdapter';
import {
  type RibbonBinding,
  type RibbonCommandHandler,
  type RibbonCommandId,
  type RibbonDispatchScope,
  dispatchRibbonCommand,
  ribbonCommand,
} from '../ribbonCommands';
import type { TextColumnContext } from '../textColumns';
import { type WatchReadFrame, createScopedWatchReader } from '../watchAdapter';
import type { WorkbookCommand, WorkbookSheet } from '../../core/contracts';
import { createFormulaTraceController } from '../formulaTrace';
import { useEffect, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react';

export interface CommandPlumbingInputs {
  activationEpoch: React.MutableRefObject<number>;
  active: boolean;
  documentId: string | undefined;
  fallbackDocumentId: string;
  functionViewMode: React.MutableRefObject<{ mode: 'normal' | 'pageLayout' | 'pageBreak'; epoch: number }>;
  navigationGeneration: React.MutableRefObject<number>;
  normalMode: React.MutableRefObject<boolean>;
  onCommandBatch: ((commands: WorkbookCommand[]) => Promise<void>) | undefined;
  ownerIdentity: string | undefined;
  sheet: WorkbookSheet;
  viewActive: React.MutableRefObject<boolean>;
}

export function useCommandPlumbing({
  activationEpoch,
  active,
  documentId,
  fallbackDocumentId,
  functionViewMode,
  navigationGeneration,
  normalMode,
  onCommandBatch,
  ownerIdentity,
  sheet,
  viewActive,
}: CommandPlumbingInputs) {
  const [jump, setJump] = useState(''),
    [jumpError, setJumpError] = useState('');
  const [gridError, setGridError] = useState('');
  const [viewNotice, setViewNotice] = useState('');
  const calculationRun = useRef<{ key: string } | null>(null);
  const [calculationBusy, setCalculationBusy] = useState(false);
  const [calculationNotice, setCalculationNotice] = useState<{
    key: string;
    phase: 'running' | 'completed' | 'failed';
    text: string;
  } | null>(null);
  const calculationMounted = useRef(true);
  useEffect(() => {
    calculationMounted.current = true;
    return () => {
      calculationMounted.current = false;
    };
  }, []);
  // Accepted head advancement is expected after this command. Scope the outcome
  // to its original runtime/worksheet/activation, never a replacement owner.
  const calculationOwnerAvailable = typeof ownerIdentity === 'string' && ownerIdentity.trim().length > 0;
  const calculationKey = JSON.stringify([
    documentId ?? fallbackDocumentId,
    ownerIdentity ?? null,
    sheet?.name ?? '',
    activationEpoch.current,
    functionViewMode.current.epoch,
    navigationGeneration.current,
  ]);
  const calculationLive = useRef({ key: calculationKey, active, normal: normalMode.current });
  calculationLive.current = { key: calculationKey, active, normal: normalMode.current };
  const calculationFrame = useRef<CalculationFrame | null>(null);
  const calculationController = useMemo(() => createCalculationJobController(() => calculationFrame.current), []);
  const calculationJob = useSyncExternalStore(
    calculationController.subscribe,
    calculationController.getSnapshot,
    calculationController.getSnapshot
  );
  const calculationDisposal = useRef<ReturnType<typeof setTimeout> | null>(null);
  useEffect(() => {
    if (calculationDisposal.current !== null) clearTimeout(calculationDisposal.current);
    return () => {
      calculationDisposal.current = setTimeout(() => calculationController.dispose(), 0);
    };
  }, [calculationController]);
  const commandBinding = useRef<RibbonBinding | null>(null);
  const commandAvailability = useRef<ReturnType<RibbonDispatchScope['availability']>>({
    profile: 'xlsx',
    active: false,
    canEdit: false,
    busy: false,
  });
  const [textColumnContext, setTextColumnContext] = useState<TextColumnContext | null>(null);
  const errorScanFrame = useRef<ErrorScanFrame | null>(null);
  const errorScanner = useMemo(() => createScopedErrorScanner(() => errorScanFrame.current), []);
  const watchReadFrame = useRef<WatchReadFrame | null>(null);
  const watchReader = useMemo(() => createScopedWatchReader(() => watchReadFrame.current), []);
  const reviewReadFrame = useRef<ReviewOwnerFrame | null>(null);
  const reviewReader = useMemo(() => createScopedReviewReader(() => reviewReadFrame.current), []);
  const traceFrame = useRef<FormulaTraceFrame | null>(null);
  const traceController = useMemo(
    () => createFormulaTraceController(createScopedFormulaTraceQuery(() => traceFrame.current)),
    []
  );
  const [traceRoot, setTraceRoot] = useState<{ sheet: string; cell: string } | null>(null);
  const traceDisposal = useRef<ReturnType<typeof setTimeout> | null>(null);
  useEffect(() => {
    if (traceDisposal.current !== null) clearTimeout(traceDisposal.current);
    return () => {
      traceController.cancel();
      traceDisposal.current = setTimeout(() => traceController.dispose(), 0);
    };
  }, [traceController]);
  useLayoutEffect(() => {
    if (traceFrame.current) traceController.setScope(traceFrame.current.scope);
    else traceController.cancel();
  });
  const pivotFrame = useRef<PivotContextFrame | null>(null),
    pivotFlight = useRef(false);
  const [pivotRunning, setPivotRunning] = useState(false),
    [pivotNotice, setPivotNotice] = useState<{ key: string; text: string } | null>(null);
  const strictBatch = useRef(onCommandBatch);
  strictBatch.current = onCommandBatch;
  const commandScope = useMemo<RibbonDispatchScope>(
    () => ({
      captureBinding: () => {
        if (!commandBinding.current) throw Error('No worksheet is active.');
        return commandBinding.current;
      },
      isCurrent: (binding) => {
        const current = commandBinding.current;
        return (
          !!current &&
          viewActive.current &&
          Object.keys(binding).every(
            (key) => binding[key as keyof RibbonBinding] === current[key as keyof RibbonBinding]
          )
        );
      },
      availability: (id) =>
        calculationFrame.current && ['formulas.calculate-now', 'formulas.calculate-sheet'].includes(id)
          ? { ...commandAvailability.current, canEdit: !calculationFrame.current.scope.readOnly }
          : !strictBatch.current && ribbonCommand(id).kind === 'workbook'
            ? { ...commandAvailability.current, capability: { state: 'disabled', reason: COMMAND_BATCH_UNAVAILABLE } }
            : commandAvailability.current,
    }),
    []
  );
  const onCommandResult = (_id: RibbonCommandId, result: Awaited<ReturnType<typeof dispatchRibbonCommand>>): void => {
    // Ribbon's selection binding intentionally lacks runtime owner identity.
    // Preserve a refused result without showing its obsolete error in a new owner.
    if (
      result.outcome === 'failed' &&
      (result.error instanceof CalculationJobRetired ||
        (_id === 'formulas.calculate-now' && result.error instanceof CalculationViewRetired))
    )
      return;
    if (result.outcome === 'failed' && result.current)
      setGridError(
        String(result.error instanceof Error ? result.error.message : result.error)
          .replace(/^Error: /, '')
      );
    else if (result.outcome === 'refused' && viewActive.current) setGridError(result.reason);
  };
  const invoke = (id: RibbonCommandId, handler: RibbonCommandHandler): void => {
    void dispatchRibbonCommand(id, commandScope, handler)
      .then((result) => onCommandResult(id, result))
      .catch((error) => {
        if (viewActive.current) setGridError(String(error));
      });
  };
  return {
    jump,
    setJump,
    jumpError,
    setJumpError,
    gridError,
    setGridError,
    viewNotice,
    setViewNotice,
    calculationRun,
    calculationBusy,
    setCalculationBusy,
    calculationNotice,
    setCalculationNotice,
    calculationMounted,
    calculationOwnerAvailable,
    calculationKey,
    calculationLive,
    calculationFrame,
    calculationController,
    calculationJob,
    commandBinding,
    commandAvailability,
    textColumnContext,
    setTextColumnContext,
    errorScanFrame,
    errorScanner,
    watchReadFrame,
    watchReader,
    reviewReadFrame,
    reviewReader,
    traceFrame,
    traceController,
    traceRoot,
    setTraceRoot,
    pivotFrame,
    pivotFlight,
    pivotRunning,
    setPivotRunning,
    pivotNotice,
    setPivotNotice,
    commandScope,
    onCommandResult,
    invoke,
  } as const;
}
