import { type CalculationFrame, type WorkbookCalculationControls, calculationFrameKey } from '../calculationJobs';
import { type FormulaPreviewControls, type FormulaPreviewFrame, formulaPreviewScopeKey } from '../formulaDraftPreview';
import type { FormulaTraceFrame } from '../formulaTraceAdapter';
import type { GridPaintStyleRef, GridRanges, GridSetNamePicker } from './gridTypes';
import type { HomeBinding, HomeFrame } from '../homeCommands';
import type { ReviewOwnerFrame } from '../reviewScanAdapter';
import type { RibbonAvailability, RibbonBinding } from '../ribbonCommands';
import type { WorkbookCommand, WorkbookFilterValuePage, WorkbookSheet } from '../../core/contracts';
import { filterDefinitionKey } from '../filterSelection';
import { useEffect, useLayoutEffect } from 'react';

export interface RenderFramesInputs {
  acceptedCommandState: string | undefined;
  activationEpoch: React.MutableRefObject<number>;
  active: boolean;
  activeRange: () => string;
  busy: boolean;
  calculationControls: WorkbookCalculationControls | undefined;
  calculationFrame: React.MutableRefObject<CalculationFrame | null>;
  canEditRef: React.MutableRefObject<boolean>;
  commandAvailability: React.MutableRefObject<Omit<RibbonAvailability, 'hasHandler'>>;
  commandBinding: React.MutableRefObject<RibbonBinding | null>;
  documentId: string | undefined;
  fallbackDocumentId: string;
  filterGeneration: React.MutableRefObject<number>;
  filterPage: WorkbookFilterValuePage | null;
  filterPageScope: string;
  filterScope: React.MutableRefObject<string>;
  formulaPreviewControls: FormulaPreviewControls | undefined;
  formulaPreviewFrame: React.MutableRefObject<FormulaPreviewFrame | null>;
  functionViewMode: React.MutableRefObject<{ mode: 'normal' | 'pageLayout' | 'pageBreak'; epoch: number }>;
  homeFrameRef: React.MutableRefObject<HomeFrame | null>;
  homeMetadata: React.MutableRefObject<{ value: WorkbookSheet[]; epoch: number }>;
  homeSpecial: React.MutableRefObject<{ value: Set<string> | null; epoch: number }>;
  literal: boolean | undefined;
  namedDraftLive: React.MutableRefObject<{ key: string; editable: boolean }>;
  namedRanges: { name: string; formula: string; sheet?: string; scopeInvalid?: true }[];
  namePickerSerial: React.MutableRefObject<number>;
  navigationGeneration: React.MutableRefObject<number>;
  normalMode: React.MutableRefObject<boolean>;
  onBatch: ((commands: WorkbookCommand[]) => void) | undefined;
  onCommandBatch: ((commands: WorkbookCommand[]) => Promise<void>) | undefined;
  ownerIdentity: string | undefined;
  paintHomeBinding: React.MutableRefObject<HomeBinding | null>;
  paintStyleRef: GridPaintStyleRef;
  passiveNotice: string | undefined;
  ranges: GridRanges;
  readOnly: boolean;
  resetFilter: () => void;
  resetFind: () => void;
  reviewReadFrame: React.MutableRefObject<ReviewOwnerFrame | null>;
  selected: string;
  setGridError: React.Dispatch<React.SetStateAction<string>>;
  setNamePicker: GridSetNamePicker;
  setPaintStyleId: React.Dispatch<React.SetStateAction<number | null>>;
  sheet: WorkbookSheet;
  specialSelection: Set<string> | null;
  storageKey: string | undefined;
  traceFrame: React.MutableRefObject<FormulaTraceFrame | null>;
  viewMode: 'normal' | 'pageLayout' | 'pageBreak';
}

export function useRenderFrames({
  acceptedCommandState,
  activationEpoch,
  active,
  activeRange,
  busy,
  calculationControls,
  calculationFrame,
  canEditRef,
  commandAvailability,
  commandBinding,
  documentId,
  fallbackDocumentId,
  filterGeneration,
  filterPage,
  filterPageScope,
  filterScope,
  formulaPreviewControls,
  formulaPreviewFrame,
  functionViewMode,
  homeFrameRef,
  homeMetadata,
  homeSpecial,
  literal,
  namedDraftLive,
  namedRanges,
  namePickerSerial,
  navigationGeneration,
  normalMode,
  onBatch,
  onCommandBatch,
  ownerIdentity,
  paintHomeBinding,
  paintStyleRef,
  passiveNotice,
  ranges,
  readOnly,
  resetFilter,
  resetFind,
  reviewReadFrame,
  selected,
  setGridError,
  setNamePicker,
  setPaintStyleId,
  sheet,
  specialSelection,
  storageKey,
  traceFrame,
  viewMode,
}: RenderFramesInputs) {
  commandBinding.current = {
    documentId: documentId ?? fallbackDocumentId,
    state: storageKey ?? sheet?.calculationRevision ?? 'local',
    sheet: sheet?.name ?? '',
    selection: activeRange(),
    activationEpoch: activationEpoch.current,
  };
  const homeRenderFrame: HomeFrame = {
    binding: {
      ...commandBinding.current,
      ownerIdentity: ownerIdentity ?? null,
      state: acceptedCommandState ?? storageKey ?? sheet?.calculationRevision ?? `local:${homeMetadata.current.epoch}`,
      selectedCell: selected,
      navigationGeneration: navigationGeneration.current,
      modeGeneration: functionViewMode.current.epoch,
      specialSelectionEpoch: homeSpecial.current.epoch,
      mode: viewMode,
      specialSelection: specialSelection ? [...specialSelection] : null,
    },
    profile: literal ? 'delimited' : passiveNotice ? 'passive' : 'xlsx',
    active,
    canEdit: !readOnly,
    busy: !!busy,
    acceptBatch: onCommandBatch,
  };
  homeFrameRef.current = homeRenderFrame;
  const filterRenderScope = JSON.stringify([
    homeRenderFrame.binding,
    homeRenderFrame.profile,
    homeRenderFrame.active,
    homeRenderFrame.canEdit,
    filterDefinitionKey(sheet?.filter),
  ]);
  filterScope.current = filterRenderScope;
  const currentFilterPage = filterPageScope === filterRenderScope ? filterPage : null;
  const filterFormGeneration = filterGeneration.current;
  const filterRuntimeKey = JSON.stringify([
    homeRenderFrame.binding.documentId,
    homeRenderFrame.binding.ownerIdentity,
    homeRenderFrame.binding.sheet,
    homeRenderFrame.binding.activationEpoch,
    homeRenderFrame.binding.modeGeneration,
    homeRenderFrame.binding.mode,
    homeRenderFrame.active,
  ]);
  useEffect(resetFilter, [filterRenderScope]);
  const homeWorkbookKey = JSON.stringify([
    homeRenderFrame.binding.documentId,
    homeRenderFrame.binding.ownerIdentity,
    homeRenderFrame.binding.state,
    homeRenderFrame.binding.activationEpoch,
  ]);
  useLayoutEffect(() => {
    resetFind();
    paintStyleRef.current = null;
    paintHomeBinding.current = null;
    setPaintStyleId(null);
  }, [homeWorkbookKey]);
  const namedDraftKey = JSON.stringify([
    commandBinding.current,
    documentId ?? fallbackDocumentId,
    ownerIdentity,
    navigationGeneration.current,
    functionViewMode.current,
    !!active,
    !!specialSelection,
    namedRanges,
  ]);
  namedDraftLive.current = {
    key: namedDraftKey,
    editable: !!active && !readOnly && !busy && !literal && !passiveNotice && !specialSelection && normalMode.current,
  };
  useLayoutEffect(() => {
    namePickerSerial.current++;
    setNamePicker(null);
  }, [namedDraftKey]);
  commandAvailability.current = {
    profile: literal ? 'delimited' : passiveNotice ? 'passive' : 'xlsx',
    active,
    canEdit: !readOnly && !!(onCommandBatch ?? onBatch),
    busy: !!busy,
  };
  calculationFrame.current = calculationControls
    ? {
        controls: calculationControls,
        scope: {
          documentId: documentId ?? fallbackDocumentId,
          ownerIdentity: ownerIdentity ?? '',
          state: commandBinding.current.state,
          sheet: sheet?.name ?? '',
          activationEpoch: activationEpoch.current,
          modeEpoch: functionViewMode.current.epoch,
          navigationGeneration: navigationGeneration.current,
          active,
          normal: normalMode.current,
          readOnly,
          busy: !!busy,
          profile: literal ? 'delimited' : passiveNotice ? 'passive' : 'xlsx',
        },
      }
    : null;
  formulaPreviewFrame.current =
    formulaPreviewControls && sheet
      ? {
          controls: formulaPreviewControls,
          scope: {
            documentId: documentId ?? fallbackDocumentId,
            ownerIdentity: ownerIdentity ?? '',
            state: commandBinding.current.state,
            sheet: sheet.name,
            cell: selected,
            navigationGeneration: navigationGeneration.current,
            activationEpoch: activationEpoch.current,
            modeEpoch: functionViewMode.current.epoch,
            active: active && !specialSelection,
            normal: normalMode.current,
            busy: !!busy,
            profile: literal ? 'delimited' : passiveNotice ? 'passive' : 'xlsx',
            receipt: formulaPreviewControls.receipt,
          },
        }
      : null;
  const formulaPreviewKey = formulaPreviewFrame.current
    ? formulaPreviewScopeKey(formulaPreviewFrame.current.scope)
    : null;
  const renderedCalculationKey = calculationFrame.current ? calculationFrameKey(calculationFrame.current) : null;
  canEditRef.current = active && !readOnly && !busy;
  useEffect(() => {
    if (ranges.error) setGridError(ranges.error);
  }, [ranges.error]);
  traceFrame.current = null;
  reviewReadFrame.current = null;
  return {
    homeRenderFrame,
    filterRenderScope,
    currentFilterPage,
    filterFormGeneration,
    filterRuntimeKey,
    homeWorkbookKey,
    namedDraftKey,
    formulaPreviewKey,
    renderedCalculationKey,
  } as const;
}
