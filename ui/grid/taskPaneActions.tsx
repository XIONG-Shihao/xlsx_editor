import { type CreateNamesBinding, type CreateNamesContext, createNamesBindingKey } from '../createNamesFromSelection';
import type { DuplicateContext } from '../duplicateSelection';
import type { FormulaDraftHandle } from '../namedFormulaDraft';
import { type FunctionBinding, type FunctionInsertionContext, functionBindingKey } from '../functionInsertion';
import type { FunctionCategory } from '../functionCatalog';
import type {
  GridActionsParkedDraft,
  GridConditionalControls3,
  GridCreateNamesParkedDraft,
  GridDuplicateParkedDraft,
  GridLiveFunctionFrame,
  GridLiveNameInventory,
  GridLocalExecutionReason,
  GridReviewParkedDraft,
  GridSetActionsIntent,
  GridSetFunctionBrowse,
  GridSortParkedDraft,
  GridTables,
  GridTextColumnsParkedDraft,
  GridValidationControls2,
} from './gridTypes';
import { type HomeFrame, homeBindingMatches } from '../homeCommands';
import type { IllustrationMode } from '../WorkbookShapes';
import type { LocalExecutionAction } from '../../core/local-action-execution.mjs';
import type { ReadRange } from '../rangeCells';
import type { RibbonBinding, RibbonDispatchScope } from '../ribbonCommands';
import type { SortContext } from '../sortSelection';
import type { SpecialSelectionKind } from '../specialSelection';
import { type TextColumnBinding, type TextColumnContext, textColumnBindingKey } from '../textColumns';
import type { TraceDirection } from '../formulaTrace';
import type { WorkbookActionsControls } from '../WorkbookActions';
import type { WorkbookCommand, WorkbookPivotDefinition, WorkbookSheet } from '../../core/contracts';
import type { WorkbookQueryControls } from '../WorkbookQueries';
import type { WorkbookReviewFrame, WorkbookReviewMode } from '../workbookReviewController';
import type { WorkbookRibbonTool, WorkbookTool } from '../SpreadsheetRibbon';
import { inRange } from './rangeAddress';
import { textColumnProtection } from '../textColumnsProtection';

export interface TaskPaneActionsInputs {
  actionsControls: WorkbookActionsControls | undefined;
  actionsIntent: { serial: number; mode: 'new' | 'list' | 'run'; action?: LocalExecutionAction };
  actionsParkedDraft: GridActionsParkedDraft;
  active: boolean;
  activeTool: WorkbookTool | null;
  busy: boolean;
  cancelCellFocus: () => void;
  captureDuplicateContext: () => DuplicateContext;
  captureSortContext: () => SortContext;
  closeTaskPane: () => void;
  commandBinding: React.MutableRefObject<RibbonBinding | null>;
  commandScope: RibbonDispatchScope;
  conditionalControls: GridConditionalControls3;
  createNamesParkedDraft: GridCreateNamesParkedDraft;
  documentId: string | undefined;
  duplicateDisabledReason: (action?: 'preview' | 'remove') => string | undefined;
  duplicateParkedDraft: GridDuplicateParkedDraft;
  fallbackDocumentId: string;
  formulaDraftHandle: React.MutableRefObject<FormulaDraftHandle | null>;
  functionViewMode: React.MutableRefObject<{ mode: 'normal' | 'pageLayout' | 'pageBreak'; epoch: number }>;
  homeFrameRef: React.MutableRefObject<HomeFrame | null>;
  homeRenderFrame: HomeFrame;
  literal: boolean | undefined;
  liveFunctionFrame: GridLiveFunctionFrame;
  liveNameInventory: GridLiveNameInventory;
  localExecutionReason: GridLocalExecutionReason;
  modeGeneration: React.MutableRefObject<number>;
  namedDraftKey: string;
  navigationGeneration: React.MutableRefObject<number>;
  normalMode: React.MutableRefObject<boolean>;
  onCommandBatch: ((commands: WorkbookCommand[]) => Promise<void>) | undefined;
  openAccessibility: (binding: RibbonBinding) => void;
  openCharts: (binding: RibbonBinding, mode: 'manage' | 'recommendations') => void;
  openIllustrations: (binding: RibbonBinding, mode: IllustrationMode) => void;
  openNotes: (binding: RibbonBinding) => void;
  openPageStyles: (binding: RibbonBinding, mode: 'themes' | 'backgrounds') => void;
  openReplacement: () => void;
  openSeries: () => void;
  openSpecial: (kind?: SpecialSelectionKind) => void;
  ownerBusy: boolean;
  ownerIdentity: string | undefined;
  passiveNotice: string | undefined;
  pivots: WorkbookPivotDefinition[];
  queriesMode: 'individual' | 'all';
  queryAllReason: 'Return to Normal view to refresh queries.' | 'This workbook is read-only.' | undefined;
  queryControls: WorkbookQueryControls | undefined;
  readOnly: boolean;
  readRange: ReadRange | undefined;
  reviewFrame: WorkbookReviewFrame | null;
  reviewParkedDraft: GridReviewParkedDraft;
  selected: string;
  setActionsIntent: GridSetActionsIntent;
  setActiveTool: React.Dispatch<React.SetStateAction<WorkbookTool | null>>;
  setCreateNamesContext: React.Dispatch<React.SetStateAction<CreateNamesContext | null>>;
  setDuplicateContext: React.Dispatch<React.SetStateAction<DuplicateContext | null>>;
  setFunctionBrowse: GridSetFunctionBrowse;
  setFunctionContext: React.Dispatch<React.SetStateAction<FunctionInsertionContext | null>>;
  setGridError: React.Dispatch<React.SetStateAction<string>>;
  setQueriesAllIntent: React.Dispatch<React.SetStateAction<number>>;
  setQueriesMode: React.Dispatch<React.SetStateAction<'individual' | 'all'>>;
  setReviewIntent: React.Dispatch<React.SetStateAction<{ serial: number; mode: WorkbookReviewMode }>>;
  setSortContext: React.Dispatch<React.SetStateAction<SortContext | null>>;
  setTextColumnContext: React.Dispatch<React.SetStateAction<TextColumnContext | null>>;
  setVisitedTools: React.Dispatch<React.SetStateAction<WorkbookTool[]>>;
  sheet: WorkbookSheet;
  sortDisabledReason: () => string | undefined;
  sortParkedDraft: GridSortParkedDraft;
  sortPending: boolean;
  specialSelection: Set<string> | null;
  specialSelectionMode: React.MutableRefObject<{ selection: Set<string> | null }>;
  startTrace: (binding: RibbonBinding, direction: TraceDirection) => Promise<void>;
  storageKey: string | undefined;
  tables: GridTables;
  textColumnsParkedDraft: GridTextColumnsParkedDraft;
  toolFocusGeneration: React.MutableRefObject<number>;
  tools: WorkbookRibbonTool[];
  validationControls: GridValidationControls2;
  viewActive: React.MutableRefObject<boolean>;
}

export function buildTaskPaneActions({
  actionsControls,
  actionsIntent,
  actionsParkedDraft,
  active,
  activeTool,
  busy,
  cancelCellFocus,
  captureDuplicateContext,
  captureSortContext,
  closeTaskPane,
  commandBinding,
  commandScope,
  conditionalControls,
  createNamesParkedDraft,
  documentId,
  duplicateDisabledReason,
  duplicateParkedDraft,
  fallbackDocumentId,
  formulaDraftHandle,
  functionViewMode,
  homeFrameRef,
  homeRenderFrame,
  literal,
  liveFunctionFrame,
  liveNameInventory,
  localExecutionReason,
  modeGeneration,
  namedDraftKey,
  navigationGeneration,
  normalMode,
  onCommandBatch,
  openAccessibility,
  openCharts,
  openIllustrations,
  openNotes,
  openPageStyles,
  openReplacement,
  openSeries,
  openSpecial,
  ownerBusy,
  ownerIdentity,
  passiveNotice,
  pivots,
  queriesMode,
  queryAllReason,
  queryControls,
  readOnly,
  readRange,
  reviewFrame,
  reviewParkedDraft,
  selected,
  setActionsIntent,
  setActiveTool,
  setCreateNamesContext,
  setDuplicateContext,
  setFunctionBrowse,
  setFunctionContext,
  setGridError,
  setQueriesAllIntent,
  setQueriesMode,
  setReviewIntent,
  setSortContext,
  setTextColumnContext,
  setVisitedTools,
  sheet,
  sortDisabledReason,
  sortParkedDraft,
  sortPending,
  specialSelection,
  specialSelectionMode,
  startTrace,
  storageKey,
  tables,
  textColumnsParkedDraft,
  toolFocusGeneration,
  tools,
  validationControls,
  viewActive,
}: TaskPaneActionsInputs) {
  const openActions = (binding: RibbonBinding, mode: 'new' | 'list'): void => {
    if (!commandScope.isCurrent(binding) || !normalMode.current || !actionsControls)
      throw Error('Activate this workbook in Normal view to manage local actions.');
    const handle = formulaDraftHandle.current;
    if (handle?.ownsFocus()) {
      const snapshot = handle.capture();
      if (snapshot.composing) throw Error('Finish the current input before opening Local Actions.');
      handle.park();
      actionsParkedDraft.current = { key: namedDraftKey, handle, snapshot };
    }
    cancelCellFocus();
    toolFocusGeneration.current++;
    setVisitedTools((previous) => (previous.includes('actions') ? previous : [...previous, 'actions']));
    setActiveTool('actions');
    setActionsIntent((previous) => ({ serial: previous.serial + 1, mode }));
  };
  const openQueries = (binding: RibbonBinding, mode: 'individual' | 'all'): void => {
    if (
      !commandScope.isCurrent(binding) ||
      !normalMode.current ||
      !queryControls ||
      (mode === 'all' && !queryControls.refreshAll)
    )
      throw Error('Activate this workbook in Normal view to use queries.');
    if (mode === 'all' && (queryAllReason || ownerBusy))
      throw Error(queryAllReason ?? 'Wait for the workbook operation to finish.');
    const focused = document.activeElement as HTMLElement | null;
    if (focused?.getAttribute('data-editing') === 'true' || formulaDraftHandle.current?.ownsFocus())
      throw Error('Finish or cancel the cell edit before opening queries.');
    if (activeTool === 'queries' && queriesMode === mode) {
      closeTaskPane();
      return;
    }
    cancelCellFocus();
    toolFocusGeneration.current++;
    setQueriesMode(mode);
    setVisitedTools((previous) => (previous.includes('queries') ? previous : [...previous, 'queries']));
    setActiveTool('queries');
    if (mode === 'all') setQueriesAllIntent((serial) => serial + 1);
  };
  const openLocalExecution = (binding: RibbonBinding, action: LocalExecutionAction): void => {
    if (!commandScope.isCurrent(binding) || localExecutionReason)
      throw Error(localExecutionReason ?? 'The original action worksheet changed.');
    const focus = document.activeElement as HTMLElement | null;
    if (focus?.getAttribute('data-editing') === 'true' || formulaDraftHandle.current?.ownsFocus())
      throw Error('Finish or cancel the cell edit before running a local action.');
    if (activeTool === 'actions' && actionsIntent.mode === 'run' && actionsIntent.action === action) {
      closeTaskPane();
      return;
    }
    cancelCellFocus();
    toolFocusGeneration.current++;
    setVisitedTools((previous) => (previous.includes('actions') ? previous : [...previous, 'actions']));
    setActiveTool('actions');
    setActionsIntent((previous) => ({ serial: previous.serial + 1, mode: 'run', action }));
  };
  const openReview = (binding: RibbonBinding, mode: WorkbookReviewMode): void => {
    if (!commandScope.isCurrent(binding) || !normalMode.current || !reviewFrame || specialSelection)
      throw Error('Activate a regular cell in Normal view to review this workbook.');
    const handle = formulaDraftHandle.current;
    if (handle?.ownsFocus()) {
      const snapshot = handle.capture();
      if (snapshot.composing) throw Error('Finish the current input before opening Review.');
      handle.park();
      reviewParkedDraft.current = { key: namedDraftKey, handle, snapshot };
    }
    cancelCellFocus();
    toolFocusGeneration.current++;
    setVisitedTools((previous) => (previous.includes('review') ? previous : [...previous, 'review']));
    setActiveTool('review');
    setReviewIntent((previous) => ({ serial: previous.serial + 1, mode }));
  };
  const toggleTool = (
    tool: WorkbookTool,
    initiatingBinding?: RibbonBinding,
    browse?: 'all' | 'recent' | FunctionCategory
  ): void => {
    if (!active || !tools.some((item) => item.id === tool)) return;
    if ((tool === 'pivot' || tool === 'pivotDesign') && initiatingBinding && !commandScope.isCurrent(initiatingBinding))
      return;
    if (tool === 'accessibility') {
      openAccessibility(initiatingBinding ?? commandScope.captureBinding());
      return;
    }
    if (tool === 'notes') {
      openNotes(initiatingBinding ?? commandScope.captureBinding());
      return;
    }
    if (tool === 'objects') {
      openCharts(initiatingBinding ?? commandScope.captureBinding(), 'manage');
      return;
    }
    if (tool === 'illustrations') {
      openIllustrations(initiatingBinding ?? commandScope.captureBinding(), 'shape');
      return;
    }
    if (tool === 'pageStyles') {
      openPageStyles(initiatingBinding ?? commandScope.captureBinding(), 'themes');
      return;
    }
    if (tool === 'replace') {
      if (initiatingBinding && !commandScope.isCurrent(initiatingBinding)) return;
      openReplacement();
      return;
    }
    if (tool === 'fill') {
      if (initiatingBinding && !commandScope.isCurrent(initiatingBinding)) return;
      openSeries();
      return;
    }
    if (tool === 'special') {
      if (initiatingBinding && !commandScope.isCurrent(initiatingBinding)) return;
      openSpecial();
      return;
    }
    if (tool === 'trace') {
      void startTrace(initiatingBinding ?? commandScope.captureBinding(), 'precedents').catch((error) => {
        if (viewActive.current) setGridError(String(error));
      });
      return;
    }
    if (tool === 'actions') {
      openActions(initiatingBinding ?? commandScope.captureBinding(), 'list');
      return;
    }
    if (tool === 'review') {
      openReview(initiatingBinding ?? commandScope.captureBinding(), 'statistics');
      return;
    }
    if (['hyperlink', 'checkbox', 'symbols'].includes(tool)) {
      const focused = document.activeElement as HTMLElement | null,
        handle = formulaDraftHandle.current;
      if ((focused?.closest('.sheet-grid') && focused.getAttribute('data-editing') === 'true') || handle?.ownsFocus())
        throw Error('Finish or cancel the cell/formula edit before opening cell insertion tools.');
    }
    if (
      tool === 'textColumns' &&
      (!homeFrameRef.current?.active ||
        !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding) ||
        !normalMode.current)
    )
      return;
    if (
      tool === 'duplicates' &&
      (!homeFrameRef.current?.active ||
        !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding) ||
        duplicateDisabledReason())
    )
      return;
    if (
      tool === 'sort' &&
      (!homeFrameRef.current?.active ||
        !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding) ||
        sortDisabledReason() ||
        sortPending)
    )
      return;
    cancelCellFocus();
    if (activeTool === tool && !browse) {
      closeTaskPane();
      return;
    }
    if (tool === 'duplicates') {
      const focused = document.activeElement as HTMLElement | null,
        handle = formulaDraftHandle.current;
      if (
        (focused?.matches('input,textarea,select,[contenteditable="true"]') &&
          !handle?.ownsFocus() &&
          !(focused.closest('.sheet-grid') && focused.getAttribute('data-editing') === 'false')) ||
        (handle?.ownsFocus() && handle.capture().composing)
      )
        throw Error('Finish or cancel the current input before comparing rows.');
      if (handle?.ownsFocus()) {
        const snapshot = handle.capture();
        handle.park();
        duplicateParkedDraft.current = { key: namedDraftKey, handle, snapshot };
      } else duplicateParkedDraft.current = null;
      setDuplicateContext(captureDuplicateContext());
    }
    if (tool === 'sort') {
      if (
        !homeFrameRef.current?.active ||
        !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding) ||
        sortDisabledReason() ||
        sortPending
      )
        return;
      const focused = document.activeElement as HTMLElement | null,
        handle = formulaDraftHandle.current;
      if (
        (focused?.closest('.sheet-grid') && focused.getAttribute('data-editing') === 'true') ||
        (handle?.ownsFocus() && handle.capture().composing)
      )
        throw Error('Finish or cancel the current input before sorting.');
      if (handle?.ownsFocus()) {
        const snapshot = handle.capture();
        handle.park();
        sortParkedDraft.current = { key: namedDraftKey, handle, snapshot };
      }
      setSortContext(captureSortContext());
    }
    if (tool === 'createNames') {
      const binding = initiatingBinding ?? commandScope.captureBinding(),
        handle = formulaDraftHandle.current,
        focused = document.activeElement as HTMLElement | null;
      if (
        !commandScope.isCurrent(binding) ||
        !readRange ||
        !onCommandBatch ||
        !ownerIdentity ||
        readOnly ||
        busy ||
        literal ||
        passiveNotice ||
        specialSelection ||
        !normalMode.current
      )
        return;
      if (
        focused?.matches('input,textarea,select,[contenteditable="true"]') &&
        !handle?.ownsFocus() &&
        !(focused.closest('.sheet-grid') && focused.getAttribute('data-editing') === 'false')
      )
        throw Error('Finish or cancel the current input before creating names.');
      if (handle?.ownsFocus()) {
        const snapshot = handle.capture();
        if (snapshot.composing) throw Error('Finish composing the current input first.');
        handle.park();
        createNamesParkedDraft.current = { key: namedDraftKey, handle, snapshot };
      } else createNamesParkedDraft.current = null;
      const navigation = navigationGeneration.current,
        selectionMode = specialSelectionMode.current,
        mode = modeGeneration.current,
        runtimeMode = functionViewMode.current,
        reader = readRange,
        batch = onCommandBatch,
        merges = [...sheet.merges];
      const captured: CreateNamesBinding = Object.freeze({
        documentId: binding.documentId,
        ownerIdentity,
        state: binding.state,
        sheet: binding.sheet,
        selection: binding.selection,
        navigationGeneration: navigation,
        activationEpoch: binding.activationEpoch,
        modeEpoch: runtimeMode.epoch,
      });
      const runtimeCurrent = (value: CreateNamesBinding): boolean => {
        const current = commandBinding.current,
          frame = liveFunctionFrame.current;
        return (
          createNamesBindingKey(value) === createNamesBindingKey(captured) &&
          viewActive.current &&
          !!current &&
          (['documentId', 'sheet', 'selection', 'activationEpoch'] as const).every(
            (key) => current[key] === binding[key]
          ) &&
          frame.ownerIdentity === captured.ownerIdentity &&
          navigationGeneration.current === navigation &&
          specialSelectionMode.current === selectionMode &&
          functionViewMode.current === runtimeMode
        );
      };
      setCreateNamesContext({
        binding: captured,
        isCurrent: (value) =>
          runtimeCurrent(value) && commandScope.isCurrent(binding) && modeGeneration.current === mode,
        isRuntimeCurrent: runtimeCurrent,
        availability: () => {
          const current = commandScope.availability('names.create-from-selection');
          return {
            active: current.active,
            readOnly: !current.canEdit,
            busy: current.busy,
            literal: current.profile !== 'xlsx',
            specialSelection: !!specialSelectionMode.current.selection,
          };
        },
        read: (range) => reader(binding.sheet, range),
        inventory: async () => structuredClone(liveNameInventory.current),
        merges,
        accept: batch,
      });
    }
    if (tool === 'textColumns') {
      const handle = formulaDraftHandle.current,
        focused = document.activeElement as HTMLElement | null;
      if (
        (focused?.matches('input,textarea,select,[contenteditable="true"]') &&
          !handle?.ownsFocus() &&
          !(focused.closest('.sheet-grid') && focused.getAttribute('data-editing') === 'false')) ||
        (handle?.ownsFocus() && handle.capture().composing)
      )
        throw Error('Finish or cancel the current input before splitting text.');
      if (handle?.ownsFocus()) {
        const snapshot = handle.capture();
        handle.park();
        textColumnsParkedDraft.current = { key: namedDraftKey, handle, snapshot };
      }
      const binding = initiatingBinding ?? commandScope.captureBinding();
      if (
        !commandScope.isCurrent(binding) ||
        !readRange ||
        !onCommandBatch ||
        !ownerIdentity ||
        !conditionalControls ||
        !validationControls
      )
        return;
      const navigation = navigationGeneration.current,
        selectionMode = specialSelectionMode.current,
        mode = modeGeneration.current,
        runtimeMode = functionViewMode.current,
        originalSheet = sheet,
        reader = readRange,
        batch = onCommandBatch,
        conditional = conditionalControls,
        validation = validationControls,
        originalTables = tables,
        originalPivots = pivots;
      const captured: TextColumnBinding = Object.freeze({
        documentId: binding.documentId,
        grantDocumentId: documentId ?? fallbackDocumentId,
        ownerIdentity,
        ownerKey: storageKey ?? 'local',
        state: binding.state,
        sheet: binding.sheet,
        selection: binding.selection,
        navigationGeneration: navigation,
        activationEpoch: binding.activationEpoch,
      });
      const runtimeCurrent = (value: TextColumnBinding): boolean => {
        const current = commandBinding.current,
          frame = liveFunctionFrame.current;
        return (
          textColumnBindingKey(value) === textColumnBindingKey(captured) &&
          viewActive.current &&
          !!current &&
          (['documentId', 'sheet', 'selection', 'activationEpoch'] as const).every(
            (key) => current[key] === binding[key]
          ) &&
          frame.grantDocumentId === captured.grantDocumentId &&
          frame.ownerIdentity === captured.ownerIdentity &&
          navigationGeneration.current === navigation &&
          specialSelectionMode.current === selectionMode &&
          functionViewMode.current === runtimeMode
        );
      };
      const calendar = originalSheet.date1904;
      setTextColumnContext({
        binding: captured,
        date1904: calendar,
        isCurrent: (value) =>
          runtimeCurrent(value) &&
          commandScope.isCurrent(binding) &&
          modeGeneration.current === mode &&
          liveFunctionFrame.current.sheet.date1904 === calendar,
        isRuntimeCurrent: runtimeCurrent,
        availability: () => {
          const current = commandScope.availability('data.text-to-columns');
          return {
            active: current.active,
            readOnly: !current.canEdit,
            busy: current.busy,
            literal: current.profile !== 'xlsx',
            specialSelection: !!specialSelectionMode.current.selection,
          };
        },
        read: (range) => reader(binding.sheet, range),
        inspect: async (source, destination) => {
          const [conditionalDefinition, validationDefinition] = await Promise.all([
            conditional.inventory(binding.sheet),
            validation.inventory(binding.sheet),
          ]);
          return textColumnProtection(
            {
              sheet: originalSheet,
              tables: originalTables,
              pivots: originalPivots,
              conditional: conditionalDefinition,
              validation: validationDefinition,
            },
            source,
            destination
          );
        },
        accept: batch,
      });
    }
    if (tool === 'functions') {
      const binding = initiatingBinding ?? commandScope.captureBinding();
      if (!commandScope.isCurrent(binding)) return;
      const target = selected,
        originalSheet = sheet,
        reader = readRange,
        batch = onCommandBatch;
      const navigation = navigationGeneration.current,
        selectionMode = specialSelectionMode.current,
        mode = modeGeneration.current,
        runtimeMode = functionViewMode.current;
      const captured: FunctionBinding = Object.freeze({
        documentId: binding.documentId,
        grantDocumentId: documentId ?? fallbackDocumentId,
        ownerKey: storageKey ?? 'local',
        ownerIdentity: ownerIdentity ?? storageKey ?? 'local',
        state: binding.state,
        sheet: binding.sheet,
        target,
        selection: binding.selection,
        navigationGeneration: navigation,
        activationEpoch: binding.activationEpoch,
      });
      const runtimeCurrent = (value: FunctionBinding): boolean => {
        const current = commandBinding.current,
          frame = liveFunctionFrame.current;
        return (
          functionBindingKey(value) === functionBindingKey(captured) &&
          viewActive.current &&
          !!current &&
          (['documentId', 'sheet', 'selection', 'activationEpoch'] as const).every(
            (key) => current[key] === binding[key]
          ) &&
          frame.grantDocumentId === captured.grantDocumentId &&
          frame.ownerIdentity === captured.ownerIdentity &&
          navigationGeneration.current === navigation &&
          specialSelectionMode.current === selectionMode &&
          functionViewMode.current === runtimeMode
        );
      };
      setFunctionBrowse((previous) => ({ mode: browse ?? 'all', serial: previous.serial + 1 }));
      setFunctionContext({
        binding: captured,
        isCurrent: (value) =>
          functionBindingKey(value) === functionBindingKey(captured) &&
          commandScope.isCurrent(binding) &&
          navigationGeneration.current === navigation &&
          specialSelectionMode.current === selectionMode &&
          modeGeneration.current === mode,
        isRuntimeCurrent: runtimeCurrent,
        currentState: () => commandBinding.current?.state ?? 'local',
        availability: () => {
          const current = commandScope.availability('formulas.insert-function');
          return {
            active: current.active,
            readOnly: !current.canEdit,
            busy: current.busy,
            literal: current.profile !== 'xlsx',
            specialSelection: !!specialSelectionMode.current.selection,
          };
        },
        readTarget: async () => {
          const cells = reader
            ? await reader(binding.sheet, target + ':' + target)
            : { [target]: originalSheet.cells[target] };
          if (
            !cells ||
            typeof cells !== 'object' ||
            Array.isArray(cells) ||
            Object.keys(cells).some((ref) => ref !== target)
          )
            throw Error('The target reply did not match the captured cell.');
          const merged = originalSheet.merges.some((range) => inRange(range, target));
          return {
            cell: cells[target],
            ...(merged ? { protectedReason: 'Choose an unmerged cell for this function picker.' } : {}),
          };
        },
        readAcceptedTarget: async (formula) => {
          if (!runtimeCurrent(captured)) throw Error('The insertion view changed.');
          const head = commandBinding.current!.state,
            frame = liveFunctionFrame.current;
          const cells = frame.reader
            ? await frame.reader(binding.sheet, target + ':' + target)
            : { [target]: frame.sheet?.cells[target] };
          if (!runtimeCurrent(captured) || commandBinding.current?.state !== head)
            throw Error('The accepted result view changed.');
          if (
            !cells ||
            typeof cells !== 'object' ||
            Array.isArray(cells) ||
            Object.keys(cells).some((ref) => ref !== target) ||
            cells[target]?.formula !== formula
          )
            throw Error('The target changed after insertion.');
          return { state: head, cell: cells[target] };
        },
        accept: batch,
      });
    }
    toolFocusGeneration.current++;
    setVisitedTools((current) => (current.includes(tool) ? current : [...current, tool]));
    setActiveTool(tool);
  };
  return { openActions, openQueries, openLocalExecution, openReview, toggleTool } as const;
}
