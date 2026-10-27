import { COMMAND_BATCH_UNAVAILABLE } from './commandBatch';
import {
  type FilterAction,
  type FilterContext,
  acceptFilterCommand,
  acceptedFilterCommand,
  acceptedFilterReason,
  filterDefinitionKey,
  validateFilterValuePage,
} from '../filterSelection';
import type { FormulaDraftHandle } from '../namedFormulaDraft';
import type { GridFilterBounds, GridFilterValues, GridSortBounds } from './gridTypes';
import { type HomeBinding, type HomeFrame, captureHomeBinding, homeBindingMatches } from '../homeCommands';
import type {
  WorkbookCommand,
  WorkbookEdit,
  WorkbookFilterDefinition,
  WorkbookFilterValuePage,
  WorkbookSheet,
} from '../../core/contracts';
import { colName } from '../sheetModel';

export interface FilterActionsInputs {
  busy: boolean;
  currentFilterPage: WorkbookFilterValuePage | null;
  distinctFilterValues: { key: string; value: string | number | boolean | null }[];
  filterApplying: boolean;
  filterBounds: GridFilterBounds;
  filterChoices: React.RefObject<HTMLSelectElement>;
  filterColumn: string;
  filterFormGeneration: number;
  filterGeneration: React.MutableRefObject<number>;
  filterMounted: React.MutableRefObject<boolean>;
  filterPageScope: string;
  filterPhysical: React.MutableRefObject<boolean>;
  filterReadPending: boolean;
  filterRenderScope: string;
  filterRuntimeKey: string;
  filterScope: React.MutableRefObject<string>;
  filterSelected: Map<string, { key: string; value: string | number | boolean | null; count: number }>;
  filterText: string;
  filterValues: GridFilterValues;
  filterWritePhysical: React.MutableRefObject<boolean>;
  formulaDraftHandle: React.MutableRefObject<FormulaDraftHandle | null>;
  homeFrameRef: React.MutableRefObject<HomeFrame | null>;
  homeRenderFrame: HomeFrame;
  homeSameTarget: (binding: HomeBinding) => boolean;
  literal: boolean | undefined;
  liveFilterDefinition: React.MutableRefObject<WorkbookFilterDefinition | undefined>;
  newSheetName: string;
  onCommandBatch: ((commands: WorkbookCommand[]) => Promise<void>) | undefined;
  pendingSheetSelection: React.MutableRefObject<string | null>;
  readOnly: boolean;
  resetFilter: () => void;
  send: (commands: WorkbookCommand[], allowSpecial?: boolean) => void;
  setFilterApplying: React.Dispatch<React.SetStateAction<boolean>>;
  setFilterBusy: React.Dispatch<React.SetStateAction<boolean>>;
  setFilterNotice: React.Dispatch<React.SetStateAction<{ key: string; text: string } | null>>;
  setFilterOffset: React.Dispatch<React.SetStateAction<number>>;
  setFilterPage: React.Dispatch<React.SetStateAction<WorkbookFilterValuePage | null>>;
  setFilterPageScope: React.Dispatch<React.SetStateAction<string>>;
  setFilterReadPending: React.Dispatch<React.SetStateAction<boolean>>;
  setGridError: React.Dispatch<React.SetStateAction<string>>;
  setNewSheetName: React.Dispatch<React.SetStateAction<string>>;
  setSheetAction: React.Dispatch<React.SetStateAction<'addSheet' | 'renameSheet' | 'duplicateSheet' | null>>;
  sheet: WorkbookSheet;
  sheetAction: 'addSheet' | 'renameSheet' | 'duplicateSheet' | null;
  sheets: WorkbookSheet[];
  sortBounds: GridSortBounds;
}

export function buildFilterActions({
  busy,
  currentFilterPage,
  distinctFilterValues,
  filterApplying,
  filterBounds,
  filterChoices,
  filterColumn,
  filterFormGeneration,
  filterGeneration,
  filterMounted,
  filterPageScope,
  filterPhysical,
  filterReadPending,
  filterRenderScope,
  filterRuntimeKey,
  filterScope,
  filterSelected,
  filterText,
  filterValues,
  filterWritePhysical,
  formulaDraftHandle,
  homeFrameRef,
  homeRenderFrame,
  homeSameTarget,
  literal,
  liveFilterDefinition,
  newSheetName,
  onCommandBatch,
  pendingSheetSelection,
  readOnly,
  resetFilter,
  send,
  setFilterApplying,
  setFilterBusy,
  setFilterNotice,
  setFilterOffset,
  setFilterPage,
  setFilterPageScope,
  setFilterReadPending,
  setGridError,
  setNewSheetName,
  setSheetAction,
  sheet,
  sheetAction,
  sheets,
  sortBounds,
}: FilterActionsInputs) {
  const filterDisabledReason = (write = true, requireRectangle = true): string | undefined => {
    const frame = homeFrameRef.current;
    if (!frame?.active) return 'Activate this workbook before filtering.';
    if (frame.profile !== 'xlsx') return 'Filtering requires an XLSX workbook.';
    if (frame.binding.mode !== 'normal') return 'Return to Normal view to filter rows.';
    if (requireRectangle && frame.binding.specialSelection) return 'Select a rectangular range before filtering.';
    if (frame.busy) return 'Wait for the workbook operation to finish.';
    if (!frame.binding.ownerIdentity) return 'Filtering is unavailable in this editor.';
    if (write && !frame.canEdit) return 'This workbook is read-only.';
    if (write && !frame.acceptBatch) return COMMAND_BATCH_UNAVAILABLE;
  };
  const filterActionReason = (action: FilterAction): string | undefined =>
    filterDisabledReason(true, false) ??
    (filterApplying ? 'Wait for the current filter command to finish.' : undefined) ??
    acceptedFilterReason(sheet.filter, action);
  const captureFilterContext = (requireRectangle = true): FilterContext => {
    if (!homeFrameRef.current?.active || !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding))
      throw Error('The workbook or filter selection changed. Load the values again.');
    const binding = captureHomeBinding(homeRenderFrame.binding),
      definition = sheet.filter ? structuredClone(sheet.filter) : undefined,
      definitionKey = filterDefinitionKey(definition),
      batch = onCommandBatch!;
    const isCurrent = () =>
      !!homeFrameRef.current?.active &&
      homeBindingMatches(binding, homeFrameRef.current.binding) &&
      filterDefinitionKey(liveFilterDefinition.current) === definitionKey;
    return {
      binding,
      definition,
      isCurrent,
      isRuntimeCurrent: () => homeSameTarget(binding),
      disabledReason: () => filterDisabledReason(true, requireRectangle),
      accept: async (commands) => {
        if (!isCurrent()) throw Error('The workbook or filter definition changed. Try again.');
        const reason = filterDisabledReason(true, requireRectangle);
        if (reason) throw Error(reason);
        if (filterWritePhysical.current) throw Error('Wait for the current filter command to finish.');
        filterWritePhysical.current = true;
        setFilterApplying(true);
        try {
          const pending = batch(commands);
          if (!pending || typeof pending.then !== 'function')
            throw Error('Filtering requires an acknowledged workbook command.');
          await pending;
        } finally {
          filterWritePhysical.current = false;
          if (filterMounted.current) setFilterApplying(false);
        }
      },
    };
  };
  const invokeFilter = async (
    builder: (context: FilterContext) => WorkbookCommand,
    notice: string,
    requireRectangle = true
  ): Promise<void> => {
    if (
      !homeFrameRef.current?.active ||
      !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding) ||
      filterScope.current !== filterRenderScope
    )
      return;
    let context: FilterContext | undefined;
    try {
      context = captureFilterContext(requireRectangle);
      const focused = document.activeElement as HTMLElement | null,
        handle = formulaDraftHandle.current;
      if (
        (focused?.closest('.sheet-grid') && focused.getAttribute('data-editing') === 'true') ||
        (handle?.ownsFocus() && handle.capture().composing)
      ) {
        setGridError('Finish or cancel the current input before filtering.');
        return;
      }
      if (handle?.ownsFocus()) handle.park();
      const result = await acceptFilterCommand(context, builder(context));
      if (result.runtimeCurrent) {
        setFilterNotice({ key: filterRuntimeKey, text: notice });
        setGridError('');
        resetFilter();
      }
    } catch (error) {
      if (
        context?.isCurrent() ||
        (!context &&
          homeFrameRef.current?.active &&
          homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding))
      )
        setGridError(error instanceof Error ? error.message : String(error));
    }
  };
  const applyAcceptedFilter = (action: FilterAction): Promise<void> =>
    invokeFilter(
      (context) => acceptedFilterCommand(context.binding.sheet, context.definition, action),
      action === 'reapply' ? 'Filter reapplied.' : action === 'clear' ? 'Filter cleared.' : 'Filter removed.',
      false
    );
  const applyFilterChoices = (): Promise<void> =>
    invokeFilter((context) => {
      if (!sortBounds || !filterChoices.current) throw Error('Select a rectangular filter range.');
      const keys = new Set(Array.from(filterChoices.current.selectedOptions, (option) => option.value));
      if (!keys.size) throw Error('Select at least one value to filter.');
      return {
        sheet: context.binding.sheet,
        operation: 'filterRows',
        range: context.binding.selection,
        column: filterColumn,
        values: distinctFilterValues.filter((item) => keys.has(item.key)).map((item) => item.value),
      };
    }, 'Filter applied.');
  const loadFilterValues = async (offset = 0): Promise<void> => {
    if (
      !homeFrameRef.current?.active ||
      !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding) ||
      filterDisabledReason(false) ||
      !filterBounds ||
      filterBounds.rows < 2 ||
      !filterValues ||
      filterPhysical.current ||
      filterGeneration.current !== filterFormGeneration ||
      filterScope.current !== filterRenderScope
    )
      return;
    const context = captureFilterContext(),
      reader = filterValues,
      queryText = filterText,
      column = filterColumn,
      generation = ++filterGeneration.current,
      scope = filterRenderScope;
    const dataRange = `${colName(filterBounds.first.col)}${filterBounds.first.row + 2}:${colName(filterBounds.last.col)}${filterBounds.last.row + 1}`;
    filterPhysical.current = true;
    setFilterReadPending(true);
    setFilterBusy(true);
    try {
      const result = await reader(context.binding.sheet, dataRange, column, queryText, offset);
      if (generation !== filterGeneration.current || scope !== filterScope.current || !context.isCurrent()) return;
      const page = validateFilterValuePage(result, offset);
      setFilterPage(page);
      setFilterPageScope(scope);
      setFilterOffset(offset);
      setFilterBusy(false);
      setGridError('');
    } catch (error) {
      if (generation === filterGeneration.current && scope === filterScope.current && context.isCurrent()) {
        setFilterBusy(false);
        setGridError(error instanceof Error ? error.message : String(error));
      }
    } finally {
      filterPhysical.current = false;
      if (filterMounted.current) setFilterReadPending(false);
    }
  };
  const protectedFilter = currentFilterPage
    ? Object.values(currentFilterPage.protected).some((count) => count > 0)
    : false;
  const applyPagedFilter = (): Promise<void> => {
    if (filterGeneration.current !== filterFormGeneration) return Promise.resolve();
    return invokeFilter((context) => {
      if (
        !filterBounds ||
        filterReadPending ||
        protectedFilter ||
        !currentFilterPage ||
        filterPageScope !== filterRenderScope
      )
        throw Error('Load current values for this range before applying a filter.');
      if (!filterSelected.size) throw Error('Select at least one value to filter.');
      return {
        sheet: context.binding.sheet,
        operation: 'filterRows',
        range: context.binding.selection,
        column: filterColumn,
        values: [...filterSelected.values()].map((item) => item.value),
      };
    }, 'Filter applied.');
  };
  const submitSheetAction = (): void => {
    if (!sheetAction || readOnly || busy || literal) return;
    const name = newSheetName.trim();
    if (!name || name.length > 31 || /[\\/:?*\[\]]/.test(name) || name.startsWith("'") || name.endsWith("'")) {
      setGridError('Use a worksheet name of 1–31 characters without \\, /, :, ?, *, [ or ].');
      return;
    }
    if (
      sheets.some(
        (item) =>
          item.name.toLocaleLowerCase() === name.toLocaleLowerCase() &&
          (sheetAction !== 'renameSheet' || item.name !== sheet.name)
      )
    ) {
      setGridError('A worksheet already has that name.');
      return;
    }
    const command: WorkbookEdit =
      sheetAction === 'addSheet'
        ? { sheet: name, operation: 'addSheet' }
        : { sheet: sheet.name, operation: sheetAction, newName: name };
    pendingSheetSelection.current = name;
    send([command]);
    setSheetAction(null);
    setNewSheetName('');
    setGridError('');
  };
  return {
    filterDisabledReason,
    filterActionReason,
    invokeFilter,
    applyAcceptedFilter,
    applyFilterChoices,
    loadFilterValues,
    protectedFilter,
    applyPagedFilter,
    submitSheetAction,
  } as const;
}
