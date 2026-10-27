import type { FilterAction } from '../filterSelection';
import type {
  GridFilterBounds,
  GridFilterValues,
  GridInvokeFilter,
  GridSetFilterSelected,
  GridSortBounds,
  GridToggleTool,
} from './gridTypes';
import { type HomeFrame, homeBindingMatches } from '../homeCommands';
import { Icon } from '../host';
import type { RibbonCommandHandler, RibbonCommandId } from '../ribbonCommands';
import { RibbonGroup } from '../SpreadsheetRibbon';
import type { WorkbookCell, WorkbookFilterValuePage, WorkbookSheet } from '../../core/contracts';
import { filterValueLabel } from './filterValues';

export interface SortFilterCommandsInputs {
  applyAcceptedFilter: (action: FilterAction) => Promise<void>;
  applyFilterChoices: () => Promise<void>;
  applyPagedFilter: () => Promise<void>;
  currentFilterPage: WorkbookFilterValuePage | null;
  distinctFilterValues: { key: string; value: string | number | boolean | null }[];
  filterActionReason: (action: FilterAction) => string | undefined;
  filterApplying: boolean;
  filterBounds: GridFilterBounds;
  filterBusy: boolean;
  filterChoices: React.RefObject<HTMLSelectElement>;
  filterColumn: string;
  filterDisabledReason: (write?: boolean, requireRectangle?: boolean) => string | undefined;
  filterFormGeneration: number;
  filterGeneration: React.MutableRefObject<number>;
  filterNotice: { key: string; text: string } | null;
  filterOffset: number;
  filterReadPending: boolean;
  filterRenderScope: string;
  filterRuntimeKey: string;
  filterScope: React.MutableRefObject<string>;
  filterSelected: Map<string, { key: string; value: string | number | boolean | null; count: number }>;
  filterText: string;
  filterValues: GridFilterValues;
  homeFrameRef: React.MutableRefObject<HomeFrame | null>;
  homeRenderFrame: HomeFrame;
  invoke: (id: RibbonCommandId, handler: RibbonCommandHandler) => void;
  invokeFilter: GridInvokeFilter;
  literal: boolean | undefined;
  loadFilterValues: (offset?: number) => Promise<void>;
  protectedFilter: boolean;
  selectedCell: WorkbookCell;
  setFilterBusy: React.Dispatch<React.SetStateAction<boolean>>;
  setFilterOffset: React.Dispatch<React.SetStateAction<number>>;
  setFilterPage: React.Dispatch<React.SetStateAction<WorkbookFilterValuePage | null>>;
  setFilterPageScope: React.Dispatch<React.SetStateAction<string>>;
  setFilterSelected: GridSetFilterSelected;
  setFilterText: React.Dispatch<React.SetStateAction<string>>;
  setGridError: React.Dispatch<React.SetStateAction<string>>;
  setSortHeader: React.Dispatch<React.SetStateAction<boolean>>;
  setSortPrimary: React.Dispatch<React.SetStateAction<string>>;
  sheet: WorkbookSheet;
  sortBounds: GridSortBounds;
  sortColumns: string[];
  sortDisabledReason: () => string | undefined;
  sortHeader: boolean;
  sortPending: boolean;
  sortPrimary: string;
  sortSelected: (direction: 'asc' | 'desc') => Promise<void>;
  taskPaneId: string;
  toggleTool: GridToggleTool;
  words: (en: string) => string;
}

export function buildSortFilterCommands({
  applyAcceptedFilter,
  applyFilterChoices,
  applyPagedFilter,
  currentFilterPage,
  distinctFilterValues,
  filterActionReason,
  filterApplying,
  filterBounds,
  filterBusy,
  filterChoices,
  filterColumn,
  filterDisabledReason,
  filterFormGeneration,
  filterGeneration,
  filterNotice,
  filterOffset,
  filterReadPending,
  filterRenderScope,
  filterRuntimeKey,
  filterScope,
  filterSelected,
  filterText,
  filterValues,
  homeFrameRef,
  homeRenderFrame,
  invoke,
  invokeFilter,
  literal,
  loadFilterValues,
  protectedFilter,
  selectedCell,
  setFilterBusy,
  setFilterOffset,
  setFilterPage,
  setFilterPageScope,
  setFilterSelected,
  setFilterText,
  setGridError,
  setSortHeader,
  setSortPrimary,
  sheet,
  sortBounds,
  sortColumns,
  sortDisabledReason,
  sortHeader,
  sortPending,
  sortPrimary,
  sortSelected,
  taskPaneId,
  toggleTool,
  words,
}: SortFilterCommandsInputs) {
  const sortFilterCommands = !literal ? (
    <RibbonGroup name="Sort & Filter" icon="sort" collapseBelow={0}>
      <div className="sheet-ribbon-body" data-formula-draft-zone>
        <details className="sheet-ribbon-menu">
          <summary>
            <Icon name="sort" size={25} />
            <span>{words('Sort & Filter')}</span> <Icon name="chevron" size={11} />
          </summary>
          <div>
            <label>
              {words('Primary column')}
              <select
                aria-label={words('Primary sort column')}
                value={sortColumns.includes(sortPrimary) ? sortPrimary : filterColumn}
                onChange={(event) => setSortPrimary(event.target.value)}
              >
                {sortColumns.map((name) => (
                  <option key={name} value={name}>
                    {name}
                  </option>
                ))}
              </select>
            </label>
            <label>
              <input type="checkbox" checked={sortHeader} onChange={(event) => setSortHeader(event.target.checked)} />
              {words('Header row')}
            </label>
            <button
              data-workbook-tool="sort"
              data-ribbon-command="data.sort"
              aria-controls={taskPaneId}
              disabled={!!sortDisabledReason() || sortPending}
              title={sortDisabledReason()}
              onClick={() => {
                if (
                  !homeFrameRef.current?.active ||
                  !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding)
                )
                  return;
                invoke('data.sort', () => toggleTool('sort'));
              }}
            >
              {words('Sort…')}
            </button>
            <button
              disabled={!!sortDisabledReason() || sortPending}
              title={sortDisabledReason()}
              data-ribbon-command="data.sort"
              onClick={() => void sortSelected('asc')}
            >
              {words('Sort ascending')}
            </button>
            <button
              disabled={!!sortDisabledReason() || sortPending}
              title={sortDisabledReason()}
              data-ribbon-command="data.sort"
              onClick={() => void sortSelected('desc')}
            >
              {words('Sort descending')}
            </button>
            {!filterValues && (
              <>
                <button
                  disabled={!!filterDisabledReason() || filterApplying || !sortBounds || sortBounds.rows < 2}
                  title={filterDisabledReason()}
                  data-ribbon-command="data.filter"
                  onClick={() =>
                    void invokeFilter(
                      (context) => ({
                        sheet: context.binding.sheet,
                        operation: 'filterRows',
                        range: context.binding.selection,
                        column: filterColumn,
                        values: [selectedCell?.value ?? null],
                      }),
                      'Filter applied.'
                    )
                  }
                >
                  {words('Filter selected value')}
                </button>
                <label>
                  {words('Distinct values')}
                  <select
                    key={filterRenderScope}
                    ref={filterChoices}
                    aria-label={words('Distinct filter values')}
                    multiple
                    size={Math.min(5, Math.max(2, distinctFilterValues.length))}
                  >
                    {distinctFilterValues.slice(0, 50).map((item) => (
                      <option key={item.key} value={item.key}>
                        {String(item.value ?? '(blank)')}
                      </option>
                    ))}
                  </select>
                </label>
                <button
                  disabled={
                    !!filterDisabledReason() || filterApplying || !sortBounds || distinctFilterValues.length > 50
                  }
                  title={filterDisabledReason()}
                  data-ribbon-command="data.filter"
                  onClick={() => void applyFilterChoices()}
                >
                  {words('Apply selected values')}
                </button>
                {distinctFilterValues.length > 50 && (
                  <small>{words('More than 50 values; narrow the selected range.')}</small>
                )}
              </>
            )}
            {sheet.filter && (
              <p aria-label={words('Current filter')} data-filter-revision={sheet.filter.revision ?? undefined}>
                {words('Filter')}: {sheet.name}!{sheet.filter.range}
                {sheet.filter.column && (
                  <>
                    {' '}
                    · {words('Column')} {sheet.filter.column}
                  </>
                )}
                {sheet.filter.owned &&
                  Array.isArray(sheet.filter.values) &&
                  sheet.filter.values.every(
                    (value) =>
                      value === null ||
                      typeof value === 'string' ||
                      typeof value === 'boolean' ||
                      (typeof value === 'number' && Number.isFinite(value))
                  ) && (
                    <>
                      {' '}
                      · {sheet.filter.values.length} {words('selected values')}:{' '}
                      {sheet.filter.values
                        .slice(0, 3)
                        .map((value) => filterValueLabel(value, words))
                        .join(', ')}
                      {sheet.filter.values.length > 3 ? '…' : ''}
                    </>
                  )}
              </p>
            )}
            <button
              disabled={!!filterActionReason('reapply')}
              title={filterActionReason('reapply')}
              data-ribbon-command="data.filter"
              data-filter-action="reapply"
              onClick={() => void applyAcceptedFilter('reapply')}
            >
              {words('Reapply filter')}
            </button>
            <button
              disabled={!!filterActionReason('clear')}
              title={filterActionReason('clear')}
              data-ribbon-command="data.filter"
              data-filter-action="clear"
              onClick={() => void applyAcceptedFilter('clear')}
            >
              {words('Clear filter')}
            </button>
            <button
              disabled={!!filterActionReason('remove')}
              title={filterActionReason('remove')}
              data-ribbon-command="data.filter"
              data-filter-action="remove"
              onClick={() => void applyAcceptedFilter('remove')}
            >
              {words('Remove filter')}
            </button>
            {filterNotice?.key === filterRuntimeKey && (
              <p role="status" aria-label={words('Filter status')}>
                {words(filterNotice.text)}
              </p>
            )}
            {filterValues && (
              <div className="sheet-filter-values" aria-label={words('Paged filter values')}>
                <small>
                  {words('First row is the header. Filter column:')} {filterColumn}
                </small>
                <input
                  aria-label={words('Search filter values')}
                  value={filterText}
                  onChange={(event) => {
                    if (
                      !homeFrameRef.current?.active ||
                      !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding) ||
                      filterGeneration.current !== filterFormGeneration
                    )
                      return;
                    filterGeneration.current++;
                    setFilterBusy(false);
                    setFilterPage(null);
                    setFilterPageScope('');
                    setFilterOffset(0);
                    setFilterText(event.target.value);
                  }}
                  onKeyDown={(event) => {
                    if (event.key === 'Enter') {
                      event.preventDefault();
                      void loadFilterValues();
                    }
                  }}
                />
                <button
                  disabled={
                    !!filterDisabledReason(false) || filterReadPending || !filterBounds || filterBounds.rows < 2
                  }
                  title={filterDisabledReason(false)}
                  onClick={(event) => {
                    event.stopPropagation();
                    void loadFilterValues();
                  }}
                >
                  {words('Load filter values')}
                </button>
                {filterBusy && (
                  <button
                    onClick={(event) => {
                      event.stopPropagation();
                      if (
                        filterScope.current !== filterRenderScope ||
                        filterGeneration.current !== filterFormGeneration
                      )
                        return;
                      filterGeneration.current++;
                      setFilterBusy(false);
                    }}
                  >
                    {words('Cancel value loading')}
                  </button>
                )}
                {currentFilterPage && (
                  <>
                    <span role="status">
                      {currentFilterPage.values.length ? filterOffset + 1 : 0}–
                      {filterOffset + currentFilterPage.values.length} / {currentFilterPage.total} {words('values')} ·{' '}
                      {filterSelected.size} {words('selected')}
                    </span>
                    {protectedFilter && (
                      <small role="status">
                        {words('Filtering is protected for this range:')} {currentFilterPage.protected.errors}{' '}
                        {words('errors')}, {currentFilterPage.protected.unresolved} {words('unresolved formulas')},{' '}
                        {currentFilterPage.protected.special} {words('special formulas')},{' '}
                        {currentFilterPage.protected.unsafeNumbers} {words('values with stored precision')}
                      </small>
                    )}
                    <div>
                      {currentFilterPage.values.map((item) => (
                        <label key={item.key}>
                          <input
                            type="checkbox"
                            aria-label={words('Filter value') + ' ' + filterValueLabel(item.value, words)}
                            data-filter-key={item.key}
                            checked={filterSelected.has(item.key)}
                            disabled={filterReadPending || protectedFilter}
                            onChange={(event) => {
                              if (
                                filterScope.current !== filterRenderScope ||
                                filterGeneration.current !== filterFormGeneration ||
                                !homeFrameRef.current?.active ||
                                !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding)
                              )
                                return;
                              const checked = event.target.checked;
                              if (checked && filterSelected.size >= 1000) {
                                setGridError('A filter supports up to 1,000 selected values.');
                                return;
                              }
                              setFilterSelected((prior) => {
                                const next = new Map(prior);
                                if (checked) next.set(item.key, item);
                                else next.delete(item.key);
                                return next;
                              });
                            }}
                          />
                          {filterValueLabel(item.value, words)} ({item.count})
                        </label>
                      ))}
                    </div>
                    <button
                      disabled={filterReadPending || filterOffset === 0}
                      onClick={(event) => {
                        event.stopPropagation();
                        void loadFilterValues(Math.max(0, filterOffset - 100));
                      }}
                    >
                      {words('Previous values')}
                    </button>
                    <button
                      disabled={filterReadPending || currentFilterPage.nextOffset === null}
                      onClick={(event) => {
                        event.stopPropagation();
                        if (currentFilterPage.nextOffset !== null) void loadFilterValues(currentFilterPage.nextOffset);
                      }}
                    >
                      {words('More values')}
                    </button>
                    <button
                      disabled={
                        !!filterDisabledReason() ||
                        filterApplying ||
                        filterReadPending ||
                        protectedFilter ||
                        !filterSelected.size
                      }
                      title={filterDisabledReason()}
                      data-ribbon-command="data.filter"
                      onClick={() => void applyPagedFilter()}
                    >
                      {words('Apply selected values')}
                    </button>
                    <button
                      disabled={filterReadPending || !filterSelected.size}
                      onClick={(event) => {
                        event.stopPropagation();
                        if (
                          filterScope.current === filterRenderScope &&
                          filterGeneration.current === filterFormGeneration
                        )
                          setFilterSelected(new Map());
                      }}
                    >
                      {words('Reset selected values')}
                    </button>
                  </>
                )}
              </div>
            )}
          </div>
        </details>
      </div>
      <small>{words('Sort & Filter')}</small>
    </RibbonGroup>
  ) : undefined;
  return { sortFilterCommands } as const;
}
