import type { AcceptedPrintPreview } from '../useAcceptedPrintPreview';
import type { BackgroundTile } from '../backgroundTiles';
import type { FormulaDraftHandle } from '../namedFormulaDraft';
import type { FormulaTraceController, FormulaTraceScope, TraceEndpoint } from '../formulaTrace';
import type { FormulaTraceFrame, FormulaTraceReceipt } from '../formulaTraceAdapter';
import { Fragment } from 'react';
import type {
  GridFillDragging,
  GridHomeInvoke,
  GridInkLive,
  GridInkReason,
  GridObjectControls2,
  GridSetViewport,
} from './gridTypes';
import { GridRow } from './GridRow';
import { type HomeCommandId, type HomeFrame, homeBindingMatches } from '../homeCommands';
import type { InkPenId } from '../../core/ink-stroke.mjs';
import type { ObjectAxes } from '../worksheetObjectGeometry';
import type { PageBoundaryModel } from '../pageBoundaries';
import { PageBoundaryOverlay } from '../WorkbookPageBreakPreview';
import type { ViewGeometry, ViewVirtualWindow } from '../viewGeometry';
import type { ViewPreferences } from '../viewPreferences';
import type { WorkbookCommand, WorkbookGeometryProfile, WorkbookNormalFont, WorkbookSheet } from '../../core/contracts';
import { WorkbookPageLayout } from '../WorkbookPageLayout';
import { WorkbookTraceOverlay } from '../WorkbookTraceOverlay';
import { WorksheetInkLayer } from '../WorksheetInkLayer';
import { WorksheetObjectLayer } from '../WorksheetObjectLayer';
import { colName, formulaRowVersion } from '../sheetModel';
import { geometryFontFamily } from '../worksheetFontMetrics';
import { rectangle } from './rangeAddress';

export interface GridViewportInputs {
  activationEpoch: React.MutableRefObject<number>;
  active: boolean;
  activeRange: () => string;
  boundaryModel: PageBoundaryModel | null;
  busy: boolean;
  columnOffsets: Float64Array<ArrayBufferLike>;
  columnWindow: ViewVirtualWindow;
  dragEnter: (cell: string) => void;
  dragStart: (cell: string, extend: boolean) => void;
  drawMode: 'draw' | 'select' | 'erase' | null;
  fillDragging: GridFillDragging;
  fillPreview: string | null;
  firstRenderedColumn: number;
  firstRenderedRow: number;
  formulaDraftHandle: React.MutableRefObject<FormulaDraftHandle | null>;
  frozenColumnCount: number;
  frozenColumnGap: number;
  frozenRowCount: number;
  frozenRowGap: number;
  functionViewMode: React.MutableRefObject<{ mode: 'normal' | 'pageLayout' | 'pageBreak'; epoch: number }>;
  geometry: ViewGeometry;
  geometryKey: string;
  geometryProfile: WorkbookGeometryProfile;
  gridScroll: React.RefObject<HTMLDivElement>;
  homeDisabled: (id: HomeCommandId) => boolean;
  homeFrameRef: React.MutableRefObject<HomeFrame | null>;
  homeInvoke: GridHomeInvoke;
  homeRenderFrame: HomeFrame;
  inkKey: string;
  inkLive: GridInkLive;
  inkPreset: InkPenId;
  inkReason: GridInkReason;
  lastRenderedColumn: number;
  lastRenderedRow: number;
  literal: boolean | undefined;
  moveAfterCommit: (from: string, key: 'Enter' | 'Tab', shift: boolean) => void;
  navigateTrace: (target: TraceEndpoint) => void;
  normalFont: Readonly<WorkbookNormalFont>;
  objectAxes: ObjectAxes;
  objectControls: GridObjectControls2;
  objectFamily: string;
  objectFamilyEpoch: React.MutableRefObject<{ family: string; epoch: number }>;
  onCommandBatch: ((commands: WorkbookCommand[]) => Promise<void>) | undefined;
  openTracePane: () => void;
  passiveNotice: string | undefined;
  printPreview: AcceptedPrintPreview;
  range: string;
  readOnly: boolean;
  renderedColumnIndices: number[];
  renderedRows: number[];
  returnNormal: () => void;
  rowOffsets: Float64Array<ArrayBufferLike>;
  rowWindow: ViewVirtualWindow;
  select: (cell: string, extend?: boolean) => void;
  selectColumn: (col: number) => void;
  selected: string;
  selectedRef: React.MutableRefObject<string>;
  selectRow: (row: number) => void;
  setGridError: React.Dispatch<React.SetStateAction<string>>;
  setRange: React.Dispatch<React.SetStateAction<string>>;
  setSelectionEnd: React.Dispatch<React.SetStateAction<string>>;
  setValue: (cell: string, value: string) => void;
  setViewport: GridSetViewport;
  sheet: WorkbookSheet;
  specialSelection: Set<string> | null;
  specialSelectionMode: React.MutableRefObject<{ selection: Set<string> | null }>;
  storageKey: string | undefined;
  tableColumnCount: number;
  traceController: FormulaTraceController;
  traceControls: { receipt: FormulaTraceReceipt; trace: FormulaTraceFrame['trace'] } | undefined;
  traceScope: FormulaTraceScope;
  updateObjectExtent: (family: string, extent: { rows: number; columns: number }) => void;
  view: ViewPreferences;
  viewActive: React.MutableRefObject<boolean>;
  viewMode: 'normal' | 'pageLayout' | 'pageBreak';
  viewport: { width: number; height: number; scrollTop: number; scrollLeft: number };
  visibleColumns: number;
  visibleRows: number;
  words: (en: string) => string;
  worksheetBackground: BackgroundTile | undefined;
}

export function GridViewport({
  activationEpoch,
  active,
  activeRange,
  boundaryModel,
  busy,
  columnOffsets,
  columnWindow,
  dragEnter,
  dragStart,
  drawMode,
  fillDragging,
  fillPreview,
  firstRenderedColumn,
  firstRenderedRow,
  formulaDraftHandle,
  frozenColumnCount,
  frozenColumnGap,
  frozenRowCount,
  frozenRowGap,
  functionViewMode,
  geometry,
  geometryKey,
  geometryProfile,
  gridScroll,
  homeDisabled,
  homeFrameRef,
  homeInvoke,
  homeRenderFrame,
  inkKey,
  inkLive,
  inkPreset,
  inkReason,
  lastRenderedColumn,
  lastRenderedRow,
  literal,
  moveAfterCommit,
  navigateTrace,
  normalFont,
  objectAxes,
  objectControls,
  objectFamily,
  objectFamilyEpoch,
  onCommandBatch,
  openTracePane,
  passiveNotice,
  printPreview,
  range,
  readOnly,
  renderedColumnIndices,
  renderedRows,
  returnNormal,
  rowOffsets,
  rowWindow,
  select,
  selectColumn,
  selected,
  selectedRef,
  selectRow,
  setGridError,
  setRange,
  setSelectionEnd,
  setValue,
  setViewport,
  sheet,
  specialSelection,
  specialSelectionMode,
  storageKey,
  tableColumnCount,
  traceController,
  traceControls,
  traceScope,
  updateObjectExtent,
  view,
  viewActive,
  viewMode,
  viewport,
  visibleColumns,
  visibleRows,
  words,
  worksheetBackground,
}: GridViewportInputs): JSX.Element {
  return (
    <>
      <div className="sheet-grid-viewport" hidden={viewMode === 'pageLayout'}>
        <div
          className="sheet-grid-scroll"
          ref={gridScroll}
          onScroll={(event) => {
            const scroll = event.currentTarget;
            setViewport({
              width: scroll.clientWidth,
              height: scroll.clientHeight,
              scrollTop: scroll.scrollTop,
              scrollLeft: scroll.scrollLeft,
            });
          }}
        >
          <table
            className="sheet-grid"
            ref={(table) => {
              if (table) table.inert = viewMode === 'pageBreak' || !!drawMode;
            }}
            data-gridlines={view.gridlines}
            data-headings={view.headings}
            data-zoom={view.zoom}
            data-worksheet-background={worksheetBackground?.id}
            data-show-formulas={view.showFormulas}
            data-geometry-profile={literal ? 'literal' : geometryKey}
            data-maximum-digit-width={literal ? undefined : geometryProfile.maximumDigitWidth}
            style={
              {
                width: geometry.rowHeaderWidth + columnOffsets[visibleColumns],
                fontFamily: literal
                  ? undefined
                  : geometryProfile.source === 'substituted-font'
                    ? 'sans-serif'
                    : geometryFontFamily(normalFont),
                fontWeight: literal ? undefined : normalFont.bold ? 700 : 400,
                fontStyle: literal ? undefined : normalFont.italic ? 'italic' : 'normal',
                fontSize: literal ? 12 * geometry.scale : `${normalFont.fontSize * geometry.scale}pt`,
                '--sheet-scale': geometry.scale,
                '--sheet-header-height': `${geometry.columnHeaderHeight}px`,
              } as React.CSSProperties
            }
          >
            <colgroup>
              {view.headings && <col style={{ width: geometry.rowHeaderWidth }} />}
              {renderedColumnIndices.map((c) => (
                <Fragment key={c}>
                  {c === columnWindow.indices[0] && firstRenderedColumn > frozenColumnCount && (
                    <col
                      style={{
                        width: columnOffsets[firstRenderedColumn] - columnOffsets[frozenColumnCount],
                      }}
                    />
                  )}
                  <col style={{ width: columnOffsets[c + 1] - columnOffsets[c] }} />
                </Fragment>
              ))}
              {frozenColumnGap > 0 && <col style={{ width: frozenColumnGap }} />}
              {lastRenderedColumn < visibleColumns && (
                <col style={{ width: columnOffsets[visibleColumns] - columnOffsets[lastRenderedColumn] }} />
              )}
            </colgroup>
            {view.headings && (
              <thead>
                <tr>
                  <th
                    data-grid-corner
                    style={{
                      width: geometry.rowHeaderWidth,
                      position: 'sticky',
                      top: 0,
                      left: 0,
                      zIndex: 5,
                    }}
                    title={words('Select used sheet range')}
                    onClick={() => {
                      const last = `${colName(Math.max(0, sheet.columns - 1))}${Math.max(sheet.rows, 30)}`;
                      select('A1');
                      setSelectionEnd(last);
                      setRange(rectangle('A1', last));
                    }}
                  ></th>
                  {renderedColumnIndices.map((c) => (
                    <Fragment key={c}>
                      {c === columnWindow.indices[0] && firstRenderedColumn > frozenColumnCount && (
                        <th
                          className="sheet-spacer-column"
                          style={{
                            width: columnOffsets[firstRenderedColumn] - columnOffsets[frozenColumnCount],
                            minWidth: columnOffsets[firstRenderedColumn] - columnOffsets[frozenColumnCount],
                          }}
                        />
                      )}
                      <th
                        title={words('Select used column')}
                        style={{
                          width: columnOffsets[c + 1] - columnOffsets[c],
                          ...(c < frozenColumnCount
                            ? {
                                left: geometry.rowHeaderWidth + columnOffsets[c],
                                zIndex: 4,
                                position: 'sticky' as const,
                              }
                            : {}),
                        }}
                        onClick={() => selectColumn(c)}
                      >
                        {colName(c)}
                      </th>
                    </Fragment>
                  ))}
                  {frozenColumnGap > 0 && (
                    <th className="sheet-spacer-column" style={{ width: frozenColumnGap, minWidth: frozenColumnGap }} />
                  )}
                  {lastRenderedColumn < visibleColumns && (
                    <th
                      className="sheet-spacer-column"
                      style={{
                        width: columnOffsets[visibleColumns] - columnOffsets[lastRenderedColumn],
                        minWidth: columnOffsets[visibleColumns] - columnOffsets[lastRenderedColumn],
                      }}
                    />
                  )}
                </tr>
              </thead>
            )}
            <tbody>
              {renderedRows.map((r) => (
                <Fragment key={r}>
                  {r === rowWindow.indices[0] && firstRenderedRow > frozenRowCount && (
                    <tr className="sheet-spacer-row">
                      <td
                        colSpan={tableColumnCount}
                        style={{ height: rowOffsets[firstRenderedRow] - rowOffsets[frozenRowCount] }}
                      />
                    </tr>
                  )}
                  <GridRow
                    r={r}
                    sheet={sheet}
                    columnPage={firstRenderedColumn}
                    lastRenderedColumn={lastRenderedColumn}
                    renderedRowIndices={renderedRows}
                    geometry={geometry}
                    showFormulas={view.showFormulas}
                    frozenColumnGap={frozenColumnGap}
                    frozenColumnCount={frozenColumnCount}
                    renderedColumnIndices={renderedColumnIndices}
                    visibleColumns={visibleColumns}
                    columnOffsets={columnOffsets}
                    firstRenderedRow={r < frozenRowCount ? 0 : firstRenderedRow}
                    lastRenderedRow={r < frozenRowCount ? frozenRowCount : lastRenderedRow}
                    selected={selected}
                    readOnly={readOnly}
                    allowFill={!literal && !homeDisabled('home.fill-range')}
                    range={range}
                    specialSelection={specialSelection}
                    fillPreview={fillPreview}
                    busy={busy}
                    setValue={setValue}
                    onSelect={(cell) => {
                      if (cell === selectedRef.current && range === cell && !specialSelectionMode.current.selection)
                        return;
                      select(cell);
                    }}
                    onDragStart={dragStart}
                    onDragEnter={dragEnter}
                    onFillStart={(event) => {
                      event.preventDefault();
                      event.stopPropagation();
                      if (
                        !homeFrameRef.current?.active ||
                        !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding) ||
                        homeDisabled('home.fill-range')
                      )
                        return;
                      const focused = document.activeElement as HTMLElement | null,
                        formula = formulaDraftHandle.current;
                      if (
                        (focused?.closest('.sheet-grid') && focused.getAttribute('data-editing') === 'true') ||
                        (formula?.ownsFocus() && formula.capture().composing)
                      ) {
                        setGridError('Finish or cancel the current input before using Home commands.');
                        return;
                      }
                      if (formula?.ownsFocus()) formula.park();
                      const source = activeRange();
                      fillDragging.current = {
                        source,
                        target: null,
                        dispatch: (range) => homeInvoke('home.fill-range', { source, range }),
                      };
                    }}
                    onMove={moveAfterCommit}
                    onSelectRow={selectRow}
                    formulaVersion={formulaRowVersion(sheet, r)}
                  />
                </Fragment>
              ))}
              {frozenRowGap > 0 && (
                <tr className="sheet-spacer-row">
                  <td colSpan={tableColumnCount} style={{ height: frozenRowGap }} />
                </tr>
              )}
              {lastRenderedRow < visibleRows && (
                <tr className="sheet-spacer-row">
                  <td
                    colSpan={tableColumnCount}
                    style={{ height: rowOffsets[visibleRows] - rowOffsets[lastRenderedRow] }}
                  />
                </tr>
              )}
            </tbody>
          </table>
        </div>
        {viewMode === 'pageBreak' && boundaryModel && (
          <PageBoundaryOverlay model={boundaryModel} geometry={geometry} viewport={viewport} words={words} />
        )}
        {objectControls && !literal && !passiveNotice && (
          <WorksheetObjectLayer
            keyValue={JSON.stringify([
              objectFamily,
              objectFamilyEpoch.current.epoch,
              objectControls.revision,
              storageKey ?? '',
              activationEpoch.current,
              functionViewMode.current.epoch,
            ])}
            family={objectFamily}
            gestureKey={inkKey}
            inkMode={drawMode === 'select' || drawMode === 'erase' ? drawMode : null}
            sheet={sheet.name}
            active={active && viewMode !== 'pageLayout'}
            controls={{ ...objectControls, onBatch: onCommandBatch ?? objectControls.onBatch }}
            axes={objectAxes}
            view={{
              ...viewport,
              scale: geometry.scale,
              rowHeaderWidth: geometry.rowHeaderWidth,
              columnHeaderHeight: geometry.columnHeaderHeight,
            }}
            reason={
              viewMode === 'pageBreak'
                ? 'Drawing edits are unavailable in this read-only page map.'
                : !onCommandBatch
                  ? 'Worksheet drawing gestures are unavailable in this host.'
                  : readOnly
                    ? 'This workbook is read-only.'
                    : busy
                      ? 'Wait for the workbook operation to finish.'
                      : sheet.freeze?.rows || sheet.freeze?.columns
                        ? 'Unfreeze panes before using worksheet drawing gestures.'
                        : undefined
            }
            onExtent={updateObjectExtent}
            words={words}
          />
        )}
        {objectControls && !literal && !passiveNotice && (
          <WorksheetInkLayer
            width={viewport.width}
            height={viewport.height}
            words={words}
            frame={{
              key: inkKey,
              sheet: sheet.name,
              active: active && drawMode === 'draw' && viewMode === 'normal',
              reason: inkReason,
              pen: inkPreset,
              axes: objectAxes,
              view: {
                ...viewport,
                scale: geometry.scale,
                rowHeaderWidth: geometry.rowHeaderWidth,
                columnHeaderHeight: geometry.columnHeaderHeight,
                frozenRows: sheet.freeze?.rows,
                frozenColumns: sheet.freeze?.columns,
              },
              accept: async (commands) => {
                if (
                  inkLive.current.key !== inkKey ||
                  !inkLive.current.active ||
                  inkLive.current.reason ||
                  !viewActive.current ||
                  !onCommandBatch
                )
                  throw Error('The original drawing worksheet or viewport changed.');
                await onCommandBatch(commands);
              },
            }}
            onWheel={(dx, dy) => {
              const scroll = gridScroll.current;
              if (scroll) {
                scroll.scrollLeft += dx;
                scroll.scrollTop += dy;
              }
            }}
          />
        )}
        {traceControls && (
          <WorkbookTraceOverlay
            controller={traceController}
            scope={traceScope}
            geometry={geometry}
            viewport={viewport}
            sheet={sheet}
            onNavigate={navigateTrace}
            onOpenList={openTracePane}
            words={words}
          />
        )}
      </div>
      {viewMode === 'pageLayout' && (
        <WorkbookPageLayout preview={printPreview} zoom={view.zoom} onNormal={returnNormal} words={words} />
      )}
    </>
  );
}
