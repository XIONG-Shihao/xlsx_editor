import { BooleanCellControls, CalculatedCell, ConditionalDefinition, OwnedPivotPresentation } from './gridContexts';
import { CellInput } from './CellInput';
import { Fragment, memo, useContext } from 'react';
import type { ViewGeometry } from '../viewGeometry';
import type { WorkbookSheet } from '../../core/contracts';
import { conditionalFill as acceptedConditionalFill } from '../../core/rules-model.mjs';
import { addressPosition, colName, displayDetails, inputDisplay } from '../sheetModel';
import { backgroundCellStyle, backgroundTile } from '../backgroundTiles';
import { booleanCellValue } from '../../core/cell-insertions.mjs';
import { conditionalFill, tableAppearance } from './cellAppearance';
import { geometryFontFamily } from '../worksheetFontMetrics';
import { inRange } from './rangeAddress';
import { pivotCellStyle } from '../pivotStyles';
import { useDocumentWords } from '../words';

export interface RowProps {
  r: number;
  sheet: WorkbookSheet;
  columnPage: number;
  lastRenderedColumn: number;
  renderedRowIndices: number[];
  geometry: ViewGeometry;
  showFormulas: boolean;
  frozenColumnGap: number;
  frozenColumnCount: number;
  renderedColumnIndices: number[];
  visibleColumns: number;
  columnOffsets: Float64Array;
  firstRenderedRow: number;
  lastRenderedRow: number;
  selected: string;
  readOnly: boolean;
  allowFill: boolean;
  setValue: (cell: string, value: string) => void;
  onSelect: (cell: string, extend?: boolean) => void;
  onDragStart: (cell: string, extend: boolean) => void;
  onDragEnter: (cell: string) => void;
  onMove: (cell: string, key: 'Enter' | 'Tab', shift: boolean) => void;
  onSelectRow: (row: number) => void;
  range: string;
  specialSelection: Set<string> | null;
  fillPreview: string | null;
  onFillStart: (event: React.MouseEvent) => void;
  busy?: boolean;
  formulaVersion: number;
}
export const GridRow = memo(
  function GridRow({
    r,
    sheet,
    columnPage,
    lastRenderedColumn,
    renderedRowIndices,
    geometry,
    showFormulas,
    frozenColumnGap,
    frozenColumnCount,
    renderedColumnIndices,
    visibleColumns,
    columnOffsets,
    firstRenderedRow,
    lastRenderedRow,
    selected,
    readOnly,
    allowFill,
    setValue,
    onSelect,
    onDragStart,
    onDragEnter,
    onMove,
    onSelectRow,
    range,
    specialSelection,
    fillPreview,
    onFillStart,
    busy,
  }: RowProps): JSX.Element {
    const worksheetBackground = backgroundTile(sheet, geometry.scale);
    const conditionalDefinition = useContext(ConditionalDefinition);
    const booleanControls = useContext(BooleanCellControls);
    const ownedPivots = useContext(OwnedPivotPresentation);
    const words = useDocumentWords();
    const merged = (
      ref: string
    ): { anchorRow: number; anchorCol: number; endRow: number; endCol: number; range: string } | null => {
      const p = addressPosition(ref);
      for (const range of sheet.merges) {
        const [from, to] = range.split(':'),
          a = addressPosition(from),
          b = addressPosition(to ?? from);
        if (p.row >= a.row && p.row <= b.row && p.col >= a.col && p.col <= b.col)
          return {
            anchorRow: a.row,
            anchorCol: a.col,
            endRow: b.row,
            endCol: b.col,
            range,
          };
      }
      return null;
    };
    return (
      <tr style={{ height: geometry.rowOffsets[r + 1] - geometry.rowOffsets[r] }}>
        {geometry.headings && (
          <th
            style={{
              padding: 0,
              width: geometry.rowHeaderWidth,
              top: r < geometry.frozenRows ? geometry.columnHeaderHeight + geometry.rowOffsets[r] : 'auto',
              zIndex: r < geometry.frozenRows ? 4 : 2,
            }}
            title="Select used row"
            onClick={() => onSelectRow(r)}
          >
            <div style={{ height: geometry.rowOffsets[r + 1] - geometry.rowOffsets[r], overflow: 'hidden' }}>
              {r + 1}
            </div>
          </th>
        )}
        {renderedColumnIndices.map((c) => {
          const ref = `${colName(c)}${r + 1}`,
            merge = merged(ref);
          const columnWindowStart = c < frozenColumnCount ? 0 : columnPage;
          const columnWindowEnd = c < frozenColumnCount ? frozenColumnCount : lastRenderedColumn;
          const firstMergeRow = merge
            ? renderedRowIndices.find(
                (row) =>
                  row >= Math.max(merge.anchorRow, firstRenderedRow) && row <= merge.endRow && row < lastRenderedRow
              )
            : r;
          const firstMergeColumn = merge
            ? renderedColumnIndices.find(
                (col) =>
                  col >= Math.max(merge.anchorCol, columnWindowStart) && col <= merge.endCol && col < columnWindowEnd
              )
            : c;
          const representative = !merge || (r === firstMergeRow && c === firstMergeColumn);
          const leadingSpacer =
            c === renderedColumnIndices.find((col) => col >= columnPage) && columnPage > frozenColumnCount ? (
              <td
                className="sheet-spacer-column"
                style={{
                  width: columnOffsets[columnPage] - columnOffsets[frozenColumnCount],
                  minWidth: columnOffsets[columnPage] - columnOffsets[frozenColumnCount],
                }}
              />
            ) : null;
          if (!representative) return leadingSpacer ? <Fragment key={c}>{leadingSpacer}</Fragment> : null;
          const anchorRef = merge ? `${colName(merge.anchorCol)}${merge.anchorRow + 1}` : ref;
          const projected = !!merge && anchorRef !== ref;
          const cell = sheet.cells[anchorRef];
          const displayed = displayDetails(sheet, anchorRef),
            showingFormula = showFormulas && !!cell?.formula;
          const formatColor = showingFormula ? undefined : displayed.color;
          const formatNotice =
            !showingFormula && displayed.formatStatus === 'unsupported'
              ? words('Number format is not supported; showing the unformatted value.')
              : undefined;
          const width = sheet.widths.find((w) => c + 1 >= w.min && c + 1 <= w.max);
          const legacyHighlight = conditionalDefinition === undefined ? conditionalFill(sheet, anchorRef) : undefined;
          const highlight = conditionalDefinition
            ? acceptedConditionalFill(conditionalDefinition.rules, anchorRef, cell, conditionalDefinition.duplicates)
            : legacyHighlight
              ? legacyHighlight.startsWith('#')
                ? legacyHighlight
                : '#' + legacyHighlight
              : undefined;
          const tableStyle = tableAppearance(sheet, anchorRef);
          const pivotStyle = pivotCellStyle(ownedPivots, sheet.name, anchorRef);
          let booleanValue: boolean | null = null;
          try {
            if (booleanControls && !merge) booleanValue = booleanCellValue(cell ?? null);
          } catch {
            /* Ordinary cell presentation is retained outside this profile. */
          }
          return (
            <Fragment key={c}>
              {leadingSpacer}
              <td
                data-pivot-style={pivotStyle?.sample}
                data-pivot-role={pivotStyle?.role}
                data-pivot-cell={pivotStyle ? anchorRef : undefined}
                colSpan={
                  merge
                    ? renderedColumnIndices.filter((col) => col >= c && col <= merge.endCol && col < columnWindowEnd)
                        .length
                    : undefined
                }
                rowSpan={
                  merge
                    ? renderedRowIndices.filter((row) => row >= r && row <= merge.endRow && row < lastRenderedRow)
                        .length
                    : undefined
                }
                className={
                  (specialSelection ? specialSelection.has(anchorRef) : inRange(range, anchorRef))
                    ? 'selected'
                    : fillPreview && inRange(fillPreview, anchorRef)
                      ? 'fill-preview'
                      : ''
                }
                aria-label={projected ? `Merged continuation of ${merge.range}` : undefined}
                aria-selected={specialSelection ? specialSelection.has(anchorRef) : inRange(range, anchorRef)}
                onMouseDown={(event) => {
                  if (event.shiftKey) event.preventDefault();
                  onDragStart(anchorRef, event.shiftKey);
                }}
                onMouseEnter={() => onDragEnter(anchorRef)}
                style={{
                  width: columnOffsets[Math.min(merge ? merge.endCol + 1 : c + 1, columnWindowEnd)] - columnOffsets[c],
                  background:
                    highlight ??
                    (cell?.fillColor
                      ? `#${cell.fillColor}`
                      : (pivotStyle?.background ?? tableStyle?.background ?? 'var(--bg-panel, white)')),
                  ...backgroundCellStyle(
                    worksheetBackground,
                    columnOffsets[c],
                    geometry.rowOffsets[r],
                    !!(highlight || cell?.fillColor || pivotStyle?.background || tableStyle?.background)
                  ),
                  boxShadow: cell?.border === 'all' ? 'inset 0 0 0 1px currentColor' : undefined,
                  verticalAlign: cell?.verticalAlign === 'center' ? 'middle' : cell?.verticalAlign,
                  position: r < geometry.frozenRows || c < geometry.frozenColumns ? 'sticky' : undefined,
                  top: r < geometry.frozenRows ? geometry.columnHeaderHeight + geometry.rowOffsets[r] : undefined,
                  left: c < geometry.frozenColumns ? geometry.rowHeaderWidth + columnOffsets[c] : undefined,
                  zIndex:
                    r < geometry.frozenRows && c < geometry.frozenColumns
                      ? 3
                      : r < geometry.frozenRows || c < geometry.frozenColumns
                        ? 2
                        : undefined,
                }}
              >
                {projected ? (
                  <span aria-hidden="true" />
                ) : (
                  <>
                    <div
                      className={cell?.wrapText ? 'sheet-cell-wrapped' : undefined}
                      style={{
                        height:
                          geometry.rowOffsets[Math.min(merge ? merge.endRow + 1 : r + 1, lastRenderedRow)] -
                          geometry.rowOffsets[r],
                        overflow: 'hidden',
                      }}
                    >
                      {booleanControls && booleanValue !== null && (
                        <input
                          type="checkbox"
                          className="sheet-boolean-toggle"
                          aria-label={`${sheet.name} ${ref} checkbox`}
                          title={words('Editor TRUE/FALSE checkbox')}
                          checked={booleanValue}
                          disabled={booleanControls.disabled}
                          onMouseDown={(event) => {
                            event.stopPropagation();
                            onSelect(ref);
                          }}
                          onFocus={() => onSelect(ref)}
                          onKeyDown={(event) => {
                            if (event.key === ' ') event.stopPropagation();
                          }}
                          onChange={() => {
                            void booleanControls.toggle(sheet.name, ref, booleanValue!);
                          }}
                        />
                      )}
                      <CellInput
                        pending={busy}
                        gridCell
                        display={showingFormula ? inputDisplay(cell) : displayed.text}
                        data-number-format-status={showingFormula ? 'not-applicable' : displayed.formatStatus}
                        data-number-format-color={formatColor}
                        title={formatNotice}
                        aria-label={`${sheet.name} ${ref}`}
                        aria-description={
                          [
                            cell?.calculation?.pending
                              ? 'Pending calculation; displayed result is stored and not current'
                              : cell?.formula
                                ? `Formula ${cell.formula}; calculation ${cell.calculation?.status ?? 'pending'}`
                                : undefined,
                            formatNotice,
                          ]
                            .filter(Boolean)
                            .join('. ') || undefined
                        }
                        value={inputDisplay(cell)}
                        style={{
                          paddingLeft:
                            booleanControls && booleanValue !== null
                              ? '25px'
                              : pivotStyle?.paddingLeft !== undefined
                                ? `${pivotStyle.paddingLeft * geometry.scale}px`
                                : undefined,
                          fontWeight:
                            (cell?.bold ?? sheet.normalFont?.bold) || tableStyle?.header || pivotStyle?.bold
                              ? 700
                              : 400,
                          fontStyle: (cell?.italic ?? sheet.normalFont?.italic) ? 'italic' : 'normal',
                          textDecoration: cell?.underline ? 'underline' : undefined,
                          fontFamily: cell?.fontFamily
                            ? geometryFontFamily({
                                fontFamily: cell.fontFamily,
                                fontSize: cell.fontSize ?? 11,
                                bold: false,
                                italic: false,
                              })
                            : undefined,
                          fontSize: cell?.fontSize ? `${cell.fontSize * geometry.scale}pt` : undefined,
                          color: formatColor
                            ? `#${formatColor}`
                            : cell?.fontColor
                              ? `#${cell.fontColor}`
                              : (pivotStyle?.color ?? tableStyle?.color),
                          textAlign: cell?.horizontalAlign,
                        }}
                        readOnly={readOnly}
                        onFocus={() => onSelect(ref)}
                        onCommit={(value) => setValue(ref, value)}
                        onMove={(key, shift) => onMove(ref, key, shift)}
                      />
                      {cell?.wrapText && (
                        <span
                          className="sheet-wrapped-display"
                          aria-hidden="true"
                          style={{
                            fontFamily: cell.fontFamily
                              ? geometryFontFamily({
                                  fontFamily: cell.fontFamily,
                                  fontSize: cell.fontSize ?? 11,
                                  bold: false,
                                  italic: false,
                                })
                              : undefined,
                            fontSize: cell.fontSize ? `${cell.fontSize * geometry.scale}pt` : undefined,
                            fontWeight: (cell.bold ?? sheet.normalFont?.bold) ? 700 : 400,
                            fontStyle: (cell.italic ?? sheet.normalFont?.italic) ? 'italic' : 'normal',
                            color: formatColor ? `#${formatColor}` : cell.fontColor ? `#${cell.fontColor}` : undefined,
                            textAlign: cell.horizontalAlign,
                          }}
                        >
                          {showingFormula ? inputDisplay(cell) : displayed.text}
                        </span>
                      )}
                      {(cell?.calculation?.pending ||
                        cell?.formula ||
                        cell?.formulaType ||
                        (typeof cell?.value === 'number' && cell.format !== undefined && cell.format !== 0)) && (
                        <CalculatedCell sheet={sheet} cellRef={ref} />
                      )}
                    </div>
                    {ref === (range.split(':')[1] ?? range) && !readOnly && allowFill && !specialSelection && (
                      <button
                        data-home-command="home.fill-range"
                        className="sheet-fill-handle"
                        aria-label="Drag to fill selected range"
                        title="Drag to fill cells"
                        onMouseDown={onFillStart}
                      />
                    )}
                  </>
                )}
              </td>
            </Fragment>
          );
        })}
        {frozenColumnGap > 0 && (
          <td className="sheet-spacer-column" style={{ width: frozenColumnGap, minWidth: frozenColumnGap }} />
        )}
        {lastRenderedColumn < visibleColumns && (
          <td
            className="sheet-spacer-column"
            style={{
              width: columnOffsets[visibleColumns] - columnOffsets[lastRenderedColumn],
              minWidth: columnOffsets[visibleColumns] - columnOffsets[lastRenderedColumn],
            }}
          />
        )}
      </tr>
    );
  },
  (a, b) => {
    if (
      a.r !== b.r ||
      a.columnPage !== b.columnPage ||
      a.lastRenderedColumn !== b.lastRenderedColumn ||
      a.renderedRowIndices !== b.renderedRowIndices ||
      a.geometry !== b.geometry ||
      a.showFormulas !== b.showFormulas ||
      a.frozenColumnGap !== b.frozenColumnGap ||
      a.frozenColumnCount !== b.frozenColumnCount ||
      a.renderedColumnIndices !== b.renderedColumnIndices ||
      a.firstRenderedRow !== b.firstRenderedRow ||
      a.lastRenderedRow !== b.lastRenderedRow ||
      a.visibleColumns !== b.visibleColumns ||
      a.columnOffsets !== b.columnOffsets ||
      a.readOnly !== b.readOnly ||
      a.allowFill !== b.allowFill ||
      (a.busy !== b.busy && (addressPosition(a.selected).row === a.r || addressPosition(b.selected).row === b.r)) ||
      a.range !== b.range ||
      a.specialSelection !== b.specialSelection ||
      a.fillPreview !== b.fillPreview ||
      a.onFillStart !== b.onFillStart ||
      a.sheet.hiddenRows !== b.sheet.hiddenRows ||
      a.sheet.rowHeights !== b.sheet.rowHeights ||
      a.sheet.freeze !== b.sheet.freeze ||
      a.setValue !== b.setValue ||
      a.formulaVersion !== b.formulaVersion ||
      a.sheet.name !== b.sheet.name ||
      a.sheet.merges !== b.sheet.merges ||
      a.sheet.widths !== b.sheet.widths ||
      a.sheet.columns !== b.sheet.columns ||
      a.sheet.date1904 !== b.sheet.date1904
    )
      return false;
    if (
      a.selected !== b.selected &&
      (addressPosition(a.selected).row === a.r || addressPosition(b.selected).row === b.r)
    )
      return false;
    for (const col of b.renderedColumnIndices) {
      const ref = `${colName(col)}${b.r + 1}`;
      if (a.sheet.cells[ref] !== b.sheet.cells[ref]) return false;
    }
    return true;
  }
);
