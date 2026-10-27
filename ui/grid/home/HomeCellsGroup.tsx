import { DimensionInput } from '../DimensionInput';
import type { GridHomeInvoke, GridHomeProps } from '../gridTypes';
import type { HomeBinding, HomeCommandId, HomeFrame } from '../../homeCommands';
import { Icon } from '../../host';
import { RibbonGroup } from '../../SpreadsheetRibbon';
import type { WorkbookSheet } from '../../../core/contracts';
import type { WorksheetDimensions } from '../../../core/worksheet-geometry.mjs';
import { addressPosition } from '../../sheetModel';

export interface HomeCellsGroupInputs {
  busy: boolean;
  homeDisabled: (id: HomeCommandId) => boolean;
  homeFrameRef: React.MutableRefObject<HomeFrame | null>;
  homeInvoke: GridHomeInvoke;
  homeMenu: (event: React.MouseEvent, menu: string) => void;
  homeProps: GridHomeProps;
  homeRenderFrame: HomeFrame;
  homeSameTarget: (binding: HomeBinding) => boolean;
  literal: boolean | undefined;
  readOnly: boolean;
  selected: string;
  setGridError: React.Dispatch<React.SetStateAction<string>>;
  sheet: WorkbookSheet;
  structuralBlock: string;
  words: (en: string) => string;
  worksheetDimensions: WorksheetDimensions;
}

export function HomeCellsGroup({
  busy,
  homeDisabled,
  homeFrameRef,
  homeInvoke,
  homeMenu,
  homeProps,
  homeRenderFrame,
  homeSameTarget,
  literal,
  readOnly,
  selected,
  setGridError,
  sheet,
  structuralBlock,
  words,
  worksheetDimensions,
}: HomeCellsGroupInputs): JSX.Element {
  return (
    <RibbonGroup name="Cells" icon="table" collapseBelow={1160}>
      <div className="sheet-ribbon-body sheet-ribbon-stack">
        <details className="sheet-ribbon-menu">
          <summary
            data-home-command="home.menu"
            aria-disabled={homeDisabled('home.menu')}
            onClick={(event) => homeMenu(event, 'insert')}
          >
            <Icon name="insertCells" size={18} />
            {words('Insert')} <Icon name="chevron" size={11} />
          </summary>
          <div>
            <button
              {...homeProps('home.insert-row')}
              disabled={
                homeDisabled('home.insert-row') || readOnly || busy || literal || !sheet.capabilities?.insertRows
              }
              onClick={() => void homeInvoke('home.insert-row')}
            >
              {words('Insert row')}
            </button>
            <button
              {...homeProps('home.insert-column')}
              disabled={
                homeDisabled('home.insert-column') ||
                readOnly ||
                busy ||
                literal ||
                (!!structuralBlock && !sheet.capabilities?.referenceColumns)
              }
              onClick={() => void homeInvoke('home.insert-column')}
            >
              {words('Insert column')}
            </button>
          </div>
        </details>
        <details className="sheet-ribbon-menu">
          <summary
            data-home-command="home.menu"
            aria-disabled={homeDisabled('home.menu')}
            onClick={(event) => homeMenu(event, 'delete')}
          >
            <Icon name="deleteCells" size={18} />
            {words('Delete')} <Icon name="chevron" size={11} />
          </summary>
          <div>
            <button
              {...homeProps('home.delete-row')}
              disabled={
                homeDisabled('home.delete-row') ||
                readOnly ||
                busy ||
                literal ||
                (!!structuralBlock && !sheet.capabilities?.deleteRows)
              }
              onClick={() => void homeInvoke('home.delete-row')}
            >
              {words('Delete row')}
            </button>
            <button
              {...homeProps('home.delete-column')}
              disabled={
                homeDisabled('home.delete-column') ||
                readOnly ||
                busy ||
                literal ||
                (!!structuralBlock && !sheet.capabilities?.referenceColumns)
              }
              onClick={() => void homeInvoke('home.delete-column')}
            >
              {words('Delete column')}
            </button>
          </div>
        </details>
        <details className="sheet-ribbon-menu">
          <summary
            data-home-command="home.menu"
            aria-disabled={homeDisabled('home.menu')}
            onClick={(event) => homeMenu(event, 'format')}
          >
            <Icon name="cellSize" size={18} />
            {words('Format')} <Icon name="chevron" size={11} />
          </summary>
          <div>
            <label>
              {words('Column width')}
              <DimensionInput
                key={`${sheet.name}:${selected}:width`}
                value={
                  sheet.widths.find(
                    (width) =>
                      addressPosition(selected).col + 1 >= width.min && addressPosition(selected).col + 1 <= width.max
                  )?.width ?? (literal ? 12.5 : worksheetDimensions.defaultColumnWidth)
                }
                min={0.01}
                max={255}
                disabled={homeDisabled('home.column-width') || readOnly || busy || literal}
                onError={(message) => {
                  if (
                    homeSameTarget(homeRenderFrame.binding) &&
                    homeFrameRef.current?.binding.state === homeRenderFrame.binding.state
                  )
                    setGridError(message);
                }}
                onCommit={(value) => void homeInvoke('home.column-width', { value })}
                commandId="home.column-width"
                onDraft={(value, accept) =>
                  void homeInvoke('home.dimension-draft', { dimension: 'column-width', value }, () => accept())
                }
              />
            </label>
            <label>
              {words('Row height')}
              <DimensionInput
                key={`${sheet.name}:${selected}:height`}
                value={
                  sheet.rowHeights?.find((row) => row.row === addressPosition(selected).row)?.height ??
                  (literal ? 43.5 : worksheetDimensions.defaultRowHeight)
                }
                min={0.01}
                max={409.5}
                disabled={homeDisabled('home.row-height') || readOnly || busy || literal}
                onError={(message) => {
                  if (
                    homeSameTarget(homeRenderFrame.binding) &&
                    homeFrameRef.current?.binding.state === homeRenderFrame.binding.state
                  )
                    setGridError(message);
                }}
                onCommit={(value) => void homeInvoke('home.row-height', { value })}
                commandId="home.row-height"
                onDraft={(value, accept) =>
                  void homeInvoke('home.dimension-draft', { dimension: 'row-height', value }, () => accept())
                }
              />
            </label>
          </div>
        </details>
      </div>
      <small>{words('Cells')}</small>
    </RibbonGroup>
  );
}
