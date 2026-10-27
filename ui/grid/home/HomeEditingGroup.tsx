import type { GridHomeInvoke, GridHomeProps } from '../gridTypes';
import type { HomeCommandId } from '../../homeCommands';
import { Icon } from '../../host';
import { RibbonGroup, type WorkbookTool } from '../../SpreadsheetRibbon';

export interface HomeEditingGroupInputs {
  activeTool: WorkbookTool | null;
  busy: boolean;
  fillToolMode: 'series' | 'across';
  homeDisabled: (id: HomeCommandId) => boolean;
  homeInvoke: GridHomeInvoke;
  homeMenu: (event: React.MouseEvent, menu: string) => void;
  homeProps: GridHomeProps;
  literal: boolean | undefined;
  readOnly: boolean;
  specialSelection: Set<string> | null;
  taskPaneId: string;
  words: (en: string) => string;
}

export function HomeEditingGroup({
  activeTool,
  busy,
  fillToolMode,
  homeDisabled,
  homeInvoke,
  homeMenu,
  homeProps,
  literal,
  readOnly,
  specialSelection,
  taskPaneId,
  words,
}: HomeEditingGroupInputs): JSX.Element {
  return (
    <RibbonGroup name="Editing" icon="sort" collapseBelow={1280}>
      <div className="sheet-ribbon-body sheet-ribbon-stack">
        <details className="sheet-ribbon-menu">
          <summary
            data-home-command="home.menu"
            aria-disabled={homeDisabled('home.menu')}
            onClick={(event) => homeMenu(event, 'fill')}
          >
            <Icon name="fillDown" size={18} />
            {words('Fill')} <Icon name="chevron" size={11} />
          </summary>
          <div>
            <button
              {...homeProps('home.fill-down')}
              disabled={homeDisabled('home.fill-down') || readOnly || busy || literal}
              onClick={() => void homeInvoke('home.fill-down')}
            >
              {words('Fill down')}
            </button>
            <button
              {...homeProps('home.fill-right')}
              disabled={homeDisabled('home.fill-right') || readOnly || busy || literal}
              onClick={() => void homeInvoke('home.fill-right')}
            >
              {words('Fill right')}
            </button>
            <button
              {...homeProps('home.fill-up')}
              disabled={homeDisabled('home.fill-up') || readOnly || busy || literal}
              onClick={() => void homeInvoke('home.fill-up')}
            >
              {words('Fill up')}
            </button>
            <button
              {...homeProps('home.fill-left')}
              disabled={homeDisabled('home.fill-left') || readOnly || busy || literal}
              onClick={() => void homeInvoke('home.fill-left')}
            >
              {words('Fill left')}
            </button>
            <button
              {...homeProps('home.fill-across')}
              data-workbook-tool="fill"
              aria-controls={taskPaneId}
              aria-pressed={activeTool === 'fill' && fillToolMode === 'across'}
              onClick={() => void homeInvoke('home.fill-across')}
            >
              {words('Across Worksheets…')}
            </button>
            <button
              {...homeProps('home.fill-series')}
              data-workbook-tool="fill"
              aria-controls={taskPaneId}
              aria-pressed={activeTool === 'fill' && fillToolMode === 'series'}
              disabled={homeDisabled('home.fill-series') || busy || literal}
              onClick={() => void homeInvoke('home.fill-series')}
            >
              {words('Fill series')}
            </button>
          </div>
        </details>
        <details className="sheet-ribbon-menu">
          <summary
            data-home-command="home.menu"
            aria-disabled={homeDisabled('home.menu')}
            onClick={(event) => homeMenu(event, 'clear')}
          >
            <Icon name="clear" size={18} />
            {words('Clear')} <Icon name="chevron" size={11} />
          </summary>
          <div>
            <button
              {...homeProps('home.clear-contents')}
              disabled={homeDisabled('home.clear-contents') || readOnly || busy}
              onClick={() => void homeInvoke('home.clear-contents')}
            >
              {words('Clear contents')}
            </button>
            <button
              {...homeProps('home.clear-formats')}
              disabled={homeDisabled('home.clear-formats') || readOnly || busy || literal}
              onClick={() => void homeInvoke('home.clear-formats')}
            >
              {words('Clear formats')}
            </button>
            <button
              {...homeProps('home.clear-all')}
              disabled={homeDisabled('home.clear-all') || readOnly || busy || literal}
              onClick={() => void homeInvoke('home.clear-all')}
            >
              {words('Clear all')}
            </button>
            <button {...homeProps('home.clear-hyperlinks')} onClick={() => void homeInvoke('home.clear-hyperlinks')}>
              {words('Clear Hyperlinks')}
            </button>
          </div>
        </details>
        <details className="sheet-ribbon-menu">
          <summary
            data-home-command="home.menu"
            aria-disabled={homeDisabled('home.menu')}
            onClick={(event) => homeMenu(event, 'find')}
          >
            <Icon name="search" size={25} />
            <span>{words('Find & Select')}</span> <Icon name="chevron" size={11} />
          </summary>
          <div>
            <button {...homeProps('home.find-open')} onClick={() => void homeInvoke('home.find-open')}>
              {words('Find and replace')}
            </button>
            <button
              {...homeProps('home.special-open')}
              data-workbook-tool="special"
              aria-controls={taskPaneId}
              aria-pressed={activeTool === 'special'}
              onClick={() => void homeInvoke('home.special-open')}
            >
              {words('Go To Special…')}
            </button>
            <button {...homeProps('home.select-blanks')} onClick={() => void homeInvoke('home.select-blanks')}>
              {words('Select blanks')}
            </button>
            <button {...homeProps('home.select-formulas')} onClick={() => void homeInvoke('home.select-formulas')}>
              {words('Select formulas')}
            </button>
            <button
              {...homeProps('home.clear-special-selection')}
              disabled={homeDisabled('home.clear-special-selection') || !specialSelection}
              onClick={() => void homeInvoke('home.clear-special-selection')}
            >
              {words('Clear special selection')}
            </button>
          </div>
        </details>
      </div>
      <small>{words('Editing')}</small>
    </RibbonGroup>
  );
}
