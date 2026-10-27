import type { GridHomeInvoke, GridHomeProps } from '../gridTypes';
import type { HomeCommandId } from '../../homeCommands';
import { Icon } from '../../host';
import { RibbonGroup } from '../../SpreadsheetRibbon';
import type { WorkbookCell, WorkbookSheet } from '../../../core/contracts';

export interface HomeAlignmentGroupInputs {
  activeRange: () => string;
  busy: boolean;
  homeDisabled: (id: HomeCommandId) => boolean;
  homeInvoke: GridHomeInvoke;
  homeMenu: (event: React.MouseEvent, menu: string) => void;
  homeProps: GridHomeProps;
  homeReason: (id: HomeCommandId) => string | undefined;
  literal: boolean | undefined;
  mixedStyle: (key: keyof WorkbookCell) => boolean;
  readOnly: boolean;
  sheet: WorkbookSheet;
  styleCell: WorkbookCell;
  words: (en: string) => string;
}

export function HomeAlignmentGroup({
  activeRange,
  busy,
  homeDisabled,
  homeInvoke,
  homeMenu,
  homeProps,
  homeReason,
  literal,
  mixedStyle,
  readOnly,
  sheet,
  styleCell,
  words,
}: HomeAlignmentGroupInputs): JSX.Element {
  return (
    <RibbonGroup name="Alignment" icon="alignLeft" collapseBelow={1020}>
      <div className="sheet-ribbon-body sheet-ribbon-columns">
        <div className="sheet-ribbon-row">
          <button
            {...homeProps('home.align-top')}
            title={homeReason('home.align-top') ?? words('Align top')}
            aria-label={words('Align top')}
            aria-pressed={mixedStyle('verticalAlign') ? 'mixed' : styleCell?.verticalAlign === 'top'}
            disabled={homeDisabled('home.align-top') || readOnly || busy || literal}
            onClick={() => void homeInvoke('home.align-top')}
          >
            <Icon name="alignTop" size={17} />
          </button>
          <button
            {...homeProps('home.align-middle')}
            title={homeReason('home.align-middle') ?? words('Align middle')}
            aria-label={words('Align middle')}
            aria-pressed={mixedStyle('verticalAlign') ? 'mixed' : styleCell?.verticalAlign === 'center'}
            disabled={homeDisabled('home.align-middle') || readOnly || busy || literal}
            onClick={() => void homeInvoke('home.align-middle')}
          >
            <Icon name="alignMiddle" size={17} />
          </button>
          <button
            {...homeProps('home.align-bottom')}
            title={homeReason('home.align-bottom') ?? words('Align bottom')}
            aria-label={words('Align bottom')}
            aria-pressed={mixedStyle('verticalAlign') ? 'mixed' : styleCell?.verticalAlign === 'bottom'}
            disabled={homeDisabled('home.align-bottom') || readOnly || busy || literal}
            onClick={() => void homeInvoke('home.align-bottom')}
          >
            <Icon name="alignBottom" size={17} />
          </button>
          <button
            {...homeProps('home.wrap')}
            className={mixedStyle('wrapText') ? 'mixed' : styleCell?.wrapText ? 'active' : ''}
            aria-pressed={mixedStyle('wrapText') ? 'mixed' : !!styleCell?.wrapText}
            disabled={homeDisabled('home.wrap') || readOnly || busy || literal}
            onClick={() => void homeInvoke('home.wrap', { value: !styleCell?.wrapText })}
          >
            <Icon name="wrap" size={17} />
            {words('Wrap Text')}
          </button>
        </div>
        <div className="sheet-ribbon-row">
          <button
            {...homeProps('home.align-left')}
            title={homeReason('home.align-left') ?? words('Align left')}
            aria-label={words('Align left')}
            aria-pressed={mixedStyle('horizontalAlign') ? 'mixed' : styleCell?.horizontalAlign === 'left'}
            disabled={homeDisabled('home.align-left') || readOnly || busy || literal}
            onClick={() => void homeInvoke('home.align-left')}
          >
            <Icon name="alignLeft" size={17} />
          </button>
          <button
            {...homeProps('home.align-center')}
            title={homeReason('home.align-center') ?? words('Align center')}
            aria-label={words('Align center')}
            aria-pressed={mixedStyle('horizontalAlign') ? 'mixed' : styleCell?.horizontalAlign === 'center'}
            disabled={homeDisabled('home.align-center') || readOnly || busy || literal}
            onClick={() => void homeInvoke('home.align-center')}
          >
            <Icon name="alignCenter" size={17} />
          </button>
          <button
            {...homeProps('home.align-right')}
            title={homeReason('home.align-right') ?? words('Align right')}
            aria-label={words('Align right')}
            aria-pressed={mixedStyle('horizontalAlign') ? 'mixed' : styleCell?.horizontalAlign === 'right'}
            disabled={homeDisabled('home.align-right') || readOnly || busy || literal}
            onClick={() => void homeInvoke('home.align-right')}
          >
            <Icon name="alignRight" size={17} />
          </button>
          <details className="sheet-ribbon-menu">
            <summary
              data-home-command="home.menu"
              aria-disabled={homeDisabled('home.menu')}
              onClick={(event) => homeMenu(event, 'merge')}
            >
              <Icon name="merge" size={17} />
              {words('Merge & Center')} <Icon name="chevron" size={11} />
            </summary>
            <div>
              <button
                {...homeProps('home.merge-center')}
                disabled={
                  homeDisabled('home.merge-center') || readOnly || busy || literal || !activeRange().includes(':')
                }
                onClick={() => void homeInvoke('home.merge-center')}
              >
                {words('Merge & Center')}
              </button>
              <button
                {...homeProps('home.unmerge')}
                disabled={
                  homeDisabled('home.unmerge') || readOnly || busy || literal || !sheet.merges.includes(activeRange())
                }
                onClick={() => void homeInvoke('home.unmerge')}
              >
                {words('Unmerge cells')}
              </button>
            </div>
          </details>
        </div>
      </div>
      <small>{words('Alignment')}</small>
    </RibbonGroup>
  );
}
