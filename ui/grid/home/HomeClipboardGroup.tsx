import type { GridHomeInvoke, GridHomeProps, GridImportPaste } from '../gridTypes';
import type { HomeCommandId } from '../../homeCommands';
import { Icon } from '../../host';
import { RibbonGroup } from '../../SpreadsheetRibbon';

export interface HomeClipboardGroupInputs {
  busy: boolean;
  homeDisabled: (id: HomeCommandId) => boolean;
  homeInvoke: GridHomeInvoke;
  homeMenu: (event: React.MouseEvent, menu: string) => void;
  homeProps: GridHomeProps;
  homeReason: (id: HomeCommandId) => string | undefined;
  importPaste: GridImportPaste;
  literal: boolean | undefined;
  paintStyleId: number | null;
  readOnly: boolean;
  words: (en: string) => string;
}

export function HomeClipboardGroup({
  busy,
  homeDisabled,
  homeInvoke,
  homeMenu,
  homeProps,
  homeReason,
  importPaste,
  literal,
  paintStyleId,
  readOnly,
  words,
}: HomeClipboardGroupInputs): JSX.Element {
  return (
    <RibbonGroup name="Clipboard" icon="paste" collapseBelow={720}>
      <div className="sheet-ribbon-body">
        <div className="sheet-ribbon-paste">
          <button
            {...homeProps('home.paste')}
            className="sheet-ribbon-large"
            title={homeReason('home.paste') ?? words('Paste with ⌘V')}
            onClick={() => void homeInvoke('home.paste')}
            disabled={homeDisabled('home.paste') || readOnly || busy}
          >
            <Icon name="paste" size={26} />
            {words('Paste')}
          </button>
          <details className="sheet-ribbon-menu">
            <summary
              aria-label={words('Paste Special')}
              title={words('Paste Special')}
              data-home-command="home.menu"
              aria-disabled={homeDisabled('home.menu')}
              onClick={(event) => homeMenu(event, 'paste')}
            >
              <Icon name="chevron" size={11} />
            </summary>
            <div>
              <small>
                {words(
                  importPaste && !literal
                    ? 'Paste: up to 1,000,000 cells. Paste Special: up to 1,000 cells. Each action has one undo step.'
                    : 'Paste: up to 10,000 cells. Paste Special: up to 1,000 cells. Each action has one undo step.'
                )}
              </small>
              <small>{words('Paste as text: up to 10,000 cells and 2 MiB per action. Each entry stays text.')}</small>
              <button
                {...homeProps('home.paste-text')}
                disabled={homeDisabled('home.paste-text') || readOnly || busy}
                onClick={() => void homeInvoke('home.paste-text')}
              >
                {words('Paste as text')}
              </button>
              <button
                {...homeProps('home.paste-values')}
                disabled={homeDisabled('home.paste-values') || readOnly || busy || literal}
                onClick={() => void homeInvoke('home.paste-values')}
              >
                {words('Paste values')}
              </button>
              <button
                {...homeProps('home.paste-formulas')}
                disabled={homeDisabled('home.paste-formulas') || readOnly || busy || literal}
                onClick={() => void homeInvoke('home.paste-formulas')}
              >
                {words('Paste formulas')}
              </button>
              <button
                {...homeProps('home.paste-formatting')}
                disabled={homeDisabled('home.paste-formatting') || readOnly || busy || literal}
                onClick={() => void homeInvoke('home.paste-formatting')}
              >
                {words('Paste formatting')}
              </button>
            </div>
          </details>
        </div>
        <div className="sheet-ribbon-stack">
          <button
            {...homeProps('home.cut')}
            title={homeReason('home.cut') ?? words('Cut selected range')}
            onClick={() => void homeInvoke('home.cut')}
            disabled={homeDisabled('home.cut') || readOnly || busy}
          >
            <Icon name="cut" size={16} />
            {words('Cut')}
          </button>
          <button
            {...homeProps('home.copy')}
            title={homeReason('home.copy') ?? words('Copy selected range')}
            onClick={() => void homeInvoke('home.copy')}
          >
            <Icon name="copy" size={15} />
            {words('Copy')}
          </button>
          <button
            {...homeProps('home.format-painter')}
            className={paintStyleId !== null ? 'active' : ''}
            title={
              homeReason('home.format-painter') ?? words('Copy selected cell formatting, then click a destination cell')
            }
            disabled={homeDisabled('home.format-painter') || readOnly || busy || literal}
            onClick={() => void homeInvoke('home.format-painter')}
          >
            <Icon name="paint" size={16} />
            {words('Format Painter')}
          </button>
        </div>
      </div>
      <small>{words('Clipboard')}</small>
    </RibbonGroup>
  );
}
