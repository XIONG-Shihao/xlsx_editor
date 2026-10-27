import type { GridHomeInvoke, GridHomeProps } from '../gridTypes';
import type { HomeCommandId } from '../../homeCommands';
import { Icon } from '../../host';
import { RibbonGroup } from '../../SpreadsheetRibbon';
import type { WorkbookCell, WorkbookNormalFont } from '../../../core/contracts';
import { homeFontSamples } from '../../homeFormattingSamples';

export interface HomeFontGroupInputs {
  busy: boolean;
  homeDisabled: (id: HomeCommandId) => boolean;
  homeInvoke: GridHomeInvoke;
  homeProps: GridHomeProps;
  homeReason: (id: HomeCommandId) => string | undefined;
  literal: boolean | undefined;
  mixedStyle: (key: keyof WorkbookCell) => boolean;
  normalFont: Readonly<WorkbookNormalFont>;
  readOnly: boolean;
  styleCell: WorkbookCell;
  words: (en: string) => string;
}

export function HomeFontGroup({
  busy,
  homeDisabled,
  homeInvoke,
  homeProps,
  homeReason,
  literal,
  mixedStyle,
  normalFont,
  readOnly,
  styleCell,
  words,
}: HomeFontGroupInputs): JSX.Element {
  return (
    <RibbonGroup name="Font" icon="font" collapseBelow={620}>
      <div className="sheet-ribbon-body sheet-ribbon-columns">
        <div className="sheet-ribbon-row">
          <select
            {...homeProps('home.font-family')}
            aria-label={words('Font family')}
            disabled={homeDisabled('home.font-family') || readOnly || busy || literal}
            value={mixedStyle('fontFamily') ? '' : (styleCell?.fontFamily ?? normalFont.fontFamily)}
            onChange={(event) => void homeInvoke('home.font-family', { value: event.target.value })}
          >
            {mixedStyle('fontFamily') && <option value="">{words('Mixed')}</option>}
            {[
              ...new Set([
                styleCell?.fontFamily ?? normalFont.fontFamily,
                ...homeFontSamples,
                'Aptos Narrow',
                'Calibri',
                'Times New Roman',
              ]),
            ].map((name) => (
              <option key={name}>{name}</option>
            ))}
          </select>
          <select
            {...homeProps('home.font-size')}
            aria-label={words('Font size')}
            disabled={homeDisabled('home.font-size') || readOnly || busy || literal}
            value={mixedStyle('fontSize') ? '' : (styleCell?.fontSize ?? normalFont.fontSize)}
            onChange={(event) => void homeInvoke('home.font-size', { value: Number(event.target.value) })}
          >
            {mixedStyle('fontSize') && <option value="">{words('Mixed')}</option>}
            {[...new Set([styleCell?.fontSize ?? normalFont.fontSize, 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36])]
              .sort((a, b) => a - b)
              .map((size) => (
                <option key={size} value={size}>
                  {size}
                </option>
              ))}
          </select>
          <button
            {...homeProps('home.font-size-increase')}
            aria-label={words('Increase font size')}
            title={homeReason('home.font-size-increase') ?? words('Increase font size')}
            disabled={homeDisabled('home.font-size-increase') || readOnly || busy || literal}
            onClick={() => void homeInvoke('home.font-size-increase')}
          >
            <Icon name="fontIncrease" size={18} />
          </button>
          <button
            {...homeProps('home.font-size-decrease')}
            aria-label={words('Decrease font size')}
            title={homeReason('home.font-size-decrease') ?? words('Decrease font size')}
            disabled={homeDisabled('home.font-size-decrease') || readOnly || busy || literal}
            onClick={() => void homeInvoke('home.font-size-decrease')}
          >
            <Icon name="fontDecrease" size={18} />
          </button>
        </div>
        <div className="sheet-ribbon-row">
          <button
            {...homeProps('home.bold')}
            className={mixedStyle('bold') ? 'mixed' : (styleCell?.bold ?? normalFont.bold) ? 'active' : ''}
            aria-pressed={mixedStyle('bold') ? 'mixed' : !!(styleCell?.bold ?? normalFont.bold)}
            aria-label={words('Bold')}
            title={homeReason('home.bold') ?? words('Bold')}
            disabled={homeDisabled('home.bold') || readOnly || busy || literal}
            onClick={() => void homeInvoke('home.bold', { value: !(styleCell?.bold ?? normalFont.bold) })}
          >
            <b>B</b>
          </button>
          <button
            {...homeProps('home.italic')}
            className={mixedStyle('italic') ? 'mixed' : (styleCell?.italic ?? normalFont.italic) ? 'active' : ''}
            aria-pressed={mixedStyle('italic') ? 'mixed' : !!(styleCell?.italic ?? normalFont.italic)}
            aria-label={words('Italic')}
            title={homeReason('home.italic') ?? words('Italic')}
            disabled={homeDisabled('home.italic') || readOnly || busy || literal}
            onClick={() => void homeInvoke('home.italic', { value: !(styleCell?.italic ?? normalFont.italic) })}
          >
            <i>I</i>
          </button>
          <button
            {...homeProps('home.underline')}
            className={mixedStyle('underline') ? 'mixed' : styleCell?.underline ? 'active' : ''}
            aria-pressed={mixedStyle('underline') ? 'mixed' : !!styleCell?.underline}
            aria-label={words('Underline')}
            title={homeReason('home.underline') ?? words('Underline')}
            disabled={homeDisabled('home.underline') || readOnly || busy || literal}
            onClick={() => void homeInvoke('home.underline', { value: !styleCell?.underline })}
          >
            <u>U</u>
          </button>
          <button
            {...homeProps('home.borders')}
            className={mixedStyle('border') ? 'mixed' : styleCell?.border === 'all' ? 'active' : ''}
            aria-pressed={mixedStyle('border') ? 'mixed' : styleCell?.border === 'all'}
            aria-label={words('Borders')}
            title={homeReason('home.borders') ?? words('All borders')}
            disabled={homeDisabled('home.borders') || readOnly || busy || literal}
            onClick={() => void homeInvoke('home.borders', { value: styleCell?.border === 'all' ? 'none' : 'all' })}
          >
            <Icon name="border" size={17} />
          </button>
          <label
            className={mixedStyle('fillColor') ? 'sheet-color-control mixed' : 'sheet-color-control'}
            title={mixedStyle('fillColor') ? words('Mixed fill colors') : words('Fill color')}
            aria-label={words('Fill color')}
          >
            <Icon name="fill" size={17} />
            <span className="sheet-color-swatch" style={{ background: `#${styleCell?.fillColor ?? 'FFFF00'}` }} />
            <input
              {...homeProps('home.fill-color')}
              type="color"
              aria-label={words('Fill color')}
              value={`#${styleCell?.fillColor ?? 'FFFF00'}`}
              disabled={homeDisabled('home.fill-color') || readOnly || busy || literal}
              onChange={(event) => void homeInvoke('home.fill-color', { value: event.target.value.slice(1) })}
            />
          </label>
          <label
            className={mixedStyle('fontColor') ? 'sheet-color-control mixed' : 'sheet-color-control'}
            title={mixedStyle('fontColor') ? words('Mixed font colors') : words('Font color')}
            aria-label={words('Font color')}
          >
            <span className="sheet-font-color-letter">A</span>
            <span className="sheet-color-swatch" style={{ background: `#${styleCell?.fontColor ?? '000000'}` }} />
            <input
              {...homeProps('home.font-color')}
              type="color"
              aria-label={words('Font color')}
              value={`#${styleCell?.fontColor ?? '000000'}`}
              disabled={homeDisabled('home.font-color') || readOnly || busy || literal}
              onChange={(event) => void homeInvoke('home.font-color', { value: event.target.value.slice(1) })}
            />
          </label>
        </div>
      </div>
      <small>{words('Font')}</small>
    </RibbonGroup>
  );
}
