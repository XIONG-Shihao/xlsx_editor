import type { GridHomeInvoke, GridHomeProps } from '../gridTypes';
import type { HomeCommandId } from '../../homeCommands';
import { Icon } from '../../host';
import { RibbonGroup } from '../../SpreadsheetRibbon';
import type { WorkbookCell, WorkbookEdit } from '../../../core/contracts';
import { editableNumberFormat } from '../editableFormats';

export interface HomeNumberGroupInputs {
  busy: boolean;
  homeDisabled: (id: HomeCommandId) => boolean;
  homeInvoke: GridHomeInvoke;
  homeProps: GridHomeProps;
  homeReason: (id: HomeCommandId) => string | undefined;
  literal: boolean | undefined;
  mixedStyle: (key: keyof WorkbookCell) => boolean;
  readOnly: boolean;
  styleCell: WorkbookCell;
  words: (en: string) => string;
}

export function HomeNumberGroup({
  busy,
  homeDisabled,
  homeInvoke,
  homeProps,
  homeReason,
  literal,
  mixedStyle,
  readOnly,
  styleCell,
  words,
}: HomeNumberGroupInputs): JSX.Element {
  return (
    <RibbonGroup name="Number" icon="number" collapseBelow={840}>
      <div className="sheet-ribbon-body sheet-ribbon-columns">
        <select
          {...homeProps('home.number-format')}
          aria-label={words('Number format')}
          disabled={homeDisabled('home.number-format') || readOnly || busy || literal}
          value={mixedStyle('format') ? 'mixed' : (editableNumberFormat(styleCell?.format) ?? 'general')}
          onChange={(event) =>
            void homeInvoke('home.number-format', {
              value: event.target.value as NonNullable<WorkbookEdit['numberFormat']>,
            })
          }
        >
          {mixedStyle('format') && <option value="mixed">{words('Mixed')}</option>}
          {[
            'general',
            'number',
            'percent',
            'date',
            'text',
            'currency',
            'accounting',
            'thousands',
            'decimal0',
            'decimal1',
            'decimal2',
          ].map((format) => (
            <option key={format} value={format}>
              {words(
                (
                  {
                    general: 'General',
                    number: 'Number',
                    percent: 'Percentage',
                    date: 'Date',
                    text: 'Text',
                    currency: 'Currency',
                    accounting: 'Accounting',
                    thousands: 'Thousands separator',
                    decimal0: '0 decimals',
                    decimal1: '1 decimal',
                    decimal2: '2 decimals',
                  } as Record<string, string>
                )[format]
              )}
            </option>
          ))}
        </select>
        <div className="sheet-ribbon-row">
          <button
            {...homeProps('home.accounting')}
            disabled={homeDisabled('home.accounting') || readOnly || busy || literal}
            title={homeReason('home.accounting') ?? words('Accounting')}
            onClick={() => void homeInvoke('home.accounting')}
          >
            $
          </button>
          <button
            {...homeProps('home.percent')}
            disabled={homeDisabled('home.percent') || readOnly || busy || literal}
            title={homeReason('home.percent') ?? words('Percent')}
            onClick={() => void homeInvoke('home.percent')}
          >
            %
          </button>
          <button
            {...homeProps('home.thousands')}
            disabled={homeDisabled('home.thousands') || readOnly || busy || literal}
            title={homeReason('home.thousands') ?? words('Thousands separator')}
            onClick={() => void homeInvoke('home.thousands')}
          >
            ,
          </button>
          <button
            {...homeProps('home.decimals-decrease')}
            disabled={
              homeDisabled('home.decimals-decrease') ||
              readOnly ||
              busy ||
              literal ||
              ![undefined, 0, 1, 2, '0.0', '0.00'].includes(styleCell?.format)
            }
            title={homeReason('home.decimals-decrease') ?? words('Decrease decimals')}
            onClick={() => void homeInvoke('home.decimals-decrease')}
          >
            <Icon name="decimalDecrease" size={19} />
          </button>
          <button
            {...homeProps('home.decimals-increase')}
            disabled={
              homeDisabled('home.decimals-increase') ||
              readOnly ||
              busy ||
              literal ||
              ![undefined, 0, 1, 2, '0.0', '0.00'].includes(styleCell?.format)
            }
            title={homeReason('home.decimals-increase') ?? words('Increase decimals')}
            onClick={() => void homeInvoke('home.decimals-increase')}
          >
            <Icon name="decimalIncrease" size={19} />
          </button>
        </div>
      </div>
      <small>{words('Number')}</small>
    </RibbonGroup>
  );
}
