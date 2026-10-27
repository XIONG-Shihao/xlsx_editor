import type { GridConditionalControls2, GridHomeInvoke, GridHomeProps } from '../gridTypes';
import type { HomeCommandId } from '../../homeCommands';
import { Icon } from '../../host';
import { RibbonGroup, type WorkbookTool } from '../../SpreadsheetRibbon';

export interface HomeStylesGroupInputs {
  active: boolean;
  activeTool: WorkbookTool | null;
  busy: boolean;
  conditionalColor: string;
  conditionalControls: GridConditionalControls2;
  conditionalRule: 'greaterThan' | 'lessThan' | 'duplicateValues';
  conditionalThreshold: string;
  homeDisabled: (id: HomeCommandId) => boolean;
  homeInvoke: GridHomeInvoke;
  homeMenu: (event: React.MouseEvent, menu: string) => void;
  homeProps: GridHomeProps;
  literal: boolean | undefined;
  readOnly: boolean;
  tableName: string;
  taskPaneId: string;
  words: (en: string) => string;
}

export function HomeStylesGroup({
  active,
  activeTool,
  busy,
  conditionalColor,
  conditionalControls,
  conditionalRule,
  conditionalThreshold,
  homeDisabled,
  homeInvoke,
  homeMenu,
  homeProps,
  literal,
  readOnly,
  tableName,
  taskPaneId,
  words,
}: HomeStylesGroupInputs): JSX.Element {
  return (
    <RibbonGroup name="Styles" icon="cellStyle" collapseBelow={1380}>
      <div className="sheet-ribbon-body sheet-ribbon-stack">
        {conditionalControls && !literal ? (
          <button
            {...homeProps('styles.conditional')}
            type="button"
            data-workbook-tool="conditional"
            aria-label={words('Conditional formatting')}
            aria-pressed={activeTool === 'conditional'}
            aria-controls={taskPaneId}
            disabled={homeDisabled('styles.conditional') || !active}
            data-ribbon-command="styles.conditional"
            onClick={() => void homeInvoke('styles.conditional')}
          >
            <Icon name="conditional" size={17} />
            {words('Conditional formatting')}
          </button>
        ) : (
          <details className="sheet-ribbon-menu">
            <summary
              data-home-command="home.menu"
              aria-disabled={homeDisabled('home.menu')}
              onClick={(event) => homeMenu(event, 'conditional')}
            >
              <Icon name="conditional" size={17} />
              {words('Conditional Formatting')} <Icon name="chevron" size={11} />
            </summary>
            <div>
              <select
                {...homeProps('home.conditional-draft')}
                aria-label={words('Conditional rule')}
                value={conditionalRule}
                onChange={(event) =>
                  void homeInvoke('home.conditional-draft', { field: 'rule', value: event.target.value })
                }
              >
                <option value="greaterThan">{words('Greater than')}</option>
                <option value="lessThan">{words('Less than')}</option>
                <option value="duplicateValues">{words('Duplicate values')}</option>
              </select>
              {conditionalRule !== 'duplicateValues' && (
                <label>
                  {words('Threshold')}
                  <input
                    {...homeProps('home.conditional-draft')}
                    type="number"
                    aria-label={words('Conditional threshold')}
                    value={conditionalThreshold}
                    onChange={(event) =>
                      void homeInvoke('home.conditional-draft', {
                        field: 'threshold',
                        value: event.target.value,
                      })
                    }
                  />
                </label>
              )}
              <label>
                {words('Highlight color')}
                <input
                  {...homeProps('home.conditional-draft')}
                  type="color"
                  aria-label={words('Conditional highlight color')}
                  value={`#${conditionalColor}`}
                  onChange={(event) =>
                    void homeInvoke('home.conditional-draft', {
                      field: 'fillColor',
                      value: event.target.value.slice(1).toUpperCase(),
                    })
                  }
                />
              </label>
              <button
                {...homeProps('home.conditional-apply')}
                disabled={homeDisabled('home.conditional-apply') || readOnly || busy || literal}
                onClick={() =>
                  void homeInvoke('home.conditional-apply', {
                    rule: conditionalRule,
                    ...(conditionalRule === 'duplicateValues' ? {} : { threshold: Number(conditionalThreshold) }),
                    fillColor: conditionalColor,
                  })
                }
              >
                {words('Apply rule')}
              </button>
            </div>
          </details>
        )}
        <details className="sheet-ribbon-menu">
          <summary
            data-home-command="home.menu"
            aria-disabled={homeDisabled('home.menu')}
            onClick={(event) => homeMenu(event, 'table')}
          >
            <Icon name="table" size={17} />
            {words('Format as Table')} <Icon name="chevron" size={11} />
          </summary>
          <div>
            <label>
              {words('Table name')}
              <input
                {...homeProps('home.table-draft')}
                aria-label={words('Table name')}
                value={tableName}
                onChange={(event) => void homeInvoke('home.table-draft', { name: event.target.value })}
              />
            </label>
            <small>{words('First row must contain unique text headers. Uses a blue striped table style.')}</small>
            <button
              {...homeProps('home.create-table')}
              disabled={homeDisabled('home.create-table') || readOnly || busy || literal}
              onClick={() => void homeInvoke('home.create-table', { name: tableName })}
            >
              {words('Create table')}
            </button>
          </div>
        </details>
        <details className="sheet-ribbon-menu">
          <summary
            data-home-command="home.menu"
            aria-disabled={homeDisabled('home.menu')}
            onClick={(event) => homeMenu(event, 'styles')}
          >
            <Icon name="cellStyle" size={17} />
            {words('Cell Styles')} <Icon name="chevron" size={11} />
          </summary>
          <div>
            <button
              {...homeProps('home.style-good')}
              disabled={homeDisabled('home.style-good') || readOnly || busy || literal}
              onClick={() => void homeInvoke('home.style-good')}
            >
              <span className="sheet-style-sample good">{words('Good')}</span>
            </button>
            <button
              {...homeProps('home.style-bad')}
              disabled={homeDisabled('home.style-bad') || readOnly || busy || literal}
              onClick={() => void homeInvoke('home.style-bad')}
            >
              <span className="sheet-style-sample bad">{words('Bad')}</span>
            </button>
            <button
              {...homeProps('home.style-heading')}
              disabled={homeDisabled('home.style-heading') || readOnly || busy || literal}
              onClick={() => void homeInvoke('home.style-heading')}
            >
              <span className="sheet-style-sample heading">{words('Heading')}</span>
            </button>
          </div>
        </details>
      </div>
      <small>{words('Styles')}</small>
    </RibbonGroup>
  );
}
