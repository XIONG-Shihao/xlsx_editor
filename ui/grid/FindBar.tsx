import type { GridHomeInvoke, GridHomeProps } from './gridTypes';
import type { HomeCommandId } from '../homeCommands';
import type { SearchCursor } from '../searchPages';
import type { WorkbookTool } from '../SpreadsheetRibbon';

export interface FindBarInputs {
  active: boolean;
  activeTool: WorkbookTool | null;
  find: () => Promise<void>;
  findBusy: boolean;
  findIndex: number;
  findMatches: { sheet: string; cell: string }[];
  findMode: 'display' | 'raw' | 'formula';
  findNextCursor: SearchCursor | null;
  findOpen: boolean;
  findPageIndex: number;
  findPages: React.MutableRefObject<{ cursor: SearchCursor; offset: number }[]>;
  findScope: 'sheet' | 'workbook';
  findText: string;
  findTotal: number;
  homeDisabled: (id: HomeCommandId) => boolean;
  homeInvoke: GridHomeInvoke;
  homeProps: GridHomeProps;
  replacementDisabledReason: (action?: 'preview' | 'apply') => string | undefined;
  replaceText: string;
  taskPaneId: string;
  words: (en: string) => string;
}

export function FindBar({
  active,
  activeTool,
  find,
  findBusy,
  findIndex,
  findMatches,
  findMode,
  findNextCursor,
  findOpen,
  findPageIndex,
  findPages,
  findScope,
  findText,
  findTotal,
  homeDisabled,
  homeInvoke,
  homeProps,
  replacementDisabledReason,
  replaceText,
  taskPaneId,
  words,
}: FindBarInputs): JSX.Element {
  return (
    <>
      {findOpen && (
        <div className="sheet-find" role="search" aria-label={words('Find and replace')}>
          <input
            aria-label={words('Find')}
            value={findText}
            {...homeProps('home.find-draft')}
            onChange={(event) => void homeInvoke('home.find-draft', { field: 'text', value: event.target.value })}
            onKeyDown={(event) => {
              if (event.key === 'Enter') {
                event.preventDefault();
                void find();
              }
            }}
          />
          <select
            aria-label={words('Search scope')}
            value={findScope}
            {...homeProps('home.find-draft')}
            onChange={(event) => void homeInvoke('home.find-draft', { field: 'scope', value: event.target.value })}
          >
            <option value="sheet">{words('Current sheet')}</option>
            <option value="workbook">{words('All sheets')}</option>
          </select>
          <select
            aria-label={words('Search content')}
            value={findMode}
            {...homeProps('home.find-draft')}
            onChange={(event) => void homeInvoke('home.find-draft', { field: 'mode', value: event.target.value })}
          >
            <option value="display">{words('Displayed values')}</option>
            <option value="raw">{words('Raw values')}</option>
            <option value="formula">{words('Formulas')}</option>
          </select>
          <button
            {...homeProps('home.find-search')}
            onClick={() => void homeInvoke('home.find-search')}
            disabled={homeDisabled('home.find-search') || findBusy}
          >
            {findBusy ? words('Searching…') : words('Find')}
          </button>
          {findBusy && (
            <button {...homeProps('home.find-cancel')} onClick={() => void homeInvoke('home.find-cancel')}>
              {words('Cancel search')}
            </button>
          )}
          <span role="status">
            {findMatches.length
              ? `${(findPages.current[findPageIndex]?.offset ?? 0) + findIndex + 1} / ${findTotal}${findNextCursor ? '+' : ''}`
              : findBusy
                ? words('Searching…')
                : words('No matches')}
          </span>
          {findMatches.length > 0 && (
            <select
              aria-label={words('Found cell')}
              value={findIndex}
              {...homeProps('home.find-match')}
              disabled={homeDisabled('home.find-match') || findBusy}
              onChange={(event) => void homeInvoke('home.find-match', { index: Number(event.target.value) })}
            >
              {findMatches.map((match, index) => (
                <option key={`${match.sheet}\0${match.cell}`} value={index}>
                  {match.sheet}!{match.cell}
                </option>
              ))}
            </select>
          )}
          <button
            {...homeProps('home.find-next')}
            disabled={homeDisabled('home.find-next') || findBusy || !findMatches.length}
            onClick={() => void homeInvoke('home.find-next')}
          >
            {words('Next')}
          </button>
          <button
            {...homeProps('home.find-page')}
            disabled={homeDisabled('home.find-page') || findBusy || findPageIndex === 0}
            onClick={() => void homeInvoke('home.find-page', { direction: 'previous' })}
          >
            {words('Previous results')}
          </button>
          <button
            {...homeProps('home.find-page')}
            disabled={homeDisabled('home.find-page') || findBusy || !findNextCursor}
            onClick={() => void homeInvoke('home.find-page', { direction: 'next' })}
          >
            {words('More results')}
          </button>
          <input
            aria-label={words('Replace with')}
            value={replaceText}
            {...homeProps('home.find-draft')}
            onChange={(event) =>
              void homeInvoke('home.find-draft', { field: 'replacement', value: event.target.value })
            }
            disabled={!active}
          />
          <button
            {...homeProps('home.replace-preview')}
            data-workbook-tool="replace"
            aria-controls={taskPaneId}
            aria-pressed={activeTool === 'replace'}
            disabled={homeDisabled('home.replace-preview') || findBusy || !!replacementDisabledReason()}
            onClick={() => void homeInvoke('home.replace-preview')}
          >
            {words('Preview replacements')}
          </button>
          <button
            {...homeProps('home.replace-one')}
            disabled={homeDisabled('home.replace-one') || findBusy || !findMatches.length}
            onClick={() => void homeInvoke('home.replace-one')}
          >
            {words('Replace')}
          </button>
          <button
            {...homeProps('home.find-close')}
            aria-label={words('Close find')}
            onClick={() => void homeInvoke('home.find-close')}
          >
            ×
          </button>
        </div>
      )}
    </>
  );
}
