import type { GridTables } from './gridTypes';
import { RetainedWorkbookTool } from './RetainedWorkbookTool';
import type { WorkbookCommand, WorkbookSheet } from '../../core/contracts';
import type { WorkbookTool } from '../SpreadsheetRibbon';

export interface NameManagerPaneInputs {
  activeRange: () => string;
  activeTool: WorkbookTool | null;
  busy: boolean;
  goToName: (formula: string) => void;
  literal: boolean | undefined;
  nameCommands: boolean;
  namedRanges: { name: string; formula: string; sheet?: string; scopeInvalid?: true }[];
  namePage: number;
  nameScope: 'workbook' | 'worksheet';
  navigateMetadataRange: (name: string, target: string) => void;
  newRangeName: string;
  readOnly: boolean;
  send: (commands: WorkbookCommand[], allowSpecial?: boolean) => void;
  setNamePage: React.Dispatch<React.SetStateAction<number>>;
  setNameScope: React.Dispatch<React.SetStateAction<'workbook' | 'worksheet'>>;
  setNewRangeName: React.Dispatch<React.SetStateAction<string>>;
  setTablePage: React.Dispatch<React.SetStateAction<number>>;
  sheet: WorkbookSheet;
  tablePage: number;
  tables: GridTables;
  visitedTools: WorkbookTool[];
  words: (en: string) => string;
}

export function NameManagerPane({
  activeRange,
  activeTool,
  busy,
  goToName,
  literal,
  nameCommands,
  namedRanges,
  namePage,
  nameScope,
  navigateMetadataRange,
  newRangeName,
  readOnly,
  send,
  setNamePage,
  setNameScope,
  setNewRangeName,
  setTablePage,
  sheet,
  tablePage,
  tables,
  visitedTools,
  words,
}: NameManagerPaneInputs): JSX.Element {
  return (
    <>
      {visitedTools.includes('names') && !literal && (
        <RetainedWorkbookTool tool="names" visible={activeTool === 'names'}>
          <section className="sheet-name-manager" aria-label={words('Names and tables')}>
            <form
              onSubmit={(event) => {
                event.preventDefault();
                if (!nameCommands || readOnly || busy) return;
                send([
                  {
                    operation: 'defineName',
                    sheet: sheet.name,
                    name: newRangeName.trim(),
                    range: activeRange(),
                    nameScope,
                    expectedNameFormula: null,
                  },
                ]);
              }}
            >
              <label>
                {words('Range name')}
                <input
                  aria-label={words('Range name')}
                  maxLength={255}
                  value={newRangeName}
                  onChange={(event) => setNewRangeName(event.target.value)}
                />
              </label>
              <label>
                {words('Name scope')}
                <select
                  aria-label={words('Name scope')}
                  value={nameScope}
                  onChange={(event) => setNameScope(event.target.value as typeof nameScope)}
                >
                  <option value="workbook">{words('Workbook')}</option>
                  <option value="worksheet">{sheet.name}</option>
                </select>
              </label>
              <span>{activeRange()}</span>
              <button disabled={readOnly || busy || !nameCommands || !newRangeName.trim()}>
                {words('Create range name')}
              </button>
            </form>
            {!nameCommands && <small>{words('Named-range authoring requires an indexed XLSX owner.')}</small>}
            <div aria-label={words('Named ranges')}>
              {namedRanges.slice(namePage * 100, (namePage + 1) * 100).map((item, index) => (
                <div key={`${item.scopeInvalid ? 'invalid' : (item.sheet ?? '')}\0${item.name}\0${index}`}>
                  <strong>{item.name}</strong>{' '}
                  <span>
                    {item.scopeInvalid ? words('Invalid worksheet scope') : (item.sheet ?? words('Workbook'))} ·{' '}
                    {item.formula}
                  </span>{' '}
                  <button disabled={busy} onClick={() => goToName(item.formula)}>
                    {words('Go to name')} {item.name}
                  </button>
                  <button
                    disabled={
                      readOnly ||
                      busy ||
                      !nameCommands ||
                      item.scopeInvalid ||
                      /^_xlnm\./i.test(item.name) ||
                      (!!item.sheet && item.sheet !== sheet.name)
                    }
                    title={
                      item.sheet && item.sheet !== sheet.name
                        ? words('Switch to ' + item.sheet + ' to use its worksheet-local name selection.')
                        : undefined
                    }
                    aria-description={
                      item.sheet && item.sheet !== sheet.name
                        ? words('Switch to ' + item.sheet + ' to use its worksheet-local name selection.')
                        : undefined
                    }
                    onClick={() => {
                      if (item.scopeInvalid || (item.sheet && item.sheet !== sheet.name)) return;
                      send([
                        {
                          operation: 'updateName',
                          sheet: item.sheet ?? sheet.name,
                          name: item.name,
                          range: activeRange(),
                          nameScope: item.sheet ? 'worksheet' : 'workbook',
                          expectedNameFormula: item.formula,
                        },
                      ]);
                    }}
                  >
                    {words('Use selection for name')} {item.name}
                  </button>
                  <button
                    disabled={readOnly || busy || !nameCommands || item.scopeInvalid || /^_xlnm\./i.test(item.name)}
                    onClick={() => {
                      if (item.scopeInvalid) return;
                      send([
                        {
                          operation: 'removeName',
                          sheet: item.sheet ?? sheet.name,
                          name: item.name,
                          nameScope: item.sheet ? 'worksheet' : 'workbook',
                          expectedNameFormula: item.formula,
                        },
                      ]);
                    }}
                  >
                    {words('Remove name')} {item.name}
                  </button>
                </div>
              ))}
            </div>
            <button disabled={namePage === 0} onClick={() => setNamePage((page) => page - 1)}>
              {words('Previous names')}
            </button>
            <span>
              {namedRanges.length ? Math.min(namePage * 100 + 1, namedRanges.length) : 0}–
              {Math.min((namePage + 1) * 100, namedRanges.length)} / {namedRanges.length}
            </span>
            <button
              disabled={(namePage + 1) * 100 >= namedRanges.length}
              onClick={() => setNamePage((page) => page + 1)}
            >
              {words('More names')}
            </button>
            <div aria-label={words('Workbook tables')}>
              {tables.slice(tablePage * 100, (tablePage + 1) * 100).map((item) => (
                <div key={item.name}>
                  <strong>{item.name}</strong>{' '}
                  <span>
                    {item.sheet}!{item.ref} ·{' '}
                    {item.columns
                      .slice(0, 20)
                      .map((column) => column.name)
                      .join(', ')}
                    {item.columns.length > 20 ? '…' : ''}
                  </span>
                  <button disabled={busy} onClick={() => navigateMetadataRange(item.sheet, item.ref)}>
                    {words('Go to table')} {item.name}
                  </button>
                </div>
              ))}
            </div>
            <button disabled={tablePage === 0} onClick={() => setTablePage((page) => page - 1)}>
              {words('Previous tables')}
            </button>
            <span>
              {tables.length ? Math.min(tablePage * 100 + 1, tables.length) : 0}–
              {Math.min((tablePage + 1) * 100, tables.length)} / {tables.length}
            </span>
            <button disabled={(tablePage + 1) * 100 >= tables.length} onClick={() => setTablePage((page) => page + 1)}>
              {words('More tables')}
            </button>
            <small>
              {words(
                'Static name ranges can be edited, and unused names can be removed. Names used in formulas cannot be removed. Dynamic names and unsupported object references stay protected.'
              )}
            </small>
          </section>
        </RetainedWorkbookTool>
      )}
    </>
  );
}
