import type { AcceptedPrintPreview } from '../useAcceptedPrintPreview';
import type { SelectionSummary } from '../../core/indexed-session.mjs';
import type { WorkbookSheet } from '../../core/contracts';

export interface SheetTabsFooterInputs {
  activeFrame: (work: () => void) => number;
  busy: boolean;
  firstRenderedRow: number;
  lastRenderedRow: number;
  literal: boolean | undefined;
  newSheetName: string;
  printPreview: AcceptedPrintPreview;
  readOnly: boolean;
  scrollCell: (owner: WorkbookSheet, row: number, col: number, align?: boolean) => void;
  select: (cell: string, extend?: boolean) => void;
  selectionSummary: SelectionSummary | null;
  setNewSheetName: React.Dispatch<React.SetStateAction<string>>;
  setSheetAction: React.Dispatch<React.SetStateAction<'addSheet' | 'renameSheet' | 'duplicateSheet' | null>>;
  setSheetName: React.Dispatch<React.SetStateAction<string>>;
  sheet: WorkbookSheet;
  sheetAction: 'addSheet' | 'renameSheet' | 'duplicateSheet' | null;
  sheets: WorkbookSheet[];
  status: string | undefined;
  submitSheetAction: () => void;
  summaryError: string;
  viewMode: 'normal' | 'pageLayout' | 'pageBreak';
  visibleRows: number;
  words: (en: string) => string;
}

export function SheetTabsFooter({
  activeFrame,
  busy,
  firstRenderedRow,
  lastRenderedRow,
  literal,
  newSheetName,
  printPreview,
  readOnly,
  scrollCell,
  select,
  selectionSummary,
  setNewSheetName,
  setSheetAction,
  setSheetName,
  sheet,
  sheetAction,
  sheets,
  status,
  submitSheetAction,
  summaryError,
  viewMode,
  visibleRows,
  words,
}: SheetTabsFooterInputs): JSX.Element {
  return (
    <>
      <div className="sheet-tabs" aria-label={words('Worksheets')}>
        <button
          aria-label={words('Add worksheet')}
          title={words('Add worksheet')}
          disabled={readOnly || busy || literal}
          onClick={() => {
            setSheetAction('addSheet');
            setNewSheetName(`Sheet${sheets.length + 1}`);
          }}
        >
          ＋
        </button>
        {sheets
          .filter((item) => item.state === 'visible')
          .map((item) => (
            <button
              key={item.name}
              className={item.name === sheet.name ? 'active' : ''}
              onClick={() => {
                setSheetName(item.name);
                select('A1');
                activeFrame(() => scrollCell(item, 0, 0, true));
              }}
            >
              {item.name}
            </button>
          ))}
        <button
          disabled={readOnly || busy || literal}
          onClick={() => {
            setSheetAction('renameSheet');
            setNewSheetName(sheet.name);
          }}
        >
          {words('Rename')}
        </button>
        <button
          disabled={readOnly || busy || literal}
          onClick={() => {
            setSheetAction('duplicateSheet');
            setNewSheetName(`${sheet.name} Copy`.slice(0, 31));
          }}
        >
          {words('Duplicate')}
        </button>
        <select
          aria-label={words('Worksheet')}
          value={sheet.name}
          onChange={(event) => {
            const owner = sheets.find((item) => item.name === event.target.value);
            setSheetName(event.target.value);
            select('A1');
            if (owner) activeFrame(() => scrollCell(owner, 0, 0, true));
          }}
        >
          {sheets.map((item) => (
            <option key={item.name} value={item.name}>
              {item.name}
              {item.state !== 'visible' ? ' (hidden)' : ''}
            </option>
          ))}
        </select>
      </div>
      {sheetAction && (
        <form
          className="sheet-tab-action"
          onSubmit={(event) => {
            event.preventDefault();
            submitSheetAction();
          }}
        >
          <label>
            {sheetAction === 'addSheet'
              ? words('New worksheet name')
              : sheetAction === 'renameSheet'
                ? words('Rename worksheet')
                : words('Duplicate worksheet as')}
            <input
              autoFocus
              maxLength={31}
              value={newSheetName}
              onChange={(event) => setNewSheetName(event.target.value)}
            />
          </label>
          <button type="submit" disabled={readOnly || busy}>
            {words('Apply')}
          </button>
          <button type="button" onClick={() => setSheetAction(null)}>
            {words('Cancel')}
          </button>
        </form>
      )}
      <footer>
        {selectionSummary && (
          <span
            aria-label="Selection summary"
            title="Numeric statistics exclude text, booleans, errors and unresolved formulas."
          >
            Count: {selectionSummary.count} · Numbers: {selectionSummary.numericCount} · Sum:{' '}
            {selectionSummary.sum ?? 'Unavailable'} · Average: {selectionSummary.average ?? '—'}
            {!!(selectionSummary.errors + selectionSummary.unresolved + selectionSummary.unsafeNumbers) &&
              ` · Excluded: ${selectionSummary.errors + selectionSummary.unresolved + selectionSummary.unsafeNumbers}`}
          </span>
        )}
        {summaryError && <span title={summaryError}>Selection summary unavailable</span>}
        <span className="sheet-footer-status" title={status}>
          {status}
        </span>
        <span aria-label={words('View position')}>
          {viewMode === 'pageLayout' ? (
            printPreview.job ? (
              <>
                {words('Page')} {printPreview.number} {words('of')} {printPreview.job.pageCount}
              </>
            ) : (
              words('Page Layout preview')
            )
          ) : (
            <>
              {words('Rows')} {firstRenderedRow + 1}–{lastRenderedRow} {words('of')} {visibleRows}
            </>
          )}
        </span>
      </footer>
    </>
  );
}
