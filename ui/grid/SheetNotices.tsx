import type { AcceptedPrintPreview } from '../useAcceptedPrintPreview';
import type { GridCalculationController, GridSetViewport } from './gridTypes';
import type { PageBoundaryModel } from '../pageBoundaries';
import { PageBoundaryToolbar } from '../WorkbookPageBreakPreview';
import { type WorkbookCalculationControls, calculationPendingObjectCount } from '../calculationJobs';
import { WorkbookCalculationStatus } from '../WorkbookCalculation';
import type { WorkbookGeometryProfile, WorkbookSheet } from '../../core/contracts';

export interface SheetNoticesInputs {
  active: boolean;
  boundaryModel: PageBoundaryModel | null;
  calculationController: GridCalculationController;
  calculationControls: WorkbookCalculationControls | undefined;
  calculationKey: string;
  calculationNotice: { key: string; phase: 'running' | 'completed' | 'failed'; text: string } | null;
  columnOffsets: Float64Array<ArrayBufferLike>;
  geometryProfile: WorkbookGeometryProfile;
  gridError: string;
  gridScroll: React.RefObject<HTMLDivElement>;
  literal: boolean | undefined;
  normalMode: React.MutableRefObject<boolean>;
  passiveNotice: string | undefined;
  printPreview: AcceptedPrintPreview;
  returnNormal: () => void;
  rowOffsets: Float64Array<ArrayBufferLike>;
  selectedFontProfile: WorkbookGeometryProfile;
  setViewNotice: React.Dispatch<React.SetStateAction<string>>;
  setViewport: GridSetViewport;
  sheet: WorkbookSheet;
  specialSelection: Set<string> | null;
  viewMode: 'normal' | 'pageLayout' | 'pageBreak';
  viewNotice: string;
  words: (en: string) => string;
}

export function SheetNotices({
  active,
  boundaryModel,
  calculationController,
  calculationControls,
  calculationKey,
  calculationNotice,
  columnOffsets,
  geometryProfile,
  gridError,
  gridScroll,
  literal,
  normalMode,
  passiveNotice,
  printPreview,
  returnNormal,
  rowOffsets,
  selectedFontProfile,
  setViewNotice,
  setViewport,
  sheet,
  specialSelection,
  viewMode,
  viewNotice,
  words,
}: SheetNoticesInputs): JSX.Element {
  return (
    <>
      {gridError && (
        <div className="sheet-grid-error" role="alert">
          {gridError}
        </div>
      )}
      {calculationControls && <WorkbookCalculationStatus controller={calculationController} words={words} />}
      {calculationControls && !literal && !passiveNotice && active && (
        <div
          className="sheet-view-notice"
          role="status"
          aria-label={words('Workbook calculation policy')}
          data-calculation-mode={calculationControls.policy.mode}
          data-pending-formulas={calculationControls.policy.pendingFormulaCells}
          data-pending-cache-objects={calculationPendingObjectCount(calculationControls.policy)}
        >
          {words(calculationControls.policy.mode === 'manual' ? 'Manual calculation' : 'Automatic calculation')} ·{' '}
          {calculationControls.policy.pendingFormulaCells} {words('formulas pending calculation')}
          {!!calculationPendingObjectCount(calculationControls.policy) && (
            <>
              {' '}
              · {calculationPendingObjectCount(calculationControls.policy)}{' '}
              {words('chart/pivot caches pending calculation')}
            </>
          )}
        </div>
      )}
      {!calculationControls &&
        calculationNotice &&
        active &&
        normalMode.current &&
        calculationNotice.key === calculationKey && (
          <div
            className="sheet-view-notice"
            role="status"
            aria-label={words('Calculation status')}
            data-calculation-phase={calculationNotice.phase}
          >
            {words(calculationNotice.text)}
          </div>
        )}
      {!literal && active && (
        <div
          className="sheet-view-notice"
          role="status"
          aria-label={words('Selected cell font rendering')}
          data-requested-font={selectedFontProfile.font.fontFamily}
          data-font-source={selectedFontProfile.source}
          data-font-availability={selectedFontProfile.availability}
        >
          {selectedFontProfile.font.fontFamily}:{' '}
          {words(
            selectedFontProfile.source === 'substituted-font'
              ? 'This font is unavailable here; the selected cell uses a local substitute.'
              : selectedFontProfile.source === 'reference-fallback'
                ? 'Local font measurements are unavailable; a substitute may be used.'
                : 'The selected cell uses local font rendering; installed-family availability has not been independently verified.'
          )}
        </div>
      )}
      {!literal && active && (
        <div
          className="sheet-view-notice"
          role="status"
          aria-label={words('Worksheet font rendering')}
          data-geometry-font-source={geometryProfile.source}
          data-geometry-font-availability={geometryProfile.availability}
        >
          {words(
            geometryProfile.source === 'reference-fallback'
              ? 'Worksheet dimensions use reference font sizes because local font measurements are unavailable.'
              : geometryProfile.source === 'substituted-font'
                ? 'The worksheet font is unavailable here; dimensions use a local substitute.'
                : 'Worksheet dimensions use local font rendering; the requested font has not been independently verified.'
          )}
          {sheet.geometryIssues?.length ? (
            <> {words('Some source dimensions could not be read; display defaults are used.')}</>
          ) : null}
        </div>
      )}
      {viewNotice && (
        <div className="sheet-view-notice" role="status">
          {words(viewNotice)}{' '}
          <button type="button" aria-label={words('Dismiss view message')} onClick={() => setViewNotice('')}>
            ×
          </button>
        </div>
      )}
      {specialSelection && (
        <div className="sheet-special-status" role="status">
          {specialSelection.size} {words('matching cells selected')} ·{' '}
          {words('Format or Clear those cells; click a cell to exit.')}
        </div>
      )}
      {viewMode === 'pageBreak' && (
        <PageBoundaryToolbar
          preview={printPreview}
          model={boundaryModel}
          onNormal={returnNormal}
          frozen={!!(sheet.freeze?.rows || sheet.freeze?.columns)}
          words={words}
          onPage={(number) => {
            const page = boundaryModel?.pages.find((page) => page.number === number),
              scroll = gridScroll.current;
            if (!page || !scroll) return;
            scroll.scrollLeft = columnOffsets[page.firstColumn];
            scroll.scrollTop = rowOffsets[page.firstRow];
            setViewport({
              width: scroll.clientWidth,
              height: scroll.clientHeight,
              scrollLeft: scroll.scrollLeft,
              scrollTop: scroll.scrollTop,
            });
          }}
        />
      )}
    </>
  );
}
