import type { GridPrintControls3 } from './gridTypes';
import type { PrintFocusRequest } from '../WorkbookPrint';
import type { WorkbookGeometryProfile, WorkbookSheet } from '../../core/contracts';
import { useAcceptedPrintPreview } from '../useAcceptedPrintPreview';
import { useLayoutEffect, useRef, useState } from 'react';

export interface PrintAndViewModeInputs {
  activationEpoch: React.MutableRefObject<number>;
  active: boolean;
  documentId: string | undefined;
  fallbackDocumentId: string;
  geometryKey: string;
  geometryProfile: WorkbookGeometryProfile;
  gridScroll: React.RefObject<HTMLDivElement>;
  printControls: GridPrintControls3;
  sheet: WorkbookSheet;
  storageKey: string | undefined;
  viewMode: 'normal' | 'pageLayout' | 'pageBreak';
}

export function usePrintAndViewMode({
  activationEpoch,
  active,
  documentId,
  fallbackDocumentId,
  geometryKey,
  geometryProfile,
  gridScroll,
  printControls,
  sheet,
  storageKey,
  viewMode,
}: PrintAndViewModeInputs) {
  useLayoutEffect(() => {
    if (gridScroll.current) gridScroll.current.inert = viewMode === 'pageLayout';
  }, [viewMode]);
  const [printFocus, setPrintFocus] = useState<PrintFocusRequest | null>(null),
    printFocusSerial = useRef(0);
  const printScope = {
    documentId: printControls?.documentId ?? documentId ?? fallbackDocumentId,
    grantDocumentId: documentId ?? fallbackDocumentId,
    state: printControls?.revision ?? storageKey ?? 'local',
    sheet: sheet?.name ?? '',
    ownerKey: storageKey ?? 'local',
    activationEpoch: activationEpoch.current,
    active,
    geometryKey,
  };
  const printPreview = useAcceptedPrintPreview({
    scope: printScope,
    callbacks: printControls
      ? {
          preview: (settings, id) => printControls.preview(sheet?.name ?? '', settings, id, geometryProfile),
          page: printControls.page,
          release: printControls.release,
          cancel: printControls.cancel,
          publish: printControls.publish,
        }
      : null,
  });
  const normalViewportKey = JSON.stringify([
    printScope.documentId,
    printScope.grantDocumentId,
    printScope.state,
    printScope.sheet,
    printScope.ownerKey,
    printScope.activationEpoch,
    active,
    geometryKey,
  ]);
  const normalViewportBookmark = useRef<{ key: string; left: number; top: number; zoom: number } | null>(null),
    normalReturnPending = useRef(false);
  return {
    printFocus,
    setPrintFocus,
    printFocusSerial,
    printScope,
    printPreview,
    normalViewportKey,
    normalViewportBookmark,
    normalReturnPending,
  } as const;
}
