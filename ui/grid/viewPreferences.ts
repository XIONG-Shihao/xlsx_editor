import { DEFAULT_VIEW_PREFERENCES, type ViewPreferencesPatch, createViewPreferenceStore } from '../viewPreferences';
import { ViewPresetStore } from '../viewPresets';
import type { WorkbookSheet } from '../../core/contracts';
import { addressPosition } from '../sheetModel';
import { useLayoutEffect, useMemo, useRef, useState } from 'react';

export interface ViewPreferencesInputs {
  activationEpoch: React.MutableRefObject<number>;
  activeFrame: (work: () => void) => number;
  documentId: string | undefined;
  fallbackDocumentId: string;
  gridScroll: React.RefObject<HTMLDivElement>;
  navigationGeneration: React.MutableRefObject<number>;
  ownerIdentity: string | undefined;
  sheet: WorkbookSheet;
  viewActive: React.MutableRefObject<boolean>;
  viewPreferenceId: string | undefined;
  /** Read when the code runs, like the closed-over variables they stand for. */
  late: {
    readonly scrollCell: (owner: WorkbookSheet, row: number, col: number, align?: boolean) => void;
    readonly selectedRef: React.MutableRefObject<string>;
    readonly setGridError: React.Dispatch<React.SetStateAction<string>>;
    readonly sheetRef: React.MutableRefObject<WorkbookSheet>;
  };
}

export function useViewPreferences({
  activationEpoch,
  activeFrame,
  documentId,
  fallbackDocumentId,
  gridScroll,
  navigationGeneration,
  ownerIdentity,
  sheet,
  viewActive,
  viewPreferenceId,
  late,
}: ViewPreferencesInputs) {
  const preferenceDocumentId = viewPreferenceId ?? documentId;
  const preferenceStore = useMemo(() => {
    try {
      return createViewPreferenceStore(window.localStorage);
    } catch {
      return createViewPreferenceStore(null);
    }
  }, []);
  const presetStore = useMemo(() => {
    try {
      return new ViewPresetStore(window.localStorage);
    } catch {
      return new ViewPresetStore(null);
    }
  }, []);
  const [preferenceVersion, setPreferenceVersion] = useState(0);
  const preferenceBinding = useRef({ documentId, preferenceDocumentId });
  const previousPreference = preferenceBinding.current;
  const pendingPreferenceMove =
    !!documentId &&
    documentId === previousPreference.documentId &&
    !!preferenceDocumentId &&
    !!previousPreference.preferenceDocumentId &&
    preferenceDocumentId !== previousPreference.preferenceDocumentId;
  // Keep the accepted local view during first Save's transition render. The
  // complete sheet preference transaction runs before paint, never in render.
  const preferenceReadId = pendingPreferenceMove ? previousPreference.preferenceDocumentId : preferenceDocumentId;
  const view = useMemo(
    () =>
      sheet
        ? preferenceStore.get(preferenceReadId, sheet.name, sheet.sourceView ?? {})
        : { ...DEFAULT_VIEW_PREFERENCES },
    [preferenceStore, preferenceReadId, sheet?.name, sheet?.sourceView, preferenceVersion]
  );
  useLayoutEffect(() => {
    const prior = preferenceBinding.current;
    preferenceBinding.current = { documentId, preferenceDocumentId };
    if (prior.preferenceDocumentId === preferenceDocumentId) return;
    if (documentId && documentId === prior.documentId && prior.preferenceDocumentId && preferenceDocumentId) {
      preferenceStore.moveDocument(prior.preferenceDocumentId, preferenceDocumentId);
      presetStore.moveDocument(prior.preferenceDocumentId, preferenceDocumentId);
    }
    setPreferenceVersion((version) => version + 1);
  }, [documentId, preferenceDocumentId, preferenceStore, presetStore]);
  const [viewport, setViewport] = useState({ width: 1000, height: 600, scrollTop: 0, scrollLeft: 0 });
  useLayoutEffect(() => {
    const scroll = gridScroll.current;
    if (!scroll) return;
    const measure = () =>
      setViewport((previous) => {
        const next = {
          width: scroll.clientWidth,
          height: scroll.clientHeight,
          scrollTop: scroll.scrollTop,
          scrollLeft: scroll.scrollLeft,
        };
        return Object.keys(next).every(
          (key) => next[key as keyof typeof next] === previous[key as keyof typeof previous]
        )
          ? previous
          : next;
      });
    measure();
    const observer = new ResizeObserver(measure);
    observer.observe(scroll);
    return () => observer.disconnect();
  }, [!!sheet]);
  const changeView = (patch: ViewPreferencesPatch): void => {
    if (!viewActive.current || !sheet) return;
    try {
      preferenceStore.patch(preferenceDocumentId, sheet.name, patch, sheet.sourceView ?? {});
      setPreferenceVersion((version) => version + 1);
      const navigation = navigationGeneration.current;
      if (patch.zoom !== undefined || patch.headings !== undefined)
        activeFrame(() => {
          if (navigationGeneration.current !== navigation) return;
          const position = addressPosition(late.selectedRef.current);
          late.scrollCell(late.sheetRef.current, position.row, position.col);
        });
    } catch (error) {
      late.setGridError(String(error instanceof Error ? error.message : error));
    }
  };
  const presetDocumentId = preferenceReadId ?? fallbackDocumentId;
  const namedViewPresets = useMemo(
    () => (sheet ? presetStore.list(presetDocumentId, sheet.name) : []),
    [presetStore, presetDocumentId, sheet?.name, preferenceVersion]
  );
  const presetBindingKey = JSON.stringify([
    presetDocumentId,
    sheet?.name,
    documentId,
    ownerIdentity,
    activationEpoch.current,
  ]);
  const presetLive = useRef(presetBindingKey);
  presetLive.current = presetBindingKey;
  const presetAction = (action: () => void): void => {
    if (!viewActive.current || presetLive.current !== presetBindingKey || !sheet) return;
    action();
    setPreferenceVersion((version) => version + 1);
  };
  return {
    preferenceDocumentId,
    preferenceStore,
    presetStore,
    view,
    viewport,
    setViewport,
    changeView,
    presetDocumentId,
    namedViewPresets,
    presetBindingKey,
    presetAction,
  } as const;
}
