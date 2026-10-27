import type { SelectionSummary } from '../../core/indexed-session.mjs';
import type { WorkbookSheet } from '../../core/contracts';
import { useEffect, useRef, useState } from 'react';

export interface SelectionSummaryInputs {
  range: string;
  selected: string;
  sheet: WorkbookSheet;
  specialSelection: Set<string> | null;
  storageKey: string | undefined;
  summarize: ((sheet: string, range: string) => Promise<SelectionSummary>) | undefined;
  /** Read when the code runs, like the closed-over variables they stand for. */
  late: {
    readonly activeRange: () => string;
    readonly busy: boolean;
  };
}

export function useSelectionSummary({
  range,
  selected,
  sheet,
  specialSelection,
  storageKey,
  summarize,
  late,
}: SelectionSummaryInputs) {
  const summaryReader = useRef(summarize);
  summaryReader.current = summarize;
  const [selectionSummary, setSelectionSummary] = useState<SelectionSummary | null>(null);
  const [summaryError, setSummaryError] = useState('');
  useEffect(() => {
    let cancelled = false;
    setSelectionSummary(null);
    setSummaryError('');
    if (!summaryReader.current || !sheet || late.busy || specialSelection) return;
    const timer = setTimeout(() => {
      void summaryReader.current!(sheet.name, late.activeRange())
        .then((result) => {
          if (!cancelled) setSelectionSummary(result);
        })
        .catch((error) => {
          if (!cancelled) setSummaryError(String(error));
        });
    }, 150);
    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
  }, [storageKey, sheet?.name, range, selected, late.busy, specialSelection]);
  return { selectionSummary, summaryError } as const;
}
