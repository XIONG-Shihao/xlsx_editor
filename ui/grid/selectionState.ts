import type { FormulaDraftHandle, FormulaDraftSnapshot, NamedFormulaChoice } from '../namedFormulaDraft';
import { type FormulaPreviewFrame, createFormulaDraftPreviewController } from '../formulaDraftPreview';
import type { HomeBinding, HomeFrame } from '../homeCommands';
import type { WorkbookCommand, WorkbookSheet } from '../../core/contracts';
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';

export interface SelectionStateInputs {
  metadata: WorkbookSheet[];
  setJump: React.Dispatch<React.SetStateAction<string>>;
  setJumpError: React.Dispatch<React.SetStateAction<string>>;
  sheet: WorkbookSheet;
}

export function useSelectionState({ metadata, setJump, setJumpError, sheet }: SelectionStateInputs) {
  const [range, setRange] = useState('A1');
  const [specialSelection, setSpecialSelection] = useState<Set<string> | null>(null);
  const specialSelectionMode = useRef({ selection: specialSelection });
  if (specialSelectionMode.current.selection !== specialSelection)
    specialSelectionMode.current = { selection: specialSelection };
  const [selectionEnd, setSelectionEnd] = useState('A1');
  const [objectExtent, setObjectExtent] = useState({ family: '', rows: 0, columns: 0 });
  const updateObjectExtent = useCallback(
    (family: string, extent: { rows: number; columns: number }) =>
      setObjectExtent((previous) =>
        previous.family === family && previous.rows === extent.rows && previous.columns === extent.columns
          ? previous
          : { family, ...extent }
      ),
    []
  );
  const dragging = useRef(false);
  const fillDragging = useRef<{
    source: string;
    target: string | null;
    dispatch?: (range: string) => Promise<void>;
  } | null>(null);
  const [fillPreview, setFillPreview] = useState<string | null>(null);
  const homeFrameRef = useRef<HomeFrame | null>(null),
    homePhysical = useRef(false),
    [homePending, setHomePending] = useState(false);
  const homeMetadata = useRef({ value: metadata, epoch: 0 });
  if (homeMetadata.current.value !== metadata)
    homeMetadata.current = { value: metadata, epoch: homeMetadata.current.epoch + 1 };
  const homeSpecial = useRef({ value: specialSelection, epoch: 0 });
  if (homeSpecial.current.value !== specialSelection)
    homeSpecial.current = { value: specialSelection, epoch: homeSpecial.current.epoch + 1 };
  const findReceipt = useRef(''),
    paintHomeBinding = useRef<HomeBinding | null>(null),
    homePaintTarget = useRef<string | null>(null);
  const homePaintRunner = useRef<
    ((cell: string, format: Extract<WorkbookCommand, { type: 'formatRange' }>['format']) => void) | null
  >(null);
  const [paintStyleId, setPaintStyleId] = useState<number | null>(null);
  const paintStyleRef = useRef<Extract<WorkbookCommand, { type: 'formatRange' }>['format'] | null>(null);
  const internalClipboard = useRef<{
    sheet: string;
    range: string;
    text: string;
    cut: boolean;
    binding?: HomeBinding;
  } | null>(null);
  const [selected, setSelected] = useState('A1');
  const formulaDraftHandle = useRef<FormulaDraftHandle | null>(null);
  const [formulaComposing, setFormulaComposing] = useState(false);
  const [formulaPreviewDraft, setFormulaPreviewDraft] = useState<{
    scopeKey: string;
    snapshot: FormulaDraftSnapshot;
    active: boolean;
  } | null>(null);
  const formulaPreviewFrame = useRef<FormulaPreviewFrame | null>(null);
  const formulaPreviewController = useMemo(
    () => createFormulaDraftPreviewController(() => formulaPreviewFrame.current),
    []
  );
  const previewDisposal = useRef<ReturnType<typeof setTimeout> | null>(null);
  useEffect(() => {
    if (previewDisposal.current !== null) clearTimeout(previewDisposal.current);
    return () => {
      previewDisposal.current = setTimeout(() => formulaPreviewController.dispose(), 0);
    };
  }, [formulaPreviewController]);
  const reportFormulaPreviewDraft = useCallback(
    (scopeKey: string, snapshot: FormulaDraftSnapshot, active: boolean) =>
      setFormulaPreviewDraft({ scopeKey, snapshot, active }),
    []
  );
  const [namePicker, setNamePicker] = useState<{
    key: string;
    target: string;
    snapshot: FormulaDraftSnapshot;
    choices: readonly NamedFormulaChoice[];
    handle: FormulaDraftHandle;
  } | null>(null);
  const namePickerSerial = useRef(0);
  const namedDraftLive = useRef({ key: '', editable: false });
  useLayoutEffect(() => {
    setJump('');
    setJumpError('');
  }, [sheet?.name, selected, range]);
  const selectedRef = useRef(selected);
  selectedRef.current = selected;
  return {
    range,
    setRange,
    specialSelection,
    setSpecialSelection,
    specialSelectionMode,
    selectionEnd,
    setSelectionEnd,
    objectExtent,
    updateObjectExtent,
    dragging,
    fillDragging,
    fillPreview,
    setFillPreview,
    homeFrameRef,
    homePhysical,
    homePending,
    setHomePending,
    homeMetadata,
    homeSpecial,
    findReceipt,
    paintHomeBinding,
    homePaintTarget,
    homePaintRunner,
    paintStyleId,
    setPaintStyleId,
    paintStyleRef,
    internalClipboard,
    selected,
    setSelected,
    formulaDraftHandle,
    formulaComposing,
    setFormulaComposing,
    formulaPreviewDraft,
    formulaPreviewFrame,
    formulaPreviewController,
    reportFormulaPreviewDraft,
    namePicker,
    setNamePicker,
    namePickerSerial,
    namedDraftLive,
    selectedRef,
  } as const;
}
