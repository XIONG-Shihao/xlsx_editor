import type { AccessibilityFrame } from '../accessibilityController';
import type { ChartRecommendationFrame } from '../chartRecommendationsController';
import type { CreateNamesContext } from '../createNamesFromSelection';
import type { DuplicateContext, DuplicateReceipt } from '../duplicateSelection';
import type { FillAcrossContext } from '../WorkbookFillAcross';
import type { FillSeriesContext } from '../WorkbookFill';
import type { FormulaDraftHandle, FormulaDraftSnapshot } from '../namedFormulaDraft';
import type { FunctionCategory } from '../functionCatalog';
import type { FunctionInsertionContext } from '../functionInsertion';
import type { GridTables } from './gridTypes';
import type { IllustrationMode } from '../WorkbookShapes';
import type { LocalExecutionAction } from '../../core/local-action-execution.mjs';
import type { NoteFrame } from '../noteController';
import type { PageStylesFrame } from '../pageStylesController';
import type { ReplaceContext } from '../WorkbookReplace';
import type { SortContext } from '../sortSelection';
import type { SpecialSelectionContext } from '../WorkbookSpecialSelection';
import type { WorkbookReviewMode } from '../workbookReviewController';
import type { WorkbookTool } from '../SpreadsheetRibbon';
import type { WorksheetObjectFrame } from '../worksheetObjectController';
import { useId, useLayoutEffect, useRef, useState } from 'react';

export interface ToolPaneStateInputs {
  active: boolean;
  duplicatePreviewReceipt: DuplicateReceipt | undefined;
  namedRanges: { name: string; formula: string; sheet?: string; scopeInvalid?: true }[];
  tables: GridTables;
  viewActive: React.MutableRefObject<boolean>;
}

export function useToolPaneState({
  active,
  duplicatePreviewReceipt,
  namedRanges,
  tables,
  viewActive,
}: ToolPaneStateInputs) {
  const [sheetAction, setSheetAction] = useState<'addSheet' | 'renameSheet' | 'duplicateSheet' | null>(null);
  const [newSheetName, setNewSheetName] = useState('');
  const [pageStylesMode, setPageStylesMode] = useState<'themes' | 'backgrounds'>('themes'),
    pageStylesLive = useRef<PageStylesFrame | null>(null);
  const [illustrationMode, setIllustrationMode] = useState<IllustrationMode>('shape');
  const illustrationLive = useRef<WorksheetObjectFrame | null>(null);
  const [chartToolsMode, setChartToolsMode] = useState<'manage' | 'recommendations'>('manage');
  const chartLive = useRef<ChartRecommendationFrame | null>(null);
  const accessibilityLive = useRef<AccessibilityFrame | null>(null);
  const notesLive = useRef<NoteFrame | null>(null);
  const [activeTool, setActiveTool] = useState<WorkbookTool | null>(null);
  const [visitedTools, setVisitedTools] = useState<WorkbookTool[]>([]);
  const [reviewIntent, setReviewIntent] = useState<{ serial: number; mode: WorkbookReviewMode }>({
    serial: 0,
    mode: 'statistics',
  });
  const [queriesMode, setQueriesMode] = useState<'individual' | 'all'>('individual'),
    [queriesAllIntent, setQueriesAllIntent] = useState(0);
  const [actionsIntent, setActionsIntent] = useState<{
    serial: number;
    mode: 'new' | 'list' | 'run';
    action?: LocalExecutionAction;
  }>({ serial: 0, mode: 'list' });
  const actionsParkedDraft = useRef<{ key: string; handle: FormulaDraftHandle; snapshot: FormulaDraftSnapshot } | null>(
    null
  );
  const reviewParkedDraft = useRef<{ key: string; handle: FormulaDraftHandle; snapshot: FormulaDraftSnapshot } | null>(
    null
  );
  const [duplicateContext, setDuplicateContext] = useState<DuplicateContext | null>(null),
    duplicateSerial = useRef(0),
    duplicateWritePhysical = useRef(false);
  const [fillToolMode, setFillToolMode] = useState<'series' | 'across'>('series');
  const [acrossContext, setAcrossContext] = useState<FillAcrossContext | null>(null),
    acrossSerial = useRef(0),
    acrossWritePhysical = useRef(false);
  const acrossLive = useRef({ key: '', visible: false });
  acrossLive.current = {
    key: acrossContext?.key ?? '',
    visible: active && activeTool === 'fill' && fillToolMode === 'across',
  };
  const [seriesContext, setSeriesContext] = useState<FillSeriesContext | null>(null),
    seriesSerial = useRef(0),
    seriesWritePhysical = useRef(false);
  const seriesLive = useRef({ key: '', visible: false });
  seriesLive.current = {
    key: seriesContext?.key ?? '',
    visible: active && activeTool === 'fill' && fillToolMode === 'series',
  };
  const seriesParkedDraft = useRef<{ key: string; handle: FormulaDraftHandle; snapshot: FormulaDraftSnapshot } | null>(
    null
  );
  const [replacementContext, setReplacementContext] = useState<ReplaceContext | null>(null),
    replacementSerial = useRef(0),
    replacementWritePhysical = useRef(false);
  const replacementLive = useRef({ key: '', visible: false });
  replacementLive.current = { key: replacementContext?.key ?? '', visible: active && activeTool === 'replace' };
  const replacementParkedDraft = useRef<{
    key: string;
    handle: FormulaDraftHandle;
    snapshot: FormulaDraftSnapshot;
  } | null>(null);
  const [specialContext, setSpecialContext] = useState<SpecialSelectionContext | null>(null),
    specialSerial = useRef(0),
    specialSelectPhysical = useRef(false);
  const specialLive = useRef({ key: '', visible: false });
  specialLive.current = { key: specialContext?.key ?? '', visible: active && activeTool === 'special' };
  const specialParkedDraft = useRef<{ key: string; handle: FormulaDraftHandle; snapshot: FormulaDraftSnapshot } | null>(
    null
  );
  const textColumnsParkedDraft = useRef<{
    key: string;
    handle: FormulaDraftHandle;
    snapshot: FormulaDraftSnapshot;
  } | null>(null);
  const duplicateParkedDraft = useRef<{
    key: string;
    handle: FormulaDraftHandle;
    snapshot: FormulaDraftSnapshot;
  } | null>(null);
  const liveDuplicateReceipt = useRef(duplicatePreviewReceipt);
  liveDuplicateReceipt.current = duplicatePreviewReceipt;
  const [sortContext, setSortContext] = useState<SortContext | null>(null),
    sortSerial = useRef(0),
    sortPhysical = useRef(false),
    [sortPending, setSortPending] = useState(false);
  const sortParkedDraft = useRef<{ key: string; handle: FormulaDraftHandle; snapshot: FormulaDraftSnapshot } | null>(
    null
  );
  const [createNamesContext, setCreateNamesContext] = useState<CreateNamesContext | null>(null);
  const liveNameInventory = useRef({ namedRanges, tables });
  liveNameInventory.current = { namedRanges, tables };
  const createNamesParkedDraft = useRef<{
    key: string;
    handle: FormulaDraftHandle;
    snapshot: FormulaDraftSnapshot;
  } | null>(null);
  const [functionContext, setFunctionContext] = useState<FunctionInsertionContext | null>(null);
  const [functionBrowse, setFunctionBrowse] = useState<{ mode: 'all' | 'recent' | FunctionCategory; serial: number }>({
    mode: 'all',
    serial: 0,
  });
  const taskPaneId = useId(),
    taskPaneTitleId = useId();
  const taskPaneHeading = useRef<HTMLHeadingElement>(null);
  const booleanWritePending = useRef(false);
  const toolFocusGeneration = useRef(0);
  useLayoutEffect(() => {
    if (activeTool && viewActive.current) taskPaneHeading.current?.focus({ preventScroll: true });
  }, [activeTool]);
  return {
    sheetAction,
    setSheetAction,
    newSheetName,
    setNewSheetName,
    pageStylesMode,
    setPageStylesMode,
    pageStylesLive,
    illustrationMode,
    setIllustrationMode,
    illustrationLive,
    chartToolsMode,
    setChartToolsMode,
    chartLive,
    accessibilityLive,
    notesLive,
    activeTool,
    setActiveTool,
    visitedTools,
    setVisitedTools,
    reviewIntent,
    setReviewIntent,
    queriesMode,
    setQueriesMode,
    queriesAllIntent,
    setQueriesAllIntent,
    actionsIntent,
    setActionsIntent,
    actionsParkedDraft,
    reviewParkedDraft,
    duplicateContext,
    setDuplicateContext,
    duplicateSerial,
    duplicateWritePhysical,
    fillToolMode,
    setFillToolMode,
    acrossContext,
    setAcrossContext,
    acrossSerial,
    acrossWritePhysical,
    acrossLive,
    seriesContext,
    setSeriesContext,
    seriesSerial,
    seriesWritePhysical,
    seriesLive,
    seriesParkedDraft,
    replacementContext,
    setReplacementContext,
    replacementSerial,
    replacementWritePhysical,
    replacementLive,
    replacementParkedDraft,
    specialContext,
    setSpecialContext,
    specialSerial,
    specialSelectPhysical,
    specialLive,
    specialParkedDraft,
    textColumnsParkedDraft,
    duplicateParkedDraft,
    liveDuplicateReceipt,
    sortContext,
    setSortContext,
    sortSerial,
    sortPhysical,
    sortPending,
    setSortPending,
    sortParkedDraft,
    createNamesContext,
    setCreateNamesContext,
    liveNameInventory,
    createNamesParkedDraft,
    functionContext,
    setFunctionContext,
    functionBrowse,
    setFunctionBrowse,
    taskPaneId,
    taskPaneTitleId,
    taskPaneHeading,
    booleanWritePending,
    toolFocusGeneration,
  } as const;
}
