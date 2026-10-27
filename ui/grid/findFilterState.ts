import type { FilterValue, FilterValuePage } from './filterValues';
import type { SearchCursor } from '../searchPages';
import type { WorkbookCell, WorkbookSheet } from '../../core/contracts';
import { useEffect, useRef, useState } from 'react';

export interface FindFilterStateInputs {
  metadata: WorkbookSheet[];
  sheet: WorkbookSheet;
  storageKey: string | undefined;
}

export function useFindFilterState({ metadata, sheet, storageKey }: FindFilterStateInputs) {
  const [findOpen, setFindOpen] = useState(false);
  const [findText, setFindText] = useState('');
  const [replaceText, setReplaceText] = useState('');
  const [findScope, setFindScope] = useState<'sheet' | 'workbook'>('sheet');
  const [findMode, setFindMode] = useState<'display' | 'raw' | 'formula'>('display');
  const [findMatches, setFindMatches] = useState<Array<{ sheet: string; cell: string }>>([]);
  const [findTotal, setFindTotal] = useState(0);
  const [findPageIndex, setFindPageIndex] = useState(0);
  const [findNextCursor, setFindNextCursor] = useState<SearchCursor | null>(null);
  const [findIndex, setFindIndex] = useState(0);
  const [findBusy, setFindBusy] = useState(false);
  const replacementConfigurationKey = JSON.stringify([findOpen, findText, replaceText, findScope, findMode]);
  const replacementConfigurationLive = useRef(replacementConfigurationKey);
  replacementConfigurationLive.current = replacementConfigurationKey;
  const [sortPrimary, setSortPrimary] = useState('');
  const [conditionalRule, setConditionalRule] = useState<'greaterThan' | 'lessThan' | 'duplicateValues'>('greaterThan');
  const [conditionalThreshold, setConditionalThreshold] = useState('0');
  const [conditionalColor, setConditionalColor] = useState('FFF2CC');
  const [tableName, setTableName] = useState('Table1');
  const [sortHeader, setSortHeader] = useState(true);
  const filterChoices = useRef<HTMLSelectElement>(null);
  const [filterText, setFilterText] = useState('');
  const [filterPage, setFilterPage] = useState<FilterValuePage | null>(null);
  const [filterOffset, setFilterOffset] = useState(0);
  const [filterBusy, setFilterBusy] = useState(false),
    [filterReadPending, setFilterReadPending] = useState(false),
    [filterApplying, setFilterApplying] = useState(false);
  const [filterPageScope, setFilterPageScope] = useState('');
  const filterPhysical = useRef(false),
    filterWritePhysical = useRef(false),
    filterMounted = useRef(true);
  const [filterNotice, setFilterNotice] = useState<{ key: string; text: string } | null>(null);
  useEffect(() => {
    filterMounted.current = true;
    return () => {
      filterMounted.current = false;
      filterGeneration.current++;
    };
  }, []);
  const [filterSelected, setFilterSelected] = useState(new Map<string, FilterValue>());
  const filterGeneration = useRef(0);
  const filterScope = useRef('');
  const liveFilterDefinition = useRef(sheet?.filter);
  liveFilterDefinition.current = sheet?.filter;
  const resetFilter = (): void => {
    filterGeneration.current++;
    setFilterBusy(false);
    setFilterPage(null);
    setFilterPageScope('');
    setFilterOffset(0);
    setFilterSelected(new Map());
  };
  const findGeneration = useRef(0);
  const foundValues = useRef(new Map<string, WorkbookCell>());
  const findConfig = useRef<{ text: string; mode: 'display' | 'raw' | 'formula'; sheets: string[] } | null>(null);
  const findPages = useRef<Array<{ cursor: SearchCursor; offset: number }>>([]);
  const resetFind = (): void => {
    findGeneration.current++;
    setFindBusy(false);
    setFindMatches([]);
    setFindTotal(0);
    setFindPageIndex(0);
    setFindNextCursor(null);
    foundValues.current.clear();
    findConfig.current = null;
    findPages.current = [];
  };
  useEffect(resetFind, [storageKey, storageKey === undefined ? metadata : undefined]);
  useEffect(
    () => () => {
      findGeneration.current++;
    },
    []
  );
  return {
    findOpen,
    setFindOpen,
    findText,
    setFindText,
    replaceText,
    setReplaceText,
    findScope,
    setFindScope,
    findMode,
    setFindMode,
    findMatches,
    setFindMatches,
    findTotal,
    setFindTotal,
    findPageIndex,
    setFindPageIndex,
    findNextCursor,
    setFindNextCursor,
    findIndex,
    setFindIndex,
    findBusy,
    setFindBusy,
    replacementConfigurationKey,
    replacementConfigurationLive,
    sortPrimary,
    setSortPrimary,
    conditionalRule,
    setConditionalRule,
    conditionalThreshold,
    setConditionalThreshold,
    conditionalColor,
    setConditionalColor,
    tableName,
    setTableName,
    sortHeader,
    setSortHeader,
    filterChoices,
    filterText,
    setFilterText,
    filterPage,
    setFilterPage,
    filterOffset,
    setFilterOffset,
    filterBusy,
    setFilterBusy,
    filterReadPending,
    setFilterReadPending,
    filterApplying,
    setFilterApplying,
    filterPageScope,
    setFilterPageScope,
    filterPhysical,
    filterWritePhysical,
    filterMounted,
    filterNotice,
    setFilterNotice,
    filterSelected,
    setFilterSelected,
    filterGeneration,
    filterScope,
    liveFilterDefinition,
    resetFilter,
    findGeneration,
    foundValues,
    findConfig,
    findPages,
    resetFind,
  } as const;
}
