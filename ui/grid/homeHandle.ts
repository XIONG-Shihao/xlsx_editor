import { BULK_CELL_LIMIT } from './gridLimits';
import type {
  GridFindConfig,
  GridHomeFormatCommands,
  GridHomeReadCells,
  GridInternalClipboard,
  GridPaintStyleRef,
  GridScan,
  GridToggleTool,
} from './gridTypes';
import type { HomeBinding, HomeCommandId, HomeCommandLease as HomeLease } from '../homeCommands';
import { type SearchCursor, type SearchQuery, loadSearchPage, scanSearchQuery } from '../searchPages';
import type { SpecialSelectionKind } from '../specialSelection';
import type { WorkbookCell, WorkbookCommand, WorkbookNormalFont, WorkbookSheet } from '../../core/contracts';
import { addressPosition, colName, displayValue } from '../sheetModel';
import { clearHyperlinksTargets } from '../clearHyperlinksTargets';
import { editableCellFormat } from './editableFormats';
import { editorHost } from '../host';
import { homeStyleSamples } from '../homeFormattingSamples';
import { rectangle, selectedBounds } from './rangeAddress';
import { replacementForMatch } from '../replacePreview';

export interface HomeHandleInputs {
  conditionalThreshold: string;
  findBusy: boolean;
  findConfig: GridFindConfig;
  findGeneration: React.MutableRefObject<number>;
  findIndex: number;
  findMatches: { sheet: string; cell: string }[];
  findMode: 'display' | 'raw' | 'formula';
  findNextCursor: SearchCursor | null;
  findPageIndex: number;
  findPages: React.MutableRefObject<{ cursor: SearchCursor; offset: number }[]>;
  findReceipt: React.MutableRefObject<string>;
  findScope: 'sheet' | 'workbook';
  findText: string;
  findTotal: number;
  foundValues: React.MutableRefObject<Map<string, WorkbookCell>>;
  homeClipboardText: (target: string, cells: Record<string, WorkbookCell>) => string;
  homeFormatCommands: GridHomeFormatCommands;
  homePaintTarget: React.MutableRefObject<string | null>;
  homePaste: (lease: HomeLease, text: string) => Promise<void>;
  homePasteAsText: (lease: HomeLease, text: string) => Promise<void>;
  homePasteSpecial: (lease: HomeLease, mode: 'values' | 'formulas' | 'formatting') => Promise<void>;
  homeReadCells: GridHomeReadCells;
  homeSameTarget: (binding: HomeBinding) => boolean;
  internalClipboard: GridInternalClipboard;
  literal: boolean | undefined;
  normalFont: Readonly<WorkbookNormalFont>;
  onCommandRedo: (() => Promise<void>) | undefined;
  onCommandUndo: (() => Promise<void>) | undefined;
  paintHomeBinding: React.MutableRefObject<HomeBinding | null>;
  paintStyleRef: GridPaintStyleRef;
  pasteAbort: React.MutableRefObject<AbortController | null>;
  query: SearchQuery | undefined;
  replaceText: string;
  resetFind: () => void;
  scan: GridScan;
  select: (cell: string, extend?: boolean) => void;
  setConditionalColor: React.Dispatch<React.SetStateAction<string>>;
  setConditionalRule: React.Dispatch<React.SetStateAction<'greaterThan' | 'lessThan' | 'duplicateValues'>>;
  setConditionalThreshold: React.Dispatch<React.SetStateAction<string>>;
  setFindBusy: React.Dispatch<React.SetStateAction<boolean>>;
  setFindIndex: React.Dispatch<React.SetStateAction<number>>;
  setFindMatches: React.Dispatch<React.SetStateAction<{ sheet: string; cell: string }[]>>;
  setFindMode: React.Dispatch<React.SetStateAction<'display' | 'raw' | 'formula'>>;
  setFindNextCursor: React.Dispatch<React.SetStateAction<SearchCursor | null>>;
  setFindOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setFindPageIndex: React.Dispatch<React.SetStateAction<number>>;
  setFindScope: React.Dispatch<React.SetStateAction<'sheet' | 'workbook'>>;
  setFindText: React.Dispatch<React.SetStateAction<string>>;
  setFindTotal: React.Dispatch<React.SetStateAction<number>>;
  setPaintStyleId: React.Dispatch<React.SetStateAction<number | null>>;
  setRange: React.Dispatch<React.SetStateAction<string>>;
  setReplaceText: React.Dispatch<React.SetStateAction<string>>;
  setSpecialSelection: React.Dispatch<React.SetStateAction<Set<string> | null>>;
  setTableName: React.Dispatch<React.SetStateAction<string>>;
  sheet: WorkbookSheet;
  sheets: WorkbookSheet[];
  storageKey: string | undefined;
  /** Read when the code runs, like the closed-over variables they stand for. */
  late: {
    readonly goToMatch: any;
    readonly homeWorkbookKey: string;
    readonly openAcross: () => void;
    readonly openReplacement: () => void;
    readonly openSeries: () => void;
    readonly openSpecial: (kind?: SpecialSelectionKind) => void;
    readonly structuralBlock: string;
    readonly styleCell: WorkbookCell;
    readonly toggleTool: GridToggleTool;
  };
}

export function buildHomeHandle({
  conditionalThreshold,
  findBusy,
  findConfig,
  findGeneration,
  findIndex,
  findMatches,
  findMode,
  findNextCursor,
  findPageIndex,
  findPages,
  findReceipt,
  findScope,
  findText,
  findTotal,
  foundValues,
  homeClipboardText,
  homeFormatCommands,
  homePaintTarget,
  homePaste,
  homePasteAsText,
  homePasteSpecial,
  homeReadCells,
  homeSameTarget,
  internalClipboard,
  literal,
  normalFont,
  onCommandRedo,
  onCommandUndo,
  paintHomeBinding,
  paintStyleRef,
  pasteAbort,
  query,
  replaceText,
  resetFind,
  scan,
  select,
  setConditionalColor,
  setConditionalRule,
  setConditionalThreshold,
  setFindBusy,
  setFindIndex,
  setFindMatches,
  setFindMode,
  setFindNextCursor,
  setFindOpen,
  setFindPageIndex,
  setFindScope,
  setFindText,
  setFindTotal,
  setPaintStyleId,
  setRange,
  setReplaceText,
  setSpecialSelection,
  setTableName,
  sheet,
  sheets,
  storageKey,
  late,
}: HomeHandleInputs) {
  const homeSearch = async (lease: HomeLease, index: number, cursor: SearchCursor, offset: number): Promise<void> => {
    const config = findConfig.current;
    if (!config) return;
    const generation = ++findGeneration.current;
    setFindBusy(true);
    try {
      const queryPage =
        query ??
        scanSearchQuery({
          scan,
          cells: (name) => sheets.find((item) => item.name === name)?.cells ?? {},
          content: (name, ref, value, mode) =>
            mode === 'formula'
              ? value.formula
                ? '=' + value.formula
                : ''
              : mode === 'raw'
                ? String(value.rawValue ?? value.value ?? '')
                : displayValue({ ...sheets.find((item) => item.name === name)!, cells: { [ref]: value } }, ref),
        });
      const result = await lease.read(() =>
        loadSearchPage({
          sheets: config.sheets,
          text: config.text,
          mode: config.mode,
          query: queryPage,
          cursor,
          cancelled: () => generation !== findGeneration.current || !lease.isCurrent(),
          yieldWork: () => new Promise<void>((resolve) => requestAnimationFrame(() => resolve())),
        })
      );
      if (!result || generation !== findGeneration.current) return;
      if (!result.matches.length && index > 0) {
        setFindNextCursor(null);
        return;
      }
      foundValues.current.clear();
      const found = result.matches.map(({ sheet, cell, value }) => {
        foundValues.current.set(`${sheet}\0${cell}`, value);
        return { sheet, cell };
      });
      findPages.current[index] = { cursor, offset };
      findPages.current.length = index + 1;
      setFindMatches(found);
      setFindTotal(offset + found.length);
      setFindPageIndex(index);
      setFindNextCursor(result.next);
      setFindIndex(0);
      findReceipt.current = late.homeWorkbookKey;
      if (found[0]) late.goToMatch(found[0]);
    } finally {
      if (generation === findGeneration.current) setFindBusy(false);
    }
  };
  const homeHandle = async (id: HomeCommandId, lease: HomeLease, args: unknown): Promise<void> => {
    const a = (args ?? {}) as {
      value?: any;
      format?: any;
      field?: string;
      name?: string;
      rule?: any;
      threshold?: number;
      fillColor?: string;
      direction?: string;
      text?: string;
      index?: number;
      source?: string;
      range?: string;
    };
    const b = lease.binding,
      position = addressPosition(b.selectedCell);
    const formats: Partial<Record<HomeCommandId, Extract<WorkbookCommand, { type: 'formatRange' }>['format']>> = {
      'home.font-family': { fontFamily: a.value },
      'home.font-size': { fontSize: a.value },
      'home.bold': { bold: a.value },
      'home.italic': { italic: a.value },
      'home.underline': { underline: a.value },
      'home.borders': { border: a.value },
      'home.fill-color': { fillColor: a.value },
      'home.font-color': { fontColor: a.value },
      'home.wrap': { wrapText: a.value },
      'home.number-format': { numberFormat: a.value },
      'home.font-size-increase': { fontSize: Math.min(72, (late.styleCell?.fontSize ?? normalFont.fontSize) + 1) },
      'home.font-size-decrease': { fontSize: Math.max(6, (late.styleCell?.fontSize ?? normalFont.fontSize) - 1) },
      'home.align-top': { verticalAlign: 'top' },
      'home.align-middle': { verticalAlign: 'center' },
      'home.align-bottom': { verticalAlign: 'bottom' },
      'home.align-left': { horizontalAlign: 'left' },
      'home.align-center': { horizontalAlign: 'center' },
      'home.align-right': { horizontalAlign: 'right' },
      'home.accounting': { numberFormat: 'accounting' },
      'home.percent': { numberFormat: 'percent' },
      'home.thousands': { numberFormat: 'thousands' },
      'home.style-good': homeStyleSamples.good,
      'home.style-bad': homeStyleSamples.bad,
      'home.style-heading': homeStyleSamples.heading,
    };
    if (formats[id]) {
      await lease.accept(homeFormatCommands(lease, formats[id]!));
      return;
    }
    if (id === 'home.decimals-increase' || id === 'home.decimals-decrease') {
      if (![undefined, 0, 1, 2, '0.0', '0.00'].includes(late.styleCell?.format))
        throw Error('Choose an ordinary numeric format before changing decimals.');
      const old =
          late.styleCell?.format === 1
            ? 0
            : late.styleCell?.format === '0.0'
              ? 1
              : late.styleCell?.format === 2 || late.styleCell?.format === '0.00'
                ? 2
                : 0,
        next = Math.max(0, Math.min(2, old + (id === 'home.decimals-increase' ? 1 : -1)));
      await lease.accept(
        homeFormatCommands(lease, { numberFormat: (['decimal0', 'decimal1', 'decimal2'] as const)[next] })
      );
      return;
    }
    switch (id) {
      case 'home.copy':
      case 'home.cut': {
        const cells = await homeReadCells(lease, b.sheet, b.selection),
          text = homeClipboardText(b.selection, cells);
        await lease.read(() => editorHost().api.copyText(text));
        internalClipboard.current = { sheet: b.sheet, range: b.selection, text, cut: id === 'home.cut', binding: b };
        return;
      }
      case 'home.paste': {
        const text = a.text ?? (await lease.read(() => editorHost().api.readClipboardText()));
        await homePaste(lease, text);
        return;
      }
      case 'home.paste-text': {
        const text = a.text ?? (await lease.read(() => editorHost().api.readClipboardText()));
        await homePasteAsText(lease, text);
        return;
      }
      case 'home.paste-values':
      case 'home.paste-formulas':
      case 'home.paste-formatting':
        await homePasteSpecial(
          lease,
          id === 'home.paste-values' ? 'values' : id === 'home.paste-formulas' ? 'formulas' : 'formatting'
        );
        return;
      case 'home.cancel-paste':
        pasteAbort.current?.abort();
        return;
      case 'home.format-painter': {
        const ref = b.specialSelection?.[0] ?? b.selectedCell,
          cells = await homeReadCells(lease, b.sheet, ref);
        paintStyleRef.current = editableCellFormat(cells[ref]);
        paintHomeBinding.current = b;
        setPaintStyleId(cells[ref]?.styleId ?? 0);
        return;
      }
      case 'home.paint-format':
        await lease.accept([
          { type: 'formatRange', sheet: b.sheet, range: homePaintTarget.current ?? b.selectedCell, format: a.format },
        ]);
        return;
      case 'home.merge-center':
        selectedBounds(b.selection, 1000);
        if (!b.selection.includes(':')) throw Error('Select a range to merge.');
        await lease.accept([
          { sheet: b.sheet, operation: 'mergeCells', range: b.selection },
          {
            type: 'formatRange',
            sheet: b.sheet,
            range: b.selection.split(':')[0],
            format: { horizontalAlign: 'center', verticalAlign: 'center' },
          },
        ]);
        return;
      case 'home.unmerge':
        if (!sheet.merges.includes(b.selection)) throw Error('Select the exact merged range to unmerge.');
        await lease.accept([{ sheet: b.sheet, operation: 'unmergeCells', range: b.selection }]);
        return;
      case 'styles.conditional':
        late.toggleTool('conditional');
        return;
      case 'home.conditional-draft':
        if (a.field === 'rule') setConditionalRule(a.value);
        else if (a.field === 'threshold') setConditionalThreshold(a.value);
        else setConditionalColor(a.value);
        return;
      case 'home.conditional-apply':
        if (a.rule !== 'duplicateValues' && (!conditionalThreshold.trim() || !Number.isFinite(a.threshold)))
          throw Error('Enter a numeric threshold.');
        selectedBounds(b.selection, BULK_CELL_LIMIT);
        await lease.accept([
          {
            sheet: b.sheet,
            operation: 'addConditionalRule',
            range: b.selection,
            rule: a.rule,
            ...(a.rule === 'duplicateValues' ? {} : { threshold: a.threshold }),
            fillColor: a.fillColor,
          },
        ]);
        return;
      case 'home.table-draft':
        setTableName(a.name!);
        return;
      case 'home.create-table': {
        const bounds = selectedBounds(b.selection, BULK_CELL_LIMIT);
        if (bounds.rows < 2) throw Error('Select a header row and at least one data row.');
        await lease.accept([{ sheet: b.sheet, operation: 'createTable', range: b.selection, tableName: a.name! }]);
        return;
      }
      case 'home.insert-row':
      case 'home.delete-row':
      case 'home.insert-column':
      case 'home.delete-column': {
        const row = id.endsWith('row'),
          insert = id.includes('insert');
        if (
          (insert && row && !sheet.capabilities?.insertRows) ||
          (!insert && row && !!late.structuralBlock && !sheet.capabilities?.deleteRows) ||
          (!row && !!late.structuralBlock && !sheet.capabilities?.referenceColumns)
        )
          throw Error(late.structuralBlock ?? 'This structural edit is protected.');
        await lease.accept([
          {
            sheet: b.sheet,
            operation: insert ? (row ? 'insertRow' : 'insertColumn') : row ? 'deleteRow' : 'deleteColumn',
            index: row ? position.row : position.col,
          },
        ]);
        return;
      }
      case 'home.column-width':
      case 'home.row-height':
        await lease.accept([
          {
            sheet: b.sheet,
            operation: id === 'home.column-width' ? 'resizeColumn' : 'resizeRow',
            index: id === 'home.column-width' ? position.col : position.row,
            ...(id === 'home.column-width' ? { width: a.value } : { height: a.value }),
          },
        ]);
        return;
      case 'home.fill-range':
        await lease.accept([{ type: 'fillRange', sheet: b.sheet, source: a.source!, range: a.range! }]);
        if (homeSameTarget(b)) {
          const [first] = a.source!.split(':'),
            last = a.range!.split(':')[1] ?? a.range!;
          select(first);
          setRange(rectangle(first, last));
        }
        return;
      case 'home.fill-down':
      case 'home.fill-right':
      case 'home.fill-up':
      case 'home.fill-left': {
        const bounds = selectedBounds(b.selection, BULK_CELL_LIMIT),
          first = b.selection.split(':')[0],
          last = b.selection.split(':')[1] ?? first;
        let source: string, target: string;
        if (id === 'home.fill-down' || id === 'home.fill-up') {
          if (bounds.rows < 2)
            throw Error(`Select at least two rows to fill ${id === 'home.fill-down' ? 'down' : 'up'}.`);
          if (id === 'home.fill-down') {
            source = `${first}:${colName(bounds.last.col)}${bounds.first.row + 1}`;
            target = `${colName(bounds.first.col)}${bounds.first.row + 2}:${last}`;
          } else {
            source = `${colName(bounds.first.col)}${bounds.last.row + 1}:${last}`;
            target = `${first}:${colName(bounds.last.col)}${bounds.last.row}`;
          }
        } else {
          if (bounds.columns < 2)
            throw Error(`Select at least two columns to fill ${id === 'home.fill-right' ? 'right' : 'left'}.`);
          if (id === 'home.fill-right') {
            source = `${first}:${colName(bounds.first.col)}${bounds.last.row + 1}`;
            target = `${colName(bounds.first.col + 1)}${bounds.first.row + 1}:${last}`;
          } else {
            source = `${colName(bounds.last.col)}${bounds.first.row + 1}:${last}`;
            target = `${first}:${colName(bounds.last.col - 1)}${bounds.last.row + 1}`;
          }
        }
        await lease.accept([{ type: 'fillRange', sheet: b.sheet, source, range: target }]);
        return;
      }
      case 'home.fill-across':
        late.openAcross();
        return;
      case 'home.fill-across-preview':
      case 'home.fill-across-apply':
        throw Error('Preview the destinations in Across Worksheets before applying it.');
      case 'home.fill-series':
        late.openSeries();
        return;
      case 'home.fill-series-apply':
        throw Error('Preview the series in the Fill series pane before applying it.');
      case 'home.clear-contents':
      case 'home.clear-all':
      case 'home.clear-formats': {
        const bounds = selectedBounds(b.selection, BULK_CELL_LIMIT),
          commands: WorkbookCommand[] = [];
        if (id !== 'home.clear-formats') {
          if (storageKey && !b.specialSelection && !literal)
            commands.push({ type: 'clearRange', sheet: b.sheet, range: b.selection });
          else
            commands.push({
              type: 'setCells',
              sheet: b.sheet,
              cells: (
                b.specialSelection ??
                Array.from(
                  { length: bounds.rows * bounds.columns },
                  (_, i) =>
                    `${colName(bounds.first.col + (i % bounds.columns))}${bounds.first.row + Math.floor(i / bounds.columns) + 1}`
                )
              ).map((cell) => ({ cell, value: null })),
            });
        }
        if (id !== 'home.clear-contents') commands.push(...homeFormatCommands(lease, { styleId: 0 }));
        await lease.accept(commands);
        return;
      }
      case 'home.clear-hyperlinks': {
        if (!storageKey) throw Error('Clear Hyperlinks requires a native XLSX owner.');
        await lease.accept([clearHyperlinksTargets(b)]);
        return;
      }
      case 'home.special-open':
      case 'home.select-blanks':
      case 'home.select-formulas':
        late.openSpecial(
          id === 'home.select-blanks' ? 'blanks' : id === 'home.select-formulas' ? 'formulas' : 'constants'
        );
        return;
      case 'home.special-preview':
      case 'home.special-select':
        throw Error('Preview matching cells in the Go To Special pane first.');
      case 'home.clear-special-selection':
        setSpecialSelection(null);
        return;
      case 'home.find-open':
        setFindOpen(true);
        return;
      case 'home.find-close':
        resetFind();
        setFindOpen(false);
        return;
      case 'home.find-draft':
        if (a.field !== 'replacement') resetFind();
        if (a.field === 'text') setFindText(a.value);
        else if (a.field === 'replacement') setReplaceText(a.value);
        else if (a.field === 'scope') setFindScope(a.value);
        else setFindMode(a.value);
        return;
      case 'home.find-search':
        resetFind();
        if (findText) {
          findConfig.current = {
            text: findText,
            mode: findMode,
            sheets: (findScope === 'sheet' ? [sheet] : sheets).map((item) => item.name),
          };
          await homeSearch(lease, 0, { sheet: 0 }, 0);
        }
        return;
      case 'home.find-cancel':
        findGeneration.current++;
        setFindBusy(false);
        return;
      case 'home.find-page': {
        if (findReceipt.current !== late.homeWorkbookKey) throw Error('The search workbook changed. Find again.');
        const index = a.direction === 'previous' ? findPageIndex - 1 : findPageIndex + 1,
          prior = a.direction === 'previous' ? findPages.current[index] : null,
          cursor = prior?.cursor ?? findNextCursor;
        if (cursor) await homeSearch(lease, index, cursor, prior?.offset ?? findTotal);
        return;
      }
      case 'home.find-match':
        if (findReceipt.current !== late.homeWorkbookKey || !findMatches[a.index!])
          throw Error('The search match changed. Find again.');
        setFindIndex(a.index!);
        late.goToMatch(findMatches[a.index!]);
        return;
      case 'home.dimension-draft':
        return;
      case 'home.find-next':
        if (findReceipt.current !== late.homeWorkbookKey) throw Error('The search workbook changed. Find again.');
        if (findMatches.length) {
          const index = (findIndex + 1) % findMatches.length;
          setFindIndex(index);
          late.goToMatch(findMatches[index]);
        }
        return;
      case 'home.replace-preview':
        late.openReplacement();
        return;
      case 'home.replace-all':
        throw Error('Preview all replacements in the replacement pane before applying them.');
      case 'home.replace-one': {
        if (findReceipt.current !== late.homeWorkbookKey || findBusy)
          throw Error('Find current matches before replacing.');
        const match = findMatches[findIndex],
          prior = match && foundValues.current.get(`${match.sheet}\0${match.cell}`);
        if (!match || !prior) throw Error('Find a current match before replacing.');
        const row = replacementForMatch(
          { ...match, value: prior },
          {
            text: findText,
            replacement: replaceText,
            mode: findMode,
            sheets: [match.sheet],
            literal: !!literal,
            date1904: !!sheet.date1904,
          }
        );
        if (!row.edit) throw Error(row.reason ?? 'No editable replacement.');
        const { sheet: name, ...edit } = row.edit;
        await lease.accept([{ type: 'setCells', sheet: name, cells: [edit] }]);
        if (homeSameTarget(b)) resetFind();
        return;
      }
      case 'home.undo':
      case 'home.redo': {
        const history = id === 'home.undo' ? onCommandUndo : onCommandRedo;
        if (!history) throw Error('Workbook history commands are unavailable in this editor.');
        await lease.perform(history);
        return;
      }
      case 'home.menu':
        return;
    }
  };
  return { homeHandle } as const;
}
