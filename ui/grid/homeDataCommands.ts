import { BULK_CELL_LIMIT } from './gridLimits';
import type { GridImportPaste, GridInternalClipboard } from './gridTypes';
import {
  type HomeBinding,
  type HomeFrame,
  type HomeCommandLease as HomeLease,
  normalizeHomeBatch,
} from '../homeCommands';
import type { ReadRange } from '../rangeCells';
import type { WorkbookCell, WorkbookCommand, WorkbookEdit, WorkbookSheet } from '../../core/contracts';
import { addressPosition, colName, inputDisplay, maxColumnCount, parseGridInput } from '../sheetModel';
import { clearHyperlinksTargets } from '../clearHyperlinksTargets';
import { createValuePasteSourceContext, normalizeValuePasteCell } from '../../core/value-paste.mjs';
import { editableCellFormat } from './editableFormats';
import { inspectPaste } from '../largePaste';
import { literalClipboardCells, serializeClipboardText } from '../clipboardText';
import { parseDelimited } from '../../core/delimited';
import { pasteBytes } from '../../core/paste-protocol.mjs';
import { rangesOverlap, rectangle, selectedBounds } from './rangeAddress';

export interface HomeDataCommandsInputs {
  homeFrameRef: React.MutableRefObject<HomeFrame | null>;
  importPaste: GridImportPaste;
  internalClipboard: GridInternalClipboard;
  literal: boolean | undefined;
  metadata: WorkbookSheet[];
  pasteAbort: React.MutableRefObject<AbortController | null>;
  range: string;
  readRange: ReadRange | undefined;
  selected: string;
  setPasteProgress: React.Dispatch<React.SetStateAction<{ completed: number; total: number } | null>>;
  setRange: React.Dispatch<React.SetStateAction<string>>;
  sheets: WorkbookSheet[];
  storageKey: string | undefined;
}

export function buildHomeDataCommands({
  homeFrameRef,
  importPaste,
  internalClipboard,
  literal,
  metadata,
  pasteAbort,
  range,
  readRange,
  selected,
  setPasteProgress,
  setRange,
  sheets,
  storageKey,
}: HomeDataCommandsInputs) {
  const activeRange = (): string => {
    try {
      const [a, b = a] = range.toUpperCase().split(':');
      addressPosition(a);
      addressPosition(b);
      return rectangle(a, b);
    } catch {
      return selected;
    }
  };
  // Every Home invocation captures this render, while currentFrame observes
  // live ownership. Never combine an old DOM handler with a newer workbook.
  const homeSameTarget = (binding: HomeBinding): boolean => {
    const current = homeFrameRef.current?.binding;
    return (
      !!current &&
      current.documentId === binding.documentId &&
      current.ownerIdentity === binding.ownerIdentity &&
      current.sheet === binding.sheet &&
      current.selection === binding.selection &&
      current.selectedCell === binding.selectedCell &&
      current.activationEpoch === binding.activationEpoch &&
      current.navigationGeneration === binding.navigationGeneration &&
      current.modeGeneration === binding.modeGeneration &&
      current.specialSelectionEpoch === binding.specialSelectionEpoch &&
      current.mode === binding.mode
    );
  };
  const homeReadCells = async (
    lease: HomeLease,
    name: string,
    target: string,
    signal?: AbortSignal
  ): Promise<Record<string, WorkbookCell>> => {
    const area = selectedBounds(target, BULK_CELL_LIMIT),
      result: Record<string, WorkbookCell> = {};
    const reader = readRange,
      owner = sheets.find((item) => item.name === name);
    if (!owner) throw Error('The original worksheet is unavailable.');
    for (let col = area.first.col; col <= area.last.col; col += 1000) {
      const width = Math.min(1000, area.last.col - col + 1),
        step = Math.floor(1000 / width);
      for (let row = area.first.row; row <= area.last.row; row += step) {
        const lastRow = Math.min(area.last.row, row + step - 1),
          lastCol = col + width - 1,
          window = `${colName(col)}${row + 1}:${colName(lastCol)}${lastRow + 1}`;
        if (signal?.aborted) throw Error('The fill preview was cancelled.');
        const reply = await lease.read(() =>
          reader
            ? reader(name, window)
            : Promise.resolve(
                (() => {
                  const cells: Record<string, WorkbookCell> = {};
                  for (let r = row; r <= lastRow; r++)
                    for (let c = col; c <= lastCol; c++) {
                      const ref = `${colName(c)}${r + 1}`,
                        value = owner.cells[ref];
                      if (value !== undefined) cells[ref] = value;
                    }
                  return cells;
                })()
              )
        );
        if (signal?.aborted) throw Error('The fill preview was cancelled.');
        if (!reply || typeof reply !== 'object' || Array.isArray(reply))
          throw Error('The cell read did not return the original rectangle.');
        for (const [ref, cell] of Object.entries(reply)) {
          const p = addressPosition(ref);
          if (
            p.row < row ||
            p.row > lastRow ||
            p.col < col ||
            p.col > lastCol ||
            !cell ||
            typeof cell !== 'object' ||
            Array.isArray(cell)
          )
            throw Error('The cell read returned a foreign or malformed cell.');
          result[ref] = structuredClone(cell);
        }
      }
    }
    return result;
  };
  const homeFormatCommands = (
    lease: HomeLease,
    format: Extract<WorkbookCommand, { type: 'formatRange' }>['format']
  ): WorkbookCommand[] => {
    selectedBounds(lease.binding.selection, BULK_CELL_LIMIT);
    return lease.binding.specialSelection
      ? [...lease.binding.specialSelection].map((cell) => ({
          type: 'formatRange',
          sheet: lease.binding.sheet,
          range: cell,
          format,
        }))
      : [{ type: 'formatRange', sheet: lease.binding.sheet, range: lease.binding.selection, format }];
  };
  const clearHyperlinksDisabledReason = (): string | undefined => {
    if (!storageKey) return 'Clear Hyperlinks requires a native XLSX owner.';
    try {
      clearHyperlinksTargets(homeFrameRef.current!.binding);
    } catch (error) {
      return error instanceof Error ? error.message : String(error);
    }
  };
  const homeClipboardText = (target: string, cells: Record<string, WorkbookCell>): string => {
    const b = selectedBounds(target, BULK_CELL_LIMIT),
      text = serializeClipboardText(
        Array.from({ length: b.rows }, (_, row) =>
          Array.from({ length: b.columns }, (_, col) =>
            inputDisplay(cells[`${colName(b.first.col + col)}${b.first.row + row + 1}`])
          )
        )
      );
    if (pasteBytes(text) > 16 * 1024 * 1024) throw Error('Copy text exceeds 16 MiB. Copy a smaller rectangle.');
    return text;
  };
  const homePasteAsText = async (lease: HomeLease, text: string): Promise<void> => {
    if (!text) return;
    const own = internalClipboard.current;
    if (own?.text === text && own.cut)
      throw Error('Paste as text does not move a cut range. Use Paste to complete the move, or Copy the source first.');
    if (
      own?.text === text &&
      own.binding?.documentId === lease.binding.documentId &&
      own.binding?.ownerIdentity === lease.binding.ownerIdentity &&
      own.binding.state !== lease.binding.state
    )
      throw Error('The copied workbook changed. Copy the source again before pasting.');
    const paste = literalClipboardCells(text, { sheet: lease.binding.sheet, start: lease.binding.selectedCell });
    const commands = normalizeHomeBatch([{ type: 'setCells', sheet: lease.binding.sheet, cells: paste.cells }]);
    if (pasteBytes(JSON.stringify(commands)) > 2 * 1024 * 1024)
      throw Error('Paste as text exceeds its 2 MiB action limit. Split the text into smaller rectangles.');
    await lease.accept(commands);
    if (homeSameTarget(lease.binding)) setRange(paste.range);
  };
  const homePaste = async (lease: HomeLease, text: string): Promise<void> => {
    if (!text) return;
    const inspection = inspectPaste(text, {
      sheet: lease.binding.sheet,
      start: lease.binding.selectedCell,
      policy: literal ? 'text' : 'grid',
    });
    if (inspection.count > BULK_CELL_LIMIT) {
      if (!importPaste || literal)
        throw Error(
          'Paste supports up to 10,000 cells in this editor. Nothing was changed; split the data into smaller rectangles.'
        );
      const imported = importPaste,
        controller = new AbortController();
      pasteAbort.current = controller;
      setPasteProgress({ completed: 0, total: inspection.count });
      try {
        await lease.perform(() =>
          imported(text, inspection, {
            signal: controller.signal,
            onProgress: (completed, total) => {
              if (pasteAbort.current === controller && homeSameTarget(lease.binding))
                setPasteProgress({ completed, total });
            },
          })
        );
        if (homeSameTarget(lease.binding)) {
          const start = addressPosition(inspection.shape.start);
          setRange(
            rectangle(
              inspection.shape.start,
              `${colName(start.col + inspection.shape.columns - 1)}${start.row + inspection.shape.rows}`
            )
          );
        }
      } finally {
        if (pasteAbort.current === controller) {
          pasteAbort.current = null;
          setPasteProgress(null);
        }
      }
      return;
    }
    const rows = parseDelimited(text, '\t').rows,
      start = addressPosition(lease.binding.selectedCell),
      width = maxColumnCount(rows),
      count = rows.length * width;
    if (count > BULK_CELL_LIMIT || start.row + rows.length > 1048576 || start.col + width > 16384)
      throw Error('Paste target exceeds worksheet or action limits.');
    const target = rectangle(lease.binding.selectedCell, `${colName(start.col + width - 1)}${start.row + rows.length}`),
      targetCells = await homeReadCells(lease, lease.binding.sheet, target),
      own = internalClipboard.current;
    const sameClipboard =
      own?.text === text &&
      own.binding?.documentId === lease.binding.documentId &&
      own.binding?.ownerIdentity === lease.binding.ownerIdentity;
    if (own?.text === text && own.binding && sameClipboard && own.binding.state !== lease.binding.state)
      throw Error('The copied workbook changed. Copy the source again before pasting.');
    if (!literal && sameClipboard && own!.sheet !== lease.binding.sheet) {
      if (own!.cut) throw Error('Moving a range between worksheets is not yet qualified; copy it instead.');
      const source = await homeReadCells(lease, own!.sheet, own!.range);
      if (Object.values(source).some((cell) => cell.formula || cell.formulaType))
        throw Error('Cross-sheet formula copy is not yet qualified; use Paste Values.');
    }
    if (!literal && sameClipboard && own!.sheet === lease.binding.sheet) {
      if (rangesOverlap(own!.range, target)) throw Error('Paste target overlaps its source; choose a separate range.');
      const commands: WorkbookCommand[] = [
        { type: 'fillRange', sheet: lease.binding.sheet, source: own!.range, range: target },
      ];
      if (own!.cut) {
        const b = selectedBounds(own!.range, BULK_CELL_LIMIT);
        commands.push(
          storageKey
            ? { type: 'clearRange', sheet: lease.binding.sheet, range: own!.range }
            : {
                type: 'setCells',
                sheet: lease.binding.sheet,
                cells: Array.from({ length: b.rows * b.columns }, (_, i) => ({
                  cell: `${colName(b.first.col + (i % b.columns))}${b.first.row + Math.floor(i / b.columns) + 1}`,
                  value: null,
                })),
              }
        );
      }
      await lease.accept(commands);
      if (own!.cut && internalClipboard.current === own) internalClipboard.current = null;
    } else {
      if (literal && sameClipboard && own!.cut && rangesOverlap(own!.range, target))
        throw Error('Paste target overlaps its source; choose a separate range.');
      const cells = rows.flatMap((row, r) =>
        row.map((value, c) => {
          const cell = `${colName(start.col + c)}${start.row + r + 1}`;
          return {
            cell,
            ...parseGridInput(
              value,
              !!literal || targetCells[cell]?.format === 49 || targetCells[cell]?.format === '@'
            ),
          };
        })
      );
      if (literal && sameClipboard && own!.cut) {
        const source = selectedBounds(own!.range, BULK_CELL_LIMIT);
        for (let row = source.first.row; row <= source.last.row; row++)
          for (let col = source.first.col; col <= source.last.col; col++)
            cells.push({ cell: `${colName(col)}${row + 1}`, value: null });
        await lease.accept([{ type: 'setCells', sheet: lease.binding.sheet, cells }]);
        if (internalClipboard.current === own) internalClipboard.current = null;
      } else await lease.accept([{ type: 'setCells', sheet: lease.binding.sheet, cells }]);
    }
    if (homeSameTarget(lease.binding)) setRange(target);
  };
  const homePasteSpecial = async (lease: HomeLease, mode: 'values' | 'formulas' | 'formatting'): Promise<void> => {
    const source = internalClipboard.current;
    if (!source || source.cut) throw Error('Copy a range before using Paste Special.');
    if (
      !source.binding ||
      source.binding.documentId !== lease.binding.documentId ||
      source.binding.ownerIdentity !== lease.binding.ownerIdentity ||
      source.binding.state !== lease.binding.state
    )
      throw Error('The copied workbook changed. Copy the source again.');
    const b = selectedBounds(source.range, 1000),
      start = addressPosition(lease.binding.selectedCell);
    if (start.row + b.rows > 1048576 || start.col + b.columns > 16384)
      throw Error('Paste target is outside worksheet limits.');
    const target = rectangle(lease.binding.selectedCell, `${colName(start.col + b.columns - 1)}${start.row + b.rows}`);
    if (mode === 'values' && storageKey)
      await lease.accept([
        {
          type: 'pasteValues',
          sheet: lease.binding.sheet,
          sourceSheet: source.sheet,
          source: source.range,
          range: target,
        },
      ]);
    else if (mode === 'formulas') {
      if (source.sheet !== lease.binding.sheet) throw Error('Paste Formulas currently requires the same worksheet.');
      if (rangesOverlap(source.range, target)) throw Error('Paste Formulas target overlaps the copied range.');
      await lease.accept([{ type: 'pasteFormulas', sheet: lease.binding.sheet, source: source.range, range: target }]);
    } else {
      const values = await homeReadCells(lease, source.sheet, source.range),
        cells: Array<Omit<WorkbookEdit, 'sheet'>> = [],
        formats: WorkbookCommand[] = [];
      const sourceSheet = metadata.find((item) => item.name === source.sheet),
        sourceContext = mode === 'values' && sourceSheet ? createValuePasteSourceContext(sourceSheet) : null;
      if (mode === 'values' && !sourceContext) throw Error('The complete accepted Paste Values source is unavailable.');
      for (let row = 0; row < b.rows; row++)
        for (let col = 0; col < b.columns; col++) {
          const sourceRef = `${colName(b.first.col + col)}${b.first.row + row + 1}`,
            from = values[sourceRef],
            cell = `${colName(start.col + col)}${start.row + row + 1}`;
          if (mode === 'formatting')
            formats.push({
              type: 'formatRange',
              sheet: lease.binding.sheet,
              range: cell,
              format: editableCellFormat(from),
            });
          else cells.push({ cell, ...normalizeValuePasteCell(from, { ...sourceContext!, cell: sourceRef }) });
        }
      await lease.accept(mode === 'formatting' ? formats : [{ type: 'setCells', sheet: lease.binding.sheet, cells }]);
    }
    if (homeSameTarget(lease.binding)) setRange(target);
  };
  return {
    activeRange,
    homeSameTarget,
    homeReadCells,
    homeFormatCommands,
    clearHyperlinksDisabledReason,
    homeClipboardText,
    homePasteAsText,
    homePaste,
    homePasteSpecial,
  } as const;
}
