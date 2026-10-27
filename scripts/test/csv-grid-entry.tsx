import './grid-host'
import React, { useState } from 'react';

import { createRoot } from 'react-dom/client';
import { SpreadsheetGrid } from '@office-editors/xlsx-editor/ui/SpreadsheetGrid';
import { SheetHistory } from '@office-editors/xlsx-editor/ui/sheetHistory';
import { csvSheet } from '@office-editors/xlsx-editor/ui/csvModel';
import { WorkbookWorker } from '@office-editors/xlsx-editor/ui/workbookWorker';
import { indexCsv, applyCsvCommands, serializeCsv } from '@office-editors/xlsx-editor/core/csv-workbook.mjs';
import type { WorkbookView, WorkbookCommand } from '@office-editors/xlsx-editor/contracts';
const indexed = indexCsv('id,note\r\n' + Array.from({ length: 130000 }, (_, i) => `${i},"original-${i}"`).join('\r\n') + '\r\n'), history = new SheetHistory([csvSheet(indexed)]);
history.current.csvIndex = indexed;
const scope = globalThis as any;
function Host() {
  const [frame, setFrame] = useState(history.current), [error, setError] = useState('');
  scope.__csv = () => ({ rows: frame.sheets[0].rows, tail: frame.sheets[0].cells.B130001.value, edits: frame.edits, serializedTail: serializeCsv(frame.csvIndex!).slice(-100), error });
  const apply = (commands: WorkbookCommand[]) => {
    try {
      const result = applyCsvCommands(history.current.csvIndex!, commands), next = history.batch(result.edits, [csvSheet(result.index)]);
      next.csvIndex = result.index;
      setFrame(next);
    }
    catch (e) {
      setError(String(e));
    }
  };
  return <><SpreadsheetGrid literal sheets={frame.sheets} readOnly={false} onEdit={edit => apply([edit])} onBatch={apply} onUndo={() => setFrame(history.undo())} onRedo={() => setFrame(history.redo())} canUndo={history.canUndo} canRedo={history.canRedo} /><p role="alert">{error}</p></>;
}
createRoot(document.getElementById('root')!).render(<Host />);
scope.__workerProbe = async () => {
  const cells: Record<string, any> = {};
  for (let i = 0; i < 50000; i++)
    cells[`A${i + 1}`] = { value: i, dataType: 'number', rawValue: String(i) };
  cells.B2 = { value: '#DIV/0!', dataType: 'error', error: '#DIV/0!' };
  const base: WorkbookView = { revision: 'original', readOnly: false, limitations: [], calculation: '', sheets: [{ name: 'Main', state: 'visible', path: '', cells, rows: 50000, columns: 2, widths: [], merges: [] }] }, worker = new WorkbookWorker();
  try {
    const prepared = await worker.prepare(base);
    if (prepared.sheets !== base.sheets)
      throw Error('Formula-free preparation should retain renderer model');
    const edited = await worker.apply(prepared, [{ sheet: 'Main', cell: 'A50000', value: 7 }]);
    const formula = await worker.apply(edited.view, [{ sheet: 'Main', cell: 'B50000', formula: 'A50000+2' }]);
    const dependent = await worker.apply(formula.view, [{ sheet: 'Main', cell: 'A50000', value: 10 }]);
    const undone = await worker.apply(edited.view, [{ sheet: 'Main', cell: 'B50000', formula: 'A50000+3' }]);
    return { rows: dependent.view.sheets[0].rows, preparedRevision: prepared.revision, originalTail: base.sheets[0].cells.A50000.value, firstFormula: formula.view.sheets[0].cells.B50000.calculation, dependent: dependent.view.sheets[0].cells.B50000.calculation, historyResync: undone.view.sheets[0].cells.B50000.calculation, errorPreserved: dependent.view.sheets[0].cells.B2, unchangedIdentity: edited.view.sheets[0].cells.A1 === base.sheets[0].cells.A1 };
  }
  finally {
    worker.dispose();
  }
};
