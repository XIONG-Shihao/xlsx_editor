import type { WorkbookConditionalDefinition, WorkbookSheet, WorkbookView } from '../../core/contracts';
import { createContext, useContext } from 'react';
import { displayValue } from '../sheetModel';
import { useDocumentWords } from '../words';

export const OwnedPivotPresentation = createContext<NonNullable<WorkbookView['pivots']>>([]);
export const BooleanCellControls = createContext<{
  disabled: boolean;
  toggle: (sheet: string, cell: string, value: boolean) => Promise<void>;
} | null>(null);
export const CalculationRevision = createContext<string | undefined>(undefined);
export const ConditionalDefinition = createContext<WorkbookConditionalDefinition | null | undefined>(undefined);
export function CalculatedCell({ sheet, cellRef }: { sheet: WorkbookSheet; cellRef: string }): JSX.Element {
  const revision = useContext(CalculationRevision),
    words = useDocumentWords(),
    cell = sheet.cells[cellRef];
  return (
    <small
      className="sheet-a11y-status"
      role="status"
      aria-label={`${sheet.name} ${cellRef} calculated result`}
      data-status={cell?.calculation?.status ?? (cell?.formula || cell?.formulaType ? 'pending' : 'literal')}
      data-revision={revision ?? cell?.calculation?.revision}
      data-value={cell?.calculation?.status === 'calculated' ? JSON.stringify(cell.calculation.value) : undefined}
      title={cell?.calculation?.pending ? 'Stored result; pending calculation' : cell?.calculation?.error}
    >
      {words(String(displayValue(sheet, cellRef)))}
    </small>
  );
}
