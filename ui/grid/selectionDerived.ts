import { BULK_CELL_LIMIT, FILTER_CELL_LIMIT } from './gridLimits';
import type { GridFilterValues } from './gridTypes';
import type { WorkbookNormalFont, WorkbookSheet, WorkbookValidationDefinition } from '../../core/contracts';
import { addressPosition, colName, formulaCount } from '../sheetModel';
import { ruleContains } from '../../core/rules-model.mjs';
import { selectedBounds } from './rangeAddress';
import { useWorksheetGeometryProfile } from '../worksheetFontMetrics';

export interface SelectionDerivedInputs {
  acceptedValidation: { scope: string; definition: WorkbookValidationDefinition } | null;
  activeRange: () => string;
  filterValues: GridFilterValues;
  normalFont: Readonly<WorkbookNormalFont>;
  selected: string;
  sheet: WorkbookSheet;
  sheets: WorkbookSheet[];
  specialSelection: Set<string> | null;
  structuralReason: string | undefined;
  validationScope: string;
  words: (en: string) => string;
}

export function useSelectionDerived({
  acceptedValidation,
  activeRange,
  filterValues,
  normalFont,
  selected,
  sheet,
  sheets,
  specialSelection,
  structuralReason,
  validationScope,
  words,
}: SelectionDerivedInputs) {
  const structuralBlock =
    structuralReason ||
    (sheets.some((s) => s.merges.length || formulaCount(s) > 0)
      ? words('Remove formulas or merged cells before inserting or deleting rows/columns.')
      : '');
  const selectedCell = sheet.cells[selected];
  const selectedValidation =
    acceptedValidation?.scope === validationScope
      ? acceptedValidation.definition.rules.find((rule) => ruleContains(rule, selected))
      : undefined;
  const firstSpecialCell = specialSelection?.values().next().value;
  const styleCell = firstSpecialCell ? sheet.cells[firstSpecialCell] : selectedCell;
  const selectedFontProfile = useWorksheetGeometryProfile({
    normalFont: {
      fontFamily: styleCell?.fontFamily ?? normalFont.fontFamily,
      fontSize: styleCell?.fontSize ?? normalFont.fontSize,
      bold: styleCell?.bold ?? normalFont.bold,
      italic: styleCell?.italic ?? normalFont.italic,
    },
  });
  const sortBounds = (() => {
    try {
      return selectedBounds(activeRange(), BULK_CELL_LIMIT);
    } catch {
      return null;
    }
  })();
  const filterBounds = (() => {
    try {
      return selectedBounds(activeRange(), filterValues ? FILTER_CELL_LIMIT : BULK_CELL_LIMIT);
    } catch {
      return null;
    }
  })();
  const sortColumns = sortBounds
    ? Array.from({ length: sortBounds.columns }, (_, index) => colName(sortBounds.first.col + index))
    : [];
  const filterColumn = colName(addressPosition(selected).col);
  const distinctFilterValues: Array<{ key: string; value: string | number | boolean | null }> = [];
  if (!filterValues && sortBounds && sortBounds.rows > 1 && sortBounds.columns <= 30) {
    const seen = new Set<string>();
    for (let row = sortBounds.first.row + 1; row <= sortBounds.last.row; row++) {
      const cell = sheet.cells[`${filterColumn}${row + 1}`];
      const value = cell?.formula
        ? cell.calculation?.status === 'calculated'
          ? (cell.calculation.value ?? null)
          : null
        : (cell?.value ?? null);
      const key = JSON.stringify(value);
      if (!seen.has(key)) {
        seen.add(key);
        distinctFilterValues.push({ key, value });
      }
      if (distinctFilterValues.length > 50) break;
    }
  }
  return {
    structuralBlock,
    selectedCell,
    selectedValidation,
    styleCell,
    selectedFontProfile,
    sortBounds,
    filterBounds,
    sortColumns,
    filterColumn,
    distinctFilterValues,
  } as const;
}
