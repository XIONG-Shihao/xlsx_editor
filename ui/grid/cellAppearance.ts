import { type WorkbookSheet } from '../../core/contracts';
import { addressPosition, colName } from '../sheetModel';
import { inRange, selectedBounds } from './rangeAddress';

export const duplicateRuleCounts = new WeakMap<WorkbookSheet, Map<string, Map<string, number>>>();
export function tableAppearance(
  sheet: WorkbookSheet,
  ref: string
): { background: string; color?: string; header: boolean } | undefined {
  const position = addressPosition(ref);
  for (const table of sheet.tableRanges ?? []) {
    if (!inRange(table.ref, ref)) continue;
    const first = addressPosition(table.ref.split(':')[0]);
    if (position.row === first.row) return { background: '#4472C4', color: '#FFFFFF', header: true };
    return { background: (position.row - first.row) % 2 === 0 ? '#D9E5F3' : '#FFFFFF', header: false };
  }
  return undefined;
}
export function conditionalFill(sheet: WorkbookSheet, ref: string): string | undefined {
  for (const rule of sheet.conditionalRules ?? []) {
    if (!inRange(rule.range, ref)) continue;
    const cell = sheet.cells[ref];
    const value = cell?.formula
      ? cell.calculation?.status === 'calculated'
        ? cell.calculation.value
        : null
      : cell?.value;
    if (rule.rule === 'greaterThan' && typeof value === 'number' && value > (rule.threshold ?? Infinity))
      return rule.fillColor;
    if (rule.rule === 'lessThan' && typeof value === 'number' && value < (rule.threshold ?? -Infinity))
      return rule.fillColor;
    if (rule.rule === 'duplicateValues' && value !== undefined && value !== null && value !== '') {
      if (sheet.duplicateValues?.[rule.range]) {
        if (sheet.duplicateValues[rule.range].includes(JSON.stringify(value))) return rule.fillColor;
        continue;
      }
      let byRange = duplicateRuleCounts.get(sheet);
      if (!byRange) {
        byRange = new Map();
        duplicateRuleCounts.set(sheet, byRange);
      }
      let counts = byRange.get(rule.range);
      if (!counts) {
        counts = new Map();
        try {
          const bounds = selectedBounds(rule.range, 10000);
          for (let row = 0; row < bounds.rows; row++)
            for (let col = 0; col < bounds.columns; col++) {
              const candidate = sheet.cells[`${colName(bounds.first.col + col)}${bounds.first.row + row + 1}`];
              const entry = candidate?.formula
                ? candidate.calculation?.status === 'calculated'
                  ? candidate.calculation.value
                  : null
                : candidate?.value;
              if (entry !== undefined && entry !== null && entry !== '') {
                const key = JSON.stringify(entry);
                counts.set(key, (counts.get(key) ?? 0) + 1);
              }
            }
        } catch {
          return undefined;
        }
        byRange.set(rule.range, counts);
      }
      if ((counts.get(JSON.stringify(value)) ?? 0) > 1) return rule.fillColor;
    }
  }
  return undefined;
}
