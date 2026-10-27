import { type WorkbookCell, type WorkbookCommand, type WorkbookEdit } from '../../core/contracts';

export function editableNumberFormat(format: string | number | undefined): WorkbookEdit['numberFormat'] {
  if (format === 0 || format === undefined) return 'general';
  if (format === 1) return 'decimal0';
  if (format === 2) return 'number';
  if (format === 3) return 'thousands';
  if (format === 10) return 'percent';
  if (format === 14) return 'date';
  if (format === 44) return 'accounting';
  if (format === 49) return 'text';
  if (format === '0.0') return 'decimal1';
  if (typeof format === 'string' && format.includes('$')) return 'currency';
  return undefined;
}
export function editableCellFormat(
  cell: WorkbookCell | undefined
): Extract<WorkbookCommand, { type: 'formatRange' }>['format'] {
  return {
    styleId: cell?.styleId ?? 0,
    bold: cell?.bold,
    italic: cell?.italic,
    underline: cell?.underline,
    fontFamily: cell?.fontFamily,
    fontSize: cell?.fontSize,
    fontColor: cell?.fontColor,
    fillColor: cell?.fillColor,
    horizontalAlign: cell?.horizontalAlign,
    verticalAlign: cell?.verticalAlign,
    wrapText: cell?.wrapText,
    border: cell?.border,
    numberFormat: editableNumberFormat(cell?.format),
  };
}
