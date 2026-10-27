import { rangeBounds, assertXmlText } from './workbook-commands.mjs';
import { refusal } from './refusal-codes.mjs';
export const LOCAL_EXECUTION_ACTIONS = Object.freeze(['unhide-all', 'freeze-selection']);
/** Built-ins are typed commands, never private stored definitions or arbitrary code. */
export function localExecutionPreview(action, scope) {
  if (!LOCAL_EXECUTION_ACTIONS.includes(action)) throw refusal('query-bad-input', 'Choose a supported local action');
  if (
    !scope ||
    Object.keys(scope).sort().join(',') !== 'selection,sheet' ||
    typeof scope.sheet !== 'string' ||
    !scope.sheet.trim() ||
    scope.sheet.length > 31 ||
    /[\\/:?*\[\]]/.test(scope.sheet) ||
    typeof scope.selection !== 'string' ||
    scope.selection.length > 64
  )
    throw refusal('query-malformed', 'Invalid local action target');
  assertXmlText(scope.sheet);
  const bounds = rangeBounds(scope.selection),
    cell = bounds.first;
  if (action === 'unhide-all')
    return {
      action,
      label: 'Unhide Rows and Columns',
      sheet: scope.sheet,
      selection: scope.selection,
      commands: [{ operation: 'unhideAll', sheet: scope.sheet }],
    };
  if (cell.row > 100 || cell.col > 30)
    throw refusal('query-limit', 'Freeze Selection supports up to 100 rows and 30 columns. Choose a nearer cell.');
  if (cell.row === 0 && cell.col === 0)
    throw refusal('query-bad-input', 'Choose a cell below or right of A1 to freeze preceding rows or columns.');
  return {
    action,
    label: 'Freeze Selection',
    sheet: scope.sheet,
    selection: scope.selection,
    commands: [{ operation: 'freeze', sheet: scope.sheet, rows: cell.row, columns: cell.col }],
  };
}
