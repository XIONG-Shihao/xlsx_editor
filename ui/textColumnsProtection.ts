/** Known selected intersections refuse; formulas elsewhere are
 * recalculated/admitted by the existing strict owner command. No graph guessing. */
import type {
  WorkbookSheet,
  WorkbookView,
  WorkbookConditionalDefinition,
  WorkbookValidationDefinition,
} from '../core/contracts';
function bounds(range: string) {
  const pieces = range.replaceAll('$', '').toUpperCase().split(':');
  const cell = (value: string) => {
    const match = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(value);
    if (!match) throw Error('Unknown range');
    const col = [...match[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1,
      row = Number(match[2]) - 1;
    if (col >= 16384 || row >= 1048576) throw Error('Unknown range');
    return { row, col };
  };
  if (pieces.length > 2) throw Error('Unknown range');
  const first = cell(pieces[0]),
    last = cell(pieces[1] ?? pieces[0]);
  if (first.row > last.row || first.col > last.col) throw Error('Unknown range');
  return { first, last };
}
function intersects(a: string, b: string): boolean {
  const x = bounds(a),
    y = bounds(b);
  return (
    x.first.row <= y.last.row && x.last.row >= y.first.row && x.first.col <= y.last.col && x.last.col >= y.first.col
  );
}
export function textColumnProtection(
  input: {
    sheet: WorkbookSheet;
    tables: NonNullable<WorkbookView['tables']>;
    pivots: NonNullable<WorkbookView['pivots']>;
    conditional: WorkbookConditionalDefinition;
    validation: WorkbookValidationDefinition;
  },
  source: string,
  destination: string
): { reason: string | null } {
  const hit = (range: string) => intersects(range, source) || intersects(range, destination),
    unknown = () => ({ reason: 'The workbook protection inventory is unavailable or unsupported.' });
  try {
    if (
      !input ||
      !input.sheet ||
      !Array.isArray(input.sheet.merges) ||
      !Array.isArray(input.tables) ||
      !Array.isArray(input.pivots)
    )
      return unknown();
    if (input.sheet.merges.some(hit)) return { reason: 'The source or destination intersects merged cells.' };
    if (input.sheet.filter && hit(input.sheet.filter.range))
      return { reason: 'The source or destination intersects a filtered range.' };
    for (const table of input.tables) {
      if (typeof table.sheet !== 'string' || typeof table.ref !== 'string') return unknown();
      if (table.sheet === input.sheet.name && hit(table.ref))
        return { reason: 'The source or destination intersects a table.' };
    }
    for (const pivot of input.pivots) {
      if (
        typeof pivot.sheet !== 'string' ||
        typeof pivot.sourceSheet !== 'string' ||
        !pivot.outputRange ||
        !pivot.sourceRange
      )
        return unknown();
      bounds(pivot.outputRange);
      bounds(pivot.sourceRange);
      if (
        (pivot.sheet === input.sheet.name && pivot.outputRange && hit(pivot.outputRange)) ||
        (pivot.sourceSheet === input.sheet.name && pivot.sourceRange && hit(pivot.sourceRange))
      )
        return { reason: 'The source or destination intersects a pivot source or output.' };
    }
    for (const definition of [input.conditional, input.validation]) {
      if (
        !definition ||
        definition.sheet !== input.sheet.name ||
        typeof definition.canEdit !== 'boolean' ||
        !Array.isArray(definition.rules) ||
        !Array.isArray(definition.protectedReasons) ||
        definition.protectedReasons.some((reason) => typeof reason !== 'string') ||
        (definition.evaluationReasons !== undefined &&
          (!Array.isArray(definition.evaluationReasons) ||
            definition.evaluationReasons.some((reason) => typeof reason !== 'string')))
      )
        return unknown();
      // Definition editing and duplicate-render evaluation are separate authorities.
      // A protected collection is excludable only with complete trusted ranges.
      if ((!definition.canEdit || definition.protectedReasons.length) && !Array.isArray(definition.protectedRanges))
        return unknown();
      if (definition.protectedRanges !== undefined && !Array.isArray(definition.protectedRanges)) return unknown();
      const protectedRanges = definition.protectedRanges ?? [];
      if (definition.rules.some((rule) => hit(rule.range)) || protectedRanges.some(hit))
        return { reason: 'The source or destination intersects conditional formatting or input rules.' };
    }
    return { reason: null };
  } catch {
    return unknown();
  }
}
