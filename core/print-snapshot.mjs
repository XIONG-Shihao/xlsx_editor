/** Readonly immutable accepted print snapshot, bounded before host publication. */
import { createHash } from 'node:crypto';
import { displayCell } from './display.mjs';
import { rangeBounds, columnName } from './workbook-commands.mjs';
import { indexedPrintDefinition, indexedPrintGeometry } from './indexed-print.mjs';
import { PRINT_LIMITS, validatePrintSettings } from './print-model.mjs';
import { indexedRulesDefinition, indexedRuleDuplicates } from './indexed-rules.mjs';
import { conditionalFill, acceptedRuleValue, ruleContains } from './rules-model.mjs';
const encoded = (value) => Buffer.byteLength(JSON.stringify(value));
function refuse(code, details = {}) {
  const error = Error(code.replaceAll('_', ' '));
  error.code = code;
  Object.assign(error, details);
  throw error;
}
function styleOf(cell, ref) {
  const style = {};
  for (const key of [
    'bold',
    'italic',
    'underline',
    'fontFamily',
    'fontSize',
    'fontColor',
    'fillColor',
    'horizontalAlign',
    'verticalAlign',
    'wrapText',
    'border',
  ])
    if (cell?.[key] !== undefined) style[key] = cell[key];
  if (
    style.fontFamily !== undefined &&
    (typeof style.fontFamily !== 'string' || !/^[\p{L}\p{N} ._-]{1,80}$/u.test(style.fontFamily))
  )
    refuse('PRINT_UNSUPPORTED_FONT', { cell: ref });
  if (style.fontSize !== undefined && (!Number.isFinite(style.fontSize) || style.fontSize < 1 || style.fontSize > 72))
    refuse('PRINT_UNSUPPORTED_FONT_SIZE', { cell: ref });
  for (const key of ['fontColor', 'fillColor'])
    if (style[key] !== undefined && !/^#?[0-9a-fA-F]{6}$/.test(style[key]))
      refuse('PRINT_UNSUPPORTED_COLOR', { cell: ref });
  if (style.horizontalAlign === 'general') delete style.horizontalAlign;
  if (
    (style.horizontalAlign !== undefined && !['left', 'center', 'right'].includes(style.horizontalAlign)) ||
    (style.verticalAlign !== undefined && !['top', 'center', 'bottom'].includes(style.verticalAlign)) ||
    (style.border !== undefined && !['none', 'all'].includes(style.border))
  )
    refuse('PRINT_UNSUPPORTED_ALIGNMENT_BORDER', { cell: ref });
  for (const key of ['bold', 'italic', 'underline', 'wrapText'])
    if (style[key] !== undefined && typeof style[key] !== 'boolean') refuse('PRINT_UNSUPPORTED_STYLE', { cell: ref });
  const value = cell?.formula ? cell.calculation?.value : cell?.value;
  style.horizontalAlign ??= typeof value === 'number' ? 'right' : typeof value === 'boolean' ? 'center' : 'left';
  return style;
}
function display(cell, sheet, ref, date1904, ownership) {
  const owned =
    ownership?.kind === 'consumer'
      ? (cell?.calculation?.spillConsumer || cell?.calculation?.scalarArray?.schema === 1) &&
        JSON.stringify(ownership.calculation) === JSON.stringify(cell.calculation)
      : ['master', 'follower'].includes(ownership?.kind) &&
        cell?.spill &&
        JSON.stringify(ownership.spill) === JSON.stringify(cell.spill);
  if ((((cell?.formulaType && cell.formulaType !== 'shared') || cell?.spill) && !owned) || cell?.scenarioResult)
    refuse('PRINT_SPECIAL_FORMULA', { cell: sheet + '!' + ref });
  if (cell?.formula) {
    const calculation = cell.calculation;
    if (!calculation || !['calculated', 'error', 'circular'].includes(calculation.status))
      refuse('PRINT_UNRESOLVED_CALCULATION', { cell: sheet + '!' + ref });
    if (
      calculation.status === 'calculated' &&
      calculation.value !== undefined &&
      (!['string', 'number', 'boolean'].includes(typeof calculation.value) ||
        (typeof calculation.value === 'number' && !Number.isFinite(calculation.value)))
    )
      refuse('PRINT_UNQUALIFIED_CALCULATION', { cell: sheet + '!' + ref });
  }
  const value = displayCell(cell, date1904);
  if (value.length > 32767) refuse('PRINT_TEXT_BOUND', { cell: sheet + '!' + ref });
  return value;
}
export async function createPrintSnapshot({
  sessionId,
  state,
  revision,
  metadataBytes,
  view,
  sheet,
  settings: requested,
  geometryProfile,
  readRange,
  objects,
  object,
  qualifyCells,
  check = () => {},
}) {
  check();
  const definition = indexedPrintDefinition(metadataBytes, view, sheet);
  if (!definition.canEdit) refuse('PRINT_PROTECTED_SETTINGS', { reasons: definition.protectedReasons });
  const settings = validatePrintSettings(requested ?? definition.settings),
    geometry = indexedPrintGeometry(metadataBytes, view, sheet, settings, geometryProfile),
    source = view.sheets.find((item) => item.name === sheet),
    bounds = rangeBounds(settings.area),
    cells = [];
  let size = 0;
  // Read the entire closed rule collection from this immutable owner's metadata.
  // Duplicate counts cover their declared source ranges, including off-page cells.
  const conditional = indexedRulesDefinition(metadataBytes, view, sheet, 'conditional');
  if (!conditional.canEdit)
    refuse('PRINT_UNQUALIFIED_CONDITIONAL_FORMATTING', { reasons: conditional.protectedReasons });
  let duplicates;
  try {
    duplicates = await indexedRuleDuplicates(conditional.rules, sheet, readRange, check);
  } catch (error) {
    check();
    refuse('PRINT_UNQUALIFIED_CONDITIONAL_FORMATTING', { reasons: [String(error.message ?? error)] });
  }
  check();
  if (typeof qualifyCells !== 'function') refuse('PRINT_FORMAT_QUALIFICATION_REQUIRED');
  const hiddenRows = new Set(geometry.rows.filter((item) => item.hidden).map((item) => item.index)),
    hiddenColumns = new Set(geometry.columns.filter((item) => item.hidden).map((item) => item.index)),
    columnsPerRead = Math.min(1000, bounds.columns),
    rowsPerRead = Math.max(1, Math.floor(1000 / columnsPerRead));
  for (let row = bounds.first.row; row <= bounds.last.row; row += rowsPerRead)
    for (let col = bounds.first.col; col <= bounds.last.col; col += columnsPerRead) {
      const lastRow = Math.min(bounds.last.row, row + rowsPerRead - 1),
        lastCol = Math.min(bounds.last.col, col + columnsPerRead - 1),
        range = columnName(col) + (row + 1) + ':' + columnName(lastCol) + (lastRow + 1);
      check();
      const proof = await qualifyCells(range);
      check();
      if (proof?.qualified !== true) refuse('PRINT_FORMAT_QUALIFICATION_REQUIRED');
      if (
        proof.normalFont &&
        (!geometry.geometryProfile?.font ||
          Object.keys(proof.normalFont).sort().join(',') !== 'bold,fontFamily,fontSize,italic' ||
          Object.keys(proof.normalFont).some((key) => proof.normalFont[key] !== geometry.geometryProfile.font[key]))
      )
        refuse('PRINT_UNRESOLVED_GEOMETRY', {
          reasons: ['Native theme Normal font differs from the accepted geometry.'],
        });
      if (
        proof.pageTheme &&
        (Object.keys(proof.pageTheme).sort().join(',') !== 'revision,sample' ||
          proof.pageTheme.sample !== view.pageTheme?.sample ||
          proof.pageTheme.revision !== view.pageTheme?.revision)
      )
        refuse('PRINT_UNSUPPORTED_THEME', { reasons: ['Native theme identity differs from the accepted owner.'] });
      const reply = await readRange(sheet, range);
      check();
      for (let r = row; r <= lastRow; r++)
        for (let c = col; c <= lastCol; c++) {
          const ref = columnName(c) + (r + 1),
            cell = reply.cells?.[ref],
            visible = !hiddenRows.has(r) && !hiddenColumns.has(c),
            text = visible ? display(cell, sheet, ref, !!source.date1904, proof.specialCells?.[ref]) : '',
            printed = { ref, text, style: visible ? styleOf(cell, ref) : {} };
          if (visible && conditional.rules.some((rule) => ruleContains(rule, ref))) {
            if (!acceptedRuleValue(cell).qualified)
              refuse('PRINT_UNQUALIFIED_CONDITIONAL_FORMATTING', { cell: sheet + '!' + ref });
            const fill = conditionalFill(conditional.rules, ref, cell, duplicates);
            if (fill) printed.style.fillColor = fill;
          }
          if (printed.text.length > 32767) refuse('PRINT_TEXT_BOUND', { cell: sheet + '!' + ref });
          size += encoded(printed);
          if (size > PRINT_LIMITS.snapshotBytes) refuse('PRINT_SNAPSHOT_BOUND');
          cells.push(printed);
        }
    }
  const previews = [],
    width = geometry.columns.filter((item) => !item.hidden).reduce((sum, item) => sum + item.widthPt, 0),
    height = geometry.rows.filter((item) => !item.hidden).reduce((sum, item) => sum + item.heightPt, 0);
  let after;
  do {
    check();
    const page = await objects(sheet, { after, limit: 50 });
    check();
    for (const entry of page.objects) {
      if (!entry.anchor) refuse('PRINT_UNSUPPORTED_OBJECT', { objectId: entry.id });
      const box = geometry.objectBox(entry.anchor);
      if (!box || box[0] >= width || box[1] >= height || box[0] + box[2] <= 0 || box[1] + box[3] <= 0) continue;
      if (entry.protectedReason) refuse('PRINT_UNSUPPORTED_OBJECT', { objectId: entry.id });
      if (previews.length >= PRINT_LIMITS.objects) refuse('PRINT_OBJECT_LIMIT');
      const preview = await object(sheet, entry.id);
      check();
      const frozen = { ...preview, boxPt: box };
      size += encoded(frozen);
      if (size > PRINT_LIMITS.snapshotBytes) refuse('PRINT_SNAPSHOT_BOUND');
      previews.push(frozen);
    }
    after = page.nextCursor ?? undefined;
  } while (after !== undefined);
  const snapshot = {
    schema: 1,
    sessionId,
    state: String(state),
    revision,
    sheet,
    date1904: !!source.date1904,
    settings,
    definitionRevision: definition.revision,
    conditionalRevision: conditional.revision,
    geometryProfile: geometry.geometryProfile,
    rows: geometry.rows,
    columns: geometry.columns,
    merges: geometry.merges,
    cells,
    objects: previews,
  };
  const json = JSON.stringify(snapshot);
  if (Buffer.byteLength(json) > PRINT_LIMITS.snapshotBytes) refuse('PRINT_SNAPSHOT_BOUND');
  snapshot.fingerprint = createHash('sha256').update(json).digest('hex');
  check();
  return snapshot;
}
