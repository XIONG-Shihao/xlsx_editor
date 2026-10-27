/** Home routing only: builders and host I/O remain with the captured Grid.
 * Native/owner admission remains authoritative for every workbook mutation. */
import type { WorkbookCommand, WorkbookEdit, WorkbookFillAcrossSheetsMode } from '../core/contracts';
import type { RibbonBinding, RibbonCapability, RibbonProfile } from './ribbonCommands';

export type NoArgumentHomeId =
  | 'home.copy'
  | 'home.cut'
  | 'home.paste-values'
  | 'home.paste-formulas'
  | 'home.paste-formatting'
  | 'home.format-painter'
  | 'home.cancel-paste'
  | 'home.font-size-increase'
  | 'home.font-size-decrease'
  | 'home.align-top'
  | 'home.align-middle'
  | 'home.align-bottom'
  | 'home.align-left'
  | 'home.align-center'
  | 'home.align-right'
  | 'home.merge-center'
  | 'home.unmerge'
  | 'home.accounting'
  | 'home.percent'
  | 'home.thousands'
  | 'home.decimals-increase'
  | 'home.decimals-decrease'
  | 'styles.conditional'
  | 'home.style-good'
  | 'home.style-bad'
  | 'home.style-heading'
  | 'home.insert-row'
  | 'home.insert-column'
  | 'home.delete-row'
  | 'home.delete-column'
  | 'home.fill-down'
  | 'home.fill-right'
  | 'home.fill-up'
  | 'home.fill-left'
  | 'home.fill-series'
  | 'home.fill-series-apply'
  | 'home.fill-across'
  | 'home.clear-contents'
  | 'home.clear-formats'
  | 'home.clear-all'
  | 'home.clear-hyperlinks'
  | 'home.find-open'
  | 'home.find-search'
  | 'home.find-cancel'
  | 'home.find-next'
  | 'home.replace-preview'
  | 'home.replace-one'
  | 'home.replace-all'
  | 'home.find-close'
  | 'home.special-open'
  | 'home.special-preview'
  | 'home.special-select'
  | 'home.select-blanks'
  | 'home.select-formulas'
  | 'home.clear-special-selection'
  | 'home.undo'
  | 'home.redo';
export type HomeCellFormat = Pick<
  WorkbookEdit,
  | 'bold'
  | 'italic'
  | 'underline'
  | 'fontFamily'
  | 'fontSize'
  | 'fontColor'
  | 'fillColor'
  | 'horizontalAlign'
  | 'verticalAlign'
  | 'wrapText'
  | 'border'
  | 'styleId'
  | 'numberFormat'
>;
export type HomeMenu =
  | 'paste'
  | 'merge'
  | 'conditional'
  | 'table'
  | 'styles'
  | 'insert'
  | 'delete'
  | 'format'
  | 'fill'
  | 'clear'
  | 'find';
export interface HomeFillAcrossOptions {
  readonly targets: readonly string[];
  readonly mode: WorkbookFillAcrossSheetsMode;
}
export type HomeArgumentsById = { [I in NoArgumentHomeId]: undefined } & {
  'home.paste': undefined | { text: string };
  'home.paste-text': undefined | { text: string };
  'home.paint-format': { format: HomeCellFormat };
  'home.font-family': { value: string };
  'home.font-size': { value: number };
  'home.bold': { value: boolean };
  'home.italic': { value: boolean };
  'home.underline': { value: boolean };
  'home.borders': { value: 'none' | 'all' };
  'home.fill-color': { value: string };
  'home.font-color': { value: string };
  'home.wrap': { value: boolean };
  'home.number-format': { value: NonNullable<WorkbookEdit['numberFormat']> };
  'home.conditional-draft': { field: 'rule' | 'threshold' | 'fillColor'; value: string };
  'home.conditional-apply': {
    rule: 'greaterThan' | 'lessThan' | 'duplicateValues';
    threshold?: number;
    fillColor: string;
  };
  'home.table-draft': { name: string };
  'home.create-table': { name: string };
  'home.column-width': { value: number };
  'home.row-height': { value: number };
  'home.fill-range': { source: string; range: string };
  'home.fill-across-preview': HomeFillAcrossOptions;
  'home.fill-across-apply': HomeFillAcrossOptions;
  'home.dimension-draft': { dimension: 'column-width' | 'row-height'; value: string };
  'home.find-draft': { field: 'text' | 'replacement' | 'scope' | 'mode'; value: string };
  'home.find-page': { direction: 'previous' | 'next' };
  'home.find-match': { index: number };
  'home.menu': { value: HomeMenu };
};
export type HomeCommandId = keyof HomeArgumentsById;
export type HomeArguments = HomeArgumentsById[HomeCommandId];
export interface HomeCommandDefinition {
  readonly id: HomeCommandId;
  readonly group: string;
  readonly order: number;
  readonly label: string;
  readonly kind: 'workbook' | 'hostJob' | 'localView' | 'tool';
  readonly supportedProfiles: readonly RibbonProfile[];
  readonly requiresEdit: boolean;
  readonly requiresBatch: boolean;
  readonly allowSpecial: boolean;
}
const all: readonly RibbonProfile[] = ['xlsx', 'delimited', 'passive'],
  editable: readonly RibbonProfile[] = ['xlsx', 'delimited'],
  xlsx: readonly RibbonProfile[] = ['xlsx'];
const entry = (
  id: HomeCommandId,
  group: string,
  order: number,
  label: string,
  kind: HomeCommandDefinition['kind'] = 'workbook',
  profiles: readonly RibbonProfile[] = xlsx,
  requiresEdit = true,
  allowSpecial = false,
  requiresBatch = kind === 'workbook'
): HomeCommandDefinition =>
  Object.freeze({
    id,
    group,
    order,
    label,
    kind,
    supportedProfiles: Object.freeze([...profiles]),
    requiresEdit,
    requiresBatch,
    allowSpecial,
  });
/** IDs inventory real Home controls, never missing screenshot IDs. */
export const homeCommandDefinitions: readonly HomeCommandDefinition[] = Object.freeze([
  entry('home.paste', 'Clipboard', 1, 'Paste', 'hostJob', editable, true, false, true),
  entry('home.paste-values', 'Clipboard', 2, 'Paste values'),
  entry('home.paste-formulas', 'Clipboard', 3, 'Paste formulas'),
  entry('home.paste-formatting', 'Clipboard', 4, 'Paste formatting'),
  entry('home.paste-text', 'Clipboard', 4.5, 'Paste as text', 'hostJob', editable, true, false, true),
  entry('home.cut', 'Clipboard', 5, 'Cut', 'hostJob', editable),
  entry('home.copy', 'Clipboard', 6, 'Copy', 'hostJob', all, false),
  entry('home.format-painter', 'Clipboard', 7, 'Format Painter', 'localView', xlsx, true, false, true),
  entry('home.paint-format', 'Clipboard', 8, 'Apply copied formatting'),
  entry('home.cancel-paste', 'Clipboard', 9, 'Cancel paste', 'localView', editable, false, true, false),
  entry('home.font-family', 'Font', 20, 'Font family', 'workbook', xlsx, true, true),
  entry('home.font-size', 'Font', 21, 'Font size', 'workbook', xlsx, true, true),
  entry('home.font-size-increase', 'Font', 22, 'Increase font size', 'workbook', xlsx, true, true),
  entry('home.font-size-decrease', 'Font', 23, 'Decrease font size', 'workbook', xlsx, true, true),
  ...(['bold', 'italic', 'underline', 'borders', 'fill-color', 'font-color'] as const).map((key, index) =>
    entry(`home.${key}`, 'Font', 24 + index, key.replaceAll('-', ' '), 'workbook', xlsx, true, true)
  ),
  ...(['top', 'middle', 'bottom', 'left', 'center', 'right'] as const).map((value, index) =>
    entry(`home.align-${value}`, 'Alignment', 40 + index, `Align ${value}`, 'workbook', xlsx, true, true)
  ),
  entry('home.wrap', 'Alignment', 46, 'Wrap Text', 'workbook', xlsx, true, true),
  entry('home.merge-center', 'Alignment', 47, 'Merge & Center'),
  entry('home.unmerge', 'Alignment', 48, 'Unmerge cells'),
  entry('home.number-format', 'Number', 60, 'Number format', 'workbook', xlsx, true, true),
  ...(['accounting', 'percent', 'thousands', 'decimals-increase', 'decimals-decrease'] as const).map((key, index) =>
    entry(`home.${key}`, 'Number', 61 + index, key.replaceAll('-', ' '), 'workbook', xlsx, true, true)
  ),
  entry('styles.conditional', 'Styles', 80, 'Conditional formatting', 'tool', xlsx, false, true, false),
  entry('home.conditional-draft', 'Styles', 81, 'Conditional formatting draft', 'localView', xlsx, false, true, false),
  entry('home.conditional-apply', 'Styles', 82, 'Apply conditional rule'),
  entry('home.table-draft', 'Styles', 83, 'Table name draft', 'localView', xlsx, false, true, false),
  entry('home.create-table', 'Styles', 84, 'Create table'),
  ...(['good', 'bad', 'heading'] as const).map((style, index) =>
    entry(`home.style-${style}`, 'Styles', 85 + index, `Style ${style}`, 'workbook', xlsx, true, true)
  ),
  ...(['insert-row', 'insert-column', 'delete-row', 'delete-column', 'column-width', 'row-height'] as const).map(
    (key, index) => entry(`home.${key}`, 'Cells', 100 + index, key.replaceAll('-', ' '))
  ),
  entry('home.dimension-draft', 'Cells', 106, 'Dimension draft', 'localView', xlsx, false, true, false),
  ...(['fill-down', 'fill-right'] as const).map((key, index) =>
    entry(`home.${key}`, 'Editing', 120 + index, key.replaceAll('-', ' '))
  ),
  entry('home.fill-up', 'Editing', 121.25, 'Fill up'),
  entry('home.fill-left', 'Editing', 121.5, 'Fill left'),
  entry('home.fill-across', 'Editing', 121.75, 'Across Worksheets', 'tool', xlsx, true, false, false),
  entry(
    'home.fill-across-preview',
    'Editing',
    121.8,
    'Preview Across Worksheets',
    'localView',
    xlsx,
    true,
    false,
    false
  ),
  entry('home.fill-across-apply', 'Editing', 121.9, 'Apply Across Worksheets'),
  entry('home.fill-series', 'Editing', 122, 'Fill series', 'tool', xlsx, false, false, false),
  entry('home.fill-series-apply', 'Editing', 122.5, 'Apply series'),
  entry('home.fill-range', 'Editing', 119, 'Fill selected extension'),
  entry('home.clear-contents', 'Editing', 123, 'Clear contents', 'workbook', editable, true, true),
  entry('home.clear-formats', 'Editing', 124, 'Clear formats', 'workbook', xlsx, true, true),
  entry('home.clear-all', 'Editing', 125, 'Clear all', 'workbook', xlsx, true, true),
  entry('home.clear-hyperlinks', 'Editing', 125.5, 'Clear Hyperlinks', 'workbook', xlsx, true, true),
  ...(
    [
      'find-open',
      'find-draft',
      'find-search',
      'find-cancel',
      'find-next',
      'find-page',
      'find-close',
      'select-blanks',
      'select-formulas',
      'clear-special-selection',
      'menu',
    ] as const
  ).map((key, index) =>
    entry(`home.${key}`, 'Editing', 130 + index, key.replaceAll('-', ' '), 'localView', all, false, true, false)
  ),
  entry('home.replace-preview', 'Editing', 136, 'Preview replacements', 'tool', editable, false, false, false),
  entry('home.find-match', 'Editing', 142, 'Go to found match', 'localView', all, false, true, false),
  entry('home.special-open', 'Editing', 143, 'Go To Special', 'tool', editable, false, false, false),
  entry(
    'home.special-preview',
    'Editing',
    143.25,
    'Preview matching cells',
    'localView',
    editable,
    false,
    false,
    false
  ),
  entry('home.special-select', 'Editing', 143.5, 'Select matching cells', 'localView', editable, false, false, false),
  entry('home.replace-one', 'Editing', 150, 'Replace', 'workbook', editable, true, false),
  entry('home.replace-all', 'Editing', 151, 'Apply replacements', 'workbook', editable, true, false),
  entry('home.undo', 'Editing', 160, 'Undo', 'hostJob', editable, true, true, false),
  entry('home.redo', 'Editing', 161, 'Redo', 'hostJob', editable, true, true, false),
]);
const definitions = new Map(homeCommandDefinitions.map((item) => [item.id, item]));
export const isHomeCommandId = (id: string): id is HomeCommandId => definitions.has(id as HomeCommandId);
export function homeCommandDefinition(id: HomeCommandId): HomeCommandDefinition {
  const found = definitions.get(id);
  if (!found) throw Error('Unknown Home command');
  return found;
}
const plain = (v: unknown): v is Record<string, unknown> =>
  !!v &&
  typeof v === 'object' &&
  !Array.isArray(v) &&
  (Object.getPrototypeOf(v) === Object.prototype || Object.getPrototypeOf(v) === null);
const exact = (v: unknown, required: string[], optional: string[] = []): v is Record<string, unknown> =>
  plain(v) &&
  required.every((key) => Object.hasOwn(v, key)) &&
  Object.keys(v).every((key) => required.includes(key) || optional.includes(key));
const choice = (v: unknown, values: readonly unknown[]) => values.includes(v);
const text = (v: unknown, max = 32767, empty = true): v is string =>
  typeof v === 'string' &&
  v.length <= max &&
  (empty || !!v.trim()) &&
  !/[\u0000-\u0008\u000b\u000c\u000e-\u001f\uD800-\uDFFF\uFFFE\uFFFF]/u.test(v);
const finite = (v: unknown, min: number, max: number): v is number =>
  typeof v === 'number' && Number.isFinite(v) && v >= min && v <= max;
const colors = (v: unknown): v is string => typeof v === 'string' && /^[0-9a-f]{6}$/i.test(v);
const formats = [
  'general',
  'number',
  'percent',
  'date',
  'text',
  'currency',
  'accounting',
  'thousands',
  'decimal0',
  'decimal1',
  'decimal2',
] as const;
const rules = ['greaterThan', 'lessThan', 'duplicateValues'] as const;
const menus: readonly HomeMenu[] = [
  'paste',
  'merge',
  'conditional',
  'table',
  'styles',
  'insert',
  'delete',
  'format',
  'fill',
  'clear',
  'find',
];
function boundedRectangle(value: unknown, limit: number): value is string {
  if (!range(value)) return false;
  const [first, last = first] = value.split(':'),
    a = /^([A-Z]+)(\d+)$/.exec(first)!,
    b = /^([A-Z]+)(\d+)$/.exec(last)!;
  const column = (text: string) => [...text].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0),
    rows = Number(b[2]) - Number(a[2]) + 1,
    columns = column(b[1]) - column(a[1]) + 1;
  return rows > 0 && columns > 0 && rows * columns <= limit;
}
function format(value: unknown): HomeCellFormat {
  const keys = [
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
    'styleId',
    'numberFormat',
  ];
  if (!plain(value) || !Object.keys(value).length || Object.keys(value).some((key) => !keys.includes(key)))
    throw Error('Invalid Home formatting arguments');
  for (const [key, v] of Object.entries(value)) {
    const valid = ['bold', 'italic', 'underline', 'wrapText'].includes(key)
      ? typeof v === 'boolean'
      : key === 'fontFamily'
        ? text(v, 100, false)
        : key === 'fontSize'
          ? finite(v, 6, 72)
          : ['fontColor', 'fillColor'].includes(key)
            ? colors(v)
            : key === 'horizontalAlign'
              ? choice(v, ['left', 'center', 'right'])
              : key === 'verticalAlign'
                ? choice(v, ['top', 'center', 'bottom'])
                : key === 'border'
                  ? choice(v, ['all', 'none'])
                  : key === 'styleId'
                    ? Number.isSafeInteger(v) && Number(v) >= 0
                    : choice(v, formats);
    if (!valid) throw Error('Invalid Home formatting arguments');
  }
  return Object.freeze({ ...value }) as HomeCellFormat;
}
/** Copies closed UI arguments. Draft strings may be incomplete; apply arguments
 * must be typed. This does not replace native style/table/dependency admission. */
export function normalizeHomeArguments<I extends HomeCommandId>(id: I, args?: unknown): HomeArgumentsById[I] {
  homeCommandDefinition(id);
  let result: unknown;
  if (id === 'home.paste' || id === 'home.paste-text') {
    if (args === undefined) return undefined as HomeArgumentsById[I];
    if (exact(args, ['text']) && typeof args.text === 'string' && args.text.length <= 64 * 1024 * 1024)
      result = { text: args.text };
  } else if (id === 'home.paint-format') {
    if (exact(args, ['format'])) result = { format: format(args.format) };
  } else if (['home.bold', 'home.italic', 'home.underline', 'home.wrap'].includes(id)) {
    if (exact(args, ['value']) && typeof args.value === 'boolean') result = { value: args.value };
  } else if (id === 'home.font-family') {
    if (exact(args, ['value']) && text(args.value, 100, false)) result = { value: args.value };
  } else if (id === 'home.font-size' || id === 'home.column-width' || id === 'home.row-height') {
    if (
      exact(args, ['value']) &&
      finite(
        args.value,
        id === 'home.font-size' ? 6 : 0.01,
        id === 'home.font-size' ? 72 : id === 'home.column-width' ? 255 : 409.5
      )
    )
      result = { value: args.value };
  } else if (id === 'home.borders') {
    if (exact(args, ['value']) && choice(args.value, ['none', 'all'])) result = { value: args.value };
  } else if (id === 'home.fill-color' || id === 'home.font-color') {
    if (exact(args, ['value']) && colors(args.value)) result = { value: args.value };
  } else if (id === 'home.number-format') {
    if (exact(args, ['value']) && choice(args.value, formats)) result = { value: args.value };
  } else if (id === 'home.conditional-draft') {
    if (
      exact(args, ['field', 'value']) &&
      text(args.value) &&
      choice(args.field, ['rule', 'threshold', 'fillColor']) &&
      (args.field !== 'rule' || choice(args.value, rules))
    )
      result = { field: args.field, value: args.value };
  } else if (id === 'home.conditional-apply') {
    if (
      exact(args, ['rule', 'fillColor'], ['threshold']) &&
      choice(args.rule, rules) &&
      colors(args.fillColor) &&
      (args.rule === 'duplicateValues'
        ? !Object.hasOwn(args, 'threshold')
        : typeof args.threshold === 'number' && Number.isFinite(args.threshold))
    )
      result = {
        rule: args.rule,
        fillColor: args.fillColor,
        ...(Object.hasOwn(args, 'threshold') ? { threshold: args.threshold } : {}),
      };
  } else if (id === 'home.table-draft' || id === 'home.create-table') {
    if (exact(args, ['name']) && text(args.name, 255, id === 'home.table-draft')) result = { name: args.name };
  } else if (id === 'home.find-draft') {
    if (
      exact(args, ['field', 'value']) &&
      choice(args.field, ['text', 'replacement', 'scope', 'mode']) &&
      text(args.value) &&
      (args.field !== 'scope' || choice(args.value, ['sheet', 'workbook'])) &&
      (args.field !== 'mode' || choice(args.value, ['display', 'raw', 'formula']))
    )
      result = { field: args.field, value: args.value };
  } else if (id === 'home.find-page') {
    if (exact(args, ['direction']) && choice(args.direction, ['previous', 'next']))
      result = { direction: args.direction };
  } else if (id === 'home.fill-across-preview' || id === 'home.fill-across-apply') {
    if (
      exact(args, ['targets', 'mode']) &&
      Array.isArray(args.targets) &&
      args.targets.length > 0 &&
      args.targets.length <= 10000 &&
      args.targets.every((value) => text(value, 31, false)) &&
      new Set(args.targets).size === args.targets.length &&
      choice(args.mode, ['all', 'contents', 'formats'])
    )
      result = { targets: Object.freeze([...args.targets]), mode: args.mode };
  } else if (id === 'home.fill-range') {
    if (exact(args, ['source', 'range']) && boundedRectangle(args.source, 1000) && boundedRectangle(args.range, 10000))
      result = { source: args.source, range: args.range };
  } else if (id === 'home.find-match') {
    if (exact(args, ['index']) && Number.isInteger(args.index) && Number(args.index) >= 0 && Number(args.index) <= 999)
      result = { index: args.index };
  } else if (id === 'home.dimension-draft') {
    if (
      exact(args, ['dimension', 'value']) &&
      choice(args.dimension, ['column-width', 'row-height']) &&
      text(args.value)
    )
      result = { dimension: args.dimension, value: args.value };
  } else if (id === 'home.menu') {
    if (exact(args, ['value']) && choice(args.value, menus)) result = { value: args.value };
  } else if (args === undefined) return undefined as HomeArgumentsById[I];
  if (result === undefined) throw Error('Invalid Home command arguments');
  return Object.freeze(result) as HomeArgumentsById[I];
}

export interface HomeBinding extends RibbonBinding {
  readonly ownerIdentity: string | null;
  readonly selectedCell: string;
  readonly navigationGeneration: number;
  readonly modeGeneration: number;
  readonly specialSelectionEpoch: number;
  readonly mode: 'normal' | 'pageLayout' | 'pageBreak';
  readonly specialSelection: readonly string[] | null;
}
export interface HomeFrame {
  readonly binding: HomeBinding;
  readonly profile: RibbonProfile;
  readonly active: boolean;
  readonly canEdit: boolean;
  readonly busy: boolean;
  readonly acceptBatch?: (commands: WorkbookCommand[]) => Promise<void>;
  /** Only a literal portable owner whose strict callback captures document,
   * generation and head. It does not manufacture a native session identity. */
  readonly portableOwnerBound?: boolean;
}
export interface HomeDispatchScope {
  captureFrame: () => HomeFrame;
  currentFrame: () => HomeFrame | null;
}
export class HomeCommandRetired extends Error {
  readonly code = 'HOME_COMMAND_RETIRED';
  constructor() {
    super('The initiating workbook, revision or target changed.');
    this.name = 'HomeCommandRetired';
  }
}
const keys = [
  'documentId',
  'state',
  'sheet',
  'selection',
  'activationEpoch',
  'ownerIdentity',
  'selectedCell',
  'navigationGeneration',
  'modeGeneration',
  'specialSelectionEpoch',
  'mode',
  'specialSelection',
] as const;
function cell(value: unknown): value is string {
  if (typeof value !== 'string') return false;
  const match = /^([A-Z]{1,3})([1-9]\d*)$/.exec(value);
  return (
    !!match && Number(match[2]) <= 1048576 && [...match[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) <= 16384
  );
}
function range(value: unknown): value is string {
  if (typeof value !== 'string') return false;
  const parts = value.split(':');
  return parts.length <= 2 && parts.every(cell);
}
export function captureHomeBinding(value: HomeBinding): HomeBinding {
  if (
    !exact(value, [...keys]) ||
    !['documentId', 'state', 'sheet'].every((key) => text(value[key], 4096, false)) ||
    !range(value.selection) ||
    !cell(value.selectedCell) ||
    (value.ownerIdentity !== null && !text(value.ownerIdentity, 4096, false)) ||
    !choice(value.mode, ['normal', 'pageLayout']) ||
    !['activationEpoch', 'navigationGeneration', 'modeGeneration', 'specialSelectionEpoch'].every(
      (key) => Number.isSafeInteger(value[key]) && Number(value[key]) >= 0
    )
  )
    throw Error('Home command requires an accepted owner binding');
  const selected = value.specialSelection;
  if (
    selected !== null &&
    (!Array.isArray(selected) ||
      !selected.length ||
      selected.length > 10000 ||
      !selected.every(cell) ||
      new Set(selected).size !== selected.length)
  )
    throw Error('Invalid Home special selection');
  return Object.freeze({ ...value, specialSelection: selected === null ? null : Object.freeze([...selected].sort()) });
}
export function homeBindingMatches(a: HomeBinding, b: HomeBinding | null, runtimeOnly = false): boolean {
  if (!b) return false;
  const aSpecial = a.specialSelection === null ? null : [...a.specialSelection].sort(),
    bSpecial = b.specialSelection === null ? null : [...b.specialSelection].sort();
  return keys.every((key) => {
    if (runtimeOnly && key === 'state') return true;
    if (key === 'specialSelection')
      return (
        (aSpecial === null && bSpecial === null) ||
        (!!aSpecial &&
          !!bSpecial &&
          aSpecial.length === bSpecial.length &&
          aSpecial.every((value, index) => value === bSpecial[index]))
      );
    return a[key] === b[key];
  });
}
export function homeCommandCapability(id: HomeCommandId, frame: HomeFrame, hasHandler = true): RibbonCapability {
  const command = homeCommandDefinition(id);
  if (!hasHandler) return { state: 'hidden', reason: 'This host has not supplied this Home command.' };
  if (!command.supportedProfiles.includes(frame.profile))
    return { state: 'hidden', reason: 'This workbook format does not support this command.' };
  if (!frame.active) return { state: 'disabled', reason: 'Activate this workbook first.' };
  if (command.requiresEdit && frame.binding.mode !== 'normal')
    return { state: 'disabled', reason: 'Return to Normal view to edit the worksheet.' };
  if (command.requiresEdit && !frame.canEdit) return { state: 'disabled', reason: 'This workbook is read-only.' };
  if (command.requiresEdit && frame.busy)
    return { state: 'disabled', reason: 'Wait for the workbook operation to finish.' };
  if (
    command.requiresEdit &&
    !frame.binding.ownerIdentity &&
    !(frame.profile === 'delimited' && frame.portableOwnerBound === true && frame.acceptBatch)
  )
    return { state: 'disabled', reason: 'This editor has not bound an accepted workbook owner.' };
  if (command.requiresBatch && !frame.acceptBatch)
    return { state: 'disabled', reason: 'Workbook commands are unavailable in this editor.' };
  if (frame.binding.specialSelection && !command.allowSpecial)
    return {
      state: 'disabled',
      reason: 'Special selection supports formatting and Clear. Click a cell for other actions.',
    };
  return { state: 'enabled' };
}
export interface HomeCommandLease {
  readonly binding: HomeBinding;
  isCurrent: () => boolean;
  assertCurrent: () => void;
  read: <T>(reader: () => Promise<T>) => Promise<T>;
  accept: (commands: WorkbookCommand[]) => Promise<void>;
  perform: <T>(job: () => Promise<T>) => Promise<T>;
}
export type HomeCommandHandler<I extends HomeCommandId, T = unknown> = (
  lease: HomeCommandLease,
  args: HomeArgumentsById[I]
) => T | Promise<T>;
export type HomeDispatchResult<T = unknown> =
  | { outcome: 'refused'; binding: HomeBinding; reason: string; current: boolean; runtimeCurrent: boolean }
  | { outcome: 'stale'; binding: HomeBinding; current: false; runtimeCurrent: boolean }
  | { outcome: 'completed'; binding: HomeBinding; value: T; current: boolean; runtimeCurrent: boolean }
  | { outcome: 'failed'; binding: HomeBinding; error: unknown; current: boolean; runtimeCurrent: boolean };
function acknowledged<T>(result: Promise<T>): Promise<T> {
  if (!result || typeof result.then !== 'function')
    throw Error('This command did not supply an awaited acknowledgement.');
  return result;
}
export function normalizeHomeBatch(commands: WorkbookCommand[]): WorkbookCommand[] {
  if (!Array.isArray(commands) || !commands.length) throw Error('Home command has no workbook changes');
  const copied = structuredClone(commands);
  return copied.flatMap((command) =>
    'type' in command && command.type === 'setCells' && command.cells.length > 1000
      ? Array.from({ length: Math.ceil(command.cells.length / 1000) }, (_, index) => ({
          ...command,
          cells: command.cells.slice(index * 1000, (index + 1) * 1000),
        }))
      : [command]
  );
}
/** captureFrame must be bound to the initiating render, currentFrame to live
 * state. Readers/job closures likewise come from that captured render. */
export async function dispatchHomeCommand<I extends HomeCommandId, T>(
  id: I,
  scope: HomeDispatchScope,
  handler?: HomeCommandHandler<I, T>,
  args?: HomeArgumentsById[I]
): Promise<HomeDispatchResult<T>> {
  const original = scope.captureFrame(),
    binding = captureHomeBinding(original.binding),
    frame = Object.freeze({ ...original, binding }),
    command = homeCommandDefinition(id);
  const matches = (runtimeOnly = false): boolean => {
    const current = scope.currentFrame();
    return (
      !!current &&
      current.active &&
      current.profile === frame.profile &&
      homeBindingMatches(binding, current.binding, runtimeOnly) &&
      (binding.ownerIdentity !== null ||
        frame.profile !== 'delimited' ||
        (current.portableOwnerBound === frame.portableOwnerBound && current.acceptBatch === frame.acceptBatch))
    );
  };
  const state = () => ({ current: matches(), runtimeCurrent: matches(true) });
  if (!matches()) return { outcome: 'stale', binding, current: false, runtimeCurrent: matches(true) };
  const capability = homeCommandCapability(id, frame, !!handler);
  if (capability.state !== 'enabled')
    return {
      outcome: 'refused',
      binding,
      reason: capability.reason ?? 'This Home command is unavailable.',
      ...state(),
    };
  const assertCurrent = (): void => {
    if (!matches()) throw new HomeCommandRetired();
    const current = scope.currentFrame()!;
    const available = homeCommandCapability(id, { ...current, acceptBatch: frame.acceptBatch });
    if (available.state !== 'enabled') throw Error(available.reason ?? 'This Home command is unavailable.');
  };
  const perform = async <U>(job: () => Promise<U>): Promise<U> => {
    assertCurrent();
    return await acknowledged(job());
  };
  const lease: HomeCommandLease = Object.freeze({
    binding,
    isCurrent: matches,
    assertCurrent,
    read: async <U>(reader: () => Promise<U>) => {
      assertCurrent();
      const result = await acknowledged(reader());
      assertCurrent();
      return result;
    },
    perform,
    accept: async (commands: WorkbookCommand[]) => {
      if (!command.requiresEdit || !frame.acceptBatch)
        throw Error('Workbook commands are unavailable for this action.');
      assertCurrent();
      await acknowledged(frame.acceptBatch(normalizeHomeBatch(commands)));
    },
  });
  try {
    const normalized = normalizeHomeArguments(id, args);
    const value = await handler!(lease, normalized);
    return { outcome: 'completed', binding, value, ...state() };
  } catch (error) {
    return { outcome: 'failed', binding, error, ...state() };
  }
}
