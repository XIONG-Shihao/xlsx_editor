/** Shared command metadata and lifetime-aware dispatch. This module never owns
 * workbook mutations, file grants, publication or arbitrary script execution. */
import { inkPen, type InkPenId } from '../core/ink-stroke.mjs';
import type { FunctionCategory } from './functionCatalog';
import {
  homeCommandDefinitions,
  normalizeHomeArguments,
  type HomeCommandId,
  type HomeArguments,
} from './homeCommands.ts';
export type WorkbookTool =
  | 'accessibility'
  | 'notes'
  | 'illustrations'
  | 'pageStyles'
  | 'checkbox'
  | 'symbols'
  | 'hyperlink'
  | 'pivot'
  | 'pivotDesign'
  | 'names'
  | 'objects'
  | 'print'
  | 'conditional'
  | 'validation'
  | 'references'
  | 'queries'
  | 'functions'
  | 'watch'
  | 'errors'
  | 'textColumns'
  | 'trace'
  | 'calculation'
  | 'createNames'
  | 'sort'
  | 'duplicates'
  | 'review'
  | 'actions'
  | 'fill'
  | 'replace'
  | 'special';
export type RibbonTab =
  | 'Home'
  | 'Insert'
  | 'Draw'
  | 'Page Layout'
  | 'Formulas'
  | 'Data'
  | 'Review'
  | 'View'
  | 'Automate'
  | 'PivotTable Analyze'
  | 'Design';
export type RibbonProfile = 'xlsx' | 'delimited' | 'passive';
export type RibbonCommandKind = 'tool' | 'localView' | 'workbook' | 'hostJob';
export type RibbonCommandId =
  | 'draw.pen'
  | 'draw.select-cells'
  | 'draw.select-strokes'
  | 'draw.eraser'
  | 'draw.preset'
  | 'insert.recommended-charts'
  | 'insert.model3d'
  | 'insert.smartart'
  | 'insert.wordart'
  | 'insert.icons'
  | 'insert.shapes'
  | 'insert.text-box'
  | 'insert.checkbox'
  | 'insert.symbols'
  | 'insert.link'
  | 'pivots.context'
  | 'pivots.design'
  | 'pivots.refresh'
  | 'pivots.go-source'
  | HomeCommandId
  | 'formulas.error-checking'
  | 'formulas.watch'
  | 'formulas.insert-function'
  | 'formulas.recently-used'
  | 'names.manager'
  | 'names.define'
  | 'names.use-in-formula'
  | 'names.create-from-selection'
  | 'objects.manage'
  | 'references.manage'
  | 'formulas.financial'
  | 'formulas.logical'
  | 'formulas.text'
  | 'formulas.date-time'
  | 'formulas.lookup-reference'
  | 'formulas.math-trig'
  | 'formulas.more-functions'
  | 'formulas.trace-precedents'
  | 'formulas.trace-dependents'
  | 'formulas.remove-arrows'
  | 'formulas.calculate-now'
  | 'formulas.calculate-sheet'
  | 'formulas.calculation-options'
  | 'page.themes'
  | 'page.background'
  | 'print.setup'
  | 'print.margins'
  | 'print.orientation'
  | 'print.paper'
  | 'print.area'
  | 'print.titles'
  | 'print.scale'
  | 'data.refresh-all'
  | 'data.remove-duplicates'
  | 'data.text-to-columns'
  | 'data.import-file'
  | 'data.queries'
  | 'data.sort'
  | 'data.filter'
  | 'data.validation'
  | 'styles.conditional'
  | 'formulas.auto-sum'
  | 'review.accessibility'
  | 'insert.note'
  | 'review.notes'
  | 'review.statistics'
  | 'review.spelling'
  | 'actions.new'
  | 'actions.library'
  | 'actions.unhide-all'
  | 'actions.freeze-selection'
  | 'view.preset-apply'
  | 'view.preset-save'
  | 'view.preset-update'
  | 'view.preset-delete'
  | 'view.source-defaults'
  | 'view.page-break-preview'
  | 'view.page-layout'
  | 'view.normal'
  | 'view.zoom'
  | 'view.zoom-100'
  | 'view.zoom-selection'
  | 'view.gridlines'
  | 'view.headings'
  | 'view.formula-bar'
  | 'view.show-formulas'
  | 'view.freeze';
export interface RibbonCommandDefinition {
  readonly id: RibbonCommandId;
  readonly matrixIds: readonly string[];
  readonly tab: RibbonTab;
  readonly group: string;
  readonly order: number;
  readonly label: string;
  readonly hint: string;
  readonly kind: RibbonCommandKind;
  readonly supportedProfiles: readonly RibbonProfile[];
  readonly requiresEdit: boolean;
  readonly argument: 'none' | 'zoom' | 'toggle' | 'freeze' | 'home' | 'inkPreset';
}
const all: readonly RibbonProfile[] = ['xlsx', 'delimited', 'passive'];
const workbook: readonly RibbonProfile[] = ['xlsx'];
const definition = (
  id: RibbonCommandId,
  matrixIds: string[],
  tab: RibbonTab,
  group: string,
  order: number,
  label: string,
  kind: RibbonCommandKind,
  supportedProfiles: readonly RibbonProfile[],
  requiresEdit = false,
  argument: RibbonCommandDefinition['argument'] = 'none',
  hint = label
): RibbonCommandDefinition =>
  Object.freeze({
    id,
    matrixIds: Object.freeze(matrixIds),
    tab,
    group,
    order,
    label,
    hint,
    kind,
    supportedProfiles: Object.freeze([...supportedProfiles]),
    requiresEdit,
    argument,
  });

/** Potentially supported commands, not a declaration that every host implements
 * every handler. Planned/excluded screenshot inventory stays in FEATURE_MATRIX. */
export const ribbonCommands: readonly RibbonCommandDefinition[] = Object.freeze([
  definition(
    'draw.preset',
    ['XL-D04'],
    'Draw',
    'Pens',
    30,
    'Pen presets',
    'localView',
    workbook,
    true,
    'inkPreset',
    'Choose one of three fixed pen/highlighter samples, remembered privately in this app profile.'
  ),
  definition(
    'draw.pen',
    ['XL-D01'],
    'Draw',
    'Drawing tools',
    10,
    'Draw',
    'tool',
    workbook,
    true,
    'none',
    'Draw an anchored stroke using the selected pen. Release to save one Undo step; Escape cancels the draft.'
  ),
  definition(
    'draw.select-strokes',
    ['XL-D03'],
    'Draw',
    'Drawing tools',
    15,
    'Select Strokes',
    'tool',
    workbook,
    true,
    'none',
    'Select an owned line, drag to move it, or press Delete to remove the whole stroke.'
  ),
  definition(
    'draw.eraser',
    ['XL-D02'],
    'Draw',
    'Drawing tools',
    16,
    'Eraser',
    'tool',
    workbook,
    true,
    'none',
    'Click an owned line to erase that whole stroke on release. Escape cancels.'
  ),
  definition(
    'draw.select-cells',
    [],
    'Draw',
    'Drawing tools',
    20,
    'Select Cells',
    'localView',
    workbook,
    false,
    'none',
    'Leave drawing mode and return to ordinary cell editing.'
  ),
  definition(
    'actions.unhide-all',
    ['XL-A03'],
    'Automate',
    'Built-in actions',
    30,
    'Unhide Rows and Columns',
    'tool',
    workbook,
    true,
    'none',
    'Preview and explicitly unhide the captured worksheet in one Undo step.'
  ),
  definition(
    'actions.freeze-selection',
    ['XL-A04'],
    'Automate',
    'Built-in actions',
    40,
    'Freeze Selection',
    'tool',
    workbook,
    true,
    'none',
    'Preview and freeze rows/columns preceding the captured first selected cell.'
  ),
  definition(
    'actions.new',
    ['XL-A01'],
    'Automate',
    'Local actions',
    10,
    'New Action',
    'tool',
    all,
    false,
    'none',
    'Define a local action for your private library. Saving it leaves the workbook unchanged.'
  ),
  definition(
    'actions.library',
    ['XL-A02'],
    'Automate',
    'Local actions',
    20,
    'Local Actions',
    'tool',
    all,
    false,
    'none',
    'Manage saved action definitions in this application profile.'
  ),
  definition(
    'review.accessibility',
    ['XL-R03'],
    'Review',
    'Accessibility',
    25,
    'Check Accessibility',
    'tool',
    workbook,
    false,
    'none',
    'Inspect supported workbook content without editing it.'
  ),
  definition(
    'review.notes',
    ['XL-R07'],
    'Review',
    'Notes',
    30,
    'Notes',
    'tool',
    workbook,
    false,
    'none',
    'View, add, edit or delete plain workbook notes.'
  ),
  definition(
    'insert.note',
    ['XL-I14'],
    'Insert',
    'Comments',
    32,
    'New Note',
    'tool',
    workbook,
    false,
    'none',
    'Add or edit a plain note at the selected cell.'
  ),
  definition(
    'review.statistics',
    ['XL-R02'],
    'Review',
    'Workbook',
    20,
    'Workbook Statistics',
    'tool',
    workbook,
    false,
    'none',
    'Inspect accepted records across all worksheets, including hidden rows and sheets.'
  ),
  definition(
    'review.spelling',
    ['XL-R01'],
    'Review',
    'Proofing',
    10,
    'Spelling',
    'tool',
    workbook,
    false,
    'none',
    'Review literal text using an installed offline dictionary. Confirm each correction separately.'
  ),
  definition(
    'formulas.calculation-options',
    ['XL-F13'],
    'Formulas',
    'Calculation',
    150,
    'Calculation Options',
    'tool',
    workbook,
    false,
    'none',
    'Inspect or change Automatic, Manual and recalculation before saving.'
  ),
  definition(
    'formulas.calculate-sheet',
    ['XL-F14'],
    'Formulas',
    'Calculation',
    170,
    'Calculate Sheet',
    'workbook',
    workbook,
    true,
    'none',
    'Calculate the active worksheet and required inputs. This creates one Undo step.'
  ),
  definition(
    'formulas.calculate-now',
    ['XL-F14'],
    'Formulas',
    'Calculation',
    160,
    'Calculate Now',
    'workbook',
    workbook,
    true,
    'none',
    'Recalculate all worksheets in the current workbook. This creates one Undo step.'
  ),
  definition(
    'formulas.trace-precedents',
    ['XL-F09'],
    'Formulas',
    'Formula auditing',
    121,
    'Trace Precedents',
    'tool',
    workbook,
    false,
    'none',
    'Show known references feeding the selected cell; repeat to show another level.'
  ),
  definition(
    'formulas.trace-dependents',
    ['XL-F09'],
    'Formulas',
    'Formula auditing',
    122,
    'Trace Dependents',
    'tool',
    workbook,
    false,
    'none',
    'Show formulas referencing the selected cell; repeat to show another level.'
  ),
  definition(
    'formulas.remove-arrows',
    ['XL-F09'],
    'Formulas',
    'Formula auditing',
    123,
    'Remove Arrows',
    'localView',
    workbook,
    false,
    'none',
    'Clear the current formula trace.'
  ),
  definition(
    'names.manager',
    ['XL-F06'],
    'Formulas',
    'Defined names',
    110,
    'Name Manager',
    'tool',
    ['xlsx', 'passive'],
    false,
    'none',
    'Inspect and manage workbook names.'
  ),
  definition(
    'names.define',
    ['XL-F06'],
    'Formulas',
    'Defined names',
    120,
    'Define Name',
    'tool',
    workbook,
    true,
    'none',
    'Define a name for the selected range.'
  ),
  definition(
    'names.use-in-formula',
    ['XL-F07'],
    'Formulas',
    'Defined names',
    125,
    'Use in Formula',
    'localView',
    workbook,
    true,
    'none',
    'Insert a name in the formula draft. Press Enter to apply it.'
  ),
  definition(
    'names.create-from-selection',
    ['XL-F08'],
    'Formulas',
    'Defined names',
    126,
    'Create from Selection',
    'tool',
    workbook,
    true,
    'none',
    'Create workbook names from selected labels in one Undo step.'
  ),
  definition(
    'insert.checkbox',
    ['XL-I06'],
    'Insert',
    'Controls',
    20,
    'Checkbox',
    'tool',
    workbook,
    false,
    'none',
    'Insert or edit the editor’s ordinary TRUE/FALSE checkbox profile.'
  ),
  definition(
    'insert.symbols',
    ['XL-I16'],
    'Insert',
    'Symbols',
    40,
    'Symbols',
    'tool',
    workbook,
    false,
    'none',
    'Append Omega, Plus/minus or Degree to a literal text cell.'
  ),
  definition(
    'insert.link',
    ['XL-I13'],
    'Insert',
    'Links',
    30,
    'Link',
    'tool',
    workbook,
    false,
    'none',
    'Insert, edit or remove an ordinary web or worksheet-cell hyperlink.'
  ),
  definition(
    'insert.icons',
    ['XL-I05'],
    'Insert',
    'Illustrations',
    16,
    'Icons',
    'tool',
    workbook,
    false,
    'none',
    'Insert or recolor Check, Star and Plus vector samples.'
  ),
  definition(
    'insert.shapes',
    ['XL-I05'],
    'Insert',
    'Illustrations',
    15,
    'Shapes',
    'tool',
    workbook,
    false,
    'none',
    'Insert or edit Rectangle, Ellipse and Right arrow samples.'
  ),
  definition(
    'insert.model3d',
    ['XL-I05'],
    'Insert',
    'Illustrations',
    18,
    '3D Models',
    'tool',
    workbook,
    false,
    'none',
    'Insert or edit Cube, Pyramid and Octahedron solid models.'
  ),
  definition(
    'insert.smartart',
    ['XL-I05'],
    'Insert',
    'Illustrations',
    17,
    'SmartArt',
    'tool',
    workbook,
    false,
    'none',
    'Insert or edit Process, Hierarchy and Cycle node/layout samples.'
  ),
  definition(
    'insert.wordart',
    ['XL-I15'],
    'Insert',
    'Text',
    36,
    'WordArt',
    'tool',
    workbook,
    false,
    'none',
    'Insert or edit Blue Fill, Orange Outline and Purple Shadow text samples.'
  ),
  definition(
    'insert.text-box',
    ['XL-I15'],
    'Insert',
    'Text',
    35,
    'Text Box',
    'tool',
    workbook,
    false,
    'none',
    'Insert or edit Plain and Note text-box samples.'
  ),
  definition(
    'insert.recommended-charts',
    ['XL-I07'],
    'Insert',
    'Charts',
    20,
    'Recommended Charts',
    'tool',
    workbook,
    false,
    'none',
    'Preview three charts from the selected supported data.'
  ),
  definition(
    'objects.manage',
    ['XL-I05', 'XL-I08'],
    'Insert',
    'Illustrations',
    10,
    'Charts and images',
    'tool',
    ['xlsx', 'passive'],
    false,
    'none',
    'Inspect or edit qualified charts and images.'
  ),
  definition(
    'pivots.context',
    [],
    'PivotTable Analyze',
    'PivotTable',
    10,
    'Field List',
    'tool',
    workbook,
    false,
    'none',
    'Inspect the selected pivot source and fields.'
  ),
  definition(
    'pivots.design',
    [],
    'Design',
    'PivotTable Styles',
    10,
    'PivotTable Styles',
    'tool',
    workbook,
    false,
    'none',
    'Choose one of three owned pivot layout/style samples and grand totals.'
  ),
  definition(
    'pivots.refresh',
    [],
    'PivotTable Analyze',
    'Data',
    20,
    'Refresh',
    'workbook',
    workbook,
    true,
    'none',
    'Refresh the selected owned pivot in one Undo step.'
  ),
  definition(
    'pivots.go-source',
    [],
    'PivotTable Analyze',
    'Data',
    30,
    'Go to Source',
    'localView',
    workbook,
    false,
    'none',
    'Select the source range of this pivot.'
  ),
  definition(
    'references.manage',
    ['XL-I01', 'XL-I02'],
    'Insert',
    'Tables',
    20,
    'Tables and pivots',
    'tool',
    workbook,
    false,
    'none',
    'Create and manage qualified tables and pivots.'
  ),
  definition(
    'page.themes',
    ['XL-P01'],
    'Page Layout',
    'Themes',
    10,
    'Themes',
    'tool',
    workbook,
    false,
    'none',
    'Preview and apply three workbook color/font samples.'
  ),
  definition(
    'page.background',
    ['XL-P06'],
    'Page Layout',
    'Page setup',
    45,
    'Background',
    'tool',
    workbook,
    false,
    'none',
    'Preview, apply or remove a repeating worksheet background; backgrounds do not print.'
  ),
  definition(
    'print.setup',
    ['XL-P08'],
    'Page Layout',
    'Page setup',
    65,
    'Page Setup',
    'tool',
    workbook,
    false,
    'none',
    'Set up pages, preview, print and export PDF.'
  ),
  definition('print.margins', ['XL-P02'], 'Page Layout', 'Page setup', 20, 'Margins', 'tool', workbook),
  definition('print.orientation', ['XL-P03'], 'Page Layout', 'Page setup', 30, 'Orientation', 'tool', workbook),
  definition('print.paper', ['XL-P03'], 'Page Layout', 'Page setup', 40, 'Size', 'tool', workbook),
  definition('print.area', ['XL-P04'], 'Page Layout', 'Page setup', 50, 'Print Area', 'tool', workbook),
  definition('print.titles', ['XL-P07'], 'Page Layout', 'Page setup', 60, 'Print Titles', 'tool', workbook),
  definition('print.scale', ['XL-P09'], 'Page Layout', 'Scale to fit', 70, 'Scale to Fit', 'tool', workbook),
  definition('data.import-file', ['XL-DA01'], 'Data', 'Get data', 10, 'Import CSV/TSV…', 'hostJob', workbook, true),
  definition(
    'data.refresh-all',
    ['XL-DA02'],
    'Data',
    'Queries and connections',
    15,
    'Refresh All',
    'tool',
    workbook,
    true,
    'none',
    'Preview saved local connections, then explicitly refresh ready results with one Undo per accepted connection.'
  ),
  definition(
    'data.queries',
    ['XL-DA01', 'XL-DA03', 'XL-DA04'],
    'Data',
    'Queries and connections',
    20,
    'Queries and refresh',
    'tool',
    workbook,
    false,
    'none',
    'Manage local file queries and refresh their results.'
  ),
  definition('data.sort', ['XL-DA06'], 'Data', 'Sort & Filter', 30, 'Sort', 'workbook', workbook, true),
  definition('data.filter', ['XL-DA07'], 'Data', 'Sort & Filter', 40, 'Filter', 'workbook', workbook, true),
  definition(
    'data.text-to-columns',
    ['XL-DA08'],
    'Data',
    'Data tools',
    45,
    'Text to Columns',
    'tool',
    workbook,
    false,
    'none',
    'Preview a literal-text split and its destination cells.'
  ),
  definition(
    'data.remove-duplicates',
    [],
    'Data',
    'Data tools',
    47,
    'Remove Duplicates',
    'tool',
    workbook,
    false,
    'none',
    'Preview matching rows and keep the first row with the selected values.'
  ),
  definition(
    'data.validation',
    ['XL-DA10'],
    'Data',
    'Data tools',
    50,
    'Data validation',
    'tool',
    workbook,
    false,
    'none',
    'Inspect or edit input rules, choices and prompts.'
  ),
  // Home identities inventory existing controls; no pictured feature ID is invented.
  ...homeCommandDefinitions
    .filter((command) => command.id !== 'styles.conditional')
    .map((command) =>
      definition(
        command.id,
        [],
        'Home',
        command.group,
        command.order,
        command.label,
        command.kind,
        command.supportedProfiles,
        command.requiresEdit,
        'home'
      )
    ),
  definition(
    'styles.conditional',
    [],
    'Home',
    'Styles',
    10,
    'Conditional formatting',
    'tool',
    workbook,
    false,
    'none',
    'Inspect or edit conditional formatting for cells.'
  ),
  definition(
    'formulas.insert-function',
    ['XL-F01', 'XL-F04', 'XL-F05'],
    'Formulas',
    'Function library',
    10,
    'Insert Function',
    'tool',
    workbook,
    false,
    'none',
    'Choose a qualified function and insert one formula into the captured cell.'
  ),
  definition(
    'formulas.auto-sum',
    ['XL-F02'],
    'Formulas',
    'Function library',
    20,
    'AutoSum',
    'workbook',
    workbook,
    true
  ),
  definition(
    'formulas.recently-used',
    ['XL-F03'],
    'Formulas',
    'Function library',
    30,
    'Recently Used',
    'tool',
    workbook,
    false,
    'none',
    'Choose a function previously inserted into a workbook.'
  ),
  definition(
    'formulas.financial',
    ['XL-F04'],
    'Formulas',
    'Function library',
    40,
    'Financial',
    'tool',
    workbook,
    false,
    'none',
    'Browse payment, present value and future value functions.'
  ),
  definition(
    'formulas.logical',
    ['XL-F04'],
    'Formulas',
    'Function library',
    50,
    'Logical',
    'tool',
    workbook,
    false,
    'none',
    'Browse logical functions.'
  ),
  definition(
    'formulas.text',
    ['XL-F04'],
    'Formulas',
    'Function library',
    60,
    'Text',
    'tool',
    workbook,
    false,
    'none',
    'Browse text functions.'
  ),
  definition(
    'formulas.date-time',
    ['XL-F05'],
    'Formulas',
    'Function library',
    70,
    'Date & Time',
    'tool',
    workbook,
    false,
    'none',
    'Browse date functions.'
  ),
  definition(
    'formulas.lookup-reference',
    ['XL-F05'],
    'Formulas',
    'Function library',
    80,
    'Lookup & Reference',
    'tool',
    workbook,
    false,
    'none',
    'Browse reference functions.'
  ),
  definition(
    'formulas.math-trig',
    ['XL-F05'],
    'Formulas',
    'Function library',
    90,
    'Math & Trig',
    'tool',
    workbook,
    false,
    'none',
    'Browse mathematical functions.'
  ),
  definition(
    'formulas.more-functions',
    ['XL-F05'],
    'Formulas',
    'Function library',
    100,
    'More Functions',
    'tool',
    workbook,
    false,
    'none',
    'Browse additional functions.'
  ),
  definition('view.preset-apply', ['XL-V01', 'XL-V03'], 'View', 'Local views', 5, 'Local view', 'localView', all),
  definition('view.preset-save', ['XL-V01', 'XL-V03'], 'View', 'Local views', 6, 'Save local view', 'localView', all),
  definition('view.preset-update', ['XL-V03'], 'View', 'Local views', 7, 'Update saved view', 'localView', all),
  definition('view.preset-delete', ['XL-V03'], 'View', 'Local views', 8, 'Delete saved view', 'localView', all),
  definition('view.source-defaults', ['XL-V04'], 'View', 'Local views', 9, 'Use workbook view', 'localView', all),
  definition(
    'view.page-break-preview',
    ['XL-V01'],
    'View',
    'Workbook views',
    20,
    'Page Break Preview',
    'hostJob',
    ['xlsx'],
    false,
    'none',
    'Read-only accepted page boundaries over the worksheet; break dragging is unavailable.'
  ),
  definition(
    'view.page-layout',
    ['XL-V02'],
    'View',
    'Workbook views',
    15,
    'Page Layout',
    'hostJob',
    ['xlsx'],
    false,
    'none',
    'Read-only accepted physical pages; edit in Normal view.'
  ),
  definition('view.normal', ['XL-V02'], 'View', 'Workbook views', 10, 'Normal', 'localView', all),
  definition('view.gridlines', ['XL-V04'], 'View', 'Show', 20, 'Gridlines', 'localView', all, false, 'toggle'),
  definition('view.headings', ['XL-V04'], 'View', 'Show', 30, 'Headings', 'localView', all, false, 'toggle'),
  definition('view.formula-bar', ['XL-V04'], 'View', 'Show', 40, 'Formula bar', 'localView', all, false, 'toggle'),
  definition(
    'formulas.error-checking',
    ['XL-F11'],
    'Formulas',
    'Formula auditing',
    130,
    'Error Checking',
    'tool',
    workbook,
    false,
    'none',
    'Find cell errors and unresolved formula results in this worksheet.'
  ),
  definition(
    'view.show-formulas',
    ['XL-F10', 'XL-V04'],
    'Formulas',
    'Formula auditing',
    140,
    'Show formulas',
    'localView',
    all,
    false,
    'toggle'
  ),
  definition(
    'formulas.watch',
    ['XL-F12'],
    'Formulas',
    'Formula auditing',
    150,
    'Watch Window',
    'tool',
    workbook,
    false,
    'none',
    'Watch up to 32 cell values and formulas.'
  ),
  definition('view.zoom', ['XL-V05'], 'View', 'Zoom', 60, 'Zoom', 'localView', all, false, 'zoom'),
  definition('view.zoom-100', ['XL-V05'], 'View', 'Zoom', 70, '100%', 'localView', all),
  definition('view.zoom-selection', ['XL-V05'], 'View', 'Zoom', 80, 'Zoom to selection', 'localView', all),
  definition('view.freeze', ['XL-V07'], 'View', 'Window', 90, 'Freeze Panes', 'workbook', workbook, true, 'freeze'),
]);
const byId = new Map(ribbonCommands.map((command) => [command.id, command]));
export function ribbonCommand(id: RibbonCommandId): RibbonCommandDefinition {
  const command = byId.get(id);
  if (!command) throw Error('Unknown ribbon command');
  return command;
}
export const ribbonToolCommands: Readonly<Record<WorkbookTool, RibbonCommandId>> = Object.freeze({
  accessibility: 'review.accessibility',
  notes: 'review.notes',
  illustrations: 'insert.shapes',
  pageStyles: 'page.themes',
  checkbox: 'insert.checkbox',
  symbols: 'insert.symbols',
  hyperlink: 'insert.link',
  actions: 'actions.library',
  review: 'review.statistics',
  fill: 'home.fill-series',
  replace: 'home.replace-preview',
  special: 'home.special-open',
  duplicates: 'data.remove-duplicates',
  sort: 'data.sort',
  createNames: 'names.create-from-selection',
  calculation: 'formulas.calculation-options',
  names: 'names.manager',
  objects: 'objects.manage',
  print: 'print.setup',
  conditional: 'styles.conditional',
  pivot: 'pivots.context',
  pivotDesign: 'pivots.design',
  validation: 'data.validation',
  references: 'references.manage',
  queries: 'data.queries',
  functions: 'formulas.insert-function',
  watch: 'formulas.watch',
  errors: 'formulas.error-checking',
  textColumns: 'data.text-to-columns',
  trace: 'formulas.trace-precedents',
});
export const ribbonFunctionCategories = Object.freeze({
  'formulas.financial': 'Financial',
  'formulas.logical': 'Logical',
  'formulas.text': 'Text',
  'formulas.date-time': 'DateTime',
  'formulas.lookup-reference': 'Lookup',
  'formulas.math-trig': 'Math',
  'formulas.more-functions': 'More',
} satisfies Partial<Record<RibbonCommandId, FunctionCategory>>);
export const ribbonCommandTools: Readonly<Partial<Record<RibbonCommandId, WorkbookTool>>> = Object.freeze({
  'review.accessibility': 'accessibility',
  'insert.note': 'notes',
  'review.notes': 'notes',
  'insert.recommended-charts': 'objects',
  'insert.model3d': 'illustrations',
  'insert.smartart': 'illustrations',
  'insert.wordart': 'illustrations',
  'insert.icons': 'illustrations',
  'insert.shapes': 'illustrations',
  'insert.text-box': 'illustrations',
  'page.themes': 'pageStyles',
  'page.background': 'pageStyles',
  'insert.checkbox': 'checkbox',
  'insert.symbols': 'symbols',
  'insert.link': 'hyperlink',
  'pivots.context': 'pivot',
  'pivots.design': 'pivotDesign',
  'home.fill-series': 'fill',
  'home.replace-preview': 'replace',
  'home.special-open': 'special',
  'actions.new': 'actions',
  'actions.library': 'actions',
  'review.statistics': 'review',
  'review.spelling': 'review',
  'formulas.calculation-options': 'calculation',
  'formulas.trace-precedents': 'trace',
  'formulas.trace-dependents': 'trace',
  ...Object.fromEntries(Object.keys(ribbonFunctionCategories).map((id) => [id, 'functions' as const])),
  'formulas.error-checking': 'errors',
  'formulas.watch': 'watch',
  'formulas.insert-function': 'functions',
  'formulas.recently-used': 'functions',
  'names.create-from-selection': 'createNames',
  'names.manager': 'names',
  'names.define': 'names',
  'objects.manage': 'objects',
  'references.manage': 'references',
  'print.setup': 'print',
  'print.margins': 'print',
  'print.orientation': 'print',
  'print.paper': 'print',
  'print.area': 'print',
  'print.titles': 'print',
  'print.scale': 'print',
  'data.remove-duplicates': 'duplicates',
  'data.sort': 'sort',
  'data.text-to-columns': 'textColumns',
  'data.refresh-all': 'queries',
  'data.queries': 'queries',
  'data.validation': 'validation',
  'styles.conditional': 'conditional',
});
export const ribbonTabOrder: readonly RibbonTab[] = Object.freeze([
  'Home',
  'Insert',
  'Draw',
  'Page Layout',
  'Formulas',
  'Data',
  'Review',
  'View',
  'Automate',
  'PivotTable Analyze',
  'Design',
]);

export interface RibbonCapability {
  state: 'enabled' | 'disabled' | 'hidden';
  reason?: string;
}
export interface RibbonAvailability {
  profile: RibbonProfile;
  active: boolean;
  canEdit: boolean;
  busy: boolean;
  hasHandler: boolean;
  capability?: RibbonCapability;
}
export function ribbonCapability(id: RibbonCommandId, scope: RibbonAvailability): RibbonCapability {
  const command = ribbonCommand(id);
  if (!scope.hasHandler) return { state: 'hidden', reason: 'This host has not supplied this command.' };
  if (!command.supportedProfiles.includes(scope.profile))
    return { state: 'hidden', reason: 'This workbook format does not support this command.' };
  if (!scope.active) return { state: 'disabled', reason: 'Activate this workbook first.' };
  if (scope.capability?.state === 'hidden' || scope.capability?.state === 'disabled') return scope.capability;
  if (command.requiresEdit && !scope.canEdit) return { state: 'disabled', reason: 'This workbook is read-only.' };
  if (command.requiresEdit && scope.busy)
    return { state: 'disabled', reason: 'Wait for the workbook operation to finish.' };
  // Inspection and local display preferences remain reachable during owner jobs.
  return { state: 'enabled' };
}
export interface RibbonBinding {
  readonly documentId: string;
  readonly state: string;
  readonly sheet: string;
  readonly selection: string;
  readonly activationEpoch: number;
}
export interface RibbonDispatchScope {
  captureBinding: () => RibbonBinding;
  isCurrent: (binding: RibbonBinding) => boolean;
  availability: (id: RibbonCommandId) => Omit<RibbonAvailability, 'hasHandler'>;
}
export type RibbonArguments =
  | undefined
  | { pen: InkPenId }
  | { zoom: number }
  | { value: boolean }
  | { mode: 'selection' | 'row' | 'column' | 'none' }
  | HomeArguments;
export type RibbonCommandHandler<T = unknown> = (binding: RibbonBinding, args: RibbonArguments) => T | Promise<T>;
export type RibbonDispatchResult<T> =
  | { outcome: 'refused'; reason: string }
  | { outcome: 'stale'; binding: RibbonBinding }
  | { outcome: 'completed'; binding: RibbonBinding; value: T; current: boolean }
  | { outcome: 'failed'; binding: RibbonBinding; error: unknown; current: boolean };
function capture(scope: RibbonDispatchScope): RibbonBinding {
  const binding = scope.captureBinding();
  if (
    ['documentId', 'state', 'sheet', 'selection'].some(
      (key) =>
        typeof binding[key as keyof RibbonBinding] !== 'string' ||
        !(binding[key as keyof RibbonBinding] as string).length
    ) ||
    !Number.isSafeInteger(binding.activationEpoch) ||
    binding.activationEpoch < 0
  )
    throw Error('Ribbon command requires an accepted document binding');
  return Object.freeze({
    documentId: binding.documentId,
    state: binding.state,
    sheet: binding.sheet,
    selection: binding.selection,
    activationEpoch: binding.activationEpoch,
  });
}
function argumentsFor(command: RibbonCommandDefinition, args: RibbonArguments): RibbonArguments {
  if (command.argument === 'home') return normalizeHomeArguments(command.id as HomeCommandId, args);
  const keys = args && Object.keys(args);
  if (command.argument === 'inkPreset' && keys?.length === 1 && 'pen' in args!) {
    return Object.freeze({ pen: inkPen(args.pen).id });
  }
  if (command.argument === 'none' && args === undefined) return undefined;
  if (
    command.argument === 'zoom' &&
    keys?.length === 1 &&
    'zoom' in args! &&
    Number.isInteger(args.zoom) &&
    args.zoom >= 10 &&
    args.zoom <= 400
  )
    return Object.freeze({ zoom: args.zoom });
  if (command.argument === 'toggle' && keys?.length === 1 && 'value' in args! && typeof args.value === 'boolean')
    return Object.freeze({ value: args.value });
  if (
    command.argument === 'freeze' &&
    keys?.length === 1 &&
    'mode' in args! &&
    ['selection', 'row', 'column', 'none'].includes(args.mode)
  )
    return Object.freeze({ mode: args.mode as 'selection' | 'row' | 'column' | 'none' });
  throw Error('Invalid ribbon command arguments');
}
/** Capture before any async handler. A completed accepted write remains completed
 * when its normal revision advance makes the initiating UI binding non-current. */
export async function dispatchRibbonCommand<T>(
  id: RibbonCommandId,
  scope: RibbonDispatchScope,
  handler?: RibbonCommandHandler<T>,
  args?: RibbonArguments
): Promise<RibbonDispatchResult<T>> {
  const capability = ribbonCapability(id, { ...scope.availability(id), hasHandler: !!handler });
  if (capability.state !== 'enabled')
    return { outcome: 'refused', reason: capability.reason ?? 'This command is unavailable.' };
  const binding = capture(scope),
    normalized = argumentsFor(ribbonCommand(id), args);
  if (!scope.isCurrent(binding)) return { outcome: 'stale', binding };
  try {
    const value = await handler!(binding, normalized);
    return { outcome: 'completed', binding, value, current: scope.isCurrent(binding) };
  } catch (error) {
    return { outcome: 'failed', binding, error, current: scope.isCurrent(binding) };
  }
}
