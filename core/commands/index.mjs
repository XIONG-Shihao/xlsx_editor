/** The command modules of the indexed session and the order they are tried in.
 *
 * A module owns one family of commands. It exports:
 *   name     - a short label
 *   handles  - the command `type`s or `operation`s it owns (used by the dispatch test)
 *   matches  - does it own this whole command list (first match wins)
 *   run      - `run(ctx, commands)`, the check and the action for those commands
 * `ctx` is built by the session for one request (see `#commandContext` in
 * indexed-session.mjs); it is how a module reaches the native engine and the
 * session's private state without that state being public. */
import * as sheetBackground from './sheet-background.mjs';
import * as pageTheme from './page-theme.mjs';
import * as booleanCell from './boolean-cell.mjs';
import * as notes from './notes.mjs';
import * as hyperlinks from './hyperlinks.mjs';
import * as clearHyperlinks from './clear-hyperlinks.mjs';
import * as calculationOptions from './calculation-options.mjs';
import * as recalculate from './recalculate.mjs';
import * as rules from './rules.mjs';
import * as importCells from './import-cells.mjs';
import * as print from './print.mjs';
import * as objects from './objects.mjs';
import * as pivots from './pivots.mjs';
import * as tables from './tables.mjs';
import * as names from './names.mjs';
import * as structure from './structure.mjs';
import * as sortFilter from './sort-filter.mjs';
import * as sheets from './sheets.mjs';
import * as layout from './layout.mjs';
import * as cellEdits from './cell-edits.mjs';

export { admitAgainstMetadata } from './guards.mjs';

/** Commands that run straight away, before the pivot and validation guards. */
export const directModules = [
  sheetBackground,
  pageTheme,
  booleanCell,
  notes,
  hyperlinks,
  clearHyperlinks,
  calculationOptions,
  recalculate,
];
/** Commands that run only after the pivot and validation guards have passed. */
export const guardedModules = [
  rules,
  importCells,
  print,
  objects,
  pivots,
  tables,
  names,
  structure,
  sortFilter,
  sheets,
  layout,
  cellEdits,
];
export const commandModules = [...directModules, ...guardedModules];

/** The first module that owns this command list, or undefined. */
export const routeCommands = (modules, commands) => modules.find((module) => module.matches(commands));
