/** Typing, fill, paste, clear and formats: ordinary cell content. */
import { refusal, reworded } from '../refusal-codes.mjs';
import { assertValidationEdits } from '../rules-model.mjs';
import { withSpillDimensions } from '../indexed-spills.mjs';
import { projectCellCommands } from '../workbook-projection.mjs';
import { cellPosition, columnName, rangeBounds, resolveWorkbookEdits } from '../workbook-commands.mjs';
import { nativeValue, empty, own } from './shared.mjs';
export const name = 'cell-edits';
export const handles = {
  type: [
    'formatRange',
    'clearRange',
    'fillRange',
    'fillSeries',
    'fillAcrossSheets',
    'pasteValues',
    'pasteFormulas',
    'setCells',
    'updateRows',
    undefined,
  ],
};
export const matches = (commands) =>
  commands.every((command) => !command.operation && handles.type.includes(command.type));
export async function run(ctx, commands) {
  const { current } = ctx;
  try {
    const groups = [];
    for (const command of commands) {
      const kind =
        command.type === 'fillAcrossSheets'
          ? 'across'
          : command.type === 'updateRows'
            ? 'keys'
            : command.type === 'formatRange'
              ? 'format'
              : ['clearRange', 'fillRange', 'fillSeries', 'pasteValues', 'pasteFormulas'].includes(command.type)
                ? 'bulk'
                : 'cells';
      if (!['bulk', 'keys'].includes(kind) && groups.at(-1)?.kind === kind) groups.at(-1).commands.push(command);
      else groups.push({ kind, commands: [command] });
    }
    let working = { ...current, view: empty(current.view) },
      allEdits = [],
      contentChanged = false,
      formulaChanged = false;
    const changedRanges = [];
    for (const group of groups) {
      if (group.kind === 'bulk' || group.kind === 'across') {
        // Compact recovery replays previously separate user actions in one
        // batch. Paste Values must observe earlier formula/input edits.
        if (
          contentChanged &&
          group.commands.some((c) => c.type === 'pasteValues') &&
          working.view.sheets.some((s) => s.storedFormulaCount > 0)
        ) {
          const calculated = await ctx.calculate({
            root: working.root,
            version: working.version,
            view: working.view,
            changedRanges,
            formulaChanged,
            revision: `${current.index.revision}:draft:${ctx.id}:${ctx.peekNextState()}`,
          });
          working.version = calculated.version;
        }
        const normalized = group.commands.map((command) => {
          const sheet = working.view.sheets.find((s) => s.name === command.sheet),
            source = working.view.sheets.find((s) => s.name === (command.sourceSheet ?? command.sheet));
          if (!sheet || !source) throw refusal('internal-unknown-reference', 'Unknown worksheet');
          return { ...command, sheet: sheet.path, ...(command.sourceSheet ? { sourceSheet: source.path } : {}) };
        });
        const request = {
          action: 'workbook',
          operation: 'bulk',
          root: working.root,
          version: working.version,
          commands: normalized,
        };
        let result;
        result = await ctx.native(request);
        if (group.kind === 'across') {
          const changesContent = group.commands[0].mode !== 'formats';
          if (typeof result.contentChanged !== 'boolean' || result.contentChanged !== changesContent)
            throw refusal('internal-invalid-field', 'Invalid native Across Worksheets content-change receipt.');
          if (
            !changesContent &&
            (result.formulaTouched !== false ||
              !result.formulaDelta ||
              Object.values(result.formulaDelta).some((delta) => delta !== 0))
          )
            throw refusal('session-stale', 'Native Formats receipt changed formula content.');
        }
        working.version = result.version;
        formulaChanged ||= result.formulaTouched;
        for (const sheet of working.view.sheets) {
          const delta = result.formulaDelta[sheet.path] ?? 0;
          sheet.storedFormulaCount += delta;
          formulaChanged ||= delta !== 0;
        }
        const changesContent = group.kind !== 'across' || result.contentChanged;
        if (changesContent) changedRanges.push(...group.commands.map((c) => ({ sheet: c.sheet, range: c.range })));
        for (const command of group.commands) {
          const sheet = working.view.sheets.find((s) => s.name === command.sheet),
            bounds = rangeBounds(command.range);
          sheet.rows = Math.max(sheet.rows, bounds.last.row + 1);
          sheet.columns = Math.max(sheet.columns, bounds.last.col + 1);
        }
        allEdits.push(...structuredClone(group.commands));
        contentChanged ||= changesContent;
      } else if (group.kind === 'format') {
        const normalized = group.commands.map((command) => {
          if (Object.keys(command).some((key) => !['type', 'sheet', 'range', 'format'].includes(key)))
            throw refusal('internal-invalid-field', 'Invalid range format command');
          const sheet = working.view.sheets.find((sheet) => sheet.name === command.sheet);
          if (!sheet) throw refusal('internal-unknown-reference', 'Unknown worksheet');
          const bounds = rangeBounds(command.range);
          return {
            sheet: sheet.path,
            range: `${columnName(bounds.first.col)}${bounds.first.row + 1}:${columnName(bounds.last.col)}${bounds.last.row + 1}`,
            format: command.format,
          };
        });
        const result = await ctx.native({
          action: 'workbook',
          operation: 'format',
          root: working.root,
          version: working.version,
          commands: normalized,
        });
        working.version = result.version;
        for (const command of group.commands) {
          const sheet = working.view.sheets.find((sheet) => sheet.name === command.sheet),
            bounds = rangeBounds(command.range);
          sheet.rows = Math.max(sheet.rows, bounds.last.row + 1);
          sheet.columns = Math.max(sheet.columns, bounds.last.col + 1);
        }
        // These compact commands also replay during recovery alongside literal
        // edits. No intermediate native version is published to the caller.
        allEdits.push(...structuredClone(group.commands));
      } else {
        if (group.kind === 'keys') group.commands = await ctx.keyedCommands(working, group.commands[0]);
        let view = await ctx.inputs(group.commands, working);
        const shared = [];
        let cellCommands = group.commands,
          validated;
        const releases = new Map();
        for (const command of group.commands)
          for (const edit of command.cells ?? [command])
            if (own(edit, 'value') || own(edit, 'formula')) {
              const sheet = view.sheets.find((s) => s.name === command.sheet),
                p = cellPosition(edit.cell),
                ref = `${columnName(p.col)}${p.row + 1}`,
                cell = sheet.cells[ref];
              if (!cell?.spill) continue;
              if (cell.spill.schema !== 1 || cell.spill.role !== 'master' || cell.spill.master !== ref)
                throw refusal('session-unsupported', 'A spill follower cannot be edited independently of its master');
              releases.set(`${sheet.path}\0${ref}`, { sheet: sheet.path, cell: ref, name: sheet.name });
            }
        if (releases.size) {
          // Validate expected-old guards against the accepted master before
          // clearing its private complete geometry, then reproject blanks.
          validated ??= resolveWorkbookEdits(view, group.commands);
          cellCommands = validated.map(({ expectedValue, expectedFormula, expectedRawValue, ...edit }) => edit);
          for (const release of releases.values()) {
            const result = await ctx.native({
              action: 'workbook',
              operation: 'spill-release',
              root: working.root,
              version: working.version,
              sheet: release.sheet,
              cell: release.cell,
            });
            working.version = result.version;
            working.view.sheets.find((sheet) => sheet.name === release.name).storedFormulaCount--;
            changedRanges.push({ sheet: release.name, range: result.releasedRange });
            formulaChanged = true;
          }
          view = await ctx.inputs(cellCommands, working);
        }
        for (const command of group.commands)
          for (const edit of command.cells ?? [command])
            if (own(edit, 'value') || own(edit, 'formula')) {
              const sheet = view.sheets.find((s) => s.name === command.sheet),
                p = cellPosition(edit.cell),
                ref = `${columnName(p.col)}${p.row + 1}`;
              if (sheet.cells[ref]?.formulaType === 'shared') shared.push({ sheet: sheet.path, range: ref });
            }
        if (shared.length) {
          // Validate against the accepted representation before expanding storage.
          validated ??= resolveWorkbookEdits(view, group.commands);
          cellCommands = validated.map(({ expectedValue, expectedFormula, expectedRawValue, ...edit }) => edit);
          const normalized = await ctx.native({
            action: 'workbook',
            operation: 'normalize-shared',
            root: working.root,
            version: working.version,
            ranges: shared,
          });
          if (normalized.version !== working.version) {
            working.version = normalized.version;
            formulaChanged = true;
            view = await ctx.inputs(group.commands, working);
          }
        }
        const projected = projectCellCommands(view, cellCommands);
        if (validated) projected.edits = validated;
        for (const sheet of working.view.sheets) {
          const rules = ctx.ruleDefinitions.validation[sheet.name];
          if (rules?.canEdit)
            assertValidationEdits(
              rules.rules,
              projected.edits.filter((edit) => edit.sheet === sheet.name)
            );
        }
        const byCell = new Map();
        for (const edit of projected.edits) {
          const key = `${edit.sheet}\0${edit.cell}`;
          if (!byCell.has(key)) byCell.set(key, []);
          byCell.get(key).push(edit);
        }
        const changes = projected.changes.map((change) => ({
          ...change,
          sheet: view.sheets.find((sheet) => sheet.name === change.sheet).path,
          edits: byCell.get(`${change.sheet}\0${change.cell}`),
        }));
        for (let start = 0; start < changes.length; start += 256) {
          const result = await ctx.native({
            action: 'workbook',
            operation: 'patch',
            root: working.root,
            version: working.version,
            changes: nativeValue(changes.slice(start, start + 256)),
          });
          working.version = result.version;
        }
        for (const change of projected.changes) {
          const sheet = working.view.sheets.find((sheet) => sheet.name === change.sheet),
            old = view.sheets.find((sheet) => sheet.name === change.sheet).cells[change.cell];
          sheet.storedFormulaCount +=
            Number(!!(change.value.formula || change.value.formulaType)) - Number(!!(old?.formula || old?.formulaType));
          const p = cellPosition(change.cell);
          sheet.rows = Math.max(sheet.rows, p.row + 1);
          sheet.columns = Math.max(sheet.columns, p.col + 1);
        }
        allEdits.push(...projected.edits);
        for (const edit of projected.edits)
          if (own(edit, 'value') || own(edit, 'formula')) {
            contentChanged = true;
            changedRanges.push({ sheet: edit.sheet, range: edit.cell });
            formulaChanged ||=
              own(edit, 'formula') || !!view.sheets.find((s) => s.name === edit.sheet).cells[edit.cell]?.formula;
          }
      }
      ctx.recheck();
    }
    const next = ctx.nextState();
    if (
      contentChanged &&
      (working.view.sheets.some((sheet) => sheet.storedFormulaCount > 0) ||
        working.view.calculationPolicy?.mode === 'manual')
    ) {
      const calculated = await ctx.calculate({
        root: working.root,
        version: working.version,
        view: working.view,
        changedRanges,
        formulaChanged,
        revision: `${current.index.revision}:draft:${ctx.id}:${next}`,
      });
      working.version = calculated.version;
      working.view = withSpillDimensions(working.view, calculated);
      working.spillOwnership ||= !!calculated.spillDimensions;
    }
    ctx.recheck();
    ctx.publish(next, { ...working, journal: { previous: current.journal, edits: allEdits } });
    return { view: ctx.view(next), edits: allEdits };
  } catch (error) {
    throw reworded(
      error,
      String(error).replace('requires the shared structural writer', 'is not qualified for native preservation')
    );
  }
}
