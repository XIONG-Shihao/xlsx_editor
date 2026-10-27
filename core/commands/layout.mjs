/** Column width, row height, freeze panes, unhide and merge. */
import fs from 'node:fs/promises';
import path from 'node:path';
import { indexedLayoutFragment, indexedMergeFragment } from '../workbook-package.mjs';
import { projectWorkbookCommands } from '../workbook-projection.mjs';
import { rangeBounds } from '../workbook-commands.mjs';
import { empty, matchesHandles } from './shared.mjs';
import { refusal } from '../refusal-codes.mjs';
export const name = 'layout';
export const handles = {
  operation: ['resizeRow', 'resizeColumn', 'freeze', 'unhideAll', 'mergeCells', 'unmergeCells'],
};
export const matches = (commands) => matchesHandles(handles, commands);
export async function run(ctx, commands) {
  const { current } = ctx;
  const command = commands[0],
    merge = ['mergeCells', 'unmergeCells'].includes(command.operation);
  if (merge && rangeBounds(command.range).size < 2)
    throw refusal('session-bad-input', 'Select at least two cells to merge or unmerge');
  const projected =
    merge || command.operation === 'unhideAll'
      ? null
      : await projectWorkbookCommands(empty(current.view), commands, {});
  const source = await ctx.rewriteSource(false),
    sheet = current.view.sheets.find((s) => s.name === command.sheet);
  let layout;
  if (merge) {
    const xml = indexedMergeFragment(await fs.readFile(path.join(current.root, 'metadata.xlsx')), command);
    await ctx.native({
      action: 'workbook',
      operation: 'artifact-write',
      root: current.root,
      name: 'layout-fragment.xml',
      data: xml,
    });
    layout = {
      tag: 'mergeCells',
      kind: command.operation,
      range: command.range,
      // So that a merge beside a table is allowed and one inside it is not.
      tableRanges: (current.view.tables ?? [])
        .filter((table) => table.sheet === command.sheet)
        .map((table) => table.ref),
    };
  } else if (command.operation === 'unhideAll') {
    if (Object.keys(command).sort().join(',') !== 'operation,sheet')
      throw refusal('internal-invalid-field', 'Invalid Unhide command');
    layout = {};
  } else if (command.operation === 'resizeRow') layout = { row: command.index, height: command.height };
  else {
    const fragment = indexedLayoutFragment(
      await fs.readFile(path.join(current.root, 'metadata.xlsx')),
      command,
      projected.view
    );
    await ctx.native({
      action: 'workbook',
      operation: 'artifact-write',
      root: current.root,
      name: 'layout-fragment.xml',
      data: fragment.xml,
    });
    layout = { tag: fragment.tag };
  }
  const result = await ctx.native({
    action: 'workbook',
    operation: 'layout',
    root: current.root,
    ...(source.candidate.enabled ? { version: current.version, command: structuredClone(command) } : {}),
    source: path.basename(source.path),
    ...ctx.rewriteFields(source),
    sheet: sheet.path,
    ...layout,
  });
  ctx.rewriteCheck(source.candidate);
  const { next, root } = source.candidate;
  const indexed = await ctx.reindexRewrite(source, result, current.view.calculationPolicy);
  ctx.recheck();
  ctx.publish(next, indexed);
  return { view: ctx.view(next), edits: structuredClone(commands) };
}
