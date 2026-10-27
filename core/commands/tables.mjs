/** Create tables; rename, resize and change table columns. */
import { indexedTablePlan } from '../indexed-tables.mjs';
import fs from 'node:fs/promises';
import path from 'node:path';
import { indexedPackageEdit } from '../workbook-package.mjs';
import { columnName, rangeBounds } from '../workbook-commands.mjs';
import { matchesHandles } from './shared.mjs';
import { finishPackageEdit } from './package-edit.mjs';
import { refusal } from '../refusal-codes.mjs';
export const name = 'tables';
export const handles = {
  operation: [
    'createTable',
    'renameTable',
    'resizeTable',
    'renameTableColumn',
    'insertTableColumn',
    'deleteTableColumn',
  ],
};
export const matches = (commands) => matchesHandles(handles, commands);
export async function run(ctx, commands) {
  const { current } = ctx;
  const transform = commands[0];
  if (transform.operation === 'createTable') return createTable(ctx, commands);
  const metadata = await indexedTablePlan(
    await fs.readFile(path.join(current.root, 'metadata.xlsx')),
    current.view,
    transform,
    { readRange: (sheet, range) => ctx.read(current, sheet, range) }
  );
  ctx.recheck();
  await ctx.native({
    action: 'workbook',
    operation: 'artifact-write',
    root: current.root,
    name: 'package-metadata.json',
    data: JSON.stringify(metadata),
  });
  const source = await ctx.rewriteSource(),
    result = await ctx.native({
      action: 'workbook',
      operation: 'package-edit',
      root: current.root,
      version: current.version,
      source: path.basename(source.path),
      ...ctx.rewriteFields(source),
    });
  ctx.rewriteCheck(source.candidate);
  const { next, root } = source.candidate;
  const indexed = await ctx.reindexRewrite(source, result, current.view.calculationPolicy);
  const accepted = indexed.view.tables?.find((table) => table.path === metadata.tablePart);
  if (
    !accepted ||
    accepted.name !== (metadata.renameTable?.new ?? metadata.tableName) ||
    accepted.ref !== (metadata.range ?? metadata.priorRef)
  )
    throw refusal('session-stale', 'Table candidate changed its admitted definition');
  if (
    metadata.renameTableColumn &&
    accepted.columns?.[metadata.renameTableColumn.index]?.name !== metadata.renameTableColumn.new
  )
    throw refusal('session-stale', 'Table candidate changed its admitted column label');
  if (metadata.columnChange) {
    const prior = current.view.tables.find((t) => t.path === metadata.tablePart).columns.map((c) => c.name);
    prior.splice(
      metadata.columnChange.index,
      metadata.columnChange.remove ? 1 : 0,
      ...(metadata.columnChange.remove ? [] : [metadata.columnChange.label])
    );
    if (JSON.stringify(accepted.columns?.map((c) => c.name)) !== JSON.stringify(prior))
      throw refusal('session-stale', 'Table candidate changed its admitted column structure');
  }
  if (metadata.columnChange || metadata.renameTableColumn) {
    const bounds = rangeBounds(accepted.ref),
      headerRange =
        columnName(bounds.first.col) +
        (bounds.first.row + 1) +
        ':' +
        columnName(bounds.last.col) +
        (bounds.first.row + 1),
      header = (await ctx.read(indexed, transform.sheet, headerRange)).cells;
    for (let i = 0; i < accepted.columns.length; i++) {
      const cell = header[columnName(bounds.first.col + i) + (bounds.first.row + 1)];
      if (!cell || cell.formula || cell.formulaType || cell.value !== accepted.columns[i].name)
        throw refusal('internal-mismatch', 'Table candidate literal headers differ from their admitted definitions');
    }
  }
  indexed.journal = { previous: null, edits: [] };
  ctx.recheck();
  ctx.publish(next, indexed);
  return { view: ctx.view(next), edits: structuredClone(commands), commands: structuredClone(commands) };
}
async function createTable(ctx, commands) {
  const { current } = ctx;
  const transform = commands[0];
  let headers = [];
  const bounds = rangeBounds(transform.range);
  if (bounds.size > 10000 || bounds.rows < 2) throw refusal('internal-invalid-field', 'Invalid table range');
  const row = bounds.first.row + 1,
    range = `${columnName(bounds.first.col)}${row}:${columnName(bounds.last.col)}${row}`;
  const read = await ctx.read(current, transform.sheet, range);
  headers = Object.entries(read.cells).map(([cell, value]) => ({
    sheet: transform.sheet,
    cell,
    value: value.value,
  }));
  const metadataBytes = await fs.readFile(path.join(current.root, 'metadata.xlsx'));
  const metadata = indexedPackageEdit(metadataBytes, transform, current.view, headers);
  return finishPackageEdit(ctx, commands, { metadata, base: current });
}
