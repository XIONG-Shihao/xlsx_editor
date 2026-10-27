/** Conditional formatting and data validation rules. */
import { indexedRulesDefinition, indexedRulesPlan } from '../indexed-rules.mjs';
import fs from 'node:fs/promises';
import path from 'node:path';
import { indexedPackageEdit } from '../workbook-package.mjs';
import { matchesHandles } from './shared.mjs';
import { finishPackageEdit } from './package-edit.mjs';
import { refusal } from '../refusal-codes.mjs';
export const name = 'rules';
export const handles = { operation: ['setConditionalRules', 'setValidationRules', 'addConditionalRule'] };
export const matches = (commands) => matchesHandles(handles, commands);
export async function run(ctx, commands) {
  const { current } = ctx;
  const transform = commands[0];
  if (transform.operation === 'addConditionalRule') return addRule(ctx, commands);
  const bytes = await fs.readFile(path.join(current.root, 'metadata.xlsx'));
  let prepared = indexedRulesPlan(bytes, current.view, transform);
  ctx.recheck();
  const source = await ctx.rewriteSource();
  if (path.basename(source.path) === 'export.xlsx' && transform.operation === 'setConditionalRules') {
    const exported = JSON.parse(await fs.readFile(path.join(current.root, 'export-metadata.json'), 'utf8'));
    if (typeof exported.replacements?.['xl/styles.xml'] === 'string')
      prepared = indexedRulesPlan(bytes, current.view, transform, {
        stylesXml: exported.replacements['xl/styles.xml'],
      });
  }
  ctx.recheck();
  await ctx.native({
    action: 'workbook',
    operation: 'artifact-write',
    root: current.root,
    name: 'package-metadata.json',
    data: JSON.stringify(prepared.plan),
  });
  const result = await ctx.native({
    action: 'workbook',
    operation: 'package-edit',
    root: current.root,
    version: current.version,
    source: path.basename(source.path),
    ...ctx.rewriteFields(source),
  });
  ctx.rewriteCheck(source.candidate);
  const { next, root } = source.candidate;
  const indexed = await ctx.reindexRewrite(source, result, current.view.calculationPolicy),
    definition = indexedRulesDefinition(
      await fs.readFile(path.join(root, 'metadata.xlsx')),
      indexed.view,
      transform.sheet,
      transform.operation === 'setConditionalRules' ? 'conditional' : 'validation'
    );
  if (!definition.canEdit) throw refusal('session-unsupported', 'Rule candidate is not qualified');
  indexed.journal = { previous: null, edits: [] };
  ctx.recheck();
  ctx.publish(next, indexed);
  return { view: ctx.view(next), edits: [prepared.command], commands: [prepared.command] };
}
async function addRule(ctx, commands) {
  const { current } = ctx;
  const transform = commands[0];
  const metadataBytes = await fs.readFile(path.join(current.root, 'metadata.xlsx'));
  const metadata = indexedPackageEdit(metadataBytes, transform, current.view, []);
  return finishPackageEdit(ctx, commands, { metadata, base: current });
}
