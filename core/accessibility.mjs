import { FORMAT, matchesFormat } from './format.mjs';
/** Read-only, captured workbook content findings. This is not a compliance certificate. */
import { noteMembershipCell, noteMembershipSheet } from './note-membership.mjs';
import { refusal } from './refusal-codes.mjs';
export const ACCESSIBILITY_LIMITS = Object.freeze({
  tables: 64,
  columns: 1000,
  headerCells: 16000,
  findings: 16256,
  skips: 512,
});
export const ACCESSIBILITY_RULES = Object.freeze(['missing-alt-text', 'blank-table-header', 'duplicate-table-header']);
const fields = (v, keys, label) => {
  if (!v || typeof v !== 'object' || Array.isArray(v) || Object.keys(v).sort().join(',') !== [...keys].sort().join(','))
    throw refusal('internal-invalid-field', 'Invalid ' + label + ' fields.');
};
const text = (v, max, label) => {
  if (typeof v !== 'string' || !v.length || v.length > max)
    throw refusal('internal-invalid-field', 'Invalid ' + label + '.');
  return v;
};
const count = (v, max) => {
  if (!Number.isSafeInteger(v) || v < 0 || v > max)
    throw refusal('internal-invalid-field', 'Invalid accessibility coverage count.');
};
export function accessibilityFields(v, { transport = false } = {}) {
  fields(v, ['operation', 'documentId', 'state', ...(transport ? ['kind', 'sessionId'] : [])], 'accessibility request');
  if (v.operation !== 'accessibility' || (transport && v.kind !== 'indexed'))
    throw refusal('internal-invalid-field', 'Invalid accessibility operation.');
  for (const k of ['documentId', 'state', ...(transport ? ['sessionId'] : [])]) text(v[k], 128, k);
  return v;
}
export function accessibilityReport(v) {
  fields(v, ['schema', 'scope', 'status', 'rules', 'findings', 'coverage'], 'accessibility report');
  if (
    !matchesFormat(v.schema, 'SCHEMA_ACCESSIBILITY_V1') ||
    v.scope !== 'workbook' ||
    !['complete', 'partial'].includes(v.status) ||
    JSON.stringify(v.rules) !== JSON.stringify(ACCESSIBILITY_RULES)
  )
    throw refusal('internal-invalid-field', 'Invalid accessibility scope/rules.');
  if (!Array.isArray(v.findings) || v.findings.length > ACCESSIBILITY_LIMITS.findings)
    throw refusal('note-limit', 'Accessibility findings exceed their bound.');
  const seen = new Set();
  for (const f of v.findings) {
    fields(f, ['rule', 'message', 'sheet', 'cell', 'objectId', 'tableName', 'relatedCells'], 'accessibility finding');
    if (!ACCESSIBILITY_RULES.includes(f.rule))
      throw refusal('internal-unknown-reference', 'Unknown accessibility rule.');
    text(f.message, 1024, 'finding message');
    noteMembershipSheet(f.sheet);
    noteMembershipCell(f.cell);
    if (f.objectId !== null) text(f.objectId, 128, 'object identity');
    if (f.tableName !== null) text(f.tableName, 255, 'table name');
    if (
      f.rule === 'missing-alt-text'
        ? f.objectId === null || f.tableName !== null
        : f.objectId !== null || f.tableName === null
    )
      throw refusal('note-malformed', 'Finding has the wrong location kind.');
    if (!Array.isArray(f.relatedCells) || f.relatedCells.length > ACCESSIBILITY_LIMITS.columns)
      throw refusal('internal-invalid-field', 'Invalid related header locations.');
    for (const c of f.relatedCells) noteMembershipCell(c);
    if (
      new Set(f.relatedCells).size !== f.relatedCells.length ||
      (f.rule === 'duplicate-table-header' && (!f.relatedCells.includes(f.cell) || f.relatedCells.length < 2)) ||
      (f.rule !== 'duplicate-table-header' && f.relatedCells.length)
    )
      throw refusal('internal-invalid-field', 'Invalid related locations.');
    const key = JSON.stringify([f.rule, f.sheet, f.cell, f.objectId, f.tableName]);
    if (seen.has(key)) throw refusal('note-duplicate', 'Duplicate accessibility finding.');
    seen.add(key);
  }
  fields(v.coverage, ['objects', 'tables', 'headerCellsRead', 'skipped'], 'accessibility coverage');
  for (const key of ['objects', 'tables']) {
    const c = v.coverage[key];
    fields(c, ['total', 'checked', 'skipped'], key + ' coverage');
    count(c.checked, 65536);
    count(c.skipped, 65536);
    if (c.total !== null) {
      count(c.total, 65536);
      if (c.checked + c.skipped !== c.total)
        throw refusal('note-bad-input', 'Coverage does not account for every item.');
    }
  }
  count(v.coverage.headerCellsRead, ACCESSIBILITY_LIMITS.headerCells);
  if (!Array.isArray(v.coverage.skipped) || v.coverage.skipped.length > ACCESSIBILITY_LIMITS.skips)
    throw refusal('note-limit', 'Skipped inventory exceeds its bound.');
  for (const s of v.coverage.skipped) {
    fields(s, ['kind', 'sheet', 'name', 'reason'], 'skipped item');
    if (!['objects', 'table'].includes(s.kind)) throw refusal('internal-invalid-field', 'Invalid skipped kind.');
    if (s.sheet !== null) noteMembershipSheet(s.sheet);
    if (s.name !== null) text(s.name, 255, 'skipped name');
    text(s.reason, 1024, 'skip reason');
  }
  const partial =
    v.coverage.skipped.length > 0 ||
    Object.values({ objects: v.coverage.objects, tables: v.coverage.tables }).some(
      (c) => c.total === null || c.skipped > 0
    );
  if (v.status !== (partial ? 'partial' : 'complete'))
    throw refusal('note-unsupported', 'Accessibility status hides incomplete coverage.');
  return v;
}
function receipt(v) {
  fields(v, ['documentId', 'ownerId', 'state', 'revision'], 'accessibility receipt');
  for (const k of Object.keys(v)) text(v[k], k === 'revision' ? 512 : 128, k);
  return v;
}
export function bindAccessibility(v, binding) {
  accessibilityReport(v);
  receipt(binding);
  return { ...structuredClone(v), ...binding };
}
export function validateAccessibility(v, binding) {
  receipt(binding);
  const keys = ['schema', 'scope', 'status', 'rules', 'findings', 'coverage'];
  fields(v, [...keys, ...Object.keys(binding)], 'bound accessibility report');
  for (const k of Object.keys(binding))
    if (v[k] !== binding[k]) throw refusal('note-stale', 'Accessibility findings belong to a retired owner/head.');
  accessibilityReport(Object.fromEntries(keys.map((k) => [k, v[k]])));
  return v;
}
