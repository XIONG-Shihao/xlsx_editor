/** Qualified ASCII defined-name authoring; no label conversion or owner writes. */
import { refusal } from './refusal-codes.mjs';

export function validateDefinedName(name) {
  if (
    typeof name !== 'string' ||
    name.length < 1 ||
    name.length > 255 ||
    !/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ||
    /^_xlnm\./i.test(name) ||
    /^(?:R|C|TRUE|FALSE|[A-Za-z]{1,3}\d+|R\d+C\d+)$/i.test(name)
  )
    throw refusal(
      'name-unsupported',
      'Use an ASCII static range name without cell addresses, Boolean literals or reserved names.'
    );
  return name;
}
/** scopeSheet omitted means workbook scope. visibleSheet additionally refuses
 * its local shadow when a workflow is creating a useful workbook-scoped name. */
export function definedNameConflict(name, inventory, { scopeSheet, visibleSheet } = {}) {
  validateDefinedName(name);
  if (
    !inventory ||
    !Array.isArray(inventory.namedRanges) ||
    !Array.isArray(inventory.tables) ||
    (scopeSheet !== undefined && typeof scopeSheet !== 'string') ||
    (visibleSheet !== undefined && typeof visibleSheet !== 'string')
  )
    throw refusal('internal-invalid-field', 'Invalid defined-name inventory or scope.');
  const fold = (value) => value.replace(/[a-z]/g, (char) => char.toUpperCase()),
    key = fold(name);
  if (inventory.tables.some((table) => typeof table.name !== 'string'))
    throw refusal('internal-invalid-field', 'Invalid defined-name table inventory.');
  if (inventory.tables.some((table) => fold(table.name) === key)) return 'A table already uses that name.';
  for (const existing of inventory.namedRanges) {
    if (typeof existing.name !== 'string') throw refusal('internal-invalid-field', 'Invalid defined-name inventory.');
    if (fold(existing.name) !== key) continue;
    if (existing.scopeInvalid || (Object.hasOwn(existing, 'sheet') && typeof existing.sheet !== 'string'))
      return 'A name with an invalid scope already uses that spelling.';
    if (existing.sheet === scopeSheet) return 'A name already exists in that scope.';
    if (scopeSheet === undefined && visibleSheet !== undefined && existing.sheet === visibleSheet)
      return 'A worksheet-local name would shadow this workbook name.';
  }
  return null;
}
