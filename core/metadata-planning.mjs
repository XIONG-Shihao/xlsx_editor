// Fixed metadata planners run in a non-Node VM. Rust supplies cell-free XML,
// owned header values and commands; these results never publish a document.
import { zipSync, strToU8 } from 'fflate';
import { DOMParser } from '@xmldom/xmldom';
import { readWorkbookMetadata, indexedPackageEdit } from './workbook-package.mjs';
import { indexedRulesDefinition, indexedRulesPlan } from './indexed-rules.mjs';
import { refusal } from './refusal-codes.mjs';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export function metadataPackage(parts) {
  if (!parts || typeof parts !== 'object' || Array.isArray(parts) || Object.keys(parts).length > 4096)
    throw refusal('internal-invalid-field', 'validation: Invalid metadata parts');
  let bytes = 0;
  const encoded = Object.create(null);
  for (const [name, xml] of Object.entries(parts)) {
    if (
      !name ||
      name.length > 4096 ||
      name.startsWith('/') ||
      name.includes('\\') ||
      name.split('/').some((part) => !part || part === '.' || part === '..') ||
      typeof xml !== 'string' ||
      /<!DOCTYPE|<!ENTITY/i.test(xml)
    )
      throw refusal('package-unsupported', 'validation: Unqualified metadata part');
    const data = strToU8(xml);
    bytes += data.length;
    if (data.length > 16 * 1024 * 1024 || bytes > 32 * 1024 * 1024)
      throw refusal('workbook-limit', 'validation: Metadata exceeds its byte budget');
    const fail = (message) => {
      throw refusal('package-malformed', 'validation: Invalid metadata XML: ' + message);
    };
    const doc = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(
      xml,
      'application/xml'
    );
    if (
      doc.documentElement.namespaceURI === NS &&
      doc.documentElement.localName === 'worksheet' &&
      doc.getElementsByTagNameNS(NS, 'c').length
    )
      throw refusal('workbook-unsupported', 'validation: Metadata planning cannot receive worksheet cells');
    encoded[name] = data;
  }
  return zipSync(encoded, { mtime: new Date(2000, 0, 1) });
}
function packagePlan(parts, command, headers = []) {
  const fields = {
    addSheet: ['operation', 'sheet'],
    renameSheet: ['operation', 'sheet', 'newName'],
    duplicateSheet: ['operation', 'sheet', 'newName'],
    createTable: ['operation', 'sheet', 'range', 'tableName'],
    addConditionalRule: ['operation', 'sheet', 'range', 'rule', 'threshold', 'fillColor'],
  };
  const allowed = fields[command?.operation];
  if (
    !allowed ||
    Object.keys(command).some((field) => !allowed.includes(field)) ||
    !Array.isArray(headers) ||
    headers.length > 16384
  )
    throw refusal('workbook-unsupported', 'validation: Unqualified metadata command');
  if (headers.length && command.operation !== 'createTable')
    throw refusal('table-bad-input', 'validation: Only table planning accepts header inputs');
  for (const header of headers)
    if (
      !header ||
      Object.keys(header).some((key) => !['cell', 'value'].includes(key)) ||
      typeof header.cell !== 'string' ||
      typeof header.value !== 'string'
    )
      throw refusal('internal-invalid-field', 'validation: Invalid table header');
  const bytes = metadataPackage(parts),
    { view } = readWorkbookMetadata(bytes);
  return indexedPackageEdit(
    bytes,
    command,
    view,
    headers.map((header) => ({ ...header, sheet: command.sheet }))
  );
}
function rulesDefinition(parts, sheet, kind) {
  if (!['conditional', 'validation'].includes(kind))
    throw refusal('internal-invalid-field', 'validation: Invalid rule kind');
  const bytes = metadataPackage(parts),
    { view } = readWorkbookMetadata(bytes);
  return indexedRulesDefinition(bytes, view, sheet, kind);
}
function rulesPlan(parts, command, stylesXml) {
  if (
    stylesXml !== undefined &&
    (typeof stylesXml !== 'string' ||
      strToU8(stylesXml).length > 16 * 1024 * 1024 ||
      /<!DOCTYPE|<!ENTITY/i.test(stylesXml))
  )
    throw refusal('style-unsupported', 'validation: Unqualified publication styles');
  if (!['setConditionalRules', 'setValidationRules'].includes(command?.operation))
    throw refusal('internal-invalid-field', 'validation: Invalid rule collection command');
  const bytes = metadataPackage(parts),
    { view } = readWorkbookMetadata(bytes);
  return indexedRulesPlan(bytes, view, command, { stylesXml });
}

function qualified(action) {
  try {
    return action();
  } catch (error) {
    const message = String(error?.message ?? error);
    if (/^(?:validation|conflict|timeout|missing_item|unavailable):/.test(message)) throw error;
    throw refusal('workbook-refused', (/^Conflict:/i.test(message) ? 'conflict: ' : 'validation: ') + message);
  }
}
export function planPackageMetadata(...args) {
  return qualified(() => packagePlan(...args));
}
export function readMetadataRules(...args) {
  return qualified(() => rulesDefinition(...args));
}
export function planMetadataRules(...args) {
  return qualified(() => rulesPlan(...args));
}
