// Portable source-rule interpretation and literal checks. Native owners supply
// bounded XML/edits; no document paths or Node APIs are available here.
import { DOMParser } from '@xmldom/xmldom';
import { readValidationRules, ruleChildren } from './rules-xml.mjs';
import { validateValidationRules, assertValidationEdits } from './rules-model.mjs';
import { rangeBounds, columnName } from './workbook-commands.mjs';
import { refusal } from './refusal-codes.mjs';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
function worksheet(xml) {
  if (typeof xml !== 'string' || xml.length > 16 * 1024 * 1024 || /<!DOCTYPE|<!ENTITY/i.test(xml))
    throw refusal('rules-malformed', 'validation: Invalid validation XML');
  const fail = (message) => {
    throw refusal('rules-malformed', 'validation: Invalid validation XML: ' + message);
  };
  const doc = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(
    xml,
    'application/xml'
  );
  if (doc.documentElement?.namespaceURI !== NS || doc.documentElement?.localName !== 'worksheet')
    throw refusal('internal-invalid-field', 'validation: Invalid validation worksheet');
  return doc;
}
function protectedRanges(doc) {
  try {
    if (ruleChildren(doc.documentElement).some((node) => node.localName === 'extLst')) return null;
    const containers = ruleChildren(doc.documentElement).filter((node) => node.localName === 'dataValidations');
    if (containers.length !== 1 || containers[0].namespaceURI !== NS) return null;
    const nodes = ruleChildren(containers[0]);
    if (!nodes.length || nodes.length > 128) return null;
    const ranges = nodes.flatMap((node) => {
      if (node.namespaceURI !== NS || node.localName !== 'dataValidation')
        throw refusal('rules-protected', 'Unqualified protected range');
      return node
        .getAttribute('sqref')
        .split(/\s+/)
        .map((range) => {
          const b = rangeBounds(range);
          return columnName(b.first.col) + (b.first.row + 1) + ':' + columnName(b.last.col) + (b.last.row + 1);
        });
    });
    return ranges.length <= 128 ? ranges : null;
  } catch {
    return null;
  }
}
export function readSourceValidation(xml) {
  const doc = worksheet(xml);
  try {
    if (ruleChildren(doc.documentElement).some((node) => node.localName === 'extLst'))
      throw refusal('rules-unsupported', 'Extended worksheet rules require qualification');
    return { supported: true, rules: readValidationRules(doc), protectedRanges: null };
  } catch {
    return { supported: false, rules: [], protectedRanges: protectedRanges(doc) };
  }
}
export function checkSourceValidation(rules, edits) {
  try {
    assertValidationEdits(validateValidationRules(rules), edits);
    return null;
  } catch (error) {
    throw refusal('rules-refused', 'validation: ' + String(error?.message ?? error));
  }
}
