/** Passive package admission only. Capabilities come from the OPC graph, not
 * the external extension or conventional ActiveX/signature filenames.
 *
 * Macros do not make a workbook read-only. The macro part (vbaProject.bin), its
 * signature parts and the macro-enabled content type are carried byte for byte,
 * as Excel does when it edits a macro workbook. Controls and package signatures
 * stay guarded: an edit would have to move the controls or would break the
 * signature. */
import { DOMParser } from '@xmldom/xmldom';
import { decodeXmlUtf8 } from './zip-parts.mjs';
import { refusal } from './refusal-codes.mjs';
const TYPES = 'http://schemas.openxmlformats.org/package/2006/content-types';
const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const guardedType = (type) =>
  /^(?:application\/vnd\.ms-office\.activex|application\/vnd\.openxmlformats-package\.digital-signature-)/i.test(type);
const guardedRelation = (type) =>
  /\/(?:activexcontrolbinary|control)$/i.test(type) || /\/digital-signature\//i.test(type);
function document(bytes, namespace, root) {
  const source = decodeXmlUtf8(bytes);
  if (/<!DOCTYPE|<!ENTITY/i.test(source)) throw refusal('package-malformed', 'Unsupported package entities');
  const fail = (message) => {
      throw refusal('package-malformed', message);
    },
    xml = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(
      source,
      'application/xml'
    ),
    element = xml.documentElement;
  if (!element || element.namespaceURI !== namespace || element.localName !== root)
    throw refusal('package-malformed', 'Invalid package graph namespace/root');
  return element;
}
function children(element) {
  return Array.from(element.childNodes).filter((node) => {
    if ([3, 4].includes(node.nodeType) && node.data.trim())
      throw refusal('internal-invalid-field', 'Invalid package graph mixed text');
    return node.nodeType === 1;
  });
}

/** True when the package carries a macro project (VBA). The editor never changes
 * it, so a change that would need it to change is refused (see `macro-protected`). */
export const hasMacroProject = (parts) => Object.keys(parts).some((name) => /(?:^|\/)vbaProject\.bin$/i.test(name));

export function packageReadOnlyReasons(parts) {
  const reasons = [],
    guard = () => reasons.push('ActiveX controls and signed workbooks are read-only');
  if (Object.keys(parts).some((name) => /activeX|_xmlsignatures/i.test(name))) guard();
  try {
    if (!parts['[Content_Types].xml']) throw refusal('internal-missing-field', 'Missing package content types');
    const root = document(parts['[Content_Types].xml'], TYPES, 'Types'),
      seen = new Set();
    for (const child of children(root)) {
      if (child.namespaceURI !== TYPES || !['Default', 'Override'].includes(child.localName) || children(child).length)
        throw refusal('package-malformed', 'Invalid package content type entry');
      const key = child.localName + ':' + child.getAttribute(child.localName === 'Default' ? 'Extension' : 'PartName'),
        type = child.getAttribute('ContentType');
      if (
        !type ||
        type.trim() !== type ||
        !/^[^\s;/]+\/[^\s;]+(?:;[^\r\n]*)?$/.test(type) ||
        key.endsWith(':') ||
        seen.has(key)
      )
        throw refusal('internal-missing-field', 'Missing, malformed or duplicate package content type');
      seen.add(key);
      if (guardedType(type)) guard();
    }
  } catch {
    reasons.push('The package content type graph is invalid; editing is not qualified');
  }
  for (const [name, bytes] of Object.entries(parts))
    if (name.endsWith('.rels')) {
      try {
        const root = document(bytes, RELS, 'Relationships'),
          seen = new Set();
        for (const child of children(root)) {
          if (child.namespaceURI !== RELS || child.localName !== 'Relationship' || children(child).length)
            throw refusal('package-malformed', 'Invalid package relationship entry');
          const id = child.getAttribute('Id'),
            type = child.getAttribute('Type'),
            target = child.getAttribute('Target');
          if (!id || !type || !/^[A-Za-z][A-Za-z0-9+.-]*:[^\s]+$/.test(type) || !target || seen.has(id))
            throw refusal('internal-missing-field', 'Missing, malformed or duplicate package relationship');
          seen.add(id);
          // Guard even a missing, external or malformed sensitive target. This
          // reader never interprets or executes the target and cannot edit it.
          if (guardedRelation(type)) guard();
        }
      } catch {
        reasons.push('The package relationship graph is invalid; editing is not qualified');
      }
    }
  return [...new Set(reasons)];
}
