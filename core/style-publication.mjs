// Pure style planning: the native owner supplies XML and unique signatures.
// This module has no filesystem, process or other Node host dependencies.
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { createWorkbookStyleWriter } from './workbook-styles.mjs';
import { refusal } from './refusal-codes.mjs';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const fields = [
  'bold',
  'italic',
  'underline',
  'fontFamily',
  'fontSize',
  'fontColor',
  'fillColor',
  'horizontalAlign',
  'verticalAlign',
  'wrapText',
  'border',
  'numberFormat',
];
export function resolveWorkbookStyles(stylesXml, signatures) {
  if (
    typeof stylesXml !== 'string' ||
    stylesXml.length > 16 * 1024 * 1024 ||
    !Array.isArray(signatures) ||
    signatures.length > 65536
  )
    throw refusal('style-limit', 'validation: Style publication exceeds its bounds');
  if (/<!DOCTYPE|<!ENTITY/i.test(stylesXml))
    throw refusal('style-malformed', 'validation: Style entities are unsupported');
  const fail = (message) => {
    throw refusal('style-malformed', 'validation: Invalid style XML: ' + message);
  };
  const styles = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(
    stylesXml,
    'application/xml'
  );
  if (styles.documentElement.namespaceURI !== NS || styles.documentElement.localName !== 'styleSheet')
    throw refusal('internal-invalid-field', 'validation: Invalid style document');
  const xfs = Array.from(styles.documentElement.childNodes).filter(
    (node) => node.nodeType === 1 && node.namespaceURI === NS && node.localName === 'cellXfs'
  );
  if (xfs.length !== 1) throw refusal('internal-missing-field', 'validation: Missing or duplicate cell styles');
  const sourceCount = Array.from(xfs[0].childNodes).filter(
    (node) => node.nodeType === 1 && node.namespaceURI === NS && node.localName === 'xf'
  ).length;
  const apply = createWorkbookStyleWriter(styles),
    seen = new Set(),
    styleIds = [];
  for (const { id, format } of signatures) {
    if (
      !Number.isSafeInteger(id) ||
      id <= 0 ||
      seen.has(id) ||
      !format ||
      typeof format !== 'object' ||
      Array.isArray(format)
    )
      throw refusal('internal-invalid-field', 'validation: Invalid style signature');
    seen.add(id);
    if (
      !Number.isSafeInteger(format.styleId) ||
      format.styleId < 0 ||
      format.styleId >= sourceCount ||
      Object.keys(format).some((key) => key !== 'styleId' && !fields.includes(key))
    )
      throw refusal('internal-invalid-field', 'validation: Invalid source style');
    const cell = styles.createElementNS(NS, 'c');
    cell.setAttribute('s', String(format.styleId));
    const edit = Object.fromEntries(fields.filter((key) => format[key] !== undefined).map((key) => [key, format[key]]));
    if (Object.keys(edit).length) apply(cell, edit);
    styleIds.push({ id, styleId: Number(cell.getAttribute('s')) });
  }
  return { styleIds, stylesXml: new XMLSerializer().serializeToString(styles) };
}
