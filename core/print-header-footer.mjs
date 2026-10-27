/** Plain-text page furniture; inline font/image/date/page fields stay protected. */
import { refusal } from './refusal-codes.mjs';

export const PRINT_HEADER_FOOTER_LIMIT = 255;
const sides = ['left', 'center', 'right'],
  codes = { left: 'L', center: 'C', right: 'R' };
export function emptyPrintHeaderFooter() {
  return {
    header: { left: '', center: '', right: '' },
    footer: { left: '', center: '', right: '' },
    scaleWithDoc: true,
  };
}
function exact(value, keys) {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.keys(value).length !== keys.length ||
    Object.keys(value).some((key) => !keys.includes(key))
  )
    throw refusal('internal-invalid-field', 'Invalid plain-text header/footer');
}
function text(value) {
  if (typeof value !== 'string' || /[\u0000-\u001f\u007f\ufffe\uffff]/u.test(value) || /\p{Surrogate}/u.test(value))
    throw refusal('print-bad-input', 'Use single-line XML-safe header/footer text');
  return value;
}
export function encodePrintSections(sections) {
  exact(sections, sides);
  const encoded = sides
    .map((side) => (sections[side] ? '&' + codes[side] + text(sections[side]).replaceAll('&', '&&') : ''))
    .join('');
  for (const side of sides) text(sections[side]);
  if (encoded.length > PRINT_HEADER_FOOTER_LIMIT)
    throw refusal('print-limit', 'Header/footer exceeds 255 encoded characters');
  return encoded;
}
export function decodePrintSections(encoded) {
  if (
    typeof encoded !== 'string' ||
    encoded.length > PRINT_HEADER_FOOTER_LIMIT ||
    (encoded && !/^&[LCR]/.test(encoded))
  )
    throw refusal('print-unsupported', 'Unqualified header/footer sections');
  const sections = { left: '', center: '', right: '' };
  let side = 'left';
  for (let at = 0; at < encoded.length; at++) {
    const ch = encoded[at];
    if (ch !== '&') {
      sections[side] += ch;
      continue;
    }
    const code = encoded[++at];
    if (code === '&') sections[side] += '&';
    else if (['L', 'C', 'R'].includes(code)) side = { L: 'left', C: 'center', R: 'right' }[code];
    else throw refusal('print-protected', 'Formatted, dynamic or image headers/footers are protected');
  }
  for (const side of sides) text(sections[side]);
  return sections;
}
export function validatePrintHeaderFooter(value) {
  exact(value, ['header', 'footer', 'scaleWithDoc']);
  if (typeof value.scaleWithDoc !== 'boolean') throw refusal('internal-invalid-field', 'Invalid header/footer scaling');
  encodePrintSections(value.header);
  encodePrintSections(value.footer);
  return { header: { ...value.header }, footer: { ...value.footer }, scaleWithDoc: value.scaleWithDoc };
}
