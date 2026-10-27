/** Explicit literal conversion shared by Text to Columns and future typed data.
 * No locale guessing, formula parsing, clock, owner state or workbook writes. */
import { hasExactNumericValue } from './calculation.mjs';
import { calculateDateScalar } from './date-calculation.mjs';

export const TEXT_CONVERSION_POLICIES = Object.freeze(['text', 'number', 'boolean', 'iso-date']);
export const TEXT_CONVERSION_LIMITS = Object.freeze({ textLength: 32767 });
const decimal = /^([+-]?)(?:(\d+)(?:\.(\d*))?|\.(\d+))(?:[eE]([+-]?\d+))?$/;
const invalidXml = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE\uFFFF]/u;
const spelling = (text) => (text.length <= 80 ? text : `${text.slice(0, 64)}… (${text.length} characters)`);
const numberText = (value) => (Object.is(value, -0) ? '-0' : String(value));
// Normalize only decimal spelling before the existing precision predicate. A
// single regexp removes long runs of zeros, avoiding repeated string slicing.
function decimalMeaning(match) {
  const fraction = match[3] ?? match[4] ?? '',
    digits = ((match[2] ?? '') + fraction).replace(/^0+/, '');
  if (!digits) return match[1] === '-' ? '-0' : '0';
  const significant = digits.replace(/0+$/, ''),
    power = Number(match[5] ?? 0) - fraction.length + digits.length - significant.length;
  return `${match[1] === '-' ? '-' : ''}${significant}e${power}`;
}
function civilDate(text) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
  if (!match) return null;
  const [year, month, day] = match.slice(1).map(Number),
    leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  const length = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1];
  return year >= 1 && year <= 9999 && month >= 1 && month <= 12 && day >= 1 && day <= length
    ? [year, month, day]
    : null;
}

/** Empty source text and ragged padding remain explicit empty text for every
 * policy. Only nonempty typed fields undergo conversion. */
export function convertTextField(raw, policy = 'text', options = {}) {
  const empty = typeof raw === 'string' && raw.length === 0,
    padded = options?.padded === true,
    notices = [];
  const reject = (reason) =>
    Object.freeze({ status: 'rejected', reason, notices: Object.freeze([...notices]), empty, padded });
  const converted = (value, kind, extra = {}) =>
    Object.freeze({ status: 'converted', kind, value, notices: Object.freeze([...notices]), empty, padded, ...extra });
  if (!TEXT_CONVERSION_POLICIES.includes(policy)) return reject('Choose Text, Number, Boolean or ISO Date.');
  if (
    !options ||
    typeof options !== 'object' ||
    Array.isArray(options) ||
    Object.keys(options).some((key) => !['date1904', 'padded'].includes(key)) ||
    (options.date1904 !== undefined && typeof options.date1904 !== 'boolean') ||
    (options.padded !== undefined && typeof options.padded !== 'boolean')
  )
    return reject('Invalid text conversion options.');
  if (typeof raw !== 'string') return reject('Conversion requires literal text.');
  if (raw.length > TEXT_CONVERSION_LIMITS.textLength || invalidXml.test(raw))
    return reject('Text contains an unsupported character or exceeds 32,767 characters.');
  if (padded && !empty) return reject('Padding must be an empty field.');
  if (empty) {
    if (padded) notices.push('Padding kept as empty text.');
    else if (policy !== 'text') notices.push('Empty field kept as text.');
    return converted('', 'text');
  }
  if (policy === 'text') return converted(raw, 'text');
  const text = raw.trim();
  if (text !== raw) notices.push('Outer whitespace removed.');
  if (!text) return reject('Whitespace-only fields need Text; they are not an empty numeric, Boolean or date value.');
  if (policy === 'number') {
    const match = decimal.exec(text);
    if (!match) return reject('Number requires an ASCII decimal, with an optional sign and exponent.');
    const value = Number(text);
    if (!Number.isFinite(value)) return reject('Number is outside the finite numeric range.');
    if (Number.isInteger(value) && !Number.isSafeInteger(value))
      return reject('Integer is outside the exact safe-integer range.');
    if (!hasExactNumericValue(value, decimalMeaning(match)))
      return reject('Number would round or underflow and change its decimal meaning.');
    const canonical = numberText(value);
    if (text !== canonical) notices.push(`Numeric spelling changes: ${spelling(text)} → ${canonical}.`);
    return converted(value, 'number');
  }
  if (policy === 'boolean') {
    if (!/^(?:true|false)$/i.test(text)) return reject('Boolean accepts TRUE or FALSE only.');
    const value = /^true$/i.test(text),
      canonical = value ? 'TRUE' : 'FALSE';
    if (text !== canonical) notices.push(`Boolean spelling changes: ${spelling(text)} → ${canonical}.`);
    return converted(value, 'boolean');
  }
  const parts = civilDate(text);
  if (!parts) return reject('ISO Date requires a real civil date in YYYY-MM-DD form.');
  if (typeof options.date1904 !== 'boolean') return reject('The workbook date system is unavailable; choose Text.');
  if (parts[0] < (options.date1904 ? 1904 : 1900))
    return reject(`Date precedes the ${options.date1904 ? '1904' : '1900'} workbook calendar.`);
  const result = calculateDateScalar('DATE', parts, options.date1904);
  if (!Number.isSafeInteger(result.value) || result.value < 0) return reject('Date is outside the workbook calendar.');
  notices.push(`ISO date stored as serial ${result.value} in the ${options.date1904 ? '1904' : '1900'} calendar.`);
  return converted(result.value, 'date', { numberFormat: 'date' });
}
