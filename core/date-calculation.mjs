/** Numeric scalar calendar slice qualified by the paired Excel 16.113.3 oracles.
 * Pure integer civil-day arithmetic; no clock, timezone or text-date parsing. */
export const QUALIFIED_DATE_FUNCTIONS = new Set(['DATE', 'YEAR', 'MONTH', 'DAY', 'EDATE', 'EOMONTH']);
const floor = Math.floor;
function civilDays(year, month, day) {
  year -= month <= 2 ? 1 : 0;
  const era = floor(year / 400),
    yoe = year - era * 400,
    mp = month + (month > 2 ? -3 : 9);
  return era * 146097 + yoe * 365 + floor(yoe / 4) - floor(yoe / 100) + floor((153 * mp + 2) / 5) + day - 1 - 719468;
}
function civilParts(days) {
  const z = days + 719468,
    era = floor(z / 146097),
    doe = z - era * 146097;
  const yoe = floor((doe - floor(doe / 1460) + floor(doe / 36524) - floor(doe / 146096)) / 365);
  const year = yoe + era * 400,
    doy = doe - (365 * yoe + floor(yoe / 4) - floor(yoe / 100)),
    mp = floor((5 * doy + 2) / 153);
  const day = doy - floor((153 * mp + 2) / 5) + 1,
    month = mp + (mp < 10 ? 3 : -9);
  return [year + (month <= 2 ? 1 : 0), month, day];
}
const origin1900 = civilDays(1899, 12, 31),
  origin1904 = civilDays(1904, 1, 1),
  march1900 = civilDays(1900, 3, 1);
const maximum1900 = 2958465,
  maximum1904 = 2957003;
export function calculateDateScalar(name, values, date1904 = false) {
  if (typeof date1904 !== 'boolean') return { unresolved: 'Invalid workbook date-system metadata' };
  if (
    !QUALIFIED_DATE_FUNCTIONS.has(name) ||
    values.length !== (name === 'DATE' ? 3 : ['EDATE', 'EOMONTH'].includes(name) ? 2 : 1)
  )
    return { unresolved: 'Unqualified date-function argument count' };
  if (values.some((value) => typeof value !== 'number' || !Number.isFinite(value)))
    return { unresolved: 'Date functions require finite numeric scalar inputs' };
  const maximum = date1904 ? maximum1904 : maximum1900;
  if (name === 'DATE') {
    let [year, month, day] = values.map(floor);
    if (year < 0 || year > 9999) return { error: '#NUM!' };
    if (Math.abs(month) > 1200 || Math.abs(day) > 36600)
      return { unresolved: 'Date normalization exceeds the qualified month/day bounds' };
    if (year <= 1899) year += 1900;
    const shift = floor((month - 1) / 12);
    year += shift;
    month -= shift * 12;
    const first = civilDays(year, month, 1);
    const serial = first - (date1904 ? origin1904 : origin1900) + (!date1904 && first >= march1900 ? 1 : 0) + day - 1;
    return serial < 0 || serial > maximum ? { error: '#NUM!' } : { value: serial };
  }
  const serial = values[0];
  if (serial < 0 || serial >= maximum + 1) return { error: '#NUM!' };
  const whole = floor(serial);
  const parts =
    !date1904 && whole === 0
      ? [1900, 1, 0]
      : !date1904 && whole === 60
        ? [1900, 2, 29]
        : civilParts((date1904 ? origin1904 : origin1900) + whole - (!date1904 && whole > 60 ? 1 : 0));
  if (name === 'EDATE' || name === 'EOMONTH') {
    const months = Math.trunc(values[1]);
    if (Math.abs(months) > 120000) return { unresolved: 'Month shift exceeds the qualified bound' };
    if (!date1904 && whole === 0 && months === -1) return { error: '#NUM!' };
    if (
      (!date1904 &&
        whole === 0 &&
        !((name === 'EDATE' && [0, 1].includes(months)) || (name === 'EOMONTH' && months === 0))) ||
      (!date1904 && whole === 60 && !((name === 'EDATE' && months === 1) || (name === 'EOMONTH' && months === 0)))
    )
      return { unresolved: 'Uncaptured fictional/origin month-date mode' };
    let [year, month, day] = parts;
    const shift = Math.floor((month - 1 + months) / 12);
    year += shift;
    month = month - 1 + months - shift * 12 + 1;
    if (year < 1 || year > 9999) return { error: '#NUM!' };
    const last = civilDays(year, month + 1, 1) - civilDays(year, month, 1);
    day = name === 'EOMONTH' ? last : Math.min(day, last);
    const target = civilDays(year, month, day),
      result = target - (date1904 ? origin1904 : origin1900) + (!date1904 && target >= march1900 ? 1 : 0);
    if (result < 0 && !(date1904 && whole === 0 && months === -1))
      return { unresolved: 'Uncaptured month-date result before calendar origin' };
    return result > maximum ? { error: '#NUM!' } : { value: result };
  }
  return { value: parts[name === 'YEAR' ? 0 : name === 'MONTH' ? 1 : 2] };
}
