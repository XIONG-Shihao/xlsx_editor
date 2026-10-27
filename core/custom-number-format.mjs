/** Documentation-based custom numeric display. Unsupported profiles return
 * null; source values, type, style code and calculation authority never change.
 * Bounds below are editor resource/runtime bounds, not Excel format limits. */
export const CUSTOM_NUMBER_FORMAT_LIMITS = Object.freeze({ codeUnits: 4096, outputUnits: 65536, fractionPlaces: 100 });
const COLORS = Object.freeze({
  black: '000000',
  blue: '0000FF',
  cyan: '00FFFF',
  green: '00FF00',
  magenta: 'FF00FF',
  red: 'FF0000',
  white: 'FFFFFF',
  yellow: 'FFFF00',
});
const literalCharacter = (char) => "$+-/(): !^&'~{}<>=".includes(char);
const glyph = (code, index) => String.fromCodePoint(code.codePointAt(index));
const raw = (tokens) => tokens.map((token) => (token.kind === 'raw' ? token.text : '')).join('');
const output = (text, kind, color) =>
  text.length > CUSTOM_NUMBER_FORMAT_LIMITS.outputUnits
    ? null
    : { text, ...(color ? { color } : {}), kind, supported: true };
function parse(code) {
  const sections = [{ tokens: [], color: undefined, condition: undefined }];
  let section = sections[0];
  for (let at = 0; at < code.length;) {
    const char = glyph(code, at);
    at += char.length;
    if (char === ';') {
      if (sections.length === 4) return null;
      section = { tokens: [], color: undefined, condition: undefined };
      sections.push(section);
      continue;
    }
    if (char === '"') {
      let text = '',
        closed = false;
      while (at < code.length) {
        const next = glyph(code, at);
        at += next.length;
        if (next === '"') {
          closed = true;
          break;
        }
        text += next;
      }
      if (!closed) return null;
      section.tokens.push({ kind: 'literal', text });
      continue;
    }
    if (char === '\\') {
      if (at === code.length) return null;
      const text = glyph(code, at);
      at += text.length;
      section.tokens.push({ kind: 'literal', text });
      continue;
    }
    if (char === '_' || char === '*') {
      if (at === code.length) return null;
      const text = glyph(code, at);
      at += text.length;
      section.tokens.push({ kind: 'unsupported', text });
      continue;
    }
    if (char === '[') {
      const end = code.indexOf(']', at);
      if (end < 0) return null;
      const directive = code.slice(at, end);
      if (directive.includes('[')) return null;
      at = end + 1;
      const color = COLORS[directive.toLowerCase()];
      if (color) {
        if (section.tokens.length || section.color) return null;
        section.color = color;
        continue;
      }
      const condition = /^(<=|>=|<>|=|<|>)([+-]?(?:\d+(?:\.\d*)?|\.\d+))$/.exec(directive);
      if (condition) {
        if (section.tokens.length || section.condition) return null;
        const value = Number(condition[2]);
        if (!Number.isFinite(value)) return null;
        section.condition = { operator: condition[1], value };
        continue;
      }
      const currency = /^\$([^\[\]-]*)-(409|0409)$/i.exec(directive);
      if (currency) {
        if (currency[1]) section.tokens.push({ kind: 'literal', text: currency[1] });
        continue;
      }
      section.tokens.push({ kind: 'unsupported', text: directive });
      continue;
    }
    if (char === ']' || /[\u0000-\u0008\u000b\u000c\u000e-\u001f\uD800-\uDFFF\uFFFE\uFFFF]/u.test(char)) return null;
    section.tokens.push({
      kind: char === '@' ? 'text' : char === '%' ? 'percent' : /[0#.,]/.test(char) ? 'numeric' : 'raw',
      text: char,
    });
  }
  return sections;
}
function matches(value, condition) {
  switch (condition.operator) {
    case '<':
      return value < condition.value;
    case '>':
      return value > condition.value;
    case '<=':
      return value <= condition.value;
    case '>=':
      return value >= condition.value;
    case '=':
      return value === condition.value;
    case '<>':
      return value !== condition.value;
    default:
      return false;
  }
}
function selectNumber(sections, value) {
  const numeric = sections.slice(0, Math.min(3, sections.length));
  if (numeric.some((section) => section.condition)) {
    if (!numeric[0].condition || numeric.some((section, index) => index > 1 && section.condition)) return null;
    let fallback = false;
    for (const section of numeric) {
      if (section.condition) {
        if (fallback) return null;
        if (matches(value, section.condition)) return { section, value, conditional: true };
      } else {
        fallback = true;
        return { section, value, conditional: true };
      }
    }
    return null; // A width-dependent hash fill for unmatched criteria is unqualified.
  }
  const index = value < 0 && numeric.length > 1 ? 1 : value === 0 && numeric.length > 2 ? 2 : 0;
  return { section: numeric[index], value: index === 1 ? Math.abs(value) : value, conditional: false };
}
function literals(tokens, value, allowText = false) {
  const parts = [];
  let units = 0;
  for (const token of tokens) {
    let part;
    if (token.kind === 'literal' || (token.kind === 'raw' && literalCharacter(token.text))) part = token.text;
    else if (token.kind === 'percent') part = '%';
    else if (token.kind === 'text' && allowText) part = value;
    else return null;
    // Check before retaining/concatenating each substitution. A short format can
    // repeat a full cell value thousands of times; a final length check is late.
    if (part.length > CUSTOM_NUMBER_FORMAT_LIMITS.outputUnits - units) return null;
    units += part.length;
    parts.push(part);
  }
  return parts.join('');
}
function isoDate(tokens, value, date1904) {
  if (value < 0 || tokens.length !== 10) return null;
  const pattern = tokens
    .map((token) =>
      token.kind === 'raw' ? token.text.toLowerCase() : token.kind === 'literal' && token.text === '-' ? '-' : '?'
    )
    .join('');
  if (pattern !== 'yyyy-mm-dd') return null;
  if (!date1904 && Math.floor(value) === 60) return '1900-02-29';
  const days = date1904 ? value : value < 60 ? value + 1 : value;
  const date = new Date((date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30)) + days * 86400000);
  return Number.isFinite(date.getTime()) ? date.toISOString().slice(0, 10) : null;
}
function compileNumber(tokens) {
  const first = tokens.findIndex((token) => token.kind === 'numeric' && /[0#.]/.test(token.text));
  if (first < 0) return null;
  let last = -1;
  tokens.forEach((token, index) => {
    if (token.kind === 'numeric') last = index;
  });
  const core = tokens.slice(first, last + 1);
  if (core.some((token) => token.kind !== 'numeric')) return null;
  const mask = core.map((token) => token.text).join(''),
    pieces = mask.split('.');
  if (pieces.length > 2) return null;
  let integer = pieces[0],
    fraction = pieces[1] ?? '',
    scale = 0;
  const integerScale = integer.match(/,+$/)?.[0].length ?? 0,
    fractionScale = pieces.length === 2 ? (fraction.match(/,+$/)?.[0].length ?? 0) : 0;
  if (integerScale && fractionScale) return null;
  if (integerScale) {
    integer = integer.slice(0, -integerScale);
    scale = integerScale;
  }
  if (fractionScale) {
    fraction = fraction.slice(0, -fractionScale);
    scale = fractionScale;
  }
  const grouping = integer.includes(',');
  if (grouping && integer.split(',').some((part) => !part || !/^[0#]+$/.test(part))) return null;
  integer = integer.replaceAll(',', '');
  if (
    !/^#*0*$/.test(integer) ||
    !/^[0#]*$/.test(fraction) ||
    (!integer && !fraction) ||
    fraction.length > CUSTOM_NUMBER_FORMAT_LIMITS.fractionPlaces
  )
    return null;
  const prefix = literals(tokens.slice(0, first), '');
  const suffix = literals(tokens.slice(last + 1), '');
  if (prefix === null || suffix === null) return null;
  const percent = tokens.filter((token) => token.kind === 'percent').length;
  if (percent > 1) return null;
  return { integer, fraction, decimal: pieces.length === 2, scale, grouping, percent, prefix, suffix };
}
/** Pure primitive formatter. Strings are typed text, never parsed as numbers.
 * Raw numeric/error/calculation fences belong to the caller's cell contract. */
export function formatCustomNumber(value, code, date1904 = false) {
  if (
    typeof code !== 'string' ||
    !code ||
    code.length > CUSTOM_NUMBER_FORMAT_LIMITS.codeUnits ||
    /[\u0000-\u0008\u000b\u000c\u000e-\u001f\uD800-\uDFFF\uFFFE\uFFFF]/u.test(code) ||
    typeof date1904 !== 'boolean' ||
    !['number', 'string'].includes(typeof value) ||
    (typeof value === 'number' && !Number.isFinite(value))
  )
    return null;
  const sections = parse(code);
  if (!sections) return null;
  if (typeof value === 'string') {
    const section =
      sections.length === 4
        ? sections[3]
        : sections.length === 1 && sections[0].tokens.some((token) => token.kind === 'text')
          ? sections[0]
          : null;
    if (!section) return output(value, 'text');
    if (section.condition) return null;
    const text = literals(section.tokens, value, true);
    return text === null ? null : output(text, 'text', section.color);
  }
  const selected = selectNumber(sections, value);
  if (!selected) return null;
  const { section } = selected,
    tokens = section.tokens;
  if (tokens.some((token) => token.kind === 'unsupported')) return null;
  const general = tokens.every((token) => token.kind === 'raw') && raw(tokens).toLowerCase() === 'general';
  if (general || (tokens.length === 1 && tokens[0].kind === 'text'))
    return output(String(value), 'number', section.color);
  const date = isoDate(tokens, selected.value, date1904);
  if (date !== null) return output(date, 'date', section.color);
  const literal = literals(tokens, '');
  if (literal !== null && !tokens.some((token) => token.kind === 'percent'))
    return output(literal, 'literal', section.color);
  const model = compileNumber(tokens);
  if (!model || (selected.conditional && selected.value < 0)) return null;
  const scaled = (Math.abs(selected.value) * Math.pow(100, model.percent)) / Math.pow(1000, model.scale);
  if (!Number.isFinite(scaled)) return null;
  let rounded;
  try {
    rounded = model.grouping
      ? new Intl.NumberFormat('en-US', {
          useGrouping: false,
          minimumFractionDigits: model.fraction.length,
          maximumFractionDigits: model.fraction.length,
        }).format(scaled)
      : scaled.toFixed(model.fraction.length);
  } catch {
    return null;
  }
  if (!/^\d+(?:\.\d+)?$/.test(rounded)) return null;
  let [integer, fraction = ''] = rounded.split('.');
  const required = (model.integer.match(/0/g) ?? []).length;
  integer = integer === '0' && required === 0 ? '' : integer.padStart(required, '0');
  for (
    let index = model.fraction.length - 1;
    index >= 0 && model.fraction[index] === '#' && fraction.endsWith('0');
    index--
  )
    fraction = fraction.slice(0, -1);
  if (model.grouping) integer = integer.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const decimal = model.decimal && (fraction.length > 0 || model.fraction.length === 0) ? '.' : '';
  const sign = selected.value < 0 ? '-' : '';
  return output(sign + model.prefix + integer + decimal + fraction + model.suffix, 'number', section.color);
}
