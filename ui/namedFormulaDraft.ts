/** Name insertion changes editor text only; owner acceptance belongs to Enter. */
export interface FormulaDraftSnapshot {
  text: string;
  start: number;
  end: number;
  generation: number;
  composing: boolean;
}
export interface FormulaDraftLease {
  isCurrent: () => boolean;
  accept: (text: string) => Promise<void>;
  error: (error: unknown) => void;
}
export interface FormulaDraftHandle {
  capture: () => FormulaDraftSnapshot;
  park: () => void;
  attach: (lease: FormulaDraftLease) => void;
  replace: (expected: FormulaDraftSnapshot, text: string, caret: number) => void;
  focus: () => void;
  ownsFocus: () => boolean;
  /** A same-scope draft may already be parked after ribbon/menu focus moved. */
  isParked?: () => boolean;
}
export interface NamedFormulaChoice {
  id: string;
  name: string;
  formula: string;
  sheet?: string;
  scope: string;
  disabledReason?: string;
}
type SourceName = { name: string; formula: string; sheet?: string; scopeInvalid?: boolean };
const namePattern = /^[A-Za-z_][A-Za-z0-9_.]*$/;
const asciiFold = (text: string) => text.replace(/[A-Z]/g, (letter) => letter.toLowerCase());
function cell(reference: string): [number, number] | null {
  const match = /^\$([A-Z]{1,3})\$([1-9]\d*)$/.exec(reference);
  if (!match) return null;
  let col = 0;
  for (const letter of match[1]) col = col * 26 + letter.charCodeAt(0) - 64;
  const row = Number(match[2]);
  return col <= 16384 && row <= 1048576 ? [row, col] : null;
}
export function namedFormulaChoices(
  names: readonly SourceName[],
  currentSheet: string,
  sheets: readonly string[]
): readonly NamedFormulaChoice[] {
  if (names.length > 1000) throw Error('This workbook has too many names for this picker.');
  const local = new Map<string, number>(),
    global = new Map<string, number>(),
    same = (a: string, b: string) => asciiFold(a) === asciiFold(b);
  for (const item of names) {
    if (
      item.scopeInvalid ||
      !namePattern.test(item.name) ||
      (Object.hasOwn(item, 'sheet') && typeof item.sheet !== 'string')
    )
      continue;
    const key = asciiFold(item.name),
      map = item.sheet && same(item.sheet, currentSheet) ? local : !item.sheet ? global : null;
    if (map) map.set(key, (map.get(key) ?? 0) + 1);
  }
  return Object.freeze(
    names.map((item, index) => {
      const key = asciiFold(item.name);
      let reason: string | undefined;
      if (item.scopeInvalid || (Object.hasOwn(item, 'sheet') && typeof item.sheet !== 'string'))
        reason = 'This name has an invalid worksheet scope.';
      else if (
        !namePattern.test(item.name) ||
        item.name.length > 255 ||
        /^_xlnm\./i.test(item.name) ||
        /^(?:TRUE|FALSE|R|C|[A-Za-z]{1,3}\d+|R\d+C\d+)$/i.test(item.name)
      )
        reason = 'This name is reserved or unsupported.';
      else if (item.sheet && !same(item.sheet, currentSheet))
        reason = 'Switch to ' + item.sheet + ' to use this worksheet-local name.';
      else if (item.sheet ? (local.get(key) ?? 0) > 1 : (global.get(key) ?? 0) > 1)
        reason = 'This name is ambiguous in its scope.';
      else if (!item.sheet && local.has(key))
        reason = 'The ' + currentSheet + ' worksheet-local name takes precedence.';
      else {
        const qualified =
            /^(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_ .]*))!(\$[A-Z]{1,3}\$[1-9]\d*(?::\$[A-Z]{1,3}\$[1-9]\d*)?)$/.exec(
              item.formula
            ),
          plain = /^\$[A-Z]{1,3}\$[1-9]\d*(?::\$[A-Z]{1,3}\$[1-9]\d*)?$/.test(item.formula);
        const target = qualified ? (qualified[1]?.replaceAll("''", "'") ?? qualified[2]) : item.sheet;
        const range = qualified?.[3] ?? (plain && item.sheet ? item.formula : '');
        const [a, b = a] = range.split(':'),
          first = cell(a),
          last = cell(b);
        if (
          !target ||
          !sheets.some((sheet) => same(sheet, target)) ||
          !first ||
          !last ||
          first[0] > last[0] ||
          first[1] > last[1]
        )
          reason = 'Only static cell and rectangular range names are available here.';
      }
      return Object.freeze({
        id: JSON.stringify([item.sheet ?? null, item.name, item.formula, index]),
        name: item.name,
        formula: item.formula,
        ...(item.sheet ? { sheet: item.sheet } : {}),
        scope: item.sheet ?? 'Workbook',
        ...(reason ? { disabledReason: reason } : {}),
      });
    })
  );
}
export function insertFormulaName(
  snapshot: FormulaDraftSnapshot,
  choice: NamedFormulaChoice
): { text: string; caret: number } {
  if (snapshot.composing) throw Error('Finish composing the formula before choosing a name.');
  if (choice.disabledReason) throw Error(choice.disabledReason);
  if (
    !namePattern.test(choice.name) ||
    choice.name.length > 255 ||
    /^_xlnm\./i.test(choice.name) ||
    /^(?:TRUE|FALSE|R|C|[A-Za-z]{1,3}\d+|R\d+C\d+)$/i.test(choice.name)
  )
    throw Error('Invalid or reserved name.');
  const { text, start, end } = snapshot;
  if (!Number.isInteger(start) || !Number.isInteger(end) || start < 0 || end < start || end > text.length)
    throw Error('The formula caret changed.');
  if (text !== '' && !text.startsWith('=')) throw Error('Start a formula with = before inserting a name.');
  if (text !== '' && start === 0 && end === 0) throw Error('Place the caret after the formula = sign.');
  // Do not splice a name into a string, quoted worksheet or structured selector.
  let quote = '',
    square = 0;
  for (let i = 0; i < start; i++) {
    const c = text[i];
    if (quote) {
      if (c === quote) {
        if (text[i + 1] === quote) i++;
        else quote = '';
      }
      continue;
    }
    if (c === '"' || c === "'") quote = c;
    else if (c === '[') square++;
    else if (c === ']') square = Math.max(0, square - 1);
  }
  if (quote || square) throw Error('Place the caret outside quoted text and table selectors.');
  if (text.slice(0, start).trimEnd().endsWith('!'))
    throw Error('Select the entire worksheet-qualified reference before inserting this name.');
  const token = /[\p{L}\p{N}_.\\$]/u;
  if ((start > 1 && token.test(text[start - 1])) || (end < text.length && token.test(text[end])))
    throw Error('Select the complete reference or place the caret between formula terms.');
  const prefix = text === '' || start === 0 ? '=' : text.slice(0, start),
    next = prefix + choice.name + text.slice(end);
  if (new TextEncoder().encode(next).byteLength > 32768) throw Error('The formula draft is too long.');
  return { text: next, caret: prefix.length + choice.name.length };
}
