/** closed authored syntax, never calculation or owner admission. */
export type FunctionCategory = 'Financial' | 'Logical' | 'Text' | 'DateTime' | 'Lookup' | 'Math' | 'More';
export type FunctionId =
  | 'PMT'
  | 'PV'
  | 'FV'
  | 'SUM'
  | 'AVERAGE'
  | 'COUNT'
  | 'MIN'
  | 'MAX'
  | 'IF'
  | 'IFERROR'
  | 'LEN'
  | 'LEFT'
  | 'TRIM'
  | 'DATE'
  | 'YEAR'
  | 'MONTH'
  | 'DAY'
  | 'ABS'
  | 'ROUND'
  | 'DEC2HEX'
  | 'ROWS'
  | 'COLUMNS';
export type ArgumentKind = 'number' | 'boolean' | 'text' | 'reference';
export type FunctionArgument =
  | { kind: 'number'; value: number }
  | { kind: 'boolean'; value: boolean }
  | { kind: 'text'; value: string }
  | { kind: 'reference'; value: string };
export type NumericRule =
  | 'finite'
  | 'nonnegative-integer'
  | 'round-digits'
  | 'dec2hex-integer'
  | 'date-year'
  | 'date-month'
  | 'date-day'
  | 'date-serial'
  | 'financial-rate'
  | 'financial-periods'
  | 'financial-cash'
  | 'financial-timing';
export interface FunctionArgumentSpec {
  readonly label: string;
  readonly allowedKinds: readonly ArgumentKind[];
  readonly shape: 'scalar' | 'rectangle';
  readonly numeric?: NumericRule;
  readonly bmpText?: boolean;
  readonly optional?: boolean;
}
export interface FunctionDefinition {
  readonly id: FunctionId;
  readonly name: FunctionId;
  readonly category: FunctionCategory;
  readonly summary: string;
  readonly route: 'native-or-compatibility';
  readonly minArgs: number;
  readonly maxArgs: number;
  readonly arguments: readonly FunctionArgumentSpec[];
  readonly repeatLast: boolean;
  readonly limitations: readonly string[];
}
export interface SerializedFunction {
  readonly functionId: FunctionId;
  readonly formula: string;
  readonly displayFormula: string;
  readonly argumentCount: number;
  readonly requiresOwnerValidation: true;
}
export const FUNCTION_LIMITS = Object.freeze({
  arguments: 16,
  referenceCells: 1000,
  formulaBytes: 8192,
  textBytes: 4096,
  maximumRow: 1048576,
  maximumColumn: 16384,
  maximumNumber: Number.MAX_SAFE_INTEGER,
});
export const FUNCTION_CATEGORIES: readonly {
  readonly id: FunctionCategory;
  readonly label: string;
  readonly notice?: string;
}[] = Object.freeze([
  Object.freeze({
    id: 'Financial' as const,
    label: 'Financial',
    notice:
      'Fixed-rate payments, present value and future value. Results can differ from Excel in the last decimal places.',
  }),
  Object.freeze({ id: 'Logical' as const, label: 'Logical' }),
  Object.freeze({ id: 'Text' as const, label: 'Text' }),
  Object.freeze({ id: 'DateTime' as const, label: 'Date & Time' }),
  Object.freeze({ id: 'Lookup' as const, label: 'Lookup & Reference' }),
  Object.freeze({ id: 'Math' as const, label: 'Math & Trig' }),
  Object.freeze({ id: 'More' as const, label: 'More Functions' }),
]);
const numeric = (label: string, rule: NumericRule = 'finite', optional = false): FunctionArgumentSpec =>
  Object.freeze({
    label,
    allowedKinds: Object.freeze(['number', 'reference'] as const),
    shape: 'scalar',
    numeric: rule,
    optional,
  });
const scalar = (label: string): FunctionArgumentSpec =>
  Object.freeze({
    label,
    allowedKinds: Object.freeze(['number', 'boolean', 'text', 'reference'] as const),
    shape: 'scalar',
  });
const text = (label: string): FunctionArgumentSpec =>
  Object.freeze({ label, allowedKinds: Object.freeze(['text', 'reference'] as const), shape: 'scalar', bmpText: true });
const rectangle = (label: string): FunctionArgumentSpec =>
  Object.freeze({ label, allowedKinds: Object.freeze(['reference'] as const), shape: 'rectangle' });
const common = Object.freeze([
  'Syntax is not a result proof: the actual captured owner must validate types, dependencies, calendar, protections and typed calculation status.',
  'No external/3D references, names, formulas-as-arguments, array constants, spills, volatile functions or arbitrary fragments are admitted by this first serializer.',
]);
function entry(
  id: FunctionId,
  category: FunctionCategory,
  summary: string,
  args: readonly FunctionArgumentSpec[],
  minArgs: number,
  maxArgs = minArgs,
  limits: readonly string[] = [],
  repeatLast = false
): FunctionDefinition {
  return Object.freeze({
    id,
    name: id,
    category,
    summary,
    route: 'native-or-compatibility',
    minArgs,
    maxArgs,
    arguments: Object.freeze([...args]),
    repeatLast,
    limitations: Object.freeze([...common, ...limits]),
  });
}
export const FUNCTION_CATALOG: readonly FunctionDefinition[] = Object.freeze([
  ...(['SUM', 'AVERAGE', 'COUNT', 'MIN', 'MAX'] as const).map((id) =>
    entry(
      id,
      'Math',
      {
        SUM: 'Add numeric values.',
        AVERAGE: 'Find the numeric average.',
        COUNT: 'Count numeric values.',
        MIN: 'Find the smallest numeric value.',
        MAX: 'Find the largest numeric value.',
      }[id],
      [
        Object.freeze({
          label: 'Values',
          allowedKinds: Object.freeze(['number', 'reference'] as const),
          shape: 'rectangle' as const,
          numeric: 'finite' as const,
        }),
      ],
      1,
      16,
      ['Reference blank/text/Boolean/error rules belong to the selected function and owner, not this serializer.'],
      true
    )
  ),
  entry(
    'IF',
    'Logical',
    'Choose one scalar result from a condition.',
    [
      Object.freeze({
        label: 'Condition',
        allowedKinds: Object.freeze(['number', 'boolean', 'reference'] as const),
        shape: 'scalar',
        numeric: 'finite',
      }),
      scalar('If true'),
      scalar('If false'),
    ],
    3,
    3,
    [
      'The condition accepts numeric/Boolean scalars. Comparisons and nested calls require a later closed expression model. Unqualified branches cannot hide behind IF.',
    ]
  ),
  entry(
    'IFERROR',
    'Logical',
    'Use a scalar fallback for an actual typed error.',
    [scalar('Value'), scalar('If error')],
    2,
    2,
    ['Unsupported/unknown/protected calculations stay unresolved. Error-looking literal text is not a typed error.']
  ),
  entry('LEN', 'Text', 'Count characters in the qualified BMP text profile.', [text('Text')], 1, 1, [
    'Supplementary Unicode and locale-dependent conversions are outside this first picker profile.',
  ]),
  entry(
    'LEFT',
    'Text',
    'Take characters from the start of qualified BMP text.',
    [text('Text'), numeric('Characters', 'nonnegative-integer', true)],
    1,
    2,
    ['Omitting Characters preserves the function default; an empty argument is not emitted.']
  ),
  entry('TRIM', 'Text', 'Collapse ordinary ASCII spaces in BMP text.', [text('Text')], 1, 1, [
    'Tabs, nonbreaking spaces and general whitespace are not promised to be trimmed.',
  ]),
  entry(
    'DATE',
    'DateTime',
    'Construct a serial date using the workbook calendar.',
    [numeric('Year', 'date-year'), numeric('Month', 'date-month'), numeric('Day', 'date-day')],
    3,
    3,
    [
      'Numeric arguments use the qualified floor/normalization rules. No text date/locale conversion. Owner calendar and fictional/origin boundaries remain authoritative.',
    ]
  ),
  ...(['YEAR', 'MONTH', 'DAY'] as const).map((id) =>
    entry(
      id,
      'DateTime',
      'Extract ' + id.toLowerCase() + ' using the workbook calendar.',
      [numeric('Serial date', 'date-serial')],
      1,
      1,
      [
        'The serializer does not choose the 1900/1904 calendar or interpret references. Its widest serial bound must still be narrowed by the actual owner.',
      ]
    )
  ),
  entry('ABS', 'Math', 'Find the magnitude of a numeric scalar.', [numeric('Number')], 1),
  entry(
    'ROUND',
    'Math',
    'Round a numeric scalar to qualified decimal places.',
    [numeric('Number'), numeric('Decimal places', 'round-digits')],
    2,
    2,
    [
      'This first picker accepts integral decimal places 0..12; negative or wider digits are not claimed as native support.',
    ]
  ),
  entry(
    'DEC2HEX',
    'More',
    'Convert an admitted nonnegative integer to hexadecimal text.',
    [numeric('Number', 'dec2hex-integer')],
    1,
    1,
    [
      'Exactly one numeric scalar integer 0..2147483647; places, negative values and coercions are separate compatibility profiles not enabled here.',
    ]
  ),
  ...(['ROWS', 'COLUMNS'] as const).map((id) =>
    entry(
      id,
      'Lookup',
      'Count ' + id.toLowerCase() + ' in an ordinary rectangle.',
      [rectangle('Reference')],
      1,
      1,
      ['One same-sheet rectangle up to1000 positions; no generated arrays, names or whole-row/column syntax.']
    )
  ),
  ...(['PMT', 'PV', 'FV'] as const).map((id) =>
    entry(
      id,
      'Financial',
      {
        PMT: 'Find a constant payment using a fixed period rate.',
        PV: 'Find the present value of constant payments.',
        FV: 'Find the future value of constant payments.',
      }[id],
      [
        numeric('Period rate', 'financial-rate'),
        numeric('Number of periods', 'financial-periods'),
        numeric(id === 'PMT' ? 'Present value' : 'Payment each period', 'financial-cash'),
        numeric(
          id === 'FV' ? 'Present value (optional; default 0)' : 'Future value (optional; default 0)',
          'financial-cash',
          true
        ),
        numeric('Payment timing (optional; 0 = end, 1 = beginning)', 'financial-timing', true),
      ],
      3,
      5,
      [
        'Use a period rate from 0 to 1 and a whole number of periods from 1 to 1,200. Rate and period count must use the same time unit.',
        'Cash inputs can be up to 1,000,000,000,000 in magnitude. Outgoing cash is negative and incoming cash positive. Optional balance and payment timing default to 0; enter the balance before timing.',
        'Use a single ordinary numeric input cell on this worksheet. Formula cells, array results, scenario cells, text, TRUE/FALSE, empty cells and errors are not supported for these inputs. Values with unsupported precision are refused.',
        'The calculation must stay finite and within supported number bounds. Some high-rate and long-period combinations cannot be calculated yet.',
        'Results can differ from Excel in the last decimal places.',
      ]
    )
  ),
]);
const byId = new Map<string, FunctionDefinition>(FUNCTION_CATALOG.map((definition) => [definition.id, definition]));
const categories = new Set<string>(FUNCTION_CATEGORIES.map((category) => category.id));
export function functionDefinition(id: string): FunctionDefinition {
  const definition = typeof id === 'string' ? byId.get(id) : undefined;
  if (!definition) throw Error('This function has no qualified picker profile.');
  return definition;
}
export function listFunctions(
  options: { category?: FunctionCategory; query?: string } = {}
): readonly FunctionDefinition[] {
  if (
    !options ||
    typeof options !== 'object' ||
    Array.isArray(options) ||
    Object.keys(options).some((key) => !['category', 'query'].includes(key))
  )
    throw Error('Invalid catalog search.');
  if (options.category !== undefined && !categories.has(options.category)) throw Error('Unknown function category.');
  if (options.query !== undefined && (typeof options.query !== 'string' || options.query.length > 100))
    throw Error('Catalog search exceeds its bound.');
  const query = (options.query ?? '').trim().toLowerCase();
  return Object.freeze(
    FUNCTION_CATALOG.filter(
      (definition) =>
        (!options.category || definition.category === options.category) &&
        (!query || (definition.name + ' ' + definition.summary).toLowerCase().includes(query))
    )
  );
}
const bytes = (value: string) => new TextEncoder().encode(value).byteLength;
function plain(value: unknown): asserts value is FunctionArgument {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    ![Object.prototype, null].includes(Object.getPrototypeOf(value))
  )
    throw Error('Arguments must use the closed typed representation.');
  const descriptors = Object.getOwnPropertyDescriptors(value),
    keys = Reflect.ownKeys(descriptors);
  if (
    keys.length !== 2 ||
    !Object.hasOwn(descriptors, 'kind') ||
    !Object.hasOwn(descriptors, 'value') ||
    keys.some((key) => typeof key !== 'string' || !Object.hasOwn(descriptors[key], 'value'))
  )
    throw Error('Unexpected argument fields or accessors.');
}
function finite(value: unknown): asserts value is number {
  if (
    typeof value !== 'number' ||
    !Number.isFinite(value) ||
    Object.is(value, -0) ||
    Math.abs(value) > FUNCTION_LIMITS.maximumNumber
  )
    throw Error('Enter a finite bounded number; negative zero and unsafe magnitudes are not admitted.');
}
function numericMeaning(text: string): string {
  const match = /^([+-]?)(\d*)(?:\.(\d*))?(?:[eE]([+-]?\d+))?$/.exec(text)!;
  let digits = (match[2] + (match[3] ?? '')).replace(/^0+/, '');
  if (!digits) return '0';
  let power = Number(match[4] ?? 0) - (match[3]?.length ?? 0);
  while (digits.endsWith('0')) {
    digits = digits.slice(0, -1);
    power++;
  }
  return (match[1] === '-' ? '-' : '') + digits + 'E' + power;
}
/** Numeric widgets must validate raw text before discarding its spelling. */
export function parseNumberLiteral(raw: string): number {
  if (typeof raw !== 'string' || raw.length > 80) throw Error('Numeric input exceeds its bound.');
  const text = raw.trim();
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(text))
    throw Error('Enter a numeric literal, not a formula, date or formatted number.');
  const value = Number(text);
  finite(value);
  if (numericMeaning(text) !== numericMeaning(String(value)))
    throw Error('Numeric input would lose its stored decimal meaning.');
  return value;
}
function numericRule(value: number, rule: NumericRule): void {
  const whole = Math.floor(value);
  if (
    (rule === 'nonnegative-integer' && (!Number.isSafeInteger(value) || value < 0)) ||
    (rule === 'round-digits' && (!Number.isInteger(value) || value < 0 || value > 12)) ||
    (rule === 'dec2hex-integer' && (!Number.isInteger(value) || value < 0 || value > 2147483647)) ||
    (rule === 'date-year' && (whole < 0 || whole > 9999)) ||
    (rule === 'date-month' && Math.abs(whole) > 1200) ||
    (rule === 'date-day' && Math.abs(whole) > 36600) ||
    (rule === 'date-serial' && (value < 0 || value >= 2958466)) ||
    (rule === 'financial-rate' && (value < 0 || value > 1)) ||
    (rule === 'financial-periods' && (!Number.isInteger(value) || value < 1 || value > 1200)) ||
    (rule === 'financial-cash' && Math.abs(value) > 1e12) ||
    (rule === 'financial-timing' && value !== 0 && value !== 1)
  )
    throw Error('Number is outside this function argument profile.');
}
function reference(value: unknown, shape: FunctionArgumentSpec['shape']): string {
  if (typeof value !== 'string' || value.length > 40) throw Error('Reference exceeds its bound.');
  const match = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d{0,6})(?::(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d{0,6}))?$/.exec(
    value.trim()
  );
  if (!match) throw Error('Choose one canonical same-sheet A1 reference or rectangle.');
  const column = (text: string) =>
    [...text.toUpperCase()].reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0);
  const c0 = column(match[2]),
    r0 = Number(match[4]),
    c1 = match[6] ? column(match[6]) : c0,
    r1 = match[8] ? Number(match[8]) : r0;
  if (
    c0 > FUNCTION_LIMITS.maximumColumn ||
    c1 > FUNCTION_LIMITS.maximumColumn ||
    r0 > FUNCTION_LIMITS.maximumRow ||
    r1 > FUNCTION_LIMITS.maximumRow ||
    r1 < r0 ||
    c1 < c0 ||
    (r1 - r0 + 1) * (c1 - c0 + 1) > FUNCTION_LIMITS.referenceCells
  )
    throw Error('Reference is outside the worksheet or picker range bound.');
  if (shape === 'scalar' && match[6]) throw Error('This argument requires one scalar cell, not a range.');
  return (
    match[1] +
    match[2].toUpperCase() +
    match[3] +
    match[4] +
    (match[6] ? ':' + match[5] + match[6].toUpperCase() + match[7] + match[8] : '')
  );
}
function argument(value: unknown, spec: FunctionArgumentSpec): string {
  plain(value);
  if (!spec.allowedKinds.includes(value.kind)) throw Error('Argument type is outside this function profile.');
  if (value.kind === 'number') {
    finite(value.value);
    numericRule(value.value, spec.numeric ?? 'finite');
    return String(value.value).replace('e', 'E');
  }
  if (value.kind === 'boolean') {
    if (typeof value.value !== 'boolean') throw Error('Boolean argument is invalid.');
    return value.value ? 'TRUE' : 'FALSE';
  }
  if (value.kind === 'reference') return reference(value.value, spec.shape);
  if (
    typeof value.value !== 'string' ||
    bytes(value.value) > FUNCTION_LIMITS.textBytes ||
    /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/u.test(value.value) ||
    [...value.value].some((char) => char.codePointAt(0)! >= 0xd800 && char.codePointAt(0)! <= 0xdfff)
  )
    throw Error('Text exceeds its bound or contains invalid XML characters.');
  if (spec.bmpText && [...value.value].some((char) => char.codePointAt(0)! > 0xffff))
    throw Error('Supplementary Unicode is outside this text function profile.');
  return '"' + value.value.replaceAll('"', '""') + '"';
}
export function serializeFunction(id: string, args: readonly unknown[]): SerializedFunction {
  const definition = functionDefinition(id);
  if (
    !Array.isArray(args) ||
    args.length < definition.minArgs ||
    args.length > definition.maxArgs ||
    args.length > FUNCTION_LIMITS.arguments ||
    Reflect.ownKeys(args).some(
      (key) =>
        key !== 'length' && (typeof key !== 'string' || !/^(?:0|[1-9]\d*)$/.test(key) || Number(key) >= args.length)
    )
  )
    throw Error('Function argument count or container is invalid.');
  const values: string[] = [];
  for (let index = 0; index < args.length; index++) {
    const descriptor = Object.getOwnPropertyDescriptor(args, String(index));
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) throw Error('Missing argument or argument accessor.');
    const spec = definition.arguments[index] ?? (definition.repeatLast ? definition.arguments.at(-1) : undefined);
    if (!spec) throw Error('Unexpected function argument.');
    values.push(argument(descriptor.value, spec));
  }
  const formula = definition.name + '(' + values.join(',') + ')';
  if (bytes(formula) > FUNCTION_LIMITS.formulaBytes) throw Error('Formula exceeds the picker byte bound.');
  return Object.freeze({
    functionId: definition.id,
    formula,
    displayFormula: '=' + formula,
    argumentCount: args.length,
    requiresOwnerValidation: true,
  });
}
