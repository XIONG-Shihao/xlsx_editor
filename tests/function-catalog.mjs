import './support/scratch-dir.mjs';
/** Pure production catalog/syntax checks, not evaluator or actual owner qualification. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import { build } from 'esbuild';
const component = fileURLToPath(new URL('../', import.meta.url)),
  directory = path.join(component, '.build/tests');
await fs.mkdir(directory, { recursive: true });
const source = path.join(component, 'ui/functionCatalog.ts');
const before = await fs.readFile(source),
  sha = (bytes) => createHash('sha256').update(bytes).digest('hex');
const built = await build({
    entryPoints: [source],
    bundle: true,
    write: false,
    platform: 'node',
    format: 'esm',
    target: 'es2022',
    logLevel: 'silent',
  }),
  compiled = { code: built.outputFiles[0].text, warnings: built.warnings };
const model = await import('data:text/javascript;base64,' + Buffer.from(compiled.code).toString('base64'));
const failures = [];
function check(name, work) {
  test(name, async () => {
    try {
      await work();
    } catch (error) {
      failures.push({ name, message: String(error?.message ?? error) });
      throw error;
    }
  });
}
after(async () => {
  const unchanged = sha(await fs.readFile(source)) === sha(before);
  if (!unchanged) failures.push({ name: 'source-unchanged', message: 'Production module changed during pure checks' });
});
const n = (value) => ({ kind: 'number', value }),
  b = (value) => ({ kind: 'boolean', value }),
  s = (value) => ({ kind: 'text', value }),
  r = (value) => ({ kind: 'reference', value });
const serialize = (id, args) => model.serializeFunction(id, args);
check(
  'catalog exposes twenty-six declared profiles, immutable categories and three enabled financial functions',
  () => {
    assert.equal(model.FUNCTION_CATALOG.length, 26);
    assert.equal(new Set(model.FUNCTION_CATALOG.map((x) => x.id)).size, 26);
    assert.deepEqual(
      model.FUNCTION_CATEGORIES.map((x) => x.id),
      ['Financial', 'Logical', 'Text', 'DateTime', 'Lookup', 'Math', 'More']
    );
    assert.deepEqual(
      model.listFunctions({ category: 'Financial' }).map((item) => item.id),
      ['PMT', 'PV', 'FV']
    );
    assert.match(model.FUNCTION_CATEGORIES[0].notice, /Results can differ from Excel/);
    for (const id of [
      'NPV',
      'IRR',
      'FILTER',
      'SEQUENCE',
      'ROUNDUP',
      'PLATFORM_B7_ADD',
      'SUM);NOW(',
      'sum',
      'constructor',
    ])
      assert.throws(() => model.functionDefinition(id));
    assert.throws(() => model.FUNCTION_CATALOG.push({}));
    assert.throws(() => {
      model.functionDefinition('SUM').arguments[0].allowedKinds.push('text');
    });
  }
);
check('category/search browsing returns actual closed profiles and refuses unknown or excessive metadata', () => {
  assert.deepEqual(
    model.listFunctions({ category: 'Lookup' }).map((x) => x.id),
    ['ROWS', 'COLUMNS', 'VLOOKUP', 'XLOOKUP']
  );
  assert.deepEqual(
    model.listFunctions({ query: '  dec2hex  ' }).map((x) => x.id),
    ['DEC2HEX']
  );
  assert.deepEqual(
    model.listFunctions({ category: 'Text', query: 'spaces' }).map((x) => x.id),
    ['TRIM']
  );
  for (const options of [{ category: 'Cloud' }, { query: 'x'.repeat(101) }, { query: 7 }, { extra: true }, null, []])
    assert.throws(() => model.listFunctions(options));
});
check(
  'typed scalar calls preserve literal meaning, omit optional arguments and make owner qualification explicit',
  () => {
    assert.deepEqual(serialize('IF', [b(true), n(10), n(0)]), {
      functionId: 'IF',
      formula: 'IF(TRUE,10,0)',
      displayFormula: '=IF(TRUE,10,0)',
      argumentCount: 3,
      requiresOwnerValidation: true,
    });
    assert.equal(serialize('IF', [n(0), s('yes'), b(false)]).formula, 'IF(0,"yes",FALSE)');
    assert.equal(serialize('LEFT', [s('abc')]).formula, 'LEFT("abc")');
    assert.equal(serialize('LEFT', [s('abc'), n(0)]).formula, 'LEFT("abc",0)');
    assert.equal(serialize('DATE', [n(2026), n(10), n(1)]).formula, 'DATE(2026,10,1)');
    assert.equal(serialize('DEC2HEX', [n(255)]).formula, 'DEC2HEX(255)');
    assert.equal(serialize('ABS', [n(-1.25)]).formula, 'ABS(-1.25)');
    assert.equal(serialize('ROUND', [n(-2.675), n(2)]).formula, 'ROUND(-2.675,2)');
  }
);
check('literal formula-looking text is quoted and escaped rather than emitted as an executable fragment', () => {
  const samples = ['He said "go"', '=SUM(A1)', 'x"),NOW(),("y', '#N/A', 'a,b;C1', 'first\nsecond\tline', '中文'];
  for (const value of samples) {
    const formula = serialize('IFERROR', [r('A1'), s(value)]).formula;
    assert.equal(formula, 'IFERROR(A1,"' + value.replaceAll('"', '""') + '")');
    const encoded = formula.slice('IFERROR(A1,"'.length, -2);
    assert.equal(encoded.replaceAll('""', '"'), value);
  }
  assert.throws(() => serialize('SUM', [s('1),NOW()')]));
  assert.throws(() => serialize('IF', [s('TRUE'), n(1), n(2)]));
});
check(
  'canonical same-sheet references retain anchors and reject expressions, foreign contexts and scalar-range ambiguity',
  () => {
    assert.equal(serialize('SUM', [r(' $a$1:$b$2 '), r('c3'), n(5)]).formula, 'SUM($A$1:$B$2,C3,5)');
    assert.equal(serialize('YEAR', [r('$a1')]).formula, 'YEAR($A1)');
    const invalid = [
      'Sheet1!A1',
      "'Other Sheet'!A1",
      '[other.xlsx]Sheet1!A1',
      'Name',
      'Items[Value]',
      'A1#',
      '@A1',
      '1:3',
      'A:C',
      'A1+B1',
      'SUM(A1)',
      'A1,B1',
      'A1:B1:B2',
      'A0',
      'A01',
      'XFE1',
      'A1048577',
      'B2:A1',
    ];
    for (const value of invalid) assert.throws(() => serialize('SUM', [r(value)]), value);
    for (const id of ['YEAR', 'LEN', 'IFERROR'])
      assert.throws(() => serialize(id, id === 'IFERROR' ? [r('A1:A1'), n(0)] : [r('A1:A1')]), id);
  }
);
check(
  'rectangle limits independently admit edge coordinates and exact thousand positions, refusing thousand-and-one',
  () => {
    assert.equal(serialize('ROWS', [r('A1:A1000')]).formula, 'ROWS(A1:A1000)');
    assert.equal(serialize('COLUMNS', [r('A1:J100')]).formula, 'COLUMNS(A1:J100)');
    assert.equal(serialize('SUM', [r('XFD1048576')]).formula, 'SUM(XFD1048576)');
    assert.throws(() => serialize('SUM', [r('A1:A1001')]));
    assert.throws(() => serialize('ROWS', [r('A1:K100')]));
    assert.throws(() => serialize('SUM', [r('A1:ZZZ1')]));
  }
);
check(
  'literal numeric domain and function-specific bounds are closed without coercion or signed-zero rewriting',
  () => {
    for (const value of [NaN, Infinity, -Infinity, -0, Number.MAX_SAFE_INTEGER + 1, '2', true, null])
      assert.throws(() => serialize('ABS', [n(value)]));
    assert.equal(serialize('ABS', [n(1e-7)]).formula, 'ABS(1E-7)');
    for (const value of [0, 12]) assert.equal(serialize('ROUND', [n(1.25), n(value)]).argumentCount, 2);
    for (const value of [-1, 13, 1.5]) assert.throws(() => serialize('ROUND', [n(1.25), n(value)]));
    for (const value of [0, 2147483647]) assert.equal(serialize('DEC2HEX', [n(value)]).argumentCount, 1);
    for (const value of [-1, 2147483648, 2.5]) assert.throws(() => serialize('DEC2HEX', [n(value)]));
    for (const value of [-1, 1.5]) assert.throws(() => serialize('LEFT', [s('text'), n(value)]));
  }
);
check(
  'raw numeric widget text refuses empty/coercing/rounded values before conversion and preserves equivalent decimal spellings',
  () => {
    for (const [raw, expected] of [
      [' .5 ', 0.5],
      ['+2', 2],
      ['1E-7', 1e-7],
      ['0.100000000000000000', 0.1],
      ['120.00', 120],
      ['2.5e2', 250],
      ['0', 0],
    ])
      assert.equal(model.parseNumberLiteral(raw), expected);
    for (const raw of [
      '',
      ' ',
      '-0',
      '-0.000',
      '0x10',
      '12%',
      '1,000',
      '3/12',
      'DATE(2026,1,1)',
      '2026/10/1',
      '1_000',
      'Infinity',
      '1e400',
      '1e-10000',
      '0.123456789012345678901',
      '9007199254740993',
    ])
      assert.throws(() => model.parseNumberLiteral(raw), raw);
    assert.throws(() => model.parseNumberLiteral(5));
    assert.throws(() => model.parseNumberLiteral('0'.repeat(81)));
  }
);
check('date literal normalization bounds do not guess a calendar or resolved reference type', () => {
  assert.equal(serialize('DATE', [n(9999.9), n(-1200), n(36600)]).formula, 'DATE(9999.9,-1200,36600)');
  assert.equal(serialize('YEAR', [n(60.75)]).formula, 'YEAR(60.75)');
  assert.equal(serialize('DAY', [r('A1')]).requiresOwnerValidation, true);
  for (const args of [
    [n(-1), n(1), n(1)],
    [n(10000), n(1), n(1)],
    [n(2026), n(-1200.1), n(1)],
    [n(2026), n(1), n(36601)],
  ])
    assert.throws(() => serialize('DATE', args));
  for (const value of [-1, 2958466]) assert.throws(() => serialize('MONTH', [n(value)]));
  assert.throws(() => serialize('DATE', [s('2026'), n(10), n(1)]));
});
check('text byte/XML/BMP guards cover multibyte and invalid Unicode without silently modifying content', () => {
  assert.equal(serialize('LEN', [s('a'.repeat(4096))]).formula.length, 4103);
  assert.equal(serialize('LEN', [s('中'.repeat(1365))]).requiresOwnerValidation, true);
  for (const value of ['x'.repeat(4097), '中'.repeat(1366), '\u0000', '\u000B', '\uFFFF', '\uD800', '\uDC00'])
    assert.throws(() => serialize('LEN', [s(value)]));
  assert.throws(() => serialize('LEN', [s('🙂')]));
  assert.throws(() => serialize('LEFT', [s('🙂')]));
  assert.equal(serialize('IFERROR', [r('A1'), s('🙂')]).formula, 'IFERROR(A1,"🙂")');
  assert.equal(
    new TextEncoder().encode(serialize('IF', [b(true), s('a'.repeat(4000)), s('b'.repeat(4000))]).formula).length,
    8014
  );
  assert.throws(() => serialize('IF', [b(true), s('"'.repeat(3000)), s('"'.repeat(3000))]), /Formula exceeds/);
});
check(
  'arity, holes, extra fields and accessors refuse before invoking caller code; accepted input stays unchanged',
  () => {
    for (const args of [
      [],
      Array.from({ length: 17 }, () => n(1)),
      [undefined],
      [null],
      ['A1'],
      [{ kind: 'number', value: 1, formula: 'NOW()' }],
    ])
      assert.throws(() => serialize('SUM', args));
    assert.throws(() => serialize('DEC2HEX', [n(255), n(4)]));
    assert.throws(() => serialize('IF', [b(true), n(1)]));
    const hole = [];
    hole.length = 1;
    assert.throws(() => serialize('SUM', hole));
    let invoked = 0;
    const accessor = {
      kind: 'number',
      get value() {
        invoked++;
        return 1;
      },
    };
    assert.throws(() => serialize('SUM', [accessor]));
    assert.equal(invoked, 0);
    const array = [];
    Object.defineProperty(array, '0', {
      get() {
        invoked++;
        return n(1);
      },
    });
    assert.throws(() => serialize('SUM', array));
    assert.equal(invoked, 0);
    const extras = [n(1)];
    extras.unexpected = true;
    assert.throws(() => serialize('SUM', extras));
    const inputs = [r(' $a$1:$a$3 '), n(3)],
      snapshot = structuredClone(inputs);
    serialize('SUM', inputs);
    assert.deepEqual(inputs, snapshot);
  }
);
check(
  'aggregate syntax supports independent mixed argument vectors while lookup/Boolean/text profiles stay narrow',
  () => {
    assert.equal(serialize('SUM', [r('A1:A3'), n(2), r('$C$8')]).formula, 'SUM(A1:A3,2,$C$8)');
    assert.equal(
      serialize(
        'COUNT',
        Array.from({ length: 16 }, (_, i) => n(i))
      ).argumentCount,
      16
    );
    for (const id of ['SUM', 'AVERAGE', 'COUNT', 'MIN', 'MAX']) assert.throws(() => serialize(id, [b(true)]));
    assert.throws(() => serialize('ROWS', [n(3)]));
    assert.throws(() => serialize('LEN', [n(12)]));
    assert.throws(() => serialize('IFERROR', [{ kind: 'formula', value: '1/0' }, n(99)]));
    for (const id of ['XLOOKUP', 'EDATE', 'EOMONTH', 'SUM(SEQUENCE)', 'SUM(UNIQUE)'])
      assert.throws(() => serialize(id, [n(1)]));
  }
);
