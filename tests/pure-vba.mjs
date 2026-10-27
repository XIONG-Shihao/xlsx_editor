import test from 'node:test';
import assert from 'node:assert/strict';
import { compilePureVba, pureVbaFunctions, evaluatePureVba, PURE_VBA_LIMITS } from '../core/pure-vba.mjs';
const arithmetic =
  'Attribute VB_Name = "Arithmetic"\r\nOption Explicit\r\nPublic Function Calc(ByVal a As Double, ByVal b As Double) As Double\r\n Dim total As Double\r\n total = a * 3# + b / 2#\r\n Calc = -(total - 1#)\r\nEnd Function\r\n';
test('explicit numeric AST matches qualified Double arithmetic and Integer constant semantics', () => {
  assert.deepEqual(pureVbaFunctions(arithmetic), [{ name: 'Calc', parameters: ['a', 'b'], returnType: 'Double' }]);
  assert.equal(evaluatePureVba(arithmetic, 'cAlC', [2, 8]).value, -9);
  assert.equal(evaluatePureVba(arithmetic, 'Calc', [-0, 0]).value, 1);
  assert.equal(evaluatePureVba('Function MyFunc()\nMyFunc = 42\nEnd Function', 'MyFunc', []).value, 42);
  assert.equal(
    evaluatePureVba('Function D() As Double\nD = 1# / 10# + 2# / 10#\nEnd Function', 'D', []).value,
    0.1 + 0.2
  );
  assert.equal(evaluatePureVba('Function D() As Double\nDim x As Double\nD = x + 1#\nEnd Function', 'D', []).value, 1);
  assert.equal(evaluatePureVba('Function D() As Double\nD = -0#\nEnd Function', 'D', []).value, -0);
});
test('host, event, mutation, coercion and unsupported syntax are closed refusals', () => {
  for (const source of [
    'Sub Auto_Open()\nEnd Sub',
    'Declare Function Shell Lib "a" () As Double',
    'Function D() As Double\nD = Application.Cells(1,1)\nEnd Function',
    'Function D() As Double\nD = Shell("a")\nEnd Function',
    'Function D() As Double\nD = 2# ^ 3#\nEnd Function',
    'Function D() As Double\nD = "2"\nEnd Function',
    'Function D() As Double\nD = 30000 + 30000\nEnd Function',
    'Function D(ByRef x As Double) As Double\nD = x\nEnd Function',
    'Function D(ByVal x As Double) As Double\nx = 1#\nD = x\nEnd Function',
    'Function D()\nD = 32768\nEnd Function',
    'Function D() As Double\nD = 1#\nApplication = 1#\nEnd Function',
    'Function D() As Double\nIf True Then\nD = 1#\nEnd Function',
    'Function D() As Double\nD = 1#\nEnd Function\nDim global As Double',
    'Function D() As Double\nD = 1#\nEnd Function\nPrivate Sub event()\nEnd Sub',
  ])
    assert.throws(() => compilePureVba(source), { code: 'workbook-refused' });
  assert.throws(() => evaluatePureVba(arithmetic, 'Calc', [2]), /exact finite/);
  assert.throws(() => evaluatePureVba(arithmetic, 'Calc', [Infinity, 1]), /exact finite/);
  assert.throws(() => evaluatePureVba(arithmetic, 'Shell', [2, 8]), /identifier/);
  assert.throws(() => evaluatePureVba(arithmetic, 'Unknown', [2, 8]), /not declared/);
});
test('numeric and resource boundaries fail before returning an invalid result', () => {
  assert.throws(() => compilePureVba(' '.repeat(PURE_VBA_LIMITS.sourceBytes + 1)), /budget/);
  assert.throws(
    () => evaluatePureVba('Function D() As Double\nD = 1# / 0#\nEnd Function', 'D', []),
    /division by zero/
  );
  assert.throws(
    () => evaluatePureVba('Function D() As Double\nD = 1E308# * 1E308#\nEnd Function', 'D', []),
    /overflow/
  );
  assert.throws(
    () => compilePureVba('Function D() As Double\nD = ' + '('.repeat(40) + '1#' + ')'.repeat(40) + '\nEnd Function'),
    /depth/
  );
  assert.throws(
    () => evaluatePureVba('Function D() As Double\nD = ' + Array(40).fill('1#').join('+') + '\nEnd Function', 'D', []),
    /work budget/
  );
  assert.throws(
    () =>
      compilePureVba(
        'Function D() As Double\n' +
          Array(PURE_VBA_LIMITS.statements + 1)
            .fill('D = 1#')
            .join('\n') +
          '\nEnd Function'
      ),
    /statement budget/
  );
});
