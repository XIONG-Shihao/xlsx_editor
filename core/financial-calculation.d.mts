/** Closed syntax declarations; actual numeric/source admission belongs to the accepted owner. */
export type FinancialName = 'PMT' | 'PV' | 'FV';
export type FinancialOperand =
  | { kind: 'number'; value: number }
  | { kind: 'reference'; reference: string }
  | { kind: 'unary'; operator: '+' | '-'; value: FinancialOperand }
  | { kind: 'binary'; operator: '+' | '-' | '*' | '/'; left: FinancialOperand; right: FinancialOperand };
export interface FinancialCall {
  name: FinancialName;
  args: FinancialOperand[];
  index: number;
  start: number;
  end: number;
  references: string[];
  values?: number[];
}
export interface FinancialExpressionSyntax {
  text: string;
  ast: FinancialOperand;
  nodes: number;
  references: string[];
  requiresOwnerValidation: true;
}
export const FINANCIAL_FUNCTIONS: ReadonlySet<FinancialName>;
export const FINANCIAL_REASON: string;
export function financialArgument(text: string): FinancialOperand;
export function financialExpressionSyntax(raw: string): FinancialExpressionSyntax;
export function financialCalls(formula: string): FinancialCall[];
export function rewriteFinancialFormula(formula: string, calls: FinancialCall[]): string;
export function calculateFinancialScalar(
  name: FinancialName,
  input: readonly number[]
): { value: number; unresolved?: never } | { value?: never; unresolved: string };
export function checkFinancialFormula(
  formula: string,
  cells: Readonly<Record<string, unknown>>
): { calls: FinancialCall[]; unresolved?: never } | { calls?: never; unresolved: string };
