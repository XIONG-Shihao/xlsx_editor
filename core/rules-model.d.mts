import type { WorkbookConditionalRule, WorkbookValidationRule } from './contracts.js';
export const RULE_LIMITS: { rules: number; positions: number; listCharacters: number };
export function validateConditionalRules(input: unknown): WorkbookConditionalRule[];
export function validateValidationRules(input: unknown): WorkbookValidationRule[];
export function rangesOverlap(a: string, b: string): boolean;
export function ruleContains(rule: { range: string }, ref: string): boolean;
export function acceptedRuleValue(cell: any): { qualified: boolean; value?: string | number | boolean | null };
export function duplicateRuleKey(value: unknown): string;
export function conditionalRuleMatches(
  rule: WorkbookConditionalRule,
  cell: any,
  duplicates?: Record<string, number>
): boolean;
export function conditionalFill(
  rules: WorkbookConditionalRule[],
  ref: string,
  cell: any,
  duplicates?: Record<string, Record<string, number>>
): string | undefined;
export function assertValidationEdits(rules: WorkbookValidationRule[], edits: any[]): void;
export function containsFormula(range: string, text: string): string;
