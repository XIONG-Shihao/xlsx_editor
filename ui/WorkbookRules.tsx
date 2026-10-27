import React, { useEffect, useRef, useState } from 'react';
import type {
  WorkbookConditionalRule,
  WorkbookConditionalDefinition,
  WorkbookValidationRule,
  WorkbookValidationDefinition,
  WorkbookRulesCommand,
} from '../core/contracts';
import { validateConditionalRules, validateValidationRules } from '../core/rules-model.mjs';
type Shared = {
  documentId: string;
  revision: string;
  sheet: string;
  selection: string;
  readOnly?: boolean;
  busy?: boolean;
  onBatch: (commands: WorkbookRulesCommand[]) => Promise<unknown> | unknown;
  words?: (text: string) => string;
};
export type ConditionalRulesProps = Shared & {
  inventory: (sheet: string) => Promise<WorkbookConditionalDefinition>;
  onInventory?: (definition: WorkbookConditionalDefinition) => void;
};
export type DataValidationProps = Shared & {
  inventory: (sheet: string) => Promise<WorkbookValidationDefinition>;
  onInventory?: (definition: WorkbookValidationDefinition) => void;
};
const blankConditional = (range: string): WorkbookConditionalRule => ({
  range,
  kind: 'greaterThan',
  values: [0],
  fillColor: '#FFF2CC',
  stopIfTrue: false,
});
const blankValidation = (range: string): WorkbookValidationRule => ({
  range,
  kind: 'list',
  values: ['Yes', 'No'],
  allowBlank: true,
  showDropdown: true,
  promptTitle: '',
  prompt: '',
  errorTitle: 'Invalid value',
  error: 'Choose a permitted value.',
});
function RulePanel(props: (ConditionalRulesProps | DataValidationProps) & { kind: 'conditional' | 'validation' }) {
  const { documentId, revision, sheet, selection, readOnly, busy, onBatch, inventory, kind } = props,
    words = props.words ?? ((value) => value),
    scope = documentId + '\0' + revision + '\0' + sheet,
    live = useRef(scope),
    generation = useRef(0),
    [reload, setReload] = useState(0),
    [definition, setDefinition] = useState<WorkbookConditionalDefinition | WorkbookValidationDefinition>(),
    [error, setError] = useState(''),
    [working, setWorking] = useState(false),
    [edit, setEdit] = useState<number | null>(null),
    [draft, setDraft] = useState<WorkbookConditionalRule | WorkbookValidationRule>(() =>
      kind === 'conditional' ? blankConditional(selection) : blankValidation(selection)
    );
  live.current = scope;
  useEffect(() => {
    const id = ++generation.current;
    setDefinition(undefined);
    setError('');
    setWorking(false);
    setEdit(null);
    setDraft(kind === 'conditional' ? blankConditional(selection) : blankValidation(selection));
    Promise.resolve(inventory(sheet)).then(
      (value) => {
        if (live.current === scope && generation.current === id) {
          setDefinition(value);
          (props.onInventory as ((definition: typeof value) => void) | undefined)?.(value);
        }
      },
      (reason) => {
        if (live.current === scope && generation.current === id) setError(String(reason.message ?? reason));
      }
    );
    return () => {
      generation.current++;
    };
  }, [scope, reload, inventory]);
  const disabled = !!readOnly || !!busy || working || !definition?.canEdit;
  const change = (patch: object) => setDraft((value) => ({ ...value, ...patch }) as typeof value);
  const conditional = draft as WorkbookConditionalRule,
    validation = draft as WorkbookValidationRule;
  async function publish(rules: (WorkbookConditionalRule | WorkbookValidationRule)[]) {
    if (disabled || !definition) return;
    const requestScope = scope;
    try {
      setWorking(true);
      setError('');
      const command: WorkbookRulesCommand =
        kind === 'conditional'
          ? {
              operation: 'setConditionalRules',
              sheet,
              expectedRulesRevision: definition.revision,
              rules: validateConditionalRules(rules),
            }
          : {
              operation: 'setValidationRules',
              sheet,
              expectedRulesRevision: definition.revision,
              rules: validateValidationRules(rules),
            };
      await onBatch([command]);
      if (live.current === requestScope) {
        setReload((value) => value + 1);
        setEdit(null);
        setDraft(kind === 'conditional' ? blankConditional(selection) : blankValidation(selection));
      }
    } catch (reason) {
      if (live.current === requestScope) setError(String((reason as Error).message ?? reason));
    } finally {
      if (live.current === requestScope) setWorking(false);
    }
  }
  function pick(index: number) {
    const rule = definition?.rules[index];
    if (rule) {
      setEdit(index);
      setDraft(structuredClone(rule));
    }
  }
  function move(index: number, delta: number) {
    if (!definition) return;
    const rules = [...definition.rules],
      to = index + delta;
    if (to < 0 || to >= rules.length) return;
    [rules[index], rules[to]] = [rules[to], rules[index]];
    void publish(rules);
  }
  function setKind(value: string) {
    if (kind === 'conditional') {
      const next = blankConditional(draft.range);
      next.kind = value as WorkbookConditionalRule['kind'];
      next.fillColor = conditional.fillColor;
      next.stopIfTrue = conditional.stopIfTrue;
      if (value === 'textContains') {
        delete next.values;
        next.text = 'text';
      } else if (value === 'duplicateValues') delete next.values;
      else next.values = value === 'between' ? [0, 100] : [0];
      setDraft(next);
    } else {
      const next = blankValidation(draft.range);
      next.kind = value as WorkbookValidationRule['kind'];
      if (value !== 'list') {
        delete next.values;
        next.minimum = 0;
        next.maximum = value === 'textLength' ? 255 : 100;
        next.showDropdown = false;
      }
      setDraft(next);
    }
  }
  return (
    <section
      className="xlsx-rules-panel"
      aria-label={words(kind === 'conditional' ? 'Conditional formatting' : 'Data validation')}
    >
      <p>
        {words(
          kind === 'conditional'
            ? 'First matching fill has priority. Stop if true prevents later rules.'
            : 'Rules check future literal input. Existing values are preserved. Lists use exact text.'
        )}
      </p>
      {!definition && !error && <p role="status">{words('Loading rules…')}</p>}
      {[...(definition?.protectedReasons ?? []), ...(definition?.evaluationReasons ?? [])].map((reason) => (
        <p role="status" key={reason}>
          {reason}
        </p>
      ))}
      <ol>
        {definition?.rules.map((rule, index) => (
          <li key={index}>
            <span>
              {rule.range} · {words(rule.kind)}
              {kind === 'conditional' ? ' · ' + (rule as WorkbookConditionalRule).fillColor : ''}
            </span>{' '}
            <button type="button" disabled={disabled} onClick={() => pick(index)}>
              {words('Edit rule')} {index + 1}
            </button>{' '}
            <button
              type="button"
              disabled={disabled}
              onClick={() => void publish(definition.rules.filter((_, at) => at !== index))}
            >
              {words('Delete rule')} {index + 1}
            </button>
            {kind === 'conditional' && (
              <>
                <button
                  type="button"
                  aria-label={words('Move rule up') + ' ' + (index + 1)}
                  disabled={disabled || index === 0}
                  onClick={() => move(index, -1)}
                >
                  ↑
                </button>
                <button
                  type="button"
                  aria-label={words('Move rule down') + ' ' + (index + 1)}
                  disabled={disabled || index === definition.rules.length - 1}
                  onClick={() => move(index, 1)}
                >
                  ↓
                </button>
              </>
            )}
          </li>
        ))}
      </ol>
      <form
        onSubmit={(event) => {
          event.preventDefault();
          if (!definition) return;
          const rules = [...definition.rules];
          if (edit === null) rules.push(draft as never);
          else rules[edit] = draft as never;
          void publish(rules);
        }}
      >
        <fieldset disabled={disabled}>
          <legend>
            {words(edit === null ? 'Add rule' : 'Edit rule')} {edit === null ? '' : edit + 1}
          </legend>
          <label>
            {words('Rule range')}{' '}
            <input
              aria-label={words('Rule range')}
              value={draft.range}
              onChange={(event) => change({ range: event.target.value })}
            />
          </label>{' '}
          <button type="button" onClick={() => change({ range: selection })}>
            {words('Use selection')}
          </button>
          <label>
            {words('Rule type')}{' '}
            <select
              aria-label={words('Rule type')}
              value={draft.kind}
              onChange={(event) => setKind(event.target.value)}
            >
              {(kind === 'conditional'
                ? ['greaterThan', 'lessThan', 'equal', 'between', 'textContains', 'duplicateValues']
                : ['list', 'wholeNumber', 'decimal', 'textLength']
              ).map((value) => (
                <option key={value} value={value}>
                  {words(value)}
                </option>
              ))}
            </select>
          </label>
          {kind === 'conditional' ? (
            <>
              {conditional.values?.map((value, index) => (
                <label key={index}>
                  {words(index ? 'Maximum' : 'Threshold')}{' '}
                  <input
                    aria-label={words(index ? 'Maximum' : 'Threshold')}
                    type="number"
                    step="any"
                    value={value}
                    onChange={(event) =>
                      change({
                        values: conditional.values!.map((item, at) =>
                          at === index ? Number(event.target.value) : item
                        ),
                      })
                    }
                  />
                </label>
              ))}
              {conditional.kind === 'textContains' && (
                <label>
                  {words('Contains text')}{' '}
                  <input
                    aria-label={words('Contains text')}
                    maxLength={128}
                    value={conditional.text ?? ''}
                    onChange={(event) => change({ text: event.target.value })}
                  />
                </label>
              )}
              <label>
                {words('Fill color')}{' '}
                <input
                  aria-label={words('Fill color')}
                  type="color"
                  value={conditional.fillColor}
                  onChange={(event) => change({ fillColor: event.target.value })}
                />
              </label>
              <label>
                <input
                  type="checkbox"
                  checked={conditional.stopIfTrue}
                  onChange={(event) => change({ stopIfTrue: event.target.checked })}
                />
                {words('Stop if true')}
              </label>
            </>
          ) : (
            <>
              {validation.kind === 'list' ? (
                <label>
                  {words('List values (one per line)')}{' '}
                  <textarea
                    aria-label={words('List values (one per line)')}
                    value={validation.values?.join('\n') ?? ''}
                    onChange={(event) => change({ values: event.target.value.split('\n') })}
                  />
                </label>
              ) : (
                <>
                  {['minimum', 'maximum'].map((key) => (
                    <label key={key}>
                      {words(key === 'minimum' ? 'Minimum' : 'Maximum')}{' '}
                      <input
                        aria-label={words(key === 'minimum' ? 'Minimum' : 'Maximum')}
                        type="number"
                        step={validation.kind === 'decimal' ? 'any' : '1'}
                        value={validation[key as 'minimum' | 'maximum'] ?? 0}
                        onChange={(event) => change({ [key]: Number(event.target.value) })}
                      />
                    </label>
                  ))}
                </>
              )}
              <label>
                <input
                  type="checkbox"
                  checked={validation.allowBlank}
                  onChange={(event) => change({ allowBlank: event.target.checked })}
                />
                {words('Allow blank')}
              </label>
              {validation.kind === 'list' && (
                <label>
                  <input
                    type="checkbox"
                    checked={validation.showDropdown}
                    onChange={(event) => change({ showDropdown: event.target.checked })}
                  />
                  {words('Show dropdown')}
                </label>
              )}
              {(['promptTitle', 'prompt', 'errorTitle', 'error'] as const).map((key) => (
                <label key={key}>
                  {words(
                    {
                      promptTitle: 'Input title',
                      prompt: 'Input message',
                      errorTitle: 'Error title',
                      error: 'Error message',
                    }[key]
                  )}{' '}
                  <input
                    aria-label={words(
                      {
                        promptTitle: 'Input title',
                        prompt: 'Input message',
                        errorTitle: 'Error title',
                        error: 'Error message',
                      }[key]
                    )}
                    maxLength={key === 'promptTitle' || key === 'errorTitle' ? 32 : key === 'prompt' ? 255 : 225}
                    value={validation[key]}
                    onChange={(event) => change({ [key]: event.target.value })}
                  />
                </label>
              ))}
            </>
          )}
          <button type="submit">{words(edit === null ? 'Add rule' : 'Apply rule')}</button>
          <button
            type="button"
            onClick={() => {
              setEdit(null);
              setDraft(kind === 'conditional' ? blankConditional(selection) : blankValidation(selection));
            }}
          >
            {words('Reset form')}
          </button>
        </fieldset>
      </form>
      {error && <p role="alert">{error}</p>}
    </section>
  );
}
export function ConditionalRules(props: ConditionalRulesProps) {
  return <RulePanel {...props} kind="conditional" />;
}
export function DataValidation(props: DataValidationProps) {
  return <RulePanel {...props} kind="validation" />;
}
