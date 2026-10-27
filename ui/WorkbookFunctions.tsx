/** Retained function pane with closed Financial arithmetic inputs. */
import { useLayoutEffect, useMemo, useRef, useState } from 'react';
import {
  FUNCTION_CATEGORIES,
  FUNCTION_CATALOG,
  functionDefinition,
  listFunctions,
  parseNumberLiteral,
  serializeFunction,
  type FunctionArgument,
  type FunctionCategory,
  type FunctionId,
  type FunctionDefinition,
  type NumericRule,
} from './functionCatalog';
import {
  FunctionInsertionCancelled,
  FunctionInsertionRetired,
  functionBindingKey,
  inspectFunctionTarget,
  insertFunction,
  targetDescription,
  targetFingerprint,
  targetOccupied,
  targetProtection,
  type FunctionInsertionContext,
  type FunctionInsertionResult,
  type FunctionTarget,
} from './functionInsertion';
import { createFunctionRecentStore } from './functionRecents';
import { FormulaDraftPreview } from './WorkbookFormulaPreview';
import type { FormulaDraftPreviewController, FormulaPreviewScope } from './formulaDraftPreview';

type ArgumentKind = FunctionArgument['kind'];
interface ArgumentDraft {
  kind: ArgumentKind;
  value: string;
}
export interface WorkbookFunctionsProps {
  context: FunctionInsertionContext;
  visible: boolean;
  browse?: { mode: 'all' | 'recent' | FunctionCategory; serial: number };
  onRestart?: () => void;
  words?: (text: string) => string;
  onOutcome?: (result: FunctionInsertionResult) => void;
  formulaPreview?: { controller: FormulaDraftPreviewController; scope: FormulaPreviewScope };
}
const identity = (value: string) => value;
function initialArgument(id: FunctionId, index: number, context: FunctionInsertionContext): ArgumentDraft {
  const definition = functionDefinition(id),
    spec = definition.arguments[Math.min(index, definition.arguments.length - 1)];
  const kind: ArgumentKind = spec.allowedKinds.includes('reference') ? 'reference' : spec.allowedKinds[0];
  return { kind, value: kind === 'boolean' ? 'false' : '' };
}
function argumentValue(draft: ArgumentDraft): FunctionArgument {
  if (draft.kind === 'number') {
    return { kind: 'number', value: parseNumberLiteral(draft.value) };
  }
  if (draft.kind === 'boolean') return { kind: 'boolean', value: draft.value === 'true' };
  return { kind: draft.kind, value: draft.value };
}
function typedArguments(id: FunctionId, drafts: ArgumentDraft[]): FunctionArgument[] {
  const definition = functionDefinition(id),
    values = [...drafts];
  while (values.length > definition.minArgs) {
    const last = values.at(-1)!,
      spec = definition.arguments[values.length - 1];
    if (
      (last.kind === 'number' || last.kind === 'reference' || last.kind === 'expression') &&
      !last.value.trim() &&
      (spec?.optional || definition.repeatLast)
    )
      values.pop();
    else break;
  }
  return values.map(argumentValue);
}
function supportedHints(definition: FunctionDefinition): string[] {
  const numeric: Record<NumericRule, string> = {
    finite: 'Use a finite number within the supported precision.',
    'nonnegative-integer': 'Use a whole number of zero or greater.',
    'round-digits': 'Use a whole number from 0 to 12.',
    'dec2hex-integer': 'Use a whole number from 0 to 2,147,483,647.',
    'date-year': 'After rounding down, the year must be from 0 to 9,999.',
    'date-month': 'After rounding down, the month must be within −1,200 to 1,200.',
    'date-day': 'After rounding down, the day must be within −36,600 to 36,600.',
    'date-serial': 'Use a serial date from 0 to less than 2,958,466; the workbook date system determines the calendar.',
    'financial-rate': 'Use a period rate from 0 to 1. Enter 0.005 for 0.5% per period.',
    'financial-periods': 'Use a whole number from 1 to 1,200; match the time unit used by Period rate.',
    'financial-cash':
      'Use cash with magnitude at most 1,000,000,000,000. Incoming cash is positive; outgoing cash is negative.',
    'column-index': 'Use a whole column number starting at 1 within the selected table.',
    'financial-timing':
      'Use numeric 0 for end of period, or numeric 1 for beginning of period; omission defaults to 0.',
  };
  const hints = definition.arguments.flatMap((spec) => [
    ...(spec.numeric ? [spec.label + ': ' + numeric[spec.numeric]] : []),
    ...(spec.bmpText ? [spec.label + ': Supplementary Unicode characters, such as 🙂, are not supported.'] : []),
    ...(spec.shape === 'scalar' && spec.allowedKinds.includes('reference')
      ? [spec.label + ': Cell references must select one cell.']
      : []),
  ]);
  if (definition.category === 'Financial')
    hints.push(
      'Optional balance defaults to 0. To set payment timing, supply the balance first; no empty middle argument is emitted.',
      'Use a single ordinary numeric input cell on this worksheet. Formula cells, empty cells, text, TRUE/FALSE, errors, array results and scenario cells are not supported for these inputs.',
      'The calculation must stay finite and within supported bounds. Some high rate and long period combinations cannot be calculated yet.',
      'Results can differ from Excel in the last decimal places.'
    );
  if (definition.category === 'Financial')
    hints.push(
      'Choose Expression for arithmetic such as B1/12 or B2*12. Only numbers, single-cell inputs, parentheses and + - * / are supported. Nested functions, names, ranges and workbook links are not supported.',
      'An expression allows up to 1,024 bytes, 32 terms and operators and eight levels of grouping; a call allows up to 32 cell references. Use uppercase cell addresses such as B1.'
    );
  if (definition.id === 'VLOOKUP')
    hints.push('Exact match: FALSE is included automatically. The first matching row is returned.');
  if (definition.id === 'XLOOKUP')
    hints.push(
      'Exact match, first to last. Both ranges must be single rows or columns with the same number of cells. Blank and formula-derived inputs are not supported yet.'
    );
  if (definition.id === 'SUMIF' || definition.id === 'COUNTIF')
    hints.push(
      'Choose Text for a condition such as Sales, >=10 or A*. Choose Number for a numeric condition, or Reference for a criterion cell.'
    );
  if (definition.id === 'TRIM')
    hints.push('Only ordinary spaces are trimmed. Tabs and nonbreaking spaces stay in the text.');
  return hints;
}
export function WorkbookFunctions({
  context,
  visible,
  browse,
  words = identity,
  onOutcome,
  onRestart,
  formulaPreview,
}: WorkbookFunctionsProps): JSX.Element {
  const key = functionBindingKey(context.binding),
    current = context.isCurrent(context.binding);
  const runtimeCurrent = context.isRuntimeCurrent?.(context.binding) ?? current;
  const acknowledgedState = context.currentState?.() ?? context.binding.state;
  const live = useRef({ key, current, runtimeCurrent, visible });
  live.current = { key, current, runtimeCurrent, visible };
  const generation = useRef(0),
    abort = useRef<AbortController | null>(null),
    acceptance = useRef<{ key: string; serial: number } | null>(null);
  const recentStore = useMemo(() => {
    try {
      return createFunctionRecentStore(window.localStorage);
    } catch {
      return createFunctionRecentStore(null);
    }
  }, []);
  const [, setRecentVersion] = useState(0),
    recentIds = recentStore.read();
  const [category, setCategory] = useState<FunctionCategory | 'all' | 'recent'>('all'),
    [query, setQuery] = useState(''),
    [id, setId] = useState<FunctionId>(FUNCTION_CATALOG[0].id);
  const [args, setArgs] = useState<ArgumentDraft[]>(() =>
    Array.from({ length: functionDefinition(id).minArgs }, (_, index) => initialArgument(id, index, context))
  );
  const [target, setTarget] = useState<FunctionTarget | null>(null),
    [checking, setChecking] = useState(false),
    [accepting, setAccepting] = useState(false),
    [status, setStatus] = useState(''),
    [error, setError] = useState('');
  const [accepted, setAccepted] = useState<FunctionInsertionResult | null>(null),
    [acceptedTarget, setAcceptedTarget] = useState<Awaited<
      ReturnType<NonNullable<FunctionInsertionContext['readAcceptedTarget']>>
    > | null>(null);
  const [resultLoading, setResultLoading] = useState(false),
    resultSerial = useRef(0);
  const acceptedVisible = !!accepted && functionBindingKey(accepted.binding) === key && runtimeCurrent;
  const operationCurrent = (serial: number) =>
    live.current.key === key && live.current.current && live.current.visible && generation.current === serial;
  const changeFunction = (next: FunctionId): void => {
    setId(next);
    setArgs(
      Array.from({ length: functionDefinition(next).minArgs }, (_, index) => initialArgument(next, index, context))
    );
    setError('');
    setStatus('');
  };
  const matching = (selected: typeof category, text = query) => {
    const found = listFunctions({
      ...(selected === 'all' || selected === 'recent' ? {} : { category: selected }),
      query: text,
    });
    return selected === 'recent'
      ? recentIds.map(functionDefinition).filter((item) => found.some((match) => match.id === item.id))
      : found;
  };
  const functions = matching(category),
    definition = functionDefinition(id);
  const categories = FUNCTION_CATEGORIES.filter((item) =>
    FUNCTION_CATALOG.some((member) => member.category === item.id)
  );
  let formula = '',
    formulaError = '';
  try {
    formula = serializeFunction(id, typedArguments(id, args)).displayFormula;
  } catch (failure) {
    formulaError = failure instanceof Error ? failure.message : String(failure);
  }
  const previewDraft = useRef({ text: '', generation: 0 });
  const previewText = JSON.stringify([key, formula, formulaError]);
  if (previewDraft.current.text !== previewText)
    previewDraft.current = { text: previewText, generation: previewDraft.current.generation + 1 };
  const scope = context.availability(),
    protectedReason = target ? targetProtection(target) : null;
  const unavailable = !current
    ? 'The initiating workbook or target changed. Reopen Insert Function.'
    : !visible
      ? 'Open Insert Function before inserting.'
      : !scope.active
        ? 'Activate this workbook first.'
        : !context.accept
          ? 'Workbook commands are unavailable in this editor.'
          : scope.readOnly
            ? 'This workbook is read-only.'
            : scope.literal
              ? 'This command requires an editable XLSX workbook.'
              : scope.specialSelection
                ? 'Click a cell to leave the special selection.'
                : scope.busy
                  ? 'Wait for the workbook operation to finish.'
                  : protectedReason;
  const readPreview = (): void => {
    const controller = new AbortController(),
      serial = ++generation.current;
    abort.current?.abort();
    abort.current = controller;
    setTarget(null);
    setChecking(true);
    setError('');
    setStatus('');
    void inspectFunctionTarget(context, controller.signal)
      .then((value) => {
        if (operationCurrent(serial)) setTarget(value);
      })
      .catch((failure) => {
        if (
          operationCurrent(serial) &&
          !(failure instanceof FunctionInsertionCancelled) &&
          !(failure instanceof FunctionInsertionRetired)
        )
          setError(failure instanceof Error ? failure.message : String(failure));
      })
      .finally(() => {
        if (operationCurrent(serial)) {
          setChecking(false);
          abort.current = null;
        }
      });
  };
  useLayoutEffect(() => {
    generation.current++;
    abort.current?.abort();
    abort.current = null;
    if (!current) {
      setTarget(null);
      setChecking(false);
      setAccepting(false);
      setError('');
      setStatus('');
      return;
    }
    if (!visible) {
      setChecking(false);
      return;
    }
    readPreview();
    return () => {
      abort.current?.abort();
      generation.current++;
    };
  }, [key, current, visible]);
  useLayoutEffect(() => {
    setCategory('all');
    setQuery('');
    changeFunction(FUNCTION_CATALOG[0].id);
    setAccepting(!!acceptance.current);
    setAccepted(null);
  }, [key]);
  useLayoutEffect(() => {
    if (!browse) return;
    const mode = browse.mode,
      next =
        mode === 'recent'
          ? (recentStore.read()[0] ?? FUNCTION_CATALOG[0].id)
          : mode === 'all'
            ? FUNCTION_CATALOG[0].id
            : listFunctions({ category: mode })[0]?.id;
    setAccepted(null);
    setCategory(mode);
    setQuery('');
    changeFunction(next ?? FUNCTION_CATALOG[0].id);
  }, [browse?.mode, browse?.serial, key]);
  const checkAcceptedResult = (): void => {
    const serial = ++resultSerial.current;
    setAcceptedTarget(null);
    if (!acceptedVisible || !accepted || !context.readAcceptedTarget) {
      setResultLoading(false);
      return;
    }
    setResultLoading(true);
    void context
      .readAcceptedTarget(accepted.formula)
      .then((value) => {
        if (live.current.key === key && live.current.runtimeCurrent && resultSerial.current === serial)
          setAcceptedTarget(value);
      })
      .catch(() => {})
      .finally(() => {
        if (live.current.key === key && resultSerial.current === serial) setResultLoading(false);
      });
  };
  useLayoutEffect(() => {
    checkAcceptedResult();
    return () => {
      resultSerial.current++;
    };
  }, [acceptedVisible, accepted, current, visible, key, acknowledgedState]);
  const cancel = (): void => {
    if (accepting) return;
    generation.current++;
    abort.current?.abort();
    abort.current = null;
    setChecking(false);
    setError('');
    setStatus('Insertion cancelled. The workbook was not changed.');
  };
  const submit = async (): Promise<void> => {
    if (unavailable || checking || accepting || acceptance.current || !target || formulaError) return;
    const controller = new AbortController(),
      serial = ++generation.current,
      capturedOutcome = onOutcome,
      accepted = { key, serial };
    abort.current = controller;
    setChecking(true);
    setError('');
    setStatus('Checking the accepted target…');
    try {
      const result = await insertFunction(
        context,
        {
          id,
          args: typedArguments(id, args),
          ...(targetOccupied(target) ? { replaceFingerprint: targetFingerprint(target) } : {}),
        },
        {
          signal: controller.signal,
          beforeAcceptance: () => {
            if (!operationCurrent(serial)) throw new FunctionInsertionRetired();
            acceptance.current = accepted;
            setChecking(false);
            setAccepting(true);
            abort.current = null;
            setStatus('Inserting formula…');
          },
        }
      );
      recentStore.recordAccepted(result.functionId);
      setRecentVersion((value) => value + 1);
      setAccepted(result);
      try {
        capturedOutcome?.(result);
      } catch (failure) {
        console.warn('xlsx.function-outcome', failure);
      }
      if (operationCurrent(serial)) {
        setStatus('Inserted into ' + context.binding.sheet + '!' + context.binding.target + '.');
        setTarget({ cell: { value: null, formula: result.formula } });
      }
    } catch (failure) {
      if (
        operationCurrent(serial) &&
        !(failure instanceof FunctionInsertionCancelled) &&
        !(failure instanceof FunctionInsertionRetired)
      ) {
        setStatus('');
        setError(failure instanceof Error ? failure.message : String(failure));
      }
    } finally {
      if (acceptance.current === accepted) {
        acceptance.current = null;
        setAccepting(false);
      }
      if (operationCurrent(serial)) {
        setChecking(false);
        abort.current = null;
      }
    }
  };
  if (acceptedVisible && accepted) {
    const result = acceptedTarget?.state === acknowledgedState ? acceptedTarget : null,
      calculation = result?.cell.calculation;
    let text = resultLoading
      ? 'Reading the accepted result…'
      : 'Result unavailable. Check the target to read its current accepted value.';
    if (calculation?.status === 'calculated' && calculation.value !== undefined)
      text =
        'Calculated (' +
        typeof calculation.value +
        '): ' +
        (typeof calculation.value === 'string' ? JSON.stringify(calculation.value) : String(calculation.value));
    else if (calculation?.status === 'error')
      text = 'Error: ' + (calculation.error ?? result?.cell.error ?? 'Formula error');
    else if (calculation?.status === 'circular')
      text = 'Circular reference: ' + (calculation.error ?? 'The formula refers to itself.');
    else if (calculation?.status === 'unresolved')
      text = 'Unresolved formula: ' + (calculation.error ?? 'This input cannot be calculated yet.');
    return (
      <section
        className="sheet-functions"
        aria-label={words('Insert Function')}
        data-function-document={accepted.binding.documentId}
        data-function-state={result?.state ?? accepted.binding.state}
        data-function-target={accepted.binding.target}
      >
        <p role="status">{words('Inserted into ' + accepted.binding.sheet + '!' + accepted.binding.target + '.')}</p>
        <output aria-label={words('Inserted formula')}>{'=' + accepted.formula}</output>
        <p
          role="status"
          aria-label={words('Accepted formula result')}
          data-status={calculation?.status ?? 'unavailable'}
        >
          {words(text)}
        </p>
        <button type="button" disabled={resultLoading || !context.readAcceptedTarget} onClick={checkAcceptedResult}>
          {words('Check result')}
        </button>
        {onRestart && (
          <button type="button" onClick={onRestart}>
            {words('Start another insertion')}
          </button>
        )}
      </section>
    );
  }
  if (!current)
    return (
      <section className="sheet-functions" aria-label={words('Insert Function')}>
        <p role="status">
          {words(
            runtimeCurrent && acceptance.current
              ? 'Finishing insertion…'
              : 'The initiating workbook or target changed. Reopen Insert Function.'
          )}
        </p>
      </section>
    );
  const frozen = checking || accepting;
  return (
    <section
      className="sheet-functions"
      aria-label={words('Insert Function')}
      data-function-document={context.binding.documentId}
      data-function-state={context.binding.state}
      data-function-target={context.binding.target}
    >
      <p>
        <strong>
          {context.binding.sheet}!{context.binding.target}
        </strong>{' '}
        · {words('The target stays fixed while this pane is open.')}
      </p>
      <label>
        {words('Search functions')}
        <input
          aria-label={words('Search functions')}
          maxLength={100}
          value={query}
          disabled={frozen}
          onChange={(event) => {
            const next = event.target.value;
            setQuery(next);
            const matches = matching(category, next);
            if (matches.length) changeFunction(matches[0].id);
          }}
        />
      </label>
      <label>
        {words('Function category')}
        <select
          aria-label={words('Function category')}
          value={category}
          disabled={frozen}
          onChange={(event) => {
            const next = event.target.value as typeof category;
            setCategory(next);
            const matches = matching(next);
            if (matches.length) changeFunction(matches[0].id);
          }}
        >
          <option value="all">{words('All supported categories')}</option>
          <option value="recent">{words('Recently Used')}</option>
          {categories.map((item) => (
            <option key={item.id} value={item.id}>
              {words(item.label)}
            </option>
          ))}
        </select>
      </label>
      <label>
        {words('Function')}
        <select
          aria-label={words('Function')}
          value={functions.some((item) => item.id === id) ? id : ''}
          disabled={frozen || !functions.length}
          onChange={(event) => changeFunction(event.target.value as FunctionId)}
        >
          {!functions.length && (
            <option value="">
              {words(category === 'recent' ? 'No recently used functions' : 'No matching functions')}
            </option>
          )}
          {functions.map((item) => (
            <option key={item.id} value={item.id}>
              {item.name}
            </option>
          ))}
        </select>
      </label>
      <p>{words(definition.summary)}</p>
      <details className="sheet-function-limits">
        <summary>{words('Supported inputs')}</summary>
        <p>
          {words(
            'Use the argument types shown below. Cell references stay on this worksheet and cover up to 1,000 cells. Financial inputs also allow the declared arithmetic Expression type. Nested functions and workbook links are not accepted by this picker.'
          )}
        </p>
        <ul>
          {supportedHints(definition).map((hint) => (
            <li key={hint}>{words(hint)}</li>
          ))}
        </ul>
      </details>
      <form
        onSubmit={(event) => {
          event.preventDefault();
          void submit();
        }}
      >
        <fieldset disabled={frozen || !functions.length}>
          <legend>{words('Function arguments')}</legend>
          {args.map((argument, index) => {
            const spec = definition.arguments[Math.min(index, definition.arguments.length - 1)],
              label =
                spec.label +
                (definition.repeatLast && index >= definition.arguments.length - 1 ? ' ' + (index + 1) : '');
            return (
              <div key={index}>
                <label>
                  {words(label + ' type')}
                  <select
                    aria-label={words(label + ' type')}
                    value={argument.kind}
                    onChange={(event) => {
                      const kind = event.target.value as ArgumentKind;
                      setArgs((values) =>
                        values.map((value, at) =>
                          at === index ? { kind, value: kind === 'boolean' ? 'false' : '' } : value
                        )
                      );
                    }}
                  >
                    {spec.allowedKinds.map((kind) => (
                      <option key={kind} value={kind}>
                        {words(kind)}
                      </option>
                    ))}
                  </select>
                </label>
                <label>
                  {words(label)}
                  {argument.kind === 'boolean' ? (
                    <select
                      aria-label={words(label)}
                      value={argument.value}
                      onChange={(event) =>
                        setArgs((values) =>
                          values.map((value, at) => (at === index ? { ...value, value: event.target.value } : value))
                        )
                      }
                    >
                      <option value="false">FALSE</option>
                      <option value="true">TRUE</option>
                    </select>
                  ) : (
                    <input
                      aria-label={words(label)}
                      value={argument.value}
                      maxLength={argument.kind === 'text' || argument.kind === 'expression' ? 1024 : 128}
                      placeholder={
                        argument.kind === 'reference'
                          ? spec.shape === 'scalar'
                            ? 'A1'
                            : 'A1:A10'
                          : argument.kind === 'number'
                            ? '0'
                            : argument.kind === 'expression'
                              ? 'B1/12'
                              : ''
                      }
                      onChange={(event) =>
                        setArgs((values) =>
                          values.map((value, at) => (at === index ? { ...value, value: event.target.value } : value))
                        )
                      }
                    />
                  )}
                </label>
              </div>
            );
          })}
          <button
            type="button"
            disabled={args.length >= definition.maxArgs}
            onClick={() => setArgs((values) => [...values, initialArgument(id, values.length, context)])}
          >
            {words('Add argument')}
          </button>
          <button
            type="button"
            disabled={args.length <= definition.minArgs}
            onClick={() => setArgs((values) => values.slice(0, -1))}
          >
            {words('Remove last argument')}
          </button>
        </fieldset>
        <p>
          {words('Formula preview')}{' '}
          <output aria-label={words('Formula preview')}>{formula || words(formulaError)}</output>
        </p>
        {formulaPreview && (
          <FormulaDraftPreview
            controller={formulaPreview.controller}
            consumer="functions"
            input={{
              scope: formulaPreview.scope,
              formula: formula.startsWith('=') ? formula.slice(1) : '',
              generation: previewDraft.current.generation,
              active: visible && current && scope.active && !accepting,
              invalidReason: formulaError || undefined,
            }}
            words={words}
          />
        )}
        <p>
          {words('Existing target')}{' '}
          <output aria-label={words('Existing target')}>
            {target
              ? words(targetDescription(target))
              : words(checking ? 'Reading target…' : 'Target has not been read.')}
          </output>
        </p>
        <button
          type="button"
          disabled={frozen || !scope.active || scope.literal || scope.specialSelection || scope.busy}
          onClick={readPreview}
        >
          {words('Check target')}
        </button>
        <p>
          {words(
            formulaPreview
              ? 'The preview leaves the workbook unchanged. Insert explicitly to write the formula.'
              : 'The workbook calculates the formula after insertion. Unsupported input is refused or kept explicitly unresolved.'
          )}
        </p>
        {unavailable && <p role="status">{words(unavailable)}</p>}
        {status && <p role="status">{words(status)}</p>}
        {error && <p role="alert">{words(error)}</p>}
        <button
          type="submit"
          disabled={!!unavailable || frozen || !target || !!formulaError || !functions.length}
          title={unavailable ? words(unavailable) : undefined}
        >
          {words(
            target && targetOccupied(target)
              ? 'Replace ' + context.binding.target + ' with formula'
              : 'Insert formula into ' + context.binding.target
          )}
        </button>
        <button type="button" disabled={accepting} onClick={cancel}>
          {words(accepting ? 'Finishing…' : 'Cancel insertion')}
        </button>
      </form>
    </section>
  );
}
