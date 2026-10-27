import { useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react';
import {
  createErrorCheckingController,
  errorScanScopeKey,
  type ErrorScanner,
  type ErrorScanScope,
  type CellIssue,
} from './errorChecking';
export interface WorkbookErrorsProps {
  scope: ErrorScanScope;
  scan: ErrorScanner;
  goToIssue: (reference: string) => void;
  words?: (text: string) => string;
}
const identity = (text: string) => text;
const issueLabel = (kind: CellIssue['kind']): string =>
  ({
    error: 'Cell error',
    circular: 'Circular formula',
    unresolved: 'Unresolved formula',
    stored: 'Stored formula result',
  })[kind];
export function WorkbookErrors({ scope, scan, goToIssue, words = identity }: WorkbookErrorsProps): JSX.Element {
  const key = errorScanScopeKey(scope),
    live = useRef({ key, scope, scan, goToIssue, generation: 0 });
  const generation = live.current.key === key ? live.current.generation : live.current.generation + 1;
  live.current = { key, scope, scan, goToIssue, generation };
  const controller = useMemo(
    () => createErrorCheckingController((bound, after, signal) => live.current.scan(bound, after, signal)),
    []
  );
  const snapshot = useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot);
  const [feedback, setFeedback] = useState(''),
    [selected, setSelected] = useState(-1),
    [pending, setPending] = useState(false);
  const intent = useRef(0),
    lifetime = useRef(0);
  const current = () => live.current.key === key && live.current.generation === generation && live.current.scope.active;
  const fresh = scope.active && !!snapshot.scope && errorScanScopeKey(snapshot.scope) === key;
  // Scope masking is immediate in render, before effect cleanup can run.
  const issues = fresh ? snapshot.issues : [],
    complete = fresh && snapshot.complete,
    found = fresh ? snapshot.issuesFound : 0,
    cursor = fresh ? snapshot.cursor : null;
  useLayoutEffect(() => {
    intent.current++;
    setPending(false);
    setSelected(-1);
    setFeedback('');
    try {
      controller.setScope(scope);
    } catch {
      setFeedback('Error Checking is unavailable for this worksheet.');
    }
    return () => {
      intent.current++;
      controller.cancel();
    };
  }, [key]);
  useLayoutEffect(() => {
    const epoch = ++lifetime.current;
    return () => {
      intent.current++;
      queueMicrotask(() => {
        if (lifetime.current === epoch) controller.dispose();
      });
    };
  }, []);
  const scanPage = async (restart = false): Promise<void> => {
    if (!current()) return;
    const action = ++intent.current;
    setFeedback('');
    setPending(true);
    setSelected(-1);
    try {
      if (restart) controller.restart();
      const result = await controller.scanNext();
      if (intent.current !== action || !current()) return;
      if (result.outcome === 'retired') setFeedback('The previous scan stopped. Start or continue the scan again.');
      else if (result.outcome === 'paused') setFeedback('Activate this worksheet to scan it.');
    } catch {
      if (intent.current === action && current())
        setFeedback('The worksheet could not be scanned. Continue or restart to try again.');
    } finally {
      if (intent.current === action && current()) setPending(false);
    }
  };
  const go = (index: number): void => {
    if (!current() || !fresh || controller.getSnapshot() !== snapshot || !issues[index]) return;
    setSelected(index);
    try {
      goToIssue(issues[index].reference);
    } catch {
      if (current()) setFeedback('The cell could not be selected. Scan again to try it.');
    }
  };
  const busy = fresh && (pending || snapshot.phase === 'scanning');
  const summary = !scope.active
    ? 'Paused. Reopen this pane in Normal view to scan.'
    : busy
      ? 'Scanning this worksheet…'
      : complete
        ? found
          ? 'Scan complete: ' + found + ' issues found.'
          : 'Scan complete: no issues found.'
        : 'Scan incomplete.';
  const limit = fresh ? snapshot.incompleteReason : null;
  return (
    <section
      className="sheet-errors"
      aria-label={words('Error Checking')}
      data-error-document={scope.documentId}
      data-error-owner={scope.ownerIdentity}
      data-error-state={scope.state}
      data-error-sheet={scope.sheet}
      data-error-cursor={cursor ?? ''}
    >
      <p>
        {words(
          'Checks cell errors and formula result status in this worksheet. It does not change formulas or values.'
        )}
      </p>
      <div className="sheet-error-actions">
        <button type="button" disabled={!scope.active || busy || complete} onClick={() => void scanPage()}>
          {words(cursor ? 'Continue scan' : 'Scan worksheet')}
        </button>
        <button type="button" disabled={!scope.active || busy} onClick={() => void scanPage(true)}>
          {words('Restart scan')}
        </button>
        {busy && (
          <button
            type="button"
            onClick={() => {
              if (current()) {
                intent.current++;
                controller.cancel();
                setPending(false);
                setFeedback('Scan cancelled. Results already shown are kept.');
              }
            }}
          >
            {words('Cancel scan')}
          </button>
        )}
        <button
          type="button"
          disabled={!scope.active || !fresh || !issues.length || selected >= issues.length - 1}
          onClick={() => go(selected + 1)}
        >
          {words('Next issue')}
        </button>
      </div>
      <p role="status">{words(summary)}</p>
      {fresh && (
        <p role="status">
          {issues.length} {words('issues shown')} · {found} {words('found so far')} · {snapshot.processedCells}{' '}
          {words('cells checked')}
        </p>
      )}
      {limit === 'issue-limit' && (
        <p role="status">
          {words('This page reached 100 issues. Continue scan resumes after')} {cursor}.
        </p>
      )}
      {limit === 'page-limit' && (
        <p role="status">{words('This step reached the scan limit. Continue scan checks the remaining cells.')}</p>
      )}
      {limit === 'cancelled' && !feedback && (
        <p role="status">{words('Scan cancelled; results already shown are kept.')}</p>
      )}
      {complete && !issues.length && found > 0 && (
        <p>{words('No issues on this page. Earlier pages contain the reported issues.')}</p>
      )}
      {feedback && scope.active && <p role="alert">{words(feedback)}</p>}
      <ol aria-label={words('Worksheet issues')}>
        {issues.map((issue, index) => (
          <li
            key={issue.reference}
            data-error-reference={issue.reference}
            data-error-kind={issue.kind}
            data-selected={selected === index}
          >
            <strong>
              {scope.sheet}!{issue.reference}
            </strong>
            <span>{words(issueLabel(issue.kind))}</span>
            <output aria-label={words('Issue reason') + ' ' + scope.sheet + '!' + issue.reference}>
              {issue.kind === 'stored' ? words('Stored result; not recalculated.') : issue.reason}
            </output>
            {issue.formula && (
              <output aria-label={words('Issue formula') + ' ' + scope.sheet + '!' + issue.reference}>
                {'=' + issue.formula}
              </output>
            )}
            <button
              type="button"
              disabled={!scope.active || !fresh}
              aria-label={words('Go to issue') + ' ' + scope.sheet + '!' + issue.reference}
              onClick={() => go(index)}
            >
              {words('Go to cell')}
            </button>
          </li>
        ))}
      </ol>
    </section>
  );
}
